# -*- coding: utf-8 -*-
"""
Versión en inglés de assemble.py. Misma arquitectura (anclas por regex + regeneración completa de
cada párrafo cubierto, sin find/replace de texto viejo), pero:
  - las anclas apuntan al texto en INGLÉS (verificadas contra Earning_Report_June_2026_27_08_26.docx,
    el Word en inglés real de junio 2026 — el mismo trimestre que ar_data/ar.docx);
  - usa ar_engine_en.py / periods_en.py en vez de ar_engine.py / periods.py.

Limitación actual (ver README / mensaje a Melissa): las TABLAS (imágenes) todavía no se regeneran en
inglés — collect_and_render_tables_en reutiliza las imágenes que ya trae el Word base en inglés tal
cual, así que por ahora este motor solo produce un resultado 100% correcto si el Word base en inglés
corresponde al MISMO trimestre que el Excel (como Jun26, el único trimestre con ambos disponibles).
Generar un trimestre distinto al del Word base en inglés regenerará todo el texto correctamente, pero
las tablas (imágenes) quedarán con los valores del Word base. Falta: tablas en inglés con header/
formato numérico en-US (ver render_tables_v5.py) y el traductor Excel ES->EN.
"""
import re, copy, io, os, sys, zipfile, tempfile, shutil, difflib
import docx
from docx.enum.text import WD_COLOR_INDEX
import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
import ar_engine_en as E
import periods_en as P
import assemble as ASM_ES  # motor en español: se reutiliza collect_causes_by_highlight (ver generate_en)
import deepl_translate as DT  # traducción del texto narrativo (ver translate_narrative_from_definitive_es)

MARK_RE = re.compile(r'\[\[([^\]]+)\]\]|\{\{([^}]+)\}\}')

def set_paragraph_text(paragraph, text, force_bold=None, force_underline=None):
    # force_bold/force_underline: para los pocos títulos/encabezados de HEADER_TEMPLATES_EN donde se
    # conoce con certeza (verificado contra el Word en inglés real de Hortifrut de Jun26) si deben ir
    # en negrita/subrayado o no — se usa en vez de heredar del run actual porque, en el Word base en
    # inglés que Melissa ha ido subiendo trimestre a trimestre, varios de estos títulos YA perdieron su
    # negrita en una vuelta anterior de este mismo pipeline (antes de este fix), así que heredar del
    # run ya no basta para corregirlos esta vez (ver apply_header_templates).
    #
    # Conserva fuente/tamaño/subrayado/NEGRITA del primer run original. La negrita se agregó acá
    # (feedback Melissa, Oct26, punto 6: "En inglés falta tener el mismo formato de subtitulos y
    # titulos que el texto en español, como por ejemplo el resaltado en negrita o subrayado") — a
    # diferencia de assemble.py (lado español), donde los subtítulos de sección fijos (ej.
    # "Principales Componentes de los Costos de Ventas", en negrita en la plantilla base) casi nunca
    # se regeneran (no están cubiertos por ningún ancla de SINGLE_TEMPLATES/HEADER_TEMPLATES, así que
    # conservan intacto su run original), en inglés translate_narrative_from_definitive_es SÍ pasa por
    # TODO el texto narrativo del documento (por pedido explícito de Melissa de traducir línea a línea
    # TODO, ver docstring de esa función) — incluidos esos mismos subtítulos fijos, que en inglés
    # también vienen en negrita en la plantilla base. Antes esta función siempre creaba el run nuevo
    # SIN negrita (solo heredaba el subrayado), así que cada vez que translate_narrative_from_
    # definitive_es tocaba uno de esos subtítulos, perdía la negrita que si tenía el lado español.
    # Nota: se probó además un "último recurso" que leía <w:b/> de la marca de párrafo (w:pPr/w:rPr)
    # cuando el run ya no traía negrita directa, pensado para auto-reparar subtítulos que ya la habían
    # perdido en una vuelta anterior de este mismo pipeline (antes de este fix). Se descartó: en los
    # .docx reales de Melissa esa marca de párrafo también queda, con la misma pinta (bold+underline
    # juntos), en párrafos que NO deberían ser negrita — una viñeta del bridge ("Mayor gasto por
    # depreciación...") y una oración narrativa normal ("Reducción en activos no corrientes...",
    # insertada por el ajuste de estructura párrafo a párrafo, ver generate_en) — ambas confirmadas con
    # negrita de sobra al probarlo. Mejor conservar solo lo que el run de texto YA trae correctamente
    # (de aquí en adelante no se vuelve a perder) que arriesgarse a negritas de más en cualquier parte
    # del documento.
    base_font_name = None; base_size = None; base_underline = None; base_bold = None
    if paragraph.runs:
        r0 = paragraph.runs[0]
        base_font_name = r0.font.name; base_size = r0.font.size
        base_underline = r0.font.underline
        base_bold = r0.bold
    if force_bold is not None:
        base_bold = force_bold
    if force_underline is not None:
        base_underline = force_underline
    for r in list(paragraph.runs):
        r._element.getparent().remove(r._element)

    pos = 0
    for m in MARK_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], base_font_name, base_size, base_bold, False, base_underline)
        if m.group(1) is not None:
            _add_run(paragraph, m.group(1), base_font_name, base_size, base_bold, True, base_underline)
        else:
            _add_run(paragraph, m.group(2), base_font_name, base_size, True, False, base_underline)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], base_font_name, base_size, base_bold, False, base_underline)

def _add_run(paragraph, text, font_name, size, bold, highlight, underline=None):
    run = paragraph.add_run(text)
    if font_name: run.font.name = font_name
    if size: run.font.size = size
    if bold is not None: run.bold = bold
    if underline: run.font.underline = True
    if highlight:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return run

def find_paragraph(doc, anchor_regex, start_from=0):
    rx = re.compile(anchor_regex)
    for i, p in enumerate(doc.paragraphs):
        if i < start_from: continue
        if rx.search(p.text):
            return i, p
    return None, None

# ---------- Extracción de causas de negocio ya escritas en el Word base en inglés ----------
# Verificadas contra Earning_Report_June_2026_27_08_26.docx (ver /tmp/en_june_paragraphs.txt). Igual
# que en assemble.py: si el patrón no calza en un Word base futuro (porque cambió la redacción, p.ej.
# porque el Word base ya es una salida previa de esta misma herramienta en "estilo Jun26"), simplemente
# no se encuentra causa y el párrafo queda solo con el CAUSA_TAG para completar a mano.

def _extract_deterioro_6m_causes_en(text):
    cause_cur = None; cause_prior = None
    m = re.search(r'million,\s*(mainly associated.*?),\s*compared with', text)
    if m: cause_cur = m.group(1).strip()
    m2 = re.search(r'as of \w+\d+,\s*(mainly associated.*?)\.\s*$', text)
    if m2: cause_prior = m2.group(1).strip()
    return cause_cur, cause_prior

def _extract_gastosadmin_6m_cause_en(text):
    m = re.search(r'compared to \w+\d+,\s*(mainly explained.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_costoventa_12m_cause_en(text):
    m = re.search(r'compared with [\d.]+%\s*in the \d+M S\d+/\d+,\s*([^\.]+)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_razon_endeudamiento_cause_en(text):
    m = re.search(r'total liabilities,\s*(mainly associated.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_cobertura_gf_cause_en(text):
    m = re.search(r'compared to \w+\d+,\s*(mainly associated.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_bridge_deterioro_cause_en(text):
    m = re.search(r'million\s+(mainly associated.*?),\s*compared with', text)
    return m.group(1).strip() if m else None

def _extract_bridge_depreciacion_cause_en(text):
    m = re.search(r'\),\s*(mainly explained.*?|explained.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_bridge_costosfin_cause_en(text):
    m = re.search(r'million,\s*(mainly associated.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

LABEL_KEYWORDS_EN = {
    'Deterioro': ['impairment'],
    'Depreciación y amortizaciones': ['depreciation and amortization'],
    'Costos Financieros Netos': ['net financial expenses', 'net financial costs'],
    'Diferencia de cambio': ['exchange rate difference'],
    'Gasto por impuestos a las ganancias': ['taxes'],
    'Otros ingresos (gastos) no operacionales': ['other income', 'other non-operating expenses', 'other non-operating income'],
    'Participación en asociadas': ['share of profit of associates', 'share in associates'],
}

def classify_bridge_label(text):
    t = text.lower()
    for label, kws in LABEL_KEYWORDS_EN.items():
        if any(kw in t for kw in kws):
            return label
    return None

# Igual que en assemble.py: los templates siempre se generan en "estilo Jun26" (12M, 6 meses) y
# periods_en.localize_en() los convierte al conteo real del cierre. Las anclas, en cambio, deben
# aceptar cualquier conteo de meses porque anclan sobre el texto del documento BASE.
SINGLE_TEMPLATES_EN = [
    (r"^EBITDA as of \w+\d+ reached US\$", E.p_ebitda_6m),
    (r"^Accumulated EBITDA as of \w+\d+ without the fair value effect", E.p_ebitda_sinfv_6m),
    (r"^During the \d+ months ending \w+\d+,", E.p_ganancia_controladora_6m),
    (r"^During the \d+M S\d+/\d+, EBITDA reached US\$", E.p_ebitda_12m_intro),
    (r"^(Increase|Decrease) in income by [\d.]+% associated", E.p_ebitda_12m_bullet_ingresos),
    (r"^Costs and expenses without including depreciation recorded", E.p_ebitda_12m_bullet_costos),
    (r"^Profit \(loss\) attributable to the parent company shareholders was US\$[\d\.,\-]+ million in the \d+M", E.p_ganancia_controladora_12m),
    (r"^Income from operating activities reached US\$[\d,\.]+ million as of \w+\d+,", E.p_ingresos_totales_6m),
    (r"^Sales from the Fresh Fruit segment as of \w+\d+", E.p_frutafresca_6m),
    (r"^Value added products recorded (a|an) [\d.]+% increase in sales income as of \w+\d+", E.p_valoragregado_6m),
    (r"^Income from operating activities reached US\$[\d,\.]+ million in the \d+M", E.p_ingresos_totales_12m),
    (r"^The Fresh Fruit sales segment in the \d+M", E.p_frutafresca_12m),
    (r"^Also, the value added products segment recorded", E.p_valoragregado_12m),
    (r"^Sales costs as of \w+\d+ totaled US\$", E.p_costoventa_6m),
    (r"^Sales costs in the \d+M S\d+/\d+ reached US\$", E.p_costoventa_12m),
    (r"^Administrative expenses as of \w+\d+ reached US\$", E.p_gastosadmin_6m),
    (r"^Other expenses, per function \(excluding impairment", E.p_otrosgastos_6m),
    (r"^The expense due to impairment in the value of assets as of \w+\d+ was", E.p_deterioro_6m),
    (r"^The other components of the income statement recorded", E.p_otroscomponentes_6m),
    (r"^As of \w+\d+, (a|the) (gains|income) tax expense", E.p_impuesto_6m),
    (r"^The rotation of assets as of \w+ \d+, \d+", E.p_rotacion_activos),
    (r"^Also, the rotation ratio of inventories", E.p_rotacion_inventarios),
    (r"^The Company.s net financial debt (was|has been)", E.p_dfn),
    (r"^Current liquidity was [\d.]+ times as of", E.p_liquidez),
    (r"^In the meantime, the acid ratio reached", E.p_razon_acida),
    (r"^The debt ratio (decreased|increased)", E.p_razon_endeudamiento),
    (r"^During this period, the company (maintained|modified|refinanced)", E.p_estructura_deuda),
    (r"^The financial expense hedge index is at", E.p_cobertura_gf),
    (r"^The profitability of the parent company equity", E.p_rentabilidad_patrim_controladora),
    (r"^Also, the profitability of total equity", E.p_rentabilidad_patrim_total),
    (r"^As of \w+ \d+, \d+, total assets (decreased|increased)", E.p_activos_totales_intro),
    (r"(?i)^(Reduction|Increase) in non-current assets", E.p_activos_nocorrientes_bridge),
    (r"(?i)^(Reduction|Increase) in current assets", E.p_activos_corrientes_bridge),
    (r"^Total liabilities (decreased|increased)", E.p_pasivos_bridge),
    (r"^The Company.s total equity (increased|decreased)", E.p_patrimonio_bridge),
]

# Preámbulo y títulos de sección: no dependen del Excel (solo del período), toman `periods` directamente.
# Cada entrada trae además (negrita, subrayado) verificados contra el Word en inglés real de Hortifrut
# de Jun26 (Earning_Report_June_2026_27_08_26.docx) — se fuerzan explícitamente (ver
# set_paragraph_text/force_bold) en vez de heredarlos del run actual porque, en el Word base en inglés
# que Melissa ha ido subiendo trimestre a trimestre, varios de estos títulos YA perdieron su negrita en
# una vuelta anterior de este mismo pipeline (feedback Melissa, Oct26, punto 6) — heredar del run ya
# no basta para corregirlos esta vez, solo para no perderlos de nuevo en adelante.
HEADER_TEMPLATES_EN = [
    (r"^As of \w+ \d+, \d+\s*$", E.p_titulo_fecha, True, True),
    (r"^The current reasoned analysis has been prepared for the period ending", E.p_preambulo_1, False, None),
    (r"^Since the Company administers its operations with an agricultural season", E.p_preambulo_2, False, None),
    (r"^Accumulated EBITDA analysis as of \w+", E.p_header_ebitda_acumulado, False, True),
    (r"^Income Statement analysis as of \w+", E.p_header_resultado_calendario, False, True),
    (r"^Income Statement Analysis .* season\s*$", E.p_header_resultado_temporada, False, True),
    (r"^Accumulated income analysis as of \w+", E.p_header_ingresos_acumulados, False, True),
    # Feedback Melissa (Oct26): el Word base en inglés traía este título como "Sensitivity analysis"
    # (genérico, de una plantilla antigua) en vez del que de verdad corresponde a esta sección y que
    # calza con el español ("Efectos en una Potencial Variación del Tipo de Cambio en las Cuentas
    # Monetarias de Balances") — ver p_header_sensibilidad_tc. El regex acepta también el título ya
    # correcto (como en el Word de referencia Jun26) para que no haga nada si ya viene bien.
    (r"^(Sensitivity [Aa]nalysis|Effects of a Potential Exchange Rate Fluctuation on the Monetary "
     r"Accounts of the Balance Sheet)\s*$", E.p_header_sensibilidad_tc, True, False),
]

def apply_header_templates(doc, periods, log):
    for anchor, fn, force_bold, force_underline in HEADER_TEMPLATES_EN:
        idx, para = find_paragraph(doc, anchor)
        if para is None:
            log.append(f"[AVISO] (EN) No encontré párrafo para patrón: {anchor[:60]}")
            continue
        try:
            set_paragraph_text(para, fn(periods), force_bold=force_bold, force_underline=force_underline)
            log.append(f"[OK] (EN) Actualizado (encabezado/preámbulo): {anchor[:50]}")
        except Exception as ex:
            log.append(f"[ERROR] (EN) {anchor[:50]}: {ex}")

# STATIC_SUBTITLES_EN: subtítulos/títulos de sección que NO dependen del período (texto idéntico todos
# los trimestres, a diferencia de HEADER_TEMPLATES_EN de arriba) — por eso no necesitan una función
# fn(periods) que genere el texto, solo re-forzar negrita/subrayado sobre el texto que YA está ahí (ver
# apply_static_subtitle_formatting_en). Lista armada comparando contra el Word en inglés oficial de
# referencia, nunca tocado por este pipeline (Marzo 2026, 648d7711): cada tupla trae (patrón,
# force_bold, force_underline) verificados ahí mismo.
#
# Por qué además hace falta EXCLUIRLOS de _segment_narrative (ver su uso más abajo, en
# translate_narrative_from_definitive_es): se confirmó con un caso real que, sin excluirlo, un título
# como "RISK FACTOR ANALYSIS" puede terminar reutilizado como párrafo DONANTE para una viñeta española
# totalmente distinta (ej. "Reducción en activos no corrientes...") — el emparejamiento por FORMA
# (_segment_signature, ver _align_segments) no distingue un título real de una viñeta común, así que un
# título de una sola viñeta puede calzar por pura coincidencia estructural con el conteo de un bloque
# de contenido de otra sección completamente distinta. El resultado observado fue el título
# desapareciendo del documento por completo (su texto quedaba sobrescrito) — no solo un problema de
# negrita. Excluirlos de la segmentación los deja protegidos igual que ya protege a HEADER_TEMPLATES_EN.
STATIC_SUBTITLES_EN = [
    (r"^REASONED ANALYSIS OF HORTIFRUT", True, True),
    (r"^HIGHLIGHTS OF THE PERIOD AND SUBSEQUENT FACTS\s*$", True, True),
    (r"^SUMMARY OF THE PERIOD\s*$", True, True),
    (r"^INCOME STATEMENT ANALYSIS\s*$", True, True),
    (r"^Main Components of Income\s*$", True, False),
    (r"^Main Components of Costs and Expenses\s*$", True, False),
    (r"^Main Components of Sales Costs\s*$", True, True),
    (r"^Main Components of Administrative Expenses\s*$", True, True),
    (r"^Main Components of Other Expenses per function\s*$", True, True),
    (r"^Impairment in the value of assets\s*$", True, True),
    (r"^Other Components of Income Statement\s*$", True, False),
    (r"^(Income|Gains) tax expense\s*$", True, False),
    (r"^Other Result Indicators\s*$", True, False),
    (r"^Activity Indicators:?\s*$", True, False),
    (r"^COMPARATIVE ANALYSIS OF MAIN TENDENCIES\s*$", True, False),
    (r"^Net Financial Debt\s*$", True, False),
    (r"^[a-zA-Z]\)\s*Main Financial Indicators\s*$", True, False),
    (r"^\d+\.\s*ANALYSIS OF STATEMENT OF FINANCIAL POSITION\s*$", True, True),
    (r"^Main items of the Consolidated Statement of Financial Position\s*$", True, True),
    (r"^MAIN COMPONENTS OF NET OPERATING INVESTMENT AND FINANCING CASH FLOWS\s*$", True, True),
    (r"^RISK FACTOR ANALYSIS\s*$", True, True),
    (r"^Liquidity and Financing Risk\s*$", True, False),
    (r"^Exchange Rate Risk\s*$", True, False),
    (r"^Exposure in currencies different to the functional currency\s*$", True, False),
    (r"^Interest Rate Risk\s*$", True, False),
    (r"^Operating Risks\s*$", True, False),
    (r"^Insurance\s*$", True, False),
    (r"^Risk in the Estimations\s*$", True, False),
    (r"^Risk Associated to Merger and Acquisition Process\s*$", True, False),
]

def apply_static_subtitle_formatting_en(doc, log):
    """Refuerza negrita/subrayado (STATIC_SUBTITLES_EN) sobre los subtítulos/títulos de sección que no
    cambian de un trimestre a otro, SIN tocar su texto (a diferencia de apply_header_templates, que sí
    reescribe el texto porque esos sí dependen del período). Se corre después del barrido narrativo
    (translate_narrative_from_definitive_es) para restaurar la negrita/subrayado perdida en una vuelta
    anterior de este mismo pipeline (feedback Melissa, Oct26, punto 6), sin depender de heredar el
    formato de un párrafo reutilizado (ver docstring de STATIC_SUBTITLES_EN para el motivo)."""
    compiled = [(re.compile(pat), b, u) for pat, b, u in STATIC_SUBTITLES_EN]
    n = 0
    for para in doc.paragraphs:
        t = para.text.strip()
        if not t:
            continue
        for rx, force_bold, force_underline in compiled:
            if rx.search(t):
                for r in para.runs:
                    if r.text.strip():
                        r.bold = force_bold
                        r.underline = force_underline
                n += 1
                break
    if n:
        log.append(
            f"[OK] (EN) Formato de {n} subtítulo(s)/título(s) de sección reforzado (negrita/subrayado "
            "verificados contra el Word en inglés oficial de referencia), sin tocar su texto."
        )

def collect_original_causes(doc, log):
    causes = {}
    anchor_by_fn = {fn: anchor for anchor, fn in SINGLE_TEMPLATES_EN}

    def _text_for(fn):
        anchor = anchor_by_fn.get(fn)
        if not anchor:
            return None
        _, para = find_paragraph(doc, anchor)
        return para.text if para is not None else None

    t = _text_for(E.p_deterioro_6m)
    if t:
        cur, prior = _extract_deterioro_6m_causes_en(t)
        causes['deterioro_6m_cur'] = cur; causes['deterioro_6m_prior'] = prior
        if cur or prior:
            log.append("[OK] (EN) Causa original preservada: deterioro de activos 6M")

    t = _text_for(E.p_gastosadmin_6m)
    if t:
        cause = _extract_gastosadmin_6m_cause_en(t)
        if cause:
            causes['gastosadmin_6m'] = cause
            log.append("[OK] (EN) Causa original preservada: gastos de administración 6M")

    t = _text_for(E.p_costoventa_12m)
    if t:
        cause = _extract_costoventa_12m_cause_en(t)
        if cause:
            causes['costoventa_12m'] = cause
            log.append("[OK] (EN) Causa original preservada: costo de ventas 12M")

    t = _text_for(E.p_razon_endeudamiento)
    if t:
        cause = _extract_razon_endeudamiento_cause_en(t)
        if cause:
            causes['razon_endeudamiento'] = cause
            log.append("[OK] (EN) Causa original preservada: razón de endeudamiento")

    t = _text_for(E.p_cobertura_gf)
    if t:
        cause = _extract_cobertura_gf_cause_en(t)
        if cause:
            causes['cobertura_gf'] = cause
            log.append("[OK] (EN) Causa original preservada: cobertura de gastos financieros")

    causes['bridge_12m'] = {}
    idx, intro = find_paragraph(doc, r"^Profit \(loss\) attributable to the parent company shareholders was US\$")
    if intro is not None:
        pos_paras, after_pos = _next_list_block(doc, idx + 1, 14)
        j = after_pos
        n = len(doc.paragraphs)
        while j < n and doc.paragraphs[j].text.strip() == '':
            j += 1
        sep_idx = j if (j < n and 'offset' in doc.paragraphs[j].text.lower()) else None
        neg_paras = []
        if sep_idx is not None:
            neg_paras, _ = _next_list_block(doc, sep_idx + 1, 14)
        for p in (pos_paras + neg_paras):
            label = classify_bridge_label(p.text)
            cause = None
            if label == 'Deterioro':
                cause = _extract_bridge_deterioro_cause_en(p.text)
            elif label == 'Depreciación y amortizaciones':
                cause = _extract_bridge_depreciacion_cause_en(p.text)
            elif label == 'Costos Financieros Netos':
                cause = _extract_bridge_costosfin_cause_en(p.text)
            if label and cause:
                causes['bridge_12m'][label] = cause
        if causes['bridge_12m']:
            log.append(f"[OK] (EN) Causas originales preservadas en bridge 12M: {', '.join(causes['bridge_12m'].keys())}")
    return causes

def _call_single_template(fn, wb, causes):
    if fn is E.p_deterioro_6m:
        return fn(wb, causes.get('deterioro_6m_cur'), causes.get('deterioro_6m_prior'))
    if fn is E.p_gastosadmin_6m:
        return fn(wb, causes.get('gastosadmin_6m'))
    if fn is E.p_costoventa_12m:
        return fn(wb, causes.get('costoventa_12m'))
    if fn is E.p_razon_endeudamiento:
        return fn(wb, causes.get('razon_endeudamiento'))
    if fn is E.p_cobertura_gf:
        return fn(wb, causes.get('cobertura_gf'))
    return fn(wb)

POSITIONAL_FALLBACK_AFTER = {
    E.p_ebitda_12m_bullet_costos: E.p_ebitda_12m_bullet_ingresos,
}

def apply_single_templates(doc, wb, periods, causes, log):
    matched_idx = {}
    for anchor, fn in SINGLE_TEMPLATES_EN:
        idx, para = find_paragraph(doc, anchor)
        if para is None and fn in POSITIONAL_FALLBACK_AFTER:
            ref_idx = matched_idx.get(POSITIONAL_FALLBACK_AFTER[fn])
            if ref_idx is not None:
                j = ref_idx + 1
                while j < len(doc.paragraphs) and doc.paragraphs[j].text.strip() == '':
                    j += 1
                if j < len(doc.paragraphs):
                    idx, para = j, doc.paragraphs[j]
        if para is None:
            log.append(f"[AVISO] (EN) No encontré párrafo para patrón: {anchor[:60]}")
            continue
        matched_idx[fn] = idx
        try:
            # Nota: la mayoría de las funciones de ar_engine_en.py usan los tokens placeholder
            # "estilo Jun26" que localize_en() sustituye recién aquí (igual que el motor en español).
            # Pero los 5 preámbulo/header ya reciben `periods` directo (no pasan por localize_en).
            text = P.localize_en(_call_single_template(fn, wb, causes), periods)
            set_paragraph_text(para, text)
            log.append(f"[OK] (EN) Actualizado: {anchor[:50]}")
        except Exception as ex:
            log.append(f"[ERROR] (EN) {anchor[:50]}: {ex}")

# ---------- Bridge bullets (Ganancia Atribuible) ----------
def _next_list_block(doc, start, max_scan=14):
    n = len(doc.paragraphs)
    j = start
    while j < n and j < start + max_scan and doc.paragraphs[j].text.strip() == '':
        j += 1
    items = []
    while j < n and j < start + max_scan:
        p = doc.paragraphs[j]
        style_name = p.style.name if p.style else ''
        if 'List' in style_name and p.text.strip() != '':
            items.append(p)
            j += 1
        else:
            break
    return items, j

def _apply_numpr_from_source(target_paras, source_paras):
    """Copia el numPr (numId/ilvl — bullet simple vs. "a) b) c)", etc.) de cada párrafo de
    `source_paras` (español, ya con su propio numPr corregido vía apply_bridge_bullets/
    _normalize_bullet_numbering del lado ES) al párrafo correspondiente de `target_paras` (inglés),
    uno a uno. Se usa en vez de ASM_ES._normalize_bullet_numbering (que normaliza una lista CONTRA SÍ
    MISMA, usando su propio primer párrafo como referencia) porque en inglés cada viñeta narrativa
    suele quedar aislada en su propio segmento de un solo párrafo (separada de sus vecinas por varias
    líneas en blanco en el Word base, ver docstring de translate_narrative_from_definitive_es) — sin
    ningún hermano en el MISMO segmento con el que comparar, normalizar-contra-sí-mismo no corrige
    nada. Copiar del español es además justo lo que pidió Melissa (Oct26): "la estructura debe ser
    la misma entre el texto definitivo en español y la versión traducida en inglés". Si
    `target_paras` tiene más párrafos que `source_paras` (una viñeta de más en inglés), los
    sobrantes se alinean con el último numPr conocido de `source_paras`."""
    if not source_paras:
        ASM_ES._normalize_bullet_numbering(target_paras)
        return
    last_ref = None
    for i, p in enumerate(target_paras):
        src = source_paras[i] if i < len(source_paras) else source_paras[-1]
        ref = ASM_ES._capture_numpr(src)
        if ref is not None:
            last_ref = ref
        ASM_ES._apply_numpr(p, last_ref)

def _replace_list_items(bullet_paras, items_texts, numpr_source=None, neutralize_existing=False):
    # numpr_source (lista de párrafos en ESPAÑOL alineados 1 a 1, opcional): si se pasa, el numId/
    # ilvl de cada párrafo en inglés se copia del español correspondiente en vez de normalizarse
    # contra el primer párrafo de esta misma lista (ver _apply_numpr_from_source) — mismo bug
    # reportado por Melissa en el lado español (un "efecto" puntual quedaba con numeración "a) b)
    # c)" en vez de viñeta simple, arrastrado desde el Word base) puede venir igual en el Word base
    # en inglés, y set_paragraph_text nunca toca el numPr, solo el texto.
    #
    # neutralize_existing: usado por translate_narrative_from_definitive_es (alineación por FORMA,
    # no por contenido, ver _align_segments/_segment_signature). Ahí un párrafo "existente" que se
    # reutiliza no es necesariamente del mismo tipo de contenido que la viñeta española que le toca:
    # confirmado con un caso real donde la viñeta en inglés reutilizada para "Reducción en activos no
    # corrientes..." resultó ser, por puro accidente de alineación estructural, el encabezado
    # "RISK FACTOR ANALYSIS " (negrita+subrayado, de una sección totalmente distinta del documento) —
    # mantener su negrita/subrayado como si aplicara al contenido NUEVO fue el bug. Con
    # neutralize_existing=True se despoja ese párrafo de negrita/subrayado directo ANTES de
    # reescribirle el texto, igual que ya se hace con los párrafos recién clonados (ver
    # _neutralize_cloned_paragraph_format). No se activa por defecto porque otros llamadores (ej. las
    # viñetas EBITDA/bridge en _replace_bridge_span_split) sí reutilizan viñetas contiguas del MISMO
    # bloque, donde conservar su formato original es correcto.
    n_existing = len(bullet_paras)
    n_target = len(items_texts)
    for k in range(min(n_existing, n_target)):
        if neutralize_existing:
            _neutralize_cloned_paragraph_format(bullet_paras[k])
        set_paragraph_text(bullet_paras[k], items_texts[k])
    if n_target > n_existing:
        final = list(bullet_paras)
        last_el = bullet_paras[-1]._p
        for k in range(n_existing, n_target):
            new_el = copy.deepcopy(last_el)
            last_el.addnext(new_el)
            new_para = docx.text.paragraph.Paragraph(new_el, bullet_paras[-1]._parent)
            # Igual que en _clone_paragraph_after/_before: un ítem de viñeta recién clonado no debe
            # heredar negrita/subrayado directo del run del ÚLTIMO ítem existente (que puede traer esa
            # marca por un accidente de formato de otro trimestre, no porque corresponda a contenido
            # nuevo) — confirmado con el mismo caso real ("Reducción en activos no corrientes...").
            _neutralize_cloned_paragraph_format(new_para)
            set_paragraph_text(new_para, items_texts[k])
            final.append(new_para)
            last_el = new_el
        result = final
    elif n_existing > n_target:
        result = list(bullet_paras[:n_target])
        for p in bullet_paras[n_target:]:
            p._p.getparent().remove(p._p)
    else:
        result = list(bullet_paras)
    if numpr_source is not None:
        _apply_numpr_from_source(result, numpr_source)
    else:
        ASM_ES._normalize_bullet_numbering(result)
    return result

def build_bridge_items_6m(wb, max_items=6):
    ebitda_text = E.p_ebitda_bullet_6m(wb)
    items = E.bridge_bullets(wb, '6m', threshold_abs_mus=3.0)[:max_items]
    rest = [E.bridge_factor_text(lbl, cur, prior, delta, '6m') for lbl, cur, prior, delta in items]
    return [ebitda_text] + rest

def _is_list_para(p):
    style_name = p.style.name if p.style else ''
    return 'List' in style_name and p.text.strip() != ''

def _collect_bridge_span(doc, start, max_scan=24):
    n = len(doc.paragraphs)
    j = start
    while j < n and j < start + max_scan and doc.paragraphs[j].text.strip() == '':
        j += 1
    span = []
    while j < n and j < start + max_scan:
        p = doc.paragraphs[j]
        if p.text.strip() == '':
            j += 1
            continue
        if _is_list_para(p):
            span.append(p); j += 1
            continue
        k = j + 1
        while k < n and k < start + max_scan and doc.paragraphs[k].text.strip() == '':
            k += 1
        if k < n and k < start + max_scan and _is_list_para(doc.paragraphs[k]):
            span.append(p); j += 1
            continue
        break
    return span, j

def _replace_bridge_span_split(doc, start, pos_texts, neg_texts, sep_sentence, intro_para, max_scan=24):
    span, _ = _collect_bridge_span(doc, start, max_scan)
    bullet_paras = [p for p in span if _is_list_para(p)]
    sep_paras = [p for p in span if not _is_list_para(p)]
    if not bullet_paras:
        return False

    reusable_sep = None
    if sep_paras:
        reusable_sep = sep_paras[0]
        for p in sep_paras[1:]:
            p._p.getparent().remove(p._p)

    if not neg_texts:
        _replace_list_items(bullet_paras, pos_texts)
        if reusable_sep is not None:
            reusable_sep._p.getparent().remove(reusable_sep._p)
        return True

    n_pos = len(pos_texts)
    pos_paras = bullet_paras[:n_pos]
    neg_paras = bullet_paras[n_pos:]
    if not pos_paras:
        return False
    pos_paras_final = _replace_list_items(pos_paras, pos_texts)

    last_pos_el = pos_paras_final[-1]._p
    if reusable_sep is not None:
        sep_para = reusable_sep
        last_pos_el.addnext(sep_para._p)
    else:
        new_el = copy.deepcopy(intro_para._p)
        last_pos_el.addnext(new_el)
        sep_para = docx.text.paragraph.Paragraph(new_el, pos_paras_final[-1]._parent)
    set_paragraph_text(sep_para, sep_sentence)

    if neg_paras:
        _replace_list_items(neg_paras, neg_texts)
    else:
        last_el = sep_para._p
        for text in neg_texts:
            new_el = copy.deepcopy(last_el)
            last_el.addnext(new_el)
            new_para = docx.text.paragraph.Paragraph(new_el, sep_para._parent)
            set_paragraph_text(new_para, text)
            last_el = new_el
    return True

def _rebuild_calendar_bridge(doc, items_texts, default_intro="The main variations are explained below:"):
    idx_gc, p_gc = find_paragraph(doc, r"^During the \d+ months ending \w+\d+,")
    idx_end, p_end = find_paragraph(doc, r"^Income Statement Analysis .* season\s*$")
    if p_gc is None or p_end is None or idx_end <= idx_gc:
        return False, "no ubiqué el tramo (intro o header de temporada no encontrados)"

    zone = [p for p in doc.paragraphs[idx_gc + 1:idx_end] if p.text.strip() != '']
    if not zone:
        return False, "tramo vacío entre el resultado calendario y el header de temporada"

    split_at = next((i for i, p in enumerate(zone) if 'US$' in p.text), None)
    if split_at is None:
        return False, "no encontré viñetas con cifras en el tramo"

    lead = zone[:split_at]
    old_bridge_paras = zone[split_at:]
    needs_intro = not (lead and lead[-1].text.strip().endswith(':'))

    template_el = old_bridge_paras[0]._p
    parent = old_bridge_paras[0]._parent
    insert_after_el = lead[-1]._p if lead else p_gc._p

    if needs_intro:
        intro_el = copy.deepcopy(template_el)
        insert_after_el.addnext(intro_el)
        set_paragraph_text(docx.text.paragraph.Paragraph(intro_el, parent), default_intro)
        insert_after_el = intro_el

    last_el = insert_after_el
    for text in items_texts:
        new_el = copy.deepcopy(template_el)
        last_el.addnext(new_el)
        set_paragraph_text(docx.text.paragraph.Paragraph(new_el, parent), text)
        last_el = new_el

    for p in old_bridge_paras:
        p._p.getparent().remove(p._p)
    return True, None

def apply_bridge_bullets(doc, wb, periods, causes, log):
    try:
        items_6m = build_bridge_items_6m(wb)
        ok, reason = _rebuild_calendar_bridge(doc, items_6m)
        log.append(f"[{'OK' if ok else 'AVISO'}] (EN) Bridge calendario ({len(items_6m)} factores)" + (f": {reason}" if not ok else ""))
    except Exception as ex:
        log.append(f"[ERROR] (EN) Bridge calendario: {ex}")

    try:
        bridge_12m_causes = causes.get('bridge_12m', {})
        idx, intro = find_paragraph(doc, r"^Profit \(loss\) attributable to the parent company shareholders was US\$")
        if intro is None:
            log.append("[AVISO] (EN) No encontré la intro del bridge de Ganancia Atribuible de temporada")
            return
        items = E.bridge_bullets(wb, '12m', threshold_abs_mus=3.0)
        pos = sorted([it for it in items if it[3] >= 0], key=lambda x: abs(x[3]), reverse=True)
        neg = sorted([it for it in items if it[3] < 0], key=lambda x: abs(x[3]), reverse=True)
        pos_texts = [E.p_ebitda_bullet_12m(wb)] + [E.bridge_factor_text(lbl, cur, prior, delta, '12m', cause=bridge_12m_causes.get(lbl)) for lbl, cur, prior, delta in pos]
        neg_texts = [E.bridge_factor_text(lbl, cur, prior, delta, '12m', cause=bridge_12m_causes.get(lbl)) for lbl, cur, prior, delta in neg]

        sep_sentence = "The previous positive effects were partially offset by the following effects:"
        ok = _replace_bridge_span_split(doc, idx + 1, pos_texts, neg_texts, sep_sentence, intro)
        if ok:
            log.append(f"[OK] (EN) Bridge de temporada GA ({len(pos_texts)} factores positivos, {len(neg_texts)} compensatorios)")
        else:
            log.append("[AVISO] (EN) No encontré las viñetas del bridge de Ganancia Atribuible de temporada")
    except Exception as ex:
        log.append(f"[ERROR] (EN) Bridge de temporada GA: {ex}")

# ---------- Tablas / imágenes ----------
# TODO (pendiente, ver docstring del módulo): por ahora NO se regeneran tablas en inglés. Se mantienen
# las imágenes que ya trae el Word base. Cuando existan specs de tabla en inglés (render_tables_v5.py
# con headers traducidos y formato en-US), reemplazar por una versión real de collect_and_render_tables
# análoga a la de assemble.py.

def apply_global_sweep(doc, old_periods, new_periods, log):
    if old_periods is None:
        log.append("[AVISO] (EN) No pude determinar el período del documento base (título no reconocido); se omitió el barrido final de fechas sueltas.")
        return
    n_runs = 0
    for p in doc.paragraphs:
        for r in p.runs:
            if not r.text:
                continue
            new_text = P.sweep_text_en(r.text, old_periods, new_periods)
            if new_text != r.text:
                r.text = new_text
                n_runs += 1
    log.append(f"[OK] (EN) Barrido final de fechas/etiquetas sueltas: {n_runs} fragmentos de texto corregidos")

# TABLE_ANCHORS_EN: (regex, tabla) EN EL ORDEN en que las 18 tablas-imagen aparecen en el informe,
# verificado contra el texto que antecede a cada imagen en el Word de referencia (Jun26,
# Earning_Report_June_2026). Reemplaza al viejo IMAGE_TO_TABLE (render_tables_v5.py), que asumía que
# "la imagen N (contada por nombre de archivo) es SIEMPRE la tabla N de esta lista" para CUALQUIER
# Word base que se suba — eso se rompe apenas el Word base es de un trimestre con una estructura de
# tablas distinta a la de referencia (ej. Flujo de Efectivo como imagen en vez de tabla nativa de
# Word en marzo 2026, o la tabla de Riesgos como tabla nativa en vez de imagen ese mismo trimestre):
# cada imagen de ahí en adelante terminaba recibiendo la tabla equivocada (ver aviso a Melissa
# oct-2026, ej. la tabla "Vencimientos de Préstamos" apareciendo donde debía ir Flujo de Efectivo).
# Ahora se empareja cada imagen con su tabla por el texto que la antecede (ver _build_image_table_map
# más abajo), tolerando que a un trimestre le falte o le sobre alguna imagen respecto al de
# referencia, en vez de asumir ciegamente la misma cantidad y orden.
TABLE_ANCHORS_EN = [
    (r'^The calculation of the EBITDA is detailed below', 'EBITDA'),
    (r'^Main Components of Income$', 'Ingresos'),
    (r'^Main Components of Income$', 'IngresosxSegmento'),
    (r'^Main Components of Costs and Expenses$', 'Costos'),
    (r'^Other Components of Income Statement$', 'OtrosComp'),
    (r'^Activity Indicators:?$', 'IndicadoresActividad'),
    (r"^The Company.s net financial debt (was reduced|decreased|increased)", 'DFN'),
    (r'^The evolution of the main financial indicators is as follows', 'Indicadores1'),
    (r'^(During this period, the company maintained its debt structure|The book value of the share)', 'Rentabilidad'),
    (r'^Main items of the Consolidated Statement of Financial Position$', 'Balance'),
    (r'^Below is a summary of the maturities of total financial liabilities', 'VctosPtmos'),
    (r'^The following table shows exposure in currencies different to the dollar', 'ExposicTC_tabla'),
    (r'^The potential net effect on financial assets and liabilities', 'EfectoTC'),
    (r'^As stated beforehand,.*risks identified by Hortifrut', 'RiesgosPorTipo'),
    (r'^As stated beforehand,.*risks identified by Hortifrut', 'RiesgosPorSeveridad'),
    (r'^Main critical risks faced as of', 'Riesgos'),
    (r'^The main insurances contracted as of', 'Seguros'),
    (r'^Below is the detail of the effect that a 10% reduction', 'FVfruta'),
]

def _has_image_el(p):
    return any(el.tag.endswith('}blip') for el in p._p.iter())

def _build_image_table_map(word_path, log):
    """Devuelve {índice de imagen (1-based, en el orden en que aparecen las imágenes leyendo el Word
    de arriba hacia abajo): tabla}, usando TABLE_ANCHORS_EN sobre el texto del párrafo más cercano
    ANTES de cada imagen (ver docstring de TABLE_ANCHORS_EN). Se abre `word_path` de nuevo, limpio,
    en vez de reusar el Word en inglés ya en memoria — para no depender de si el texto narrativo de
    alrededor ya fue sobreescrito por la traducción de este trimestre."""
    try:
        doc_fresh = docx.Document(word_path)
    except Exception as ex:
        log.append(f"[ERROR] (EN) No pude releer el Word base en inglés para ubicar sus tablas-imagen: {ex}")
        return {}

    paras = doc_fresh.paragraphs
    anchors_for_image = []
    for i, p in enumerate(paras):
        if _has_image_el(p):
            j = i - 1
            while j >= 0 and not paras[j].text.strip():
                j -= 1
            anchors_for_image.append(paras[j].text.strip() if j >= 0 else '')

    result = {}
    ai, ti = 0, 0
    n_img, n_tab = len(anchors_for_image), len(TABLE_ANCHORS_EN)
    unmatched_images = []
    skipped_tkeys = []  # tablas saltadas en el reacomodo (ver [AVISO] más abajo) — antes esto pasaba
    # en silencio: así fue como 'Riesgos' quedó sin su imagen sin que ningún aviso lo mencionara
    # (confirmado con los .docx reales de Melissa, oct-2026, punto 5 — ver _ensure_risks_image_en).
    while ai < n_img and ti < n_tab:
        text = anchors_for_image[ai]
        rx, tkey = TABLE_ANCHORS_EN[ti]
        if re.search(rx, text):
            result[ai + 1] = tkey
            ai += 1; ti += 1
            continue
        found_ahead = False
        for look in range(1, 6):
            if ti + look >= n_tab:
                break
            rx2, _tkey2 = TABLE_ANCHORS_EN[ti + look]
            if re.search(rx2, text):
                skipped_tkeys.extend(tk for _, tk in TABLE_ANCHORS_EN[ti:ti + look])
                ti += look
                found_ahead = True
                break
        if found_ahead:
            continue
        unmatched_images.append(ai + 1)
        ai += 1

    if unmatched_images:
        log.append(
            f"[AVISO] (EN) {len(unmatched_images)} imagen(es) del Word base en inglés (posición(es) "
            f"{unmatched_images} contando solo imágenes) no calzaron con ninguna tabla conocida por su "
            "texto de contexto — se dejaron tal cual, sin reemplazar (puede ser una tabla que ese "
            "trimestre viene como imagen en un formato distinto al de referencia)."
        )
    if skipped_tkeys:
        log.append(
            f"[AVISO] (EN) {len(skipped_tkeys)} tabla(s) se saltaron al reacomodar la coincidencia "
            f"imagen↔tabla (ver aviso de imágenes que no calzaron, arriba) y pueden haber quedado SIN "
            f"su imagen en el documento final — revísalas a mano: {skipped_tkeys}"
        )
    if ti < n_tab:
        faltantes = [tkey for _, tkey in TABLE_ANCHORS_EN[ti:]]
        log.append(
            f"[AVISO] (EN) {len(faltantes)} tabla(s) no encontraron su imagen correspondiente en el "
            f"Word base en inglés (probablemente ese trimestre no la tenía como imagen, ej. porque era "
            f"tabla nativa de Word): {faltantes}"
        )
    return result

def collect_and_render_tables_en(wb, word_path, log, doc=None):
    """Igual que collect_and_render_tables (assemble.py) pero renderiza cada tabla en inglés
    (render_table_en: mismo Excel en español, formato en-US + etiquetas traducidas vía TABLES_EN).
    doc: el Word en inglés ya abierto, si está disponible — mismo uso que en
    assemble.collect_and_render_tables (compensar el achicamiento de página de Word)."""
    rt_dir = os.path.dirname(__file__)
    sys.path.insert(0, rt_dir)
    import render_tables_v5 as RT

    usable_width_in = RT.usable_width_inches(doc) if doc is not None else None

    RT.TABLES['IngresosxSegmento']['sheet'] = 'IngresosxSegmento'
    if 'Indicadores1' not in wb.sheetnames and 'Ind. Financieros' in wb.sheetnames:
        RT.TABLES['Indicadores1']['sheet'] = 'Ind. Financieros'

    with zipfile.ZipFile(word_path, 'r') as z:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    actual = [n for n in re.findall(r'Target="media/([^"]+)"', rels) if re.match(r'image\d+', n)]
    actual = sorted(actual, key=lambda x: int(re.search(r'\d+', x).group()))

    img_idx_to_table = _build_image_table_map(word_path, log)

    img_dir = tempfile.mkdtemp(); image_map = {}; rendered = {}
    for img_idx, tkey in img_idx_to_table.items():
        if tkey not in RT.TABLES or tkey not in RT.TABLES_EN: continue
        spec = RT.TABLES[tkey]
        if spec['sheet'] not in wb.sheetnames: continue
        if img_idx > len(actual): continue
        a = actual[img_idx - 1]
        if tkey in rendered:
            image_map[a] = rendered[tkey]; continue
        out_png = os.path.join(img_dir, f'{tkey}.png')
        try:
            RT.render_table_en(wb, tkey, out_png, dpi=300, usable_width_in=usable_width_in)
            image_map[a] = out_png; rendered[tkey] = out_png
            log.append(f"[OK] (EN) Tabla renderizada: {tkey}")
        except Exception as ex:
            log.append(f"[ERROR] (EN) Tabla {tkey}: {ex}")
    return image_map, img_dir

# Llaves de `causes` que corresponden a una sola frase de causa (ver _call_single_template);
# 'bridge_12m' se maneja aparte porque es un diccionario anidado por label.
_SINGLE_CAUSE_KEYS = [
    'deterioro_6m_cur', 'deterioro_6m_prior', 'gastosadmin_6m',
    'costoventa_12m', 'razon_endeudamiento', 'cobertura_gf',
]

def _merge_causes_from_edited_es(causes, wb, word_es_edited_path, log):
    """Si se subió el Word en ESPAÑOL ya generado/editado este trimestre, extrae (por resaltado
    amarillo, ver assemble.collect_causes_by_highlight) sus causas de negocio y las usa en vez de las
    que traía el Word en inglés base (que son del trimestre ANTERIOR). Quedan marcadas con una tupla
    (texto_en_español, True) para que ar_engine_en._with_cause las deje en español, resaltadas, con el
    aviso de traducir — nunca se traducen automáticamente (ni con un servicio externo ni con un modelo
    local): es la persona quien las traduce. De paso, collect_causes_by_highlight deja en el log un
    [AVISO] por cada párrafo donde la persona haya editado a mano algo más allá de la causa resaltada
    (una cifra, una palabra), para que se revise manualmente que el inglés diga lo mismo ahí — ver
    docstring de esa función."""
    try:
        doc_es = docx.Document(word_es_edited_path)
    except Exception as ex:
        log.append(f"[AVISO] (EN) No pude abrir el Word en español editado para extraer causas: {ex}")
        return
    periods_es = ASM_ES.P.derive_periods(wb)
    causes_es = ASM_ES.collect_causes_by_highlight(doc_es, wb, periods_es, log)
    used_any = False
    for key in _SINGLE_CAUSE_KEYS:
        val = causes_es.get(key)
        if val:
            causes[key] = (val, True)
            used_any = True
    for label, val in causes_es.get('bridge_12m', {}).items():
        if val:
            causes.setdefault('bridge_12m', {})[label] = (val, True)
            used_any = True
    if used_any:
        log.append(
            "[OK] (EN) Se tomaron las causas de negocio del Word en español editado este trimestre "
            "(quedan en español, resaltadas, con aviso de traducir)"
        )
    else:
        log.append(
            "[AVISO] (EN) No encontré causas resaltadas en el Word en español que subiste; se "
            "mantienen las del Word en inglés base (si había) o solo el aviso genérico"
        )

# ---------- Flujo de Efectivo / Cash Flow (tabla nativa de Word, no una imagen) ----------
# Misma idea que assemble.update_flujo_caja_table (reutiliza su _find_fdec_header, _norm_label y
# set_cell_text, vía ASM_ES), pero traduciendo cada etiqueta del Excel (en español) a su
# equivalente en inglés antes de buscarla en la tabla del Word — porque esta tabla, en el Word en
# inglés, ya viene con las etiquetas traducidas (igual que la tabla 'Riesgos' en render_tables_v5.
# TABLES_EN). Las traducciones se verificaron contra Word en inglés reales de Hortifrut
# (ar_test_en/roundtrip_june.docx y las demás variantes de ese mismo trimestre); las pocas filas
# que no aparecieron en ningún Word real de referencia (ítems que el Excel trae en 0 casi siempre,
# así que rara vez se muestran) se tradujeron de forma directa y literal.
FDEC_LABELS_EN = {
    'cobros procedentes de las ventas de bienes y prestación de servicios':
        'Charges from the sale of goods and providing of services',
    'pago a proveedores por los suministros de bienes y servicios':
        'Payment to suppliers for supplying of goods and services',
    'pago por y a cuenta de empleados': 'Payment for and on behalf of employees',
    'intereses netos': 'Net interests',
    'impuestos': 'Taxes',
    'otros': 'Other',
    '(1) total flujo de efectivo por actividades de operación':
        '(1) Total Cash Flow for Operating Activities',
    'flujos de efectivo procedentes de la pérdida de control de subsidiarias u otros negocios':
        'Cash flows from loss of control of subsidiaries or other businesses',
    'flujos de efectivo utilizados para obtener el control de subsidiarias u otros negocios':
        'Cash flows used to obtain control of subsidiaries or other businesses',
    'flujos de efectivo utilizados en la compra de participaciones no controladoras':
        'Cash flows used in the purchase of non-controlling interests',
    'otros pagos para adquirir patrimonio o instrumentos de deuda de otras entidades':
        'Other payments to acquire equity or debt instruments of other entities',
    'venta y compra de propiedades, plantas y equipos':
        'Sale and purchase of property, plant and equipment',
    'compras de activos intangibles': 'Purchases of intangible assets',
    'compras de otros activos a largo plazo': 'Purchases of other long-term assets',
    'dividendos recibidos': 'Dividends received',
    '(2) total flujo de efectivo por actividades de inversión':
        '(2) Total Cash Flow for Investment Activities',
    'préstamos de entidades relacionadas': 'Loans from related entities',
    'ingresos por financiamiento': 'Income from financing',
    'pagos de préstamos': 'Payment of loans',
    'pagos de pasivos por arrendamiento': 'Payments of lease liabilities',
    'importes procedentes de la emisión de acciones': 'Proceeds from the issuance of shares',
    'dividendos pagados y otros flujos de inversión': 'Dividends paid and other investment flows',
    '(3) total flujo de efectivo por actividades de financiamiento':
        '(3) Total Cash Flow for Financing Activities',
    'incremento neto de efectivo y equivalentes de efectivo (1) + (2) + (3)':
        'Net Increase of Cash and Cash Equivalents (1) + (2) + (3)',
    'efecto por variación del tipo de cambio': 'Effect of exchange rate fluctuation',
    'efectivo y equivalente al efectivo al principio del periodo':
        'Cash and Cash Equivalent at the start of the period',
    'efectivo y equivalente al efectivo al final del período':
        'Cash and Cash Equivalent at the End of the Period',
}

_MESES_ES_EN_FDEC = {
    'enero': 'January', 'febrero': 'February', 'marzo': 'March', 'abril': 'April', 'mayo': 'May',
    'junio': 'June', 'julio': 'July', 'agosto': 'August', 'septiembre': 'September',
    'setiembre': 'September', 'octubre': 'October', 'noviembre': 'November', 'diciembre': 'December',
}

def _translate_fdec_header_en(s):
    """Traduce el rango de fechas del encabezado ('Enero-Junio 2026' -> 'January-June 2026'); si el
    texto no calza con el patrón esperado, se deja tal cual en vez de inventar algo."""
    if not s:
        return s
    m = re.match(r'^\s*([A-Za-zñÑáéíóúÁÉÍÓÚ]+)\s*-\s*([A-Za-zñÑáéíóúÁÉÍÓÚ]+)\s+(\d{4})\s*$', s)
    if not m:
        return s
    e1 = _MESES_ES_EN_FDEC.get(m.group(1).lower())
    e2 = _MESES_ES_EN_FDEC.get(m.group(2).lower())
    if not e1 or not e2:
        return s
    return f"{e1}-{e2} {m.group(3)}"

def update_flujo_caja_table_en(doc, wb, log):
    """Igual que assemble.update_flujo_caja_table, pero para el Word en inglés: cada etiqueta de la
    hoja 'FdeC' se traduce con FDEC_LABELS_EN antes de buscarla en la tabla (que ya viene con las
    etiquetas en inglés), y las cifras se formatean en-US (sin decimales, igual que el resto de las
    cifras en inglés del informe — ver fmt_en en render_tables_v5.py). La columna de Comments (si
    existe) nunca se toca."""
    rt_dir = os.path.dirname(__file__)
    sys.path.insert(0, rt_dir)
    import render_tables_v5 as RT

    if 'FdeC' not in wb.sheetnames:
        log.append("[AVISO] (EN) No encontré la hoja 'FdeC' en el Excel: no se actualizó el Cash Flow.")
        return False
    ws = wb['FdeC']
    hr, hc = ASM_ES._find_fdec_header(ws)
    if hr is None:
        log.append("[AVISO] (EN) No encontré el encabezado del Flujo de Efectivo en la hoja 'FdeC': no se actualizó.")
        return False
    cur_col, prior_col = hc + 1, hc + 2
    header_cur_label = _translate_fdec_header_en(str(ws.cell(hr, cur_col).value or '').strip())
    header_prior_label = _translate_fdec_header_en(str(ws.cell(hr, prior_col).value or '').strip())

    data_en = {}
    unmapped_es = []
    r = hr + 1
    while True:
        label = ws.cell(r, hc).value
        if label is None or label == '':
            break
        norm_es = ASM_ES._norm_label(label)
        en_label = FDEC_LABELS_EN.get(norm_es)
        if en_label:
            data_en[ASM_ES._norm_label(en_label)] = (ws.cell(r, cur_col).value, ws.cell(r, prior_col).value)
        else:
            unmapped_es.append(str(label).strip()[:60])
        r += 1

    target_table = None
    for t in doc.tables:
        try:
            first_cell_text = t.rows[0].cells[0].text
        except Exception:
            continue
        if re.search(r'[Mm]ain sources and uses of funds', first_cell_text):
            target_table = t
            break
    if target_table is None:
        log.append("[AVISO] (EN) No encontré la tabla de Cash Flow en el Word en inglés: no se actualizó.")
        return False

    n_updated = 0
    missing = []
    for ri, row in enumerate(target_table.rows):
        cells = row.cells
        if len(cells) < 3:
            continue
        if ri == 0:
            if header_cur_label: ASM_ES.set_cell_text(cells[1], header_cur_label)
            if header_prior_label: ASM_ES.set_cell_text(cells[2], header_prior_label)
            continue
        label = ASM_ES._norm_label(cells[0].text)
        if label in data_en:
            cur_val, prior_val = data_en[label]
            ASM_ES.set_cell_text(cells[1], RT.fmt_en(cur_val, 'num'))
            ASM_ES.set_cell_text(cells[2], RT.fmt_en(prior_val, 'num'))
            n_updated += 1
        else:
            missing.append(cells[0].text.strip()[:60])

    msg = f"[OK] (EN) Cash Flow actualizado: {n_updated} fila(s) con cifras nuevas del Excel ('FdeC')."
    if missing:
        msg += (f" {len(missing)} fila(s) del Word no encontraron su etiqueta traducida y se dejaron "
                f"tal cual (revisa si cambió la redacción): {'; '.join(missing)}")
    log.append(msg)
    if unmapped_es:
        log.append(
            "[AVISO] (EN) Filas del Excel 'FdeC' sin traducción registrada en FDEC_LABELS_EN (no se "
            f"usaron porque no tengo cómo buscarlas en el Word en inglés): {'; '.join(unmapped_es)}"
        )
    log.append("[INFO] (EN) Cash Flow: la columna de Comments no se toca — complétala tú misma cada trimestre.")
    return True

# ---------- Clonar la tabla de Cash Flow desde el español (si el Word base en inglés no trae tabla
# nativa, solo una imagen vieja) ----------
# Feedback Melissa (Oct26), con los .docx reales en mano: "La tabla del Flujo de Caja en inglés, no
# está igual a la de español. Ni en estructura, formato ni en datos." Root-cause confirmado contra el
# Word en inglés REAL de Hortifrut (Analisis_Razonado_Marzo_2026_16_06_26_English.docx, el documento
# oficial de la Compañía, no uno generado por esta app): ahí, igual que en el Word que subió Melissa,
# la sección "MAIN COMPONENTS OF NET OPERATING INVESTMENT AND FINANCING CASH FLOWS" va seguida de una
# imagen (.emf) pegada a mano, nunca de una tabla nativa de Word editable como en español — por eso
# update_flujo_caja_table_en (arriba) jamás encuentra nada que actualizar ahí: no es un bug de esta
# app, es que esa tabla, en inglés, nunca tuvo esa forma en los Word base de Hortifrut.
# Mientras Melissa no reemplace esa imagen por una tabla nativa en su plantilla base en inglés (lo
# ideal a mediano plazo, para que vuelva a quedar 100% editable igual que en español), esta función
# CLONA la tabla nativa YA ACTUALIZADA del Word en español definitivo de este trimestre
# (word_es_edited_path) dentro del Word en inglés, en el lugar de esa imagen vieja. Es seguro clonar el
# <w:tbl> completo tal cual (deepcopy, cross-documento) porque ambos Word comparten la misma plantilla
# corporativa (mismos w:styleId, verificado contra los .docx reales de Melissa) — así se conserva
# exactamente la misma estructura y formato (anchos de columna, bordes, negrita de la fila Total) que
# la tabla en español, no solo una aproximación. Luego de clonar, solo se traduce el TEXTO: las
# etiquetas de cada fila (vía FDEC_LABELS_EN — después de esto, generate_en vuelve a llamar a
# update_flujo_caja_table_en, que ahora SÍ encuentra la tabla recién clonada y le pone las cifras del
# Excel ya en formato en-US) y la columna de Comentarios (texto libre que Melissa escribe a mano cada
# trimestre: se traduce con DeepL si está configurado, igual que el resto del documento; si no, se dej
# a en español con el mismo aviso que usa el resto del texto narrativo).
def _find_cash_flow_image_anchor(doc_en):
    """Ubica, en el Word en inglés, el párrafo de encabezado estático de esta sección (no cambia de
    trimestre a trimestre, no está en HEADER_TEMPLATES_EN) y, si lo encuentra, el párrafo con la
    imagen vieja que suele venir justo después (separados por alguna línea en blanco, igual que en los
    .docx reales inspeccionados). Devuelve (anchor_para, image_para_o_None)."""
    idx, anchor_para = find_paragraph(
        doc_en, r"MAIN (COMPONENTS OF NET OPERATING|SOURCES AND USES OF FUNDS)")
    if anchor_para is None:
        return None, None
    paras = doc_en.paragraphs
    j = idx + 1
    while j < len(paras) and not paras[j].text.strip():
        if _has_image_el(paras[j]):
            return anchor_para, paras[j]
        j += 1
    return anchor_para, None

def _clone_flujo_caja_table_from_es(doc_en, word_es_edited_path, log):
    try:
        doc_es = docx.Document(word_es_edited_path)
    except Exception as ex:
        log.append(f"[AVISO] (EN) No pude abrir el Word en español para clonar la tabla de Cash Flow: {ex}")
        return False

    es_table = None
    for t in doc_es.tables:
        try:
            first_cell_text = t.rows[0].cells[0].text
        except Exception:
            continue
        if re.search(ASM_ES.FLUJO_CAJA_HEADER_ANCHOR, first_cell_text):
            es_table = t
            break
    if es_table is None:
        log.append("[AVISO] (EN) No encontré la tabla de Flujo de Efectivo en el Word en español: no "
                    "pude clonarla al inglés.")
        return False

    anchor_para, image_para = _find_cash_flow_image_anchor(doc_en)
    if anchor_para is None:
        log.append("[AVISO] (EN) No encontré el encabezado de la sección de Cash Flow en el Word en "
                    "inglés (texto esperado: 'MAIN COMPONENTS OF NET OPERATING...'): no pude insertar "
                    "la tabla clonada desde el español.")
        return False

    insert_after_el = image_para._p if image_para is not None else anchor_para._p
    if image_para is not None:
        old_p = image_para._p
        insert_after_el = old_p
    new_tbl = copy.deepcopy(es_table._tbl)
    insert_after_el.addnext(new_tbl)
    if image_para is not None:
        old_p.getparent().remove(old_p)

    new_table_obj = docx.table.Table(new_tbl, doc_en)

    n_translated = 0
    n_untranslated = 0
    for ri, row in enumerate(new_table_obj.rows):
        cells = row.cells
        if not cells:
            continue
        if ri == 0:
            # Fila de encabezado: el texto de esta celda es lo que update_flujo_caja_table_en busca
            # (regex 'Main sources and uses of funds') para reconocer la tabla — sin traducirla acá,
            # la llamada de vuelta a esa función (ver generate_en) no la encontraría nunca.
            # Texto exacto verificado contra el Word en inglés real de Hortifrut de Jun26
            # (Earning_Report_June_2026_27_08_26.docx), donde esta SÍ es una tabla nativa.
            ASM_ES.set_cell_text(cells[0], 'Main sources and uses of funds in the period (US$ million)')
            continue
        norm = ASM_ES._norm_label(cells[0].text)
        en_label = FDEC_LABELS_EN.get(norm)
        if en_label:
            ASM_ES.set_cell_text(cells[0], en_label)
            n_translated += 1
        else:
            n_untranslated += 1
            log.append(
                "[AVISO] (EN) Fila del Flujo de Efectivo sin traducción registrada en FDEC_LABELS_EN, "
                f"se dejó en español: {cells[0].text.strip()[:60]!r}"
            )

    deepl_on = DT.is_configured()
    for ri, row in enumerate(new_table_obj.rows):
        cells = row.cells
        if len(cells) < 4:
            continue
        if ri == 0:
            ASM_ES.set_cell_text(cells[3], 'Comments')
            continue
        es_comment = cells[3].text.strip()
        if not es_comment:
            continue
        translated = DT.translate_es_en(convert_es_numbers_to_en(es_comment)) if deepl_on else None
        if translated:
            ASM_ES.set_cell_text(cells[3], translated)
        elif deepl_on:
            ASM_ES.set_cell_text(cells[3], f"{NO_DEEPL_FAILED_PREFIX_EN}{es_comment}")
        else:
            ASM_ES.set_cell_text(cells[3], f"{NO_DEEPL_PREFIX_EN}{es_comment}")

    log.append(
        "[OK] (EN) Tabla de Cash Flow clonada desde el Word en español (reemplaza la imagen vieja "
        f"desactualizada) — {n_translated} fila(s) traducida(s), estructura y formato ahora calzan "
        "exactamente con la tabla en español."
    )
    return True

# Algunas plantillas Word en inglés (vistas en la de marzo 2026, pero no en la de junio) traen, debajo
# del texto "Main critical risks faced as of ... are the following:", una tabla NATIVA de Word vieja
# (encabezados "Type of Risk / Name of the Risk / Description of the Risk / Corporate Controls
# Implemented") con contenido genérico/desactualizado (ej. "Water supply, availability and quality",
# "Cybersecurity attacks") que NO existe en la versión en español equivalente — ahí esa misma sección
# solo trae la imagen con el detalle real de riesgos (tabla 'Riesgos' de render_tables_v5, ya traducida
# más abajo en TABLES_EN). Melissa reportó que el tema de riesgos críticos "queda repetido" en inglés
# por esto: la plantilla muestra la tabla vieja Y, en otra parte del documento, nuestra imagen con el
# detalle real — dos cosas hablando de lo mismo. Como esta tabla vieja es contenido estático de la
# plantilla (no algo que nuestro código inserte), no hay celdas que "actualizar" como en Flujo de
# Efectivo: la solución es borrarla directamente para igualar la estructura a la versión en español.
# El encabezado buscado es muy específico (no puede confundirse con la tabla de Flujo de Efectivo,
# cuyo encabezado es "Main sources and uses of funds...").
_STALE_RISK_TABLE_HEADER_RE = re.compile(r'type of risk', re.IGNORECASE)

def remove_stale_risk_table_en(doc, log):
    removed = 0
    for t in list(doc.tables):
        try:
            header_row_text = ' | '.join(c.text for c in t.rows[0].cells)
        except Exception:
            continue
        if _STALE_RISK_TABLE_HEADER_RE.search(header_row_text) and 'corporate controls' in header_row_text.lower():
            t._element.getparent().remove(t._element)
            removed += 1
    if removed:
        log.append(
            f"[OK] (EN) Se encontró y eliminó {removed} tabla(s) nativa(s) de riesgos obsoleta(s) en la "
            "plantilla ('Type of Risk / Name of the Risk / ...') que no existen en la versión en español: "
            "el detalle real de riesgos ya lo muestra la imagen traducida más abajo en el documento."
        )

# ---------- Red de seguridad: insertar la tabla de riesgos críticos si quedó ausente ----------
# Feedback Melissa (Oct26), punto 5: "En inglés faltó la tabla de riesgos criticos." Root-cause
# confirmado con los .docx reales: en el Word en inglés que ella sube, después de "Main critical
# risks faced as of ... are the following:" NO hay nada — ni la tabla nativa vieja (la que
# remove_stale_risk_table_en elimina arriba, asumiendo que la imagen real la reemplaza) ni ninguna
# imagen. collect_and_render_tables_en (TABLE_ANCHORS_EN) solo SUSTITUYE imágenes que ya existen en el
# Word base: si no hay ninguna imagen ahí para sustituir, la tabla 'Riesgos' queda completamente
# ausente y nada lo avisa (el único aviso que había, sobre imágenes sin calzar, habla de imágenes
# SOBRANTES, no de tablas FALTANTES en medio de la lista — ver _build_image_table_map, que ahora
# además deja constancia en el log de cualquier tabla saltada en el reacomodo). Esta función es la red
# de seguridad final: si de verdad no quedó ninguna imagen ahí, renderiza la tabla 'Riesgos' (mismo
# Excel/formato que las otras 17 tablas-imagen en inglés, ver render_tables_v5.TABLES_EN) y la inserta
# directamente, en vez de depender de que ya exista una imagen para sustituir.
    # El ancla busca tanto el texto correcto en inglés como su versión en español-con-prefijo: esta
    # oración trae la fecha del trimestre (como "Al 30 de junio de 2026" en HEADER_TEMPLATES), así
    # que SIEMPRE difiere de la del trimestre anterior y translate_narrative_from_definitive_es
    # siempre la reescribe — con DeepL apagado (caso por defecto), el resultado NO es el texto en
    # inglés sino el español con el prefijo "SPANISH — DeepL not configured...", que no calzaría con
    # un ancla que solo buscara el inglés (confirmado con los .docx reales de Melissa: así se perdió
    # el ancla la primera vez que se probó este fix).
_RISKS_HEADER_ANCHOR_RE = r"^(Main critical risks faced as of|(SPANISH.{0,80}:\s*)?Los principales riesgos cr[ií]ticos que se enfrentan al)"

def _ensure_risks_image_en(doc, wb, log):
    anchor_idx, anchor_para = find_paragraph(doc, _RISKS_HEADER_ANCHOR_RE)
    if anchor_para is None:
        log.append(
            "[AVISO] (EN) No encontré el encabezado 'Main critical risks faced as of...' en el Word "
            "en inglés: no pude verificar si la tabla de riesgos críticos quedó ausente."
        )
        return
    paras = doc.paragraphs
    j = anchor_idx + 1
    last_el = anchor_para._p
    while j < len(paras) and not paras[j].text.strip():
        if _has_image_el(paras[j]):
            return  # ya hay una imagen ahí; collect_and_render_tables_en se encarga de actualizarla
        last_el = paras[j]._p
        j += 1

    rt_dir = os.path.dirname(__file__)
    sys.path.insert(0, rt_dir)
    import render_tables_v5 as RT
    spec = RT.TABLES.get('Riesgos')
    if not spec or spec['sheet'] not in wb.sheetnames or 'Riesgos' not in RT.TABLES_EN:
        log.append(
            "[AVISO] (EN) La tabla de riesgos críticos no aparece en el Word en inglés y tampoco pude "
            "generarla (falta la hoja 'Riesgos' en el Excel o su configuración en render_tables_v5): "
            "revisa esa sección a mano."
        )
        return

    img_dir = tempfile.mkdtemp()
    out_png = os.path.join(img_dir, 'Riesgos_faltante.png')
    try:
        usable_w_emu = doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin
        usable_width_in = usable_w_emu / ASM_ES.EMU_PER_INCH
        RT.render_table_en(wb, 'Riesgos', out_png, dpi=300, usable_width_in=usable_width_in)
        from PIL import Image as PILImage
        with PILImage.open(out_png) as im:
            w_px, h_px = im.size
        new_w_emu = min(usable_w_emu, int(w_px / 300 * ASM_ES.EMU_PER_INCH))
        new_h_emu = int(new_w_emu * h_px / w_px)
        new_p = docx.oxml.parser.OxmlElement('w:p')
        last_el.addnext(new_p)
        new_para = docx.text.paragraph.Paragraph(new_p, anchor_para._parent)
        new_para.add_run().add_picture(out_png, width=docx.shared.Emu(new_w_emu), height=docx.shared.Emu(new_h_emu))
        log.append(
            "[OK] (EN) La tabla de riesgos críticos no tenía ninguna imagen que sustituir en el Word en "
            "inglés (quedaba completamente ausente): se generó y se insertó de nuevo, con el mismo "
            "detalle/formato que la versión en español."
        )
    except Exception as ex:
        log.append(f"[ERROR] (EN) No pude generar/insertar la tabla de riesgos críticos faltante: {ex}")
    finally:
        shutil.rmtree(img_dir, ignore_errors=True)

# ---------- Traducción general de TODO el texto narrativo, vía DeepL ----------
# Antes, esto solo cubría el texto libre (párrafos que no calzaban con ninguna plantilla de
# SINGLE_TEMPLATES_EN/HEADER_TEMPLATES_EN); las oraciones con cifras (EBITDA, costos, razón de
# endeudamiento, etc.) y las viñetas del puente de EBITDA/Ganancia Atribuible se generaban
# recalculando una plantilla en inglés fija (ar_engine_en.py) con cifras nuevas del Excel. Melissa
# reportó (oct-2026) algo más grave que una diferencia de redacción: las CIFRAS de esas viñetas en
# inglés no coincidían con las del español — ej. "Higher EBITDA... US$107.81 million" en inglés vs.
# "EBITDA... US$197,89 millones" en español, el mismo concepto, el mismo trimestre. La causa: el
# motor de cifras en inglés (ar_engine_en.py) calcula esos valores de forma INDEPENDIENTE del motor
# en español (ar_engine.py) — ambos leen el mismo Excel, pero son dos implementaciones separadas que
# pueden (y en este caso lo hicieron) calzar con el ancla equivocada o computar un período distinto,
# sin que nada lo detecte. Mientras existan dos motores de cifras en paralelo, este tipo de error
# puede volver a aparecer en cualquier otra oración con cifras.
#
# Por pedido explícito de Melissa: "necesito que el texto en inglés sea exactamente igual al texto
# en español... lo mejor sería la traducción línea a línea... usar el inglés del trimestre anterior
# como referencia para mantener la misma traducción a ciertos conceptos" — aplicado a TODO el
# documento (no solo a las secciones donde ya vio el problema). Esto elimina de raíz el riesgo de
# arriba: ya no existe un segundo motor de cifras en inglés para el texto narrativo — cada oración y
# viñeta en inglés sale de traducir, literalmente, la oración/viñeta ya definitiva en español (la
# misma que Melissa revisó y aprobó, con las cifras y causas de negocio que ella misma escribió), así
# que no puede haber un descalce de cifras: si el español dice 197,89, el inglés dice 197.89 — nunca
# otro número. Las funciones de SINGLE_TEMPLATES_EN/ar_engine_en.py y el merge de causas de negocio
# (collect_original_causes, _merge_causes_from_edited_es, apply_single_templates, apply_bridge_bullets
# más arriba) quedaron SIN USAR por este motivo — se dejan definidas por si hiciera falta volver atrás,
# pero generate_en() ya no las llama.
#
# Lo que SÍ sigue siendo determinístico (no pasa por acá): el título con la fecha y los encabezados
# de sección (HEADER_TEMPLATES_EN, ver apply_header_templates) — son puro cálculo de fechas/períodos
# a partir del Excel, no cifras financieras ni redacción libre, así que no tienen este riesgo y es
# mejor mantenerlos 100% determinísticos en vez de confiar en que DeepL arme bien el rango de fechas.
# Tampoco pasan por acá la tabla de Flujo de Efectivo (es una tabla nativa de Word, no texto
# narrativo, ver update_flujo_caja_table_en) ni las 18 tablas-imagen (render_tables_v5.py, con su
# propio mecanismo de traducción, ya revisado por separado).
#
# Mecánica: se recorre cada documento (el Word en español DEFINITIVO de este trimestre, el Word en
# español del trimestre ANTERIOR si se subió, y el Word en inglés recién abierto, que al llegar acá
# todavía trae el texto en inglés del trimestre ANTERIOR tal cual el Word base que se subió) y se
# agrupa en "segmentos": un párrafo normal es un segmento de un elemento; una racha de viñetas (estilo
# 'List' de Word) consecutivas es UN segmento con varios elementos (para poder agregar/quitar viñetas
# si la cantidad cambia, ver _replace_list_items). Los párrafos vacíos, los que solo traen una imagen,
# y los que calzan con un encabezado de HEADER_TEMPLATES cortan un segmento (no se tocan acá).
#
# Para no perder consistencia de redacción (el pedido de Melissa de "usar el inglés anterior como
# referencia"): se arma un diccionario español->inglés comparando, posición por posición, el texto
# del Word en español ANTERIOR contra el texto que YA trae el Word en inglés (que es, segmento por
# segmento, la traducción aprobada de ESE mismo español anterior, generada por esta misma lógica el
# trimestre pasado). Para cada oración/viñeta de ESTE trimestre: si su texto (normalizado) aparece tal
# cual en ese diccionario, se reutiliza esa traducción aprobada SIN volver a llamar a DeepL (cero
# riesgo de que DeepL traduzca el mismo texto con una palabra distinta a como quedó la vez pasada); si
# no aparece (texto nuevo o reescrito), se traduce recién vía DeepL. Esto funciona aunque una oración
# se haya corrido de posición dentro de un bloque (viñetas reordenadas, una agregada o borrada en
# medio) porque la búsqueda es por CONTENIDO, no por posición — mismo principio que la corrección
# anterior al descalce de la sección de Riesgos.

def _has_image(p):
    return any(el.tag.endswith('}blip') for el in p._p.iter())

def _is_header_para(p, header_anchors):
    t = p.text.strip()
    return any(rx.search(t) for rx in header_anchors)

def _segment_narrative(doc, header_anchors):
    """Agrupa los párrafos de `doc` en segmentos: un párrafo normal = un segmento de 1 elemento; una
    racha de viñetas consecutivas (ver _is_list_para) = un segmento de N elementos. Los párrafos
    vacíos, los que solo traen una imagen (_has_image) y los que calzan con un encabezado de
    HEADER_TEMPLATES (fecha/título de sección, ver docstring de arriba) cortan el segmento actual sin
    pasar a formar parte de ninguno — esos se actualizan aparte, de forma determinística."""
    segments = []
    cur_list = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t or _has_image(p) or _is_header_para(p, header_anchors):
            if cur_list:
                segments.append(cur_list); cur_list = []
            continue
        if _is_list_para(p):
            cur_list.append(p)
            continue
        if cur_list:
            segments.append(cur_list); cur_list = []
        segments.append([p])
    if cur_list:
        segments.append(cur_list)
    return segments

def _segment_signature(seg):
    # Solo describe la FORMA del segmento (racha de viñetas vs párrafo normal), nunca su contenido —
    # el contenido siempre sale del español (old_lookup/template_lookup/DeepL, ver más abajo). Se usa
    # solo para decidir cuántos segmentos de cada tipo hacen falta (ver _align_segments).
    #
    # Nota: se probó además distinguir párrafos normales por un bucket de largo de texto (para
    # desambiguar dos párrafos normales consecutivos sin ninguna viñeta cerca que sirva de ancla, ej.
    # el título del documento) pero empeoró MUCHO el resultado real: el español es sistemáticamente
    # ~15-25% más largo que su traducción al inglés, así que casi ningún párrafo quedaba en el mismo
    # bucket que su propia traducción, aunque fuera exactamente la misma oración — se revirtió.
    return 'L' if _is_list_para(seg[0]) else 'N'

def _neutralize_cloned_paragraph_format(paragraph):
    """Despoja un párrafo recién clonado (ver _clone_paragraph_after/_clone_paragraph_before) de
    negrita/subrayado directos en sus runs, dejando solo fuente/tamaño. Sin esto, un párrafo nuevo
    clonado de un VECINO que por pura casualidad era un título/subtítulo en negrita+subrayado (ej.
    clonar justo después de "Principales partidas del Estado de Situación Financiera Consolidado")
    heredaba esa negrita/subrayado como si el contenido NUEVO (una oración narrativa común, nunca un
    título) también debiera llevarla — confirmado: así quedaba una oración de bridge común
    ("Reducción en activos no corrientes...") en negrita y subrayada por error, tanto en los .docx
    reales de Melissa como en el Word en inglés oficial de referencia (Marzo 2026). Un párrafo recién
    insertado no es un título heredado de nadie: debe partir neutro, y set_paragraph_text se encarga
    de ponerle `{{...}}` si de verdad corresponde una frase líder en negrita."""
    for r in paragraph.runs:
        r.bold = False
        r.underline = False

def _clone_paragraph_after(anchor_para, n=1):
    """Clona `anchor_para` (su mismo estilo/formato) `n` veces, insertando los clones inmediatamente
    después de `anchor_para`. Devuelve la lista de los nuevos objetos Paragraph (todavía con el texto
    de `anchor_para`; el texto real se setea aparte vía set_paragraph_text). Mismo mecanismo que ya
    usa _replace_list_items para agregar viñetas de más, generalizado acá a cualquier párrafo."""
    new_paras = []
    last_el = anchor_para._p
    for _ in range(n):
        new_el = copy.deepcopy(last_el)
        last_el.addnext(new_el)
        new_para = docx.text.paragraph.Paragraph(new_el, anchor_para._parent)
        _neutralize_cloned_paragraph_format(new_para)
        new_paras.append(new_para)
        last_el = new_el
    return new_paras

def _clone_paragraph_before(anchor_para, n=1):
    """Como _clone_paragraph_after, pero insertando los clones ANTES de `anchor_para` (en orden). Se
    usa cuando el segmento nuevo aparece antes de cualquier párrafo en inglés todavía visto (ej. el
    documento empieza con un párrafo nuevo) y por eso no hay un párrafo previo del cual clonar hacia
    adelante — se clona el primer párrafo del documento en inglés y se inserta antes de él, para no
    perder el contenido ni el orden."""
    new_paras = []
    anchor_el = anchor_para._p
    for _ in range(n):
        new_el = copy.deepcopy(anchor_el)
        anchor_el.addprevious(new_el)
        new_para = docx.text.paragraph.Paragraph(new_el, anchor_para._parent)
        _neutralize_cloned_paragraph_format(new_para)
        new_paras.append(new_para)
    return new_paras

def _remove_segment(seg):
    for p in seg:
        p._p.getparent().remove(p._p)

def _align_segments(segs_es, segs_en, log):
    """Empareja los segmentos de `segs_es` (Word en español DEFINITIVO de este trimestre) con los de
    `segs_en` (Word en inglés base) para decidir, segmento por segmento, en qué párrafo(s) de doc_en
    escribir cada oración/viñeta traducida.

    ANTES (versión anterior): se emparejaba por POSICIÓN directa (segs_es[i] con segs_en[i], hasta
    min(len(segs_es), len(segs_en))). Eso es MUY frágil: un solo párrafo de más o de menos en
    CUALQUIER parte de alguno de los dos documentos (típico en Word mantenidos a mano trimestre a
    trimestre — ej. un título partido en 2 líneas un trimestre y en 1 línea otro) desalinea TODO lo
    que viene después para el resto del documento, mezclando el texto de una oración con el párrafo
    de OTRA — así se generaron los "datos rarísimos" que reportó Melissa en oct-2026 (párrafos con
    cifras de trimestres distintos apareciendo duplicados, uno al lado del otro).

    AHORA: se comparan las FIRMAS estructurales de ambos (_segment_signature: 'L' viñetas / 'N'
    párrafo normal) con difflib, igual que ya se hace para detectar qué cambió en una lista de
    oraciones — así un párrafo de más o de menos en cualquier parte solo desalinea ESE tramo puntual
    (se detecta como 'insert'/'delete'/'replace' local), nunca el resto del documento. Cuando a
    `segs_es` le sobra un segmento sin contraparte en inglés (párrafo nuevo este trimestre), se CLONA
    el último párrafo en inglés visto (mismo mecanismo de clonado que ya usa _replace_list_items) para
    poder escribir ahí con el formato correcto. Cuando a `segs_en` le sobra un segmento sin
    contraparte en español (quedó obsoleto / ya no existe este trimestre), se BORRA del documento en
    vez de dejarlo tal cual (que es justo lo que producía el dato viejo duplicado).

    Cada tupla devuelta en `pairs` trae un tercer elemento `trusted` (bool): True solo cuando el
    segmento en inglés viene de un tramo 'equal' del diff (misma firma estructural Y misma posición
    relativa en la secuencia — la señal más fuerte de que de verdad es la contraparte de ese segmento
    español). False para cualquier segmento en inglés reutilizado dentro de un tramo 'replace' (la
    cantidad de segmentos no calzaba ahí, así que el emparejamiento es solo por POSICIÓN dentro del
    tramo, no por contenido) o para un segmento recién clonado. Esto importa para el llamador: un
    segmento 'replace' puede terminar reutilizando un párrafo en inglés que NO tiene relación alguna de
    contenido con el español que le toca — confirmado con un caso real donde la viñeta reutilizada para
    "Reducción en activos no corrientes..." resultó ser, por puro accidente de alineación, el
    encabezado "RISK FACTOR ANALYSIS " (negrita+subrayado, de una sección totalmente distinta) — así
    que el llamador solo debe confiar en conservar negrita/subrayado existente cuando trusted=True."""
    sig_es = [_segment_signature(s) for s in segs_es]
    sig_en = [_segment_signature(s) for s in segs_en]
    sm = difflib.SequenceMatcher(None, sig_en, sig_es, autojunk=False)

    first_en_para = segs_en[0][0] if segs_en else None

    pairs = []
    n_equal = 0
    n_inserted = 0
    n_deleted = 0
    last_en_para = None
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                pairs.append((segs_es[j1 + k], segs_en[i1 + k], True))
                last_en_para = segs_en[i1 + k][-1]
                n_equal += 1
            continue
        n_en_here = i2 - i1
        n_es_here = j2 - j1
        n_both = min(n_en_here, n_es_here)
        for k in range(n_both):
            pairs.append((segs_es[j1 + k], segs_en[i1 + k], False))
            last_en_para = segs_en[i1 + k][-1]
        if n_en_here > n_both:
            # Sobran segmentos en inglés sin contraparte en español este trimestre: se borran (en vez
            # de dejarlos tal cual, con datos de un trimestre que ya no corresponde).
            for k in range(i1 + n_both, i2):
                _remove_segment(segs_en[k])
                n_deleted += 1
        elif n_es_here > n_both:
            # Sobran segmentos en español sin contraparte en inglés (párrafo/bloque nuevo este
            # trimestre): se clona el último párrafo en inglés visto para mantener el formato.
            for k in range(j1 + n_both, j2):
                if last_en_para is not None:
                    new_seg = _clone_paragraph_after(last_en_para, 1)
                    last_en_para = new_seg[-1]
                elif first_en_para is not None:
                    # Todavía no se vio ningún párrafo en inglés (el segmento nuevo está antes de
                    # cualquier coincidencia, ej. el documento empieza con contenido nuevo) — se clona
                    # el primer párrafo del documento en inglés y se inserta antes de él.
                    new_seg = _clone_paragraph_before(first_en_para, 1)
                else:
                    new_seg = []
                pairs.append((segs_es[k], new_seg, False))
                n_inserted += 1

    if n_inserted or n_deleted:
        log.append(
            f"[AVISO] (EN) La estructura de párrafos del Word en español definitivo y la del Word en "
            f"inglés base no calzaban del todo (ej. un título partido distinto, un párrafo fusionado o "
            f"nuevo): se ajustó automáticamente — se insertaron {n_inserted} párrafo(s) nuevo(s) "
            f"clonando el formato del párrafo en inglés más cercano, y se eliminaron {n_deleted} "
            "párrafo(s) en inglés sin contraparte en español este trimestre (contenido obsoleto de "
            "otro trimestre) — en vez de dejar todo desalineado como antes."
        )
    return pairs, n_equal

def _align_segments_readonly(segs_a, segs_b):
    """Como _align_segments, pero de SOLO LECTURA (no clona ni borra ningún párrafo): devuelve pares
    (segmento_a, segmento_b) para los tramos que calzaron en estructura, ignorando los que no tienen
    contraparte (sin insertar ni borrar nada — no corresponde tocar ningún documento acá). Se usa para
    construir old_lookup/template_lookup a partir del Word en español del trimestre ANTERIOR y el Word
    en inglés base, que son solo diccionarios de consulta de texto, nunca documentos a modificar."""
    sig_a = [_segment_signature(s) for s in segs_a]
    sig_b = [_segment_signature(s) for s in segs_b]
    sm = difflib.SequenceMatcher(None, sig_a, sig_b, autojunk=False)
    pairs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                pairs.append((segs_a[i1 + k], segs_b[j1 + k]))
        else:
            n_both = min(i2 - i1, j2 - j1)
            for k in range(n_both):
                pairs.append((segs_a[i1 + k], segs_b[j1 + k]))
    return pairs

# DEEPL_PROPOSAL_PREFIX_EN: ya no se usa (feedback Melissa, Oct26 — ver más abajo, en el bloque que
# arma target_texts): se dejó de anteponer este prefijo a las oraciones SÍ traducidas por DeepL,
# porque ahí el texto resultante es inglés completo y utilizable, no algo pendiente de revisar.
NO_DEEPL_PREFIX_EN = "SPANISH — DeepL not configured, translate manually: "
# NO_DEEPL_FAILED_PREFIX_EN: mismo caso de "se quedó en español, pendiente de traducir a mano", pero
# cuando DeepL SÍ está configurado y la llamada de todos modos falló (key inválida, cuota agotada,
# problema de red) — el motivo real queda en el log (DT.get_last_error()).
NO_DEEPL_FAILED_PREFIX_EN = "SPANISH — DeepL translation failed, translate manually: "

_DIFF_WS_RE = re.compile(r'\s+')

def _normalize_for_diff(s):
    """Normaliza un texto para compararlo contra el mismo texto del trimestre anterior: nunca se
    quiere traducir de nuevo (ni tratar como 'distinto') algo que solo difiere en un espacio doble,
    una comilla tipográfica vs. recta, o un salto de línea — ruido típico de guardar en Word, no un
    cambio real de redacción."""
    if s is None:
        return ''
    s = s.replace('’', "'").replace('‘', "'").replace('“', '"').replace('”', '"')
    s = s.replace('\xa0', ' ')
    return _DIFF_WS_RE.sub(' ', s.strip())

def _sanitize_for_markup(s):
    """set_paragraph_text usa [[...]] para resaltar en amarillo: si el contenido en sí trae un
    corchete suelto (del español original o de la traducción de DeepL), rompe ese parseo y el
    párrafo queda sin resaltar, con los corchetes de marcado visibles como texto literal. Para
    evitarlo, cualquier corchete que traiga el contenido se cambia por paréntesis antes de envolverlo
    en [[...]] — no se pierde información, solo el tipo de signo."""
    return s.replace('[', '(').replace(']', ')')

# Convierte cifras en formato chileno/español (punto = miles, coma = decimales) al formato en-US
# (coma = miles, punto = decimales) ANTES de mandar el texto a DeepL — no se le confía ese cambio a
# DeepL (es un traductor de texto, no un formateador de números: podría dejar la coma/punto tal cual
# vino, o interpretarlo mal), así el número que DeepL ve y traduce ya llega en el formato correcto y
# lo deja igual (los traductores no suelen tocar dígitos). Verificado con cifras reales del Excel de
# junio 2026: "107,81" -> "107.81", "1.239.108" -> "1,239,108", "1.239.108,50" -> "1,239,108.50",
# "6,10%" -> "6.10%".
_ES_NUMBER_RE = re.compile(r'\b\d{1,3}(?:\.\d{3})+(?:,\d+)?\b|\b\d+,\d+\b')

def _es_number_to_en(m):
    s = m.group(0)
    if '.' in s and ',' in s:
        int_part, dec_part = s.rsplit(',', 1)
        return f"{int_part.replace('.', ',')}.{dec_part}"
    elif '.' in s:
        return s.replace('.', ',')
    elif ',' in s:
        return s.replace(',', '.')
    return s

def convert_es_numbers_to_en(text):
    return _ES_NUMBER_RE.sub(_es_number_to_en, text)

# Mismo patrón que _ES_NUMBER_RE, pero para leer números YA en formato en-US (coma = miles, punto =
# decimales) de vuelta — se usa para ubicar, dentro de una oración en inglés ya aprobada el trimestre
# pasado, dónde están las cifras para reemplazarlas por las nuevas (ver _NUM_PLACEHOLDER más abajo).
_EN_NUMBER_RE = re.compile(r'\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b|\b\d+\.\d+\b')

# Pedido explícito de Melissa (oct-2026): "no quiero que la mayor parte del texto esté en amarillo.
# Es normal que las cifras cambien de un trimestre a otro." Antes, CUALQUIER diferencia de texto
# contra el trimestre anterior —incluido que una cifra cambiara, que es lo normal y esperado cada
# trimestre— mandaba la oración completa a traducir de nuevo vía DeepL y la resaltaba en amarillo
# para que la revisara, aunque la REDACCIÓN fuera exactamente la misma de siempre. Ahora, antes de
# tratar una oración como "cambiada", se compara su "plantilla" (el mismo texto, pero con cada cifra
# reemplazada por un comodín) contra la plantilla de la MISMA oración el trimestre anterior: si
# calzan (la redacción es idéntica y solo cambiaron las cifras, que es el caso normal de cada
# trimestre), se reutiliza la redacción en inglés YA aprobada el trimestre pasado, reemplazando sus
# cifras (en el mismo orden) por las nuevas del español de este trimestre — sin volver a llamar a
# DeepL y SIN resaltar en amarillo, porque la redacción nunca pasó por un traductor: las cifras salen
# directo del Word en español definitivo, igual que el resto del mecanismo. Solo si la plantilla en sí
# cambió (se agregó una causa, se reescribió la frase, cambió la cantidad de cifras mencionadas) se
# traduce de nuevo vía DeepL y se resalta, porque ahí sí hay redacción nueva que verificar.
_NUM_PLACEHOLDER = "\x00NUM\x00"

def _mask_numbers_es(text):
    return _ES_NUMBER_RE.sub(_NUM_PLACEHOLDER, text)

def _substitute_numbers_in_order(old_en_text, new_numbers_en):
    """Reemplaza, en orden de aparición, cada cifra de `old_en_text` (formato en-US) por la
    siguiente de `new_numbers_en` — se usa solo cuando ya se confirmó que la plantilla (texto sin
    cifras) es idéntica a la del trimestre anterior, así que la cantidad de cifras calza 1 a 1."""
    it = iter(new_numbers_en)
    def _repl(m):
        return next(it, m.group(0))
    return _EN_NUMBER_RE.sub(_repl, old_en_text)

def _flatten(segments):
    return [p for seg in segments for p in seg]

def translate_narrative_from_definitive_es(doc_en, word_es_edited_path, word_es_base_path, log,
                                            periods=None, old_periods=None):
    """Ver docstring de la sección de arriba. doc_en: el Word en inglés ya abierto (todavía con el
    texto en inglés del trimestre ANTERIOR tal cual el Word base subido, salvo encabezados/fechas y
    Flujo de Efectivo, ya actualizados antes de llamar a esto — ver generate_en). word_es_edited_path:
    ruta al Word en español definitivo de ESTE trimestre (obligatorio). word_es_base_path: ruta al
    Word en ESPAÑOL del trimestre ANTERIOR (opcional pero muy recomendado: es lo que permite reusar
    traducciones ya aprobadas en vez de traducir todo de nuevo cada vez). periods/old_periods: los
    diccionarios de períodos EN INGLÉS (periods_en.derive_periods_en / derive_old_periods_from_title_en)
    ya calculados en generate_en() — se usan solo para la coincidencia "misma redacción, cifras
    actualizadas" (ver template_lookup más abajo), para poder además actualizar menciones sueltas de
    Dec25/S24-25/etc. dentro de esa redacción reutilizada."""
    try:
        doc_es = docx.Document(word_es_edited_path)
    except Exception as ex:
        log.append(f"[ERROR] (EN) No pude abrir el Word en español definitivo para traducir el texto narrativo: {ex}")
        return

    # periods_es_curr / periods_es_prev: los mismos períodos de arriba pero en ESPAÑOL (Dic25, T25/26,
    # etc.), derivados del título de cada Word en español ("Al dd de mes de yyyy"). Se usan para
    # reescribir, dentro de una oración del trimestre ANTERIOR, cualquier mención de sus propias
    # etiquetas de período por las de ESTE trimestre antes de compararla (enmascarando además las
    # cifras) contra la oración actual — así "el EBITDA a Dic25 alcanzó..." y "el EBITDA a Jun26
    # alcanzó..." se reconocen como la MISMA redacción con cifras/período actualizados, en vez de cómo
    # texto nuevo (ver template_lookup más abajo). Sin esto, toda oración que menciona el período
    # (la gran mayoría) fallaba esta coincidencia solo por el cambio de "Dic25" a "Jun26", aunque el
    # resto de la frase fuera idéntica.
    periods_es_curr = None
    try:
        _, _title_es_curr = ASM_ES.find_paragraph(doc_es, r"^Al \d+ de \w+ de \d+\s*$")
        if _title_es_curr is not None:
            periods_es_curr = ASM_ES.P.derive_old_periods_from_title(_title_es_curr.text)
    except Exception:
        periods_es_curr = None

    header_anchors_es = [re.compile(a) for a, _ in ASM_ES.HEADER_TEMPLATES]
    # Se agregan también los patrones de STATIC_SUBTITLES_EN (ver su docstring): además de
    # HEADER_TEMPLATES_EN, estos títulos/subtítulos estáticos quedan excluidos de la segmentación para
    # que nunca puedan terminar reutilizados como donantes de contenido no relacionado.
    header_anchors_en = [re.compile(a) for a, *_ in HEADER_TEMPLATES_EN] + [
        re.compile(p) for p, *_ in STATIC_SUBTITLES_EN
    ]

    segs_es = _segment_narrative(doc_es, header_anchors_es)
    segs_en = _segment_narrative(doc_en, header_anchors_en)

    n_es, n_en = len(segs_es), len(segs_en)
    # _align_segments reemplaza el zip posicional directo de antes (ver su docstring: un solo párrafo
    # de más/menos en cualquier parte ya no desalinea el resto del documento).
    aligned_pairs, n_equal = _align_segments(segs_es, segs_en, log)
    if n_es and n_equal / n_es < 0.5:
        log.append(
            "[ALERTA] (EN) Menos de la mitad de la estructura del Word en inglés base calzó con la "
            "del Word en español definitivo. Esto casi siempre significa que el 'Word base en inglés "
            "del trimestre anterior' que subiste no es el correcto (revisa que esté en inglés y "
            "mantenga la misma estructura de siempre) — el documento generado probablemente quedó "
            "con la mayor parte del texto narrativo sin traducir bien."
        )

    # Diccionario de traducciones ya aprobadas (trimestre anterior): texto en español anterior
    # (normalizado) -> texto en inglés que YA tiene el Word en inglés en esa misma posición (que es,
    # segmento por segmento, la traducción aprobada de ese español anterior). Se arma ANTES de tocar
    # nada en doc_en (que es justo lo que trae ahora: el inglés del trimestre pasado).
    segs_es_prev = []
    periods_es_prev = None
    if word_es_base_path:
        try:
            doc_es_prev = docx.Document(word_es_base_path)
            segs_es_prev = _segment_narrative(doc_es_prev, header_anchors_es)
            try:
                _, _title_es_prev = ASM_ES.find_paragraph(doc_es_prev, r"^Al \d+ de \w+ de \d+\s*$")
                if _title_es_prev is not None:
                    periods_es_prev = ASM_ES.P.derive_old_periods_from_title(_title_es_prev.text)
            except Exception:
                periods_es_prev = None
        except Exception as ex:
            log.append(
                "[AVISO] (EN) No pude abrir el Word en español del trimestre anterior para reusar "
                f"traducciones ya aprobadas: {ex}. Se traducirá todo el texto narrativo de nuevo."
            )
    else:
        log.append(
            "[AVISO] (EN) No tengo el Word en español del trimestre anterior (sección en inglés, "
            "campo 'Word en ESPAÑOL del TRIMESTRE ANTERIOR', o reutilizado del 'Word base a "
            "actualizar' de la sección en español): se traducirá TODO el texto narrativo de nuevo en "
            "vez de reusar traducciones ya aprobadas — puede hacer que términos que antes se "
            "tradujeron de una forma salgan con una redacción levemente distinta esta vez."
        )

    old_lookup = {}
    # template_lookup: igual que old_lookup, pero con las cifras de cada oración enmascaradas (ver
    # _mask_numbers_es) — permite reconocer "misma redacción, cifras distintas" (el caso normal de
    # cada trimestre) y reutilizar la redacción en inglés ya aprobada con las cifras actualizadas, en
    # vez de tratarlo como texto nuevo (ver docstring de _NUM_PLACEHOLDER más arriba).
    template_lookup = {}
    _can_sweep_periods_es = periods_es_prev is not None and periods_es_curr is not None
    # Alineado por ESTRUCTURA (ver _align_segments_readonly), no por posición directa: un segmento de
    # más o de menos entre el español del trimestre anterior y el inglés base (título partido distinto,
    # etc. — ver docstring de _align_segments) ya no desalinea el resto de este diccionario.
    for es_prev_seg, en_prev_seg in _align_segments_readonly(segs_es_prev, segs_en):
        for p_es_prev, p_en_prev in zip(es_prev_seg, en_prev_seg):
            es_prev_text = p_es_prev.text
            en_prev_text = p_en_prev.text
            key = _normalize_for_diff(es_prev_text)
            if key:
                old_lookup[key] = en_prev_text
                # Para el template (ver docstring de la función): primero se reescriben, dentro de la
                # oración ANTERIOR, sus propias menciones de período (Dic25, T24/25, etc.) por las de
                # ESTE trimestre — igual mecanismo que sweep_text/apply_global_sweep, ya probado en el
                # resto del pipeline — y RECIÉN AHÍ se enmascaran las cifras. Así, una oración que solo
                # cambió en período y cifras calza con la de este trimestre aunque mencione el período
                # explícitamente (el caso más común: "el EBITDA a Dic25 alcanzó..." / "el EBITDA a
                # Jun26 alcanzó...").
                swept_prev_text = (
                    ASM_ES.P.sweep_text(es_prev_text, periods_es_prev, periods_es_curr)
                    if _can_sweep_periods_es else es_prev_text
                )
                template_key = _mask_numbers_es(_normalize_for_diff(swept_prev_text))
                if template_key != key:  # solo vale la pena si de verdad había alguna cifra/período
                    template_lookup[template_key] = (es_prev_text, en_prev_text)

    deepl_on = DT.is_configured()
    n_reused = 0
    n_numbers_updated = 0
    n_translated = 0
    n_fallback = 0
    n_fallback_failed = 0
    n_segs_mismatch = 0
    deepl_error_reasons = []

    n_segs_sin_destino = 0
    for es_items, en_items, trusted in aligned_pairs:
        if not en_items:
            # No había ningún párrafo en inglés disponible para clonar (caso extremo: el segmento
            # nuevo es el primero de todo el documento) — no hay dónde escribir, se deja constancia.
            n_segs_sin_destino += 1
            continue
        count_mismatch = len(es_items) != len(en_items)
        if count_mismatch:
            n_segs_mismatch += 1

        target_texts = []
        for p_es in es_items:
            es_text = p_es.text.strip()
            key = _normalize_for_diff(es_text)
            old_text = old_lookup.get(key)
            if old_text is not None:
                target_texts.append(old_text)
                n_reused += 1
                continue

            # ¿Misma redacción que el trimestre pasado, solo con cifras distintas? (ver docstring de
            # _NUM_PLACEHOLDER) — si calza, no hace falta DeepL ni resaltar: las cifras nuevas salen
            # directo de este mismo párrafo en español, igual que el resto del mecanismo.
            template_hit = template_lookup.get(_mask_numbers_es(key))
            if template_hit is not None:
                _, old_en_text = template_hit
                # Antes de sustituir las cifras, se barren también las menciones de período en INGLÉS
                # de esta redacción aprobada (Dec25, S24/25, etc. -> las de este trimestre) — mismo
                # mecanismo que apply_global_sweep, aplicado acá a esta oración puntual, porque viene de
                # la traducción aprobada del trimestre ANTERIOR y puede mencionar su propio período.
                old_en_text_swept = (
                    P.sweep_text_en(old_en_text, old_periods, periods)
                    if (periods is not None and old_periods is not None) else old_en_text
                )
                new_numbers_en = _EN_NUMBER_RE.findall(convert_es_numbers_to_en(es_text))
                old_numbers_en = _EN_NUMBER_RE.findall(old_en_text_swept)
                if len(new_numbers_en) == len(old_numbers_en):
                    target_texts.append(_substitute_numbers_in_order(old_en_text_swept, new_numbers_en))
                    n_numbers_updated += 1
                    continue
                # cantidad de cifras no calza (caso raro): no se puede sustituir con confianza, se
                # sigue de largo al camino normal de traducir de nuevo.

            # Sin resaltado (feedback Melissa, Oct26: "No es necesario el resaltado en los textos, no
            # me dice nada ya que todo está quedando resaltado. Excluir los resaltados."): como DeepL
            # está apagado por defecto (no hay DEEPL_API_KEY en los Secrets, ver docstring del módulo),
            # en la práctica casi toda oración nueva o distinta de la del trimestre anterior caía en
            # alguno de estos tres casos — con el marcador [[...]] eso resaltaba en amarillo la
            # oración COMPLETA (no un detalle puntual como la "causa" del lado español, ver
            # assemble.py), tiñendo de amarillo la mayor parte del documento. Se saca el marcador
            # [[...]] (ya no se resalta).
            #
            # Segunda vuelta (feedback Melissa, Oct26, con DeepL ya configurado): "la mayoria de los
            # textos son de DeepL translation. Quitar los textos que dicen: DeepL translation —
            # please verify: en todo el documento." Acá `translated` SÍ es una traducción al inglés
            # completa y utilizable, así que se saca también el prefijo (no solo el resaltado): no hay
            # nada que avisar, es una oración en inglés lista para usar.
            #
            # Los otros dos casos (DeepL apagado, o encendido pero la llamada falló) son distintos:
            # ahí lo que queda en el documento es TEXTO EN ESPAÑOL sin traducir — sacarle también el
            # prefijo dejaría español mezclado en silencio dentro del Word en inglés, sin ninguna
            # marca de que falta traducirlo (peor que el resaltado que se sacó: ahí al menos se nota
            # que algo falta). Por eso estos dos conservan el prefijo de texto plano (sin resaltar).
            translated = DT.translate_es_en(convert_es_numbers_to_en(es_text)) if deepl_on else None
            if translated:
                target_texts.append(_sanitize_for_markup(translated))
                n_translated += 1
            elif deepl_on:
                target_texts.append(f"{NO_DEEPL_FAILED_PREFIX_EN}{_sanitize_for_markup(es_text)}")
                n_fallback += 1
                n_fallback_failed += 1
                reason = DT.get_last_error()
                if reason and reason not in deepl_error_reasons:
                    deepl_error_reasons.append(reason)
            else:
                target_texts.append(f"{NO_DEEPL_PREFIX_EN}{_sanitize_for_markup(es_text)}")
                n_fallback += 1

        # neutralize_existing: ver docstring de _replace_list_items. Se activa si el emparejamiento no
        # es de fiar por CUALQUIERA de dos señales independientes: (a) no vino de un tramo 'equal' del
        # diff estructural (trusted=False, ver _align_segments), o (b) la cantidad de párrafos/viñetas
        # no calzaba entre español e inglés para este segmento (count_mismatch — la señal más directa:
        # si de verdad fuera la misma oración/bloque, debería traer la misma cantidad de viñetas en
        # los dos idiomas). Confirmado con un caso real: 2 viñetas en español ("Reducción en activos no
        # corrientes.../...corrientes...") terminaron emparejadas con 1 solo párrafo en inglés que
        # además pertenecía a OTRA sección ("RISK FACTOR ANALYSIS", negrita+subrayado) — ese tramo SÍ
        # venía marcado 'equal' (coincidencia de forma, no de contenido) pero el conteo no calzaba, así
        # que (b) lo atrapa aunque (a) no alcance a hacerlo.
        _replace_list_items(
            en_items, target_texts, numpr_source=es_items,
            neutralize_existing=(not trusted) or count_mismatch,
        )

    total = n_reused + n_numbers_updated + n_translated + n_fallback
    if total:
        msg = (
            f"[OK] (EN) Texto narrativo (todo el documento, salvo encabezados/fechas, Flujo de "
            f"Efectivo y tablas) comparado contra el Word en español definitivo en {len(aligned_pairs)} "
            f"segmento(s), {total} oración(es)/viñeta(s) en total: {n_reused} sin cambios (misma "
            f"redacción y cifras que el trimestre anterior), {n_numbers_updated} con la misma "
            "redacción ya aprobada pero cifras actualizadas (no necesitan revisión — las cifras "
            f"nuevas salen directo del español definitivo), {n_translated} con redacción nueva o "
            "distinta, traducida(s) vía DeepL"
        )
        if n_fallback:
            msg += f", {n_fallback} dejada(s) en español pendiente de traducir a mano"
        msg += (
            ". Por pedido de Melissa (Oct26): lo traducido por DeepL queda como texto en inglés "
            "normal, sin resaltar ni prefijo (ya no dice \"DeepL translation — please verify:\"); lo "
            "que quedó pendiente de traducir a mano sigue en español, sin resaltar, con un prefijo de "
            "texto plano (\"SPANISH — DeepL not configured, translate manually:\", etc.) — buscable "
            "con Ctrl+F."
        )
        log.append(msg)
        if not deepl_on and n_fallback:
            log.append(
                "[AVISO] (EN) DeepL no está configurado (falta DEEPL_API_KEY en los Secrets de la "
                "app): el texto narrativo nuevo o que cambió respecto al trimestre anterior quedó en "
                "español (sin resaltar, con el prefijo \"SPANISH — DeepL not configured, translate "
                "manually:\"), pendiente de traducir a mano."
            )
        if deepl_on and n_fallback_failed:
            motivos = "; ".join(deepl_error_reasons) if deepl_error_reasons else "motivo no disponible"
            log.append(
                f"[ALERTA-DEEPL] (EN) DeepL está configurado, pero {n_fallback_failed} traducción(es) "
                f"fallaron igual. Motivo reportado por DeepL: {motivos}. Esos párrafos quedaron en "
                "español (sin resaltar, con el prefijo \"SPANISH — DeepL translation failed, "
                "translate manually:\"), pendientes de traducir a mano — revisa la API key y la "
                "cuenta de DeepL (¿está verificada/activada? ¿tiene cuota disponible?)."
            )
    else:
        log.append("[AVISO] (EN) No encontré ningún segmento de texto narrativo para comparar/traducir.")
    if n_segs_mismatch:
        log.append(
            f"[AVISO] (EN) {n_segs_mismatch} de los {len(aligned_pairs)} segmentos alineados tenían "
            "distinta cantidad de párrafos/viñetas entre español e inglés (se agregaron o quitaron "
            "viñetas en inglés para que calce la cantidad) — revisa esos segmentos con más atención."
        )
    if n_segs_sin_destino:
        log.append(
            f"[ALERTA] (EN) {n_segs_sin_destino} segmento(s) de texto narrativo nuevo no tenían ningún "
            "párrafo en inglés cercano del cual copiar el formato (pasa solo si el documento no tiene "
            "casi ningún párrafo en común con el inglés base) y quedaron sin generar — revisa el "
            "'Word base en inglés del trimestre anterior' que subiste."
        )


def generate_en(excel_path, word_path, out_path, word_es_edited_path, word_es_base_path=None):
    if not word_es_edited_path:
        raise ValueError(
            "word_es_edited_path es obligatorio: hace falta el Word en español definitivo de este "
            "trimestre (ya revisado/editado) para generar el inglés — se usa tanto para las causas de "
            "negocio como para proponer la traducción del resto del texto libre del documento."
        )
    log = []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    periods = P.derive_periods_en(wb)
    log.append(f"[INFO] (EN) Período detectado: {periods['CUR_SHORT']} vs {periods['PRIOR_SHORT']} / {periods['FY_PRIOR_SHORT']}")

    doc = docx.Document(word_path)

    _, _title_para = find_paragraph(doc, r"^As of \w+ \d+, \d+\s*$")
    old_periods = P.derive_old_periods_from_title_en(_title_para.text) if _title_para is not None else None

    # collect_original_causes / _merge_causes_from_edited_es / apply_single_templates /
    # apply_bridge_bullets YA NO SE LLAMAN (ver docstring de translate_narrative_from_definitive_es,
    # más arriba): el texto narrativo con cifras y causas de negocio ahora sale de traducir
    # literalmente el Word en español definitivo, no de recalcular una plantilla en inglés aparte —
    # eso es justo lo que eliminó el descalce de cifras que reportó Melissa (oct-2026).
    apply_header_templates(doc, periods, log)
    apply_global_sweep(doc, old_periods, periods, log)
    found_native_flujo_caja = update_flujo_caja_table_en(doc, wb, log)
    if not found_native_flujo_caja:
        cloned_ok = _clone_flujo_caja_table_from_es(doc, word_es_edited_path, log)
        if cloned_ok:
            update_flujo_caja_table_en(doc, wb, log)
    remove_stale_risk_table_en(doc, log)
    translate_narrative_from_definitive_es(doc, word_es_edited_path, word_es_base_path, log,
                                            periods=periods, old_periods=old_periods)
    apply_static_subtitle_formatting_en(doc, log)

    image_map, img_dir = collect_and_render_tables_en(wb, word_path, log, doc=doc)
    from assemble import fix_image_frame_sizes, swap_media_bytes
    fix_image_frame_sizes(doc, image_map, log)
    _ensure_risks_image_en(doc, wb, log)

    tmp_text_docx = word_path + '.text_en.docx'
    doc.save(tmp_text_docx)

    swap_media_bytes(tmp_text_docx, out_path, image_map, log)
    os.unlink(tmp_text_docx)
    shutil.rmtree(img_dir)

    # Mismo chequeo que assemble.count_anchor_misses, para el lado en inglés: si casi ningún
    # encabezado/fecha calzó (ej. porque se subió por error el Word en ESPAÑOL como "Word base en
    # inglés del trimestre anterior"), avisar fuerte en vez de dejarlo pasar. Antes este chequeo
    # también contaba las ~30 plantillas de SINGLE_TEMPLATES_EN (de ahí el umbral de 10); ahora que
    # esas ya no se usan (ver arriba), solo quedan los 7 encabezados de HEADER_TEMPLATES_EN — el
    # chequeo de estructura narrativa completa (segmentos) ya lo hace translate_narrative_from_
    # definitive_es con su propio [ALERTA] cuando corresponde.
    misses = sum(1 for l in log if 'No encontré párrafo' in l)
    hits = sum(1 for l in log if l.startswith('[OK] (EN) Actualizado'))
    if misses + hits >= 3 and misses / (misses + hits) >= 0.5:
        log.append(
            f"[ALERTA] (EN) De {misses + hits} encabezados/fechas que la app busca por texto en el "
            f"Word base en inglés, {misses} NO se encontraron. Esto casi siempre significa que el "
            "Word subido en 'Word base en inglés del trimestre anterior' no es el correcto (revisa "
            "que esté en inglés y mantenga la misma estructura de siempre) — el documento generado "
            "probablemente quedó con la mayor parte del texto narrativo sin actualizar."
        )
    return log

if __name__ == '__main__':
    excel_path, word_path, out_path, word_es_edited_path = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
    word_es_base_path = sys.argv[5] if len(sys.argv) > 5 else None
    log = generate_en(excel_path, word_path, out_path, word_es_edited_path, word_es_base_path)
    print('\n'.join(log))
