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

def set_paragraph_text(paragraph, text):
    # Conserva fuente/tamaño/subrayado del primer run original (ver assemble.py: algunos subtítulos de
    # sección vienen subrayados en el Word base y deben seguir subrayados aunque se regenere el texto).
    base_font_name = None; base_size = None; base_underline = None
    if paragraph.runs:
        r0 = paragraph.runs[0]
        base_font_name = r0.font.name; base_size = r0.font.size
        base_underline = r0.font.underline
    for r in list(paragraph.runs):
        r._element.getparent().remove(r._element)

    pos = 0
    for m in MARK_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], base_font_name, base_size, None, False, base_underline)
        if m.group(1) is not None:
            _add_run(paragraph, m.group(1), base_font_name, base_size, None, True, base_underline)
        else:
            _add_run(paragraph, m.group(2), base_font_name, base_size, True, False, base_underline)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], base_font_name, base_size, None, False, base_underline)

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
HEADER_TEMPLATES_EN = [
    (r"^As of \w+ \d+, \d+\s*$", E.p_titulo_fecha),
    (r"^The current reasoned analysis has been prepared for the period ending", E.p_preambulo_1),
    (r"^Since the Company administers its operations with an agricultural season", E.p_preambulo_2),
    (r"^Accumulated EBITDA analysis as of \w+", E.p_header_ebitda_acumulado),
    (r"^Income Statement analysis as of \w+", E.p_header_resultado_calendario),
    (r"^Income Statement Analysis .* season\s*$", E.p_header_resultado_temporada),
    (r"^Accumulated income analysis as of \w+", E.p_header_ingresos_acumulados),
]

def apply_header_templates(doc, periods, log):
    for anchor, fn in HEADER_TEMPLATES_EN:
        idx, para = find_paragraph(doc, anchor)
        if para is None:
            log.append(f"[AVISO] (EN) No encontré párrafo para patrón: {anchor[:60]}")
            continue
        try:
            set_paragraph_text(para, fn(periods))
            log.append(f"[OK] (EN) Actualizado (encabezado/preámbulo): {anchor[:50]}")
        except Exception as ex:
            log.append(f"[ERROR] (EN) {anchor[:50]}: {ex}")

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

def _replace_list_items(bullet_paras, items_texts):
    n_existing = len(bullet_paras)
    n_target = len(items_texts)
    for k in range(min(n_existing, n_target)):
        set_paragraph_text(bullet_paras[k], items_texts[k])
    if n_target > n_existing:
        final = list(bullet_paras)
        last_el = bullet_paras[-1]._p
        for k in range(n_existing, n_target):
            new_el = copy.deepcopy(last_el)
            last_el.addnext(new_el)
            new_para = docx.text.paragraph.Paragraph(new_el, bullet_paras[-1]._parent)
            set_paragraph_text(new_para, items_texts[k])
            final.append(new_para)
            last_el = new_el
        return final
    elif n_existing > n_target:
        for p in bullet_paras[n_target:]:
            p._p.getparent().remove(p._p)
        return list(bullet_paras[:n_target])
    return list(bullet_paras)

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

    SKIP_KEYS = set()  # 'Riesgos' ya se traduce (ver TABLES_EN en render_tables_v5.py)

    with zipfile.ZipFile(word_path, 'r') as z:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    actual = [n for n in re.findall(r'Target="media/([^"]+)"', rels) if re.match(r'image\d+', n)]
    actual = sorted(actual, key=lambda x: int(re.search(r'\d+', x).group()))
    c2a = {}
    for i, a in enumerate(actual, 1):
        for ext in ['.emf', '.png']: c2a[f'image{i}{ext}'] = a

    img_dir = tempfile.mkdtemp(); image_map = {}; rendered = {}
    for img_name, tkey in RT.IMAGE_TO_TABLE.items():
        if tkey in SKIP_KEYS: continue
        if tkey not in RT.TABLES or tkey not in RT.TABLES_EN: continue
        spec = RT.TABLES[tkey]
        if spec['sheet'] not in wb.sheetnames: continue
        a = c2a.get(img_name)
        if not a: continue
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
        return
    ws = wb['FdeC']
    hr, hc = ASM_ES._find_fdec_header(ws)
    if hr is None:
        log.append("[AVISO] (EN) No encontré el encabezado del Flujo de Efectivo en la hoja 'FdeC': no se actualizó.")
        return
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
        return

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
        new_paras.append(docx.text.paragraph.Paragraph(new_el, anchor_para._parent))
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
        new_paras.append(docx.text.paragraph.Paragraph(new_el, anchor_para._parent))
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
    vez de dejarlo tal cual (que es justo lo que producía el dato viejo duplicado)."""
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
                pairs.append((segs_es[j1 + k], segs_en[i1 + k]))
                last_en_para = segs_en[i1 + k][-1]
                n_equal += 1
            continue
        n_en_here = i2 - i1
        n_es_here = j2 - j1
        n_both = min(n_en_here, n_es_here)
        for k in range(n_both):
            pairs.append((segs_es[j1 + k], segs_en[i1 + k]))
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
                pairs.append((segs_es[k], new_seg))
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

DEEPL_PROPOSAL_PREFIX_EN = "DeepL translation — please verify: "
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
    header_anchors_en = [re.compile(a) for a, _ in HEADER_TEMPLATES_EN]

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
    for es_items, en_items in aligned_pairs:
        if not en_items:
            # No había ningún párrafo en inglés disponible para clonar (caso extremo: el segmento
            # nuevo es el primero de todo el documento) — no hay dónde escribir, se deja constancia.
            n_segs_sin_destino += 1
            continue
        if len(es_items) != len(en_items):
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

            translated = DT.translate_es_en(convert_es_numbers_to_en(es_text)) if deepl_on else None
            if translated:
                target_texts.append(f"[[{DEEPL_PROPOSAL_PREFIX_EN}{_sanitize_for_markup(translated)}]]")
                n_translated += 1
            elif deepl_on:
                target_texts.append(f"[[{NO_DEEPL_FAILED_PREFIX_EN}{_sanitize_for_markup(es_text)}]]")
                n_fallback += 1
                n_fallback_failed += 1
                reason = DT.get_last_error()
                if reason and reason not in deepl_error_reasons:
                    deepl_error_reasons.append(reason)
            else:
                target_texts.append(f"[[{NO_DEEPL_PREFIX_EN}{_sanitize_for_markup(es_text)}]]")
                n_fallback += 1

        _replace_list_items(en_items, target_texts)

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
        msg += ". Solo lo traducido por redacción nueva queda resaltado en amarillo para que lo verifiques."
        log.append(msg)
        if not deepl_on and n_fallback:
            log.append(
                "[AVISO] (EN) DeepL no está configurado (falta DEEPL_API_KEY en los Secrets de la "
                "app): el texto narrativo nuevo o que cambió respecto al trimestre anterior quedó en "
                "español, resaltado, pendiente de traducir a mano."
            )
        if deepl_on and n_fallback_failed:
            motivos = "; ".join(deepl_error_reasons) if deepl_error_reasons else "motivo no disponible"
            log.append(
                f"[ALERTA-DEEPL] (EN) DeepL está configurado, pero {n_fallback_failed} traducción(es) "
                f"fallaron igual. Motivo reportado por DeepL: {motivos}. Esos párrafos quedaron en "
                "español, resaltados, pendientes de traducir a mano — revisa la API key y la cuenta "
                "de DeepL (¿está verificada/activada? ¿tiene cuota disponible?)."
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
    update_flujo_caja_table_en(doc, wb, log)
    remove_stale_risk_table_en(doc, log)
    translate_narrative_from_definitive_es(doc, word_es_edited_path, word_es_base_path, log,
                                            periods=periods, old_periods=old_periods)

    image_map, img_dir = collect_and_render_tables_en(wb, word_path, log, doc=doc)
    from assemble import fix_image_frame_sizes, swap_media_bytes
    fix_image_frame_sizes(doc, image_map, log)

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
