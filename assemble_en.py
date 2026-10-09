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
import re, copy, io, os, sys, zipfile, tempfile, shutil
import docx
from docx.enum.text import WD_COLOR_INDEX
import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
import ar_engine_en as E
import periods_en as P

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

def collect_and_render_tables_en(wb, word_path, log):
    """Igual que collect_and_render_tables (assemble.py) pero renderiza cada tabla en inglés
    (render_table_en: mismo Excel en español, formato en-US + etiquetas traducidas vía TABLES_EN).
    La tabla 'Riesgos' (detalle narrativo de riesgos, todavía no traducido — ver render_tables_v5.py)
    se omite aquí a propósito, así que esa imagen queda con la del Word BASE en inglés tal cual
    (texto libre en español dentro de una tabla que de por sí ya viene en inglés en el base)."""
    rt_dir = os.path.dirname(__file__)
    sys.path.insert(0, rt_dir)
    import render_tables_v5 as RT

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
            RT.render_table_en(wb, tkey, out_png, dpi=300)
            image_map[a] = out_png; rendered[tkey] = out_png
            log.append(f"[OK] (EN) Tabla renderizada: {tkey}")
        except Exception as ex:
            log.append(f"[ERROR] (EN) Tabla {tkey}: {ex}")
    return image_map, img_dir

def generate_en(excel_path, word_path, out_path):
    log = []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    periods = P.derive_periods_en(wb)
    log.append(f"[INFO] (EN) Período detectado: {periods['CUR_SHORT']} vs {periods['PRIOR_SHORT']} / {periods['FY_PRIOR_SHORT']}")

    doc = docx.Document(word_path)

    _, _title_para = find_paragraph(doc, r"^As of \w+ \d+, \d+\s*$")
    old_periods = P.derive_old_periods_from_title_en(_title_para.text) if _title_para is not None else None

    causes = collect_original_causes(doc, log)
    apply_header_templates(doc, periods, log)
    apply_single_templates(doc, wb, periods, causes, log)
    apply_bridge_bullets(doc, wb, periods, causes, log)
    apply_global_sweep(doc, old_periods, periods, log)

    image_map, img_dir = collect_and_render_tables_en(wb, word_path, log)
    from assemble import fix_image_frame_sizes, swap_media_bytes
    fix_image_frame_sizes(doc, image_map, log)

    tmp_text_docx = word_path + '.text_en.docx'
    doc.save(tmp_text_docx)

    swap_media_bytes(tmp_text_docx, out_path, image_map, log)
    os.unlink(tmp_text_docx)
    shutil.rmtree(img_dir)
    return log

if __name__ == '__main__':
    excel_path, word_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    log = generate_en(excel_path, word_path, out_path)
    print('\n'.join(log))
