# -*- coding: utf-8 -*-
"""
Ensambla el Word final: texto narrativo regenerado + tablas/imágenes actualizadas.
No usa merge_runs.py: cada párrafo cubierto se regenera completo (no es find/replace de texto viejo).
"""
import re, copy, io, os, sys, zipfile, tempfile, shutil
import docx
from docx.enum.text import WD_COLOR_INDEX
import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
import ar_engine as E
import periods as P

# [[...]] = resaltar en amarillo (causa no derivable del Excel); {{...}} = negrita (frase líder tipo
# "Reducción en activos no corrientes en US$X millones:" que algunos párrafos del original llevan en negrita).
MARK_RE = re.compile(r'\[\[([^\]]+)\]\]|\{\{([^}]+)\}\}')

def set_paragraph_text(paragraph, text):
    """Reemplaza el contenido de un párrafo por `text`, resaltando en amarillo los tramos [[...]] y
    poniendo en negrita los tramos {{...}}. Conserva fuente/tamaño del primer run original, pero no
    hereda su negrita (se indica explícitamente con {{...}} donde corresponda)."""
    base_font_name = None; base_size = None
    if paragraph.runs:
        r0 = paragraph.runs[0]
        base_font_name = r0.font.name; base_size = r0.font.size
    for r in list(paragraph.runs):
        r._element.getparent().remove(r._element)

    pos = 0
    for m in MARK_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], base_font_name, base_size, None, False)
        if m.group(1) is not None:
            _add_run(paragraph, m.group(1), base_font_name, base_size, None, True)
        else:
            _add_run(paragraph, m.group(2), base_font_name, base_size, True, False)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], base_font_name, base_size, None, False)

def _add_run(paragraph, text, font_name, size, bold, highlight):
    run = paragraph.add_run(text)
    if font_name: run.font.name = font_name
    if size: run.font.size = size
    if bold is not None: run.bold = bold
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

SINGLE_TEMPLATES = [
    (r"^El EBITDA a \w+\d+ alcanzó US\$", E.p_ebitda_6m),
    (r"^El EBITDA acumulado a \w+\d+ sin efecto de fair value", E.p_ebitda_sinfv_6m),
    (r"^Durante los 6 meses terminados en \w+\d+ se registró", E.p_ganancia_controladora_6m),
    (r"^Durante los 12M T\d+/\d+, el EBITDA ascendió", E.p_ebitda_12m_intro),
    (r"^Incremento en los ingresos del [\d,]+% asociado|^Disminución en los ingresos del [\d,]+% asociado", E.p_ebitda_12m_bullet_ingresos),
    (r"^Por su parte, los costos y gastos sin incluir depreciación registraron un", E.p_ebitda_12m_bullet_costos),
    (r"^La ganancia \(pérdida\) atribuible a los propietarios de la controladora fue de US\$[\d\.,\-]+ millones en los 12M", E.p_ganancia_controladora_12m),
    (r"^Los ingresos de actividades ordinarias alcanzaron US\$[\d\.,]+ millones a \w+\d+,", E.p_ingresos_totales_6m),
    (r"^Las ventas del segmento de Fruta Fresca a \w+\d+ ", E.p_frutafresca_6m),
    (r"^Los productos con valor agregado registraron un incremento en los ingresos por venta a \w+\d+", E.p_valoragregado_6m),
    (r"^Los ingresos de actividades ordinarias alcanzaron US\$[\d\.,]+ millones en los 12M", E.p_ingresos_totales_12m),
    (r"^Las ventas del segmento de Fruta Fresca en los 12M", E.p_frutafresca_12m),
    (r"^Por su parte, el segmento de productos con valor agregado registró un incremento", E.p_valoragregado_12m),
    (r"^Los costos de ventas a \w+\d+ totalizaron", E.p_costoventa_6m),
    (r"^Los costos de ventas de los 12M .* alcanzaron", E.p_costoventa_12m),
    (r"^Los gastos de administración a \w+\d+ alcanzaron", E.p_gastosadmin_6m),
    (r"^Los otros gastos, por función \(excluyendo el deterioro", E.p_otrosgastos_6m),
    (r"^El gasto por deterioro de valor de activos a \w+\d+ fue de", E.p_deterioro_6m),
    (r"^Los otros componentes del resultado registraron un costo", E.p_otroscomponentes_6m),
    (r"^A \w+\d+, se registraron gastos por impuesto a las ganancias", E.p_impuesto_6m),
    (r"^La rotación de los activos al \d+ de \w+ de \d+", E.p_rotacion_activos),
    (r"^Por su parte, el ratio de rotación de inventarios", E.p_rotacion_inventarios),
    (r"^La deuda financiera neta de la Sociedad", E.p_dfn),
    (r"^La liquidez corriente fue de", E.p_liquidez),
    (r"^En tanto, la razón ácida alcanzó", E.p_razon_acida),
    (r"^La razón de endeudamiento (disminuyó|aumentó)", E.p_razon_endeudamiento),
    (r"^Durante este periodo, la compañía (mantuvo|modificó) su estructura de deuda", E.p_estructura_deuda),
    (r"^El índice de cobertura de gastos financieros se ubica", E.p_cobertura_gf),
    (r"^La rentabilidad del patrimonio de la controladora", E.p_rentabilidad_patrim_controladora),
    (r"^Por su parte, la rentabilidad del patrimonio total", E.p_rentabilidad_patrim_total),
    (r"^Al \d+ de \w+ de \d+, los activos totales se", E.p_activos_totales_intro),
    (r"^(Reducción|Incremento) en activos no corrientes en", E.p_activos_nocorrientes_bridge),
    (r"^(Reducción|Incremento) en activos corrientes en", E.p_activos_corrientes_bridge),
    (r"^Los pasivos totales se (redujeron|incrementaron)", E.p_pasivos_bridge),
    (r"^El patrimonio total de la Compañía", E.p_patrimonio_bridge),
]

def apply_single_templates(doc, wb, periods, log):
    for anchor, fn in SINGLE_TEMPLATES:
        idx, para = find_paragraph(doc, anchor)
        if para is None:
            log.append(f"[AVISO] No encontré párrafo para patrón: {anchor[:60]}")
            continue
        try:
            text = P.localize(fn(wb), periods)
            set_paragraph_text(para, text)
            log.append(f"[OK] Actualizado: {anchor[:50]}")
        except Exception as ex:
            log.append(f"[ERROR] {anchor[:50]}: {ex}")

# ---------- Bridge bullets (Ganancia Atribuible) ----------
def _next_list_block(doc, start, max_scan=14):
    """Desde el índice start (inclusive): salta párrafos en blanco y luego junta los párrafos
    consecutivos con estilo 'List Paragraph' (la viñeta). Devuelve (lista_de_parrafos, índice_siguiente)."""
    n = len(doc.paragraphs)
    j = start
    while j < n and j < start + max_scan and doc.paragraphs[j].text.strip() == '':
        j += 1
    items = []
    while j < n and j < start + max_scan:
        p = doc.paragraphs[j]
        style_name = p.style.name if p.style else ''
        # un párrafo "List Paragraph" vacío se usa como separador visual, no como ítem real;
        # se trata igual que un párrafo normal (corta el bloque) para no romper el índice.
        if 'List' in style_name and p.text.strip() != '':
            items.append(p)
            j += 1
        else:
            break
    return items, j

def _replace_list_items(bullet_paras, items_texts):
    """Actualiza una lista de párrafos de viñeta ya existente para que contenga exactamente
    items_texts, clonando o eliminando párrafos según haga falta."""
    n_existing = len(bullet_paras)
    n_target = len(items_texts)
    for k in range(min(n_existing, n_target)):
        set_paragraph_text(bullet_paras[k], items_texts[k])
    if n_target > n_existing:
        last_el = bullet_paras[-1]._p
        for k in range(n_existing, n_target):
            new_el = copy.deepcopy(last_el)
            last_el.addnext(new_el)
            new_para = docx.text.paragraph.Paragraph(new_el, bullet_paras[-1]._parent)
            set_paragraph_text(new_para, items_texts[k])
            last_el = new_el
    elif n_existing > n_target:
        for p in bullet_paras[n_target:]:
            p._p.getparent().remove(p._p)

def replace_bullet_block(doc, intro_anchor_regex, items_texts, max_existing_scan=14):
    idx, intro = find_paragraph(doc, intro_anchor_regex)
    if intro is None:
        return False
    bullet_paras, _ = _next_list_block(doc, idx + 1, max_existing_scan)
    if not bullet_paras:
        return False
    _replace_list_items(bullet_paras, items_texts)
    return True

def build_bridge_items_6m(wb, max_items=6):
    ebitda_text = E.p_ebitda_bullet_6m(wb)
    items = E.bridge_bullets(wb, '6m', threshold_abs_mus=3.0)[:max_items]
    rest = [E.bridge_factor_text(lbl, cur, prior, delta, '6m') for lbl, cur, prior, delta in items]
    return [ebitda_text] + rest

def apply_bridge_bullets(doc, wb, periods, log):
    # --- 6M: una sola lista de viñetas tras "Las principales variaciones se explican a continuación" ---
    items_6m = build_bridge_items_6m(wb)
    ok = replace_bullet_block(doc, r"Las principales variaciones se explican a continuación", items_6m)
    log.append(f"[{'OK' if ok else 'AVISO'}] Bridge 6M ({len(items_6m)} factores)")

    # --- 12M: dos listas (factores positivos / factores que compensan) separadas por un párrafo fijo ---
    idx, intro = find_paragraph(doc, r"^La ganancia \(pérdida\) atribuible a los propietarios de la controladora fue de US\$")
    if intro is None:
        log.append("[AVISO] No encontré la intro del bridge de Ganancia Atribuible 12M")
        return
    items = E.bridge_bullets(wb, '12m', threshold_abs_mus=3.0)
    pos = sorted([it for it in items if it[3] >= 0], key=lambda x: abs(x[3]), reverse=True)
    neg = sorted([it for it in items if it[3] < 0], key=lambda x: abs(x[3]), reverse=True)
    pos_texts = [E.p_ebitda_bullet_12m(wb)] + [E.bridge_factor_text(lbl, cur, prior, delta, '12m') for lbl, cur, prior, delta in pos]
    neg_texts = [E.bridge_factor_text(lbl, cur, prior, delta, '12m') for lbl, cur, prior, delta in neg]

    pos_paras, after_pos = _next_list_block(doc, idx + 1, 14)
    if not pos_paras:
        log.append("[AVISO] No encontré la lista de factores positivos (12M GA)")
        return
    _replace_list_items(pos_paras, pos_texts)

    j = after_pos
    n = len(doc.paragraphs)
    while j < n and doc.paragraphs[j].text.strip() == '':
        j += 1
    sep_idx = j if (j < n and 'compensad' in doc.paragraphs[j].text.lower()) else None

    if not neg_texts:
        log.append(f"[OK] Bridge 12M GA ({len(pos_texts)} factores positivos, sin compensatorios)")
        return
    if sep_idx is None:
        log.append(f"[AVISO] Bridge 12M GA: {len(pos_texts)} factores positivos OK, pero no encontré el párrafo "
                    f"separador ('...compensados por los siguientes efectos:') para los {len(neg_texts)} compensatorios")
        return
    neg_paras, _ = _next_list_block(doc, sep_idx + 1, 14)
    if not neg_paras:
        log.append(f"[AVISO] Bridge 12M GA: {len(pos_texts)} factores positivos OK, pero no encontré la lista de "
                    f"factores compensatorios")
        return
    _replace_list_items(neg_paras, neg_texts)
    log.append(f"[OK] Bridge 12M GA ({len(pos_texts)} factores positivos, {len(neg_texts)} compensatorios)")

# ---------- Tablas / imágenes ----------
def render_all_tables(wb, docx_in_path, docx_out_path, log):
    rt_dir = os.path.dirname(__file__)
    sys.path.insert(0, rt_dir)
    import render_tables_v5 as RT

    SHEET_FIXES = {
        'IngresosxSegmento': 'IngresosxSegmento',  # ya sin "(2)" en este archivo
    }
    RT.TABLES['IngresosxSegmento']['sheet'] = 'IngresosxSegmento'
    if 'Indicadores1' in wb.sheetnames:
        pass
    elif 'Ind. Financieros' in wb.sheetnames:
        RT.TABLES['Indicadores1']['sheet'] = 'Ind. Financieros'

    SKIP_KEYS = set()
    if 'Indicadores2' not in wb.sheetnames:
        # las 3 hojas se separaron con layouts distintos; no re-especificadas aún -> no tocar esas imágenes
        SKIP_KEYS.add('IndicadoresActividad')
        SKIP_KEYS.add('Rentabilidad')

    img_dir = tempfile.mkdtemp(); image_map = {}; rendered = {}
    for img_name, tkey in RT.IMAGE_TO_TABLE.items():
        if tkey in SKIP_KEYS: continue
        if tkey not in RT.TABLES: continue
        spec = RT.TABLES[tkey]
        if spec['sheet'] not in wb.sheetnames: continue
        if tkey in rendered:
            image_map[img_name] = rendered[tkey]; continue
        out_png = os.path.join(img_dir, f'{tkey}.png')
        try:
            RT.render_table(wb, spec, out_png, dpi=300)
            image_map[img_name] = out_png; rendered[tkey] = out_png
            log.append(f"[OK] Tabla renderizada: {tkey}")
        except Exception as ex:
            log.append(f"[ERROR] Tabla {tkey}: {ex}")

    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(docx_in_path, 'r') as z: z.extractall(tmp)
    rp = os.path.join(tmp, 'word', '_rels', 'document.xml.rels')
    with open(rp, encoding='utf-8') as f: rels = f.read()
    actual = [n for n in re.findall(r'Target="media/([^"]+)"', rels) if re.match(r'image\d+', n)]
    actual = sorted(actual, key=lambda x: int(re.search(r'\d+', x).group()))
    c2a = {}
    for i, a in enumerate(actual, 1):
        for ext in ['.emf', '.png']: c2a[f'image{i}{ext}'] = a
    repl = 0
    for old, new_png in image_map.items():
        a = c2a.get(old)
        if not a: continue
        nn = re.sub(r'(\.\w+)$', '_v6.png', a)
        shutil.copy2(new_png, os.path.join(tmp, 'word', 'media', nn))
        rels = rels.replace(f'Target="media/{a}"', f'Target="media/{nn}"')
        repl += 1
    with open(rp, 'w', encoding='utf-8') as f: f.write(rels)
    ct = os.path.join(tmp, '[Content_Types].xml')
    with open(ct, encoding='utf-8') as f: cxml = f.read()
    if 'image/png' not in cxml:
        cxml = cxml.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
    with open(ct, 'w', encoding='utf-8') as f: f.write(cxml)
    if os.path.exists(docx_out_path): os.unlink(docx_out_path)
    with zipfile.ZipFile(docx_out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmp):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, tmp))
    shutil.rmtree(tmp); shutil.rmtree(img_dir)
    log.append(f"[OK] {repl} imágenes de tabla actualizadas")

def generate(excel_path, word_path, out_path):
    log = []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    periods = P.derive_periods(wb)
    log.append(f"[INFO] Período detectado: {periods['CUR_SHORT']} vs {periods['PRIOR_SHORT']} / {periods['FY_PRIOR_SHORT']}")

    doc = docx.Document(word_path)
    apply_single_templates(doc, wb, periods, log)
    apply_bridge_bullets(doc, wb, periods, log)

    tmp_text_docx = word_path + '.text.docx'
    doc.save(tmp_text_docx)

    render_all_tables(wb, tmp_text_docx, out_path, log)
    os.unlink(tmp_text_docx)
    return log

if __name__ == '__main__':
    excel_path, word_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    log = generate(excel_path, word_path, out_path)
    print('\n'.join(log))
