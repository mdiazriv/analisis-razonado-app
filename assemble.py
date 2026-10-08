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

# ---------- Extracción de causas de negocio ya escritas en el Word original ----------
# Estas funciones leen el texto ORIGINAL (antes de sobrescribirlo) y recortan la frase de causa que
# ya viene redactada, para preservarla (resaltada en amarillo) en vez de generar una instrucción
# genérica. Se diseñaron mirando el documento real Jun26; si el patrón no calza en algún trimestre
# futuro (porque cambió la redacción base), simplemente no se encuentra una causa y el párrafo queda
# solo con el tag CAUSA_TAG para que la persona responsable la complete.

def _extract_deterioro_6m_causes(text):
    # Estilo Jun26 (causas separadas cur/prior): "...millones asociado a X, en comparación con
    # ...millones a Jun25, asociado a Y." Estilo Sep/Dic (causa única combinada al final): "...millones,
    # en comparación con ...millones a Sep24, explicado por <causa combinada>." En este 2do caso no hay
    # como separar cur de prior, así que la causa combinada se asigna a cause_prior (queda en la
    # posición final de la frase, igual que en el original).
    cause_cur = None; cause_prior = None
    m = re.search(r'millones\s+(asociado.*?),\s*en comparaci[oó]n con', text)
    if m: cause_cur = m.group(1).strip()
    m2 = re.search(r'a \w+\d+,\s*((?:asociado|explicado).*?)\.\s*$', text)
    if m2: cause_prior = m2.group(1).strip()
    return cause_cur, cause_prior

def _extract_gastosadmin_6m_cause(text):
    m = re.search(r'con respecto a \w+\d+,\s*(explicado.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_costoventa_12m_cause(text):
    # Ancla en "comparado con el X% en los NM T.../.." (la 2da mención de "en los NM", la del
    # párrafo de ratio costos/ingresos) para no confundirla con la mención anterior del mismo patrón.
    # \d+M (no solo "12M") porque el documento base puede venir de cualquiera de los 4 cierres.
    # \s* (no un espacio fijo) entre \d+M y T porque Sep/Dic escriben "3MT25/26"/"6MT25/26" sin espacio.
    m = re.search(r'comparado con el [\d,]+%\s*en los \d+M\s*T\d+/\d+,\s*([^\.]+)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_razon_endeudamiento_cause(text):
    # Ancla en "pasivos totales ... en (un) X%, asociado..." en vez de solo "en un X%, asociado...":
    # Jun26 dice "la reducción de los pasivos totales en un 7,18%, asociado..."; Sep/Dic en cambio
    # traen esa misma frase en una oración aparte y sin "un": "los pasivos totales se redujeron en
    # 3,68%, asociado...". Ambas comparten "pasivos totales ... en (un) X%, asociado...".
    m = re.search(r'pasivos totales[^.]*?en (?:un )?[\d,]+%,\s*(asociado.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_cobertura_gf_cause(text):
    m = re.search(r'respecto a \w+\d+\s+(asociado.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_bridge_deterioro_cause(text):
    # Jun26: "...millones asociado a X, en comparación con los US$...". Sep/Dic: "...millones asociado
    # a X. Este deterioro se compara con...": misma frase de causa, pero cierra con punto + oración
    # nueva en vez de coma, así que se acepta cualquiera de los dos cierres.
    m = re.search(r'millones\s+(asociado.*?)(?:,\s*en comparaci[oó]n con|\.\s*Este \w+ se compara con)', text)
    return m.group(1).strip() if m else None

def _extract_bridge_depreciacion_cause(text):
    # Jun26/Dic: "(...), explicado principalmente por X." Sep: "(...) por X." (sin "explicado
    # principalmente"). Se acepta la causa con o sin ese prefijo.
    m = re.search(r'\),\s*(explicado.*?|por .*?)\.\s*$', text)
    return m.group(1).strip() if m else None

def _extract_bridge_costosfin_cause(text):
    m = re.search(r'millones,\s*(asociado.*?)\.\s*$', text)
    return m.group(1).strip() if m else None

# Palabras clave para reconocer, por su texto, a qué factor del bridge de Ganancia Atribuible
# corresponde cada viñeta original (las viñetas son posicionalmente variables de trimestre a
# trimestre, así que no se pueden ubicar por índice fijo).
LABEL_KEYWORDS = {
    'Deterioro': ['deterioro'],
    'Depreciación y amortizaciones': ['depreciación y amortización', 'depreciacion y amortizacion'],
    'Costos Financieros Netos': ['gastos financieros netos', 'costos financieros netos'],
    'Diferencia de cambio': ['diferencia de cambio'],
    'Gasto por impuestos a las ganancias': ['impuestos'],
    'Otros ingresos (gastos) no operacionales': ['otros ingresos', 'otros gastos no operacionales'],
    'Participación en asociadas': ['participación en', 'participacion en'],
}

def classify_bridge_label(text):
    t = text.lower()
    for label, kws in LABEL_KEYWORDS.items():
        if any(kw in t for kw in kws):
            return label
    return None

# Nota sobre \d+M / \d+ meses: los templates siempre se generan en "estilo Jun26" (12M, 6 meses) y
# periods.localize() los convierte al conteo real del cierre (3/6/9/12M, 3/6/9/12 meses). Pero estos
# patrones ANCLAN sobre el texto del documento BASE (el trimestre anterior, que puede ser cualquiera
# de los 4 cierres), así que deben aceptar cualquier conteo de meses, no solo "12M"/"6 meses".
SINGLE_TEMPLATES = [
    (r"^El EBITDA a \w+\d+ alcanzó US\$", E.p_ebitda_6m),
    (r"^El EBITDA acumulado a \w+\d+ sin efecto de fair value", E.p_ebitda_sinfv_6m),
    (r"^Durante los \d+ meses terminados en \w+\d+ se registró", E.p_ganancia_controladora_6m),
    (r"^Durante los \d+M T\d+/\d+, el EBITDA ascendió", E.p_ebitda_12m_intro),
    (r"^Incremento en los ingresos del [\d,]+% asociado|^Disminución en los ingresos del [\d,]+% asociado", E.p_ebitda_12m_bullet_ingresos),
    (r"^Por su parte, los costos y gastos sin incluir depreciación registraron un", E.p_ebitda_12m_bullet_costos),
    (r"^La (ganancia|pérdida)(\s*\(pérdida\))? atribuible a los propietarios de la controladora (fue de|se registr[oó] en) US\$[\d\.,\-]+ millones en los \d+M", E.p_ganancia_controladora_12m),
    (r"^Los ingresos de actividades ordinarias alcanzaron US\$[\d\.,]+ millones a \w+\d+,", E.p_ingresos_totales_6m),
    (r"^Las ventas del segmento de Fruta Fresca a \w+\d+ ", E.p_frutafresca_6m),
    (r"^Los productos con valor agregado registraron (un incremento|una disminución|una reducción) en los ingresos por venta a \w+\d+", E.p_valoragregado_6m),
    (r"^Los ingresos de actividades ordinarias alcanzaron US\$[\d\.,]+ millones en los \d+M", E.p_ingresos_totales_12m),
    (r"^Las ventas del segmento de Fruta Fresca en los \d+M", E.p_frutafresca_12m),
    (r"^Por su parte, el segmento de productos con valor agregado registró (un incremento|una reducción|una disminución)", E.p_valoragregado_12m),
    (r"^Los costos de ventas a \w+\d+ totalizaron", E.p_costoventa_6m),
    (r"^Los costos de ventas de los \d+M\s*T\d+/\d+ alcanzaron", E.p_costoventa_12m),
    (r"^Los gastos de administración a \w+\d+ alcanzaron", E.p_gastosadmin_6m),
    (r"^Los otros gastos, por función \(excluyendo el deterioro", E.p_otrosgastos_6m),
    (r"^El gasto por deterioro de valor de activos a \w+\d+ fue de|^A \w+\d+, el gasto por deterioro de valor de activos fue de", E.p_deterioro_6m),
    (r"^Los otros componentes del resultado registraron un costo", E.p_otroscomponentes_6m),
    (r"^A \w+\d+, se registr(aron gastos|ó un gasto) por impuesto a las ganancias", E.p_impuesto_6m),
    (r"^La rotación de los activos al \d+ de \w+ de \d+", E.p_rotacion_activos),
    (r"^Por su parte, el ratio de rotación de inventarios", E.p_rotacion_inventarios),
    (r"^La deuda financiera neta de la Sociedad", E.p_dfn),
    (r"^La liquidez corriente fue de", E.p_liquidez),
    (r"^En tanto, la razón ácida alcanzó", E.p_razon_acida),
    (r"^La razón de endeudamiento (disminuyó|aumentó|incrementó)", E.p_razon_endeudamiento),
    (r"^Durante este periodo, la compañía (mantuvo|modificó) su estructura de deuda|^Durante este periodo, la compañía refinanció", E.p_estructura_deuda),
    (r"^El índice de cobertura de gastos financieros se ubica", E.p_cobertura_gf),
    (r"^La rentabilidad del patrimonio de la controladora", E.p_rentabilidad_patrim_controladora),
    (r"^Por su parte, la rentabilidad del patrimonio total", E.p_rentabilidad_patrim_total),
    (r"^Al \d+ de \w+ de \d+, los activos totales se", E.p_activos_totales_intro),
    (r"^(Reducción|Incremento) en activos no corrientes en", E.p_activos_nocorrientes_bridge),
    (r"(?i)(Reducción|Incremento) en activos corrientes en", E.p_activos_corrientes_bridge),
    (r"^Los pasivos totales (se )?(redujeron|incrementaron)", E.p_pasivos_bridge),
    (r"^El patrimonio total de la Compañía", E.p_patrimonio_bridge),
]

# Preámbulo y títulos de sección: no dependen del Excel (solo del período), así que sus funciones
# toman `periods` directamente en vez de `wb`. Sin esto, estos párrafos quedarían con la fecha/mes del
# trimestre ANTERIOR (el del Word base) en vez del trimestre que se está generando — no traen cifras,
# pero si no se regeneran, el informe queda con el período equivocado en el título y los encabezados.
HEADER_TEMPLATES = [
    (r"^Al \d+ de \w+ de \d+\s*$", E.p_titulo_fecha),
    (r"^El presente análisis razonado ha sido preparado para el periodo terminado al", E.p_preambulo_1),
    (r"^Dado que la Compañía administra sus operaciones con una visión de temporada agrícola", E.p_preambulo_2),
    (r"^Análisis EBITDA acumulado a \w+", E.p_header_ebitda_acumulado),
    (r"^Análisis Resultado a \w+", E.p_header_resultado_calendario),
    (r"^Análisis Resultado temporada \w+", E.p_header_resultado_temporada),
    (r"^Análisis Ingresos acumulados a \w+", E.p_header_ingresos_acumulados),
]

def apply_header_templates(doc, periods, log):
    for anchor, fn in HEADER_TEMPLATES:
        idx, para = find_paragraph(doc, anchor)
        if para is None:
            log.append(f"[AVISO] No encontré párrafo para patrón: {anchor[:60]}")
            continue
        try:
            set_paragraph_text(para, fn(periods))
            log.append(f"[OK] Actualizado (encabezado/preámbulo): {anchor[:50]}")
        except Exception as ex:
            log.append(f"[ERROR] {anchor[:50]}: {ex}")

def collect_original_causes(doc, log):
    """Extrae, del Word ORIGINAL (antes de sobrescribir ningún párrafo), las frases de causa de
    negocio que ya vienen escritas, para preservarlas (resaltadas) en el documento regenerado en vez
    de generar una instrucción genérica. Debe llamarse ANTES de apply_single_templates y
    apply_bridge_bullets, mientras el documento todavía tiene el texto del trimestre anterior."""
    causes = {}
    anchor_by_fn = {fn: anchor for anchor, fn in SINGLE_TEMPLATES}

    def _text_for(fn):
        anchor = anchor_by_fn.get(fn)
        if not anchor:
            return None
        _, para = find_paragraph(doc, anchor)
        return para.text if para is not None else None

    t = _text_for(E.p_deterioro_6m)
    if t:
        cur, prior = _extract_deterioro_6m_causes(t)
        causes['deterioro_6m_cur'] = cur; causes['deterioro_6m_prior'] = prior
        if cur or prior:
            log.append("[OK] Causa original preservada: deterioro de activos 6M")

    t = _text_for(E.p_gastosadmin_6m)
    if t:
        cause = _extract_gastosadmin_6m_cause(t)
        if cause:
            causes['gastosadmin_6m'] = cause
            log.append("[OK] Causa original preservada: gastos de administración 6M")

    t = _text_for(E.p_costoventa_12m)
    if t:
        cause = _extract_costoventa_12m_cause(t)
        if cause:
            causes['costoventa_12m'] = cause
            log.append("[OK] Causa original preservada: costo de ventas 12M")

    t = _text_for(E.p_razon_endeudamiento)
    if t:
        cause = _extract_razon_endeudamiento_cause(t)
        if cause:
            causes['razon_endeudamiento'] = cause
            log.append("[OK] Causa original preservada: razón de endeudamiento")

    t = _text_for(E.p_cobertura_gf)
    if t:
        cause = _extract_cobertura_gf_cause(t)
        if cause:
            causes['cobertura_gf'] = cause
            log.append("[OK] Causa original preservada: cobertura de gastos financieros")

    # --- Bridge 12M (Ganancia Atribuible): factores con causa histórica real ---
    causes['bridge_12m'] = {}
    idx, intro = find_paragraph(doc, r"^La (ganancia|pérdida)(\s*\(pérdida\))? atribuible a los propietarios de la controladora (fue de|se registr[oó] en) US\$")
    if intro is not None:
        pos_paras, after_pos = _next_list_block(doc, idx + 1, 14)
        j = after_pos
        n = len(doc.paragraphs)
        while j < n and doc.paragraphs[j].text.strip() == '':
            j += 1
        sep_idx = j if (j < n and 'compensad' in doc.paragraphs[j].text.lower()) else None
        neg_paras = []
        if sep_idx is not None:
            neg_paras, _ = _next_list_block(doc, sep_idx + 1, 14)
        for p in (pos_paras + neg_paras):
            label = classify_bridge_label(p.text)
            cause = None
            if label == 'Deterioro':
                cause = _extract_bridge_deterioro_cause(p.text)
            elif label == 'Depreciación y amortizaciones':
                cause = _extract_bridge_depreciacion_cause(p.text)
            elif label == 'Costos Financieros Netos':
                cause = _extract_bridge_costosfin_cause(p.text)
            if label and cause:
                causes['bridge_12m'][label] = cause
        if causes['bridge_12m']:
            log.append(f"[OK] Causas originales preservadas en bridge 12M: {', '.join(causes['bridge_12m'].keys())}")
    return causes

# Funciones de SINGLE_TEMPLATES que necesitan una causa de negocio original (ver collect_original_causes).
# El resto se llama simplemente como fn(wb).
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

# Respaldo posicional: algunos trimestres redactan este párrafo sobre una métrica distinta (p.ej. el
# Dic25 original habla de "costos de ventas" en vez de "costos y gastos sin incluir depreciación"),
# así que el texto ya no es reconocible por patrón — pero su POSICIÓN sí es estable: siempre es el
# párrafo que sigue inmediatamente al de la viñeta de ingresos de esta misma sección. Si el ancla
# propia de `fn` no encuentra nada, se intenta ubicar por esta vía antes de darlo por no encontrado.
POSITIONAL_FALLBACK_AFTER = {
    E.p_ebitda_12m_bullet_costos: E.p_ebitda_12m_bullet_ingresos,
}

def apply_single_templates(doc, wb, periods, causes, log):
    matched_idx = {}
    for anchor, fn in SINGLE_TEMPLATES:
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
            log.append(f"[AVISO] No encontré párrafo para patrón: {anchor[:60]}")
            continue
        matched_idx[fn] = idx
        try:
            text = P.localize(_call_single_template(fn, wb, causes), periods)
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
    items_texts, clonando o eliminando párrafos según haga falta. Devuelve la lista final completa de
    párrafos (incluyendo los que se hayan clonado), en orden, para que quien llama pueda seguir
    insertando contenido justo a continuación sin perder la posición real en el documento."""
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

# ---------- Bloque de viñetas del bridge, tolerante a la estructura del documento BASE ----------
# El documento base (trimestre anterior) puede venir en cualquiera de las 4 variantes que Hortifrut
# ha usado realmente: lista plana sin separador (p.ej. el bridge de temporada de diciembre), lista
# partida en positivos/"compensados"/negativos con una frase separadora (p.ej. junio y septiembre), o
# lo que haya quedado de una ejecución previa de esta misma herramienta (que siempre deja el "estilo
# Jun26": partida con separador). _collect_bridge_span reconoce cualquiera de esas variantes como un
# solo tramo contiguo (viñetas + a lo más una frase separadora embebida), para no dejar viñetas viejas
# huérfanas sin reemplazar ni duplicar contenido.
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
        # Párrafo normal (no viñeta): se incluye como "separador embebido" solo si más adelante (tras
        # saltar blancos) viene otra viñeta — si no, ya salimos del bloque del bridge.
        k = j + 1
        while k < n and k < start + max_scan and doc.paragraphs[k].text.strip() == '':
            k += 1
        if k < n and k < start + max_scan and _is_list_para(doc.paragraphs[k]):
            span.append(p); j += 1
            continue
        break
    return span, j

def _replace_bridge_span_flat(doc, intro_anchor_regex, items_texts, max_scan=24):
    """Para el bridge de calendario (estilo Jun26: una sola lista plana, sin separador)."""
    idx, intro = find_paragraph(doc, intro_anchor_regex)
    if intro is None:
        return False
    span, _ = _collect_bridge_span(doc, idx + 1, max_scan)
    bullet_paras = [p for p in span if _is_list_para(p)]
    sep_paras = [p for p in span if not _is_list_para(p)]
    if not bullet_paras:
        return False
    _replace_list_items(bullet_paras, items_texts)
    for p in sep_paras:
        p._p.getparent().remove(p._p)
    return True

def _replace_bridge_span_split(doc, start, pos_texts, neg_texts, sep_sentence, intro_para, max_scan=24):
    """Para el bridge de temporada (estilo Jun26: positivos, frase separadora, negativos)."""
    span, _ = _collect_bridge_span(doc, start, max_scan)
    bullet_paras = [p for p in span if _is_list_para(p)]
    sep_paras = [p for p in span if not _is_list_para(p)]
    if not bullet_paras:
        return False

    # Nos quedamos con un único párrafo separador reutilizable (de estilo normal) si el documento base
    # ya traía uno (sea cual sea su posición original); si había más de uno (no debería pasar) nos
    # quedamos con el primero y eliminamos el resto.
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
        # no debería pasar (siempre hay al menos la viñeta de EBITDA), pero por seguridad no tocamos nada más
        return False
    # _replace_list_items devuelve la lista final real (con los párrafos clonados si hizo falta
    # extender), para ubicar el separador y los negativos exactamente a continuación sin perder la
    # posición en el documento aunque hayan faltado viñetas para los positivos.
    pos_paras_final = _replace_list_items(pos_paras, pos_texts)

    # El separador se ubica (o reubica, si el documento base lo traía en otra posición, p.ej. una
    # lista plana sin separador) siempre justo después del último positivo — nunca antes del primero.
    last_pos_el = pos_paras_final[-1]._p
    if reusable_sep is not None:
        sep_para = reusable_sep
        last_pos_el.addnext(sep_para._p)  # lxml reubica el elemento aunque ya esté en el árbol
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

def apply_bridge_bullets(doc, wb, periods, causes, log):
    # --- Calendario: una sola lista plana (estilo Jun26) ---
    # (los factores de este bridge no tienen causa histórica preservable: son viñetas nuevas cada
    # trimestre, no arrastran una frase de negocio puntual del Word anterior)
    # El punto de entrada al bloque varía según el trimestre del documento base: junio usa "Las
    # principales variaciones se explican a continuación"; septiembre/diciembre usan en cambio la
    # misma frase partida en positivos/negativos que junio usa para la sección de temporada
    # ("...se explica principalmente por los siguientes efectos negativos/positivos:") — cualquiera de
    # las dos formas sirve como ancla de entrada, porque _collect_bridge_span después junta todo el
    # tramo (ambos grupos y el separador, si lo hay) y lo reemplaza íntegro por la lista plana.
    # Aislado en try/except: un solo dato faltante en el Excel (p.ej. una hoja con estructura distinta
    # a la plantilla esperada) no debe impedir que se genere el resto del documento — igual que los
    # párrafos de apply_single_templates, que se procesan uno a uno.
    try:
        items_6m = build_bridge_items_6m(wb)
        ok = _replace_bridge_span_flat(
            doc,
            r"Las principales variaciones se explican a continuación|se explica principalmente por los siguientes efectos (negativos|positivos)",
            items_6m,
        )
        log.append(f"[{'OK' if ok else 'AVISO'}] Bridge calendario ({len(items_6m)} factores)")
    except Exception as ex:
        log.append(f"[ERROR] Bridge calendario: {ex}")

    # --- Temporada: dos listas (factores positivos / factores que compensan) separadas por una frase fija ---
    try:
        bridge_12m_causes = causes.get('bridge_12m', {})
        idx, intro = find_paragraph(doc, r"^La (ganancia|pérdida)(\s*\(pérdida\))? atribuible a los propietarios de la controladora (fue de|se registr[oó] en) US\$")
        if intro is None:
            log.append("[AVISO] No encontré la intro del bridge de Ganancia Atribuible de temporada")
            return
        items = E.bridge_bullets(wb, '12m', threshold_abs_mus=3.0)
        pos = sorted([it for it in items if it[3] >= 0], key=lambda x: abs(x[3]), reverse=True)
        neg = sorted([it for it in items if it[3] < 0], key=lambda x: abs(x[3]), reverse=True)
        pos_texts = [E.p_ebitda_bullet_12m(wb)] + [E.bridge_factor_text(lbl, cur, prior, delta, '12m', cause=bridge_12m_causes.get(lbl)) for lbl, cur, prior, delta in pos]
        neg_texts = [E.bridge_factor_text(lbl, cur, prior, delta, '12m', cause=bridge_12m_causes.get(lbl)) for lbl, cur, prior, delta in neg]

        sep_sentence = "Los anteriores efectos positivos, fueron parcialmente compensados por los siguientes efectos:"
        ok = _replace_bridge_span_split(doc, idx + 1, pos_texts, neg_texts, sep_sentence, intro)
        if ok:
            log.append(f"[OK] Bridge de temporada GA ({len(pos_texts)} factores positivos, {len(neg_texts)} compensatorios)")
        else:
            log.append("[AVISO] No encontré las viñetas del bridge de Ganancia Atribuible de temporada")
    except Exception as ex:
        log.append(f"[ERROR] Bridge de temporada GA: {ex}")

# ---------- Tablas / imágenes ----------
EMU_PER_INCH = 914400

def collect_and_render_tables(wb, word_path, log):
    """Decide qué imagen (image1.emf, image2.emf, ...) del Word ORIGINAL corresponde a cada tabla, y
    renderiza el PNG actualizado para cada una. Devuelve ({nombre_media_original: ruta_png_nueva}, img_dir)."""
    rt_dir = os.path.dirname(__file__)
    sys.path.insert(0, rt_dir)
    import render_tables_v5 as RT

    RT.TABLES['IngresosxSegmento']['sheet'] = 'IngresosxSegmento'
    if 'Indicadores1' not in wb.sheetnames and 'Ind. Financieros' in wb.sheetnames:
        RT.TABLES['Indicadores1']['sheet'] = 'Ind. Financieros'

    SKIP_KEYS = set()
    if 'Indicadores2' not in wb.sheetnames:
        # las 3 hojas se separaron con layouts distintos; no re-especificadas aún -> no tocar esas imágenes
        SKIP_KEYS.add('IndicadoresActividad')
        SKIP_KEYS.add('Rentabilidad')

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
        if tkey not in RT.TABLES: continue
        spec = RT.TABLES[tkey]
        if spec['sheet'] not in wb.sheetnames: continue
        a = c2a.get(img_name)
        if not a: continue
        if tkey in rendered:
            image_map[a] = rendered[tkey]; continue
        out_png = os.path.join(img_dir, f'{tkey}.png')
        try:
            RT.render_table(wb, spec, out_png, dpi=300)
            image_map[a] = out_png; rendered[tkey] = out_png
            log.append(f"[OK] Tabla renderizada: {tkey}")
        except Exception as ex:
            log.append(f"[ERROR] Tabla {tkey}: {ex}")
    return image_map, img_dir

def fix_image_frame_sizes(doc, image_map, log):
    """Ajusta el tamaño de cada imagen insertada en el Word para que use el ancho útil de la página
    (o su ancho natural si es menor) y una altura proporcional — sin esto, Word estira o aplasta la
    imagen nueva dentro del marco de tamaño fijo de la imagen anterior, y el texto se ve mucho más
    chico/apretado de lo que realmente se generó."""
    try:
        from PIL import Image as PILImage
    except ImportError:
        log.append("[AVISO] Pillow no disponible: no se pudieron ajustar los tamaños de las imágenes")
        return
    usable_w_emu = doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin
    fixed = 0
    for shape in doc.inline_shapes:
        try:
            blip = shape._inline.graphic.graphicData.pic.blipFill.blip
            rId = blip.embed
            target = os.path.basename(str(doc.part.rels[rId].target_part.partname))
        except Exception:
            continue
        new_png = image_map.get(target)
        if not new_png:
            continue
        try:
            with PILImage.open(new_png) as im:
                w_px, h_px = im.size
            new_w_emu = min(usable_w_emu, int(w_px / 300 * EMU_PER_INCH))  # dpi=300 al renderizar
            new_h_emu = int(new_w_emu * h_px / w_px)
            shape.width = new_w_emu
            shape.height = new_h_emu
            fixed += 1
        except Exception as ex:
            log.append(f"[AVISO] No pude ajustar el tamaño de {target}: {ex}")
    log.append(f"[OK] {fixed} imágenes ajustadas a su proporción y tamaño correctos")

def swap_media_bytes(docx_in_path, docx_out_path, image_map, log):
    """Reemplaza, dentro del .docx ya guardado, los bytes de cada imagen de tabla por su PNG actualizado."""
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(docx_in_path, 'r') as z: z.extractall(tmp)
    rp = os.path.join(tmp, 'word', '_rels', 'document.xml.rels')
    with open(rp, encoding='utf-8') as f: rels = f.read()
    repl = 0
    for a, new_png in image_map.items():
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
    shutil.rmtree(tmp)
    log.append(f"[OK] {repl} imágenes de tabla actualizadas")

def generate(excel_path, word_path, out_path):
    log = []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    periods = P.derive_periods(wb)
    log.append(f"[INFO] Período detectado: {periods['CUR_SHORT']} vs {periods['PRIOR_SHORT']} / {periods['FY_PRIOR_SHORT']}")

    doc = docx.Document(word_path)
    causes = collect_original_causes(doc, log)
    apply_header_templates(doc, periods, log)
    apply_single_templates(doc, wb, periods, causes, log)
    apply_bridge_bullets(doc, wb, periods, causes, log)

    image_map, img_dir = collect_and_render_tables(wb, word_path, log)
    fix_image_frame_sizes(doc, image_map, log)

    tmp_text_docx = word_path + '.text.docx'
    doc.save(tmp_text_docx)

    swap_media_bytes(tmp_text_docx, out_path, image_map, log)
    os.unlink(tmp_text_docx)
    shutil.rmtree(img_dir)
    return log

if __name__ == '__main__':
    excel_path, word_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    log = generate(excel_path, word_path, out_path)
    print('\n'.join(log))
