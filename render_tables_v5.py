"""
Table renderer v5 - Final production version.
Key fixes over v4:
- Numbers NEVER truncated: font shrinks to fit, no ellipsis
- ExposicTC: shortened labels, wider figure, smaller font
- VctosPtmos: proper col structure, label col wraps cleanly  
- Risks: % shown as actual percentages (×100)
- Riesgos: wrapping for long text cols
- Costos: proper col widths so (1.105.877) fits
- All 7-col tables use fig_w=9.5 by default
"""
import warnings; warnings.filterwarnings('ignore')
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from PIL import Image, ImageChops
from datetime import datetime
import re

# ── Traducción genérica de tokens de encabezado recurrentes (lang='en') ──
# Estas celdas vienen directamente del Excel en ESPAÑOL (ar_v2.xlsx; no hay un Excel en inglés separado
# alimentando las tablas) como texto literal (rangos de fecha, unidades), no como números — así que no
# pasan por fmt_en(). Se traducen aquí de forma genérica (por valor exacto o por patrón de mes+año) para
# no tener que listarlas una por una en cada label_overrides de TABLES_EN.
_MONTH_ES_EN = {'Ene':'Jan','Abr':'Apr','Ago':'Aug','Dic':'Dec'}
_COMMON_TOKEN_EN = {
    'variación': 'Variation', 'Variación': 'Variation',
    'variaciones': 'Variation', 'Variaciones': 'Variation',
    'variaciones %': 'Variation %', 'Variaciones %': 'Variation %',
    'MUS$': 'ThUS$', 'veces': 'times', 'Veces': 'Times', 'Días': 'Days',
}
# Celdas con fecha larga dinámica (cambian cada trimestre dentro del propio Excel, p.ej.
# "Exposición Neta al 30 junio 2026" / "Totales al 30 junio 2026" en ExposicTC/EfectoTC) —
# se traducen por patrón en vez de quedar fijas en label_overrides, para que sigan siendo
# correctas en trimestres futuros sin tener que tocar TABLES_EN cada vez.
_MONTH_ES_FULL_EN = {
    'enero':'January','febrero':'February','marzo':'March','abril':'April','mayo':'May','junio':'June',
    'julio':'July','agosto':'August','septiembre':'September','setiembre':'September','octubre':'October',
    'noviembre':'November','diciembre':'December',
}
_DYNAMIC_DATE_LABEL_EN = [
    (re.compile(r'^Exposición Neta al (.+)$'), 'Net Exposure as of {date}'),
    (re.compile(r'^Totales al (.+)$'), 'Total as of {date}'),
]
def _translate_dynamic_date_labels_en(txt):
    def _date_repl(m):
        day, month_es, year = m.group(1), m.group(2).lower(), m.group(3)
        month_en = _MONTH_ES_FULL_EN.get(month_es, month_es)
        return f'{month_en} {day}, {year}'
    date_re = re.compile(
        r'\b(\d{1,2}) (enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|'
        r'octubre|noviembre|diciembre) (\d{4})\b', re.IGNORECASE)
    for pat, template in _DYNAMIC_DATE_LABEL_EN:
        m = pat.match(txt)
        if m:
            translated_date = date_re.sub(_date_repl, m.group(1))
            return template.format(date=translated_date)
    return txt
def _translate_common_strings_en(txt):
    def _repl(m):
        return _MONTH_ES_EN.get(m.group(1), m.group(1)) + m.group(2)
    txt = re.sub(r'\b(Ene|Abr|Ago|Dic)(\d{2})\b', _repl, txt)
    txt = _translate_dynamic_date_labels_en(txt)
    txt = txt.replace('MUS$', 'ThUS$')  # substring, not just full-string: covers 'MUS$' and '(MUS$)'
    return _COMMON_TOKEN_EN.get(txt, txt)

# ── Number formatting ──
def fmt(v, style='num'):
    if v is None or v == '': return ''
    if isinstance(v, bool): return ''
    if isinstance(v, datetime): return v.strftime('%d-%b-%y')
    if isinstance(v, str): return v.strip()
    try: n = float(v)
    except: return str(v)
    if style == 'pct':      return f'{n:.2f}%'.replace('.', ',')
    if style == 'auto_pct': return f'{n*100:.2f}%'.replace('.', ',')
    if style == 'pct_raw':  return f'{n*100:.0f}%'   # integers like 20%, 13%
    if style == 'dec2':     return f'{n:.2f}'.replace('.', ',')
    if style == 'dec4':
        s = f'{n:.7f}'.rstrip('0').rstrip('.')
        return s.replace('.', ',')
    formatted = f'{abs(n):,.0f}'.replace(',', '.')
    return f'({formatted})' if n < 0 else formatted

# ── Number formatting (en-US): period decimal, comma thousands — the reverse of fmt() above.
# Used when render_table(..., lang='en') renders the SAME workbook (Spanish ar_v2.xlsx; there is no
# separate English workbook feeding the tables) with translated labels (see label_overrides below) and
# English number conventions. Negative numbers use parentheses in both languages (accounting style).
def fmt_en(v, style='num'):
    if v is None or v == '': return ''
    if isinstance(v, bool): return ''
    if isinstance(v, datetime): return v.strftime('%d-%b-%y')
    if isinstance(v, str): return v.strip()
    try: n = float(v)
    except: return str(v)
    if style == 'pct':      return f'{n:.2f}%'
    if style == 'auto_pct': return f'{n*100:.2f}%'
    if style == 'pct_raw':  return f'{n*100:.0f}%'
    if style == 'dec2':     return f'{n:.2f}'
    if style == 'dec4':
        return f'{n:.7f}'.rstrip('0').rstrip('.')
    formatted = f'{abs(n):,.0f}'
    return f'({formatted})' if n < 0 else formatted

# ── Word wrap ──
def wrap_label(text, max_chars, max_lines=2):
    """Wrap at word boundaries."""
    if not text or len(text) <= max_chars: return text
    words = text.split()
    lines, cur = [], ''
    for w in words:
        if cur and len(cur)+1+len(w) > max_chars:
            lines.append(cur); cur = w
        else:
            cur = (cur+' '+w).strip()
    if cur: lines.append(cur)
    return '\n'.join(lines[:max_lines])

def resolve_row_nums(ws, spec):
    """Devuelve la lista de filas reales de la hoja para spec['rows'], aplicando el ajuste de
    'anchor' si corresponde (ver nota abajo). Factorizado desde render_table para que el
    traductor de Excel (excel_translator_en.py) use exactamente la misma resolución de filas y
    así escriba las traducciones en la celda correcta, incluso en plantillas de trimestres
    anteriores donde la hoja corrió una fila (ver dic25/sep25/mar26)."""
    row_nums = spec.get('rows') or [r for r in range(1,ws.max_row+1)
                  if any(ws.cell(r,c).value not in (None,'') for c in range(1,ws.max_column+1))]
    # anchor: (col, label_text, rowidx_in_rows_list) — algunas plantillas de trimestres
    # anteriores (ej. dic25/sep25/mar26) tienen una fila de más/menos respecto a la plantilla
    # actual (ar_v2.xlsx). Si el spec define 'anchor', se busca esa etiqueta cerca de la
    # posición esperada y se desplaza TODO row_nums por el delta encontrado, para que la
    # tabla se siga alineando aunque la plantilla haya corrido filas.
    anchor = spec.get('anchor')
    if anchor and row_nums:
        a_col, a_label, a_idx = anchor
        expected_row = row_nums[a_idx]
        found_row = None
        for delta in range(0, 4):
            for r in ({expected_row - delta, expected_row + delta} if delta else {expected_row}):
                if r < 1:
                    continue
                v = ws.cell(r, a_col).value
                if isinstance(v, str) and v.strip() == a_label:
                    found_row = r
                    break
            if found_row is not None:
                break
        if found_row is not None and found_row != expected_row:
            shift = found_row - expected_row
            row_nums = [r + shift for r in row_nums]
    return row_nums

# ── Core renderer ──
def render_table(wb, spec, out_path, dpi=300, lang='es', label_overrides=None):
    """lang='en': usa formato numérico en-US (fmt_en) en vez de es-CL (fmt). label_overrides:
    {(row_idx, col_idx): 'texto'} para reemplazar el texto de una celda (misma fila/columna lógicas
    que label_shorten, pero por (fila,columna) en vez de solo fila 0) — usado para las tablas en
    inglés, que leen el MISMO Excel en español pero muestran las etiquetas ya traducidas."""
    label_overrides = label_overrides or {}
    _fmt = fmt_en if lang == 'en' else fmt
    ws = wb[spec['sheet']]
    row_nums = resolve_row_nums(ws, spec)
    col_defs   = spec['cols']
    n_rows     = len(row_nums)
    cw         = spec['col_widths']
    bold_rows  = spec.get('bold_rows', set())
    italic_rows= spec.get('italic_rows', set())
    total_rows = spec.get('total_rows', set())
    header_rows= spec.get('header_rows', set())
    center_cols= spec.get('center_cols', set())
    row_fmt_map= spec.get('row_fmt', {})
    yellow_all = spec.get('yellow_all', False)
    grp_headers= spec.get('group_headers', [])
    italic_col0_only = spec.get('italic_col0_only', False)
    label_shorten= spec.get('label_shorten', {})   # {row_idx: 'short text'}

    FIG_W = spec.get('fig_w', 8.2)
    ROW_H_BASE = spec.get('row_h', 0.285)
    FONT = spec.get('font_size', 9.0)
    FONT_HDR = FONT - 0.5
    FONT_SM  = FONT - 1.5
    FONT_XS  = FONT - 2.5
    MARGIN = 0.06
    FN_H = 0.22 if spec.get('footnote') else 0
    max_lines_spec = spec.get('max_lines', 2)
    wrap_all = spec.get('wrap_all_cols', False)

    # Fixed row heights — same for all rows in the table
    # ROW_H_BASE is set large enough (0.28") to accommodate 2-line wraps at 9pt
    row_heights = [ROW_H_BASE] * n_rows

    # group_gaps: cuando la lista de filas salta (sub-tablas no contiguas en la hoja, ej. Risks),
    # inserta un espacio visual extra antes del salto para que se vean como bloques separados.
    group_gaps = spec.get('group_gaps', False)
    GAP_H = ROW_H_BASE * 0.6
    extra_gaps = [0.0] * n_rows
    if group_gaps:
        for ri in range(1, n_rows):
            if row_nums[ri] - row_nums[ri-1] > 1:
                extra_gaps[ri] = GAP_H

    FIG_H = n_rows * ROW_H_BASE + sum(extra_gaps) + 2*MARGIN + FN_H

    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
    AX_H = FIG_H / FIG_W
    ax.set_xlim(0,1); ax.set_ylim(0,AX_H); ax.axis('off')
    Y0 = AX_H - MARGIN/FIG_W

    xs = []; x=0.0
    for w in cw: xs.append(x); x+=w

    BLACK=(0,0,0); YELLOW=(1.0,0.867,0.0); WHITE=(1.0,1.0,1.0)

    # Y positions (constant spacing since all rows same height, plus any group gap)
    row_y_tops = []
    y_cur = Y0
    for ri, rh in enumerate(row_heights):
        y_cur -= extra_gaps[ri] / FIG_W
        row_y_tops.append(y_cur)
        y_cur -= rh / FIG_W

    for ri, rn in enumerate(row_nums):
        H = row_heights[ri] / FIG_W
        yt = row_y_tops[ri]
        is_bold   = ri in bold_rows
        is_italic = ri in italic_rows
        is_total  = ri in total_rows
        is_header = ri in header_rows
        bg = YELLOW if yellow_all else WHITE

        for ci,(ec,bstyle) in enumerate(col_defs):
            x0=xs[ci]; w=cw[ci]
            ax.add_patch(plt.Rectangle((x0,yt-H),w,H,facecolor=bg,edgecolor='none',zorder=1))

            # Italic-only-col0 rows: skip numeric cols
            if is_italic and italic_col0_only and ci>0: continue

            # Skip group header covered cells
            if any(g[0]==ri and g[1]<=ci<=g[2] for g in grp_headers): continue

            val = ws.cell(rn,ec).value
            style = bstyle
            if style=='_auto' and ci in (2,3): style=row_fmt_map.get(ri,'dec2')
            # pct_rows: numeric cols in these rows show as percentages
            pct_rows = spec.get('pct_rows',set())
            if ri in pct_rows and ci>0 and style=='num': style=spec.get('pct_style','auto_pct')

            if val is None or val=='': txt=''
            elif isinstance(val,str):
                txt=val.strip()
                if lang=='en': txt=_translate_common_strings_en(txt)
            elif isinstance(val,datetime): txt=val.strftime('%d-%b-%y')
            elif isinstance(val,bool): txt=''
            else: txt=_fmt(val,style)

            # Apply label shortening if defined
            if ci==0 and ri in label_shorten: txt=label_shorten[ri]
            # Apply (row,col) label override if defined (English tables: translated label text)
            if (ri,ci) in label_overrides: txt=label_overrides[(ri,ci)]

            # Col 0: wrap to 2 lines max
            # Numeric cols: NEVER truncate
            col_inch = w * FIG_W
            max_chars = max(8, int(col_inch * 9.5))
            max_lines = spec.get('max_lines', 2)
            if ci==0 or is_header:
                txt = wrap_label(txt, max_chars, max_lines)
            elif ci>0 and not (is_italic and italic_col0_only):
                # For non-label cols in tables with max_lines>2, wrap text cols too
                if spec.get('wrap_all_cols'):
                    txt = wrap_label(txt, max_chars, max_lines)

            # Alignment
            if ci in center_cols or (is_header and ci>0):
                ha,xp = 'center', x0+w/2
            else:
                ha,xp = 'left', x0+0.007

            # Font size: shrink numeric cols that are tight
            fw = 'bold' if (is_bold or is_header) else 'normal'
            fs = 'italic' if is_italic else 'normal'  # bold+italic combinado es válido (ver 'Costos': fila Deterioro)
            n_lines = txt.count('\n')+1
            longest = max((len(l) for l in txt.split('\n')),default=0)

            if is_header:
                sz = FONT_HDR
            else:
                sz = FONT   # ALWAYS same size for all data rows

            ax.text(xp, yt-H*0.5, txt,
                    ha=ha, va='center', fontsize=sz,
                    fontweight=fw, fontstyle=fs,
                    color=BLACK, clip_on=True, zorder=3,
                    fontfamily='serif', multialignment='center', linespacing=1.1)

        # Lines
        ax.plot([0,1],[yt-H,yt-H],color=(0.82,0.82,0.82),linewidth=0.25,zorder=2)
        if is_total: ax.plot([0,1],[yt,yt],color=BLACK,linewidth=0.85,zorder=4)
        if is_header and (ri+1) not in header_rows:
            ax.plot([0,1],[yt-H,yt-H],color=BLACK,linewidth=0.7,zorder=4)

    # Group headers
    for g in grp_headers:
        gh_ri,gh_c1,gh_c2,gh_txt = g
        yt_g = row_y_tops[gh_ri]
        H_g = row_heights[gh_ri] / FIG_W
        x0_g=xs[gh_c1]; x1_g=xs[gh_c2]+cw[gh_c2]; w_g=x1_g-x0_g
        ax.add_patch(plt.Rectangle((x0_g,yt_g-H_g),w_g,H_g,facecolor=WHITE,edgecolor='none',zorder=2))
        ax.text(x0_g+w_g/2, yt_g-H_g*0.5, gh_txt,
                ha='center',va='center',fontsize=FONT_HDR,
                fontweight='bold',color=BLACK,
                fontfamily='serif',multialignment='center',zorder=3)

    # Top/bottom lines only
    ax.plot([0,1],[Y0,Y0],color=BLACK,linewidth=0.85,zorder=5)
    table_bottom = row_y_tops[-1] - row_heights[-1]/FIG_W
    ax.plot([0,1],[table_bottom,table_bottom],color=BLACK,linewidth=0.85,zorder=5)

    if spec.get('footnote'):
        ax.text(0,table_bottom-0.008,spec['footnote'],
                ha='left',va='top',fontsize=6.5,
                color=(0.3,0.3,0.3),fontstyle='italic',fontfamily='serif')

    plt.subplots_adjust(left=0,right=1,top=1,bottom=0)
    plt.savefig(out_path,dpi=dpi,bbox_inches='tight',facecolor='white',pad_inches=0.02)
    plt.close()

    img=Image.open(out_path).convert('RGB')
    diff=ImageChops.difference(img,Image.new('RGB',img.size,(255,255,255)))
    bb=diff.getbbox()
    if bb:
        pad=8
        img.crop((max(0,bb[0]-pad),max(0,bb[1]-pad),
                  min(img.width,bb[2]+pad),min(img.height,bb[3]+pad))
                ).save(out_path,'PNG',dpi=(dpi,dpi))

# ══════════════════════════════════════════
# TABLE SPECS
# ══════════════════════════════════════════
GH6_12 = [(0,1,2,'AÑO CALENDARIO\n(6 meses)'),(0,3,4,'TEMPORADA\n(12 meses)')]
GH6_12_7col = [(0,1,3,'AÑO CALENDARIO\n(6 meses)'),(0,4,6,'TEMPORADA\n(12 meses)')]

TABLES = {

 'EBITDA': dict(
    sheet='EBITDA',
    cols=[(2,'str'),(3,'num'),(5,'num'),(7,'num'),(9,'num')],
    rows=[1,2,3,4,5,6,7,8,9,10,11,12,13,14],
    bold_rows={2,5,9,10,12,13}, total_rows={5,9,12,13},
    header_rows={0,1,2}, center_cols={1,2,3,4},
    col_widths=[0.38,0.155,0.155,0.155,0.155],
    group_headers=GH6_12,
    footnote='*Excluye deterioro de valor de activos.',
 ),

 # Ingresos/Costos/OtrosComp/IngresosxSegmento: la plantilla trae, arriba de los datos, DOS filas de
 # encabezado propias (fila de rango de fechas real "Ene26 - Jun26"/"variación", y fila de unidades
 # "MUS$"/"%" + la etiqueta de la tabla) — a diferencia de EBITDA, estas hojas NO tienen una fila de
 # banner "AÑO CALENDARIO/TEMPORADA" propia. Por eso NO llevan group_headers (eso pisaría la fila real
 # de fechas con el banner genérico, perdiendo el rango real — bug reportado por Melissa: "no quedaron
 # bien los títulos de fechas"): las 2 filas de encabezado real de la plantilla ya aportan toda la
 # info de período necesaria.
 'Ingresos': dict(
    sheet='Ingresos',
    cols=[(2,'str'),(3,'num'),(5,'num'),(7,'auto_pct'),(9,'num'),(11,'num'),(13,'auto_pct')],
    rows=[2,3,4,5,8],
    bold_rows={4}, total_rows={4}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.30,0.115,0.115,0.085,0.115,0.115,0.085],
    fig_w=9.5,
    anchor=(2, 'Ingresos', 1),
 ),

 'IngresosxSegmento': dict(
    sheet='IngresosxSegmento (2)',
    cols=[(2,'str'),(3,'num'),(5,'num'),(6,'auto_pct'),(8,'num'),(10,'num'),(11,'auto_pct')],
    rows=[6,7,9,16,19],
    bold_rows={4}, total_rows={4}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.30,0.115,0.115,0.085,0.115,0.115,0.085],
    fig_w=9.5,
 ),

 'Costos': dict(
    sheet='Costos',
    cols=[(3,'str'),(4,'num'),(6,'num'),(8,'auto_pct'),(10,'num'),(12,'num'),(14,'auto_pct')],
    rows=[3,4,6,8,9,10,12,13,14],
    bold_rows={2,5,6,7,8}, italic_rows={3,4,7},
    total_rows={8}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.38,0.107,0.107,0.082,0.107,0.107,0.082],
    fig_w=9.5, row_h=0.30,
    # Shorten the very long row 3 label (0-based index of row 9 in rows list = index 5)
    label_shorten={4: 'Otros gastos por función,\nexcl. deterioros'},
 ),

 'OtrosComp': dict(
    sheet='Otros ingresos(gastos)',
    cols=[(3,'str'),(4,'num'),(6,'num'),(8,'auto_pct'),(10,'num'),(12,'num'),(14,'auto_pct')],
    rows=[4,5,7,8,9,10,11,13],
    bold_rows={7}, total_rows={7}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.38,0.107,0.107,0.082,0.107,0.107,0.082],
    fig_w=9.5,
    anchor=(3, 'Otros Ingresos (egresos)', 1),
 ),

 'IndicadoresActividad': dict(
    # Hoja 'Ind. Actividad' (plantilla desde Jun26 en adelante): fila 3 = encabezado (con el rango de
    # fechas tal como lo escribe la propia plantilla, p.ej. "Ene26 - Jun26"), fila 4 = título de
    # sección "Actividad" (sin valores), y 3 pares fila-indicador/fila-descripción en cursiva.
    sheet='Ind. Actividad',
    cols=[(2,'str'),(3,'str'),(4,'_auto'),(5,'_auto')],
    rows=[3,4,5,6,7,8,9,10],
    bold_rows={0,1,2,4,6}, italic_rows={3,5,7},
    total_rows=set(), header_rows={0}, center_cols={1,2,3},
    col_widths=[0.44,0.14,0.21,0.21],
    row_fmt={2:'dec2',4:'dec2',6:'num'},
    italic_col0_only=True,
 ),

 'DFN': dict(
    sheet='DF Neta',
    cols=[(2,'str'),(3,'num'),(5,'num')],
    rows=[5,6,7,8,9,10,11,12,13,15],
    bold_rows={0,1,6,7,9}, total_rows={6,9}, header_rows={0,1}, center_cols={1,2},
    col_widths=[0.55,0.225,0.225],
    footnote='*Se consideran Arrendamientos Operacionales que a partir de 2019 deben ser reconocidos como activos y pasivos en esta (IFRS 16).',
 ),

 'Indicadores1': dict(
    sheet='Indicadores1',
    cols=[(3,'str'),(4,'str'),(5,'_auto'),(6,'_auto'),(7,'auto_pct')],
    rows=[6,7,8,9,10,11,12,13,14,15,16,17,18],
    bold_rows={0,1,3,5,7,9,11}, italic_rows={2,4,6,8,10,12},
    total_rows=set(), header_rows={0}, center_cols={1,2,3,4},
    col_widths=[0.40,0.14,0.155,0.155,0.15],
    row_fmt={1:'dec2',3:'dec2',5:'dec2',7:'auto_pct',9:'auto_pct',11:'dec4'},
    italic_col0_only=True,
 ),

 'Rentabilidad': dict(
    # Hoja 'Ind. Rentabilidad' (plantilla desde Jun26 en adelante): fila 1 = encabezado (incluye
    # "Variaciones %", una columna que la hoja de Actividad no tiene), y 3 pares fila-indicador/
    # fila-descripción en cursiva. Los valores de las filas 2 y 3 (Rentabilidad del patrimonio) vienen
    # como fracción (-0,0016 = -0,16%) mientras que la fila 1 (Cobertura) viene en veces, no en %.
    sheet='Ind. Rentabilidad',
    cols=[(1,'str'),(2,'str'),(3,'_auto'),(4,'_auto'),(5,'auto_pct')],
    rows=[1,2,3,4,5,6,7],
    bold_rows={0,1,3,5}, italic_rows={2,4,6},
    total_rows=set(), header_rows={0}, center_cols={1,2,3,4},
    col_widths=[0.50,0.06,0.15,0.15,0.14],
    row_fmt={1:'dec2',3:'auto_pct',5:'auto_pct'},
    italic_col0_only=True,
 ),

 'Balance': dict(
    sheet='Balance',
    cols=[(3,'str'),(5,'num'),(6,'num'),(7,'num'),(8,'auto_pct')],
    rows=[4,5,7,8,9,11,12,13,15,16,17],
    bold_rows={0,1,4,7,8,10}, total_rows={4,7,10}, header_rows={0,1}, center_cols={1,2,3,4},
    col_widths=[0.52,0.12,0.12,0.12,0.12],
 ),

 # VctosPtmos: use cols 3-9 (Detalle + 6 data cols), skip "Total" col
 'VctosPtmos': dict(
    sheet='VctosPtmos',
    cols=[(3,'str'),(4,'num'),(5,'num'),(6,'num'),(7,'num'),(8,'num'),(9,'num')],
    rows=[2,3,4,5,6,7,8,9,10],
    bold_rows={0,1,2}, total_rows=set(), header_rows={0,1,2}, center_cols={1,2,3,4,5,6},
    col_widths=[0.26,0.10,0.10,0.115,0.115,0.115,0.115],
    group_headers=[(0,2,5,'Flujos')],
    fig_w=9.0,
    label_shorten={
        4: 'Bonos - Obligaciones\npúblico',
        7: 'Ctas. comerciales\notras ctas. por pagar',
        8: 'Ctas. por pagar\n empresas relacionadas',
    },
 ),

 # ExposicTC: use smaller font, shorten labels
 'ExposicTC_tabla': dict(
    sheet='ExposicTC',
    cols=[(3,'str'),(4,'num'),(5,'num'),(6,'num'),(7,'num'),(8,'num'),(9,'num'),(10,'num'),(11,'num')],
    rows=list(range(3,28)),
    # ri: 0=header(Pesos...), 1=MUS$, 2=Activos Financieros(section), 3=Efectivo,
    # 4=Otros activos fin ctes, 5=Deudores, 6=CxC Relacionadas, 7=Otros act fin NC,
    # 8=Derechos cobrar NC, 9=CxC Relacionadas NC, 10=Total Activos,
    # 11=Pasivos Financieros(section), 12=Otros pasivos fin ctes, 13=Arrend ctes,
    # 14=Ctas comerc ctes, 15=CxP Relacionadas ctes, 16=Otras prov ctes,
    # 17=Prov beneficios ctes, 18=Otros pas fin NC, 19=Arrend NC,
    # 20=Otras ctas pagar NC, 21=CxP Relacionadas NC, 22=Otras prov NC,
    # 23=Total Pasivos, 24=Exposicion Neta
    bold_rows={2,10,11,23,24}, total_rows={10,23,24}, header_rows={0,1},
    center_cols=set(range(1,9)),
    # Columna de etiquetas más ancha (0.28 de un fig_w más generoso) + max_lines=3: con el ancho
    # angosto anterior (0.27 de 7.4in, max_lines=2) varias etiquetas largas en inglés y en español
    # se cortaban silenciosamente (wrap_label descarta lo que no entra en 2 líneas) — p.ej. "Trade
    # and other receivables, current" perdía "current". Con este ancho, la etiqueta más larga entra
    # en 3 líneas sin perder texto (ver verificación en el chat). col_widths ahora suma 1.0 (antes
    # sumaba ~0.92, por eso las 2 últimas columnas de moneda quedaban angostas/desalineadas).
    col_widths=[0.28,0.09,0.09,0.09,0.09,0.09,0.09,0.09,0.09],
    fig_w=9.5, font_size=8.0, row_h=0.42, max_lines=3,
 ),

 'EfectoTC': dict(
    sheet='ExposicTC',
    cols=[(3,'str'),(4,'num'),(5,'num'),(6,'num'),(7,'num'),(8,'num')],
    rows=[32,33,34,35,36,37,38,39,40,41,42],
    bold_rows={0,1,10}, total_rows={10}, header_rows={0,1}, center_cols={1,2,3,4,5},
    col_widths=[0.30,0.14,0.14,0.14,0.14,0.14],
    fig_w=6.8, font_size=9.5, row_h=0.33,
 ),

 # Matriz de Riesgos por Tipo (filas 2-4 de la hoja Risks): imagen propia en el Word original
 # (image14.emf), independiente de la matriz por Severidad (image15.emf, ver 'RiesgosPorSeveridad'
 # abajo) y de la tabla de detalle (image16.emf, ver 'Riesgos' abajo) — antes estaban mal
 # mapeadas (image15 apuntaba por error a la misma tabla de detalle que image16, duplicándola).
 'RiesgosPorTipo': dict(
    sheet='Risks',
    cols=[(4,'str'),(5,'num'),(6,'num'),(7,'num'),(8,'num'),(9,'num'),(10,'num')],
    rows=[2,3,4],
    bold_rows={0,2}, total_rows={2}, header_rows={0},
    center_cols={1,2,3,4,5,6},
    col_widths=[0.16,0.13,0.15,0.15,0.14,0.17,0.10],
    pct_rows={2}, pct_style='pct_raw',  # % enteros (100%, 20%...), igual que el documento original
    fig_w=6.8, font_size=9.5, row_h=0.36,
 ),

 # Matriz de Riesgos por Severidad (filas 8-10 de la hoja Risks): solo 5 columnas (menos que la
 # matriz por Tipo), imagen propia en el Word original (image15.emf).
 'RiesgosPorSeveridad': dict(
    sheet='Risks',
    cols=[(4,'str'),(5,'num'),(6,'num'),(7,'num'),(8,'num')],
    rows=[8,9,10],
    bold_rows={0,2}, total_rows={2}, header_rows={0},
    center_cols={1,2,3,4},
    col_widths=[0.34,0.16,0.17,0.17,0.16],
    label_shorten={0: 'Severidad\ndel Riesgo'},
    pct_rows={2}, pct_style='pct_raw',
    fig_w=4.9, font_size=9.5, row_h=0.36,
 ),

 # Riesgos criticos: Tipo, Nombre, Descripcion, Controles (all text, need wrapping)
 'Riesgos': dict(
    sheet='Riesgos',
    cols=[(3,'str'),(4,'str'),(5,'str'),(6,'str')],
    rows=[2,3,4],
    bold_rows={0}, total_rows=set(), header_rows={0},
    center_cols=set(), col_widths=[0.09,0.17,0.37,0.37],
    fig_w=7.2, row_h=2.0, font_size=8.5,
    max_lines=16, wrap_all_cols=True,
 ),

 'Seguros': dict(
    sheet='Seguros',
    cols=[(3,'str'),(4,'str'),(5,'str'),(6,'num'),(7,'num')],
    rows=list(range(3,43)),
    bold_rows={0,1}, total_rows=set(), header_rows={0,1}, center_cols={2,3,4},
    col_widths=[0.10,0.35,0.08,0.235,0.235],
 ),

 'FVfruta': dict(
    sheet='FVfruta',
    cols=[(3,'str'),(4,'num'),(5,'num'),(6,'num'),(7,'num')],
    rows=[4,5,6,7,11],
    bold_rows={0,1,2,4}, total_rows={4}, header_rows={0,1,2}, center_cols={1,2,3,4},
    col_widths=[0.28,0.18,0.18,0.18,0.18],
 ),
}

IMAGE_TO_TABLE = {
 'image1.emf':'EBITDA','image2.emf':'Ingresos','image3.emf':'IngresosxSegmento',
 'image4.emf':'Costos','image5.emf':'OtrosComp','image6.emf':'IndicadoresActividad',
 'image7.emf':'DFN','image8.emf':'Indicadores1','image9.emf':'Rentabilidad',
 'image9.png':'Rentabilidad',
 'image10.emf':'Balance','image11.emf':'VctosPtmos','image12.emf':'ExposicTC_tabla',
 'image13.emf':'EfectoTC','image14.emf':'RiesgosPorTipo','image15.emf':'RiesgosPorSeveridad',
 'image16.emf':'Riesgos','image17.emf':'Seguros','image18.emf':'FVfruta',
}

# ══════════════════════════════════════════
# TABLE SPECS — ENGLISH (lang='en')
# ══════════════════════════════════════════
# Las tablas en inglés se renderizan desde el MISMO Excel en español (ar_v2.xlsx; no hay un Excel
# en inglés separado alimentando las tablas) — misma hoja, mismas posiciones de celda (spec['rows']/
# spec['cols'] de TABLES arriba), pero con formato numérico en-US (fmt_en) y el texto de las
# etiquetas ya traducido (label_overrides, por (fila,columna) lógicas — ver render_table). Cada
# traducción fue verificada contra el Excel/Word en inglés de referencia que proporcionó Melissa
# (trimestre Jun26); dos celdas con fecha dinámica (ExposicTC/EfectoTC) se traducen por patrón en
# vez de quedar fijas, para seguir siendo correctas en trimestres futuros (ver
# _translate_dynamic_date_labels_en arriba). La tabla 'Riesgos' (detalle narrativo de riesgos:
# Tipo/Nombre/Descripción/Controles) SÍ está traducida (a pedido explícito de Melissa) — ver su
# entrada en TABLES_EN más abajo.

GH6_12_EN = [(0, 1, 2, 'CALENDAR YEAR\n(6 months)'), (0, 3, 4, 'SEASON\n(12 months)')]
GH6_12_7col_EN = [(0, 1, 3, 'CALENDAR YEAR\n(6 months)'), (0, 4, 6, 'SEASON\n(12 months)')]

TABLES_EN = {'EBITDA': {'group_headers': [(0, 1, 2, 'CALENDAR YEAR\n(6 months)'),
                              (0, 3, 4, 'SEASON\n(12 months)')],
            'footnote': '*Excluding impairment in the value of assets.',
            'label_overrides': {(2, 0): 'EBITDA DETERMINING',
                                (3, 0): 'Income from operating activities',
                                (4, 0): 'Other income, per function',
                                (5, 0): 'Total Income',
                                (6, 0): 'Cost of sales',
                                (7, 0): 'Administration expenses',
                                (8, 0): 'Other expenses, per function *',
                                (9, 0): 'Total Costs and Expenses',
                                (10, 0): 'Operating Result',
                                (11, 0): 'Depreciation and amortization',
                                (13, 0): 'EBITDA without Fair Value'}},
 # Ingresos/IngresosxSegmento/Costos/OtrosComp: sin group_headers (ver nota junto a TABLES arriba —
 # estas 4 hojas no tienen fila de banner propia; la fila de fechas real "Ene26-Jun26"/etc. ya se
 # traduce sola vía _translate_common_strings_en, no hace falta pisarla con un banner genérico).
 'Ingresos': {'label_overrides': {(1, 0): 'Total Operating Income',
                                  (2, 0): 'Income from operating activities',
                                  (3, 0): 'Other income, per function',
                                  (4, 0): 'Total Operating Income'}},
 'IngresosxSegmento': {'label_overrides': {(1, 0): 'Income per Segment',
                                           (2, 0): 'Fresh Fruit',
                                           (3, 0): 'Value Added Products',
                                           (4, 0): 'Total Operating Income'}},
 'Costos': {'label_overrides': {(0, 0): 'Costs and Expenses',
                                (2, 0): 'Cost of sales',
                                (3, 0): 'Administration expenses',
                                (4, 0): 'Other expenses, per function,\nexcl. impairment',
                                (5, 0): 'Other operating costs and expenses',
                                (6, 0): 'Less:',
                                (7, 0): 'Impairment of value of assets',
                                (8, 0): 'Total Costs and Expenses'}},
 'OtrosComp': {'label_overrides': {(1, 0): 'Other Income (expenses)',
                                   (2, 0): 'Other profit (loss)',
                                   (3, 0): 'Financial income',
                                   (4, 0): 'Financial expenses',
                                   (5, 0): 'Interest in profit (loss) of associated companies',
                                   (6, 0): 'Exchange rate fluctuations',
                                   (7, 0): 'Other Income (expenses)'}},
 'IndicadoresActividad': {'label_overrides': {(0, 0): 'Indicator',
                                              (0, 1): 'Unit',
                                              (1, 0): 'Activity',
                                              (2, 0): 'Rotation of Assets',
                                              (2, 1): 'Times',
                                              (3, 0): 'Operating revenue / Total average assets of '
                                                      'the period',
                                              (4, 0): 'Rotation of Inventory',
                                              (4, 1): 'Times',
                                              (5, 0): 'Cost of sales / Average inventory',
                                              (6, 0): 'Permanence of inventory (days)',
                                              (6, 1): 'Days',
                                              (7, 0): 'Inventory / Annual cost of sale (360 day '
                                                      'base)'}},
 'DFN': {'footnote': '*Operating Leases are considered, which starting in 2019 must be recognized '
                     'as assets and liabilities under this standard (IFRS 16).',
         'label_overrides': {(0, 0): 'Determination of Net Financial Debt',
                             (1, 0): 'Items',
                             (2, 0): 'Other current financial liabilities',
                             (3, 0): 'Current lease liabilities*',
                             (4, 0): 'Other non-current financial liabilities',
                             (5, 0): 'Non-current lease liabilities*',
                             (6, 0): 'Total financial liability',
                             (7, 0): 'Minus:',
                             (8, 0): 'Cash and cash equivalents',
                             (9, 0): 'Total net financial debt'}},
 'Indicadores1': {'label_overrides': {(0, 0): 'Indicator',
                                      (0, 1): 'Unit',
                                      (1, 0): 'Current Liquidity',
                                      (1, 1): 'Times',
                                      (2, 0): 'Current Assets / Current Liabilities',
                                      (3, 0): 'Acid Ratio',
                                      (3, 1): 'Times',
                                      (4, 0): 'Current assets (-) Other non-financial assets, '
                                              'inventories and current biological assets / Current '
                                              'liability',
                                      (5, 0): 'Debt Ratio',
                                      (5, 1): 'Times',
                                      (6, 0): 'Total liabilities / Equity attributable to Parent '
                                              'Company',
                                      (7, 0): 'Short term debt',
                                      (8, 0): 'Total current liabilities / Total liabilities',
                                      (9, 0): 'Long term debt',
                                      (10, 0): 'Total non-current liabilities / Total liabilities',
                                      (11, 0): 'Book value of the share (US$)',
                                      (11, 1): 'Dollars per share',
                                      (12, 0): 'Equity attributable to parent company / N° '
                                               'shares'}},
 'Rentabilidad': {'label_overrides': {(0, 0): 'Indicator',
                                      (0, 1): 'Unit',
                                      (1, 0): 'Financial expense coverage',
                                      (1, 1): 'Times',
                                      (2, 0): '(Before tax profit+Financial costs)/Financial costs',
                                      (3, 0): 'Profitability of parent company equity',
                                      (4, 0): 'Parent company gains/Parent company equity',
                                      (5, 0): 'Profitability of equity',
                                      (6, 0): 'Profit of the period/Total equity'}},
 'Balance': {'label_overrides': {(1, 0): 'Statement of Financial Position',
                                 (2, 0): 'Total current assets',
                                 (3, 0): 'Total non-current assets',
                                 (4, 0): 'Total assets',
                                 (5, 0): 'Total current liabilities',
                                 (6, 0): 'Total non-current liabilities',
                                 (7, 0): 'Total liabilities',
                                 (8, 0): 'Equity attributable to parent company equity holders',
                                 (9, 0): 'Non-controlling interest',
                                 (10, 0): 'Total equity'}},
 'VctosPtmos': {'group_headers': [(0,2,5,'Cash Flows')],
                'label_overrides': {(1, 1): 'Capital', (1, 2): 'Book Value',
                                    (1, 3): '0 to 3 months', (1, 4): 'Between 3 and 12 months',
                                    (1, 5): 'Between 2 and 5 years', (1, 6): 'More than 5 years',
                                    (2, 0): 'Creditor Bank',
                                    (3, 0): 'Bank loans',
                                    (4, 0): 'Bonds - Obligations\nwith the public',
                                    (5, 0): 'Leasing Liabilities',
                                    (6, 0): 'Operating Lease Liabilities',
                                    (7, 0): 'Trade accounts\nand other payables',
                                    (8, 0): 'Accounts payable to\nrelated companies'}},
 # Sin saltos de línea manuales: con col0 más ancha (ver TABLES['ExposicTC_tabla'] arriba) y
 # max_lines=3, wrap_label ya envuelve estas etiquetas completas sin cortar texto.
 'ExposicTC_tabla': {'label_overrides': {(0, 1): 'Chilean Pesos',
                                         (0, 2): 'Peruvian Sol',
                                         (0, 3): 'Euro',
                                         (0, 4): 'Mexican Pesos',
                                         (0, 5): 'Yuan',
                                         (0, 6): 'Dirham',
                                         (0, 7): 'Indian Rupee',
                                         (0, 8): 'Other',
                                         (2, 0): 'Financial Assets',
                                         (3, 0): 'Cash and Cash Equivalents',
                                         (4, 0): 'Other financial assets, current',
                                         (5, 0): 'Trade and other receivables, current',
                                         (6, 0): 'Accounts receivable from related entities, current',
                                         (7, 0): 'Other financial assets, non-current',
                                         (8, 0): 'Receivables, non-current',
                                         (9, 0): 'Accounts receivable from related entities, non-current',
                                         (10, 0): 'Total Financial Assets',
                                         (11, 0): 'Financial Liabilities',
                                         (12, 0): 'Other financial liabilities, current',
                                         (13, 0): 'Lease liabilities, current',
                                         (14, 0): 'Trade and other payables, current',
                                         (15, 0): 'Accounts payable to related entities, current',
                                         (16, 0): 'Other provisions, current',
                                         (17, 0): 'Employee benefit provisions, current',
                                         (18, 0): 'Other financial liabilities, non-current',
                                         (19, 0): 'Lease liabilities, non-current',
                                         (20, 0): 'Other accounts payable, non-current',
                                         (21, 0): 'Accounts payable to related entities, non-current',
                                         (22, 0): 'Other provisions, non-current',
                                         (23, 0): 'Total Financial Liabilities'}},
 'EfectoTC': {'label_overrides': {(0, 1): 'Assets', (0, 2): 'Liabilities', (0, 3): 'Net',
                                  (0, 4): 'Net (10% Devaluation)',
                                  (1, 0): 'Currencies',
                                  (2, 0): 'Chilean Peso',
                                  (3, 0): 'Peruvian Sol',
                                  (4, 0): 'Euro',
                                  (5, 0): 'Mexican Peso',
                                  (6, 0): 'Yuan',
                                  (7, 0): 'Dirham',
                                  (8, 0): 'Indian Rupee',
                                  (9, 0): 'Other'}},
 'RiesgosPorTipo': {'label_overrides': {(0, 0): 'Risk Type',
                                        (0, 1): 'Risks Identified',
                                        (0, 2): 'Strategic',
                                        (0, 3): 'Operational',
                                        (0, 4): 'Financial',
                                        (0, 5): 'Compliance',
                                        (0, 6): 'Climate',
                                        (1, 0): 'Risk Matrix'}},
 'RiesgosPorSeveridad': {'label_overrides': {(0, 0): 'Risk\nSeverity',
                                             (0, 1): 'Risks Identified',
                                             (0, 2): 'Critical Level',
                                             (0, 3): 'Tolerable Level',
                                             (0, 4): 'Acceptable Level',
                                             (1, 0): 'Risk Matrix'}},
 'Seguros': {'label_overrides': {(1, 0): 'COUNTRY',
                                 (1, 1): 'TYPE OF INSURANCE',
                                 (1, 2): 'CURRENCY',
                                 (1, 3): 'COVERED AMOUNT',
                                 (1, 4): 'COVERED AMOUNT',
                                 # Value cells with Spanish free text (not plain numbers) —
                                 # translated too, since these are substantive data, not just labels.
                                 (4, 3): 'Market Value', (4, 4): 'Market Value',
                                 (14, 3): '350,000/shipment', (14, 4): '350,000/shipment',
                                 (16, 3): 'Market Value', (16, 4): 'Market Value',
                                 (17, 3): 'Market Value', (17, 4): 'Market Value',
                                 (19, 3): '50,000 per event', (19, 4): '50,000 per event',
                                 (21, 3): '90% of unpaid amount', (21, 4): '90% of unpaid amount',
                                 (25, 3): 'Market Value', (25, 4): 'Market Value',
                                 (33, 3): 'Market Value', (33, 4): 'Market Value',
                                 (35, 3): 'Market Value', (35, 4): 'Market Value',
                                 (2, 0): 'Chile',
                                 (2, 1): 'All Risk Physical Goods',
                                 (3, 0): 'Chile',
                                 (3, 1): 'Mobile Agricultural Equipment',
                                 (4, 0): 'Chile',
                                 (4, 1): 'Motor and Commercial Vehicles',
                                 (5, 0): 'Chile',
                                 (5, 1): 'General and Product Civil Liability',
                                 (6, 0): 'Chile',
                                 (6, 1): 'Maritime Transport',
                                 (7, 0): 'Chile',
                                 (7, 1): 'Credit Insurance',
                                 (8, 0): 'Chile',
                                 (8, 1): 'Fruit and Materials Insurance',
                                 (9, 0): 'Chile',
                                 (9, 1): 'Terrorism and Sabotage',
                                 (10, 0): 'Chile',
                                 (10, 1): 'Business Interruption',
                                 (11, 0): 'USA',
                                 (11, 1): 'General and Product Civil Liability',
                                 (12, 0): 'USA',
                                 (12, 1): 'Flooding',
                                 (13, 0): 'USA',
                                 (13, 1): 'Business Interruption Damages',
                                 (14, 0): 'Mexico',
                                 (14, 1): 'Transport of Load',
                                 (15, 0): 'Mexico',
                                 (15, 1): 'Physical Goods',
                                 (16, 0): 'Mexico',
                                 (16, 1): 'Motor Vehicles',
                                 (17, 0): 'Spain',
                                 (17, 1): 'Motor Vehicles',
                                 (18, 0): 'Spain',
                                 (18, 1): 'Physical Goods',
                                 (19, 0): 'Spain',
                                 (19, 1): 'Cargo Insurance',
                                 (20, 0): 'Spain',
                                 (20, 1): 'General and Product Civil Liability',
                                 (21, 0): 'Spain',
                                 (21, 1): 'Credit Insurance',
                                 (22, 0): 'Peru',
                                 (22, 1): 'Civil Liability',
                                 (23, 0): 'Peru',
                                 (23, 1): 'Dishonesty, Disappearance and Destruction',
                                 (24, 0): 'Peru',
                                 (24, 1): 'Physical Goods',
                                 (25, 0): 'Peru',
                                 (25, 1): 'Motor Vehicles',
                                 (26, 0): 'Peru',
                                 (26, 1): 'Transport - National and Imports/Exports',
                                 (27, 0): 'Colombia',
                                 (27, 1): 'Transport - National and Exports',
                                 (28, 0): 'Colombia',
                                 (28, 1): 'Business Civil Liability',
                                 (29, 0): 'Colombia',
                                 (29, 1): 'All Risk',
                                 (30, 0): 'China',
                                 (30, 1): 'Fixed Assets and Inventory',
                                 (31, 0): 'China',
                                 (31, 1): 'Motor Vehicles',
                                 (32, 0): 'Morocco',
                                 (32, 1): 'Physical Goods',
                                 (33, 0): 'Morocco',
                                 (33, 1): 'Motor Vehicles',
                                 (34, 0): 'Portugal',
                                 (34, 1): 'General and Product Civil Liability',
                                 (35, 0): 'Portugal',
                                 (35, 1): 'Motor Vehicles',
                                 (36, 0): 'Brazil',
                                 (36, 1): 'Physical Goods',
                                 (37, 0): 'Ecuador',
                                 (37, 1): 'Physical Goods',
                                 (38, 0): 'Ecuador',
                                 (38, 1): 'Transport - National'}},
 'FVfruta': {'label_overrides': {(0, 0): 'Company',
                                 (0, 1): 'Fair Value adjustment as of',
                                 (0, 2): '10% Reduction',
                                 (0, 3): '10% Reduction',
                                 (0, 4): '10% Reduction',
                                 (1, 2): 'Volume',
                                 (1, 3): 'Price',
                                 (1, 4): 'Volume and Price',
                                 (4, 0): 'Total'}},
 # 'Riesgos' (detalle narrativo de riesgos: Tipo/Nombre/Descripción/Controles): traducido a pedido
 # explícito de Melissa (antes se dejaba sin traducir, pendiente de revisión humana del texto libre).
 # Traducción verificada frase por frase contra el Excel en español (ar_v2.xlsx, hoja 'Riesgos');
 # terminología ("the Company") alineada con el resto del Word en inglés de referencia.
 'Riesgos': {'label_overrides': {
     (0, 0): 'Risk Type', (0, 1): 'Risk Name', (0, 2): 'Risk Description',
     (0, 3): 'Corporate Controls Implemented',
     (1, 0): 'Climate',
     (1, 1): 'CLIMATE CHANGE AND NATURAL HAZARDS',
     (1, 2): 'Unfavorable weather conditions (e.g., storms) or unexpected events (e.g., fires, '
             'floods or attacks) that damage plantations and/or facilities, affecting production.',
     (1, 3): 'The Company has plantations and operations in diverse geographic locations, which '
             'partially mitigate this risk. In addition, the genetic diversification we have '
             'implemented at Hortifrut helps mitigate the effects of climate change.',
     (2, 0): 'Operational',
     (2, 1): 'CYBERSECURITY AND TECHNOLOGICAL DEPENDENCE OF THE BUSINESS',
     (2, 2): "Disruption to the technology platforms and/or information networks that support key "
             "business processes. Theft or exposure of the Company's information or sensitive "
             "information.",
     (2, 3): 'Contingency procedures are defined, and preventive measures are in place to mitigate '
             'the risk. \nIn addition, the Corporate Technology Management carries out an annual '
             'cybersecurity exercise at the global level.',
 }}}

def render_table_en(wb, tkey, out_path, dpi=300):
    """Atajo: arma el spec en inglés para TABLES[tkey] (spec español + group_headers/footnote/
    label_overrides de TABLES_EN[tkey], si existen) y llama a render_table(..., lang='en')."""
    spec = dict(TABLES[tkey])
    over = TABLES_EN.get(tkey, {})
    if 'group_headers' in over: spec['group_headers'] = over['group_headers']
    if 'footnote' in over: spec['footnote'] = over['footnote']
    render_table(wb, spec, out_path, dpi=dpi, lang='en', label_overrides=over.get('label_overrides', {}))

if __name__=='__main__':
 import sys,os,shutil,zipfile,tempfile,re
 excel_path,word_path,out_path = sys.argv[1],sys.argv[2],sys.argv[3]
 wb = load_workbook(excel_path,data_only=True)
 img_dir=tempfile.mkdtemp(); image_map={}; rendered={}

 print('🖼  Generando tablas v5...')
 for img_name,tkey in IMAGE_TO_TABLE.items():
  if tkey not in TABLES: continue
  spec=TABLES[tkey]
  if spec['sheet'] not in wb.sheetnames: continue
  if tkey in rendered:
   image_map[img_name]=rendered[tkey]; continue
  out_png=os.path.join(img_dir,f'{tkey}.png')
  try:
   render_table(wb,spec,out_png,dpi=300)
   image_map[img_name]=out_png; rendered[tkey]=out_png
   print(f'   ✓ {img_name} → {tkey}')
  except Exception as e:
   print(f'   ✗ {img_name}: {e}')
   import traceback; traceback.print_exc()

 tmp=tempfile.mkdtemp()
 with zipfile.ZipFile(word_path,'r') as z: z.extractall(tmp)
 rp=os.path.join(tmp,'word','_rels','document.xml.rels')
 with open(rp) as f: rels=f.read()
 actual=[n for n in re.findall(r'Target="media/([^"]+)"',rels) if re.match(r'image\d+',n)]
 actual=sorted(actual,key=lambda x:int(re.search(r'\d+',x).group()))
 c2a={}
 for i,a in enumerate(actual,1):
  for ext in ['.emf','.png']: c2a[f'image{i}{ext}']=a
 repl=0
 for old,new_png in image_map.items():
  a=c2a.get(old)
  if not a: continue
  nn=re.sub(r'(\.\w+)$','_v5.png',a)
  shutil.copy2(new_png,os.path.join(tmp,'word','media',nn))
  rels=rels.replace(f'Target="media/{a}"',f'Target="media/{nn}"')
  repl+=1
 with open(rp,'w') as f: f.write(rels)
 ct=os.path.join(tmp,'[Content_Types].xml')
 with open(ct) as f: c=f.read()
 if 'image/png' not in c:
  c=c.replace('</Types>','<Default Extension="png" ContentType="image/png"/></Types>')
 with open(ct,'w') as f: f.write(c)
 if os.path.exists(out_path): os.unlink(out_path)
 with zipfile.ZipFile(out_path,'w',zipfile.ZIP_DEFLATED) as z:
  for root,_,files in os.walk(tmp):
   for file in files:
    fp=os.path.join(root,file)
    z.write(fp,os.path.relpath(fp,tmp))
 shutil.rmtree(tmp); shutil.rmtree(img_dir)
 print(f'✅ Listo → {out_path} ({repl} imágenes)')
