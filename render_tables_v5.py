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

# ── Core renderer ──
def render_table(wb, spec, out_path, dpi=300):
    ws = wb[spec['sheet']]
    row_nums = spec.get('rows') or [r for r in range(1,ws.max_row+1)
                  if any(ws.cell(r,c).value not in (None,'') for c in range(1,ws.max_column+1))]
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
            if ri in pct_rows and ci>0 and style=='num': style='auto_pct'

            if val is None or val=='': txt=''
            elif isinstance(val,str): txt=val.strip()
            elif isinstance(val,datetime): txt=val.strftime('%d-%b-%y')
            elif isinstance(val,bool): txt=''
            else: txt=fmt(val,style)

            # Apply label shortening if defined
            if ci==0 and ri in label_shorten: txt=label_shorten[ri]

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
            fs = 'italic' if (is_italic and not is_bold) else 'normal'
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

 'Ingresos': dict(
    sheet='Ingresos',
    cols=[(2,'str'),(3,'num'),(5,'num'),(7,'auto_pct'),(9,'num'),(11,'num'),(13,'auto_pct')],
    rows=[3,4,5,6,9],
    bold_rows={0,1,4}, total_rows={4}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.30,0.115,0.115,0.085,0.115,0.115,0.085],
    group_headers=GH6_12_7col, fig_w=9.5,
 ),

 'IngresosxSegmento': dict(
    sheet='IngresosxSegmento (2)',
    cols=[(2,'str'),(3,'num'),(5,'num'),(6,'auto_pct'),(8,'num'),(10,'num'),(11,'auto_pct')],
    rows=[6,7,9,16,18],
    bold_rows={0,1,4}, total_rows={4}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.30,0.115,0.115,0.085,0.115,0.115,0.085],
    group_headers=GH6_12_7col, fig_w=9.5,
 ),

 'Costos': dict(
    sheet='Costos',
    cols=[(3,'str'),(4,'num'),(6,'num'),(8,'auto_pct'),(10,'num'),(12,'num'),(14,'auto_pct')],
    rows=[3,4,6,8,9,10,12,13,14],
    bold_rows={0,1,2,5,6,8}, italic_rows={3,4},
    total_rows={8}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.38,0.107,0.107,0.082,0.107,0.107,0.082],
    group_headers=GH6_12_7col, fig_w=9.5, row_h=0.30,
    # Shorten the very long row 3 label (0-based index of row 9 in rows list = index 5)
    label_shorten={4: 'Otros gastos por función,\nexcl. deterioros'},
 ),

 'OtrosComp': dict(
    sheet='Otros ingresos(gastos)',
    cols=[(3,'str'),(4,'num'),(6,'num'),(8,'auto_pct'),(10,'num'),(12,'num'),(14,'auto_pct')],
    rows=[5,6,8,9,10,11,12,14],
    bold_rows={0,1,7}, total_rows={7}, header_rows={0,1}, center_cols={1,2,3,4,5,6},
    col_widths=[0.38,0.107,0.107,0.082,0.107,0.107,0.082],
    group_headers=GH6_12_7col, fig_w=9.5,
    label_shorten={5: 'Participación en ganancias\nde asociadas'},
 ),

 'IndicadoresActividad': dict(
    sheet='Indicadores2',
    cols=[(3,'str'),(4,'str'),(5,'dec2'),(6,'dec2')],
    rows=[14,15,16,17,18,19,20,21],
    bold_rows={0,2,4,6}, italic_rows={1,3,5,7},
    total_rows=set(), header_rows={0}, center_cols={1,2,3},
    col_widths=[0.44,0.14,0.21,0.21],
    italic_col0_only=True,
 ),

 'DFN': dict(
    sheet='DF Neta',
    cols=[(2,'str'),(3,'num'),(5,'num')],
    rows=[5,6,7,8,9,10,11,12,13,15],
    bold_rows={0,1,5,7,9}, total_rows={5,9}, header_rows={0,1}, center_cols={1,2},
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
    sheet='Indicadores2',
    cols=[(3,'str'),(4,'str'),(5,'_auto'),(6,'_auto'),(7,'auto_pct')],
    rows=[4,5,6,7,8,9,10],
    bold_rows={0,1,3,5}, italic_rows={2,4,6},
    total_rows=set(), header_rows={0}, center_cols={1,2,3,4},
    col_widths=[0.40,0.14,0.155,0.155,0.15],
    row_fmt={1:'dec2',3:'pct',5:'pct'},
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
    col_widths=[0.27,0.093,0.093,0.093,0.083,0.083,0.083,0.083,0.038],
    fig_w=7.4, font_size=8.0, row_h=0.32,
    label_shorten={
        5:  'Deudores comerc. y otras\nctas por cobrar, ctes.',
        6:  'Ctas Cobrar Entidades\nRelacionadas, ctes.',
        9:  'Ctas Cobrar Entidades\nRelacionadas, NC',
        14: 'Ctas comerc. y otras ctas\npor pagar, corrientes',
        15: 'Ctas Pagar Entidades\nRelacionadas, ctes.',
        17: 'Prov. beneficios\nempleados, ctes.',
        21: 'Ctas Pagar Entidades\nRelacionadas, NC',
    },
 ),

 'EfectoTC': dict(
    sheet='ExposicTC',
    cols=[(3,'str'),(4,'num'),(5,'num'),(6,'num'),(7,'num'),(8,'num')],
    rows=[32,33,34,35,36,37,38,39,40,41,42],
    bold_rows={0,1,10}, total_rows={10}, header_rows={0,1}, center_cols={1,2,3,4,5},
    col_widths=[0.30,0.14,0.14,0.14,0.14,0.14],
    fig_w=6.8, font_size=9.5, row_h=0.33,
 ),

 # Risks: two sub-tables (filas 2-4 y 8-10, no contiguas en la hoja) -> group_gaps las separa
 # visualmente, igual que en el Word original. % stored as decimals → show as integer %
 'Risks': dict(
    sheet='Risks',
    cols=[(4,'str'),(5,'num'),(6,'num'),(7,'num'),(8,'num'),(9,'num'),(10,'num')],
    rows=[2,3,4,8,9,10],
    bold_rows={0,2,3,5}, total_rows={2,5},
    header_rows={0,3},
    center_cols={1,2,3,4,5,6},
    col_widths=[0.24,0.13,0.12,0.12,0.12,0.12,0.15],
    label_shorten={3: 'Severidad\ndel Riesgo'},
    # % rows: ri=2 and ri=5 have decimal values → format as %
    pct_rows={2,5},  # these rows get auto_pct formatting for cols 1-6
    fig_w=6.8, font_size=9.5, row_h=0.36, group_gaps=True,
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
    col_widths=[0.10,0.35,0.08,0.235,0.235], yellow_all=True,
 ),

 'FVfruta': dict(
    sheet='FVfruta',
    cols=[(2,'str'),(4,'num'),(5,'num'),(6,'num'),(7,'num')],
    rows=[4,5,6,7,11],
    bold_rows={0,1,2,4}, total_rows={4}, header_rows={0,1,2}, center_cols={1,2,3,4},
    col_widths=[0.28,0.18,0.18,0.18,0.18], yellow_all=True,
 ),
}

IMAGE_TO_TABLE = {
 'image1.emf':'EBITDA','image2.emf':'Ingresos','image3.emf':'IngresosxSegmento',
 'image4.emf':'Costos','image5.emf':'OtrosComp','image6.emf':'IndicadoresActividad',
 'image7.emf':'DFN','image8.emf':'Indicadores1','image9.emf':'Rentabilidad',
 'image9.png':'Rentabilidad',
 'image10.emf':'Balance','image11.emf':'VctosPtmos','image12.emf':'ExposicTC_tabla',
 'image13.emf':'EfectoTC','image14.emf':'Risks','image15.emf':'Riesgos',
 'image16.emf':'Riesgos','image17.emf':'Seguros','image18.emf':'FVfruta',
}

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
