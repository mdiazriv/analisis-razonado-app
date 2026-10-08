# -*- coding: utf-8 -*-
"""
Deriva las etiquetas de período (Jun26, Jun25, Dic25, fechas largas, temporadas, cantidad de meses del
tramo calendario y de la temporada) a partir de las fechas del Excel, para que el motor de generación
sirva para cualquiera de los 4 cierres trimestrales (marzo/trimestral, junio/semestral, septiembre/9M,
diciembre/anual) sin quedar amarrado a un trimestre fijo.
"""
import datetime

MESES_ABR = ['', 'Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
MESES_LARGO = ['', 'enero','febrero','marzo','abril','mayo','junio','julio','agosto',
               'septiembre','octubre','noviembre','diciembre']
MESES_NUM_PALABRA = {3: 'tres', 6: 'seis', 9: 'nueve', 12: 'doce'}

def _short(d):
    return f"{MESES_ABR[d.month]}{d.year % 100:02d}"

def _long(d):
    return f"{d.day} de {MESES_LARGO[d.month]} de {d.year}"

def _season_labels(d):
    # Temporada agrícola: 1 jul - 30 jun. Si month<=6, la temporada que termina es (year-1)/(year).
    if d.month <= 6:
        y0, y1 = d.year - 1, d.year
    else:
        y0, y1 = d.year, d.year + 1
    cur = f"T{y0%100:02d}/{y1%100:02d}"
    prior = f"T{(y0-1)%100:02d}/{(y1-1)%100:02d}"
    return cur, prior, y0, y1

def derive_periods(wb, cur_date=None, prior_fy_date=None):
    """cur_date / prior_fy_date: datetime.date del cierre actual y del cierre de año fiscal anterior.
    Si no se entregan, se intentan leer de la hoja 'DF Neta' (C5 y E5)."""
    if cur_date is None or prior_fy_date is None:
        ws = wb['DF Neta']
        cur_date = ws['C5'].value
        prior_fy_date = ws['E5'].value
        if hasattr(cur_date, 'date'): cur_date = cur_date.date()
        if hasattr(prior_fy_date, 'date'): prior_fy_date = prior_fy_date.date()

    prior_same_period = cur_date.replace(year=cur_date.year - 1)
    season_cur, season_prior, season_y0, season_y1 = _season_labels(cur_date)

    # Tramo calendario (año en curso, Ene-cierre): marzo=3, junio=6, septiembre=9, diciembre=12.
    cal_meses = cur_date.month
    # Tramo temporada (temporada agrícola, Jul-cierre): junio=12 (temporada completa), resto = lo que
    # corresponda según cuántos meses van corridos desde el 1 de julio.
    season_meses = (cal_meses - 6) if cal_meses > 6 else (cal_meses + 6)

    # El párrafo preámbulo ("comparándose con los estados financieros al...") usa como referencia el
    # cierre anual anterior (Dic) para los cierres de semestre/año (junio y diciembre), y el mismo
    # período del año anterior para los cierres trimestrales intermedios (marzo y septiembre) — así
    # vienen redactados los informes originales de Hortifrut para cada tipo de cierre.
    if cal_meses in (6, 12):
        preambulo_ref_short = _short(prior_fy_date)
        preambulo_ref_long = _long(prior_fy_date)
    else:
        preambulo_ref_short = _short(prior_same_period)
        preambulo_ref_long = _long(prior_same_period)

    return {
        'CUR_SHORT': _short(cur_date),               # Jun26
        'PRIOR_SHORT': _short(prior_same_period),     # Jun25
        'FY_PRIOR_SHORT': _short(prior_fy_date),      # Dic25
        'CUR_LONG': _long(cur_date),                  # 30 de junio de 2026
        'FY_PRIOR_LONG': _long(prior_fy_date),         # 31 de diciembre de 2025
        'SEASON_CUR': season_cur,                     # T25/26
        'SEASON_PRIOR': season_prior,                 # T24/25
        'CAL_MESES': cal_meses,                       # 3 / 6 / 9 / 12
        'SEASON_MESES': season_meses,                 # 9 / 12 / 3 / 6 (según el cierre)
        'SEASON_MESES_PALABRA': MESES_NUM_PALABRA.get(season_meses, str(season_meses)),
        'CUR_MONTH_LOWER_YEAR': f"{MESES_LARGO[cur_date.month]} {cur_date.year}",      # "junio 2026"
        'CUR_MONTH_CAP': MESES_LARGO[cur_date.month].capitalize(),                    # "Junio"
        'SEASON_RANGE_LONG': f"julio {season_y0} – {MESES_LARGO[cur_date.month]} {cur_date.year}",
        'PREAMBULO_REF_SHORT': preambulo_ref_short,
        'PREAMBULO_REF_LONG': preambulo_ref_long,
    }

# Tokens literales usados dentro de los templates de ar_engine.py (siempre redactados en "estilo
# Jun26", que es el estilo de referencia pedido para todos los trimestres), mapeados a las claves de
# arriba. localize() los reemplaza por las etiquetas derivadas del Excel cargado, independiente de qué
# trimestre sea en realidad.
TOKEN_MAP = {
    '30 de junio de 2026': 'CUR_LONG',
    '31 de diciembre de 2025': 'FY_PRIOR_LONG',
    '12M T25/26': None,   # manejado abajo (compuesto, con el conteo de meses dinámico)
    '12M T24/25': None,
    'T25/26': 'SEASON_CUR',
    'T24/25': 'SEASON_PRIOR',
    'Jun26': 'CUR_SHORT',
    'Jun25': 'PRIOR_SHORT',
    'Dic25': 'FY_PRIOR_SHORT',
}

def localize(text, periods):
    # Importante: se reemplaza por marcadores temporales únicos antes de insertar el texto final, en
    # dos pasadas. Si se reemplazara directamente en una sola pasada, el resultado de un reemplazo
    # podría quedar "recapturado" por un reemplazo posterior cuyo patrón de búsqueda coincide con el
    # texto recién insertado — por ejemplo, en el cierre de diciembre CUR_SHORT es literalmente
    # "Dic25", el mismo texto que el token fijo 'Dic25' (que representa FY_PRIOR_SHORT) buscado más
    # abajo, lo que rompería la fecha recién puesta.
    repl = [
        ('12M T25/26', f"{periods['SEASON_MESES']}M {periods['SEASON_CUR']}"),
        ('12M T24/25', f"{periods['SEASON_MESES']}M {periods['SEASON_PRIOR']}"),
        # Tramo calendario: "6 meses" (único conteo usado en los templates de ar_engine.py) -> el
        # conteo real de este cierre (3/6/9/12).
        ('6 meses', f"{periods['CAL_MESES']} meses"),
        ('30 de junio de 2026', periods['CUR_LONG']),
        ('31 de diciembre de 2025', periods['FY_PRIOR_LONG']),
        ('T25/26', periods['SEASON_CUR']),
        ('T24/25', periods['SEASON_PRIOR']),
        ('Jun26', periods['CUR_SHORT']),
        ('Jun25', periods['PRIOR_SHORT']),
        ('Dic25', periods['FY_PRIOR_SHORT']),
    ]
    placeholders = []
    for i, (old, new) in enumerate(repl):
        ph = f"\x00{i}\x00"
        if old in text:
            text = text.replace(old, ph)
            placeholders.append((ph, new))
    for ph, new in placeholders:
        text = text.replace(ph, new)
    return text

if __name__ == '__main__':
    import sys
    sys.path.insert(0, '.')
    import openpyxl
    wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
    p = derive_periods(wb)
    for k, v in p.items():
        print(k, '=', v)
    print(localize("El EBITDA a Jun26 alcanzó X, comparado con Jun25 y Dic25. 12M T25/26 vs 12M T24/25. "
                    "Durante los 6 meses terminados en Jun26...", p))
