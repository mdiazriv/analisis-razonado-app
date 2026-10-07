# -*- coding: utf-8 -*-
"""
Deriva las etiquetas de período (Jun26, Jun25, Dic25, fechas largas, temporadas) a partir de las
fechas del Excel, para que el motor de generación no quede amarrado a un trimestre fijo.
"""
import datetime

MESES_ABR = ['', 'Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
MESES_LARGO = ['', 'enero','febrero','marzo','abril','mayo','junio','julio','agosto',
               'septiembre','octubre','noviembre','diciembre']

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
    return cur, prior

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
    season_cur, season_prior = _season_labels(cur_date)

    return {
        'CUR_SHORT': _short(cur_date),               # Jun26
        'PRIOR_SHORT': _short(prior_same_period),     # Jun25
        'FY_PRIOR_SHORT': _short(prior_fy_date),      # Dic25
        'CUR_LONG': _long(cur_date),                  # 30 de junio de 2026
        'FY_PRIOR_LONG': _long(prior_fy_date),         # 31 de diciembre de 2025
        'SEASON_CUR': season_cur,                     # T25/26
        'SEASON_PRIOR': season_prior,                 # T24/25
    }

# Tokens literales usados dentro de los templates de ar_engine.py, mapeados a las claves de arriba.
# Mientras el Excel sea del mismo trimestre de referencia (Jun26), estos tokens coinciden con las
# etiquetas reales por diseño; localize() los reemplaza por las etiquetas derivadas del Excel cargado.
TOKEN_MAP = {
    '30 de junio de 2026': 'CUR_LONG',
    '31 de diciembre de 2025': 'FY_PRIOR_LONG',
    '12M T25/26': None,   # manejado abajo (compuesto)
    '12M T24/25': None,
    'T25/26': 'SEASON_CUR',
    'T24/25': 'SEASON_PRIOR',
    'Jun26': 'CUR_SHORT',
    'Jun25': 'PRIOR_SHORT',
    'Dic25': 'FY_PRIOR_SHORT',
}

def localize(text, periods):
    # Reemplazar primero los compuestos "12M T.." para que no los rompa el reemplazo de "T25/26" suelto.
    text = text.replace('12M T25/26', f"12M {periods['SEASON_CUR']}")
    text = text.replace('12M T24/25', f"12M {periods['SEASON_PRIOR']}")
    text = text.replace('30 de junio de 2026', periods['CUR_LONG'])
    text = text.replace('31 de diciembre de 2025', periods['FY_PRIOR_LONG'])
    text = text.replace('T25/26', periods['SEASON_CUR'])
    text = text.replace('T24/25', periods['SEASON_PRIOR'])
    text = text.replace('Jun26', periods['CUR_SHORT'])
    text = text.replace('Jun25', periods['PRIOR_SHORT'])
    text = text.replace('Dic25', periods['FY_PRIOR_SHORT'])
    return text

if __name__ == '__main__':
    import sys
    sys.path.insert(0, '.')
    import openpyxl
    wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
    p = derive_periods(wb)
    for k, v in p.items():
        print(k, '=', v)
    print(localize("El EBITDA a Jun26 alcanzó X, comparado con Jun25 y Dic25. 12M T25/26 vs 12M T24/25.", p))
