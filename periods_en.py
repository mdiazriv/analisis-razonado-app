# -*- coding: utf-8 -*-
"""
Versión en inglés de periods.py: deriva las mismas etiquetas de período pero con las convenciones de
redacción en inglés usadas en el Word de referencia de Hortifrut (Earning_Report_June_2026):
  - Abreviaturas de mes en inglés (Dic -> Dec, Ene -> Jan, Abr -> Apr, Ago -> Aug; el resto coincide).
  - Fechas largas en formato "June 30, 2026" (no "30 de junio de 2026").
  - Temporada con prefijo "S" en vez de "T" ("S25/26" en vez de "T25/26").
Comparado celda a celda contra Earning_Report_June_2026_27_08_26.docx (el mismo trimestre que
ar_data/ar.docx) para confirmar que coincide exactamente.
"""
import datetime

MESES_ABR_EN = ['', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
MESES_LARGO_EN = ['', 'January','February','March','April','May','June','July','August',
                  'September','October','November','December']
MESES_NUM_PALABRA_EN = {3: 'three', 6: 'six', 9: 'nine', 12: 'twelve'}

def _short(d):
    return f"{MESES_ABR_EN[d.month]}{d.year % 100:02d}"

def _long(d):
    return f"{MESES_LARGO_EN[d.month]} {d.day}, {d.year}"

def _season_labels(d):
    if d.month <= 6:
        y0, y1 = d.year - 1, d.year
    else:
        y0, y1 = d.year, d.year + 1
    cur = f"S{y0%100:02d}/{y1%100:02d}"
    prior = f"S{(y0-1)%100:02d}/{(y1-1)%100:02d}"
    return cur, prior, y0, y1

def derive_periods_en(wb, cur_date=None, prior_fy_date=None):
    if cur_date is None or prior_fy_date is None:
        ws = wb['DF Neta']
        cur_date = ws['C5'].value
        prior_fy_date = ws['E5'].value
        if hasattr(cur_date, 'date'): cur_date = cur_date.date()
        if hasattr(prior_fy_date, 'date'): prior_fy_date = prior_fy_date.date()

    prior_same_period = cur_date.replace(year=cur_date.year - 1)
    season_cur, season_prior, season_y0, season_y1 = _season_labels(cur_date)

    cal_meses = cur_date.month
    season_meses = (cal_meses - 6) if cal_meses > 6 else (cal_meses + 6)

    if cal_meses in (6, 12):
        preambulo_ref_short = _short(prior_fy_date)
        preambulo_ref_long = _long(prior_fy_date)
    else:
        preambulo_ref_short = _short(prior_same_period)
        preambulo_ref_long = _long(prior_same_period)

    return {
        'CUR_SHORT': _short(cur_date),               # Jun26
        'PRIOR_SHORT': _short(prior_same_period),     # Jun25
        'FY_PRIOR_SHORT': _short(prior_fy_date),      # Dec25
        'CUR_LONG': _long(cur_date),                  # June 30, 2026
        'FY_PRIOR_LONG': _long(prior_fy_date),         # December 31, 2025
        'SEASON_CUR': season_cur,                     # S25/26
        'SEASON_PRIOR': season_prior,                 # S24/25
        'CAL_MESES': cal_meses,
        'SEASON_MESES': season_meses,
        'SEASON_MESES_PALABRA': MESES_NUM_PALABRA_EN.get(season_meses, str(season_meses)),
        'CUR_MONTH_YEAR': f"{MESES_LARGO_EN[cur_date.month]} {cur_date.year}",      # "June 2026"
        'CUR_MONTH_CAP': MESES_LARGO_EN[cur_date.month],                            # "June"
        'SEASON_RANGE_LONG': f"{MESES_LARGO_EN[7]} {season_y0} – {MESES_LARGO_EN[cur_date.month]} {cur_date.year}",
        'PREAMBULO_REF_SHORT': preambulo_ref_short,
        'PREAMBULO_REF_LONG': preambulo_ref_long,
    }

# Tokens literales usados dentro de los templates de ar_engine_en.py (siempre redactados en "estilo
# Jun26"), mapeados a las claves de arriba, igual que TOKEN_MAP en periods.py.
TOKEN_MAP_EN = {
    'June 30, 2026': 'CUR_LONG',
    'December 31, 2025': 'FY_PRIOR_LONG',
    'S25/26': 'SEASON_CUR',
    'S24/25': 'SEASON_PRIOR',
    'Jun26': 'CUR_SHORT',
    'Jun25': 'PRIOR_SHORT',
    'Dec25': 'FY_PRIOR_SHORT',
}

def localize_en(text, periods):
    repl = [
        ('12M S25/26', f"{periods['SEASON_MESES']}M {periods['SEASON_CUR']}"),
        ('12M S24/25', f"{periods['SEASON_MESES']}M {periods['SEASON_PRIOR']}"),
        ('6 months', f"{periods['CAL_MESES']} months"),
        ('June 30, 2026', periods['CUR_LONG']),
        ('December 31, 2025', periods['FY_PRIOR_LONG']),
        ('S25/26', periods['SEASON_CUR']),
        ('S24/25', periods['SEASON_PRIOR']),
        ('Jun26', periods['CUR_SHORT']),
        ('Jun25', periods['PRIOR_SHORT']),
        ('Dec25', periods['FY_PRIOR_SHORT']),
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

_MES_LARGO_EN_TO_NUM = {m.lower(): i for i, m in enumerate(MESES_LARGO_EN) if m}

def parse_long_date_en(text):
    """Extrae una fecha tipo 'As of March 31, 2026' o '...December 31, 2025...' del texto en inglés,
    o None si no calza el patrón. Case-insensitive en el nombre del mes."""
    import re
    m = re.search(r'(\w+) (\d{1,2}),\s*(\d{4})', text, re.IGNORECASE)
    if not m:
        return None
    mes_num = _MES_LARGO_EN_TO_NUM.get(m.group(1).lower())
    if mes_num is None:
        return None
    day = int(m.group(2)); year = int(m.group(3))
    try:
        return datetime.date(year, mes_num, day)
    except ValueError:
        return None

def derive_old_periods_from_title_en(title_text):
    old_cur_date = parse_long_date_en(title_text)
    if old_cur_date is None:
        return None
    old_prior_fy_date = datetime.date(old_cur_date.year - 1, 12, 31)
    return derive_periods_en(None, cur_date=old_cur_date, prior_fy_date=old_prior_fy_date)

def sweep_pairs_en(old_periods, new_periods):
    pairs = [
        (f"{old_periods['SEASON_MESES']}M {old_periods['SEASON_CUR']}", f"{new_periods['SEASON_MESES']}M {new_periods['SEASON_CUR']}"),
        (f"{old_periods['SEASON_MESES']}M{old_periods['SEASON_CUR']}", f"{new_periods['SEASON_MESES']}M{new_periods['SEASON_CUR']}"),
        (f"{old_periods['SEASON_MESES']}M {old_periods['SEASON_PRIOR']}", f"{new_periods['SEASON_MESES']}M {new_periods['SEASON_PRIOR']}"),
        (f"{old_periods['SEASON_MESES']}M{old_periods['SEASON_PRIOR']}", f"{new_periods['SEASON_MESES']}M{new_periods['SEASON_PRIOR']}"),
        (old_periods['CUR_LONG'], new_periods['CUR_LONG']),
        (old_periods['FY_PRIOR_LONG'], new_periods['FY_PRIOR_LONG']),
        (old_periods['SEASON_CUR'], new_periods['SEASON_CUR']),
        (old_periods['SEASON_PRIOR'], new_periods['SEASON_PRIOR']),
        (old_periods['CUR_SHORT'], new_periods['CUR_SHORT']),
        (old_periods['PRIOR_SHORT'], new_periods['PRIOR_SHORT']),
        (old_periods['FY_PRIOR_SHORT'], new_periods['FY_PRIOR_SHORT']),
    ]
    seen = set()
    out = []
    for old, new in pairs:
        if old == new or old in seen:
            continue
        seen.add(old)
        out.append((old, new))
    return out

def sweep_text_en(text, old_periods, new_periods):
    pairs = sweep_pairs_en(old_periods, new_periods)
    placeholders = []
    for i, (old, new) in enumerate(pairs):
        ph = f"\x01{i}\x01"
        if old in text:
            text = text.replace(old, ph)
            placeholders.append((ph, new))
    for ph, new in placeholders:
        text = text.replace(ph, new)
    return text
