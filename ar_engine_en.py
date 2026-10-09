# -*- coding: utf-8 -*-
"""
Versión en inglés de ar_engine.py. Mismas celdas del mismo Excel (ar_v2.xlsx), misma lógica de
cálculo/bridge/driver que el motor en español — solo cambian las plantillas de redacción (en inglés,
con formato numérico en-US: punto decimal, coma de miles) y el orden de palabras donde el traductor
profesional de Hortifrut reordena la frase (ej. "a 14.49% increase" en vez de "an increase of 14.49%").

Cada plantilla fue construida comparando, frase por frase, Earning_Report_June_2026_27_08_26.docx
(el Word en inglés real de junio 2026, el MISMO trimestre que ar_data/ar.docx) contra la salida real
del motor en español (ar_engine.py) para ese mismo Excel — no contra el Word en español a mano, que
en algunos párrafos redacta distinto al motor (ver nota en cada función de bridge).
"""
from decimal import Decimal, ROUND_HALF_UP

CAUSA_TAG_EN = "CAUSE NOT DERIVABLE FROM EXCEL: update if applicable."

def _with_cause(base_text, cause):
    text = base_text[:-1] if base_text.endswith('.') else base_text
    if cause:
        return text + f", [[{cause} ({CAUSA_TAG_EN})]]"
    return text + f", [[{CAUSA_TAG_EN}]]"

def c(wb, sheet, ref):
    return wb[sheet][ref].value

def _round(value, decimals):
    q = Decimal('1' if decimals == 0 else '1.' + '0'*decimals)
    return Decimal(repr(value)).quantize(q, rounding=ROUND_HALF_UP)

def _grp(s):
    return f"{int(s):,}"  # English thousands separator = comma

# ---------- formatting (en-US) ----------
def fmt_millones(raw_thousands, force_decimals=None):
    v = raw_thousands / 1000.0
    neg = v < 0
    av = abs(v)
    decimals = force_decimals if force_decimals is not None else (2 if av < 1000 else 0)
    q = _round(av, decimals)
    if decimals == 0:
        s = _grp(q)
    else:
        whole, frac = str(q).split('.')
        s = _grp(whole) + '.' + frac
    return ('-' if neg else '') + s

def fmt_pct(frac, decimals=2):
    v = frac * 100
    neg = v < 0
    q = _round(abs(v), decimals)
    s = str(q) if decimals else str(int(q))
    return ('-' if neg else '') + s

def fmt_veces(v, decimals=2):
    neg = v < 0
    q = _round(abs(v), decimals)
    return ('-' if neg else '') + str(q)

def _leading_num_needs_an(i):
    i = abs(int(i))
    return i == 8 or i == 11 or i == 18 or 80 <= i <= 89

def indef_article(frac, decimals=2):
    """Devuelve 'a' o 'an' según cómo se lee en inglés el entero formateado por fmt_pct para esa
    misma fracción (ocho/eight, once/eleven, dieciocho/eighteen y el rango ochenta/eighty-* se leen
    con sonido vocálico inicial). Verificado contra el Word de referencia: 'a 14.49%' pero 'an 11.65%'
    y 'an 8.01%'."""
    i = int(_round(abs(frac) * 100, decimals))
    return 'an' if _leading_num_needs_an(i) else 'a'

def indef_article_millones(raw_thousands, force_decimals=None):
    """Igual que indef_article pero para un monto formateado por fmt_millones (ej. 'a US$32.14 million'
    vs 'an US$8.01 million'), evaluando el entero de la parte formateada en millones."""
    v = raw_thousands / 1000.0
    decimals = force_decimals if force_decimals is not None else (2 if abs(v) < 1000 else 0)
    i = int(_round(abs(v), decimals))
    return 'an' if _leading_num_needs_an(i) else 'a'

def sign_str(v):
    return '+' if v >= 0 else ''

def inc_dec(v, inc='increase', dec='decrease'):
    return inc if v >= 0 else dec

def crec_caida(v, inc='growth', dec='decline'):
    return inc if v >= 0 else dec

def volumen_precio_driver(vol_growth, price_growth, min_material=0.03):
    vol_up = vol_growth > min_material
    price_up = price_growth > min_material
    vol_down = vol_growth < -min_material
    price_down = price_growth < -min_material
    if vol_up and price_up:
        return 'both'
    if vol_up and price_down:
        return 'vol_despite_price'
    if price_up and vol_down:
        return 'price_despite_vol'
    if vol_up:
        return 'vol_only'
    if price_up:
        return 'price_only'
    if vol_down and price_down:
        return 'both_down'
    return 'unclear'

# ============================================================
# SECTION: SUMMARY OF THE PERIOD (EBITDA)
# ============================================================
def p_ebitda_6m(wb):
    S = 'EBITDA'
    ebitda = c(wb, S, 'C13'); ebitda_prior = c(wb, S, 'E13'); delta = c(wb, S, 'K13'); pct = c(wb, S, 'L13')
    ing_pct = c(wb, S, 'L4'); cost_pct = c(wb, S, 'L21')
    cr_cur = c(wb, S, 'C22'); cr_prior = c(wb, S, 'E22')
    return (
        f"EBITDA as of Jun26 reached US${fmt_millones(ebitda)} million, which represents {indef_article(pct)} {fmt_pct(pct)}% "
        f"{inc_dec(pct)} compared with the US${fmt_millones(ebitda_prior)} million recorded as of Jun25 "
        f"({sign_str(delta)}US${fmt_millones(delta)} million). This {inc_dec(pct)} is mainly explained by the "
        f"{fmt_pct(ing_pct)}% {crec_caida(ing_pct)} in income, partially offset by the {fmt_pct(cost_pct)}% "
        f"{inc_dec(cost_pct)} in Costs and Expenses (excluding impairment and depreciation). Costs and expenses "
        f"without depreciation represented {fmt_pct(abs(cr_cur))}% of income from operating activities as of "
        f"Jun26, compared with {fmt_pct(abs(cr_prior))}% as of Jun25."
    )

def p_ebitda_sinfv_6m(wb):
    S = 'EBITDA'
    ebitda = c(wb, S, 'C14'); prior = c(wb, S, 'E14'); pct = c(wb, S, 'L14')
    fv_cur = c(wb, S, 'C18'); fv_prior = c(wb, S, 'E18')
    return (
        f"Accumulated EBITDA as of Jun26 without the fair value effect of fruit reached US${fmt_millones(ebitda)} "
        f"million, which represents {indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct)}, compared with the US${fmt_millones(prior)} "
        f"million obtained in the same period. The impact of net fair value as of Jun26 was US${fmt_millones(fv_cur)} "
        f"million, while as of Jun25 it was US${fmt_millones(fv_prior)} million."
    )

# ============================================================
# SECTION: INCOME STATEMENT ANALYSIS (bridge)
# ============================================================
GA_ROWS = {
    'EBITDA': 4, 'Deterioro': 5, 'Depreciación y amortizaciones': 6,
    'Otros ingresos (gastos) no operacionales': 7, 'Costos Financieros Netos': 8,
    'Participación en asociadas': 9, 'Diferencia de cambio': 10,
    'Gasto por impuestos a las ganancias': 11,
}
# Etiquetas en inglés para mostrar en el texto (las keys arriba son las usadas internamente para leer
# el Excel, igual que en ar_engine.py — no cambian).
LABEL_EN = {
    'Deterioro': 'impairment in the value of assets',
    'Depreciación y amortizaciones': 'expense due to depreciation and amortization',
    'Costos Financieros Netos': 'net financial expenses',
    'Diferencia de cambio': 'exchange rate difference',
    'Gasto por impuestos a las ganancias': 'taxes',
    'Otros ingresos (gastos) no operacionales': 'other non-operating income (expenses)',
    'Participación en asociadas': 'share of profit of associates',
}

def ga_row(wb, label, period='6m'):
    r = GA_ROWS[label]
    S = 'Ganancia Atribuible'
    if period == '6m':
        return dict(cur=c(wb,S,f'C{r}'), prior=c(wb,S,f'E{r}'), delta=c(wb,S,f'K{r}'))
    else:
        return dict(cur=c(wb,S,f'G{r}'), prior=c(wb,S,f'I{r}'), delta=c(wb,S,f'N{r}'))

def p_ganancia_controladora_6m(wb):
    S='Ganancia Atribuible'
    cur=c(wb,S,'C13'); prior=c(wb,S,'E13'); delta=c(wb,S,'K13')
    mejora_empeora = 'an improvement' if delta>=0 else 'a worsening'
    ganancia_perdida = lambda v: 'a profit' if v>=0 else 'a loss'
    loss_noun = lambda v: 'profit' if v>=0 else 'loss'
    return (
        f"During the 6 months ending Jun26, {ganancia_perdida(cur)} attributable to parent company shareholders "
        f"of US${fmt_millones(cur)} million was recorded, which is compared with the US${fmt_millones(prior)} "
        f"million controlling {loss_noun(prior)} recorded as of Jun25 (representing {mejora_empeora} of "
        f"{sign_str(delta)}US${fmt_millones(delta)} million)."
    )

def p_ebitda_bullet_6m(wb):
    S = 'EBITDA'
    cur = c(wb, S, 'C13'); prior = c(wb, S, 'E13'); delta = c(wb, S, 'K13')
    vol_g = c(wb, 'Ingresos', 'E33'); price_g = c(wb, 'Ingresos', 'D36')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'higher commercialized volumes and higher prices',
        'vol_only': 'higher commercialized volumes',
        'vol_despite_price': 'higher commercialized volumes',
        'price_only': 'higher prices',
        'price_despite_vol': 'higher prices',
        'both_down': 'lower commercialized volumes and lower prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    mayor_menor = 'Higher' if delta >= 0 else 'Lower'
    return (
        f"{mayor_menor} EBITDA, which reached US${fmt_millones(cur)} million as of Jun26, compared with "
        f"US${fmt_millones(prior)} million as of Jun25 ({sign_str(delta)}US${fmt_millones(delta)} million), "
        f"explained by {phrase}."
    )

def p_ebitda_bullet_12m(wb):
    S = 'EBITDA'
    cur = c(wb, S, 'G13'); prior = c(wb, S, 'I13'); delta = c(wb, S, 'N13')
    inc_dec_noun = 'an increase' if delta >= 0 else 'a decrease'
    return (
        f"EBITDA of the 12M S25/26 for US${fmt_millones(cur)} million, with {inc_dec_noun} compared to the "
        f"result of the 12M S24/25 of US${fmt_millones(prior)} million ({sign_str(delta)}US${fmt_millones(delta)} "
        f"million), due to the reasons already explained."
    )

# Nota: estas plantillas de bridge se compararon contra la SALIDA REAL del motor en español (no contra
# el Word en español a mano, que en el párrafo de temporada redacta distinto, ej. "Deterioro por US$X
# asociado a..." en vez de "Menor deterioro del valor de activos que se registró por US$X..."). El
# motor en español usa SIEMPRE esta redacción genérica (bridge_factor_lead_in/bridge_factor_text) para
# cualquier factor del bridge, calendario o temporada — el inglés hace exactamente lo mismo.
def bridge_factor_lead_in(label, cur, prior, delta):
    if label == 'Deterioro':
        return 'Lower impairment in the value of assets, which was recorded for' if delta >= 0 else \
               'Higher impairment in the value of assets, which was recorded for'
    if label == 'Depreciación y amortizaciones':
        return 'Higher expense due to depreciation and amortization, which was recorded for' if delta < 0 else \
               'Lower expense due to depreciation and amortization, which was recorded for'
    if label == 'Costos Financieros Netos':
        return 'Higher net financial expenses, which reached' if delta < 0 else \
               'Lower net financial expenses, which reached'
    if label == 'Diferencia de cambio':
        return 'Negative exchange rate difference of' if cur < 0 else 'Positive exchange rate difference of'
    if label == 'Gasto por impuestos a las ganancias':
        return 'Positive taxes for' if cur >= 0 else 'Negative taxes for'
    if label == 'Otros ingresos (gastos) no operacionales':
        return 'Favorable variation in other non-operating income (expenses) for' if delta >= 0 else \
               'Unfavorable variation in other non-operating income (expenses) for'
    if label == 'Participación en asociadas':
        return 'Favorable variation in the share of profit of associates for' if delta >= 0 else \
               'Unfavorable variation in the share of profit of associates for'
    return f"{LABEL_EN.get(label, label)}: variation for"

def bridge_factor_text(label, cur, prior, delta, period='6m', cause=None):
    lead = bridge_factor_lead_in(label, cur, prior, delta)
    cur_lbl = 'as of Jun26' if period == '6m' else 'in the 12M S25/26'
    prior_lbl = 'as of Jun25' if period == '6m' else 'in the 12M S24/25'
    base = (
        f"{lead} US${fmt_millones(cur)} million {cur_lbl}, compared with "
        f"US${fmt_millones(prior)} million {prior_lbl} ({sign_str(delta)}US${fmt_millones(delta)} million)."
    )
    if cause:
        base = base[:-1] + f", [[{cause} ({CAUSA_TAG_EN})]]."
    return base

def p_ganancia_controladora_12m(wb):
    S = 'Ganancia Atribuible'
    cur = c(wb, S, 'G13'); prior = c(wb, S, 'I13'); delta = c(wb, S, 'N13')
    mejora_empeora = 'improving' if delta >= 0 else 'worsening'
    return (
        f"Profit (loss) attributable to the parent company shareholders was US${fmt_millones(cur)} million in "
        f"the 12M S25/26, {mejora_empeora} in relation to the US${fmt_millones(prior)} million in the 12M S24/25 "
        f"({sign_str(delta)}US${fmt_millones(delta)} million), whose result is mainly explained by:"
    )

def p_ebitda_12m_intro(wb):
    S = 'EBITDA'
    cur = c(wb, S, 'G13'); prior = c(wb, S, 'I13')
    return (
        f"During the 12M S25/26, EBITDA reached US${fmt_millones(cur)} million, representing an increase "
        f"compared to the EBITDA of the 12M S24/25 of US${fmt_millones(prior)} million. This higher EBITDA is "
        f"explained by:"
    )

def p_ebitda_12m_bullet_ingresos(wb):
    pct = c(wb, 'Ingresos', 'M4')
    vol_g = c(wb, 'Ingresos', 'E54'); price_g = c(wb, 'Ingresos', 'E57')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'associated to higher volumes and prices',
        'vol_only': 'associated to higher volumes',
        'vol_despite_price': 'associated to higher volumes',
        'price_only': 'associated to higher prices',
        'price_despite_vol': 'associated to higher prices',
        'both_down': 'associated to lower volumes and prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    cap = 'Increase' if pct >= 0 else 'Decrease'
    return f"{cap} in income by {fmt_pct(pct)}% {phrase}."

def p_ebitda_12m_bullet_costos(wb):
    S = 'EBITDA'
    cur_ratio = c(wb, S, 'G22'); prior_ratio = c(wb, S, 'I22')
    cost_prior = c(wb, S, 'I21'); cost_delta = c(wb, S, 'N21')
    pct = abs(cost_delta) / abs(cost_prior)
    return (
        f"Costs and expenses without including depreciation recorded {indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct,'increase','decrease')}. "
        f"The total of costs and expenses without including depreciation represented {fmt_pct(abs(cur_ratio))}% of "
        f"total income in the 12M S25/26, compared with {fmt_pct(abs(prior_ratio))}% in the 12M S24/25."
    )

def bridge_bullets(wb, period='6m', top_n=None, threshold_abs_mus=3.0):
    items = []
    for label in GA_ROWS:
        if label == 'EBITDA':
            continue
        d = ga_row(wb, label, period)
        items.append((label, d['cur'], d['prior'], d['delta']))
    items.sort(key=lambda x: abs(x[3]), reverse=True)
    if top_n:
        items = items[:top_n]
    else:
        items = [it for it in items if abs(it[3])/1000.0 >= threshold_abs_mus]
    return items

# ============================================================
# SECTION: INCOME (volume/price driver)
# ============================================================
def p_ingresos_totales_6m(wb):
    S='Ingresos'
    cur=c(wb,S,'C4'); prior=c(wb,S,'E4'); pct=c(wb,S,'G4')
    vol_g = c(wb,S,'E33'); price_g = c(wb,S,'D36')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'due to the increase in volumes and sale prices',
        'vol_only': 'due to the increase in sales volumes',
        'vol_despite_price': 'due to the increase in sales volumes',
        'price_only': 'due to the increase in sale prices',
        'price_despite_vol': 'due to the increase in sale prices',
        'both_down': 'due to the decrease in volumes and sale prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    return (
        f"Income from operating activities reached US${fmt_millones(cur, force_decimals=0)} million as of Jun26, "
        f"representing {indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct)} compared to Jun25, {phrase}."
    )

def p_frutafresca_6m(wb):
    S='Ingresos'
    pct = c(wb,S,'G15')
    vol_g = c(wb,S,'E39'); price_g = c(wb,S,'E42')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'mainly explained by higher commercialized volumes and prices',
        'vol_only': 'mainly explained by higher commercialized volumes',
        'vol_despite_price': 'mainly explained by higher commercialized volumes',
        'price_only': 'mainly explained by higher prices',
        'price_despite_vol': 'mainly explained by higher prices',
        'both_down': 'mainly explained by lower commercialized volumes and prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    verbo = 'increased' if pct>=0 else 'decreased'
    return (
        f"Sales from the Fresh Fruit segment as of Jun26 {verbo} {fmt_pct(pct)}% compared to the previous "
        f"period, {phrase}."
    )

def p_valoragregado_6m(wb):
    S='Ingresos'
    pct = c(wb,S,'G22')
    vol_g = c(wb,S,'E45'); price_g = c(wb,S,'E48')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'explained by higher volumes and prices',
        'vol_only': 'mainly explained by higher volumes',
        'vol_despite_price': 'mainly explained by higher volumes',
        'price_only': 'mainly explained by higher prices',
        'price_despite_vol': 'mainly explained by higher prices',
        'both_down': 'explained by lower volumes and prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    share_cur = c(wb,S,'C27'); share_prior = c(wb,S,'E27')
    return (
        f"Value added products recorded {indef_article(pct)} {fmt_pct(pct)}% increase in sales income as of Jun26 compared to "
        f"income recorded as of Jun25, {phrase}. Income from the value added segment represented "
        f"{fmt_pct(share_cur)}% of income as of Jun26 ({fmt_pct(share_prior)}% as of Jun25)."
    )

def p_ingresos_totales_12m(wb):
    S='Ingresos'
    cur=c(wb,S,'I4'); prior=c(wb,S,'K4'); pct=c(wb,S,'M4')
    vol_g = c(wb,S,'E54'); price_g = c(wb,S,'E57')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'due to an increase in sales volumes and prices',
        'vol_only': 'due to an increase in sales volumes',
        'vol_despite_price': 'due to an increase in sales volumes',
        'price_only': 'due to an increase in sale prices',
        'price_despite_vol': 'due to an increase in sale prices',
        'both_down': 'due to a decrease in sales volumes and prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    return (
        f"Income from operating activities reached US${fmt_millones(cur, force_decimals=0)} million in the "
        f"12M S25/26, representing {indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct)} compared to the "
        f"US${fmt_millones(prior, force_decimals=0)} million in the 12M S24/25, {phrase}."
    )

def p_frutafresca_12m(wb):
    S='Ingresos'
    pct = c(wb,S,'M15')
    vol_g = c(wb,S,'E60'); price_g = c(wb,S,'E63')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'mainly explained by higher commercialized volumes and prices',
        'vol_only': 'mainly explained by higher commercialized volumes',
        'vol_despite_price': 'mainly explained by higher commercialized volumes',
        'price_only': 'mainly explained by higher prices',
        'price_despite_vol': 'mainly explained by higher prices',
        'both_down': 'mainly explained by lower commercialized volumes and prices',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    verbo = 'increased' if pct>=0 else 'decreased'
    return (
        f"The Fresh Fruit sales segment in the 12M S25/26 {verbo} {fmt_pct(pct)}% compared to the 12M S24/25, "
        f"{phrase}."
    )

def p_valoragregado_12m(wb):
    S='Ingresos'
    pct = c(wb,S,'M22')
    vol_g = c(wb,S,'E67'); price_g = c(wb,S,'E70')
    driver = volumen_precio_driver(vol_g, price_g)
    phrase = {
        'both': 'explained both by higher volumes in this segment and higher prices',
        'vol_only': 'mainly explained by higher volumes in this segment',
        'vol_despite_price': 'mainly explained by higher volumes in this segment',
        'price_only': 'mainly explained by higher prices',
        'price_despite_vol': 'mainly explained by higher prices',
        'both_down': 'explained by lower volumes and prices in this segment',
        'unclear': '[REVIEW: volume/price driver inconclusive]',
    }[driver]
    share_cur = c(wb,S,'I27'); share_prior = c(wb,S,'K27')
    return (
        f"Also, the value added products segment recorded {indef_article(pct)} {fmt_pct(pct)}% increase in income during the "
        f"12M S25/26, compared with the 12M S24/25, {phrase}. This segment represented {fmt_pct(share_cur)}% of "
        f"income in the 12M S25/26, compared with the {fmt_pct(share_prior)}% represented in the 12M S24/25."
    )

# ============================================================
# SECTION: COSTS AND EXPENSES
# ============================================================
def p_costoventa_6m(wb):
    S='Costos'
    cur=abs(c(wb,S,'D6')); prior=abs(c(wb,S,'F6')); pct=c(wb,S,'H6')
    ratio_cur=c(wb,S,'D27'); ratio_prior=c(wb,S,'F27')
    cost_per_kg_cur=c(wb,S,'D36'); cost_per_kg_prior=c(wb,S,'F36')
    cpk_up = cost_per_kg_cur < cost_per_kg_prior * 1.0 and abs(cost_per_kg_cur) > abs(cost_per_kg_prior) * 1.03
    driver = "associated to the growth in volumes and the higher unit cost" if cpk_up else "associated to the growth in volumes"
    return (
        f"Sales costs as of Jun26 totaled US${fmt_millones(cur, force_decimals=0)} million, presenting "
        f"{indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct)} compared to the US${fmt_millones(prior, force_decimals=0)} million "
        f"recorded as of Jun25, {driver}. Sales costs represented {fmt_pct(abs(ratio_cur))}% of income from "
        f"operating activities as of Jun26, while as of Jun25 they reached {fmt_pct(abs(ratio_prior))}%."
    )

def p_costoventa_12m(wb, cause=None):
    S='Costos'
    cur=abs(c(wb,S,'J6')); prior=abs(c(wb,S,'L6')); pct=c(wb,S,'N6')
    ratio_cur=c(wb,S,'J27'); ratio_prior=c(wb,S,'L27')
    base = (
        f"Sales costs in the 12M S25/26 reached US${fmt_millones(cur, force_decimals=0)} million, representing "
        f"{indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct)} compared to the US${fmt_millones(prior, force_decimals=0)} million "
        f"recorded in the 12M S24/25, also explained by higher commercialized volumes. Sales costs represented "
        f"{fmt_pct(abs(ratio_cur))}% of income from operating activities in the 12M S25/26, compared with "
        f"{fmt_pct(abs(ratio_prior))}% in the 12M S24/25."
    )
    return _with_cause(base, cause)

def p_gastosadmin_6m(wb, cause=None):
    S='Costos'
    cur=abs(c(wb,S,'D8')); prior=abs(c(wb,S,'F8')); pct=c(wb,S,'H8')
    base = (
        f"Administrative expenses as of Jun26 reached US${fmt_millones(cur)} million, representing "
        f"{indef_article(pct)} {fmt_pct(pct)}% {inc_dec(pct)} compared to Jun25."
    )
    return _with_cause(base, cause)

def p_otrosgastos_6m(wb):
    S='Costos'
    cur=abs(c(wb,S,'D9')); prior=abs(c(wb,S,'F9'))
    return (
        f"Other expenses, per function (excluding impairment in the value of assets) reached "
        f"US${fmt_millones(cur)} million as of Jun26, compared with US${fmt_millones(prior)} million as of Jun25."
    )

def p_deterioro_6m(wb, cause_cur=None, cause_prior=None):
    S='Costos'
    cur=abs(c(wb,S,'D13')); prior=abs(c(wb,S,'F13'))
    frase_cur = _with_cause(f"The expense due to impairment in the value of assets as of Jun26 was "
                             f"US${fmt_millones(cur)} million", cause_cur)
    frase_prior = _with_cause(f"compared with US${fmt_millones(prior)} million as of Jun25", cause_prior)
    return f"{frase_cur}, {frase_prior}"

def p_otroscomponentes_6m(wb):
    S='Otros ingresos(gastos)'
    cur=abs(c(wb,S,'D13')); prior=abs(c(wb,S,'F13'))
    GA='Ganancia Atribuible'
    cf_cur=c(wb,GA,'C8'); cf_prior=c(wb,GA,'E8'); cf_pct=c(wb,GA,'L8')
    return (
        f"The other components of the income statement recorded {indef_article_millones(cur)} US${fmt_millones(cur)} million cost as of "
        f"Jun26, compared with the US${fmt_millones(prior)} million as of Jun25. The main non-operating item is "
        f"net financial costs, which reached US${fmt_millones(abs(cf_cur))} million as of Jun26, compared with "
        f"US${fmt_millones(abs(cf_prior))} million as of Jun25, recording {indef_article(cf_pct)} {fmt_pct(cf_pct)}% {inc_dec(cf_pct)}."
    )

def p_impuesto_6m(wb):
    S='Ganancia Atribuible'
    cur=c(wb,S,'C11'); prior=c(wb,S,'E11')
    return (
        f"As of Jun26, a gains tax expense of US${fmt_millones(abs(cur))} million was recorded, compared with "
        f"the US${fmt_millones(abs(prior))} million recorded as of Jun25."
    )

# ============================================================
# SECTION: ACTIVITY INDICATORS
# ============================================================
def p_rotacion_activos(wb, periods=None):
    S='Ind. Actividad'
    cur=c(wb,S,'D5'); prior=c(wb,S,'E5'); ing_pct=c(wb,S,'J3')
    cur_long = periods['CUR_LONG'] if periods else 'June 30, 2026'
    verbo = 'increased' if cur-prior>=0 else 'decreased'
    return (
        f"The rotation of assets as of {cur_long} {verbo} from {fmt_veces(prior)} times as of Jun25 to "
        f"{fmt_veces(cur)} times as of Jun26, mainly explained by the {fmt_pct(ing_pct)}% increase in income."
    )

def p_rotacion_inventarios(wb):
    S='Ind. Actividad'
    cur=c(wb,S,'D7'); prior=c(wb,S,'E7'); inv_prom_pct=c(wb,S,'J8')
    verbo = 'increased' if cur-prior>=0 else 'decreased'
    nivel = 'higher' if inv_prom_pct>=0 else 'lower'
    return (
        f"Also, the rotation ratio of inventories {verbo} from {fmt_veces(prior)} times as of Jun25 to "
        f"{fmt_veces(cur)} times as of Jun26, explained by {nivel} average inventory levels."
    )

# ============================================================
# SECTION: FINANCIAL INDICATORS AND PROFITABILITY
# ============================================================
def p_liquidez(wb):
    IF='Ind. Financieros'; BAL='Balance'
    cur=c(wb,IF,'E7'); prior=c(wb,IF,'F7'); pct=c(wb,IF,'G7')
    ac_pct=c(wb,BAL,'H7'); pc_pct=c(wb,BAL,'H11')
    leve = 'slight ' if abs(pct) < 0.02 else ''
    return (
        f"Current liquidity was {fmt_veces(cur)} times as of Jun26, which represents a {leve}"
        f"{inc_dec(pct,'increase','reduction')} of {fmt_pct(abs(pct))}% compared to Dec25 ({fmt_veces(prior)} times), "
        f"explained by a greater {inc_dec(ac_pct,'increase','reduction')} in current assets of the period "
        f"({fmt_pct(ac_pct)}%) in relation to the {inc_dec(pc_pct,'increase','reduction')} in current liabilities "
        f"({fmt_pct(pc_pct)}%)."
    )

def p_razon_acida(wb):
    S='Ind. Financieros'
    cur=c(wb,S,'E9'); prior=c(wb,S,'F9')
    return (
        f"In the meantime, the acid ratio reached {fmt_veces(cur)} times, experiencing an "
        f"{inc_dec(cur-prior,'increase','decrease')} in relation to Dec25."
    )

def p_razon_endeudamiento(wb, cause=None):
    IF='Ind. Financieros'; BAL='Balance'
    cur=c(wb,IF,'E11'); pct=c(wb,IF,'G11')
    pt_pct=c(wb,BAL,'H13')
    base = (
        f"The debt ratio {inc_dec(pct,'increased','decreased')} {fmt_pct(abs(pct))}% compared to Dec25, reaching "
        f"{fmt_veces(cur)} times, mainly explained by the {fmt_pct(abs(pt_pct))}% "
        f"{inc_dec(pt_pct,'increase','reduction')} in total liabilities"
    )
    return _with_cause(base, cause)

def p_estructura_deuda(wb, periods=None):
    S='Ind. Financieros'
    cur_cp=c(wb,S,'E13'); prior_cp=c(wb,S,'F13'); cur_lp=c(wb,S,'E15'); prior_lp=c(wb,S,'F15')
    mantuvo = abs(cur_cp-prior_cp) < 0.03
    verbo = 'maintained' if mantuvo else 'modified'
    fy_prior_long = periods['FY_PRIOR_LONG'] if periods else 'December 31, 2025'
    return (
        f"During this period, the company {verbo} its debt structure. The percentage of current liabilities as "
        f"of Jun26 was {fmt_pct(cur_cp)}% compared to total liabilities, compared with the {fmt_pct(prior_cp)}% "
        f"recorded as of {fy_prior_long}. The proportion of long-term liabilities went from {fmt_pct(prior_lp)}% "
        f"as of Dec25 to {fmt_pct(cur_lp)}% as of Jun26."
    )

def p_cobertura_gf(wb, cause=None):
    IR='Ind. Rentabilidad'
    cur=c(wb,IR,'C2'); prior=c(wb,IR,'D2')
    base = (
        f"The financial expense hedge index is at {fmt_veces(cur)} times as of Jun26, "
        f"{inc_dec(cur-prior,'improving','worsening')} from the {fmt_veces(prior)} times as of Jun25, due to "
        f"higher before tax profit. Also, financial costs as of Jun26 show an increase compared to Jun25"
    )
    return _with_cause(base, cause)

def p_rentabilidad_patrim_controladora(wb):
    IR='Ind. Rentabilidad'; GA='Ganancia Atribuible'
    cur=c(wb,IR,'C4'); prior=c(wb,IR,'D4')
    gc_cur=c(wb,GA,'C13'); gc_prior=c(wb,GA,'E13')
    return (
        f"The profitability of the parent company equity {inc_dec(cur-prior,'improved','worsened')} from "
        f"{fmt_pct(prior)}% as of Jun25 to {fmt_pct(cur)}% as of Jun26, explained by the "
        f"{inc_dec(gc_cur-gc_prior,'increase','decrease')} in controlling profit, which went from "
        f"US${fmt_millones(gc_prior)} million as of Jun25 to US${fmt_millones(gc_cur)} million as of Jun26."
    )

def p_rentabilidad_patrim_total(wb):
    IR='Ind. Rentabilidad'
    cur=c(wb,IR,'C6'); prior=c(wb,IR,'D6')
    ganancia_ejercicio=c(wb,IR,'L10')
    return (
        f"Also, the profitability of total equity as of Jun26 is at {fmt_pct(cur)}%, compared with the "
        f"{fmt_pct(prior)}% profitability as of Jun25, explained by the profit of the period of "
        f"US${fmt_millones(ganancia_ejercicio)} million as of Jun26."
    )

# ============================================================
# SECTION: NET FINANCIAL DEBT
# ============================================================
def p_dfn(wb, periods=None):
    S='DF Neta'
    cur=c(wb,S,'C15'); prior=c(wb,S,'E15')
    arr_cur=c(wb,S,'H8'); arr_prior=c(wb,S,'I8')
    ifrs16_cur=c(wb,S,'C28'); ifrs16_prior=c(wb,S,'E28')
    leasing_cur=c(wb,S,'C22'); leasing_prior=c(wb,S,'E22')
    cur_long = periods['CUR_LONG'] if periods else 'June 30, 2026'
    fy_prior_long = periods['FY_PRIOR_LONG'] if periods else 'December 31, 2025'
    return (
        f"The Company’s net financial debt was reduced from US${fmt_millones(prior)} million as of "
        f"{fy_prior_long} to US${fmt_millones(cur)} million as of {cur_long}. As of {cur_long}, the lease "
        f"liability reached US${fmt_millones(arr_cur)} million, of which US${fmt_millones(ifrs16_cur)} million "
        f"correspond to lease liabilities under IFRS16 and US${fmt_millones(leasing_cur)} million were associated "
        f"to recorded leasing liabilities. Also, as of {fy_prior_long}, the lease liability reached "
        f"US${fmt_millones(arr_prior)} million, of which US${fmt_millones(ifrs16_prior)} million corresponded to "
        f"lease liabilities under IFRS16 and US${fmt_millones(leasing_prior)} million were associated to leasing "
        f"liabilities."
    )

# ============================================================
# SECTION: STATEMENT OF FINANCIAL POSITION (Balance bridge)
# ============================================================
FECU = 'EEFF Presentación Fecu'

def _fecu_row(wb, label, start=1, end=176):
    ws = wb[FECU]
    for r in range(start, end+1):
        if ws.cell(r, 3).value == label:
            cur = ws.cell(r,5).value; prior = ws.cell(r,7).value
            delta = ws.cell(r,9).value
            if delta is None and cur is not None and prior is not None:
                delta = cur - prior
            return dict(cur=cur, prior=prior, delta=delta)
    raise KeyError(label)

def p_activos_totales_intro(wb, periods=None):
    d = _fecu_row(wb, 'Total Activos')
    pct = d['delta']/d['prior']
    cur_long = periods['CUR_LONG'] if periods else 'June 30, 2026'
    fy_prior_long = periods['FY_PRIOR_LONG'] if periods else 'December 31, 2025'
    verbo = 'increased' if d['delta']>=0 else 'decreased'
    return (
        f"As of {cur_long}, total assets {verbo} by US${fmt_millones(abs(d['delta']))} million "
        f"({fmt_pct(pct)}%) compared to those existing as of {fy_prior_long}, mainly explained by:"
    )

ACTIVOS_NC_LABELS = [
    'Otros activos financieros, no corrientes', 'Otros activos no financieros, no corrientes',
    'Derechos por cobrar, no corrientes', 'Cuentas por cobrar a entidades relacionadas, no corrientes',
    'Inversiones contabilizadas utilizando el método de la participación',
    'Activos intangibles distintos de la plusvalía', 'Plusvalía', 'Propiedades, planta y equipo',
    'Activos por Derecho de Uso', 'Activos por impuestos diferidos',
]
ACTIVOS_C_LABELS = [
    'Efectivo y equivalentes al efectivo', 'Otros activos financieros, corrientes',
    'Otros activos no financieros, corrientes', 'Deudores comerciales y otras cuentas por cobrar, corrientes',
    'Cuentas por cobrar a entidades relacionadas, corrientes', 'Inventarios', 'Activos biológicos, corrientes',
    'Activos por impuestos corrientes',
]
# Traducción al inglés de cada etiqueta (claves en español = las mismas usadas para leer el Excel vía
# _fecu_row; no cambian). Términos financieros estándar.
LABEL_EN_FULL = {
    'Otros activos financieros, no corrientes': 'Other non-current financial assets',
    'Otros activos no financieros, no corrientes': 'Other non-current non-financial assets',
    'Derechos por cobrar, no corrientes': 'Non-current receivables',
    'Cuentas por cobrar a entidades relacionadas, no corrientes': 'Non-current accounts receivable from related entities',
    'Inversiones contabilizadas utilizando el método de la participación': 'Investments accounted for using the equity method',
    'Activos intangibles distintos de la plusvalía': 'Intangible assets other than goodwill',
    'Plusvalía': 'Goodwill',
    'Propiedades, planta y equipo': 'Property, Plant and Equipment',
    'Activos por Derecho de Uso': 'Right-of-use assets',
    'Activos por impuestos diferidos': 'Deferred tax assets',
    'Efectivo y equivalentes al efectivo': 'Cash and cash equivalents',
    'Otros activos financieros, corrientes': 'Other current financial assets',
    'Otros activos no financieros, corrientes': 'Other current non-financial assets',
    'Deudores comerciales y otras cuentas por cobrar, corrientes': 'Trade and other receivables, current',
    'Cuentas por cobrar a entidades relacionadas, corrientes': 'Accounts receivable from related entities, current',
    'Inventarios': 'Inventories',
    'Activos biológicos, corrientes': 'Current biological assets',
    'Activos por impuestos corrientes': 'Current tax assets',
}
SHORT_EN = {
    'Plusvalía': 'Goodwill accounts',
    'Activos intangibles distintos de la plusvalía': 'Intangible assets',
    'Cuentas por cobrar a entidades relacionadas, corrientes': 'Accounts receivable from related entities',
    'Inventarios': 'inventory level',
}

def _top_items(wb, labels, top_n=3):
    items = [(lbl, _fecu_row(wb, lbl)) for lbl in labels]
    items.sort(key=lambda x: abs(x[1]['delta']), reverse=True)
    return items[:top_n]

def p_activos_nocorrientes_bridge(wb, cause=None):
    total = _fecu_row(wb, 'Total Activos no corrientes')
    top = _top_items(wb, ACTIVOS_NC_LABELS, 3)
    ppe_label, ppe = top[0]
    ppe_label_en = LABEL_EN_FULL.get(ppe_label, ppe_label)
    rest_parts = []
    for lbl, d in top[1:]:
        short_en = SHORT_EN.get(lbl, LABEL_EN_FULL.get(lbl, lbl))
        rest_parts.append(f"a reduction in {short_en} (US${fmt_millones(d['delta'])} million)" if d['delta']<0
                           else f"an increase in {short_en} (US$+{fmt_millones(d['delta'])} million)")
    rest = ', '.join(rest_parts)
    verbo = 'Reduction' if total['delta'] < 0 else 'Increase'
    lead = f"{verbo} in non-current assets of US${fmt_millones(abs(total['delta']))} million compared to Dec25:"
    compensado_clause = "offset by investments also made during the period" if not cause else \
        _with_cause("offset by investments also made during the period", cause)
    return (
        "{{" + lead + "}} "
        f"Mainly explained by the {inc_dec(ppe['delta'],'increase','reduction')} in {ppe_label_en} "
        f"(US${fmt_millones(ppe['delta'])} million) associated to the depreciation of the period, {compensado_clause}. "
        f"Additionally, it is explained by {rest}, among other minor variations."
    )

def p_activos_corrientes_bridge(wb):
    total = _fecu_row(wb, 'Total de activos corrientes distintos de los activos o grupos de activos para su disposición clasificados como mantenidos para la venta')
    top = _top_items(wb, ACTIVOS_C_LABELS, 3)
    negs = [(lbl,d) for lbl,d in top if d['delta']<0]
    poss = [(lbl,d) for lbl,d in top if d['delta']>=0]
    neg_parts = []
    for lbl, d in negs:
        short_en = SHORT_EN.get(lbl, LABEL_EN_FULL.get(lbl, lbl))
        tag = 'lower' if 'level' not in short_en else 'a lower'
        neg_parts.append(f"{tag} {short_en} (US${fmt_millones(d['delta'])} million)")
    pos_parts = [f"higher {LABEL_EN_FULL.get(lbl, lbl)} (US$+{fmt_millones(d['delta'])} million)" for lbl, d in poss]
    verbo = 'Reduction' if total['delta'] < 0 else 'Increase'
    lead = f"{verbo} in current assets of US${fmt_millones(abs(total['delta']))} million compared to Dec25:"
    return (
        "{{" + lead + "}} "
        f"Mainly explained by {' and '.join(neg_parts)}, mainly associated to the seasonality of the business. "
        f"The above was partially offset by {' and '.join(pos_parts)}."
    )

PASIVOS_GROUPS = {
    'pasivos financieros, corrientes y no corrientes': [
        ('Otros pasivos financieros, corrientes', 1), ('Otros pasivos financieros, no corrientes', 1)],
    'cuentas comerciales por pagar, corrientes y no corrientes': [
        ('Cuentas comerciales y otras cuentas por pagar, corrientes', 1), ('Otras cuentas por pagar, no corrientes', 1)],
    'pasivos por arrendamientos': [
        ('Pasivos por arrendamientos, corrientes', 1), ('Pasivos por arrendamientos, no corrientes', 1)],
}
PASIVOS_GROUPS_EN = {
    'pasivos financieros, corrientes y no corrientes': 'current and non-current financial liabilities',
    'cuentas comerciales por pagar, corrientes y no corrientes': 'current and non-current trade accounts payable',
    'pasivos por arrendamientos': 'lease liabilities',
}

def p_pasivos_bridge(wb):
    total = _fecu_row(wb, 'Total pasivos')
    group_deltas = {}
    for g, rows in PASIVOS_GROUPS.items():
        tot = 0
        for lbl, _ in rows:
            tot += _fecu_row(wb, lbl)['delta']
        group_deltas[g] = tot
    ordered = sorted(group_deltas.items(), key=lambda x: abs(x[1]), reverse=True)
    main, rest = ordered[0], ordered[1:]
    rest_txt = ', '.join(
        f"lower {PASIVOS_GROUPS_EN[g]} (US${fmt_millones(v)} million)" if v<0 else
        f"higher {PASIVOS_GROUPS_EN[g]} (US$+{fmt_millones(v)} million)" for g,v in rest)
    verbo = 'decreased' if total['delta'] < 0 else 'increased'
    main_verbo = 'reduction' if main[1] < 0 else 'increase'
    bold_mid = f"total liabilities {verbo} by US${fmt_millones(abs(total['delta']))} million"
    return (
        "{{" + bold_mid + "}}, mainly explained "
        f"by the {main_verbo} in {PASIVOS_GROUPS_EN[main[0]]} (US${fmt_millones(main[1])} million), explained by "
        f"the seasonality of the business. Additionally, by {rest_txt}, also associated to seasonality."
    )

def p_patrimonio_bridge(wb, periods=None):
    patr = _fecu_row(wb, '   Total Patrimonio')
    ganancia_cur = c(wb, FECU, 'E119')
    fy_prior_long = periods['FY_PRIOR_LONG'] if periods else 'December 31, 2025'
    verbo = 'increased' if patr['delta']>=0 else 'decreased'
    bold_mid = f"Company’s total equity {verbo} by US${fmt_millones(abs(patr['delta']))} million"
    return (
        "The {{" + bold_mid + "}} compared to " + fy_prior_long + ", totaling "
        f"US${fmt_millones(patr['cur'])} million, mainly explained by the profit of the period of "
        f"US${fmt_millones(ganancia_cur)} million."
    )

# ============================================================
# SECTION: PREAMBLE AND SECTION HEADINGS (depend only on the period, not the Excel data)
# ============================================================
def p_titulo_fecha(periods):
    return f"As of {periods['CUR_LONG']}"

def p_preambulo_1(periods):
    return (
        f"The current reasoned analysis has been prepared for the period ending {periods['CUR_LONG']}, "
        f"compared with the financial statements as of {periods['PREAMBULO_REF_LONG']} ({periods['CUR_SHORT']} "
        f"and {periods['PREAMBULO_REF_SHORT']}, respectively)."
    )

def p_preambulo_2(periods):
    season_cur_bare = periods['SEASON_CUR'][1:]
    season_prior_bare = periods['SEASON_PRIOR'][1:]
    return (
        "Since the Company administers its operations with an agricultural season (July 01 to June 30) point "
        "of view, which is the relevant criteria for this type of business, in this analysis we also include "
        f"the twelve-month comparison of the {season_cur_bare} and {season_prior_bare} seasons "
        f"(“{periods['SEASON_CUR']}” and “{periods['SEASON_PRIOR']}”, respectively)."
    )

def p_header_ebitda_acumulado(periods):
    return f"Accumulated EBITDA analysis as of {periods['CUR_MONTH_YEAR']}"

def p_header_resultado_calendario(periods):
    return f"Income Statement analysis as of {periods['CUR_MONTH_YEAR']}"

def p_header_resultado_temporada(periods):
    return f"Income Statement Analysis {periods['SEASON_RANGE_LONG']} season"

def p_header_ingresos_acumulados(periods):
    return f"Accumulated income analysis as of {periods['CUR_MONTH_CAP']}"
