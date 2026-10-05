"""RANMAN «Flujo y PLP DRA - MAQUINA - dd-mm-aa.xlsx» -> the Flujo y PLP tablero's
flujo_plp.json shape.

Faithful port of Valoran's actualizar_flujo_plp.py (Sep-2026), working from
bytes so the workbook can be uploaded through the app. Also validates a
flujo_plp.json produced by the original script.

Where each thing comes from (sheet «Resumen» is in thousands of pesos; the
tablero shows mdp). Ranman adds and drops Resumen rows from month to month
(Jun-2026 still had «Intereses Valoran» and «Capital Valoran», Sep-2026 added
three land lines and «Flujo Ranman»), so every block is found by its label:
    Resumen  «Operativos» … «Total Operativos»                   aplicación · operativos
             «Deuda y Proyectos no Operativos» … «Total …»        aplicación · deuda y no operativos
             «Inversión» … «Total Inversión»                      aplicación · inversión
             «Origen» … «Total Origen»                            origen
             «Total Aplicación», «Total Origen», and «Flujo Ranman»
             (earlier: the «Saldo (Necesidades)» right after Total Origen)   control totals
             «Deuda Total DRA» … «Saldo (Necesidades)»             corporate debt balances
             col A   on a line: 0/1 flag (2+ picks a mode); elsewhere a parameter
                     (the sale percentages under «Porcentajes de Venta de ML»)
    PLP      rows  4-82    income statement and cash flow
             cols  C:EP    monthly block (144 months)
                   KG:KR   yearly block (12 years)

Every Resumen line is re-computed as if its own flag were 1 (its "base"
amount), with every other row as the file has it, so the tablero can switch on
lines the file has off; a line whose flag picks a scenario other than 1
(«1 Conservador - 2 Pesimista») is taken as the file worked it out. The
validation then checks that, with the flags as they come, the recomputed totals
match the file's own.
"""
import datetime as dt
import io
import json
import re

import openpyxl

# the four groups of lines, by the label that opens each one on the Resumen sheet
GRUPOS = [('op', 'Operativos', 'operativos'),
          ('de', 'Deuda y proyectos no operativos', 'deuda y proyectos no operativos'),
          ('in', 'Inversión', 'inversion'),
          ('or', 'Origen', 'origen')]
MILES = 1000.0
RENOMBRES = {'Crédito Valoran': 'Crédito VRM'}   # how each line is called in the tablero

PLP_R0, PLP_R1 = 4, 82
COLS_MES = range(3, 147)      # C .. EP
COLS_ANIO = range(293, 305)   # KG .. KR
FILAS_CUENTA = {4, 5}         # apartados y firmas: units, not thousands


def _si(cond, a, b=0.0):
    """Excel's IF for the evaluator (both branches are already numbers)."""
    return a if cond else b


def hoja(wb, nombre):
    """The sheet called `nombre`, ignoring case and surrounding spaces (Feb-2026
    names it «Resumen »); None when there is none. «Resumen Base» is not «Resumen»."""
    for s in wb.sheetnames:
        if s.strip().lower() == nombre.lower():
            return wb[s]
    return None


def fecha_de_nombre(nombre: str):
    """'Flujo y PLP DRA - MAQUINA - 31-08-26.xlsx' -> '2026-08-31' (None if absent)."""
    m = re.search(r'(\d{2})-(\d{2})-(\d{2})\.xlsx$', nombre or '')
    if not m:
        return None
    d, mes, a = (int(x) for x in m.groups())
    try:
        return dt.date(2000 + a, mes, d).isoformat()
    except ValueError:
        return None


# ── sheet «Resumen»: origen, aplicación y deuda ─────────────────────────────
def _norm(t):
    import unicodedata
    t = unicodedata.normalize('NFKD', str(t or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', t).strip().lower()


def _etiqueta(sv, r):
    """The row's label: the text in columns B-F (the amounts start further right)."""
    return ' '.join(str(v).strip() for v in (sv.cell(r, c).value for c in range(2, 7))
                    if isinstance(v, str) and v.strip())


def _es_bandera(v):
    """A line's on/off switch (0/1, or 2+ for a mode): a small whole number in column A."""
    return isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 0 <= v <= 9


def estructura(sv):
    """Where the blocks of this month's Resumen are: {'grupos': {cod: [rows]},
    'aplic', 'origen', 'flujo', 'deuda': [rows], 'saldo'}. Raises ValueError
    when a block the tablero needs is not there."""
    abre = {clave: cod for cod, _, clave in GRUPOS}
    grupos = {cod: [] for cod, _, _ in GRUPOS}
    out = {'grupos': grupos, 'aplic': None, 'origen': None, 'flujo': None, 'deuda': [], 'saldo': None}
    actual, en_deuda = None, False
    for r in range(1, (sv.max_row or 0) + 1):
        t = _norm(_etiqueta(sv, r))
        if not t:
            continue
        if en_deuda:
            if t.startswith('saldo (necesidades)'):
                out['saldo'], en_deuda = r, False
            else:
                out['deuda'].append(r)
            continue
        if t in abre:
            actual = abre[t]
            continue
        if t.startswith('total'):
            if t == 'total aplicacion':
                out['aplic'] = r
            elif t == 'total origen':
                out['origen'] = r
            actual = None
            continue
        if t == 'deuda total dra':
            en_deuda, actual = True, None
            continue
        if out['origen'] and out['flujo'] is None and (t == 'flujo ranman' or t.startswith('saldo (necesidades)')):
            out['flujo'] = r          # «Flujo Ranman», or before Sep-2026 the Saldo right after Total Origen
            continue
        if t == 'flujo ranman':
            out['flujo'] = r
            continue
        if actual and _es_bandera(sv.cell(r, 1).value):
            grupos[actual].append(r)
    falta = [n for n, k in (('Total Aplicación', 'aplic'), ('Total Origen', 'origen'),
                            ('Flujo Ranman / Saldo (Necesidades)', 'flujo')) if not out[k]]
    falta += [n for cod, n, _ in GRUPOS if not grupos[cod]]
    if falta:
        raise ValueError("Resumen: no encuentro %s — ¿cambió el formato de la hoja?" % ', '.join(falta))
    return out


def lee_resumen(wf, wv, as_of: str) -> dict:
    sf, sv, flv = hoja(wf, 'Resumen'), hoja(wv, 'Resumen'), hoja(wv, 'Flujo')
    E = estructura(sv)

    # row 1 carries the monthly block and then yearly columns: cut where dates go back
    hdr = list(sv.iter_rows(min_row=1, max_row=1, max_col=133, values_only=True))[0]
    fe = [(j, c) for j, c in enumerate(hdr) if isinstance(c, dt.datetime)]
    for k in range(1, len(fe)):
        if fe[k][1] < fe[k - 1][1]:
            fe = fe[:k]
            break
    COLS = [j + 1 for j, _ in fe]
    MESES = [c.strftime('%Y-%m') for _, c in fe]

    # the series starts at the corte's month: earlier columns come as zeros
    corte = as_of[:7]
    if MESES and MESES[0] < corte and corte in MESES:
        k = MESES.index(corte)
        COLS, MESES = COLS[k:], MESES[k:]

    LINEAS = {r for rs in E['grupos'].values() for r in rs}
    memo = {}
    num = lambda v: v if isinstance(v, (int, float)) else 0.0

    def base(r, c, prof=0):
        """The row's value as if its own flag were 1 (every other row as the file has it)."""
        if prof > 12:
            return 0.0
        k = (r, c)
        if k in memo:
            return memo[k]
        memo[k] = 0.0
        f = sf.cell(r, c).value
        flag = sv.cell(r, 1).value
        if r in LINEAS and _es_bandera(flag) and flag >= 2:
            # the switch picks a scenario other than 1 («1 Conservador - 2 Pesimista»):
            # the line is on, in that scenario, and the file already worked it out
            v = num(sv.cell(r, c).value)
        elif f is None:
            v = num(sv.cell(r, c).value)
        elif isinstance(f, (int, float)):
            v = float(f)
        else:
            s = str(f).strip().lstrip('=').lstrip('+')
            m = re.match(r'^IF\(\$A\$(\d+)=1,([^,]+),', s)    # its own flag as a mode selector
            if m and int(m.group(1)) == r:
                s = m.group(2)
            s = re.sub(r'Flujo!\$?([A-Z]{1,3})\$?(\d+)',
                       lambda m: repr(num(flv['%s%s' % (m.group(1), m.group(2))].value)), s)
            # $A$n is the line's own switch (taken as on), another line's switch
            # (as the file has it: Oct-2025 «Capital Valoran» is keyed to row 22's)
            # or a parameter such as the sale percentages (taken as it is)
            s = re.sub(r'\$A\$(\d+)',
                       lambda m: '1' if int(m.group(1)) == r and _es_bandera(sv.cell(r, 1).value)
                       else repr(num(sv.cell(int(m.group(1)), 1).value)), s)

            # the line's own earlier months recompute too (running sums); any other
            # row is taken as the file has it — a line that reads another line
            # (Bancrea «por operar» = 7,152.5 − the commercial-area sale) must see
            # that line switched as it comes, or a line that is on stops
            # matching the file
            def ref(m):
                col = openpyxl.utils.column_index_from_string(m.group(1))
                row = int(m.group(2))
                return repr(base(row, col, prof + 1) if row == r
                            else num(sv.cell(row, col).value))
            s = re.sub(r'\$?([A-Z]{1,3})\$?(\d+)', ref, s).replace('%', '/100')
            s = re.sub(r'(?<![<>!=])=(?!=)', '==', s.replace('<>', '!='))     # Excel comparisons
            try:
                v = float(eval(s, {'__builtins__': {}}, {'IF': _si}))
            except Exception:
                v = num(sv.cell(r, c).value)
        memo[k] = v
        return v

    fila = lambda r: [round(num(sv.cell(r, c).value) / MILES, 2) for c in COLS]

    lineas = []
    for cod, _, _ in GRUPOS:
        for r in E['grupos'][cod]:
            lab = _etiqueta(sv, r)
            if not lab:
                continue
            serie = [round(base(r, c) / MILES, 2) for c in COLS]
            lineas.append({'r': r, 'g': cod, 'n': lab,
                           'on': bool(sv.cell(r, 1).value not in (0, None)),
                           'v': serie,
                           'hay': any(abs(x) > 0.005 for x in serie)})

    deuda = {}
    for r in E['deuda']:
        nom = _etiqueta(sv, r)
        if nom:
            deuda[RENOMBRES.get(nom, nom)] = fila(r)

    return {'meses': MESES, 'lineas': lineas, 'deuda': deuda,
            'saldoNec': fila(E['saldo']) if E['saldo'] else [0.0] * len(COLS),
            'ref': {'aplic': fila(E['aplic']), 'origen': fila(E['origen']),
                    'flujo': fila(E['flujo'])}}


# ── sheet «PLP»: income statement and cash flow, rows 4 to 82 ───────────────
def lee_plp(wv, as_of: str) -> dict:
    ws = hoja(wv, 'PLP')
    meses = [ws.cell(3, c).value for c in COLS_MES]
    anios = [ws.cell(3, c).value for c in COLS_ANIO]
    if not all(isinstance(m, dt.datetime) for m in meses):
        raise ValueError('La hoja PLP no trae fechas en C3:EP3. ¿Se movieron las columnas?')

    # the tablero only shows the current year and the next
    a0 = int(as_of[:4])
    idx = [i for i, m in enumerate(meses) if m.year in (a0, a0 + 1)]
    num = lambda v: v if isinstance(v, (int, float)) else None

    filas = []
    for r in range(PLP_R0, PLP_R1 + 1):
        et = ws.cell(r, 2).value
        if et is None:
            filas.append({'r': r, 'sep': 1})
            continue
        fmt = ws.cell(r, 3).number_format or ''
        t = 'pct' if '%' in fmt else ('cuenta' if r in FILAS_CUENTA else 'mdp')

        def val(c, t=t, r=r):
            v = num(ws.cell(r, c).value)
            if v is None:
                return None
            return round(v / 1000, 2) if t == 'mdp' else (round(v, 4) if t == 'pct' else round(v))
        f = {'r': r, 'n': str(et).strip(), 'niv': ws.row_dimensions[r].outline_level or 0, 't': t,
             'a': [val(c) for c in COLS_ANIO],
             'm': [val(list(COLS_MES)[i]) for i in idx]}
        if ws.row_dimensions[r].hidden:
            f['plg'] = 1
        filas.append(f)

    return {'anios': [int(a) for a in anios],
            'meses': [meses[i].strftime('%Y-%m') for i in idx],
            'filas': filas}


# ── validation: recomputed totals vs the file's own ─────────────────────────
def valida(fl: dict, plp: dict):
    """(ok, lines) — every check as a human-readable line, like the script prints."""
    ok, out = True, []
    n = len(fl['meses'])
    grupos = {'aplic': {'op', 'de', 'in'}, 'origen': {'or'}}
    for clave, gs in grupos.items():
        calc = [sum(l['v'][i] for l in fl['lineas'] if l['g'] in gs and l['on']) for i in range(n)]
        dif = max(abs(a - b) for a, b in zip(calc, fl['ref'][clave])) if n else 0.0
        ok &= dif < 0.1
        out.append('%-14s dif máx contra el archivo: %6.2f mdp' % (clave, dif))
    calc = [sum(l['v'][i] for l in fl['lineas'] if l['on']) for i in range(n)]
    dif = max(abs(a - b) for a, b in zip(calc, fl['ref']['flujo'])) if n else 0.0
    ok &= dif < 0.1
    out.append('%-14s dif máx contra el archivo: %6.2f mdp' % ('flujo Ranman', dif))

    fs = plp['filas']
    for i, f in enumerate(fs):
        if f.get('sep') or f.get('niv', 1) != 0:
            continue
        h = []
        for j in range(i + 1, len(fs)):
            g = fs[j]
            if g.get('sep') or g.get('niv', 0) == 0:
                break
            h.append(g)
        if not h:
            continue
        sub = [x for x in h if x['n'] in ('Administración', 'Comercialización')] or h
        dif = max(abs((f['a'][k] or 0) - sum(x['a'][k] or 0 for x in sub))
                  for k in range(len(plp['anios'])))
        ok &= dif < 0.1
        out.append('PLP renglón %-3d %-42s dif máx %6.2f mdp' % (f['r'], f['n'][:42], dif))
    return bool(ok), out


# ── the year's opening plan (script: lee_base / primer_corte_del_anio) ──────
BASE_FILAS = {4: 'apartados', 5: 'firmas', 7: 'ventas', 39: 'uair'}


def parse_base_workbook(file_bytes: bytes, filename: str) -> dict:
    """The plan the year started with = the OLDEST PLP workbook of the year, read
    the way the script's lee_base does: rows 4/5/7/39 of the PLP sheet per year
    (yearly block found by the integer years in row 3) plus every month of the
    monthly block (dates in row 3, cut at the first date regression), so the
    year-to-date can be rebuilt for any later corte. Ventas and UAIR ÷1000 → mdp."""
    fecha = fecha_de_nombre(filename)
    if not fecha:
        raise ValueError("%s: could not read the date from the file name (expected '... dd-mm-aa.xlsx')" % filename)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    try:
        if hoja(wb, 'PLP') is None:
            raise ValueError("%s has no 'PLP' sheet" % filename)
        filas = list(hoja(wb, 'PLP').iter_rows(min_row=1, max_row=max(BASE_FILAS) + 1, max_col=340, values_only=True))
    finally:
        wb.close()
    cab = filas[2]
    cols = [j for j, v in enumerate(cab) if isinstance(v, int) and not isinstance(v, bool) and 2020 < v < 2050]
    if not cols:
        raise ValueError("%s: PLP row 3 has no yearly block" % filename)
    num = lambda v: v if isinstance(v, (int, float)) else None
    fechas = [(j, v) for j, v in enumerate(cab) if isinstance(v, dt.datetime)]
    for k in range(1, len(fechas)):
        if fechas[k][1] < fechas[k - 1][1]:
            fechas = fechas[:k]
            break
    out = {'archivo': filename, 'fecha': fecha, 'anios': [cab[j] for j in cols], 'filas': {}, 'mensual': {}}
    for r in BASE_FILAS:
        fila = filas[r - 1]
        esc = 1000.0 if r in (7, 39) else 1.0           # ventas and UAIR come in thousands
        out['filas'][str(r)] = [None if num(fila[j]) is None else round(num(fila[j]) / esc, 2) for j in cols]
        out['mensual'][str(r)] = {v.strftime('%Y-%m'): (num(fila[j]) or 0) / esc for j, v in fechas}
    return out


def base_para_corte(base: dict, as_of: str):
    """The plan as the script embeds it for one corte: yearly totals plus, for the
    corte's year, the January→corte-month accumulation (`ytd`) — comparable with
    what has really been sold and signed to date."""
    if not base:
        return None
    anio, hasta = as_of[:4], as_of[:7]
    out = {'archivo': base.get('archivo'), 'fecha': base.get('fecha'), 'anios': base.get('anios') or [],
           'filas': base.get('filas') or {}, 'ytd': {}, 'hasta': hasta, 'mes': {}}
    for r, meses in (base.get('mensual') or {}).items():
        sel = [v for m, v in meses.items() if m[:4] == anio and m <= hasta]
        if sel:
            out['ytd'][r] = round(sum(sel), 2)
        # the corte's year month by month (the Commercial tab charts the pace against it)
        out['mes'][r] = {m: round(v, 2) for m, v in sorted(meses.items()) if m[:4] == anio}
    return out


# ── by development: hidden sheet «PLP - Soporte» + the PLP's own split ─────
# «PLP - Soporte» repeats the statement once per development: a block whose
# first row carries the name in B and the months in C:EP (Jan-2024 → Dec-2035),
# then Apartados … Saldo Final de Efectivo, and a «Diferencia» row the file uses
# as its own check. It is Ranman's project-level support, NOT a split of the
# consolidated PLP: land and macrolote sales and other adjustments are booked
# per development only in the PLP sheet (rows 8-22 sales, 86-100 UAIR), so both
# are kept and the page shows the gap.
SOPORTE = 'PLP - Soporte'
SOPORTE_CUENTA = {'Apartados', 'Individualizaciones'}
SOPORTE_INICIAL = {'Saldo Inicial de Efectivo'}            # a year's value = its first month
SOPORTE_FINAL = {'Efectivo Sobrante antes de Dividendos', 'Saldo Final de Efectivo Ranman'}  # = its last month


def _bloques_soporte(ws):
    filas = list(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=150, values_only=True))
    inicios = [i for i, r in enumerate(filas) if len(r) > 2 and isinstance(r[2], dt.datetime)]
    if not inicios:
        raise ValueError("'%s' has no block headers (dates in column C)" % SOPORTE)
    cab = filas[inicios[0]]
    cols = [j for j in range(2, len(cab)) if isinstance(cab[j], dt.datetime)]
    meses = [cab[j].strftime('%Y-%m') for j in cols]
    out = []
    for k, s in enumerate(inicios):
        e = inicios[k + 1] if k + 1 < len(inicios) else len(filas)
        rows, dif = [], None
        for i in range(s + 1, e):
            et = filas[i][1]
            if not isinstance(et, str) or not et.strip():
                continue
            et = et.strip()
            vals = [filas[i][j] if isinstance(filas[i][j], (int, float)) else None for j in cols]
            if et == 'Diferencia':
                dif = max((abs(v) for v in vals if v is not None), default=0.0)
                break
            if et == 'Validación' or (et.startswith('Reserva ') and not any(vals)):
                continue
            rows.append((et, vals))
        if len(rows) >= 10:                    # «AGS - LA CARTUJA C00» is a two-row stub
            out.append({'n': str(filas[s][1]).strip(), 'on': filas[s][0] != 0, 'rows': rows, 'dif': dif})
    return meses, out


def _anualiza(rows, meses, anios):
    """Per year: flows add up, opening balances take January, closing balances
    December, and margins are recomputed from the row above over sales."""
    idx = {a: [i for i, m in enumerate(meses) if int(m[:4]) == a] for a in anios}
    ventas = next((v for et, v in rows if et == 'Ventas Totales'), None)
    tot = lambda vals, a: sum(vals[i] or 0 for i in idx[a])
    out, previo = {}, None
    for et, vals in rows:
        if '(%)' in et:
            out[et] = [round(tot(previo, a) / tot(ventas, a), 4) if previo and ventas and abs(tot(ventas, a)) > 0.5 else None
                       for a in anios]
            continue
        if et in SOPORTE_INICIAL:
            out[et] = [vals[idx[a][0]] if idx[a] else None for a in anios]
        elif et in SOPORTE_FINAL:
            out[et] = [vals[idx[a][-1]] if idx[a] else None for a in anios]
        else:
            out[et] = [tot(vals, a) for a in anios]
        previo = vals
    return out


def lee_proyectos(wv, as_of: str) -> dict:
    """Every development's statement (support sheet), with the consolidated PLP's
    own per-development sales and UAIR beside it. Yearly 2024-2035 and the
    corte's year month by month; mdp except units and margins."""
    from ranman_package import clave_desarrollo     # one key per development across reports
    a0 = int(as_of[:4])
    meses, bloques = _bloques_soporte(wv[SOPORTE])
    anios = sorted({int(m[:4]) for m in meses})
    im = [i for i, m in enumerate(meses) if int(m[:4]) == a0]

    def conv(et, v):
        if v is None:
            return None
        if '(%)' in et:
            return round(v, 4)
        return round(v) if et in SOPORTE_CUENTA else round(v / MILES, 2)

    proyectos = []
    for b in bloques:
        anual = _anualiza(b['rows'], meses, anios)
        filas = [{'n': et, 't': 'pct' if '(%)' in et else ('cuenta' if et in SOPORTE_CUENTA else 'mdp'),
                  'a': [conv(et, v) for v in anual[et]], 'm': [conv(et, vals[i]) for i in im]}
                 for et, vals in b['rows']]
        proyectos.append({'n': b['n'], 'k': clave_desarrollo(b['n']), 'on': b['on'],
                          'cuadra': b['dif'] is None or b['dif'] < 1.0,
                          'dif': None if b['dif'] is None else round(b['dif'] / MILES, 2), 'filas': filas})

    # the PLP sheet's per-development rows: sales under «Ventas Totales» (row 7)
    # and the «UAIR por Proyecto» block below the cash flow
    ws = hoja(wv, 'PLP')
    meses_plp = [ws.cell(3, c).value for c in COLS_MES]
    imp = [i for i, m in enumerate(meses_plp) if isinstance(m, dt.datetime) and m.year == a0]
    num = lambda v: v if isinstance(v, (int, float)) else None

    def fila(r):
        f = lambda c: None if num(ws.cell(r, c).value) is None else round(num(ws.cell(r, c).value) / MILES, 2)
        et = str(ws.cell(r, 2).value).strip()
        return {'n': et, 'k': clave_desarrollo(et), 'a': [f(c) for c in COLS_ANIO],
                'm': [f(list(COLS_MES)[i]) for i in imp]}
    ventas, uair, r = [], [], 8
    while r < 30 and str(ws.cell(r, 2).value or '').strip() not in ('', 'Costo de Ventas'):
        ventas.append(fila(r)); r += 1
    r0 = next((r for r in range(83, 130) if str(ws.cell(r, 2).value or '').strip() == 'UAIR por Proyecto'), None)
    total_uair = None
    if r0:
        total_uair = fila(r0 + 1)
        r = r0 + 2
        while r < r0 + 40 and str(ws.cell(r, 2).value or '').strip() not in ('', 'Validación'):
            uair.append(fila(r)); r += 1
    return {'asOf': as_of, 'anios': anios, 'meses': [meses[i] for i in im], 'proyectos': proyectos,
            'plp': {'anios': [int(ws.cell(3, c).value) for c in COLS_ANIO], 'ventas': ventas, 'uair': uair,
                    'uair_total': total_uair}}


# ── entry points ────────────────────────────────────────────────────────────
def parse_workbook(file_bytes: bytes, filename: str):
    """One «Flujo y PLP DRA - MAQUINA - dd-mm-aa.xlsx» -> (datos, ok, lines).

    datos has the tablero's flujo_plp.json shape. Raises ValueError when the
    date can't be read from the name or the sheets aren't there."""
    as_of = fecha_de_nombre(filename)
    if not as_of:
        raise ValueError("%s: could not read the date from the file name (expected '... dd-mm-aa.xlsx')" % filename)
    wf = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)   # formulas
    wv = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)    # cached values
    try:
        for nombre in ('Resumen', 'Flujo', 'PLP'):
            if hoja(wv, nombre) is None:
                raise ValueError("%s has no '%s' sheet — is it a Flujo y PLP DRA workbook?" % (filename, nombre))
        fl = lee_resumen(wf, wv, as_of)
        plp = lee_plp(wv, as_of)
        # by development: optional — a change in the support sheet must not block the corte
        proyectos, nota = None, None
        if SOPORTE in wv.sheetnames:
            try:
                proyectos = lee_proyectos(wv, as_of)
            except Exception as e:     # noqa: BLE001 — reported, never fatal
                nota = "'%s' could not be read (%s)" % (SOPORTE, e)
    finally:
        wf.close()
        wv.close()
    ok, lines = valida(fl, plp)
    if nota:
        lines.append(nota)
    datos = {'archivo': filename, 'asOf': as_of,
             'generado': dt.datetime.now().isoformat(timespec='seconds'),
             'flujo': fl, 'plp': plp}
    if proyectos:
        datos['proyectos'] = proyectos     # the app stores it apart from the tablero document
    return datos, ok, lines


def parse_json(file_bytes: bytes) -> dict:
    """A flujo_plp.json written by the original script; validated for shape."""
    data = json.loads(file_bytes.decode('utf-8'))
    if not isinstance(data, dict) or not data.get('asOf') or not isinstance(data.get('flujo'), dict) \
            or not isinstance(data.get('plp'), dict):
        raise ValueError('flujo_plp.json should carry asOf, flujo and plp.')
    dt.date.fromisoformat(data['asOf'])
    fl, plp = data['flujo'], data['plp']
    for k in ('meses', 'lineas', 'deuda', 'saldoNec', 'ref'):
        if k not in fl:
            raise ValueError("flujo_plp.json: flujo is missing '%s'." % k)
    for k in ('anios', 'meses', 'filas'):
        if k not in plp:
            raise ValueError("flujo_plp.json: plp is missing '%s'." % k)
    if not fl['meses'] or not fl['lineas'] or not plp['filas']:
        raise ValueError('flujo_plp.json has no data in it.')
    return data
