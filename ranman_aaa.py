"""RANMAN «Reporte AAA <Mes> aa.pdf» -> the monthly scorecard as data.

The AAA report is Ranman's own three-page monthly scorecard (Jan-2024 onward).
Page 1 carries every table: units, returns, P&L, margins and balance sheet on
the left (Target / Real Mes / Plan Mes / Real Acum / Plan Acum / % Var since
Oct-2024; Real Mes / Real Acum / Plan Acum / % Var / Total año anterior
before), and the cash flow, debt with cost by bank, debt and financial ratios
and the Z-score on the right. Below them, the land partners and Ranman's
numbered commentary.

The PDF has no table structure, only positioned text, so the parser rebuilds
it: words are grouped into rows by height; each section starts at a row whose
phrases are column headers; a number belongs to the header whose centre is
nearest. Digits the PDF spaces apart («1 70.1») are joined, and a minus sign
set apart from its number («- 53.5») is attached to it. Nothing is read by
fixed coordinates — the page size changed in Jul-2026 and the 2024 layout sits
elsewhere.

Checked against the report's own arithmetic before anything is stored: gross
profit less admin and selling expenses = EBIT, EBIT less financial cost =
UAIR, the banks add up to total debt, and the cash flow chains from opening
balance to the operating balance.
"""
import io
import re

from ranman_package import _norm

MESES = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7,
         'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12}

# a phrase is a column header when it reads like one of these
_COLUMNA = re.compile(r'^(target|real mes|plan mes|real acum|plan \d{4} acum|% var|total \d{4}|dic \d{2,4}|'
                      r'var vs \d{4}|var vs plan|acum|q\d \d{4}|\d{4}|real|% mezcla|por disponer|factor|'
                      r'ppto acum|plan \d{4})$')
# section titles -> ids
_SECCIONES = [('variables relevantes', 'variables'), ('rendimientos', 'rendimientos'),
              ('datos de resultados', 'resultados'), ('margenes', 'margenes'), ('datos de balance', 'balance'),
              ('cash flow', 'flujo'), ('deuda con costo', 'deuda'), ('razones de deuda', 'razones_deuda'),
              ('razones financieras', 'razones_fin'), ('z- score', 'zscore'), ('z-score', 'zscore')]
_NUM = re.compile(r'^-?[\d,]*\.?\d+%?$')


def _seccion(titulo):
    t = _norm(titulo)
    return next((sid for k, sid in _SECCIONES if t.startswith(k) or k in t), None)


def _filas(words, tol=3.0):
    """Words -> rows by height (a label may sit a point below its numbers)."""
    ws = sorted(words, key=lambda w: (w['top'], w['x0']))
    filas = []
    for w in ws:
        if filas and abs(w['top'] - filas[-1]['top']) <= tol:
            filas[-1]['w'].append(w)
        else:
            filas.append({'top': w['top'], 'w': [w]})
    for f in filas:
        f['w'].sort(key=lambda w: w['x0'])
    return filas


def _frases(ws, hueco=6.0):
    """Adjacent words -> phrases: [(text, x0, x1)]."""
    out = []
    for w in ws:
        if out and w['x0'] - out[-1][2] <= hueco:
            out[-1] = (out[-1][0] + ' ' + w['text'], out[-1][1], w['x1'])
        else:
            out.append((w['text'], w['x0'], w['x1']))
    return out


def _une_numeros(ws):
    """«1» «70.1» touching -> «170.1»; «2» «,316.1» -> «2,316.1»; «0» «.959» -> «0.959»."""
    out = []
    for w in ws:
        t = w['text']
        if out and re.fullmatch(r'[\d,.%]+', t) and re.fullmatch(r'-?[\d,.]+', out[-1]['text']) \
                and w['x0'] - out[-1]['x1'] <= 1.5:
            p = out[-1]
            out[-1] = dict(p, text=p['text'] + t, x1=w['x1'])
        else:
            out.append(dict(w))
    return out


def _valor(t):
    """'1,066.3' -> 1066.3 · '-18.6%' -> -0.186 · '-' / 'N/A' -> None. Returns (value, is_pct)."""
    t = t.strip()
    if t in ('-', 'N/A', 'n/a', '#DIV/0!'):
        return None, False
    pct = t.endswith('%')
    try:
        v = float(t.rstrip('%').replace(',', ''))
    except ValueError:
        return None, False
    return (v / 100.0 if pct else v), pct


def _celdas(ws, centros):
    """A row's numeric tokens -> one value per column (nearest header centre).
    A lone «-» is an empty cell; a «-» followed within 30 pt by a number signs it."""
    ws = _une_numeros(ws)
    toks = [w for w in ws if _NUM.match(w['text']) or w['text'] == '-' or w['text'] == 'N/A']
    vals, pcts = [None] * len(centros), [False] * len(centros)
    i = 0
    while i < len(toks):
        w = toks[i]
        if w['text'] == '-' and i + 1 < len(toks) and _NUM.match(toks[i + 1]['text']) \
                and toks[i + 1]['x0'] - w['x1'] <= 30 and not toks[i + 1]['text'].startswith('-'):
            n = toks[i + 1]
            w = dict(n, text='-' + n['text'])
            i += 1
        c = (w['x0'] + w['x1']) / 2.0
        k = min(range(len(centros)), key=lambda j: abs(centros[j] - c))
        v, p = _valor(w['text'])
        if v is not None:
            vals[k], pcts[k] = v, p
        i += 1
    return vals, pcts


def _es_marca(w):
    """Row markers the export leaves between the two halves («#», «2», «7»)."""
    return re.fullmatch(r'[#\d]', w['text']) is not None


def _lee_mitad(filas, x_min, x_max):
    """Sections in one half of the page: [{id, titulo, cols, filas: [{n, v, p}]}]."""
    secciones, actual = [], None
    for f in filas:
        ws = [w for w in f['w'] if x_min <= w['x0'] < x_max]
        if not ws:
            continue
        frases = _frases(ws)
        cols = [fr for fr in frases if _COLUMNA.match(_norm(fr[0]))]
        tit = [fr for fr in frases if not _COLUMNA.match(_norm(fr[0])) and not re.fullmatch(r'[#\d]', fr[0])]
        sid = _seccion(tit[0][0]) if tit else None
        if sid and len(cols) >= 2 and tit[0][1] < cols[0][1]:
            actual = {'id': sid, 'titulo': tit[0][0], 'cols': [c[0] for c in cols],
                      'centros': [(c[1] + c[2]) / 2.0 for c in cols], 'filas': []}
            secciones.append(actual)
            continue
        if actual is None:
            continue
        borde = actual['centros'][0] - 22
        etiqueta = ' '.join(w['text'] for w in ws if w['x1'] < borde and not _es_marca(w)).strip()
        datos = [w for w in ws if w['x1'] >= borde]
        if not etiqueta and not datos:
            continue
        if _norm(etiqueta).startswith(('comentarios', 'socios de tierra')):
            actual = None
            continue
        if _norm(etiqueta).startswith('* mayor') or 'refleja' in _norm(etiqueta):     # the Z-score legend
            continue
        if etiqueta.startswith('*'):
            etiqueta = etiqueta.lstrip('* ').strip()
            if _norm(etiqueta) == 'resultado':
                etiqueta = 'Resultado Z-Score'
        v, p = _celdas(datos, actual['centros'])
        if not etiqueta and all(x is None for x in v):
            continue
        actual['filas'].append({'n': etiqueta, 'v': [None if x is None else round(x, 6) for x in v], 'p': p})
    for s in secciones:
        s.pop('centros', None)
    return secciones


def _socios(filas, ancho):
    """«Socios de Tierra»: one column per land partner, across the whole page."""
    i0 = next((i for i, f in enumerate(filas) if _norm(' '.join(w['text'] for w in f['w'])).startswith('socios de tierra')), None)
    if i0 is None:
        return None
    frases = _frases(filas[i0]['w'])
    cols = frases[1:]
    centros = [(c[1] + c[2]) / 2.0 for c in cols]
    borde = cols[0][1] - 4 if cols else ancho
    out = []
    for f in filas[i0 + 1:]:
        ws = f['w']
        txt = _norm(' '.join(w['text'] for w in ws))
        if txt.startswith('comentarios') or not ws:
            break
        etiqueta = ' '.join(w['text'] for w in ws if w['x1'] < borde).strip()
        v, p = _celdas([w for w in ws if w['x1'] >= borde], centros)
        if etiqueta:
            out.append({'n': etiqueta, 'v': v, 'p': p})
    return {'cols': [c[0] for c in cols], 'filas': out}


def _comentarios(filas):
    i0 = next((i for i, f in enumerate(filas) if _norm(' '.join(w['text'] for w in f['w'])).startswith('comentarios')), None)
    if i0 is None:
        return []
    items = []
    for f in filas[i0 + 1:]:
        t = ' '.join(w['text'] for w in f['w']).strip()
        if not t:
            continue
        m = re.match(r'^(\d+)\.-\s*(.*)$', t)
        if m:
            items.append(m.group(2).strip())
        elif items:
            items[-1] += ' ' + t
    return items


def _fila(secs, sid, *nombres):
    s = next((x for x in secs if x['id'] == sid), None)
    if not s:
        return None, None
    for n in nombres:
        f = next((r for r in s['filas'] if _norm(r['n']).startswith(_norm(n))), None)
        if f:
            return s, f
    return s, None


def _valida(secs):
    """The report's own arithmetic, on every column the rows share. (ok, lines)"""
    ok, lines = True, []

    def cmp(etq, a, b, tol):
        nonlocal ok
        if a is None or b is None:
            return
        d = abs(a - b)
        ok &= d <= tol
        lines.append('%-40s dif %.2f' % (etq, d))

    s, ub = _fila(secs, 'resultados', 'utilidad bruta')
    _, ga = _fila(secs, 'resultados', 'gastos administracion')
    _, gv = _fila(secs, 'resultados', 'gastos de venta')
    _, eb = _fila(secs, 'resultados', 'ebit')
    _, gf = _fila(secs, 'resultados', 'gastos financieros')
    _, ua = _fila(secs, 'resultados', 'uair')
    if s and ub and ga and gv and eb:
        for j, c in enumerate(s['cols']):
            if '%' in c:
                continue
            if None not in (ub['v'][j], ga['v'][j], gv['v'][j], eb['v'][j]):
                cmp('EBIT = UB - admin - venta · %s' % c, ub['v'][j] - ga['v'][j] - gv['v'][j], eb['v'][j], 0.25)
            if gf and ua and None not in (eb['v'][j], gf['v'][j], ua['v'][j]):
                # financial cost comes positive (Oct-2024 on) or negative (before)
                cmp('UAIR = EBIT - financieros · %s' % c, eb['v'][j] - abs(gf['v'][j]), ua['v'][j], 0.25)
    s, tot = _fila(secs, 'deuda', 'total')
    if s and tot:
        bancos = [f for f in s['filas'] if f is not tot and not _norm(f['n']).startswith('total')]
        j = 0
        if bancos and tot['v'][j] is not None:
            cmp('Deuda: bancos = Total', sum(f['v'][j] or 0 for f in bancos), tot['v'][j], 0.3)
    s, si = _fila(secs, 'flujo', 'saldo inicial')
    _, fd = _fila(secs, 'flujo', 'flujo disponible')
    _, eb2 = _fila(secs, 'flujo', 'ebitda')
    if s and si and fd and eb2:
        j = 0
        if None not in (si['v'][j], eb2['v'][j], fd['v'][j]):
            cmp('Flujo: saldo inicial + EBITDA = disponible', si['v'][j] + eb2['v'][j], fd['v'][j], 0.25)
    return bool(ok), lines


def fecha_de_texto(texto):
    m = re.search(r'cierre al (\d{1,2}) de ([a-z]+) de (\d{4})', _norm(texto))
    if not m or m.group(2) not in MESES:
        return None
    return '%s-%02d-%02d' % (m.group(3), MESES[m.group(2)], int(m.group(1)))


def parse_aaa(file_bytes: bytes, filename: str, as_of: str):
    """-> (data, ok, lines). Raises ValueError when page 1 isn't an AAA scorecard."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        if not pdf.pages:
            raise ValueError('%s has no pages' % filename)
        p = pdf.pages[0]
        words = p.extract_words(x_tolerance=1.5, keep_blank_chars=False)
        ancho = float(p.width)
        texto = p.extract_text() or ''
    filas = _filas(words)
    corte = fecha_de_texto(texto) or as_of
    # the two halves end where the full-width land-partner table (or the commentary) begins
    fin = next((f['top'] for f in filas if _norm(' '.join(w['text'] for w in f['w'])).startswith(('socios de tierra', 'comentarios'))), None)
    arriba = [f for f in filas if fin is None or f['top'] < fin - 1]
    # the right half starts where its first title («Cash Flow…») does
    cash = next((w for w in words if w['text'] == 'Cash'), None)
    mitad = (cash['x0'] - 1) if cash else ancho * 0.55
    # the export leaves row markers («#», «2», «7») just left of the right half: keep them out of both
    secs = _lee_mitad(arriba, 0, mitad - 5.5) + _lee_mitad(arriba, mitad, ancho + 1)
    ids = {s['id'] for s in secs}
    if not {'variables', 'resultados'} <= ids:
        raise ValueError("%s: page 1 doesn't read as an AAA scorecard (no units or P&L table found)" % filename)
    ok, lines = _valida(secs)
    data = {'archivo': filename, 'asOf': corte, 'secciones': secs, 'socios': _socios(filas, ancho),
            'comentarios': _comentarios(filas)}
    return data, ok, lines


# ── history: the few series the Dashboard charts, one value per monthly cut ──
_COLKEY = [(r'^target$', 'target'), (r'^real mes$', 'rm'), (r'^plan mes$', 'pm'), (r'^real acum$', 'ra'),
           (r'^plan \d{4} acum$', 'pa'), (r'^% var$', 'var'), (r'^acum$', 'acum'), (r'^real$', 'real')]
SERIES = [  # key, section, row (prefix), column key
    ('firmas_rm', 'variables', 'firmadas', 'rm'), ('firmas_pm', 'variables', 'firmadas', 'pm'),
    ('firmas_ra', 'variables', 'firmadas', 'ra'), ('firmas_pa', 'variables', 'firmadas', 'pa'),
    ('apartadas_rm', 'variables', 'apartadas', 'rm'), ('gestoria_rm', 'variables', 'viviendas en gestoria', 'rm'),
    ('inv_term_rm', 'variables', 'inv. viv. terminadas', 'rm'),
    ('ventas_rm', 'resultados', 'ventas netas', 'rm'), ('ventas_pm', 'resultados', 'ventas netas', 'pm'),
    ('uair_rm', 'resultados', 'uair', 'rm'), ('uair_pm', 'resultados', 'uair', 'pm'),
    ('mb_ra', 'margenes', 'margen bruto', 'ra'), ('roic_ra', 'rendimientos', 'roic', 'ra'),
    ('deuda_total', 'deuda', 'total', 'real'), ('nd_ebitda', 'razones_deuda', 'net debt', 'real'),
    ('cobertura', 'razones_deuda', 'cobertura', 'real'), ('caja', 'flujo', 'saldo ranman vivienda', 'rm'),
]


def _colkey(c):
    n = _norm(c)
    return next((k for rx, k in _COLKEY if re.match(rx, n)), None)


def serie_de(data, seccion, fila, col):
    s = next((x for x in data.get('secciones') or [] if x['id'] == seccion), None)
    if not s:
        return None
    f = next((r for r in s['filas'] if _norm(r['n']).startswith(_norm(fila))), None)
    if not f:
        return None
    keys = [_colkey(c) for c in s['cols']]
    return f['v'][keys.index(col)] if col in keys else None


def historia(cortes):
    """[(as_of, data)] oldest first -> {meses: ['2024-01', …], series: {key: [value | None]}}."""
    out = {'meses': [], 'series': {k: [] for k, *_ in SERIES}}
    for as_of, data in cortes:
        out['meses'].append(str(as_of)[:7])
        for k, sec, fila, col in SERIES:
            out['series'][k].append(serie_de(data, sec, fila, col))
    return out
