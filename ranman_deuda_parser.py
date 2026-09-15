"""
ranman_deuda_parser.py — RANMAN "Cuadro de Riesgos" (debt with cost) parser.

A faithful port of Valoran's `actualizar_cuadro_riesgos.py` (Sep-2026, kept in
OneDrive under Archivos Ranman/_Tablero de riesgo). RANMAN publishes a debt
snapshot ("corte") every week or two as `Cuadro de riesgos Ranman <dd mmm aa>.xlsm`;
this module reads the `deuda con costo` sheet of one such file and returns the
credit-level detail plus the per-corte summary record, using the SAME
classification rules so Maquina's series is comparable with the tablero's:

  • Puentes      — debt secured by the project's own land (BBVA / BANREGIO /
                   BanCrea bridge loans, Azhala land credit, the Cúspide
                   infrastructure share of the BanBajío line). Repaid from the
                   homes they finance.
  • No-Puentes   — corporate lines (CCC 55 mdp, BBVA simples 99.8 / 100 / 28,
                   BanCrea línea 48, the corporate share of BanBajío). No own
                   source of repayment — the block that matters.
    – Amortizable (the three simples: repaid principal cannot be redrawn)
    – Revolvente  (everything else: can be redrawn — a drop is not retired debt)

Amounts are in THOUSANDS of MXN, as in the workbook (÷1000 → mdp).

Blocks per bank are located by the text `Fecha de Firma` in column B — never by
row number. A credit is any row with a name and a numeric saldo deudor (some
rows lack a signing date). Each file is reconciled against the Total it declares
on the same sheet; a mismatch is reported to the caller.

Also parses the tablero's `serie.json` / `corte_actual.json` so the 101-corte
history can be imported in one go.
"""
from __future__ import annotations

import datetime as dt
import io
import json
import re
import unicodedata
from collections import defaultdict

import openpyxl

MESES = {'ene': 1, 'feb': 2, 'mzo': 3, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
         'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12}

# Créditos simples: amortised principal can NOT be redrawn.
AMORTIZABLES = {'Simple 99.8 mdp', 'Simple 100 mdp', 'Simple 28 mdp'}

BANCO = {'Hipotecaria Nacional': 'BBVA', 'BBVA CCC 55MDP': 'BBVA',
         'BBVA SIMPLE 100mdp (99.780 MDP)': 'BBVA', 'BBVA SIMPLE 100 MDP': 'BBVA',
         'BBVA SIMPLE 28MDP': 'BBVA', 'BanBajio': 'BanBajío', 'BANREGIO': 'BANREGIO',
         'BanCrea': 'BanCrea', 'ION': 'ION', 'SANTANDER': 'SANTANDER'}

# Authorised size of the revolving No-Puentes lines (mdp), read off the line
# names — used for "capacidad redisponible" (authorised minus drawn).
REVOLVING_AUTHORIZED_MDP = {'CCC 55 mdp': 55.0, 'Línea 48 mdp': 48.0}


def _ascii(s: str) -> str:
    return unicodedata.normalize('NFKD', s.lower()).encode('ascii', 'ignore').decode()


def fecha_de_nombre(nombre: str):
    """'Cuadro de riesgos Ranman 31 ago 26.xlsm' -> '2026-08-31' (ISO) or None."""
    s = nombre.lower().replace('.xlsm', '').replace('.xlsx', '').replace('_modificado', '').replace(' vf', '')
    m = re.search(r'(\d{1,2})\s*([a-z]{3})\s*(\d{2})\s*$', s)
    if not m:
        return None
    mes = MESES.get(m.group(2))
    if not mes:
        return None
    try:
        return dt.date(2000 + int(m.group(3)), mes, int(m.group(1))).isoformat()
    except ValueError:
        return None


def desarrollo(n: str) -> str:
    s = _ascii(n)
    for k, v in [('puerta natura', 'Puerta Natura'), ('puerta de piedra', 'Corporativo'),
                 ('cantabria', 'Cantabria'), ('trive', 'Trive'), ('cartuja', 'Cartuja'),
                 ('cuspide', 'Cúspide'), ('alella', 'Alella'), ('azhala', 'Azhala'),
                 ('carmel', 'Carmel'), ('miranda', 'Celaya'), ('celaya', 'Celaya'),
                 ('gran natura', 'Corporativo'), ('mompani', 'Corporativo'),
                 ('aurea', 'Corporativo'), ('veredas', 'Corporativo'),
                 ('punta norte', 'Corporativo'), ('ximonco', 'Corporativo'),
                 ('inmuebles', 'Corporativo'), ('10k25', 'Corporativo')]:
        if k in s:
            return v
    return 'Corporativo'


def clasifica(grupo: str, nombre: str, desa: str):
    """Same rule as lámina 21: Puentes = debt secured by the project's own land;
    No-Puentes = corporate lines. Returns (línea, categoría)."""
    g = grupo.upper()
    s = _ascii(nombre)
    if 'tierra' in s:
        return 'Crédito tierra', 'Puentes'
    if grupo == 'Hipotecaria Nacional':
        return 'Crédito puente', 'Puentes'
    if re.match(r'^(pre)?p[ru]{1,2}ente', s) or 'puente' in s.split('.-')[0]:
        return 'Crédito puente', 'Puentes'
    if 'CCC' in g:
        return 'CCC 55 mdp', 'No-Puentes'
    if '99.780' in g or '99.78' in g:
        return 'Simple 99.8 mdp', 'No-Puentes'
    if g == 'BBVA SIMPLE 100 MDP':
        return 'Simple 100 mdp', 'No-Puentes'
    if '28MDP' in g:
        return 'Simple 28 mdp', 'No-Puentes'
    if g == 'BANBAJIO':
        # Shared line: the Cúspide infrastructure share is project debt; the rest corporate.
        return ('Infraestructura', 'Puentes') if desa == 'Cúspide' else ('Infraestructura', 'No-Puentes')
    if g == 'BANCREA' and 'linea48' in s:
        return 'Línea 48 mdp', 'No-Puentes'
    if g == 'BANREGIO' and 'linea' in s:
        return 'Línea BANREGIO', 'No-Puentes'
    return grupo, 'No-Puentes'


def parse_corte(file_bytes: bytes, filename: str = ''):
    """One corte workbook -> (fecha_iso, creditos, declarado_total).

    Raises ValueError when the file isn't a Cuadro de Riesgos."""
    fecha = fecha_de_nombre(filename or '')
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    try:
        if 'deuda con costo' not in wb.sheetnames:
            raise ValueError("'%s' has no 'deuda con costo' sheet — is it a Cuadro de Riesgos?" % filename)
        filas = list(wb['deuda con costo'].iter_rows(min_row=1, max_row=200, max_col=17, values_only=True))
    finally:
        wb.close()

    declarado = None
    for r in filas:
        if r[0] and str(r[0]).strip().lower() == 'total' and isinstance(r[1], (int, float)):
            declarado = float(r[1])
            break

    creds, grupo = [], None
    for r in filas:
        if r[1] == 'Fecha de Firma' and r[0]:
            grupo = str(r[0]).strip()
            continue
        if not grupo or not r[0]:
            continue
        nombre = str(r[0]).strip()
        if nombre.upper().startswith('TOTAL'):
            grupo = None
            continue
        if not isinstance(r[7], (int, float)):        # no saldo deudor -> not a credit
            continue
        desa = desarrollo(nombre)
        linea, cat = clasifica(grupo, nombre, desa)
        creds.append(dict(
            banco=BANCO.get(grupo, grupo), credito=nombre, linea=linea, cat=cat,
            desarrollo=desa, saldo=float(r[7]),
            vence=r[2].date().isoformat() if isinstance(r[2], dt.datetime) else (r[2].isoformat() if isinstance(r[2], dt.date) else None),
            porEjercer=float(r[8]) if isinstance(r[8], (int, float)) else 0.0,
            sobretasa=float(r[14]) if isinstance(r[14], (int, float)) else None,
            avance=float(r[13]) if isinstance(r[13], (int, float)) else None))
    if not creds:
        raise ValueError("'%s': no credit rows found under any 'Fecha de Firma' block." % filename)
    return fecha, creds, declarado


def _agrupa(creds, campo):
    d = defaultdict(float)
    for c in creds:
        d[c[campo]] += c['saldo']
    return {k: round(v, 2) for k, v in d.items() if abs(v) > 0.005}


def serie_record(fecha: str, archivo: str, creds: list) -> dict:
    """The per-corte summary — identical shape to the tablero's serie.json."""
    total = sum(c['saldo'] for c in creds)
    npc = [c for c in creds if c['cat'] == 'No-Puentes']
    corte12 = dt.date.fromisoformat(fecha) + dt.timedelta(days=365)
    consob = [c for c in creds if c['sobretasa'] and c['saldo']]
    base_sob = sum(c['saldo'] for c in consob)
    return dict(
        fecha=fecha, archivo=archivo, total=round(total, 2), n=len(creds),
        cat=_agrupa(creds, 'cat'), banco=_agrupa(creds, 'banco'),
        desarrollo=_agrupa(creds, 'desarrollo'), linea=_agrupa(creds, 'linea'),
        np={'Amortizable': round(sum(c['saldo'] for c in npc if c['linea'] in AMORTIZABLES), 2),
            'Revolvente': round(sum(c['saldo'] for c in npc if c['linea'] not in AMORTIZABLES), 2)},
        porEjercer=round(sum(c['porEjercer'] for c in creds), 2),
        sobretasaPond=round(sum(c['sobretasa'] * c['saldo'] for c in consob) / base_sob, 6) if base_sob else None,
        vence12=round(sum(c['saldo'] for c in creds
                          if c['vence'] and dt.date.fromisoformat(c['vence']) <= corte12), 2))


def reconcile(creds: list, declarado) -> str | None:
    """None when the parsed sum matches the sheet's own Total (±1 thousand)."""
    total = sum(c['saldo'] for c in creds)
    if declarado is not None and abs(total - declarado) > 1:
        return 'sum %.1f vs declared Total %.1f mdp' % (total / 1000, declarado / 1000)
    return None


# ── tablero JSON import ──────────────────────────────────────────────────────
def parse_serie_json(file_bytes: bytes) -> list:
    data = json.loads(file_bytes.decode('utf-8'))
    if not isinstance(data, list) or not data or not all(isinstance(r, dict) and 'fecha' in r and 'total' in r for r in data):
        raise ValueError('serie.json should be a list of corte records with fecha and total.')
    out = []
    for r in data:
        dt.date.fromisoformat(r['fecha'])          # validates
        out.append(r)
    return out


def parse_corte_actual_json(file_bytes: bytes) -> dict:
    data = json.loads(file_bytes.decode('utf-8'))
    if not isinstance(data, dict) or 'asOf' not in data or not isinstance(data.get('creditos'), list):
        raise ValueError('corte_actual.json should carry asOf and a creditos list.')
    dt.date.fromisoformat(data['asOf'])
    return data
