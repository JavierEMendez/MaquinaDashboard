"""Push RANMAN's monthly package from OneDrive to the dashboard.

Walks the last few month folders of «Archivos Ranman» (synced by OneDrive on
this computer), picks out the deliverables by file name — the same rules the
dashboard uses (ranman_package.clasifica) — and sends the new or changed ones,
one per request, to /api/ranman/package. What was sent is remembered in a small
state file, so a daily run only sends what Ranman added since.

Files go oldest month first, and within a month in order of their cut date: the
dashboard keeps the year's oldest Flujo y PLP as its opening plan, so the
January workbook has to land before the later ones. A file the dashboard
rejected (its own totals don't add up) is remembered too and isn't sent again
until it changes — Ranman re-saving it is what can fix it. After a dashboard
update reads something new, --retry received,error sends those files again.

Needs, on the computer that runs it:
    RANMAN_SYNC_URL    https://<the dashboard>/api/ranman/package
    RANMAN_SYNC_TOKEN  the same value as the RANMAN_SYNC_TOKEN set on Railway
and Python with `requests` and `openpyxl` (pip install -r requirements.txt).

    python tools/ranman_sync.py --dry-run         # list what it would send
    python tools/ranman_sync.py                   # send it
    python tools/ranman_sync.py --months 12       # look further back
    python tools/ranman_sync.py --force           # resend even if unchanged
    python tools/ranman_sync.py --retry error     # resend files rejected last time, even if unchanged
    python tools/ranman_sync.py --months 34 --only aaa,sabana   # load those reports' history
    python tools/ranman_sync.py --months 14 --only bp           # a year of business plans (~1 GB)
    python tools/ranman_sync.py --root "D:\\OneDrive\\Archivos Ranman"

Exit code 1 when a file was rejected or the server couldn't be reached, so a
scheduler can flag the run.
"""
import argparse
import json
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ranman_package  # noqa: E402  (the dashboard's own name rules)

MAX_BYTES = 32 * 1024 * 1024         # the dashboard's per-file cap on /api/ranman/package
MESES = {'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6, 'july': 7,
         'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12}
STATE = Path(os.environ.get('LOCALAPPDATA') or Path.home()) / 'ranman_sync' / 'state.json'


def raiz_por_defecto():
    """%USERPROFILE%/OneDrive - Grupo Valoran/Archivos Ranman (or the macOS equivalent)."""
    casa = Path.home()
    for c in (casa / 'OneDrive - Grupo Valoran' / 'Archivos Ranman',
              casa / 'Library' / 'CloudStorage' / 'OneDrive-GrupoValoran' / 'Archivos Ranman'):
        if c.is_dir():
            return c
    return None


def meses(raiz: Path, n: int):
    """The last n month folders under <raiz>/<year>/<m>_<Month> that hold any
    file, newest first (Ranman creates the year's twelve folders up front)."""
    out = []
    for y in raiz.iterdir():
        if not (y.is_dir() and re.fullmatch(r'20\d{2}', y.name)):
            continue
        for m in y.iterdir():
            k = re.match(r'(\d{1,2})_', m.name)
            if m.is_dir() and k and 1 <= int(k.group(1)) <= 12:
                out.append((int(y.name), int(k.group(1)), m))
    out.sort(reverse=True)
    llenos = []
    for y, m, carpeta in out:
        if next(archivos(carpeta), None) is not None:
            llenos.append((y, m, carpeta))
            if len(llenos) == n:
                break
    return llenos


def archivos(carpeta: Path):
    """Every file under a month folder, never entering backup / scenario /
    paperwork sub-folders: deep inside the BP backups the paths pass Windows'
    260-character limit, and nothing there is a deliverable anyway."""
    def poda(d):
        return any(o in ranman_package._norm(d) for o in ranman_package.OMITE_CARPETA)
    for base, dirs, files in os.walk(carpeta, topdown=True, onerror=lambda e: None):
        dirs[:] = sorted(d for d in dirs if not poda(d))
        for f in sorted(files):
            yield Path(base) / f


def candidatos(raiz: Path, n: int):
    """[(path, relative path, kind, periodo)] for every deliverable worth sending."""
    for y, m, carpeta in meses(raiz, n):
        periodo = '%d-%02d' % (y, m)
        for p in archivos(carpeta):
            rel = p.relative_to(raiz).as_posix()
            if ranman_package.omitir_ruta(rel):
                continue
            kind, _ = ranman_package.clasifica(p.name)
            if kind and kind in ranman_package.ENVIABLES:
                yield p, rel, kind, periodo


def huella(p: Path):
    st = p.stat()
    return [st.st_size, int(st.st_mtime)]


def recuerdo(v):
    """A state entry -> (fingerprint, last status). Older state files stored only
    the fingerprint of files that went through."""
    if isinstance(v, dict):
        return v.get('h'), v.get('s')
    return v, 'imported'


def orden(item):
    """Oldest first: by folder month, then by the cut date in the name, then path."""
    p, rel, kind, periodo = item
    return (periodo, ranman_package.clasifica(p.name)[1] or '', rel)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    ap.add_argument('--root', help='the «Archivos Ranman» folder (default: the OneDrive one)')
    ap.add_argument('--months', type=int, default=3, help='how many month folders back to look (default 3)')
    ap.add_argument('--url', default=os.environ.get('RANMAN_SYNC_URL'))
    ap.add_argument('--token', default=os.environ.get('RANMAN_SYNC_TOKEN'))
    ap.add_argument('--dry-run', action='store_true', help='list what would be sent, send nothing')
    ap.add_argument('--force', action='store_true', help='send even what was already sent unchanged')
    ap.add_argument('--state', default=str(STATE), help='where to remember what was sent (default %(default)s)')
    ap.add_argument('--only', help='send only these deliverables, comma-separated (e.g. aaa,sabana) — '
                                   'for loading one report\'s history without resending the rest')
    ap.add_argument('--retry', help='resend unchanged files whose last answer was one of these statuses, '
                                    'comma-separated: error, received (e.g. after the dashboard learns to read one)')
    a = ap.parse_args()
    state = Path(a.state)
    solo = {k.strip() for k in a.only.split(',')} if a.only else None
    if solo and not solo <= set(ranman_package.ENTREGABLE):
        sys.exit('--only: unknown deliverable(s) %s; use any of %s'
                 % (', '.join(sorted(solo - set(ranman_package.ENTREGABLE))), ', '.join(ranman_package.ENTREGABLE)))
    reintenta = {s.strip() for s in a.retry.split(',')} if a.retry else set()
    if reintenta - {'error', 'received', 'imported'}:
        sys.exit('--retry: use error, received or imported')

    raiz = Path(a.root) if a.root else raiz_por_defecto()
    if not raiz or not raiz.is_dir():
        sys.exit('Archivos Ranman folder not found — pass it with --root.')
    if not a.dry_run and not (a.url and a.token):
        sys.exit('Set RANMAN_SYNC_URL and RANMAN_SYNC_TOKEN (or pass --url / --token).')
    try:
        estado = json.loads(state.read_text(encoding='utf-8')) if state.exists() else {}
    except ValueError:
        estado = {}

    pendientes, rechazados = [], 0
    for p, rel, kind, periodo in candidatos(raiz, a.months):
        if solo and kind not in solo:
            continue
        h, ultimo = recuerdo(estado.get(rel))
        if not a.force and h == huella(p) and ultimo not in reintenta:
            rechazados += ultimo == 'error'
            continue
        pendientes.append((p, rel, kind, periodo))
    pendientes.sort(key=orden)
    print('%s · %d month folder(s) · %d file(s) to send' % (raiz, a.months, len(pendientes)))
    if rechazados:
        print('  %d unchanged file(s) the dashboard rejected last time are left out (--retry error sends them)' % rechazados)
    if a.dry_run:
        for p, rel, kind, periodo in pendientes:
            print('  %-15s %s  %s' % (kind, periodo, rel))
        return 0

    import requests
    fallos = 0
    for p, rel, kind, periodo in pendientes:
        if p.stat().st_size > MAX_BYTES:
            print('  SKIP   %s — over 32 MB' % rel)
            continue
        try:
            with p.open('rb') as fh:
                r = requests.post(a.url, headers={'Authorization': 'Bearer ' + a.token},
                                  files={'files': (p.name, fh)}, data={'periodo': periodo, 'ruta': rel},
                                  timeout=300)
            res = (r.json().get('results') or [{}])[0] if r.headers.get('content-type', '').startswith('application/json') else {}
        except (OSError, ValueError, requests.RequestException) as e:
            print('  FAIL   %s — %s' % (rel, e))
            fallos += 1
            continue
        if r.status_code != 200 or not res:
            print('  FAIL   %s — server answered %s %s' % (rel, r.status_code, r.text[:200]))
            fallos += 1
            if r.status_code == 401:
                break                       # wrong token: no point trying the rest
            continue
        st = res.get('status') or '?'
        print('  %-8s %s — %s' % (st.upper(), rel, res.get('message', '')))
        if st == 'error':
            fallos += 1                     # rejected by its own totals: sent again once the file changes
        # what the dashboard answered, so an unchanged file isn't sent again tomorrow
        estado[rel] = {'h': huella(p), 's': st}
        # saved as it goes: an interrupted run keeps what it already sent
        state.parent.mkdir(parents=True, exist_ok=True)
        state.write_text(json.dumps(estado, indent=1, ensure_ascii=False), encoding='utf-8')
    print('done: %d sent, %d problem(s)' % (len(pendientes) - fallos, fallos))
    return 1 if fallos else 0


if __name__ == '__main__':
    sys.exit(main())
