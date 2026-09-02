"""
kmz_parser.py — KMZ / KML → compact GeoJSON for one highway on the Meta map.

Handles the two flavours we receive:
  • Hand-drawn Google Earth files: one LineString placemark per road (the
    "Libramiento Meta.kmz" carries four roads in one file, picked by name).
  • CAD exports ("EJE DE TRAZO …"): hundreds of two-point LineString
    placemarks. These are chained end-to-end into continuous polylines.

Output is simplified with Douglas–Peucker (default ~12 m) so a whole route is
a few KB in the page payload, and the geodesic length (haversine) is reported
so it can seed the highway's length when no official figure exists.
"""
from __future__ import annotations

import io
import math
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

_EARTH_KM = 6371.0088


def _norm(s) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _hav(a, b) -> float:
    la1, lo1 = math.radians(a[1]), math.radians(a[0])
    la2, lo2 = math.radians(b[1]), math.radians(b[0])
    d = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * _EARTH_KM * math.asin(math.sqrt(d))


def length_km(lines) -> float:
    return sum(_hav(l[i], l[i + 1]) for l in lines for i in range(len(l) - 1))


# ── KML reading ──────────────────────────────────────────────────────────────
def _kml_bytes(file_bytes: bytes) -> bytes:
    try:
        z = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile:
        return file_bytes                      # plain .kml
    kmls = [n for n in z.namelist() if n.lower().endswith(".kml")]
    if not kmls:
        raise ValueError("No .kml document inside the KMZ.")
    kmls.sort(key=lambda n: (n.lower() != "doc.kml", n))
    return z.read(kmls[0])


def _coords(el):
    c = el.find(".//{*}coordinates")
    if c is None or not c.text:
        return []
    out = []
    for tok in c.text.split():
        parts = tok.split(",")
        if len(parts) >= 2:
            try:
                out.append((round(float(parts[0]), 6), round(float(parts[1]), 6)))
            except ValueError:
                pass
    return out


def parse_kmz(file_bytes: bytes) -> list:
    """[{name, folder, lines:[[(lon,lat),…]], points:[(lon,lat)]}] per Placemark."""
    try:
        root = ET.fromstring(_kml_bytes(file_bytes))
    except ET.ParseError as e:
        raise ValueError("Not a readable KML/KMZ file (%s)." % e)
    feats = []

    def walk(el, folder):
        tag = el.tag.split("}")[-1]
        if tag == "Folder":
            nm = el.find("{*}name")
            if nm is not None and nm.text:
                folder = nm.text.strip()
        if tag == "Placemark":
            nm = el.find("{*}name")
            name = (nm.text or "").strip() if nm is not None else ""
            lines, points = [], []
            for g in el.iter():
                gt = g.tag.split("}")[-1]
                if gt == "Point":
                    cs = _coords(g)
                    if cs:
                        points.append(cs[0])
                elif gt in ("LineString", "LinearRing"):
                    cs = _coords(g)
                    if len(cs) >= 2:
                        lines.append(cs)
            if lines or points:
                feats.append(dict(name=name, folder=folder, lines=lines, points=points))
            return
        for ch in el:
            walk(ch, folder)

    walk(root, "")
    if not feats:
        raise ValueError("The file has no LineString or Point placemarks.")
    return feats


# ── geometry cleanup ─────────────────────────────────────────────────────────
def chain_segments(lines, tol=1e-6) -> list:
    """Join segments that share endpoints (CAD exports) into polylines."""
    segs = [list(l) for l in lines if len(l) >= 2]
    chains = []
    while segs:
        cur = segs.pop(0)
        grew = True
        while grew:
            grew = False
            for i, s in enumerate(segs):
                if abs(s[0][0] - cur[-1][0]) < tol and abs(s[0][1] - cur[-1][1]) < tol:
                    cur += s[1:]; segs.pop(i); grew = True; break
                if abs(s[-1][0] - cur[0][0]) < tol and abs(s[-1][1] - cur[0][1]) < tol:
                    cur = s[:-1] + cur; segs.pop(i); grew = True; break
                if abs(s[-1][0] - cur[-1][0]) < tol and abs(s[-1][1] - cur[-1][1]) < tol:
                    cur += list(reversed(s))[1:]; segs.pop(i); grew = True; break
                if abs(s[0][0] - cur[0][0]) < tol and abs(s[0][1] - cur[0][1]) < tol:
                    cur = list(reversed(s))[:-1] + cur; segs.pop(i); grew = True; break
        chains.append(cur)
    return chains


def _perp(p, a, b) -> float:
    if a == b:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    dx, dy = b[0] - a[0], b[1] - a[1]
    t = max(0.0, min(1.0, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / (dx * dx + dy * dy)))
    return math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy))


def simplify(points, eps_m=12.0) -> list:
    """Douglas–Peucker on lon/lat; eps in metres (lon scaled by cos(lat))."""
    if len(points) < 3:
        return list(points)
    lat0 = math.radians(sum(p[1] for p in points) / len(points))
    kx, ky = 111320.0 * math.cos(lat0), 110540.0
    pts = [(p[0] * kx, p[1] * ky) for p in points]
    keep = [False] * len(pts)
    keep[0] = keep[-1] = True
    stack = [(0, len(pts) - 1)]
    while stack:
        i, j = stack.pop()
        if j <= i + 1:
            continue
        dmax, idx = 0.0, -1
        for k in range(i + 1, j):
            d = _perp(pts[k], pts[i], pts[j])
            if d > dmax:
                dmax, idx = d, k
        if dmax > eps_m and idx > 0:
            keep[idx] = True
            stack.append((i, idx)); stack.append((idx, j))
    return [p for p, k in zip(points, keep) if k]


# ── highway build ────────────────────────────────────────────────────────────
def pick(feats, wanted) -> list:
    """Placemarks whose name contains `wanted` (accent/case-insensitive);
    falls back to all placemarks when nothing matches or wanted is empty."""
    w = _norm(wanted)
    if not w:
        return feats
    hit = [f for f in feats if w in _norm(f["name"]) or w in _norm(f["folder"])]
    return hit or feats


def build_highway_geo(file_bytes: bytes, wanted=None, eps_m: float = 12.0) -> dict:
    feats = pick(parse_kmz(file_bytes), wanted)
    lines = [l for f in feats for l in f["lines"]]
    points = [p for f in feats for p in f["points"]]
    if not lines and not points:
        raise ValueError("No route geometry found for '%s'." % (wanted or "the file"))
    chains = chain_segments(lines)
    raw_pts = sum(len(c) for c in chains)
    km = length_km(chains)
    simp = [simplify(c, eps_m) for c in chains]
    simp = [c for c in simp if len(c) >= 2]
    features = []
    if simp:
        features.append({"type": "Feature", "properties": {"kind": "route"},
                         "geometry": {"type": "MultiLineString", "coordinates": [[[x, y] for x, y in c] for c in simp]}})
    for p in points:
        features.append({"type": "Feature", "properties": {"kind": "point"},
                         "geometry": {"type": "Point", "coordinates": [p[0], p[1]]}})
    return dict(
        geojson={"type": "FeatureCollection", "features": features},
        length_km=round(km, 2), n_points=sum(len(c) for c in simp), raw_points=raw_pts,
        n_lines=len(simp), n_markers=len(points),
        names=sorted(set(f["name"] for f in feats if f["name"]))[:8],
    )
