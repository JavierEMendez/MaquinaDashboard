"""Maquina — Portfolio Management Platform (Flask).

Single-file app following the Ember dashboard's conventions: session auth
over a `users` table, a Railway Postgres via psycopg2, schema + fake-data
seeding on first request. Pages read from the unified model in db.py and
hand fully-shaped data to Jinja templates (charts are rendered client-side
with Chart.js from embedded JSON).
"""
import os
import io
import json
import functools

import requests
import psycopg2
from PIL import Image, ImageOps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, abort, Response)
from werkzeug.security import check_password_hash, generate_password_hash

import db
import seed_data

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "maquina-dev-secret-change-me")
# Cap uploads; processed down to <=256px thumbnails before storage.
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024


@app.errorhandler(413)
def _too_large(e):
    flash("That file is too large (max 8 MB).", "error")
    return redirect(request.referrer or url_for("settings"))


def process_image(file_storage, mode="cover", size=256):
    """Normalize an upload to a small thumbnail.

    mode='cover'   -> square center-crop (profile pictures) as JPEG
    mode='contain' -> fit within size, keep aspect + transparency (logos) as PNG
    Raises on anything Pillow can't open.
    """
    img = Image.open(file_storage.stream)
    img = ImageOps.exif_transpose(img)
    if mode == "cover":
        img = img.convert("RGB")
        img = ImageOps.fit(img, (size, size), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=85, optimize=True)
        return buf.getvalue(), "image/jpeg"
    img = img.convert("RGBA")
    img.thumbnail((size, size), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return buf.getvalue(), "image/png"

BASE_YEAR = seed_data.BASE_YEAR
HIST_START = seed_data.HIST_START
PROJ_END = seed_data.PROJ_END

_booted = False


@app.before_request
def _boot():
    """Idempotently create the schema and seed fake data once per process."""
    global _booted
    if _booted:
        return
    try:
        db.init_schema()
        if not db.is_seeded():
            conn = db.get_db()
            seed_data.seed(conn)
            conn.close()
        _booted = True
    except Exception as e:  # pragma: no cover
        app.logger.error("Boot/init failed: %s", e)


# ─── AUTH ──────────────────────────────────────────────────────────
def login_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return wrapper


@app.route("/health")
def health():
    return "ok", 200


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = db.get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username = %s", (username,))
        user = cur.fetchone()
        cur.close()
        conn.close()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["is_admin"] = user["is_admin"]
            fn = (user.get("first_name") or "").strip()
            ln = (user.get("last_name") or "").strip()
            session["display_name"] = f"{fn} {ln}".strip() or user["username"]
            session["has_avatar"] = user.get("avatar") is not None
            session["avatar_ver"] = len(user["avatar"]) if user.get("avatar") else 0
            nxt = request.args.get("next") or url_for("dashboard")
            return redirect(nxt)
        error = "Invalid username or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ─── FORMATTING FILTERS ────────────────────────────────────────────
def _money(x, symbol="$"):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return "—"
    a = abs(x)
    if a >= 1e9:
        return f"{symbol}{x/1e9:.1f}B"
    if a >= 1e6:
        return f"{symbol}{x/1e6:.1f}M"
    if a >= 1e3:
        return f"{symbol}{x/1e3:.0f}K"
    return f"{symbol}{x:,.0f}"


def _signed_money(x, symbol="$"):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return "—"
    sign = "+" if x >= 0 else "-"
    return f"{sign}{_money(abs(x), symbol)}"


def _num(x, integer=False):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return "—"
    if integer or abs(x - round(x)) < 1e-9:
        return f"{x:,.0f}"
    return f"{x:,.1f}"


app.jinja_env.filters["money"] = _money
app.jinja_env.filters["smoney"] = _signed_money
app.jinja_env.filters["num"] = _num


def _fmt_kpi(value, unit, unit_label):
    if value is None:
        return "—"
    if unit == "currency":
        return _money(value, "$")
    if unit == "percent":
        return f"{value:,.1f}%"
    if unit == "count":
        return f"{value:,.0f}"
    # plain number, maybe with a trailing unit label like MW / GWh / Litros
    n = _num(value)
    return f"{n} {unit_label}" if unit_label and unit_label not in ("%",) else n


app.jinja_env.filters["kpi"] = _fmt_kpi


# ─── SHARED DATA HELPERS ───────────────────────────────────────────
def usd_mxn_rate():
    """MXN per 1 USD. Prefers a stored setting; live Banxico if a token is set."""
    token = os.environ.get("BANXICO_TOKEN", "").strip()
    if token:
        try:
            r = requests.get(
                "https://www.banxico.org.mx/SieAPIRest/service/v1/series/SF43718/datos/oportuno",
                headers={"Bmx-Token": token}, timeout=4,
            )
            v = r.json()["bmx"]["series"][0]["datos"][0]["dato"]
            return float(v.replace(",", ""))
        except Exception:
            pass
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT value FROM app_settings WHERE key='usd_mxn_rate'")
    row = cur.fetchone()
    cur.close()
    conn.close()
    try:
        return float(row["value"]) if row else 17.3328
    except (TypeError, ValueError):
        return 17.3328


def all_companies(include_archived=False):
    conn = db.get_db()
    cur = conn.cursor()
    q = ("SELECT id, slug, name, industry, country, country_code, currency, "
         "value_unit, accent, description, takeover_year, display_order, archived, "
         "(logo IS NOT NULL) AS has_logo, octet_length(logo) AS logo_ver FROM companies")
    if not include_archived:
        q += " WHERE archived = FALSE"
    q += " ORDER BY display_order, name"
    cur.execute(q)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return rows


def company_by_slug(slug):
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, slug, name, industry, country, country_code, currency, value_unit, "
        "accent, description, takeover_year, display_order, archived, "
        "(logo IS NOT NULL) AS has_logo, octet_length(logo) AS logo_ver "
        "FROM companies WHERE slug = %s", (slug,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return row


def company_cashflow_totals():
    """Per-company cumulative + per-year investment/distribution (USD)."""
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT c.id, c.slug, c.name, c.industry, c.country, c.country_code,
               c.currency, c.accent,
               (c.logo IS NOT NULL) AS has_logo, octet_length(c.logo) AS logo_ver,
               pc.year, pc.investment, pc.distribution
        FROM companies c
        LEFT JOIN portfolio_cashflows pc ON pc.company_id = c.id
        WHERE c.archived = FALSE
        ORDER BY c.display_order, pc.year
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    comps = {}
    for r in rows:
        cid = r["id"]
        if cid not in comps:
            comps[cid] = dict(
                id=cid, slug=r["slug"], name=r["name"], industry=r["industry"],
                country=r["country"], country_code=r["country_code"],
                currency=r["currency"], accent=r["accent"],
                has_logo=r["has_logo"], logo_ver=r["logo_ver"],
                investment=0.0, distribution=0.0, series={},
            )
        if r["year"] is not None:
            comps[cid]["investment"] += r["investment"] or 0
            comps[cid]["distribution"] += r["distribution"] or 0
            comps[cid]["series"][r["year"]] = dict(
                inv=r["investment"] or 0, dist=r["distribution"] or 0)
    for c in comps.values():
        c["net"] = c["distribution"] - c["investment"]
        c["roi"] = (c["distribution"] / c["investment"] * 100) if c["investment"] else 0
    return list(comps.values())


def load_company_items(cid):
    """All KPI items for a company with actual/projection series + trend."""
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM unified_items WHERE company_id = %s ORDER BY display_order", (cid,))
    items = cur.fetchall()
    cur.execute("""
        SELECT iv.item_id, iv.year, iv.value, iv.dataset_type
        FROM unified_values iv
        JOIN unified_items ui ON ui.id = iv.item_id
        WHERE ui.company_id = %s
        ORDER BY iv.year
    """, (cid,))
    vals = cur.fetchall()
    cur.close()
    conn.close()

    by_item = {}
    for v in vals:
        d = by_item.setdefault(v["item_id"], {"actual": {}, "projection": {}})
        d[v["dataset_type"]][v["year"]] = v["value"]

    out = []
    for it in items:
        series = by_item.get(it["id"], {"actual": {}, "projection": {}})
        cur_v = series["actual"].get(BASE_YEAR)
        prev_v = series["actual"].get(BASE_YEAR - 1)
        trend = None
        if cur_v is not None and prev_v not in (None, 0):
            trend = (cur_v - prev_v) / abs(prev_v) * 100
        out.append(dict(
            id=it["id"], name=it["name"], type=it["type"], category=it["category"],
            unit=it["unit"], unit_label=it["unit_label"],
            is_dashboard=it["is_dashboard"], in_chart=it["in_chart"],
            current=cur_v, prev=prev_v, trend=trend,
            actual=series["actual"], projection=series["projection"],
        ))
    return out


def chart_series_for(items):
    """Build Chart.js-ready actual/projection line series for in_chart items."""
    palette = ["--c1", "--c2", "--c3", "--c4", "--c5", "--c6"]
    years = list(range(HIST_START, PROJ_END + 1))
    series = []
    ci = 0
    for it in items:
        if not it["in_chart"]:
            continue
        color_idx = ci % len(palette)
        ci += 1
        actual = [it["actual"].get(y) for y in years]
        # projection line starts where actuals end (continuous), else None
        proj = [it["projection"].get(y) if y >= BASE_YEAR else None for y in years]
        series.append(dict(
            label=it["name"], color_idx=color_idx, category=it["category"],
            actual=actual, projection=proj, unit=it["unit"], unit_label=it["unit_label"],
        ))
    return dict(years=years, base_year=BASE_YEAR, series=series)


# ─── CONTEXT PROCESSOR (sidebar) ───────────────────────────────────
@app.context_processor
def inject_nav():
    if not session.get("user_id"):
        return {}
    comps = all_companies()
    industries = {}
    for c in comps:
        industries.setdefault(c["industry"], 0)
        industries[c["industry"]] += 1
    industry_order = ["Real Estate", "Energy", "Consumer Products"]
    industry_list = [(i, industries[i]) for i in industry_order if i in industries]
    for i, n in industries.items():
        if i not in industry_order:
            industry_list.append((i, n))
    return dict(
        nav_companies=comps,
        nav_industries=industry_list,
        current_path=request.path,
        display_name=session.get("display_name"),
        is_admin=session.get("is_admin"),
        cur_user_id=session.get("user_id"),
        cur_has_avatar=session.get("has_avatar"),
        cur_avatar_ver=session.get("avatar_ver", 0),
    )


def _avatar_initials(name):
    parts = [p for p in name.split() if p]
    if len(parts) >= 2:
        return (parts[0][0] + parts[1][0]).upper()
    return name[:2].upper()


app.jinja_env.filters["initials"] = _avatar_initials


# ─── PAGES ─────────────────────────────────────────────────────────
@app.route("/")
@login_required
def dashboard():
    rate = usd_mxn_rate()
    comps = company_cashflow_totals()
    us_usd = sum(c["investment"] for c in comps if c["country_code"] == "US")
    mx_usd = sum(c["investment"] for c in comps if c["country_code"] == "MX")
    total_usd = us_usd + mx_usd

    by_country = []
    for code, label in [("US", "United States"), ("MX", "Mexico")]:
        amt = sum(c["investment"] for c in comps if c["country_code"] == code)
        if amt:
            by_country.append(dict(code=code, label=label, usd=amt,
                                   pct=(amt / total_usd * 100) if total_usd else 0))

    by_industry = {}
    for c in comps:
        by_industry.setdefault(c["industry"], 0)
        by_industry[c["industry"]] += c["investment"]
    industry_rows = [dict(industry=k, usd=v, pct=(v / total_usd * 100) if total_usd else 0)
                     for k, v in sorted(by_industry.items(), key=lambda kv: -kv[1])]

    return render_template(
        "dashboard.html",
        rate=rate,
        us_usd=us_usd, mx_usd=mx_usd, mx_mxn=mx_usd * rate, total_usd=total_usd,
        us_count=sum(1 for c in comps if c["country_code"] == "US"),
        mx_count=sum(1 for c in comps if c["country_code"] == "MX"),
        total_count=len(comps),
        by_country=by_country, by_industry=industry_rows,
        base_year=BASE_YEAR,
    )


@app.route("/maquina-cashflow")
@login_required
def portfolio():
    comps = company_cashflow_totals()
    total_inv = sum(c["investment"] for c in comps)
    total_dist = sum(c["distribution"] for c in comps)
    net = total_dist - total_inv
    roi = (total_dist / total_inv * 100) if total_inv else 0

    year_inv = sum(c["series"].get(BASE_YEAR, {}).get("inv", 0) for c in comps)
    year_dist = sum(c["series"].get(BASE_YEAR, {}).get("dist", 0) for c in comps)

    # cumulative timeline
    years = list(range(HIST_START, PROJ_END + 1))
    cum_inv, cum_dist = [], []
    ci = cd = 0.0
    for y in years:
        ci += sum(c["series"].get(y, {}).get("inv", 0) for c in comps)
        cd += sum(c["series"].get(y, {}).get("dist", 0) for c in comps)
        cum_inv.append(round(ci))
        cum_dist.append(round(cd))

    # per-company sparkline = cumulative net by year
    for c in comps:
        run = 0.0
        spark = []
        for y in years:
            s = c["series"].get(y, {})
            run += (s.get("dist", 0) - s.get("inv", 0))
            spark.append(round(run))
        c["spark"] = spark

    by_industry = {}
    by_country = {}
    for c in comps:
        by_industry[c["industry"]] = by_industry.get(c["industry"], 0) + c["investment"]
        by_country[c["country"]] = by_country.get(c["country"], 0) + c["investment"]

    # portfolio planning matrix: per company inv/dist by year
    return render_template(
        "portfolio.html",
        comps=comps, total_inv=total_inv, total_dist=total_dist, net=net, roi=roi,
        year_inv=year_inv, year_dist=year_dist, base_year=BASE_YEAR,
        timeline=dict(years=years, inv=cum_inv, dist=cum_dist, base_year=BASE_YEAR),
        by_industry=[dict(k=k, v=v) for k, v in sorted(by_industry.items(), key=lambda x: -x[1])],
        by_country=[dict(k=k, v=v) for k, v in sorted(by_country.items(), key=lambda x: -x[1])],
        years=years,
    )


@app.route("/strategy")
@login_required
def strategy():
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM strategy_phases ORDER BY phase_order")
    phases = cur.fetchall()
    cur.execute("""
        SELECT c.id, c.slug, c.name, c.accent, c.industry, c.takeover_year,
               (c.logo IS NOT NULL) AS has_logo, octet_length(c.logo) AS logo_ver,
               r.internal_risk, r.external_risk
        FROM companies c LEFT JOIN company_risks r ON r.company_id = c.id
        WHERE c.archived = FALSE ORDER BY c.display_order
    """)
    comps = cur.fetchall()
    # current phase per company
    cur.execute("""
        SELECT ph.company_id, sp.phase_order, sp.name, sp.color, ph.start_year, ph.end_year
        FROM company_phase_history ph JOIN strategy_phases sp ON sp.id = ph.phase_id
        ORDER BY ph.company_id, sp.phase_order
    """)
    hist = cur.fetchall()
    cur.close()
    conn.close()

    cur_phase = {}
    for h in hist:
        cur_phase[h["company_id"]] = h  # last wins → highest phase_order
    comp_list = []
    for c in comps:
        cp = cur_phase.get(c["id"])
        comp_list.append(dict(
            id=c["id"], slug=c["slug"], name=c["name"], accent=c["accent"], industry=c["industry"],
            has_logo=c["has_logo"], logo_ver=c["logo_ver"],
            takeover_year=c["takeover_year"],
            internal=c["internal_risk"] or 5, external=c["external_risk"] or 5,
            phase_name=cp["name"] if cp else "—",
            phase_order=cp["phase_order"] if cp else 0,
            phase_color=cp["color"] if cp else "#8A9199",
            years_held=BASE_YEAR - c["takeover_year"] if c["takeover_year"] else None,
        ))

    # group companies under each phase
    for p in phases:
        p_companies = [c for c in comp_list if c["phase_order"] == p["phase_order"]]
        p["_companies"] = p_companies

    return render_template("strategy.html", phases=phases, comps=comp_list, base_year=BASE_YEAR)


def fetch_ember_operations():
    """Latest Ember operations report as {years, rows{label:[vals]}, totals, kpis, as_of}.
    Raw USD. Returns None on any problem so the caller falls back to seeded data."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        ecur.execute("SELECT data FROM reports WHERE report_type = 'operations' "
                     "ORDER BY uploaded_at DESC LIMIT 1")
        row = ecur.fetchone()
        ecur.close(); econn.close()
        if not row or not row.get("data"):
            return None
        d = row["data"]
        yr = d.get("yearly_rollup") or {}
        years, totals = yr.get("years") or [], yr.get("totals") or []
        if not years or not totals:
            return None
        rows = {r["label"]: r["values"] for r in (yr.get("rows") or []) if isinstance(r, dict)}
        kpis = {k.get("label"): k.get("value") for k in (d.get("kpis") or []) if isinstance(k, dict)}
        return {"years": years, "rows": rows, "totals": totals, "kpis": kpis, "as_of": d.get("date")}
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember operations fetch failed: %s", e)
        return None


@app.route("/company/<slug>")
@login_required
def company(slug):
    c = company_by_slug(slug)
    if not c or c["archived"]:
        abort(404)
    items = load_company_items(c["id"])
    dashboard_kpis = [it for it in items if it["is_dashboard"]]
    chart = chart_series_for(items)

    by_category = {"Commercial": [], "Operations": [], "Finance": []}
    for it in items:
        by_category.setdefault(it["category"], []).append(it)

    # Live Ember overlay — real operating revenues from the Ember DB (seed fallback).
    ember_live, ember_asof = False, None
    if c["slug"] == "ember":
        op = fetch_ember_operations()
        if op and op["totals"]:
            ember_live, ember_asof = True, op["as_of"]
            yrs, rows, totals = op["years"], op["rows"], op["totals"]
            yr0 = str(yrs[0])
            first = lambda lbl: (rows.get(lbl) or [None])[0]
            # Overhead proxy until Ember uploads full overhead: Project Personnel + 10%.
            overhead0 = (rows.get("Project Personnel") or [0])[0] * 1.10
            dashboard_kpis = [
                dict(name="Corporate Revenues", current=totals[0], unit="currency", unit_label=yr0, trend=None),
                dict(name="Development Fees", current=first("Development Fees"), unit="currency", unit_label=yr0, trend=None),
                dict(name="Bookkeeping Fee", current=first("Bookkeeping"), unit="currency", unit_label=yr0, trend=None),
                dict(name="Corporate Cashflow", current=(totals[0] - overhead0), unit="currency",
                     unit_label="net of est. overhead · " + yr0, trend=None),
            ]
            chart = dict(years=yrs, base_year=yrs[-1],
                         series=[dict(label=lbl, color_idx=i, category="Finance",
                                      actual=vals, projection=[None] * len(yrs),
                                      unit="currency", unit_label="USD")
                                 for i, (lbl, vals) in enumerate(rows.items())])

    # strategy tab data
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM company_risks WHERE company_id = %s", (c["id"],))
    risk = cur.fetchone()
    cur.execute("SELECT * FROM company_strategies WHERE company_id = %s ORDER BY start_year", (c["id"],))
    strategies = cur.fetchall()
    cur.execute("""
        SELECT ph.id AS hid, ph.phase_id, sp.name, sp.phase_order, sp.color, sp.year_range,
               sp.team_expectation, ph.start_year, ph.end_year, ph.is_projected
        FROM company_phase_history ph JOIN strategy_phases sp ON sp.id = ph.phase_id
        WHERE ph.company_id = %s ORDER BY ph.start_year, sp.phase_order
    """, (c["id"],))
    phase_hist = cur.fetchall()
    cur.execute("SELECT id, name, phase_order FROM strategy_phases ORDER BY phase_order")
    phases_all = cur.fetchall()
    cur.close()
    conn.close()

    return render_template(
        "company.html",
        c=c, items=items, dashboard_kpis=dashboard_kpis, by_category=by_category,
        chart=chart, risk=risk, strategies=strategies,
        phase_hist=phase_hist, phases_all=phases_all, base_year=BASE_YEAR,
        years=list(range(HIST_START, PROJ_END + 1)),
        ember_live=ember_live, ember_asof=ember_asof,
    )


# ─── COMPANY STRATEGY / RISK EDITING (admin) ───────────────────────
def _company_or_404(slug):
    c = company_by_slug(slug)
    if not c or c["archived"]:
        abort(404)
    return c


def _to_int(v, default=None):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _rating(v, default=5):
    n = _to_int(v, default)
    return max(0, min(10, n if n is not None else default))


@app.route("/company/<slug>/risk", methods=["POST"])
@login_required
def update_risk(slug):
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form
    commentary = (f.get("commentary") or "").strip() or None
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(
        """INSERT INTO company_risks
             (company_id, market_risk, team_risk, finance_risk, product_risk, internal_risk, external_risk, commentary)
           VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
           ON CONFLICT (company_id) DO UPDATE SET
             market_risk=EXCLUDED.market_risk, team_risk=EXCLUDED.team_risk,
             finance_risk=EXCLUDED.finance_risk, product_risk=EXCLUDED.product_risk,
             internal_risk=EXCLUDED.internal_risk, external_risk=EXCLUDED.external_risk,
             commentary=EXCLUDED.commentary""",
        (c["id"], _rating(f.get("market_risk")), _rating(f.get("team_risk")),
         _rating(f.get("finance_risk")), _rating(f.get("product_risk")),
         _rating(f.get("internal_risk")), _rating(f.get("external_risk")), commentary))
    conn.commit(); cur.close(); conn.close()
    flash("Risk ratings updated.", "ok")
    return redirect(url_for("company", slug=slug, tab="strategy"))


@app.route("/company/<slug>/strategy/add", methods=["POST"])
@login_required
def add_strategy(slug):
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form
    name = (f.get("name") or "").strip()
    if not name:
        flash("Initiative name is required.", "error")
        return redirect(url_for("company", slug=slug, tab="strategy"))
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(
        "INSERT INTO company_strategies (company_id,name,approach,start_year,end_year,formulation_rating,execution_rating) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s)",
        (c["id"], name, f.get("approach", "better"), _to_int(f.get("start_year")), _to_int(f.get("end_year")),
         _rating(f.get("formulation_rating")), _rating(f.get("execution_rating"))))
    conn.commit(); cur.close(); conn.close()
    flash("Initiative added.", "ok")
    return redirect(url_for("company", slug=slug, tab="strategy"))


@app.route("/company/<slug>/strategy/<int:sid>/save", methods=["POST"])
@login_required
def save_strategy(slug, sid):
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form
    if f.get("delete") == "1":
        q, args, msg = ("DELETE FROM company_strategies WHERE id=%s AND company_id=%s",
                        (sid, c["id"]), "Initiative removed.")
    else:
        q = ("UPDATE company_strategies SET name=%s, approach=%s, start_year=%s, end_year=%s, "
             "formulation_rating=%s, execution_rating=%s WHERE id=%s AND company_id=%s")
        args = ((f.get("name") or "").strip() or "Untitled", f.get("approach", "better"),
                _to_int(f.get("start_year")), _to_int(f.get("end_year")),
                _rating(f.get("formulation_rating")), _rating(f.get("execution_rating")), sid, c["id"])
        msg = "Initiative updated."
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(q, args)
    conn.commit(); cur.close(); conn.close()
    flash(msg, "ok")
    return redirect(url_for("company", slug=slug, tab="strategy"))


@app.route("/company/<slug>/phase/add", methods=["POST"])
@login_required
def add_phase(slug):
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form
    pid = _to_int(f.get("phase_id"))
    if not pid:
        flash("Pick a phase.", "error")
        return redirect(url_for("company", slug=slug, tab="strategy"))
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(
        "INSERT INTO company_phase_history (company_id,phase_id,start_year,end_year,is_projected) "
        "VALUES (%s,%s,%s,%s,%s)",
        (c["id"], pid, _to_int(f.get("start_year")), _to_int(f.get("end_year")),
         f.get("is_projected") == "1"))
    conn.commit(); cur.close(); conn.close()
    flash("Lifecycle phase added.", "ok")
    return redirect(url_for("company", slug=slug, tab="strategy"))


@app.route("/company/<slug>/phase/<int:hid>/save", methods=["POST"])
@login_required
def save_phase(slug, hid):
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form
    if f.get("delete") == "1":
        q, args, msg = ("DELETE FROM company_phase_history WHERE id=%s AND company_id=%s",
                        (hid, c["id"]), "Lifecycle phase removed.")
    else:
        q = ("UPDATE company_phase_history SET phase_id=%s, start_year=%s, end_year=%s, is_projected=%s "
             "WHERE id=%s AND company_id=%s")
        args = (_to_int(f.get("phase_id")), _to_int(f.get("start_year")), _to_int(f.get("end_year")),
                f.get("is_projected") == "1", hid, c["id"])
        msg = "Lifecycle phase updated."
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(q, args)
    conn.commit(); cur.close(); conn.close()
    flash(msg, "ok")
    return redirect(url_for("company", slug=slug, tab="strategy"))


@app.route("/manage-companies")
@login_required
def manage_companies():
    comps = all_companies(include_archived=True)
    # attach KPI + cashflow counts
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_id, COUNT(*) AS n FROM unified_items GROUP BY company_id")
    kpi_counts = {r["company_id"]: r["n"] for r in cur.fetchall()}
    cur.close()
    conn.close()
    return render_template("manage_companies.html", comps=comps, kpi_counts=kpi_counts)


@app.route("/manage-companies/<int:cid>/archive", methods=["POST"])
@login_required
def archive_company(cid):
    if not session.get("is_admin"):
        abort(403)
    archived = request.form.get("archived") == "1"
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("UPDATE companies SET archived = %s WHERE id = %s", (archived, cid))
    conn.commit()
    cur.close()
    conn.close()
    flash("Company updated.", "ok")
    return redirect(url_for("manage_companies"))


def ember_diagnostics():
    """Read-only probe of the Ember Postgres for the admin diagnostic card.
    Never writes. Returns connection status, available report types, and the
    shape of the latest 'operations' report so we can map KPIs precisely."""
    info = {"configured": bool(os.environ.get("EMBER_DATABASE_URL", "").strip()),
            "connected": False, "error": None, "report_types": [], "operations": None}
    if not info["configured"]:
        return info
    try:
        econn = db.get_ember_db()
        if not econn:
            return info
        ecur = econn.cursor()
        ecur.execute("SELECT report_type, COUNT(*) AS n, MAX(uploaded_at) AS last "
                     "FROM reports GROUP BY report_type ORDER BY report_type")
        info["report_types"] = [
            (r["report_type"], r["n"], r["last"].strftime("%Y-%m-%d") if r["last"] else "—")
            for r in ecur.fetchall()
        ]
        ecur.execute("SELECT data FROM reports WHERE report_type = 'operations' "
                     "ORDER BY uploaded_at DESC LIMIT 1")
        op = ecur.fetchone()
        if op and op.get("data"):
            d = op["data"]
            info["operations"] = {
                "fields": sorted(d.keys()),
                "kpis": [(k.get("label"), k.get("value")) for k in (d.get("kpis") or []) if isinstance(k, dict)],
                "raw": json.dumps({"yearly_rollup": d.get("yearly_rollup")}, default=str, ensure_ascii=False)[:2800],
            }
        info["connected"] = True
        ecur.close(); econn.close()
    except Exception as e:
        info["error"] = str(e)[:300]
    return info


@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings():
    conn = db.get_db()
    cur = conn.cursor()
    if request.method == "POST" and session.get("is_admin"):
        rate = request.form.get("usd_mxn_rate", "").strip()
        if rate:
            cur.execute(
                "INSERT INTO app_settings (key,value) VALUES ('usd_mxn_rate',%s) "
                "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value", (rate,))
            conn.commit()
        flash("Settings saved.", "ok")
    cur.execute("SELECT value FROM app_settings WHERE key='usd_mxn_rate'")
    row = cur.fetchone()
    # users for admin
    users = []
    if session.get("is_admin"):
        cur.execute("SELECT id, username, first_name, last_name, email, is_admin, "
                    "(avatar IS NOT NULL) AS has_avatar, octet_length(avatar) AS avatar_ver "
                    "FROM users ORDER BY id")
        users = cur.fetchall()
    cur.close()
    conn.close()
    ember = ember_diagnostics() if session.get("is_admin") else None
    return render_template("settings.html",
                           usd_mxn_rate=row["value"] if row else "17.3328",
                           live_rate=usd_mxn_rate(), users=users, ember=ember,
                           banxico_on=bool(os.environ.get("BANXICO_TOKEN", "").strip()))


@app.route("/account/password", methods=["POST"])
@login_required
def change_password():
    cur_pw = request.form.get("current_password", "")
    new_pw = request.form.get("new_password", "")
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE id = %s", (session["user_id"],))
    user = cur.fetchone()
    if user and check_password_hash(user["password_hash"], cur_pw) and len(new_pw) >= 6:
        cur.execute("UPDATE users SET password_hash = %s WHERE id = %s",
                    (generate_password_hash(new_pw), session["user_id"]))
        conn.commit()
        flash("Password updated.", "ok")
    else:
        flash("Could not update password — check your current password (new must be 6+ chars).", "error")
    cur.close()
    conn.close()
    return redirect(url_for("settings"))


# ─── IMAGES: serve + upload ────────────────────────────────────────
@app.route("/avatar/<int:uid>")
@login_required
def avatar(uid):
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("SELECT avatar, avatar_mime FROM users WHERE id = %s", (uid,))
    row = cur.fetchone(); cur.close(); conn.close()
    if not row or not row["avatar"]:
        abort(404)
    return Response(bytes(row["avatar"]), mimetype=row["avatar_mime"] or "image/jpeg",
                    headers={"Cache-Control": "public, max-age=300"})


@app.route("/company-logo/<int:cid>")
@login_required
def company_logo(cid):
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("SELECT logo, logo_mime FROM companies WHERE id = %s", (cid,))
    row = cur.fetchone(); cur.close(); conn.close()
    if not row or not row["logo"]:
        abort(404)
    return Response(bytes(row["logo"]), mimetype=row["logo_mime"] or "image/png",
                    headers={"Cache-Control": "public, max-age=300"})


@app.route("/settings/users", methods=["POST"])
@login_required
def create_user():
    if not session.get("is_admin"):
        abort(403)
    username = request.form.get("username", "").strip()
    pw = request.form.get("password", "")
    first = request.form.get("first_name", "").strip()
    last = request.form.get("last_name", "").strip()
    email = request.form.get("email", "").strip() or None
    is_admin = request.form.get("is_admin") == "1"
    if not username or len(pw) < 6:
        flash("A username and a password of at least 6 characters are required.", "error")
        return redirect(url_for("settings"))

    avatar = avatar_mime = None
    f = request.files.get("avatar")
    if f and f.filename:
        try:
            avatar, avatar_mime = process_image(f, "cover")
        except Exception:
            flash("Could not read that image — use a JPG or PNG.", "error")
            return redirect(url_for("settings"))

    conn = db.get_db(); cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash, is_admin, first_name, last_name, email, avatar, avatar_mime) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
            (username, generate_password_hash(pw), is_admin, first, last, email,
             psycopg2.Binary(avatar) if avatar else None, avatar_mime),
        )
        conn.commit()
        flash(f"Added team member “{username}”.", "ok")
    except psycopg2.errors.UniqueViolation:
        conn.rollback()
        flash("That username is already taken.", "error")
    finally:
        cur.close(); conn.close()
    return redirect(url_for("settings"))


@app.route("/settings/users/<int:uid>/delete", methods=["POST"])
@login_required
def delete_user(uid):
    if not session.get("is_admin"):
        abort(403)
    if uid == session.get("user_id"):
        flash("You can’t remove your own account.", "error")
        return redirect(url_for("settings"))
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id = %s", (uid,))
    conn.commit(); cur.close(); conn.close()
    flash("Team member removed.", "ok")
    return redirect(url_for("settings"))


@app.route("/settings/users/<int:uid>/role", methods=["POST"])
@login_required
def set_user_role(uid):
    if not session.get("is_admin"):
        abort(403)
    # Block changing your own role so there's always at least one admin
    # (you) and no one can accidentally lock themselves out.
    if uid == session.get("user_id"):
        flash("You can’t change your own role.", "error")
        return redirect(url_for("settings"))
    make_admin = request.form.get("make_admin") == "1"
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("UPDATE users SET is_admin = %s WHERE id = %s", (make_admin, uid))
    conn.commit(); cur.close(); conn.close()
    flash(f"Role updated to {'administrator' if make_admin else 'member'}.", "ok")
    return redirect(url_for("settings"))


@app.route("/account/avatar", methods=["POST"])
@login_required
def account_avatar():
    f = request.files.get("avatar")
    if not f or not f.filename:
        flash("Choose an image first.", "error")
        return redirect(request.referrer or url_for("settings"))
    try:
        data, mime = process_image(f, "cover")
    except Exception:
        flash("Could not read that image — use a JPG or PNG.", "error")
        return redirect(request.referrer or url_for("settings"))
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("UPDATE users SET avatar = %s, avatar_mime = %s WHERE id = %s",
                (psycopg2.Binary(data), mime, session["user_id"]))
    conn.commit(); cur.close(); conn.close()
    session["has_avatar"] = True
    session["avatar_ver"] = len(data)
    flash("Profile picture updated.", "ok")
    return redirect(request.referrer or url_for("settings"))


@app.route("/manage-companies/<int:cid>/logo", methods=["POST"])
@login_required
def upload_company_logo(cid):
    if not session.get("is_admin"):
        abort(403)
    f = request.files.get("logo")
    if not f or not f.filename:
        flash("Choose an image first.", "error")
        return redirect(url_for("manage_companies"))
    try:
        data, mime = process_image(f, "contain")
    except Exception:
        flash("Could not read that image — use a JPG or PNG.", "error")
        return redirect(url_for("manage_companies"))
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("UPDATE companies SET logo = %s, logo_mime = %s WHERE id = %s",
                (psycopg2.Binary(data), mime, cid))
    conn.commit(); cur.close(); conn.close()
    flash("Company logo updated.", "ok")
    return redirect(url_for("manage_companies"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)), debug=True)
