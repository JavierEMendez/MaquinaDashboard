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
import calendar
import datetime
import functools

import requests
import psycopg2
from PIL import Image, ImageOps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, abort, Response)
from werkzeug.security import check_password_hash, generate_password_hash

import db
import seed_data
import maquina_cf_parser

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
        # one-time re-categorization of existing KPIs to the new tab taxonomy
        conn = db.get_db()
        cur = conn.cursor()
        cur.execute("SELECT value FROM app_settings WHERE key = 'cat_remap_v1'")
        if not cur.fetchone():
            seed_data.remap_categories(conn)
            cur.execute("INSERT INTO app_settings (key,value) VALUES ('cat_remap_v1','1') "
                        "ON CONFLICT (key) DO NOTHING")
            conn.commit()
        cur.execute("SELECT value FROM app_settings WHERE key = 'financials_v1'")
        if not cur.fetchone():
            seed_data.seed_financials(conn)
            cur.execute("INSERT INTO app_settings (key,value) VALUES ('financials_v1','1') "
                        "ON CONFLICT (key) DO NOTHING")
            conn.commit()
        # v2: correct the initial placeholder defaults once (pre-edit)
        cur.execute("SELECT value FROM app_settings WHERE key = 'financials_v2'")
        if not cur.fetchone():
            seed_data.seed_financials(conn, overwrite=True)
            cur.execute("INSERT INTO app_settings (key,value) VALUES ('financials_v2','1') "
                        "ON CONFLICT (key) DO NOTHING")
            conn.commit()
        # Seed hold/takeover dates once: default hold start to the takeover
        # year and a 5-year target hold (both editable per company afterwards).
        cur.execute("SELECT value FROM app_settings WHERE key = 'hold_dates_v1'")
        if not cur.fetchone():
            cur.execute("UPDATE companies SET hold_start_year = takeover_year "
                        "WHERE hold_start_year IS NULL AND takeover_year IS NOT NULL")
            cur.execute("UPDATE companies SET target_hold_years = 5 WHERE target_hold_years IS NULL")
            cur.execute("INSERT INTO app_settings (key,value) VALUES ('hold_dates_v1','1') "
                        "ON CONFLICT (key) DO NOTHING")
            conn.commit()
        # Add Meta (toll-road infrastructure, Valoran) to the portfolio once.
        cur.execute("SELECT value FROM app_settings WHERE key = 'meta_company_v1'")
        if not cur.fetchone():
            cur.execute("SELECT id FROM companies WHERE slug = 'meta'")
            if not cur.fetchone():
                cur.execute(
                    "INSERT INTO companies (slug,name,industry,country,country_code,currency,"
                    "value_unit,accent,description,takeover_year,display_order) VALUES "
                    "('meta','Meta','Infrastructure','Mexico','MX','MXN','MMXN','#7A5AF8',"
                    "'Toll-road infrastructure concessions (Valoran) — highway development "
                    "and operation in Mexico.',NULL,5) RETURNING id")
                mid = cur.fetchone()["id"]
                cur.execute("INSERT INTO company_financials (company_id, valuation_model) "
                            "VALUES (%s,'ebitda') ON CONFLICT (company_id) DO NOTHING", (mid,))
            cur.execute("INSERT INTO app_settings (key,value) VALUES ('meta_company_v1','1') "
                        "ON CONFLICT (key) DO NOTHING")
            conn.commit()
        # 12-month retention on the activity ledger (once per process boot).
        cur.execute("DELETE FROM activity_log WHERE ts < NOW() - INTERVAL '12 months'")
        conn.commit()
        cur.close()
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
            _log_activity("login", user_id=user["id"], username=user["username"])
            nxt = request.args.get("next") or url_for("dashboard")
            return redirect(nxt)
        error = "Invalid username or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    _log_activity("logout", user_id=session.get("user_id"), username=session.get("username"))
    session.clear()
    return redirect(url_for("login"))


# ─── ACTIVITY TRACKING ─────────────────────────────────────────────
# Per-user page-view + login/logout capture for the admin Team Activity
# page. User-level only (no IPs, no user agents); 12-month retention is
# enforced by a once-per-boot purge in _boot().
_ACTIVITY_DENY_PREFIXES = ("/static/", "/api/", "/health", "/favicon",
                           "/login", "/logout", "/avatar/", "/company-logo/")
_ACTIVITY_DENY_SUFFIXES = (".json", ".css", ".js", ".ico", ".png", ".svg",
                           ".jpg", ".jpeg", ".gif", ".woff", ".woff2", ".map")

# Friendly labels for the paths the page surfaces. Anything not mapped
# falls back to a readable name (company pages resolve to the company).
_ACTIVITY_PATH_LABELS = {
    "/":                 "Portfolio Dashboard",
    "/strategy":         "Maquina Strategy",
    "/maquina-cashflow":  "Maquina Portfolio",
    "/cashflow":         "Maquina Cashflow",
    "/manage-companies":  "Manage Companies",
    "/settings":         "Settings",
    "/activity":         "Team Activity",
}


def _activity_label(path):
    if not path:
        return "—"
    if path in _ACTIVITY_PATH_LABELS:
        return _ACTIVITY_PATH_LABELS[path]
    if path.startswith("/company/"):
        parts = path.split("/")
        c = company_by_slug(parts[2]) if len(parts) > 2 else None
        return c["name"] if c else "Company"
    return path


def _log_activity(event_type, path=None, user_id=None, username=None):
    """Insert one activity_log row. Swallows all errors — a logging
    hiccup must never take down a real request."""
    try:
        uid = user_id if user_id is not None else session.get("user_id")
        uname = username or session.get("username")
        if not uid or not uname:
            return
        conn = db.get_db()
        cur = conn.cursor()
        cur.execute("INSERT INTO activity_log (user_id, username, event_type, path) "
                    "VALUES (%s, %s, %s, %s)", (uid, uname, event_type, path))
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:  # pragma: no cover
        app.logger.warning("activity log failed (%s): %s", event_type, e)


@app.before_request
def _capture_page_view():
    """Log GET requests to authenticated dashboard pages. Logins/logouts
    are captured by their own handlers; POSTs, AJAX and asset fetches are
    ignored so the log stays focused on 'who looked at what'."""
    if request.method != "GET" or not session.get("user_id"):
        return
    p = request.path or ""
    if p.startswith(_ACTIVITY_DENY_PREFIXES) or p.endswith(_ACTIVITY_DENY_SUFFIXES):
        return
    _log_activity("page_view", path=p)


def _sparkline(values, color="#0568B3", bold=False):
    """Server-rendered SVG sparkline (88×22)."""
    w, h, pad = 88, 22, 2
    if not values:
        return ""
    vmin = min(values)
    rng = (max(values) - vmin) or 1
    step = (w - pad * 2) / max(1, len(values) - 1)
    parts = []
    for i, v in enumerate(values):
        x = pad + i * step
        y = pad + (1 - (v - vmin) / rng) * (h - pad * 2)
        parts.append(f"{'M' if i == 0 else 'L'} {x:.1f} {y:.1f}")
    return (f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}">'
            f'<path d="{" ".join(parts)}" fill="none" stroke="{color}" '
            f'stroke-width="{1.8 if bold else 1.2}" stroke-linecap="round" '
            f'stroke-linejoin="round" opacity="{1 if bold else 0.85}"/></svg>')


@app.route("/activity")
@login_required
def activity():
    """Team Activity — admin-only per-user roll-up of logins and page
    views (7d / 30d) with a 14-day sparkline and most-visited pages."""
    if not session.get("is_admin"):
        return render_template("activity.html", forbidden=True), 403

    conn = db.get_db()
    cur = conn.cursor()

    # ── KPI band (last 7 days) ──
    cur.execute("SELECT COUNT(DISTINCT user_id) AS n FROM activity_log "
                "WHERE ts >= NOW() - INTERVAL '7 days' AND user_id IS NOT NULL")
    active_7d = (cur.fetchone() or {}).get("n") or 0

    cur.execute("SELECT COUNT(*) AS n FROM activity_log "
                "WHERE event_type = 'login' AND ts >= NOW() - INTERVAL '7 days'")
    logins_7d = (cur.fetchone() or {}).get("n") or 0

    cur.execute("SELECT path, COUNT(*) AS n FROM activity_log "
                "WHERE event_type = 'page_view' AND ts >= NOW() - INTERVAL '7 days' "
                "AND path IS NOT NULL GROUP BY path ORDER BY n DESC LIMIT 1")
    top_row = cur.fetchone()
    top_page_label = _activity_label(top_row["path"]) if top_row else "—"
    top_page_count = top_row["n"] if top_row else 0

    cur.execute("SELECT COUNT(*) AS n FROM activity_log "
                "WHERE event_type = 'page_view' AND ts >= NOW() - INTERVAL '7 days'")
    views_7d = (cur.fetchone() or {}).get("n") or 0
    avg_views = round(views_7d / active_7d, 1) if active_7d else 0

    # ── Per-user roll-up ──
    cur.execute("SELECT id AS user_id, username, is_admin, "
                "COALESCE(first_name, '') AS first_name, "
                "COALESCE(last_name, '') AS last_name FROM users ORDER BY username")
    user_rows = cur.fetchall()

    cur.execute("""
        SELECT user_id, MAX(ts) AS last_seen,
               COUNT(*) FILTER (WHERE event_type='login'     AND ts >= NOW() - INTERVAL '7 days')  AS logins_7d,
               COUNT(*) FILTER (WHERE event_type='login'     AND ts >= NOW() - INTERVAL '30 days') AS logins_30d,
               COUNT(*) FILTER (WHERE event_type='page_view' AND ts >= NOW() - INTERVAL '7 days')  AS views_7d,
               COUNT(*) FILTER (WHERE event_type='page_view' AND ts >= NOW() - INTERVAL '30 days') AS views_30d
        FROM activity_log WHERE user_id IS NOT NULL GROUP BY user_id
    """)
    by_user = {r["user_id"]: r for r in cur.fetchall()}

    cur.execute("""
        SELECT user_id, DATE_TRUNC('day', ts)::date AS d, COUNT(*) AS n
        FROM activity_log
        WHERE event_type='page_view' AND ts >= (NOW() - INTERVAL '14 days')::date
          AND user_id IS NOT NULL
        GROUP BY user_id, d
    """)
    spark_rows = cur.fetchall()

    cur.execute("""
        SELECT user_id, path, COUNT(*) AS n
        FROM activity_log
        WHERE event_type='page_view' AND ts >= NOW() - INTERVAL '30 days'
          AND user_id IS NOT NULL AND path IS NOT NULL
        GROUP BY user_id, path
    """)
    top_paths = {}
    for r in cur.fetchall():
        top_paths.setdefault(r["user_id"], []).append((r["path"], r["n"]))
    for uid in top_paths:
        top_paths[uid].sort(key=lambda t: -t[1])

    cur.close()
    conn.close()

    today = datetime.date.today()
    days = [today - datetime.timedelta(days=13 - i) for i in range(14)]
    spark_by_user = {}
    for r in spark_rows:
        d = r["d"]
        if hasattr(d, "date"):
            d = d.date()
        spark_by_user.setdefault(r["user_id"], {})[d] = int(r["n"])

    rows = []
    for u in user_rows:
        agg = by_user.get(u["user_id"]) or {}
        last_seen = agg.get("last_seen")
        if last_seen and hasattr(last_seen, "isoformat"):
            last_seen_iso = last_seen.isoformat()
            last_seen_label = last_seen.strftime("%b %d, %H:%M UTC")
        else:
            last_seen_iso, last_seen_label = None, "Never"
        spark_values = [spark_by_user.get(u["user_id"], {}).get(d, 0) for d in days]
        spark_svg = _sparkline(spark_values, "#0568B3", bold=True) if any(spark_values) else ""
        top3 = top_paths.get(u["user_id"], [])[:3]
        display = (f"{u['first_name']} {u['last_name']}".strip()) or u["username"]
        rows.append(dict(
            user_id=u["user_id"], username=u["username"], display=display,
            is_admin=u["is_admin"], last_seen=last_seen_label, last_seen_iso=last_seen_iso,
            logins_7d=int(agg.get("logins_7d") or 0), logins_30d=int(agg.get("logins_30d") or 0),
            views_7d=int(agg.get("views_7d") or 0), views_30d=int(agg.get("views_30d") or 0),
            top_pages=[{"label": _activity_label(p), "n": n} for p, n in top3],
            spark_svg=spark_svg, active=bool(any(spark_values)),
        ))
    rows.sort(key=lambda r: (-r["views_7d"], r["display"].lower()))

    kpis = [
        dict(label="Active Users",      val=str(active_7d),  sub="with any activity · last 7d"),
        dict(label="Total Logins",      val=str(logins_7d),  sub="sign-ins · last 7d"),
        dict(label="Most-Visited Page", val=top_page_label,  sub=f"{top_page_count} views · last 7d"),
        dict(label="Avg Views / User",  val=str(avg_views),  sub="page views per active user · last 7d"),
    ]
    generated_at_iso = datetime.datetime.now(datetime.timezone.utc).isoformat()
    return render_template("activity.html", forbidden=False, kpis=kpis, rows=rows,
                           generated_at_iso=generated_at_iso)


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


def _act_sum(by_year, last_actual):
    """Sum a {year: value} map across actual years only (year <= last_actual).
    Year keys arrive as strings from JSONB; coerce defensively. Used by the
    Maquina Cashflow tables for the lifetime-actuals 'Total' column."""
    total = 0.0
    for k, v in (by_year or {}).items():
        try:
            if int(k) <= int(last_actual):
                total += float(v or 0)
        except (TypeError, ValueError):
            continue
    return total


def _all_sum(by_year):
    total = 0.0
    for v in (by_year or {}).values():
        try:
            total += float(v or 0)
        except (TypeError, ValueError):
            continue
    return total


def _period_val(series, period):
    """Value of a parsed series at one display period. Month periods read
    from series['months'] ('YYYY-MM'); year periods from series['by_year']."""
    if not series or not period:
        return 0.0
    d = series.get("months") if period.get("kind") == "month" else series.get("by_year")
    try:
        return float((d or {}).get(period.get("key")) or 0)
    except (TypeError, ValueError):
        return 0.0


def _period_total(series, periods):
    return sum(_period_val(series, p) for p in (periods or []))


app.jinja_env.filters["actsum"] = _act_sum
app.jinja_env.filters["allsum"] = _all_sum
app.jinja_env.filters["pval"] = _period_val
app.jinja_env.filters["ptotal"] = _period_total


def _cf_periods(meta):
    """Display axis for the Maquina Cashflow tables/charts: the full current
    year (= last year with monthly detail) broken out by month, then every
    later year as a single annual column. Older years are dropped — the view
    is current-year-monthly + forward annual outlook."""
    la = int(meta.get("last_actual_year") or 0)
    years = sorted({int(y) for y in (meta.get("years") or [])})
    periods = []
    if la:
        yy = str(la)[2:]
        for mo in range(1, 13):
            periods.append({"key": f"{la}-{mo:02d}",
                            "label": f"{calendar.month_abbr[mo]} '{yy}",
                            "kind": "month", "proj": False})
    for y in years:
        if y > la:
            periods.append({"key": str(y), "label": str(y), "kind": "year", "proj": True})
    return periods


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
        "accent, description, takeover_year, takeover_month, hold_start_year, "
        "hold_start_month, target_hold_years, display_order, archived, "
        "(logo IS NOT NULL) AS has_logo, octet_length(logo) AS logo_ver "
        "FROM companies WHERE slug = %s", (slug,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return row


def load_financials(cid):
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM company_financials WHERE company_id = %s", (cid,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return row


def _to_usd(value, value_unit, rate):
    """Normalize a display-unit value to raw USD."""
    if value is None:
        return 0.0
    if value_unit == "KUSD":
        return value * 1000.0
    if value_unit == "MMXN":
        return value * 1e6 / rate if rate else 0.0
    if value_unit == "KMXN":
        return value * 1e3 / rate if rate else 0.0
    return value


def company_valuation(fin, promote_total_usd, rate, value_unit):
    """Two-model equity valuation. Returns display-unit pieces + normalized USD.
    Sponsor: FRE × multiple + PV(promotes) (~3yr horizon). Else: EV − net debt."""
    if not fin:
        return None
    out = {"model": fin["valuation_model"]}
    if fin["valuation_model"] == "sponsor":
        fre_value = (fin["fre"] or 0) * (fin["fre_multiple"] or 0)        # display units
        cd = fin["carry_discount"] or 0.15
        promote_pv = (promote_total_usd or 0) / ((1 + cd) ** 3)           # raw USD
        out.update(fre_value=fre_value, promote_pv=promote_pv,
                   value_usd=_to_usd(fre_value, value_unit, rate) + promote_pv)
    else:
        ev = (fin["ebitda"] or 0) * (fin["ebitda_multiple"] or 0)         # display units
        equity = ev - (fin["total_debt"] or 0)
        out.update(ev=ev, equity=equity, value_usd=_to_usd(equity, value_unit, rate))
    return out


def annual_irr(cfs):
    """IRR for an annual net-cashflow list (period 0..n). None if no sign change.
    Bisection on NPV — no numpy dependency."""
    cfs = [float(x) for x in cfs]
    if not (any(x < 0 for x in cfs) and any(x > 0 for x in cfs)):
        return None

    def npv(r):
        return sum(cf / ((1 + r) ** i) for i, cf in enumerate(cfs))

    lo, hi = -0.95, 5.0
    flo, fhi = npv(lo), npv(hi)
    if flo * fhi > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2.0
        fm = npv(mid)
        if abs(fm) < 1e-2:
            return mid
        if flo * fm < 0:
            hi = mid
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2.0


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
        c["moic"] = (c["distribution"] / c["investment"]) if c["investment"] else 0
        ys = sorted(c["series"].keys())
        if ys:
            cfs = [c["series"].get(y, {}).get("dist", 0) - c["series"].get(y, {}).get("inv", 0)
                   for y in range(ys[0], ys[-1] + 1)]
            c["irr"] = annual_irr(cfs)
        else:
            c["irr"] = None
    return list(comps.values())


def _exit_calc(series, entry_year, target_hold_years, exit_value_usd):
    """Projected gross return to a potential exit. Pure (no DB).

    series: {year: {'inv': usd, 'dist': usd}} of actual + projected cashflows.
    Exit year = entry + target hold (never in the past). Cashflows AFTER the
    exit year are ignored (you've exited); the exit equity value lands in the
    exit year. MOIC = (distributions through exit + exit value) / invested;
    IRR = annualized over the dated net-cashflow stream."""
    today_y = datetime.date.today().year
    years = sorted(series.keys())
    entry_y = entry_year or (years[0] if years else today_y)
    if target_hold_years:
        exit_y = max(entry_y + int(round(float(target_hold_years))), today_y)
    else:
        exit_y = today_y
    in_window = [y for y in years if y <= exit_y]
    invested = sum((series[y].get("inv") or 0) for y in in_window)
    dist = sum((series[y].get("dist") or 0) for y in in_window)
    if not invested:
        return None
    ev = float(exit_value_usd or 0.0)
    start_y = in_window[0] if in_window else exit_y
    cfs = []
    for y in range(start_y, exit_y + 1):
        s = series.get(y) or {}
        net = (s.get("dist") or 0) - (s.get("inv") or 0)
        if y == exit_y:
            net += ev
        cfs.append(net)
    return dict(invested=invested, dist=dist, exit_value=ev, exit_year=exit_y,
                moic=(dist + ev) / invested, dpi=dist / invested, irr=annual_irr(cfs))


# ─── META: PROMOTE / COMPARTICIÓN DE INGRESOS (highway revenue share) ─
# Extracted 1-for-1 from Valoran's UPDATED proposal workbook
# ("Libro1 1.xlsx" → "Aux. Compartición", validated to the peso).
# Series in MILLIONS of MXN. Periods: an H2-2026 stub, then annual
# 2027–2050 (concession end; the workbook's 2051+ columns are zero).
# tdc = concession-title baseline revenue for the period's YEAR in
# Dec-2024 pesos (actualized +33.07% / +29.59% from title base dates —
# the stub period is compared against its FULL-year título value, per
# the workbook's own XLOOKUP logic); model = Valoran's revised revenue
# projection (nominal, ~130–150% of título vs 2–4× in the old draft);
# index = INPC index (~3.5–3.8%/yr). Nominal baseline_t = tdc_t ×
# index_t; excess over baseline is split into hurdles vs the baseline
# (0–30–50–75%+) promoted to Maquina at 5/40/60/75%. Base-case nominal
# total: $2,849.1 mdp (V $1,996.0 + P $853.1). The client-side engine
# in company.html recomputes everything from the sliders.
META_COMP = {
    "years": [2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035, 2036, 2037, 2038, 2039, 2040, 2041, 2042, 2043, 2044, 2045, 2046, 2047, 2048, 2049, 2050],
    "labels": ["2026 H2", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035", "2036", "2037", "2038", "2039", "2040", "2041", "2042", "2043", "2044", "2045", "2046", "2047", "2048", "2049", "2050"],
    "index": [1.038, 1.076406, 1.115695, 1.154744, 1.19516, 1.236991, 1.280285, 1.325095, 1.371474, 1.419475, 1.469157, 1.520578, 1.573798, 1.628881, 1.685891, 1.744898, 1.805969, 1.869178, 1.934599, 2.00231, 2.072391, 2.144925, 2.219997, 2.297697, 2.378116],
    "ventura": {
        "passages": [2.4024, 4.8848, 5.0696, 5.2428, 5.4368, 5.638, 5.8626, 6.0629, 6.2872, 6.5198, 6.7796, 7.0112, 7.2706, 7.5397, 7.84, 8.1079, 8.4079, 8.719, 9.0664, 9.3761, 9.7231, 10.0828, 10.4845, 10.8427, 11.2439],
        "toll_auto": 89.66, "toll_blend": 197.07, "toll_when": "Jul-2026",
        "name": "Ventura – El Peyote",
        "tdc": [678.89, 711.394, 746.653, 778.562, 811.681, 845.496, 882.432, 915.271, 951.261, 988.007, 1026.543, 1060.184, 1097.358, 1135.276, 1171.352, 1201.585, 1235.587, 1270.168, 1308.912, 1341.102, 1377.476, 1414.733, 1456.986, 1453.005, 1453.005],
        "model": [491.54, 1035.335, 1114.1, 1193.809, 1281.31, 1375.223, 1480.029, 1584.205, 1700.319, 1824.944, 1964.023, 2102.267, 2256.352, 2421.732, 2606.292, 2789.743, 2994.218, 3213.679, 3458.593, 3702.036, 3973.377, 4264.606, 4589.612, 4912.665, 5272.738]},
    "pitahaya": {
        "passages": [0.0, 0.0, 1.0248, 4.2162, 4.3722, 4.534, 4.7147, 4.8757, 5.0561, 5.2432, 5.4521, 5.6384, 5.847, 6.0634, 6.3049, 6.5203, 6.7616, 7.0118, 7.2911, 7.5402, 7.8192, 8.1085, 8.4316, 8.7197, 9.0423],
        "toll_auto": 138.70, "toll_blend": 374.12, "toll_when": "Oct-2028",
        "name": "La Pitahaya – Libramiento Oriente",
        "tdc": [1122.797, 1166.857, 1216.325, 1260.855, 1310.892, 1363.119, 1421.301, 1474.097, 1533.217, 1594.918, 1641.939, 1681.028, 1725.871, 1771.797, 1824.214, 1868.073, 1917.82, 1969.444, 2027.927, 2076.881, 2132.716, 2190.25, 2196.249, 2190.25, 2190.25],
        "model": [0.0, 0.0, 383.401, 1619.958, 1738.693, 1866.131, 2008.348, 2149.712, 2307.275, 2476.387, 2665.112, 2852.704, 3061.793, 3286.207, 3536.649, 3785.586, 4063.05, 4360.851, 4693.192, 5023.536, 5391.736, 5786.923, 6227.945, 6666.317, 7154.924]},
}


# v1 reference case — the ORIGINAL proposal workbook (Libro1.xlsx,
# "Compartición" sheet), kept only for the static v1-vs-v2 comparison
# card on the Promote Analysis tab. Same units/semantics as META_COMP.
META_COMP_V1 = {
    "years": [2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035, 2036, 2037, 2038, 2039, 2040, 2041, 2042, 2043, 2044, 2045, 2046, 2047, 2048, 2049, 2050],
    "index": [1.003113, 1.041231, 1.079965, 1.123163, 1.16809, 1.214814, 1.263406, 1.313942, 1.3665, 1.42116, 1.478007, 1.537127, 1.598612, 1.662556, 1.729059, 1.798221, 1.87015, 1.944956, 2.022754, 2.103664, 2.187811, 2.275323, 2.366336, 2.46099, 2.559429, 2.661806, 2.768279],
    "ventura": {
        "tdc": [616.268, 647.086, 678.89, 711.394, 746.653, 778.562, 811.681, 845.496, 882.432, 915.271, 951.261, 988.007, 1026.543, 1060.184, 1097.358, 1135.276, 1171.352, 1201.585, 1235.587, 1270.168, 1308.912, 1341.102, 1377.476, 1414.733, 1456.986, 1453.005, 1453.005],
        "model": [0.0, 1460.097, 1711.77, 1955.562, 2221.795, 2480.762, 2777.693, 3110.452, 3492.834, 3841.887, 4703.467, 5269.347, 5902.226, 6534.313, 7215.311, 7931.869, 8715.477, 9496.574, 10331.293, 11226.384, 12187.705, 13136.902, 14167.155, 15273.952, 16466.302, 17607.056, 14954.654]},
    "pitahaya": {
        "tdc": [719.828, 864.321, 1122.797, 1166.857, 1216.325, 1260.855, 1310.892, 1363.119, 1421.301, 1474.097, 1533.217, 1594.918, 1641.939, 1681.028, 1725.871, 1771.797, 1824.214, 1868.073, 1917.82, 1969.444, 2027.927, 2076.881, 2132.716, 2190.25, 2196.249, 2190.25, 2190.25],
        "model": [0.0, 0.0, 0.0, 0.0, 804.485, 3270.647, 3834.922, 4321.547, 4883.447, 5294.338, 6516.7, 7345.837, 8230.207, 9113.28, 10064.432, 11065.143, 12159.653, 13250.961, 14416.778, 15667.802, 17011.05, 18338.067, 19778.657, 21327.96, 22996.818, 24593.447, 20892.732]},
}


def _promote_case(mc, npv_base_idx, disc=0.09, tax=0.30):
    """Run one workbook's base case through the promote waterfall.
    Returns totals, NPV, effective share of excess, hurdle mix, first
    sharing year and model-vs-baseline %. Values in mdp (millions MXN)."""
    TH, SH = (0.30, 0.50, 0.75), (0.05, 0.40, 0.60, 0.75)
    years, ix = mc["years"], mc["index"]
    n = len(years)
    promote = [0.0] * n
    hurdles = [0.0] * 4
    pct, first = {}, {}
    total = excess = npv = 0.0
    for key in ("ventura", "pitahaya"):
        R = mc[key]
        pr = []
        first[key] = None
        for i in range(n):
            b = R["tdc"][i] * ix[i]
            m = R["model"][i]
            pr.append(round(m / b * 100, 1) if (m and b) else None)
            e = max(0.0, m - b)
            l1 = min(e, b * TH[0])
            l2 = min(max(e - b * TH[0], 0), b * (TH[1] - TH[0]))
            l3 = min(max(e - b * TH[1], 0), b * (TH[2] - TH[1]))
            l4 = max(e - b * TH[2], 0)
            parts = (l1 * SH[0], l2 * SH[1], l3 * SH[2], l4 * SH[3])
            s = sum(parts)
            for k in range(4):
                hurdles[k] += parts[k]
            promote[i] += s
            total += s
            excess += e
            if s > 1e-9 and first[key] is None:
                first[key] = years[i]
            npv += s * (1 - tax) / (ix[i] / ix[npv_base_idx]) / ((1 + disc) ** max(0, years[i] - years[npv_base_idx]))
        pct[key] = pr

    def _rng(vals):
        xs = [v for v in vals if v is not None]
        return "%.0f%%–%.0f%%" % (min(xs), max(xs)) if xs else "—"
    return dict(years=years, promote=[round(x, 1) for x in promote], pct=pct,
                total_nom=round(total, 1), npv=round(npv, 1),
                share_of_excess=(round(total / excess * 100, 1) if excess else None),
                hurdle_mix=[(round(h / total * 100, 1) if total else 0) for h in hurdles],
                first=first, rng_v=_rng(pct["ventura"]), rng_p=_rng(pct["pitahaya"]))


def _build_meta_compare():
    """Static v1-vs-v2 comparison payload for the Promote Analysis tab."""
    v1 = _promote_case(META_COMP_V1, npv_base_idx=1)   # discount to 2025 (first flow year)
    v2 = _promote_case(META_COMP, npv_base_idx=0)      # discount to H2-2026
    # v2's H2-2026 stub compares a half-year of revenue against the FULL-year
    # título value (the workbook's own XLOOKUP logic) — not a comparable %.
    for k in ("ventura", "pitahaya"):
        v2["pct"][k][0] = None

    def _rng(vals):
        xs = [v for v in vals if v is not None]
        return "%.0f%%–%.0f%%" % (min(xs), max(xs)) if xs else "—"
    v2["rng_v"] = _rng(v2["pct"]["ventura"])
    v2["rng_p"] = _rng(v2["pct"]["pitahaya"])
    pad = [None] * (len(META_COMP_V1["years"]) - len(META_COMP["years"]))
    return dict(
        axis=[str(y) for y in META_COMP_V1["years"]],
        chart=dict(
            v1_v=v1["pct"]["ventura"], v1_p=v1["pct"]["pitahaya"],
            v2_v=pad + v2["pct"]["ventura"], v2_p=pad + v2["pct"]["pitahaya"]),
        v1=dict(nom=v1["total_nom"], npv=v1["npv"], soe=v1["share_of_excess"],
                mix=v1["hurdle_mix"], first=v1["first"], rng_v=v1["rng_v"], rng_p=v1["rng_p"]),
        v2=dict(nom=v2["total_nom"], npv=v2["npv"], soe=v2["share_of_excess"],
                mix=v2["hurdle_mix"], first=v2["first"], rng_v=v2["rng_v"], rng_p=v2["rng_p"]),
        d_nom=round((v2["total_nom"] / v1["total_nom"] - 1) * 100, 1),
        d_npv=round((v2["npv"] / v1["npv"] - 1) * 100, 1),
    )


META_COMPARE = _build_meta_compare()


def _meta_base_assumptions(mc=None):
    """Derive the base assumptions embedded in the workbook so the slider
    labels can state what "0% delta" / "100% level" actually mean."""
    mc = mc or META_COMP
    Y, IX = mc["years"], mc["index"]
    infl_first = IX[1] / IX[0] - 1
    infl_last = IX[-1] / IX[-2] - 1
    out = {"infl_first": infl_first * 100, "infl_last": infl_last * 100,
           "infl_cagr": ((IX[-1] / IX[0]) ** (1 / (Y[-1] - Y[0])) - 1) * 100}
    for key, si in (("ventura", 1), ("pitahaya", 3)):   # first FULL year of each road
        R = mc[key]
        a, b, yrs = R["model"][si], R["model"][-1], Y[-1] - Y[si]
        nom = (b / a) ** (1 / yrs) - 1
        ixg = (IX[-1] / IX[si]) ** (1 / yrs) - 1
        # Decompose nominal revenue growth into its two independent drivers:
        # vehicles (real volume) and tarifa (nominal price per passage).
        # Revenue growth = (1+traffic) x (1+tarifa) - 1, exactly.
        p0, p1 = R["passages"][si], R["passages"][-1]
        pas = (p1 / p0) ** (1 / yrs) - 1
        tar = ((b / p1) / (a / p0)) ** (1 / yrs) - 1
        out[key] = {"start_year": Y[si], "start": a, "end": b,
                    "nom_cagr": nom * 100, "real_cagr": ((1 + nom) / (1 + ixg) - 1) * 100,
                    "pas_start": p0, "pas_end": p1,
                    "pas_cagr": pas * 100,                       # vehicles
                    "tar_cagr": tar * 100,                       # tarifa, nominal
                    "tar_real": ((1 + tar) / (1 + ixg) - 1) * 100}  # tarifa, real (~0)
    return out


META_BASE = _meta_base_assumptions()


def exit_returns(cid, entry_year, target_hold_years, exit_value_usd):
    """DB wrapper around _exit_calc for a single company."""
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT year, COALESCE(SUM(investment),0) AS inv, "
                "COALESCE(SUM(distribution),0) AS dist FROM portfolio_cashflows "
                "WHERE company_id = %s GROUP BY year", (cid,))
    series = {r["year"]: {"inv": r["inv"] or 0, "dist": r["dist"] or 0}
              for r in cur.fetchall() if r["year"] is not None}
    cur.close()
    conn.close()
    return _exit_calc(series, entry_year, target_hold_years, exit_value_usd)


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
        if not it["in_chart"] or it.get("coming_soon"):
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


@app.context_processor
def inject_helpers():
    """Template helpers available on every page."""
    return dict(fmt_month=_fmt_month)


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

    # blended portfolio IRR (aggregate net cashflow by year) + concentration
    agg_net = [sum(c["series"].get(y, {}).get("dist", 0) - c["series"].get(y, {}).get("inv", 0)
                   for c in comps) for y in years]
    blended_irr = annual_irr(agg_net)
    top = max(comps, key=lambda c: c["investment"]) if comps else None
    concentration = (top["investment"] / total_inv * 100) if (top and total_inv) else 0

    # estimated portfolio NAV (USD) = sum of per-company equity valuations
    rate = usd_mxn_rate()
    er = fetch_ember_returns()
    promote_total = er["totals"]["promote"] if er else 0
    conn = db.get_db(); cur = conn.cursor()
    cur.execute("SELECT cf.*, c.value_unit, c.slug FROM company_financials cf "
                "JOIN companies c ON c.id = cf.company_id WHERE c.archived = FALSE")
    fins = cur.fetchall(); cur.close(); conn.close()
    nav = sum((company_valuation(fr, promote_total if fr["slug"] == "ember" else 0, rate, fr["value_unit"]) or {}).get("value_usd", 0)
              for fr in fins)

    return render_template(
        "portfolio.html",
        comps=comps, total_inv=total_inv, total_dist=total_dist, net=net, roi=roi,
        moic=(roi / 100), blended_irr=blended_irr, concentration=concentration, nav=nav,
        top_holding=(top["name"] if top else None),
        year_inv=year_inv, year_dist=year_dist, base_year=BASE_YEAR,
        timeline=dict(years=years, inv=cum_inv, dist=cum_dist, base_year=BASE_YEAR),
        by_industry=[dict(k=k, v=v) for k, v in sorted(by_industry.items(), key=lambda x: -x[1])],
        by_country=[dict(k=k, v=v) for k, v in sorted(by_country.items(), key=lambda x: -x[1])],
        years=years,
    )


@app.route("/cashflow")
@login_required
def cashflow():
    """Maquina Cashflow — beautified view of the uploaded MAQUINA CF
    workbook (CF Consolidado + Business CF US/MX). A stopgap snapshot
    until per-company figures flow in live."""
    snap = db.latest_maquina_cf()
    cf = snap["data"] if snap else None
    periods = _cf_periods(cf["meta"]) if cf else []
    return render_template(
        "cashflow.html",
        cf=cf,
        periods=periods,
        uploaded_at=(snap["uploaded_at"] if snap else None),
        uploaded_by=(snap["uploaded_by"] if snap else None),
        src_filename=(snap["filename"] if snap else None),
    )


@app.route("/cashflow/upload", methods=["POST"])
@login_required
def cashflow_upload():
    """Admin-only — accept a MAQUINA CF .xlsx, parse the three core sheets,
    and persist the parsed JSON snapshot. Latest upload wins on /cashflow."""
    if not session.get("is_admin"):
        abort(403)
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Choose a MAQUINA CF .xlsx file first.", "error")
        return redirect(url_for("cashflow"))
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("That isn't an Excel workbook — upload the .xlsx file.", "error")
        return redirect(url_for("cashflow"))
    try:
        data = maquina_cf_parser.parse_maquina_cf(f.read(), filename=f.filename)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("cashflow"))
    except Exception as e:  # pragma: no cover
        app.logger.warning("Maquina CF parse failed: %s", e)
        flash("Could not read that workbook — is it the standard MAQUINA CF file?", "error")
        return redirect(url_for("cashflow"))
    db.save_maquina_cf(json.dumps(data), f.filename, session.get("username") or "admin")
    yrs = data.get("meta", {}).get("monthly_years") or []
    flash("MAQUINA CF imported — "
          + (f"actuals {yrs[0]}–{yrs[-1]} + projections through "
             f"{data['meta']['years'][-1]}." if yrs else "snapshot saved."), "ok")
    return redirect(url_for("cashflow"))


@app.route("/strategy")
@login_required
def strategy():
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM strategy_phases ORDER BY phase_order")
    phases = cur.fetchall()
    cur.execute("""
        SELECT c.id, c.slug, c.name, c.accent, c.industry, c.value_unit,
               c.takeover_year, c.takeover_month,
               c.hold_start_year, c.hold_start_month, c.target_hold_years,
               (c.logo IS NOT NULL) AS has_logo, octet_length(c.logo) AS logo_ver,
               r.internal_risk, r.external_risk
        FROM companies c LEFT JOIN company_risks r ON r.company_id = c.id
        WHERE c.archived = FALSE ORDER BY c.display_order
    """)
    comps = cur.fetchall()
    # cashflows per company/year (USD) — powers hold weighting + exit returns
    cur.execute("SELECT company_id, year, COALESCE(SUM(investment),0) AS inv, "
                "COALESCE(SUM(distribution),0) AS dist FROM portfolio_cashflows "
                "GROUP BY company_id, year")
    series_by_co = {}
    for r in cur.fetchall():
        if r["year"] is None:
            continue
        series_by_co.setdefault(r["company_id"], {})[r["year"]] = {
            "inv": r["inv"] or 0, "dist": r["dist"] or 0}
    invested_by_co = {cid: sum(s["inv"] for s in yrs.values())
                      for cid, yrs in series_by_co.items()}
    # full phase history per company (with month precision)
    cur.execute("""
        SELECT ph.company_id, sp.phase_order, sp.name, sp.color,
               ph.start_year, ph.start_month, ph.end_year, ph.end_month
        FROM company_phase_history ph JOIN strategy_phases sp ON sp.id = ph.phase_id
        ORDER BY ph.company_id, ph.start_year, ph.start_month, sp.phase_order
    """)
    hist = cur.fetchall()
    cur.close()
    conn.close()

    # Inputs for per-company exit valuations (Ember uses the sponsor model
    # with its live promote; everyone else is EBITDA × multiple − net debt).
    rate = usd_mxn_rate()
    _er = fetch_ember_returns()
    ember_promote = (_er["totals"]["promote"] if _er and _er.get("totals") else 0)

    # Pick each company's CURRENT phase from its timeline, by date: the
    # phase whose [start, end] window contains today (latest-starting one
    # if several do); otherwise the most recent phase that has started.
    today = datetime.date.today()
    today_mi = today.year * 12 + (today.month - 1)
    hist_by_co = {}
    for h in hist:
        hist_by_co.setdefault(h["company_id"], []).append(h)

    cur_phase, phase_window = {}, {}
    for cid, hs in hist_by_co.items():
        hs.sort(key=lambda h: (_month_index(h["start_year"], h["start_month"]) or 0))
        chosen = None
        for h in hs:
            s = _month_index(h["start_year"], h["start_month"])
            e = _month_index(h["end_year"], h["end_month"])
            if s is not None and e is not None and s <= today_mi <= e:
                chosen = h  # latest-starting window that still contains today
        if chosen is None:
            started = [h for h in hs if (_month_index(h["start_year"], h["start_month"]) or 0) <= today_mi]
            chosen = started[-1] if started else (hs[0] if hs else None)
        if chosen:
            cur_phase[cid] = chosen
            phase_window[cid] = "%s – %s" % (
                _fmt_month(chosen["start_year"], chosen["start_month"]),
                _fmt_month(chosen["end_year"], chosen["end_month"]))

    comp_list = []
    for c in comps:
        cp = cur_phase.get(c["id"])
        hold_years = _period_years(c["hold_start_year"], c["hold_start_month"])
        target = c["target_hold_years"]
        # projected exit returns — exit at current valuation, on the target date
        fin = load_financials(c["id"])
        val = company_valuation(fin, ember_promote if c["slug"] == "ember" else 0,
                                rate, c["value_unit"])
        xr = _exit_calc(series_by_co.get(c["id"], {}), c["hold_start_year"],
                        target, val["value_usd"] if val else None)
        comp_list.append(dict(
            id=c["id"], slug=c["slug"], name=c["name"], accent=c["accent"], industry=c["industry"],
            has_logo=c["has_logo"], logo_ver=c["logo_ver"],
            takeover_year=c["takeover_year"],
            internal=c["internal_risk"] or 5, external=c["external_risk"] or 5,
            phase_name=cp["name"] if cp else "—",
            phase_order=cp["phase_order"] if cp else 0,
            phase_color=cp["color"] if cp else "#8A9199",
            phase_window=phase_window.get(c["id"]),
            hold_years=hold_years,
            vintage=c["hold_start_year"],
            target_hold=target,
            over_due=(hold_years is not None and target and hold_years > target),
            proj=xr,
        ))

    # group companies under each phase
    for p in phases:
        p_companies = [c for c in comp_list if c["phase_order"] == p["phase_order"]]
        p["_companies"] = p_companies

    # ── Portfolio hold analytics ──
    # Capital-weighted average hold, vintage spread, and aging buckets
    # (capital + company count by how long each asset has been held).
    held = [c for c in comp_list if c["hold_years"] is not None]
    wnum = sum(c["hold_years"] * invested_by_co.get(c["id"], 0) for c in held)
    wden = sum(invested_by_co.get(c["id"], 0) for c in held)
    if wden:
        wavg_hold = wnum / wden
    else:                                    # no invested capital → simple mean
        wavg_hold = (sum(c["hold_years"] for c in held) / len(held)) if held else None
    avg_target = ([c["target_hold"] for c in comp_list if c["target_hold"]]
                  and sum(c["target_hold"] for c in comp_list if c["target_hold"])
                  / len([c for c in comp_list if c["target_hold"]])) or None

    vintages = {}
    for c in held:
        if c["vintage"]:
            vintages.setdefault(c["vintage"], []).append(c["name"])
    vintage_list = [dict(year=y, names=vintages[y], n=len(vintages[y])) for y in sorted(vintages)]

    BUCKETS = [("<2 yrs", 0, 2), ("2–4 yrs", 2, 4), ("4–6 yrs", 4, 6), ("6+ yrs", 6, 999)]
    aging = []
    for label, lo, hi in BUCKETS:
        members = [c for c in held if lo <= c["hold_years"] < hi]
        aging.append(dict(label=label, n=len(members),
                          capital=sum(invested_by_co.get(c["id"], 0) for c in members)))
    total_invested = sum(invested_by_co.get(c["id"], 0) for c in comp_list)
    hold_analytics = dict(wavg_hold=wavg_hold, avg_target=avg_target,
                          vintages=vintage_list, aging=aging,
                          total_invested=total_invested, n_held=len(held))

    # ── Portfolio projected exit returns (blended MOIC + capital-weighted IRR) ──
    xrs = [c["proj"] for c in comp_list if c.get("proj")]
    p_inv = sum(x["invested"] for x in xrs)
    p_proceeds = sum(x["dist"] + x["exit_value"] for x in xrs)
    iw = [(x["irr"], x["invested"]) for x in xrs if x["irr"] is not None and x["invested"]]
    p_irr = (sum(r * w for r, w in iw) / sum(w for _, w in iw)) if iw else None
    proj = dict(moic=(p_proceeds / p_inv if p_inv else None), irr=p_irr,
                proceeds=p_proceeds, invested=p_inv, n=len(xrs))

    return render_template("strategy.html", phases=phases, comps=comp_list,
                           hold=hold_analytics, proj=proj, base_year=BASE_YEAR)


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
        # `raw` keeps the full report (incl. the `monthly` block) so the budget
        # revenue overlay can work month-by-month, exactly like EmberApps does.
        return {"years": years, "rows": rows, "totals": totals, "kpis": kpis,
                "as_of": d.get("date"), "raw": d}
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember operations fetch failed: %s", e)
        return None


def fetch_ember_view(name):
    """Read a finished view payload published by EmberApps
    (reports.report_type='view:<name>').

    This is the preferred source for anything EmberApps derives: the payload is
    the exact object Ember's own page renders, so consuming it makes drift
    impossible. Callers fall back to Maquina's local logic when it returns None
    (view not published yet, or Ember unreachable). See the cross-app view
    contract — never re-implement a calculation Ember already performs."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        ecur.execute("SELECT data, uploaded_at FROM reports WHERE report_type = %s "
                     "ORDER BY uploaded_at DESC LIMIT 1", ("view:" + name,))
        row = ecur.fetchone()
        ecur.close()
        econn.close()
        if not row or not row.get("data"):
            return None
        d = row["data"]
        if isinstance(d, str):
            d = json.loads(d)
        if not d:
            return None
        return d, (row["uploaded_at"] if row.get("uploaded_at") else None)
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember view:%s fetch failed: %s", name, e)
        return None


def fetch_ember_budget():
    """Latest Ember Operating Budget (firm P&L forecast) from the Ember DB
    (report_type='ember_budget', uploaded on EmberApps' /budget page). Returns
    the parsed dict (revenue/people/operations/net_income/cash_flow + kpis) or
    None when unavailable, so callers fall back to the overhead proxy. Raw USD."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        ecur.execute("SELECT data FROM reports WHERE report_type = 'ember_budget' "
                     "ORDER BY uploaded_at DESC LIMIT 1")
        row = ecur.fetchone()
        ecur.close(); econn.close()
        if not row or not row.get("data"):
            return None
        d = row["data"]
        if isinstance(d, str):
            d = json.loads(d)
        return d if (d and d.get("meta")) else None
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember budget fetch failed: %s", e)
        return None


def _ops_revenue_axis(ops):
    """From an Operating Revenues report, pull corporate revenue keyed for the
    budget axis: {by_year:{'YYYY':t}, by_month:{'YYYY-MM':t}, lines:[...]}.
    Faithful port of EmberApps' _ops_revenue_axis. None if unusable."""
    if not isinstance(ops, dict):
        return None
    yr = ops.get("yearly_rollup") or {}
    years = yr.get("years") or []
    totals = yr.get("totals") or []
    by_year = {str(y): float(t or 0) for y, t in zip(years, totals)}
    mo = ops.get("monthly") or {}
    dates = mo.get("dates") or []
    mtotals = mo.get("totals") or []
    by_month = {}
    for iso, v in zip(dates, mtotals):
        s = str(iso or "")
        if len(s) >= 7:
            by_month[s[:7]] = round(by_month.get(s[:7], 0.0) + float(v or 0), 2)
    mrows = {r.get("label"): r for r in (mo.get("rows") or []) if isinstance(r, dict)}
    lines = []
    for r in (yr.get("rows") or []):
        if not isinstance(r, dict):
            continue
        name = r.get("label")
        lby = {str(y): float(v or 0) for y, v in zip(years, r.get("values") or [])}
        lmonths = {}
        mr = mrows.get(name)
        if mr:
            for iso, v in zip(dates, (mr.get("values") or [])):
                s = str(iso or "")
                if len(s) >= 7 and v:
                    lmonths[s[:7]] = round(lmonths.get(s[:7], 0.0) + float(v or 0), 2)
        lines.append({"name": name, "by_year": lby, "months": lmonths})
    if not by_year and not by_month:
        return None
    return {"by_year": by_year, "by_month": by_month, "lines": lines}


def _apply_ops_revenue(budget, ops):
    """Replace the Operating Budget's revenue with the live Operating Revenues
    figures and re-derive Net Income (= revenue - costs), Cash Flow (shifted by
    the revenue delta) and Cumulative -- MONTHLY as well as annually, so the
    Maquina panel ties 1-for-1 with EmberApps' /budget page. Original Excel
    revenue preserved under revenue.excel. No-op if operations data is missing.
    Faithful port of EmberApps' _apply_ops_revenue."""
    if not budget:
        return budget
    rev = _ops_revenue_axis(ops)
    if not rev:
        return budget
    import copy
    b = copy.deepcopy(budget)
    meta = b.get("meta", {}) or {}
    months = meta.get("months", []) or []
    years = [str(y) for y in (meta.get("years", []) or [])]

    old_total = (b.get("revenue", {}) or {}).get("total", {}) or {}
    old_m = old_total.get("months", {}) or {}
    old_y = old_total.get("by_year", {}) or {}

    new_m = {k: rev["by_month"][k] for k in months if rev["by_month"].get(k)}
    new_y = {y: round(rev["by_year"].get(y, 0.0), 2) for y in years}
    dmonth = {k: new_m.get(k, 0.0) - float(old_m.get(k, 0.0) or 0) for k in months}
    dyear = {y: new_y.get(y, 0.0) - float(old_y.get(y, 0.0) or 0) for y in years}

    # 1) revenue (keep Excel original for reference)
    b.setdefault("revenue", {})
    b["revenue"]["excel"] = old_total
    b["revenue"]["total"] = {"months": new_m, "by_year": new_y}
    b["revenue"]["lines"] = rev["lines"]
    b["revenue"]["source"] = "operations"

    # 2) net income = revenue - total costs
    tc = b.get("total_costs", {}) or {}
    tcm = tc.get("months", {}) or {}
    tcy = tc.get("by_year", {}) or {}
    ni_m = {}
    for k in months:
        v = round(new_m.get(k, 0.0) - float(tcm.get(k, 0.0) or 0), 2)
        if v:
            ni_m[k] = v
    ni_y = {y: round(new_y.get(y, 0.0) - float(tcy.get(y, 0.0) or 0), 2) for y in years}
    b["net_income"] = {"months": ni_m, "by_year": ni_y}

    # 3) cash flow shifts by the revenue delta (costs unchanged)
    cf = b.get("cash_flow", {}) or {}
    cf_m_old = cf.get("months", {}) or {}
    cfm = dict(cf_m_old)
    cfy = dict(cf.get("by_year", {}) or {})
    for k in months:
        if dmonth.get(k):
            cfm[k] = round(float(cfm.get(k, 0.0) or 0) + dmonth[k], 2)
    for y in years:
        cfy[y] = round(float(cfy.get(y, 0.0) or 0) + dyear.get(y, 0.0), 2)
    b["cash_flow"] = {"months": cfm, "by_year": cfy}

    # 4) cumulative = beginning balance + running sum of the new cash flow
    old_cum_m = (b.get("cumulative", {}) or {}).get("months", {}) or {}
    begin = 0.0
    if months:
        f = months[0]
        begin = float(old_cum_m.get(f, 0.0) or 0) - float(cf_m_old.get(f, 0.0) or 0)
    run = begin
    cum_m, year_end = {}, {}
    for k in months:
        run = round(run + float(cfm.get(k, 0.0) or 0), 2)
        cum_m[k] = run
        year_end[k[:4]] = run
    b["cumulative"] = {"months": cum_m, "by_year": {y: year_end.get(y, 0.0) for y in years}}

    # 5) KPIs
    kp = b.get("kpis", {}) or {}
    kp["revenue_life"] = round(sum(new_y.values()), 2)
    kp["net_income_life"] = round(sum(ni_y.values()), 2)
    kp["cash_flow_life"] = round(sum(cfy.values()), 2)
    kp["revenue_source"] = "operations"
    b["kpis"] = kp
    meta["revenue_source"] = "operations"
    b["meta"] = meta
    return b


def fetch_ember_loans():
    """Latest Ember loans report → {loans:[...], totals:{...}, as_of}. Raw USD.
    Returns None on any problem (caller falls back to seeded data)."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        ecur.execute("SELECT data FROM reports WHERE report_type = 'loans' "
                     "ORDER BY uploaded_at DESC LIMIT 1")
        row = ecur.fetchone()
        ecur.close(); econn.close()
        if not row or not row.get("data"):
            return None
        d = row["data"]
        loans = []
        for grp, label in (("mpc_loans", "MPC"), ("vertical_loans", "Vertical")):
            for r in ((d.get(grp) or {}).get("rows") or []):
                if isinstance(r, dict):
                    loans.append({"group": label, **r})
        if not loans:
            return None

        def num(r, k):
            try:
                return float(r.get(k) or 0)
            except (TypeError, ValueError):
                return 0.0
        bal = sum(num(r, "Balance") for r in loans)
        # Maturity ladder (balance maturing per year) + interest-reserve runway
        maturity, reserve, soonest_ir = {}, 0.0, None
        for r in loans:
            td = r.get("Loan Term Date")
            if isinstance(td, str) and len(td) >= 4 and td[:4].isdigit():
                maturity[int(td[:4])] = maturity.get(int(td[:4]), 0.0) + num(r, "Balance")
            res, rmir = num(r, "Rem. Interest Reserve"), num(r, "Remaining Mos. of IR")
            reserve += res
            if res > 0 and rmir > 0:
                soonest_ir = rmir if soonest_ir is None else min(soonest_ir, rmir)
        totals = {
            "amount": sum(num(r, "Loan Amount") for r in loans),
            "drawn": sum(num(r, "Drawn") for r in loans),
            "balance": bal,
            "remaining": sum(num(r, "Remaining") for r in loans),
            "burn": sum(num(r, "Monthly Interest Burn") for r in loans),
            "wrate": (sum(num(r, "Today's Rate") * num(r, "Balance") for r in loans) / bal) if bal else 0.0,
            "count": len(loans),
            "maturity": sorted(maturity.items()),
            "reserve": reserve,
            "soonest_ir": soonest_ir,
        }
        return {"loans": loans, "totals": totals, "as_of": d.get("date")}
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember loans fetch failed: %s", e)
        return None


def fetch_ember_returns():
    """Latest Ember returns report → per-project LP economics + promotes.
    Returns {projects:[...], totals:{promote, lp_dist}, as_of}. $ values scaled
    to raw USD (report stores $K). None on any problem."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        ecur.execute("SELECT data FROM reports WHERE report_type = 'returns' "
                     "ORDER BY uploaded_at DESC LIMIT 1")
        row = ecur.fetchone()
        ecur.close(); econn.close()
        if not row or not row.get("data"):
            return None
        projs = row["data"].get("projects") or []
        out, tot_promote, tot_dist = [], 0.0, 0.0
        for p in projs:
            if not isinstance(p, dict):
                continue
            m = {x.get("label"): x.get("total") for x in (p.get("metrics") or []) if isinstance(x, dict)}
            promote = (m.get("Promote") or 0) * 1000.0
            lp_dist = (m.get("Total LP Distributions") or 0) * 1000.0
            tot_promote += promote
            tot_dist += lp_dist
            out.append({"name": p.get("name"), "lp_irr": m.get("LP IRR"),
                        "moic": m.get("LP Equity Multiple"), "lp_dist": lp_dist,
                        "lp_contrib": (m.get("Total LP Contributions") or 0) * 1000.0,
                        "promote": promote})
        if not out:
            return None
        return {"projects": out, "totals": {"promote": tot_promote, "lp_dist": tot_dist},
                "as_of": row["data"].get("date")}
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember returns fetch failed: %s", e)
        return None


def fetch_ember_capital():
    """Replicate the core of Ember's Capital dashboard from the returns report:
    active projects, portfolio roll-ups, and the monthly LP distributions / promotes.
    $ values scaled to raw USD (report stores $K). None on any problem."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        def _latest(rt):
            ecur.execute("SELECT data FROM reports WHERE report_type = %s "
                         "ORDER BY uploaded_at DESC LIMIT 1", (rt,))
            r = ecur.fetchone()
            return (r["data"] if r else None) or {}
        d = _latest("returns")
        pipe_blob = _latest("ember_capital_pipeline_manual")
        vis_blob = _latest("ember_capital_pipeline_visibility")
        commit_blob = _latest("ember_capital_commitments")
        ecur.close(); econn.close()
        if not d:
            return None
        K = 1000.0
        months = [str(m) for m in (d.get("months") or [])]
        today_iso = datetime.date.today().isoformat()
        cur_year = today_iso[:4]
        active, mdist, mprom = [], [0.0] * len(months), [0.0] * len(months)
        tot_eq = tot_profit = tot_promote = dist_ltd = dist_ytd = 0.0
        wnum = wden = 0.0
        for p in (d.get("projects") or []):
            if not isinstance(p, dict) or p.get("active") is False:
                continue
            name = (p.get("name") or "").strip()
            if not name:
                continue
            by = {m.get("label"): m for m in (p.get("metrics") or []) if isinstance(m, dict)}

            def t(lbl):
                v = (by.get(lbl) or {}).get("total")
                try:
                    return float(v) if v is not None else 0.0
                except (TypeError, ValueError):
                    return 0.0

            def mo(lbl):
                return (by.get(lbl) or {}).get("monthly") or []
            irr = t("LP IRR")
            irr_pct = irr * 100 if abs(irr) <= 1.5 else irr
            equity = abs(t("Total LP Contributions")) * K
            promote = t("Promote") * K
            dist_total = t("Total LP Distributions") * K
            dm, pm = mo("Total LP Distributions"), mo("Promote")
            p_ltd = dist_total
            if dm and months:
                p_ltd = sum((dm[i] or 0) for i in range(min(len(dm), len(months)))
                            if months[i] <= today_iso) * K
            active.append(dict(name=name, irr=irr_pct, em=t("LP Equity Multiple"),
                               equity=equity, profit=t("Total LP Profit") * K,
                               promote=promote, dist_ltd=p_ltd, dist_total=dist_total))
            tot_eq += equity
            tot_profit += t("Total LP Profit") * K
            tot_promote += promote
            dist_ltd += p_ltd
            if equity > 0:
                wnum += irr_pct * equity
                wden += equity
            for i in range(len(months)):
                dv = (dm[i] if i < len(dm) else 0) or 0
                pv = (pm[i] if i < len(pm) else 0) or 0
                mdist[i] += dv * K
                mprom[i] += pv * K
                if months[i][:4] == cur_year and months[i] <= today_iso:
                    dist_ytd += dv * K
        active.sort(key=lambda a: -a["equity"])
        tot_dist_ltd_local = dist_ltd

        # ── Prefer EmberApps' published capital view (cross-app view contract) ──
        # view:capital is the exact context Ember's /capital page renders, so its
        # roll-ups, per-project rows, pipeline and commitments are authoritative.
        # Units: kpis + active rows are $K; commitments are already dollars.
        # Ember does NOT publish the monthly distribution/promote series or a
        # per-project promote, so those stay locally derived (labelled in the UI).
        _cv = fetch_ember_view("capital")
        if _cv:
            v, v_asof = _cv
            K2 = 1000.0
            kp = v.get("kpis") or {}
            promote_by_name = {a["name"]: a.get("promote") for a in active}
            v_active = []
            for a in (v.get("active") or []):
                nm = a.get("name")
                v_active.append(dict(
                    name=nm, asset_class=a.get("asset_class"),
                    irr=a.get("irr") or 0, em=a.get("em") or 0,
                    equity=(a.get("equity") or 0) * K2,
                    profit=(a.get("profit") or 0) * K2,
                    dist_ltd=(a.get("to_date") or 0) * K2,
                    dist_total=(a.get("to_date") or 0) * K2,
                    promote=promote_by_name.get(nm)))
            v_pipeline = [dict(
                name=p.get("name"), location=p.get("address") or "",
                asset_class=p.get("asset_class"),
                irr=p.get("irr") or 0, em=p.get("em") or 0,
                land_price=(p.get("land_price") or 0) * K2,
                fcst_equity=(p.get("forecasted_equity") or 0) * K2,
                duration=p.get("duration"), gross_margin=p.get("gross_margin"))
                for p in (v.get("pipeline") or []) if p.get("show_in_report", True)]
            vc = v.get("commitments") or {}
            groups = vc.get("groups") or []
            vt = vc.get("totals") or {}
            ctot = {
                "mpc": vt.get("mpc") or 0, "mpc_allocated": vt.get("mpc_allocated") or 0,
                "vertical": vt.get("vertical") or 0,
                "vertical_allocated": vt.get("vertical_allocated") or 0,
                "committed": vt.get("total_committed") or 0,
                "allocated": vt.get("total_allocated") or 0,
                "available": vt.get("available") or 0}
            # Equity by asset class (mirrors Ember's capital donut)
            classes = v.get("asset_classes") or []
            eq_by = {}
            for a in v_active:
                eq_by[a.get("asset_class")] = eq_by.get(a.get("asset_class"), 0) + (a["equity"] or 0)
            tot_eq_v = sum(eq_by.values()) or 1
            by_class = [dict(id=c.get("id"), label=c.get("label"), color=c.get("color"),
                             equity=eq_by.get(c.get("id"), 0),
                             pct=round(eq_by.get(c.get("id"), 0) / tot_eq_v * 100))
                        for c in classes if eq_by.get(c.get("id"))]
            return dict(
                active=v_active, months=months, mdist=mdist, mprom=mprom,
                pipeline=v_pipeline, commitments=groups, commit_totals=ctot,
                by_class=by_class,
                as_of=v.get("as_of") or d.get("date"), src="view",
                src_asof=(v_asof.strftime("%Y-%m-%d %H:%M") if v_asof else None),
                totals=dict(
                    equity=(kp.get("total_equity") or 0) * K2,
                    profit=(kp.get("forecasted_lp_profit") or 0) * K2,
                    promote=(kp.get("forecasted_promote") or 0) * K2,
                    dist_ltd=tot_dist_ltd_local,
                    dist_ytd=(kp.get("distributed_ytd") or 0) * K2,
                    to_be_distributed=(kp.get("to_be_distributed") or 0) * K2,
                    weighted_irr=kp.get("weighted_irr") or 0,
                    lp_irr=kp.get("forecasted_lp_irr"),
                    count=kp.get("active_count") or len(v_active),
                    ytd_label=kp.get("ytd_label")))

        # ── Pipeline (manual blob, minus hidden) — contributions/distributions are raw USD ──
        hidden = set(vis_blob.get("hidden") or [])
        pipeline = []
        for mp in (pipe_blob.get("projects") or []):
            if not isinstance(mp, dict) or mp.get("id") in hidden or mp.get("name") in hidden:
                continue
            pirr = mp.get("irr") or 0
            pipeline.append(dict(
                name=mp.get("name"), location=mp.get("location"), asset_class=mp.get("asset_class"),
                irr=(pirr * 100 if abs(pirr) <= 1.5 else pirr), em=mp.get("em") or 0,
                contrib=sum(abs(float(x or 0)) for x in (mp.get("contributions_yearly") or [])),
                dist=sum(float(x or 0) for x in (mp.get("distributions_yearly") or []))))

        # ── Commitments (investor groups, raw USD) ──
        groups = [g for g in (commit_blob.get("groups") or []) if isinstance(g, dict)]
        ctot = {"mpc": 0.0, "mpc_allocated": 0.0, "vertical": 0.0, "vertical_allocated": 0.0}
        for g in groups:
            for k in ctot:
                ctot[k] += float(g.get(k) or 0)
        ctot["committed"] = ctot["mpc"] + ctot["vertical"]
        ctot["allocated"] = ctot["mpc_allocated"] + ctot["vertical_allocated"]
        ctot["available"] = ctot["committed"] - ctot["allocated"]

        return dict(active=active, months=months, mdist=mdist, mprom=mprom,
                    pipeline=pipeline, commitments=groups, commit_totals=ctot,
                    as_of=d.get("date"),
                    totals=dict(equity=tot_eq, profit=tot_profit, promote=tot_promote,
                                dist_ltd=dist_ltd, dist_ytd=dist_ytd,
                                weighted_irr=(wnum / wden if wden else 0), count=len(active)))
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember capital fetch failed: %s", e)
        return None


@app.route("/company/<slug>")
@login_required
def company(slug):
    c = company_by_slug(slug)
    if not c or c["archived"]:
        abort(404)
    items = load_company_items(c["id"])
    # Ember Units Closed / Lots Delivered aren't on the Ember dashboard yet —
    # show a "coming soon" placeholder instead of seeded figures.
    if c["slug"] == "ember":
        for it in items:
            if it["name"] in ("Units Closed", "Lots Delivered"):
                it["coming_soon"] = True
    dashboard_kpis = [it for it in items if it["is_dashboard"]]
    chart = chart_series_for(items)

    by_category = {"Commercial": [], "Operations": [], "Finance": []}
    for it in items:
        by_category.setdefault(it["category"], []).append(it)

    # Live Ember overlay (seed fallback). Operating revenues → Operations tab;
    # a cross-tab summary → Dashboard; units/lots → Commercial (coming soon).
    ember_live, ember_asof, ember_loans, ember_returns, summary = False, None, None, None, None
    cap = None        # Ember Capital (Projects tab) — Ember only
    verticals = sales = None  # Commercial tab — Ember only
    ember_budget = None  # Operating Budget (firm P&L) — Ember only, from Ember DB
    comp = None       # Compartición de Ingresos scenario model — Meta only
    if c["slug"] == "meta":
        comp = dict(META_COMP, rate=usd_mxn_rate(), compare=META_COMPARE, base=META_BASE)
    if c["slug"] == "ember":
        ember_loans = fetch_ember_loans()
        ember_returns = fetch_ember_returns()
        cap = fetch_ember_capital()
        verticals = fetch_ember_verticals()
        sales = fetch_ember_sales()
        op = fetch_ember_operations()
        # Operating Budget: prefer EmberApps' published view payload — it is the
        # exact object Ember's /budget page renders, so the two can't drift.
        # Fall back to reading the raw report + re-applying the overlay locally
        # only if the view hasn't been published (see cross-app view contract).
        _bv = fetch_ember_view("budget")
        if _bv:
            ember_budget, _bv_asof = _bv
            ember_budget = dict(ember_budget, _src="view", _src_asof=_bv_asof)
        else:
            ember_budget = fetch_ember_budget()
            if ember_budget and op:
                ember_budget = _apply_ops_revenue(ember_budget, op.get('raw'))
            if ember_budget:
                ember_budget["_src"] = "local"
        # Real net from the Ember Operating Budget (firm P&L) when uploaded;
        # else fall back to the Project-Personnel ×1.10 overhead proxy.
        budget_net, budget_net_label = None, None
        if ember_budget:
            niy = (ember_budget.get("net_income") or {}).get("by_year") or {}
            if niy:
                cy = str(BASE_YEAR)
                yk = cy if cy in niy else sorted(niy.keys())[0]
                try:
                    budget_net = float(niy.get(yk) or 0)
                    budget_net_label = "per Ember budget · " + str(yk)
                except (TypeError, ValueError):
                    budget_net = None
        if op and op["totals"]:
            ember_live, ember_asof = True, op["as_of"]
            yrs, rows, totals = op["years"], op["rows"], op["totals"]
            yr0 = str(yrs[0])
            first = lambda lbl: (rows.get(lbl) or [None])[0]
            if budget_net is not None:
                net_ocf, net_label = budget_net, budget_net_label
            else:
                # Overhead proxy until Ember uploads full overhead: Project Personnel + 10%.
                net_ocf = totals[0] - (rows.get("Project Personnel") or [0])[0] * 1.10
                net_label = "net of est. overhead · " + yr0
            by_category["Operations"] = [
                dict(name="Corporate Revenues", current=totals[0], unit="currency", unit_label=yr0, trend=None, in_chart=True),
                dict(name="Development Fees", current=first("Development Fees"), unit="currency", unit_label=yr0, trend=None, in_chart=False),
                dict(name="Bookkeeping Fee", current=first("Bookkeeping"), unit="currency", unit_label=yr0, trend=None, in_chart=False),
                dict(name="Net Operating Cashflow", current=net_ocf, unit="currency", unit_label=net_label, trend=None, in_chart=False),
            ]
            series = [dict(label=lbl, color_idx=i, category="Operations",
                           actual=vals, projection=[None] * len(yrs), unit="currency", unit_label="USD")
                      for i, (lbl, vals) in enumerate(rows.items())]
            series.append(dict(label="Operating Revenue", color_idx=0, category="summary",
                               actual=totals, projection=[None] * len(yrs), unit="currency", unit_label="USD"))
            chart = dict(years=yrs, base_year=yrs[-1], series=series)
            summary = {"revenue": totals[0], "net_ocf": net_ocf, "year": yr0,
                       "debt": (ember_loans["totals"]["balance"] if ember_loans else None),
                       "wrate": (ember_loans["totals"]["wrate"] if ember_loans else None)}

    # financial inputs (EBITDA/FRE) + leverage + two-model equity valuation
    fin = load_financials(c["id"])
    leverage = (fin["total_debt"] / fin["ebitda"]) if (fin and fin["ebitda"] and fin["ebitda"] > 0) else None
    valuation = company_valuation(
        fin, (ember_returns["totals"]["promote"] if ember_returns else 0),
        usd_mxn_rate(), c["value_unit"])

    # strategy tab data
    conn = db.get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM company_risks WHERE company_id = %s", (c["id"],))
    risk = cur.fetchone()
    cur.execute("SELECT * FROM company_strategies WHERE company_id = %s ORDER BY start_year", (c["id"],))
    strategies = cur.fetchall()
    cur.execute("""
        SELECT ph.id AS hid, ph.phase_id, sp.name, sp.phase_order, sp.color, sp.year_range,
               sp.team_expectation, ph.start_year, ph.start_month, ph.end_year,
               ph.end_month, ph.is_projected
        FROM company_phase_history ph JOIN strategy_phases sp ON sp.id = ph.phase_id
        WHERE ph.company_id = %s ORDER BY ph.start_year, ph.start_month, sp.phase_order
    """, (c["id"],))
    phase_hist = cur.fetchall()
    cur.execute("SELECT id, name, phase_order FROM strategy_phases ORDER BY phase_order")
    phases_all = cur.fetchall()
    cur.close()
    conn.close()

    # Projected exit returns — exit at current valuation, on the target date
    exit_value_usd = valuation["value_usd"] if valuation else None
    exitr = exit_returns(c["id"], c["hold_start_year"], c["target_hold_years"], exit_value_usd)

    # Hold / control derived metrics (institutional PE model)
    hold_years = _period_years(c["hold_start_year"], c["hold_start_month"])
    hold = dict(
        hold_start_label=_fmt_month(c["hold_start_year"], c["hold_start_month"]),
        takeover_label=_fmt_month(c["takeover_year"], c["takeover_month"]),
        hold_years=hold_years,
        control_years=_period_years(c["takeover_year"], c["takeover_month"]),
        vintage=c["hold_start_year"],
        target_hold=c["target_hold_years"],
        exit_label=_exit_label(c["hold_start_year"], c["hold_start_month"], c["target_hold_years"]),
        over_due=(hold_years is not None and c["target_hold_years"]
                  and hold_years > c["target_hold_years"]),
    )

    return render_template(
        "company.html",
        c=c, items=items, dashboard_kpis=dashboard_kpis, by_category=by_category,
        chart=chart, risk=risk, strategies=strategies,
        phase_hist=phase_hist, phases_all=phases_all, base_year=BASE_YEAR,
        years=list(range(HIST_START, PROJ_END + 1)),
        ember_live=ember_live, ember_asof=ember_asof, ember_loans=ember_loans,
        ember_returns=ember_returns, summary=summary, fin=fin, leverage=leverage,
        valuation=valuation, cap=cap, hold=hold, exitr=exitr,
        verticals=verticals, sales=sales, ember_budget=ember_budget, comp=comp,
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


def _parse_month(val, default_month):
    """Parse an <input type=month> value ('YYYY-MM') → (year, month).
    Accepts a bare 'YYYY' too, falling back to default_month."""
    if not val:
        return None, None
    s = str(val).strip()
    if "-" in s:
        y, _, m = s.partition("-")
        return _to_int(y), (_to_int(m) or default_month)
    return _to_int(s), default_month


def _month_index(year, month):
    """Comparable integer for a (year, month) pair; None if no year."""
    return (int(year) * 12 + (int(month or 1) - 1)) if year else None


def _fmt_month(year, month=1):
    """'Dec 2025' style label for a (year, month) pair."""
    if not year:
        return "—"
    try:
        return "%s %d" % (calendar.month_abbr[int(month or 1)], int(year))
    except (TypeError, ValueError, IndexError):
        return str(year)


def _period_years(year, month):
    """Elapsed years (decimal) from a (year, month) anchor to today.
    None if no anchor year. Used for hold period / years under control."""
    mi = _month_index(year, month)
    if mi is None:
        return None
    today = datetime.date.today()
    return max(0.0, (today.year * 12 + (today.month - 1) - mi) / 12.0)


def _exit_label(year, month, hold_years):
    """Expected-exit month label = anchor (year, month) + hold_years."""
    mi = _month_index(year, month)
    if mi is None or hold_years in (None, ""):
        return None
    try:
        ex = mi + int(round(float(hold_years) * 12))
    except (TypeError, ValueError):
        return None
    return _fmt_month(ex // 12, ex % 12 + 1)


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


@app.route("/company/<slug>/hold", methods=["POST"])
@login_required
def update_hold(slug):
    """Edit a company's investment hold-start date, operational takeover
    date (both month-precise), and target hold horizon (years)."""
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form
    hy, hm = _parse_month(f.get("hold_start"), 1)
    ty, tm = _parse_month(f.get("takeover"), 1)
    target = f.get("target_hold_years")
    try:
        target = float(target) if target not in (None, "") else None
    except (TypeError, ValueError):
        target = None
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(
        "UPDATE companies SET hold_start_year=%s, hold_start_month=%s, "
        "takeover_year=%s, takeover_month=%s, target_hold_years=%s WHERE id=%s",
        (hy, hm, ty, tm, target, c["id"]))
    conn.commit(); cur.close(); conn.close()
    flash("Hold & control dates updated.", "ok")
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
    sy, sm = _parse_month(f.get("start"), 1)
    ey, em = _parse_month(f.get("end"), 12)
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(
        "INSERT INTO company_phase_history "
        "(company_id,phase_id,start_year,start_month,end_year,end_month,is_projected) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s)",
        (c["id"], pid, sy, sm, ey, em, f.get("is_projected") == "1"))
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
        sy, sm = _parse_month(f.get("start"), 1)
        ey, em = _parse_month(f.get("end"), 12)
        q = ("UPDATE company_phase_history SET phase_id=%s, start_year=%s, start_month=%s, "
             "end_year=%s, end_month=%s, is_projected=%s WHERE id=%s AND company_id=%s")
        args = (_to_int(f.get("phase_id")), sy, sm, ey, em,
                f.get("is_projected") == "1", hid, c["id"])
        msg = "Lifecycle phase updated."
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(q, args)
    conn.commit(); cur.close(); conn.close()
    flash(msg, "ok")
    return redirect(url_for("company", slug=slug, tab="strategy"))


@app.route("/company/<slug>/financials", methods=["POST"])
@login_required
def update_financials(slug):
    if not session.get("is_admin"):
        abort(403)
    c = _company_or_404(slug)
    f = request.form

    def num(name, d=0.0):
        try:
            return float(f.get(name))
        except (TypeError, ValueError):
            return d
    model = "sponsor" if f.get("valuation_model") == "sponsor" else "ebitda"
    conn = db.get_db(); cur = conn.cursor()
    cur.execute(
        """INSERT INTO company_financials
             (company_id, ebitda, ebitda_margin, ebitda_multiple, total_debt, fre, fre_multiple, carry_discount, valuation_model)
           VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
           ON CONFLICT (company_id) DO UPDATE SET
             ebitda=EXCLUDED.ebitda, ebitda_margin=EXCLUDED.ebitda_margin,
             ebitda_multiple=EXCLUDED.ebitda_multiple, total_debt=EXCLUDED.total_debt,
             fre=EXCLUDED.fre, fre_multiple=EXCLUDED.fre_multiple,
             carry_discount=EXCLUDED.carry_discount, valuation_model=EXCLUDED.valuation_model""",
        (c["id"], num("ebitda"), num("ebitda_margin"), num("ebitda_multiple"), num("total_debt"),
         num("fre"), num("fre_multiple"), num("carry_discount"), model))
    conn.commit(); cur.close(); conn.close()
    flash("Financial inputs updated.", "ok")
    return redirect(url_for("company", slug=slug, tab="operations"))


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


def fetch_ember_verticals():
    """Ember 'Vertical' communities from the live Ember DB (read-only):
      • LightHaven — BTR/MF leasing: vd_units occupancy + rents by floorplan
      • The Hawthorne — condo sales: vd_hawthorne_data units
    Mirrors Ember's own occupancy definition. Returns None on any problem."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()

        # ── LightHaven leasing ──
        ecur.execute("SELECT status, COUNT(*) AS n FROM vd_units "
                     "WHERE vertical = 'lighthaven' GROUP BY status")
        status = {(r["status"] or "—"): r["n"] for r in ecur.fetchall()}
        total = sum(status.values())
        occ = status.get("Tenant Occupied", 0) + status.get("Renewal", 0)
        leased = occ + status.get("Leased", 0) + status.get("Model", 0)  # Ember's definition
        ecur.execute("SELECT data FROM vd_rents_data WHERE vertical='lighthaven' "
                     "ORDER BY created_at DESC LIMIT 1")
        rr = ecur.fetchone()
        rents, rent_avg = [], None
        if rr and isinstance(rr["data"], dict):
            for fp, v in (rr["data"].get("byFloorplan") or {}).items():
                if str(fp).startswith("*") or not isinstance(v, dict):
                    continue
                if str(fp).upper() == "TOTAL":
                    rent_avg = v.get("avg")
                else:
                    rents.append(dict(fp=fp, avg=v.get("avg"), min=v.get("min"), max=v.get("max")))
        lighthaven = dict(
            total=total, occupied=occ, leased=leased, status=status, rents=rents, rent_avg=rent_avg,
            occupied_pct=(occ / total * 100 if total else None),
            leased_pct=(leased / total * 100 if total else None)) if total else None

        # ── Hawthorne condo sales ──
        ecur.execute("SELECT data FROM vd_hawthorne_data WHERE vertical='hawthorne' "
                     "ORDER BY created_at DESC LIMIT 1")
        hr = ecur.fetchone()
        hawthorne = None
        if hr and isinstance(hr["data"], dict):
            units = [u for u in (hr["data"].get("units") or []) if isinstance(u, dict)]
            by_status = {}
            sold_val = sold_ppsf = 0.0
            sold_n = 0
            for u in units:
                st = u.get("status") or "—"
                by_status[st] = by_status.get(st, 0) + 1
                if str(st).upper() == "SOLD":
                    sold_n += 1
                    sold_val += float(u.get("purchasePrice") or u.get("listPrice") or 0)
                    sold_ppsf += float(u.get("ppsf") or 0)
            hawthorne = dict(
                total=len(units), by_status=by_status, sold_n=sold_n, sold_value=sold_val,
                avg_price=(sold_val / sold_n if sold_n else None),
                avg_ppsf=(sold_ppsf / sold_n if sold_n else None),
                units=sorted(units, key=lambda u: (str(u.get("status")), str(u.get("title")))))
        ecur.close()
        econn.close()
        if not lighthaven and not hawthorne:
            return None
        return dict(lighthaven=lighthaven, hawthorne=hawthorne)
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember verticals fetch failed: %s", e)
        return None


def fetch_ember_sales():
    """Community sales (MPCs) payload mirrored into the Ember DB by the Ember
    app (reports.report_type='sales'). Returns None until the snapshot syncs."""
    try:
        econn = db.get_ember_db()
        if not econn:
            return None
        ecur = econn.cursor()
        ecur.execute("SELECT data, uploaded_at FROM reports WHERE report_type='sales' "
                     "ORDER BY uploaded_at DESC LIMIT 1")
        row = ecur.fetchone()
        ecur.close()
        econn.close()
        if not row or not row.get("data"):
            return None
        d = row["data"]
        if isinstance(d, str):
            d = json.loads(d)
        comms = d.get("communities") or {}
        tg = d.get("targets") or {}
        order = ["gpd", "hld", "wrg"] + [k for k in comms if k not in ("gpd", "hld", "wrg")]
        out = []
        for key in order:
            c = comms.get(key)
            if not isinstance(c, dict):
                continue
            t_ann = tg.get(key + "_annual")
            t_mon = tg.get(key + "_monthly")
            pace = c.get("ytd_pace")
            ytd = c.get("ytd_net")
            out.append(dict(
                key=key, name=c.get("name") or key.upper(),
                gross_total=c.get("gross_total"), total_net=c.get("total_net"),
                ytd_net=ytd, ytd_pace=pace,
                pace_prev=c.get("pace_prev"), pace_prev2=c.get("pace_prev2"),
                canc_ytd=c.get("canc_ytd"), canc_total=c.get("canc_total"),
                avg_price=c.get("avg_price"), earliest=c.get("earliest"),
                target_annual=t_ann, target_monthly=t_mon,
                # target vs actual
                pace_pct=(round(pace / t_mon * 100) if (pace and t_mon) else None),
                ytd_pct=(round(ytd / t_ann * 100) if (ytd and t_ann) else None),
                on_track=(pace is not None and t_mon and pace >= t_mon * 0.8)))
        if not out:
            return None

        # Attention flags — mirrors EmberApps' exec-summary rules exactly:
        # pace < 80% of monthly target, or YTD cancels > 30% of YTD nets.
        flags = []
        for c in out:
            if c["ytd_pace"] is not None and c["target_monthly"]:
                ratio = c["ytd_pace"] / c["target_monthly"]
                if ratio < 0.8:
                    flags.append("%s YTD pace is %d%% of target (%.1f/mo vs %d/mo)."
                                 % (c["name"], round(ratio * 100), c["ytd_pace"], c["target_monthly"]))
            if c["canc_ytd"] and c["ytd_net"] and c["canc_ytd"] > c["ytd_net"] * 0.3:
                flags.append("%s cancellations are heavy (%d YTD vs %d net)."
                             % (c["name"], c["canc_ytd"], c["ytd_net"]))

        combined = dict(
            ytd_net=sum((c["ytd_net"] or 0) for c in out),
            total_net=sum((c["total_net"] or 0) for c in out),
            canc_ytd=sum((c["canc_ytd"] or 0) for c in out),
            target_annual=tg.get("combined_annual"),
            target_monthly=tg.get("combined_monthly"))
        combined["ytd_pct"] = (round(combined["ytd_net"] / combined["target_annual"] * 100)
                               if combined["target_annual"] else None)
        return dict(communities=out, combined=combined, flags=flags, targets=tg,
                    generated_at=d.get("generated_at"),
                    as_of=(row["uploaded_at"].strftime("%Y-%m-%d") if row.get("uploaded_at") else None))
    except Exception as e:  # pragma: no cover
        app.logger.warning("Ember sales fetch failed: %s", e)
        return None


def ember_diagnostics():
    """Read-only probe of the Ember Postgres for the admin diagnostic card.
    Never writes. Returns connection status, available report types, and the
    shape of the latest 'operations' report so we can map KPIs precisely."""
    info = {"configured": bool(os.environ.get("EMBER_DATABASE_URL", "").strip()),
            "connected": False, "error": None, "report_types": [], "views": [], "fin_raw": None}
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

        # Published EmberApps view payloads (cross-app view contract) — which
        # views are flowing, and how fresh. Surfaced on Settings so a stalled
        # bridge is visible instead of silently falling back to local logic.
        try:
            ecur.execute("SELECT report_type, uploaded_at FROM reports "
                         "WHERE report_type LIKE 'view:%' ORDER BY report_type")
            info["views"] = [(r["report_type"].split(":", 1)[1],
                              r["uploaded_at"].strftime("%Y-%m-%d %H:%M") if r.get("uploaded_at") else "—")
                             for r in ecur.fetchall()]
        except Exception as e:
            info["views_error"] = str(e)[:160]

        # ── TEMP: finance detail probe ──
        try:
            out = {}
            ecur.execute("SELECT data FROM reports WHERE report_type='bva_finance' ORDER BY uploaded_at DESC LIMIT 1")
            r=ecur.fetchone(); dd=(r or {}).get("data") or {}
            if isinstance(dd,str): dd=json.loads(dd)
            ents=dd.get("entities") or {}
            k0=next(iter(ents),None)
            e0=ents.get(k0) or {}
            out["bva_finance_entity"]={"entity":k0,"keys":sorted(e0.keys())[:20]}
            for kk in ("revenueBySection","summary","totals","costs"):
                if kk in e0:
                    v=e0[kk]
                    out["bva_"+kk]= (v[:1] if isinstance(v,list) else (sorted(v.keys())[:16] if isinstance(v,dict) else str(v)[:120]))
            ecur.execute("SELECT data FROM reports WHERE report_type='loans' ORDER BY uploaded_at DESC LIMIT 1")
            r=ecur.fetchone(); ld=(r or {}).get("data") or {}
            if isinstance(ld,str): ld=json.loads(ld)
            out["mpc_loans_headers"]=(ld.get("mpc_loans") or {}).get("headers")
            out["mpc_loans_totals"]=(ld.get("mpc_loans") or {}).get("totals")
            out["mpc_loans_row0"]=((ld.get("mpc_loans") or {}).get("rows") or [None])[0]
            out["vertical_loans_headers"]=(ld.get("vertical_loans") or {}).get("headers")
            out["vertical_loans_totals"]=(ld.get("vertical_loans") or {}).get("totals")
            ds=(ld.get("debt_schedules") or [{}])[0]
            out["debt_schedule"]={"project":ds.get("project"),"months":len(ds.get("months") or []),
                                  "payment_total":ds.get("payment_total"),"total_revenues":ds.get("total_revenues")}
            info["fin_raw"]=json.dumps(out, default=str, ensure_ascii=False)[:5200]
        except Exception as e:
            info["fin_raw"]="probe error: "+str(e)[:200]

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
