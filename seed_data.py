"""Seed plausible fake data for the Maquina dashboard.

Numbers are reconciled to the Replit handoff screenshots:
  investment totals (USD M): Ember 20.1 · Ranman 17.9 · IPS 3.9 · Mezcal 2.0
  industry split: Real Estate 38.0 · Energy 3.9 · Consumer 2.0
Everything else is generated on smooth, deterministic curves so the demo
looks alive without random noise changing between runs.

Replace this module's data (or wire EMBER_DATABASE_URL in db.py) as real
figures come online.
"""
import os
import random
from werkzeug.security import generate_password_hash

BASE_YEAR = 2025          # series "current" anchor + actual/projection split
HIST_START = 2020
PROJ_END = 2035

# ── Lifecycle phases (portfolio-wide) ──────────────────────────────
PHASES = [
    dict(name="Startup / Takeover", phase_order=1, year_range="Years 0–2",
         team_expectation="Compliance", color="#C98A00",
         description="Stabilize controls, reporting and governance immediately after takeover."),
    dict(name="Stabilization", phase_order=2, year_range="Years 2–5",
         team_expectation="Commitment", color="#0568B3",
         description="Lock in operating discipline and predictable cashflow before scaling."),
    dict(name="Growth", phase_order=3, year_range="Years 5+",
         team_expectation="Commitment", color="#1F9D6B",
         description="Compound the platform — expand capacity, markets and distributions."),
]


def _kpi(name, category, current, cagr=0.10, unit="number", unit_label="",
         dashboard=False, in_chart=True, integer=False, type="kpi"):
    return dict(name=name, category=category, current=current, cagr=cagr, unit=unit,
                unit_label=unit_label, dashboard=dashboard, in_chart=in_chart,
                integer=integer, type=type)


# ── Authoritative tab taxonomy ─────────────────────────────────────
# Commercial = sales & project performance; Operations = company financial
# health (revenue sources, margins, expenses, net operating cashflow);
# Finance = lines of credit / debt. Overrides the per-KPI category above and
# is also applied to the live DB via remap_categories().
CATEGORY_BY_COMPANY = {
    "ember": {"Corporate Cashflow": "Operations", "Corporate Revenues": "Operations",
              "Bookkeeping Fee": "Operations", "Development Fees": "Operations",
              "G&A Expense": "Operations", "Units Closed": "Commercial", "Lots Delivered": "Commercial"},
    "ranman": {"Apartados": "Commercial", "Firmas": "Commercial", "Escrituras": "Commercial",
               "Inventario de Lotes": "Commercial", "Margen UAIR %": "Operations",
               "UAIR": "Operations", "Ingresos": "Operations"},
    "ips": {"Inversión": "Operations", "Inversión IPS": "Operations",
            "Inversión Ranman Energy": "Operations", "Acumulado Inversión": "Operations",
            "Capacidad Instalada": "Operations", "Generación": "Commercial", "Clientes": "Commercial"},
    "mezcal-local": {"Sell-in Mexico": "Commercial", "Sell-in USA": "Commercial",
                     "Puntos de Venta": "Commercial", "Margen Bruto": "Operations",
                     "UAIR": "Operations", "Producción": "Operations"},
}


def _category_for(slug, name, default):
    return CATEGORY_BY_COMPANY.get(slug, {}).get(name, default)


def remap_categories(conn):
    """Sync existing unified_items.category in the live DB to the taxonomy above."""
    cur = conn.cursor()
    for slug, mapping in CATEGORY_BY_COMPANY.items():
        for name, cat in mapping.items():
            cur.execute(
                "UPDATE unified_items SET category = %s WHERE name = %s AND "
                "company_id = (SELECT id FROM companies WHERE slug = %s)",
                (cat, name, slug))
    conn.commit()
    cur.close()


COMPANIES = [
    # ── EMBER ─────────────────────────────────────────────────────
    dict(
        slug="ember", name="Ember", industry="Real Estate",
        country="United States", country_code="US", currency="USD",
        value_unit="KUSD", accent="#E0701F", takeover_year=2023, display_order=1,
        description="US master-planned community developer — Houston metro.",
        inv_total=20.1, dist_total=31.5, inv_start=2023, inv_years=3, dist_start=2024,
        current_phase=2, internal_risk=3.5, external_risk=4.0,
        risk=dict(market_risk=4, team_risk=3, finance_risk=3, product_risk=4),
        strategies=[
            dict(name="Bookkeeping & Fee Platform Scale", approach="better",
                 start_year=2024, end_year=2027, formulation_rating=8, execution_rating=7),
            dict(name="In-house Vertical Development", approach="different",
                 start_year=2025, end_year=2030, formulation_rating=6, execution_rating=5),
        ],
        kpis=[
            _kpi("Corporate Cashflow", "Finance", 1004, 0.14, "currency", "KUSD", dashboard=True),
            _kpi("Corporate Revenues", "Finance", 2360, 0.12, "currency", "KUSD", dashboard=True),
            _kpi("Bookkeeping Fee", "Commercial", 420, 0.10, "currency", "KUSD", dashboard=True),
            _kpi("Development Fees", "Commercial", 880, 0.16, "currency", "KUSD", dashboard=True),
            _kpi("G&A Expense", "Finance", 1180, 0.08, "currency", "KUSD", in_chart=False),
            _kpi("Units Closed", "Operations", 312, 0.11, "count", "units", integer=True),
            _kpi("Lots Delivered", "Operations", 540, 0.09, "count", "lots", integer=True),
        ],
    ),
    # ── RANMAN ────────────────────────────────────────────────────
    dict(
        slug="ranman", name="Ranman", industry="Real Estate",
        country="Mexico", country_code="MX", currency="MXN",
        value_unit="MMXN", accent="#0568B3", takeover_year=2020, display_order=2,
        description="Mexican residential land development — lot sales & master plans.",
        inv_total=17.9, dist_total=67.4, inv_start=2020, inv_years=4, dist_start=2022,
        current_phase=3, internal_risk=4.0, external_risk=6.0,
        risk=dict(market_risk=6, team_risk=4, finance_risk=5, product_risk=4),
        strategies=[
            dict(name="Lot Absorption Acceleration", approach="better",
                 start_year=2022, end_year=2026, formulation_rating=8, execution_rating=8),
            dict(name="Cancún Master Plan", approach="different",
                 start_year=2024, end_year=2031, formulation_rating=7, execution_rating=6),
        ],
        kpis=[
            _kpi("Apartados", "Commercial", 571, 0.08, "count", "apartados", dashboard=True, integer=True),
            _kpi("Firmas", "Commercial", 380, 0.09, "count", "firmas", dashboard=True, integer=True),
            _kpi("Margen UAIR %", "Finance", 2.0, 0.05, "percent", "%", dashboard=True, in_chart=False),
            _kpi("UAIR", "Finance", 13.0, 0.15, "currency", "MMXN", dashboard=True),
            _kpi("Ingresos", "Finance", 640, 0.12, "currency", "MMXN"),
            _kpi("Escrituras", "Operations", 295, 0.10, "count", "escrituras", integer=True),
            _kpi("Inventario de Lotes", "Operations", 1240, -0.04, "count", "lotes", integer=True, in_chart=False),
        ],
    ),
    # ── IPS ───────────────────────────────────────────────────────
    dict(
        slug="ips", name="IPS", industry="Energy",
        country="Mexico", country_code="MX", currency="MXN",
        value_unit="MMXN", accent="#1F9D6B", takeover_year=2024, display_order=3,
        description="Energy infrastructure — generation & distribution in Mexico.",
        inv_total=3.9, dist_total=0.0, inv_start=2024, inv_years=3, dist_start=2028,
        current_phase=1, internal_risk=6.5, external_risk=5.0,
        risk=dict(market_risk=5, team_risk=6, finance_risk=7, product_risk=6),
        strategies=[
            dict(name="Solar Capacity Buildout", approach="different",
                 start_year=2024, end_year=2030, formulation_rating=6, execution_rating=5),
            dict(name="Ranman Energy Integration", approach="better",
                 start_year=2025, end_year=2028, formulation_rating=7, execution_rating=6),
        ],
        kpis=[
            _kpi("Inversión", "Finance", 67.6, 0.18, "currency", "MMXN", dashboard=True),
            _kpi("Inversión IPS", "Finance", 41.0, 0.20, "currency", "MMXN", dashboard=True),
            # Series intentionally hidden from the chart per the handoff (card only).
            _kpi("Inversión Ranman Energy", "Finance", 26.6, 0.16, "currency", "MMXN", dashboard=True, in_chart=False),
            _kpi("Acumulado Inversión", "Finance", 67.6, 0.18, "currency", "MMXN", dashboard=True, in_chart=False),
            _kpi("Capacidad Instalada", "Operations", 12.5, 0.22, "number", "MW"),
            _kpi("Generación", "Operations", 38.0, 0.20, "number", "GWh"),
            _kpi("Clientes", "Commercial", 18, 0.25, "count", "clientes", integer=True),
        ],
    ),
    # ── MEZCAL LOCAL ──────────────────────────────────────────────
    dict(
        slug="mezcal-local", name="Mezcal Local", industry="Consumer Products",
        country="Mexico", country_code="MX", currency="MXN",
        value_unit="KMXN", accent="#B07A2E", takeover_year=2024, display_order=4,
        description="Premium mezcal brand — Mexico & US distribution.",
        inv_total=2.0, dist_total=0.2, inv_start=2024, inv_years=3, dist_start=2027,
        current_phase=1, internal_risk=7.0, external_risk=6.5,
        risk=dict(market_risk=7, team_risk=6, finance_risk=7, product_risk=5),
        strategies=[
            dict(name="US Distribution Network", approach="different",
                 start_year=2025, end_year=2031, formulation_rating=6, execution_rating=4),
            dict(name="Premium SKU Margin Mix", approach="better",
                 start_year=2025, end_year=2029, formulation_rating=7, execution_rating=6),
        ],
        kpis=[
            _kpi("Sell-in Mexico", "Commercial", 8600, 0.14, "count", "Cajas", dashboard=True, integer=True),
            _kpi("Sell-in USA", "Commercial", 3100, 0.35, "count", "Cajas", dashboard=True, integer=True),
            _kpi("Margen Bruto", "Finance", 11800, 0.18, "currency", "KMXN", dashboard=True),
            _kpi("UAIR", "Finance", -2400, 0.45, "currency", "KMXN", dashboard=True),
            _kpi("Puntos de Venta", "Operations", 240, 0.20, "count", "PdV", integer=True),
            _kpi("Producción", "Operations", 96000, 0.16, "number", "Litros", integer=True, in_chart=False),
        ],
    ),
]


def _kpi_series(current, cagr, seed, integer=False, allow_negative=False):
    r = random.Random(seed)
    actual, proj = [], []
    for y in range(HIST_START, PROJ_END + 1):
        base = current * ((1 + cagr) ** (y - BASE_YEAR))
        # Anchor the current-year value exactly so headline KPIs match the
        # reference figures; only history/projection years get the ±5% wiggle.
        v = current if y == BASE_YEAR else base * (1 + (r.random() * 2 - 1) * 0.05)
        if not allow_negative:
            v = max(0, v)
        if integer:
            v = round(v)
        else:
            v = round(v, 2)
        if y <= BASE_YEAR:
            actual.append((y, v))
        if y >= BASE_YEAR:
            proj.append((y, v))
    return actual, proj


def _cashflow_series(inv_total, dist_total, inv_start, inv_years, dist_start):
    """Front-loaded investment, ramping distributions. USD millions in, raw USD out."""
    inv = {}
    weights = [max(1, inv_years - i) for i in range(inv_years)]
    wsum = sum(weights)
    for i in range(inv_years):
        inv[inv_start + i] = inv_total * 1e6 * weights[i] / wsum
    dist = {}
    dyears = list(range(dist_start, PROJ_END + 1))
    if dist_total > 0 and dyears:
        dweights = [i + 1 for i in range(len(dyears))]
        dwsum = sum(dweights)
        for i, y in enumerate(dyears):
            dist[y] = dist_total * 1e6 * dweights[i] / dwsum
    rows = []
    for y in range(HIST_START, PROJ_END + 1):
        iv, dv = inv.get(y, 0.0), dist.get(y, 0.0)
        if iv == 0 and dv == 0:
            continue
        ds = "actual" if y <= BASE_YEAR else "projection"
        rows.append((y, round(iv, 2), round(dv, 2), ds))
    return rows


def seed(conn):
    cur = conn.cursor()

    # ── admin user (only if no users yet) ──
    cur.execute("SELECT COUNT(*) AS n FROM users")
    if cur.fetchone()["n"] == 0:
        uname = os.environ.get("ADMIN_USERNAME", "admin")
        pw = os.environ.get("ADMIN_PASSWORD", "maquina2026")
        cur.execute(
            "INSERT INTO users (username, password_hash, is_admin, first_name) "
            "VALUES (%s, %s, TRUE, %s)",
            (uname, generate_password_hash(pw), "Admin"),
        )

    # ── phases ──
    phase_ids = {}
    for p in PHASES:
        cur.execute(
            "INSERT INTO strategy_phases (name, description, phase_order, year_range, team_expectation, color) "
            "VALUES (%(name)s,%(description)s,%(phase_order)s,%(year_range)s,%(team_expectation)s,%(color)s) "
            "RETURNING id", p,
        )
        phase_ids[p["phase_order"]] = cur.fetchone()["id"]

    # ── companies + everything hanging off them ──
    for c in COMPANIES:
        cur.execute(
            """INSERT INTO companies
               (slug,name,industry,country,country_code,currency,value_unit,accent,
                description,takeover_year,display_order)
               VALUES (%(slug)s,%(name)s,%(industry)s,%(country)s,%(country_code)s,%(currency)s,
                       %(value_unit)s,%(accent)s,%(description)s,%(takeover_year)s,%(display_order)s)
               RETURNING id""", c,
        )
        cid = cur.fetchone()["id"]

        # risk
        rk = c["risk"]
        cur.execute(
            "INSERT INTO company_risks (company_id,market_risk,team_risk,finance_risk,product_risk,internal_risk,external_risk) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s)",
            (cid, rk["market_risk"], rk["team_risk"], rk["finance_risk"], rk["product_risk"],
             c["internal_risk"], c["external_risk"]),
        )

        # phase history: prior phases (projected=false) up to current
        cur_phase = c["current_phase"]
        tko = c["takeover_year"]
        span = {1: (tko, tko + 2), 2: (tko + 2, tko + 5), 3: (tko + 5, PROJ_END)}
        for po in range(1, cur_phase + 1):
            s, e = span[po]
            cur.execute(
                "INSERT INTO company_phase_history (company_id,phase_id,start_year,end_year,is_projected) "
                "VALUES (%s,%s,%s,%s,%s)",
                (cid, phase_ids[po], s, e, po > cur_phase),
            )

        # strategies
        for st in c["strategies"]:
            cur.execute(
                "INSERT INTO company_strategies (company_id,name,approach,start_year,end_year,formulation_rating,execution_rating) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (cid, st["name"], st["approach"], st["start_year"], st["end_year"],
                 st["formulation_rating"], st["execution_rating"]),
            )

        # cashflows
        for (y, iv, dv, ds) in _cashflow_series(
            c["inv_total"], c["dist_total"], c["inv_start"], c["inv_years"], c["dist_start"]
        ):
            cur.execute(
                "INSERT INTO portfolio_cashflows (company_id,year,investment,distribution,dataset_type) "
                "VALUES (%s,%s,%s,%s,%s)", (cid, y, iv, dv, ds),
            )

        # KPI items + values
        for order, k in enumerate(c["kpis"]):
            cur.execute(
                """INSERT INTO unified_items
                   (company_id,name,type,category,unit,unit_label,is_dashboard,in_chart,display_order)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (cid, k["name"], k["type"], _category_for(c["slug"], k["name"], k["category"]),
                 k["unit"], k["unit_label"], k["dashboard"], k["in_chart"], order),
            )
            iid = cur.fetchone()["id"]
            allow_neg = k["current"] < 0
            actual, proj = _kpi_series(
                k["current"], k["cagr"], seed=f"{c['slug']}:{k['name']}",
                integer=k["integer"], allow_negative=allow_neg,
            )
            for (y, v) in actual:
                cur.execute(
                    "INSERT INTO unified_values (item_id,year,value,dataset_type) VALUES (%s,%s,%s,'actual')",
                    (iid, y, v),
                )
            for (y, v) in proj:
                cur.execute(
                    "INSERT INTO unified_values (item_id,year,value,dataset_type) VALUES (%s,%s,%s,'projection')",
                    (iid, y, v),
                )

    # ── settings ──
    rate = os.environ.get("USD_MXN_RATE", "17.3328")
    cur.execute(
        "INSERT INTO app_settings (key,value) VALUES ('usd_mxn_rate',%s) "
        "ON CONFLICT (key) DO NOTHING", (rate,),
    )

    conn.commit()
    cur.close()
