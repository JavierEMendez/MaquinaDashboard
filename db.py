"""Database access + schema for the Maquina dashboard.

Mirrors the Ember dashboard's pattern: a thin psycopg2 layer over a
Railway-provided PostgreSQL, with an idempotent schema created on first
request. The data model follows the Replit handoff's "unified" model
(companies + unified_items + unified_values) plus portfolio cashflows
and strategy tables.
"""
import os
import psycopg2
import psycopg2.extras


def get_db():
    """Primary connection — this app's own database (users, companies, KPIs)."""
    return psycopg2.connect(
        os.environ["DATABASE_URL"],
        cursor_factory=psycopg2.extras.RealDictCursor,
    )


def get_ember_db():
    """Read-only connection to the EXISTING Ember dashboard's Postgres.

    Returns None when EMBER_DATABASE_URL is unset, so callers fall back to
    seeded fake data for Ember. Wire real Ember figures here later.
    """
    url = os.environ.get("EMBER_DATABASE_URL", "").strip()
    if not url:
        return None
    return psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    is_admin BOOLEAN DEFAULT FALSE,
    first_name TEXT,
    last_name TEXT,
    email TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS companies (
    id SERIAL PRIMARY KEY,
    slug TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    industry TEXT,                 -- Real Estate | Energy | Consumer Products
    country TEXT,                  -- United States | Mexico
    country_code TEXT,             -- US | MX (flag)
    currency TEXT,                 -- USD | MXN
    value_unit TEXT,               -- KUSD | MMXN | KMXN | units (dashboard hint)
    accent TEXT,                   -- hex used for charts + avatar
    description TEXT,
    takeover_year INT,
    display_order INT DEFAULT 0,
    archived BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS unified_items (
    id SERIAL PRIMARY KEY,
    company_id INT REFERENCES companies(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    type TEXT NOT NULL DEFAULT 'kpi',       -- kpi | debt | investment
    category TEXT,                          -- Commercial | Operations | Finance
    unit TEXT DEFAULT 'number',             -- currency | count | percent | ratio | number
    unit_label TEXT,                        -- KUSD | Cajas | % ...
    parent_id INT REFERENCES unified_items(id) ON DELETE CASCADE,
    is_manual BOOLEAN DEFAULT FALSE,
    is_dashboard BOOLEAN DEFAULT FALSE,     -- surfaced as a KPI card on the company dashboard
    in_chart BOOLEAN DEFAULT TRUE,          -- include as a series in the performance chart
    display_order INT DEFAULT 0
);
CREATE INDEX IF NOT EXISTS unified_items_company_idx ON unified_items(company_id);

CREATE TABLE IF NOT EXISTS unified_values (
    id SERIAL PRIMARY KEY,
    item_id INT REFERENCES unified_items(id) ON DELETE CASCADE,
    year INT NOT NULL,
    value DOUBLE PRECISION,
    dataset_type TEXT NOT NULL DEFAULT 'actual',  -- actual | projection
    UNIQUE (item_id, year, dataset_type)
);
CREATE INDEX IF NOT EXISTS unified_values_item_idx ON unified_values(item_id);

CREATE TABLE IF NOT EXISTS strategy_phases (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    description TEXT,
    phase_order INT,
    year_range TEXT,
    team_expectation TEXT,
    color TEXT
);

CREATE TABLE IF NOT EXISTS company_phase_history (
    id SERIAL PRIMARY KEY,
    company_id INT REFERENCES companies(id) ON DELETE CASCADE,
    phase_id INT REFERENCES strategy_phases(id) ON DELETE CASCADE,
    start_year INT,
    end_year INT,
    is_projected BOOLEAN DEFAULT FALSE
);

CREATE TABLE IF NOT EXISTS company_risks (
    id SERIAL PRIMARY KEY,
    company_id INT UNIQUE REFERENCES companies(id) ON DELETE CASCADE,
    market_risk INT DEFAULT 5,
    team_risk INT DEFAULT 5,
    finance_risk INT DEFAULT 5,
    product_risk INT DEFAULT 5,
    internal_risk DOUBLE PRECISION DEFAULT 5,   -- quadrant Y
    external_risk DOUBLE PRECISION DEFAULT 5    -- quadrant X
);

CREATE TABLE IF NOT EXISTS company_strategies (
    id SERIAL PRIMARY KEY,
    company_id INT REFERENCES companies(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    approach TEXT,                  -- better | different
    start_year INT,
    end_year INT,
    formulation_rating INT,
    execution_rating INT
);

CREATE TABLE IF NOT EXISTS portfolio_cashflows (
    id SERIAL PRIMARY KEY,
    company_id INT REFERENCES companies(id) ON DELETE CASCADE,
    year INT NOT NULL,
    investment DOUBLE PRECISION DEFAULT 0,    -- USD
    distribution DOUBLE PRECISION DEFAULT 0,  -- USD
    dataset_type TEXT NOT NULL DEFAULT 'actual',
    UNIQUE (company_id, year, dataset_type)
);

CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

-- Editable per-company financial / valuation inputs (Phases 3–4).
-- Values are in each company's display units; multiples/rates are unitless.
CREATE TABLE IF NOT EXISTS company_financials (
    company_id INT UNIQUE REFERENCES companies(id) ON DELETE CASCADE,
    ebitda DOUBLE PRECISION DEFAULT 0,
    ebitda_margin DOUBLE PRECISION DEFAULT 0,
    ebitda_multiple DOUBLE PRECISION DEFAULT 7,
    total_debt DOUBLE PRECISION DEFAULT 0,
    fre DOUBLE PRECISION DEFAULT 0,
    fre_multiple DOUBLE PRECISION DEFAULT 10,
    carry_discount DOUBLE PRECISION DEFAULT 0.15,
    valuation_model TEXT DEFAULT 'ebitda',   -- ebitda | sponsor
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Per-user activity ledger powering the admin Team Activity page.
-- Captures successful logins, logouts, and page views (user-level only —
-- no IPs, no user agents). Username is denormalized so the page keeps
-- working even if a user is later renamed or deleted. Retention is
-- 12 months, enforced by a once-per-boot purge in app.py's _boot().
CREATE TABLE IF NOT EXISTS activity_log (
    id BIGSERIAL PRIMARY KEY,
    user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
    username TEXT NOT NULL,
    event_type TEXT NOT NULL,           -- login | logout | page_view
    path TEXT,
    ts TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS activity_log_user_ts ON activity_log (user_id, ts DESC);
CREATE INDEX IF NOT EXISTS activity_log_ts ON activity_log (ts DESC);

-- Uploaded images live in Postgres (Railway's filesystem is ephemeral).
-- Stored as small processed thumbnails; served via dedicated routes.
ALTER TABLE users     ADD COLUMN IF NOT EXISTS avatar BYTEA;
ALTER TABLE users     ADD COLUMN IF NOT EXISTS avatar_mime TEXT;
ALTER TABLE companies ADD COLUMN IF NOT EXISTS logo BYTEA;
ALTER TABLE companies ADD COLUMN IF NOT EXISTS logo_mime TEXT;
ALTER TABLE company_risks ADD COLUMN IF NOT EXISTS commentary TEXT;

-- Maquina's ownership of each portfolio company (%). Drives the "Maquina
-- share" shown beside each company's equity valuation. Defaults to 100 for
-- wholly-owned companies.
ALTER TABLE companies ADD COLUMN IF NOT EXISTS maquina_pct DOUBLE PRECISION DEFAULT 100;

-- Risk matrix: each category is plotted as Likelihood (x) x Impact (y).
-- The existing *_risk columns are the IMPACT axis; these add LIKELIHOOD so
-- every category can be positioned (and dragged) independently.
ALTER TABLE company_risks ADD COLUMN IF NOT EXISTS market_likelihood  INT DEFAULT 5;
ALTER TABLE company_risks ADD COLUMN IF NOT EXISTS team_likelihood    INT DEFAULT 5;
ALTER TABLE company_risks ADD COLUMN IF NOT EXISTS finance_likelihood INT DEFAULT 5;
ALTER TABLE company_risks ADD COLUMN IF NOT EXISTS product_likelihood INT DEFAULT 5;

-- Month precision on lifecycle phases (e.g. Dec 2025 → Dec 2027).
-- Defaults keep legacy year-only rows as full-year spans (Jan–Dec).
ALTER TABLE company_phase_history ADD COLUMN IF NOT EXISTS start_month INT DEFAULT 1;
ALTER TABLE company_phase_history ADD COLUMN IF NOT EXISTS end_month   INT DEFAULT 12;

-- Investment hold + operational control (institutional PE model). We store
-- the anchor dates and DERIVE periods (never store "years held" — it goes
-- stale). hold_start_* = entry/investment date → hold period + vintage;
-- takeover_* = operational-control date → lifecycle phase clock + years
-- under control; target_hold_years = planned horizon → hold-vs-plan + exit ETA.
ALTER TABLE companies ADD COLUMN IF NOT EXISTS hold_start_year   INT;
ALTER TABLE companies ADD COLUMN IF NOT EXISTS hold_start_month  INT DEFAULT 1;
ALTER TABLE companies ADD COLUMN IF NOT EXISTS takeover_month    INT DEFAULT 1;
ALTER TABLE companies ADD COLUMN IF NOT EXISTS target_hold_years DOUBLE PRECISION;

-- Maquina Cashflow page — each row is one uploaded "MAQUINA CF" workbook,
-- parsed into JSON (maquina_cf_parser). Latest row wins on the dashboard;
-- history is kept so prior snapshots aren't lost. A stopgap until Ember /
-- Ranman / etc. figures flow in live.
CREATE TABLE IF NOT EXISTS maquina_cf_uploads (
    id SERIAL PRIMARY KEY,
    data JSONB NOT NULL,
    filename TEXT,
    uploaded_by TEXT,
    uploaded_at TIMESTAMP DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS maquina_cf_uploads_ts ON maquina_cf_uploads (uploaded_at DESC);
"""


def init_schema():
    conn = get_db()
    cur = conn.cursor()
    cur.execute(SCHEMA)
    conn.commit()
    cur.close()
    conn.close()


def is_seeded():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) AS n FROM companies")
    n = cur.fetchone()["n"]
    cur.close()
    conn.close()
    return n > 0


def latest_maquina_cf():
    """Return the most recent uploaded Maquina CF snapshot as a dict
    {data, filename, uploaded_by, uploaded_at}, or None if none exist."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "SELECT data, filename, uploaded_by, uploaded_at "
        "FROM maquina_cf_uploads ORDER BY uploaded_at DESC, id DESC LIMIT 1"
    )
    row = cur.fetchone()
    cur.close()
    conn.close()
    return dict(row) if row else None


def save_maquina_cf(data_json: str, filename: str, uploaded_by: str):
    """Insert a parsed Maquina CF snapshot. ``data_json`` is a JSON string."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO maquina_cf_uploads (data, filename, uploaded_by) "
        "VALUES (%s, %s, %s)",
        (data_json, filename, uploaded_by),
    )
    conn.commit()
    cur.close()
    conn.close()
