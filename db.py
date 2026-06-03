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
