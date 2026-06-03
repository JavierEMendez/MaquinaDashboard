# Maquina — Portfolio Management Platform

Internal dashboard for **Maquina Holdings**, tracking financial performance,
KPIs, investments, and strategic lifecycle across the portfolio companies:
**Ember, Ranman, IPS, and Mezcal Local**.

Built to match the **Ember dashboard's** look and feel — a Flask + Postgres app
with session auth, a permanent sidebar, and a Maquina-branded theme (charcoal +
brand blue `#0568B3`). Deploys on Railway via Docker, exactly like Ember.

> Data is **seeded with realistic fake figures** today. Real numbers get wired
> in over time — Ember can read directly from the existing Ember dashboard's
> database (see *Ember data* below).

---

## Stack

| Layer | Tech |
|---|---|
| Backend | Python 3.12 · Flask · Gunicorn |
| Database | PostgreSQL (`psycopg2`) |
| Auth | Server sessions · `werkzeug` password hashing |
| Frontend | Server-rendered Jinja · Chart.js · vanilla JS |
| Fonts | Space Grotesk (display) · Inter (body) |
| Hosting | Railway (Dockerfile builder) |

## Pages

| Route | Page |
|---|---|
| `/` | Portfolio Dashboard — investment summary by currency, donuts by country/industry |
| `/maquina-cashflow` | Maquina Portfolio — Portfolio + Planning tabs (investments, distributions, ROI, timeline) |
| `/strategy` | Strategy — lifecycle phase framework + risk quadrant |
| `/company/<slug>` | Company detail — Dashboard / Commercial / Operations / Finance / Strategy / Planning tabs |
| `/manage-companies` | Company list + archive |
| `/settings` | Exchange rate, appearance, team |
| `/login`, `/logout` | Auth |

## Project structure

```
app.py            Flask app — routes, auth, data-shaping helpers, Jinja filters
db.py             Connection + idempotent schema (the unified data model)
seed_data.py      Fake data for the four companies (companies, KPIs, cashflows, strategy)
templates/        Jinja templates (base shell, sidebar, login, pages)
static/           theme.css, processed Maquina logos, favicon
Dockerfile        Railway build
```

---

## Local development

Requires a PostgreSQL database (local or a Railway public connection string).

```bash
pip install -r requirements.txt

# configure environment
cp .env.example .env          # then edit DATABASE_URL + SECRET_KEY
export DATABASE_URL=postgresql://...     # PowerShell: $env:DATABASE_URL="..."
export SECRET_KEY=dev-secret

python app.py                 # http://localhost:8080
```

Schema creation + fake-data seeding run automatically on the first request.

Default login (first boot, override with `ADMIN_USERNAME` / `ADMIN_PASSWORD`):

```
username: admin
password: maquina2026
```

---

## Deploy to Railway

1. Push this repo to GitHub (already wired to `JavierEMendez/MaquinaDashboard`).
2. In the Railway project, create a service from the repo. The `railway.toml`
   selects the **Dockerfile** builder.
3. Add a **PostgreSQL** plugin to the project.
4. Set service variables:

| Variable | Value |
|---|---|
| `DATABASE_URL` | `${{Postgres.DATABASE_URL}}` (reference the plugin) |
| `SECRET_KEY` | a long random string |
| `ADMIN_USERNAME` | initial admin (optional, default `admin`) |
| `ADMIN_PASSWORD` | initial admin password (optional, default `maquina2026`) |
| `USD_MXN_RATE` | fallback FX, e.g. `17.3328` |
| `BANXICO_TOKEN` | *(optional)* enables live MXN/USD from Banxico |
| `EMBER_DATABASE_URL` | *(optional)* read-only Ember Postgres — see below |

Railway provides `$PORT`; `gunicorn.conf.py` binds to it.

---

## Ember data

`EMBER_DATABASE_URL` (in `db.py → get_ember_db()`) is a read-only hook to the
**existing Ember dashboard's** Postgres. Left unset, Ember shows seeded figures.
Once pointed at the live database, Ember's KPIs can be pulled from real data
while the other companies stay on the seeded model until their numbers arrive.

## The data model

Follows the handoff's "unified" model so it can grow with real data:

- `companies` — portfolio company metadata
- `unified_items` / `unified_values` — KPI definitions and their values by year (actual / projection)
- `portfolio_cashflows` — per-company investment & distribution by year (USD)
- `strategy_phases` / `company_phase_history` — lifecycle framework
- `company_risks` / `company_strategies` — strategy tab data
- `users`, `app_settings`

Replace the contents of `seed_data.py` (or wire live sources) as real figures come online.
