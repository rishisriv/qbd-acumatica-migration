# QBD Contractor → Acumatica Migration Agent

An open-source, automated ETL tool for migrating a construction contractor's
financial data from **QuickBooks Desktop** to **Acumatica ERP** — no manual
CSV juggling, no third-party migration service required.

Built for the construction industry. Tested on a real contractor migration.

---

## What It Does

```
QuickBooks Desktop  ──►  migration agent  ──►  Acumatica ERP
   (COM SDK)                                   (REST API)
```

**Extracts** from QBD via the QBXMLRP2 COM SDK (no QB plugin purchase needed):
- Chart of Accounts
- Customers + Jobs (converted to Acumatica Projects)
- Vendors (with 1099 flag)
- Items (Service, Inventory, Non-Inventory)
- Employees
- Open AR Invoices
- Open AP Bills
- Trial Balance

**Transforms** the data:
- Maps QB account types to Acumatica account types
- Maps QB payment terms to Acumatica term IDs
- Converts QB jobs to Acumatica projects (linked to parent customer)
- Maps QB job statuses to project statuses
- Auto-generates Acumatica Customer IDs from names
- Validates and warns on missing fields before loading

**Loads** into Acumatica via REST API in the correct dependency order, with
retry logic and detailed error logging.

---

## Quick Start

### Prerequisites
- Windows machine (required for QuickBooks COM SDK)
- Python 3.9+
- QuickBooks Desktop installed and company file open
- Acumatica instance configured and accessible

### Install

```bash
git clone https://github.com/rishisriv/qbd-acumatica-migration
cd qbd-acumatica-migration
pip install -r requirements.txt
cp .env.example .env
# edit .env with your Acumatica URL, company name, and credentials
```

### Run

```bash
# Step 1: Extract data from QuickBooks Desktop
python qbd_auto_extract.py

# Step 2: Dry run — validate without writing to Acumatica
python migration_agent_new.py --from-files --dry-run

# Step 3: Review migration_workspace/migration.log for warnings
# Step 4: Live migration
python migration_agent_new.py --from-files
```

---

## Files

| File | Purpose |
|---|---|
| `qbd_auto_extract.py` | Connects to QBD via COM and exports data to Excel |
| `migration_agent_new.py` | Main 4-phase ETL agent (extract → transform → load → validate) |
| `MIGRATION_RECIPE.md` | Full data mapping tables, run modes, troubleshooting |
| `PREFLIGHT_CHECKLIST.md` | Everything to configure before running |
| `.env.example` | Template for your Acumatica credentials |
| `requirements.txt` | Python dependencies |

---

## Architecture

```
Phase 1 — Extract
  ├── QBD COM SDK (QBXMLRP2)  ← primary
  └── Excel files (--from-files)  ← fallback / manual export

Phase 2 — Transform & Validate
  ├── Type mapping (QB account types → Acumatica types)
  ├── Customer ID generation
  ├── Job → Project conversion
  └── Data validation warnings

Phase 3 — Load (Acumatica REST API)
  ├── Dependency-ordered writes
  ├── Retry with exponential backoff (3 attempts)
  └── 422/404/500 error logging (non-blocking)

Phase 4 — Validate
  └── Query Acumatica and compare counts against source
```

---

## Configuration

Copy `.env.example` to `.env` and fill in your values:

```
ACUMATICA_URL=http://your-server/AcumaticaERP
ACUMATICA_COMPANY=YourCompanyName
ACUMATICA_USERNAME=admin
ACUMATICA_PASSWORD=yourpassword
ACUMATICA_ENDPOINT_VER=20.200.001
```

If you have customers already created in Acumatica, populate `EXISTING_CUSTOMER_IDS`
in `migration_agent_new.py` to reuse their existing IDs.

---

## Known Limitations

These items are not fully automated and require manual steps in Acumatica:

- **Opening balance journal** — GL batch structure varies by instance; create manually via GL301000
- **Employee records** — require Employee Class and Department pre-configured
- **AP Bills** — require AP Preferences fully configured first
- **AR Invoice cash discount date** — not present in QBD export; enter manually or set a default

See [PREFLIGHT_CHECKLIST.md](PREFLIGHT_CHECKLIST.md) for the complete list.

---

## Tested On

- QuickBooks Desktop Pro/Premier 2019–2023 (US edition)
- Acumatica 2020 R2 (endpoint version 20.200.001)
- Construction contractor with ~100 customers, 100 vendors, 100 GL accounts

---

## Contributing

This is a work in progress. The core migration (GL accounts, customers, jobs → projects, vendors) works. We're continuously improving entity coverage and would love community input, especially on:

- Acumatica endpoint version compatibility (2021+, 2022+)
- Additional QB edition support (Enterprise, Mac)
- Employee migration (currently blocked by required Acumatica defaults)
- AP Bills and AR Invoice automation (currently requires manual Acumatica pre-configuration)
- Opening balance journal automation

PRs and issue reports from consultants running real migrations are especially welcome.

---

## License

MIT — use freely, attribution appreciated.
