# Pre-Migration Checklist — QBD Contractor → Acumatica

Work through every item below **before** running the migration agent.
Items are ordered: QuickBooks first, Acumatica setup second, then validation.

---

## 1. QuickBooks Desktop (source)

- [ ] QuickBooks Desktop is installed and **open** on the machine where you will run the agent
- [ ] The correct company file is loaded (not sample data)
- [ ] You are logged in as a user with **full admin rights** in QB
- [ ] No other user is connected to the company file via multi-user mode that could block COM access
- [ ] QuickBooks Integrated Applications are **enabled** (`Edit > Preferences > Integrated Applications > Company Preferences`)
- [ ] Run a Trial Balance report in QB and note the total equity figure — you will reconcile this after migration
- [ ] Note the QB fiscal year-end date — you will need it when configuring the Acumatica ledger

---

## 2. Acumatica (target)

### Instance
- [ ] Acumatica is installed and accessible at the URL in your `.env` file
- [ ] The company / tenant is created and the name matches `ACUMATICA_COMPANY` in `.env`
- [ ] REST API endpoint is enabled (`System > Web Services > Web Service Endpoints`) and the version matches `ACUMATICA_ENDPOINT_VER`
- [ ] Login credentials in `.env` are correct — test by logging in to the Acumatica UI

### Required module configuration (must exist before migrating)
- [ ] **General Ledger** — GL Preferences configured, fiscal year created, posting period open
- [ ] **Accounts Payable** — AP Preferences configured (vendor classes, payment methods)
- [ ] **Accounts Receivable** — AR Preferences configured (customer classes, payment methods)
- [ ] **Project Accounting** — Projects module enabled if migrating jobs as projects

### Required lookups (migration will fail without these)
- [ ] **Employee Class** — at least one Employee Class exists (`Payroll > Configuration > Employee Classes`)
- [ ] **Department** — at least one Department exists
- [ ] **Sales Account** — default Sales Account set on Employee Class
- [ ] **Customer Class** — default Customer Class configured
- [ ] **Vendor Class** — default Vendor Class configured
- [ ] **Credit Terms** — common terms created (Due on Receipt, Net 10, Net 15, Net 30, Net 60) under `Finance > Accounts Receivable > Configuration > Credit Terms`

### Pre-migration backup
- [ ] **Full database backup** of the Acumatica SQL database taken and stored off-machine
- [ ] If re-running the migration, verify the target company is in a clean state (or restore the backup)

---

## 3. The Migration Machine

- [ ] Python 3.9 or later installed (`python --version`)
- [ ] Dependencies installed: `pip install -r requirements.txt`
- [ ] `.env` file created from `.env.example` and filled in
- [ ] Running on **Windows** (required for the QuickBooks COM SDK / pywin32)
- [ ] Machine is on the same network as the Acumatica server

---

## 4. Data Readiness

- [ ] Chart of Accounts reviewed — remove or merge duplicate accounts before extraction
- [ ] Customers reviewed — resolve any duplicate names in QB
- [ ] Vendors reviewed — 1099 vendor flags are set correctly in QB
- [ ] Open invoices (AR) — confirm all open invoices are accurate; fully-paid invoices should be marked paid
- [ ] Open bills (AP) — confirm all open bills are accurate
- [ ] Items list reviewed — decide which items map to Acumatica Inventory vs. Non-Inventory vs. Service

---

## 5. Run Order

Run the scripts in this order:

```
1. python qbd_auto_extract.py          # Pulls data from QBD → ./qbd_exports/
2. python migration_agent_new.py --from-files --dry-run   # Validates without writing
3. Review migration_workspace/migration.log for WARNINGs
4. python migration_agent_new.py --from-files             # Live load into Acumatica
5. Review migration_workspace/migration_report.json
6. Reconcile GL balances manually (see Post-Migration Checklist)
```

---

## 6. Known Limitations / Manual Steps

These items are **not automated** and must be done manually in Acumatica after the migration:

| Item | Reason | Where in Acumatica |
|---|---|---|
| Opening Balance Journal | GL batch structure varies by instance configuration | Finance > General Ledger > Journal Transactions |
| Employee HR records | Requires Employee Class + Department pre-configured | Payroll > Employees |
| Cash Discount Date on AR invoices | Field not present in QBD export | Finance > Accounts Receivable > Invoices |
| Payment terms endpoint | Varies by Acumatica version (Terms vs. CreditTerms) | Finance > AR > Configuration > Credit Terms |
| AP module bills | Requires AP Preferences fully configured first | Finance > Accounts Payable > Bills |

---

## 7. Post-Migration Reconciliation

- [ ] Total customer count in Acumatica matches QB
- [ ] Total vendor count in Acumatica matches QB
- [ ] Chart of Accounts count matches QB
- [ ] Trial Balance total equity in Acumatica matches QB snapshot taken in step 1.6
- [ ] At least 3 customers spot-checked (address, phone, terms, open balance)
- [ ] At least 3 projects (converted from QB jobs) spot-checked (status, description, customer link)
- [ ] No duplicate customer IDs in Acumatica
