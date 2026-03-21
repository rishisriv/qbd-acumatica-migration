# Migration Recipe — QuickBooks Desktop Contractor → Acumatica

This document describes exactly what the agent migrates, how it maps the data,
and what you need to configure before running it.

---

## What Gets Migrated

| QBD Entity | Acumatica Entity | Notes |
|---|---|---|
| Chart of Accounts | Account | Type mapping below |
| Customers | Customer | With billing address and contact |
| Customer Jobs | Project | One project per QB job; linked to parent customer |
| Vendors | Vendor | Including 1099 flag |
| Items | InventoryItem / Non-Inventory | Service items → Non-Inventory |
| Employees | Employee | Requires Employee Class pre-configured |
| Open AR Invoices | ARInvoice | Unpaid only |
| Open AP Bills | APBill | Unpaid only |
| Trial Balance | JournalTransaction | Opening balance GL entry |
| Credit Terms | Terms | Mapped to Acumatica term IDs |

---

## Account Type Mapping

| QuickBooks Type | Acumatica Type |
|---|---|
| Bank | Cash |
| Accounts Receivable | AccountsReceivable |
| Other Current Asset | Asset |
| Fixed Asset | FixedAsset |
| Other Asset | Asset |
| Accounts Payable | AccountsPayable |
| Credit Card | Liability |
| Other Current Liability | Liability |
| Long Term Liability | Liability |
| Equity | Equity |
| Income | Income |
| Cost of Goods Sold | Expense |
| Expense | Expense |
| Other Income | Income |
| Other Expense | Expense |

---

## Credit Terms Mapping

| QuickBooks Term | Acumatica Term ID |
|---|---|
| Due on receipt | DUERECEIPT |
| Net 10 | NET10 |
| Net 15 | NET15 |
| Net 30 | NET30 |
| Net 60 | NET60 |

To add more terms, edit the `QBD_TERMS_MAP` dictionary in `migration_agent_new.py`.

---

## Job Status Mapping (QB Jobs → Acumatica Projects)

| QuickBooks Job Status | Acumatica Project Status |
|---|---|
| In progress | Active |
| Awarded | Active |
| Pending | Planned |
| Closed | Completed |
| (none / unknown) | Planned |

---

## Customer ID Generation

The agent generates Acumatica Customer IDs automatically from the QB name:

- **`Last, First` format** → `LASTF` (e.g., `Smith, John` → `SMITHJ`)
- **Company name** → first 10 characters, uppercased, spaces removed (e.g., `Acme Corp` → `ACMECORP`)

If you have customers already created in Acumatica before this migration, populate
the `EXISTING_CUSTOMER_IDS` dict in the `DataTransformer` class to reuse their IDs
instead of generating new ones.

---

## How the Agent Loads Acumatica

The agent uses the Acumatica REST API (`/entity/{ENDPOINT_NAME}/{ENDPOINT_VER}/`).
Each entity is loaded in dependency order:

```
1. Credit Terms          (AR depends on these)
2. Chart of Accounts     (all modules depend on GL accounts)
3. Vendors               (AP depends on vendors)
4. Customers             (AR, Projects depend on customers)
5. Projects              (converted from QB jobs)
6. Employees
7. Opening Balance Journal  (GL entry)
8. Open AR Invoices
9. Open AP Bills
```

Each PUT/POST retries up to 3 times with backoff on 5xx errors.
422 validation errors are logged and skipped (no retry — fix the data, re-run).

---

## Run Modes

```bash
# Full run: extract from QBD then load into Acumatica
python migration_agent_new.py

# Use existing Excel exports (QBD not required to be running)
python migration_agent_new.py --from-files

# Dry run: validate everything, make no writes to Acumatica
python migration_agent_new.py --dry-run

# Most common for first pass: validate from files
python migration_agent_new.py --from-files --dry-run
```

---

## Excel Export Format

When using `--from-files`, the agent looks for these files in `./qbd_exports/`:

| File | Required Columns |
|---|---|
| `QBD_ChartOfAccounts.xlsx` | Name, FullName, AccountType, Balance |
| `QBD_Customers.xlsx` | Name, FullName, IsActive, Balance, BillAddressAddr1, BillAddressCity, BillAddressState, BillAddressPostalCode, Phone, Email, Terms, ParentRef |
| `QBD_Vendors.xlsx` | Name, IsActive, Balance, VendorAddressAddr1, Phone, Email, Is1099 |
| `QBD_Items.xlsx` | Name, FullName, Type, SalesDesc, PurchaseDesc |
| `QBD_Employees.xlsx` | Name, FirstName, LastName, Email, Phone |
| `QBD_OpenInvoices.xlsx` | TxnDate, DueDate, RefNumber, CustomerRef, BalanceRemaining, ARAccountRef |
| `QBD_OpenBills.xlsx` | TxnDate, DueDate, RefNumber, VendorRef, BalanceRemaining, APAccountRef |

These files are generated automatically by `qbd_auto_extract.py`.
You can also create them manually from a QB report export.

---

## Output Files

After the migration runs:

| File | Contents |
|---|---|
| `migration_workspace/migration.log` | Full execution log with timestamps |
| `migration_workspace/migration_report.json` | Structured JSON summary: counts, errors, warnings per phase |
| `qbd_exports/extract.log` | QBD extraction log |

---

## Troubleshooting Common Errors

| Error | Cause | Fix |
|---|---|---|
| `422 Employee Class is required` | Employee Class not configured in Acumatica | Create at least one Employee Class, then re-run |
| `422 Cash Discount Date is required` | AR Invoice missing field not in QBD | Set a default in the AR Preferences or enter manually |
| `404 Terms entity not found` | Endpoint name mismatch by Acumatica version | Try `CreditTerms` instead of `Terms` in the API call |
| `500 AP Bills` | AP Preferences not configured | Configure Accounts Payable Preferences first |
| `QBD COM connection failed` | QuickBooks not running or blocking COM | Open QBD, switch to single-user mode, try again |
| `openpyxl error reading file` | Corrupt or password-protected XLSX | Re-export from QB or use the auto-extractor |
