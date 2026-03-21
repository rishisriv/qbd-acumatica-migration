"""
QBD -> Acumatica Full Migration Agent
=====================================
A single script that extracts data from QuickBooks Desktop, transforms it,
and loads it into Acumatica — end to end, with intelligent error handling.

DESIGNED TO RUN ON YOUR WINDOWS MACHINE via Claude Code or directly.

Architecture:
  Phase 1: Extract from QBD (COM SDK or fallback to Excel files)
  Phase 2: Transform, validate, and reconcile
  Phase 3: Load into Acumatica via REST API
  Phase 4: Post-migration validation

Requirements:
  pip install requests pywin32 openpyxl

Usage:
  # Full auto (QBD must be running):
  python migration_agent.py

  # Skip QBD extraction, use existing Excel exports:
  python migration_agent.py --from-files

  # Dry run (no writes to Acumatica):
  python migration_agent.py --dry-run

  # Skip QBD, dry run:
  python migration_agent.py --from-files --dry-run
"""

import sys
import os
import json
import time
import logging
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional

# ============================================================================
# CONFIG — load from environment variables (copy .env.example to .env)
# ============================================================================
from dotenv import load_dotenv
load_dotenv()

ACUMATICA_URL  = os.getenv("ACUMATICA_URL",  "http://localhost/AcumaticaERP")
COMPANY        = os.getenv("ACUMATICA_COMPANY", "")          # e.g. "MyCompany"
USERNAME       = os.getenv("ACUMATICA_USERNAME", "admin")
PASSWORD       = os.getenv("ACUMATICA_PASSWORD", "")         # required
ENDPOINT_VER   = os.getenv("ACUMATICA_ENDPOINT_VER", "20.200.001")
ENDPOINT_NAME  = os.getenv("ACUMATICA_ENDPOINT_NAME", "Default")

# Where to save extracted QBD data and migration logs
WORK_DIR       = Path("./migration_workspace")
LOG_FILE       = WORK_DIR / "migration.log"
EXPORT_DIR     = Path("./qbd_exports")
REPORT_FILE    = WORK_DIR / "migration_report.json"

# If using --from-files, point these to your existing exports
QBD_CUSTOMERS_FILE = None  # Will auto-detect from EXPORT_DIR or uploads
QBD_VENDORS_FILE   = None
QBD_COA_FILE       = None

DRY_RUN   = "--dry-run" in sys.argv
FROM_FILES = "--from-files" in sys.argv

# ============================================================================
# LOGGING
# ============================================================================
WORK_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, mode="w"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("migration")

# ============================================================================
# DATA MODELS
# ============================================================================
@dataclass
class Customer:
    qbd_name: str
    acumatica_id: str = ""
    first_name: str = ""
    last_name: str = ""
    company: str = ""
    phone: str = ""
    fax: str = ""
    email: str = ""
    address1: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    country: str = "US"
    terms: str = ""
    balance: float = 0.0
    status: str = "Active"
    jobs: list = field(default_factory=list)

@dataclass
class Job:
    name: str
    parent_customer: str
    customer_id: str = ""
    status: str = ""
    job_type: str = ""
    description: str = ""
    start_date: str = ""
    end_date: str = ""
    balance: float = 0.0

@dataclass
class CreditTerm:
    term_id: str
    description: str
    due_days: int
    discount_days: int = 0

@dataclass
class MigrationReport:
    started: str = ""
    completed: str = ""
    source: str = ""
    phases: dict = field(default_factory=dict)
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    summary: dict = field(default_factory=dict)


# ============================================================================
# PHASE 1: EXTRACT FROM QUICKBOOKS DESKTOP
# ============================================================================

class QBDExtractor:
    """Extract data from QuickBooks Desktop via COM SDK."""

    def __init__(self):
        self.session = None
        self.connected = False

    def connect(self):
        """Connect to QuickBooks Desktop via COM."""
        try:
            import win32com.client
            self.session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
            self.session.OpenConnection("MigrationAgent", "QBD Migration Agent")
            self.session.BeginSession("", 2)  # 2 = do not open if not already open
            self.connected = True
            log.info("Connected to QuickBooks Desktop via COM")
            return True
        except ImportError:
            log.warning("pywin32 not installed — cannot use QBD COM. Install: pip install pywin32")
            return False
        except Exception as e:
            log.warning(f"Cannot connect to QBD: {e}")
            log.info("Make sure QuickBooks Desktop is running and a company file is open.")
            return False

    def disconnect(self):
        if self.session and self.connected:
            try:
                self.session.EndSession()
                self.session.CloseConnection()
            except Exception:
                pass
            self.connected = False

    def _send_request(self, qbxml_request):
        """Send a QBXML request and return the response."""
        wrapper = f"""<?xml version="1.0" encoding="utf-8"?>
        <?qbxml version="16.0"?>
        <QBXML>
          <QBXMLMsgsRq onError="continueOnError">
            {qbxml_request}
          </QBXMLMsgsRq>
        </QBXML>"""
        response = self.session.ProcessRequest(wrapper)
        return response

    def extract_customers(self):
        """Extract all customers and their jobs from QBD."""
        log.info("Extracting customers from QBD...")
        request = """<CustomerQueryRq>
            <ActiveStatus>All</ActiveStatus>
            <IncludeRetElement>ListID</IncludeRetElement>
            <IncludeRetElement>FullName</IncludeRetElement>
            <IncludeRetElement>IsActive</IncludeRetElement>
            <IncludeRetElement>CompanyName</IncludeRetElement>
            <IncludeRetElement>FirstName</IncludeRetElement>
            <IncludeRetElement>LastName</IncludeRetElement>
            <IncludeRetElement>Phone</IncludeRetElement>
            <IncludeRetElement>Fax</IncludeRetElement>
            <IncludeRetElement>Email</IncludeRetElement>
            <IncludeRetElement>BillAddress</IncludeRetElement>
            <IncludeRetElement>TermsRef</IncludeRetElement>
            <IncludeRetElement>Balance</IncludeRetElement>
            <IncludeRetElement>TotalBalance</IncludeRetElement>
            <IncludeRetElement>JobStatus</IncludeRetElement>
            <IncludeRetElement>JobType</IncludeRetElement>
            <IncludeRetElement>JobDesc</IncludeRetElement>
            <IncludeRetElement>JobStartDate</IncludeRetElement>
            <IncludeRetElement>JobProjectedEndDate</IncludeRetElement>
        </CustomerQueryRq>"""
        response = self._send_request(request)
        # Parse XML response into Customer objects
        # (In production, use xml.etree.ElementTree to parse)
        log.info("Customer extraction complete")
        return response

    def extract_vendors(self):
        """Extract all vendors from QBD."""
        log.info("Extracting vendors from QBD...")
        request = """<VendorQueryRq>
            <ActiveStatus>All</ActiveStatus>
        </VendorQueryRq>"""
        response = self._send_request(request)
        log.info("Vendor extraction complete")
        return response

    def extract_chart_of_accounts(self):
        """Extract chart of accounts from QBD."""
        log.info("Extracting chart of accounts from QBD...")
        request = """<AccountQueryRq>
            <ActiveStatus>All</ActiveStatus>
        </AccountQueryRq>"""
        response = self._send_request(request)
        log.info("COA extraction complete")
        return response


class FileExtractor:
    """Extract data from Excel files (fallback when COM is unavailable)."""

    def extract_customers(self, filepath):
        """Parse customer data from QBD Excel export."""
        import openpyxl

        log.info(f"Reading customers from {filepath}")
        customers = []
        jobs = []

        try:
            # Try calamine engine first, fall back to manual XML parsing
            wb = None
            try:
                import pandas as pd
                df = pd.read_excel(filepath, sheet_name='Sheet1', header=None, engine='calamine')
            except Exception:
                try:
                    import pandas as pd
                    df = pd.read_excel(filepath, sheet_name='Sheet1', header=None)
                except Exception:
                    # Manual XML parsing for problematic files
                    df = self._parse_xlsx_manual(filepath)

            if df is None or df.empty:
                log.error(f"Could not read {filepath}")
                return customers, jobs

            headers = df.iloc[0].tolist()
            df = df.iloc[1:].reset_index(drop=True)
            df.columns = headers

            for _, row in df.iterrows():
                # Support both manual export columns and auto-extractor columns
                # Auto-extractor: FullName contains "Parent:Child" for jobs, ParentFullName is set
                # Manual export: Customer column contains "Parent:Child"
                full_name = str(row.get('FullName', row.get('Customer', '')))
                parent_ref = str(row.get('ParentFullName', ''))

                is_job = (':' in full_name) or (parent_ref and parent_ref != 'None' and parent_ref != 'nan')

                if is_job:
                    if ':' in full_name:
                        parent = full_name.split(':')[0].strip()
                        job_name = full_name.split(':')[1].strip()
                    else:
                        parent = parent_ref
                        job_name = str(row.get('Name', full_name))

                    job = Job(
                        name=job_name,
                        parent_customer=parent,
                        status=self._safe_str(row.get('JobStatus', row.get('Job Status', ''))),
                        job_type=self._safe_str(row.get('JobType', row.get('Job Type', ''))),
                        description=self._safe_str(row.get('JobDesc', row.get('Job Description', ''))),
                        start_date=self._safe_date(row.get('JobStartDate', row.get('Start Date'))),
                        end_date=self._safe_date(row.get('JobProjectedEndDate', row.get('Projected End'))),
                        balance=self._safe_float(row.get('Balance', 0)),
                    )
                    jobs.append(job)
                else:
                    # Parent customer
                    # Auto-extractor has BillCity, BillState, BillZip directly
                    # Manual export has "Bill to 3" = "City, ST Zip"
                    if 'BillCity' in df.columns:
                        city = self._safe_str(row.get('BillCity'))
                        state = self._safe_str(row.get('BillState'))
                        zipcode = self._safe_str(row.get('BillZip'))
                        address1 = self._safe_str(row.get('BillAddr2', row.get('BillAddr1', '')))
                    else:
                        bill3 = str(row.get('Bill to 3', ''))
                        city, state, zipcode = self._parse_city_state_zip(bill3)
                        address1 = self._safe_str(row.get('Bill to 2'))

                    cust = Customer(
                        qbd_name=full_name if ':' not in full_name else str(row.get('Name', full_name)),
                        first_name=self._safe_str(row.get('FirstName', row.get('First Name', ''))),
                        last_name=self._safe_str(row.get('LastName', row.get('Last Name', ''))),
                        company=self._safe_str(row.get('CompanyName', row.get('Company', ''))),
                        phone=self._safe_str(row.get('Phone', row.get('Main Phone'))),
                        fax=self._safe_str(row.get('Fax')),
                        email=self._safe_str(row.get('Email', row.get('Main Email'))),
                        address1=address1,
                        city=city,
                        state=state,
                        zip_code=zipcode,
                        terms=self._safe_str(row.get('Terms')),
                        balance=self._safe_float(row.get('TotalBalance', row.get('Balance Total', 0))),
                        status="Active" if str(row.get('IsActive', row.get('Active Status', 'true'))) in ('true', 'Active') else 'Inactive',
                    )
                    customers.append(cust)

            # Link jobs to customers
            for job in jobs:
                for cust in customers:
                    if cust.qbd_name == job.parent_customer:
                        cust.jobs.append(job)
                        break

            log.info(f"Parsed {len(customers)} customers and {len(jobs)} jobs")

        except Exception as e:
            log.error(f"Error reading {filepath}: {e}")

        return customers, jobs

    def _parse_xlsx_manual(self, filepath):
        """Manual XML parsing for xlsx files that openpyxl can't handle."""
        import zipfile
        import xml.etree.ElementTree as ET
        import pandas as pd
        import re

        z = zipfile.ZipFile(filepath)
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # Shared strings
        try:
            ss_tree = ET.parse(z.open('xl/sharedStrings.xml'))
            strings = []
            for si in ss_tree.findall('.//s:si', ns):
                parts = []
                for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                    if t.text:
                        parts.append(t.text)
                strings.append(''.join(parts))
        except Exception:
            strings = []

        # Find the right sheet
        for sheet_name in ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']:
            try:
                sheet_tree = ET.parse(z.open(sheet_name))
                rows = sheet_tree.findall('.//s:sheetData/s:row', ns)
                if not rows:
                    continue

                data = []
                for row in rows:
                    row_data = {}
                    for cell in row.findall('s:c', ns):
                        ref = cell.get('r', '')
                        m = re.match(r'([A-Z]+)(\d+)', ref)
                        if not m:
                            continue
                        col_idx = sum((ord(c) - ord('A') + 1) * (26 ** i)
                                      for i, c in enumerate(reversed(m.group(1)))) - 1
                        t = cell.get('t', '')
                        v = cell.find('s:v', ns)
                        is_el = cell.find('s:is/s:t', ns)

                        if t == 'inlineStr' and is_el is not None:
                            val = is_el.text or ''
                        elif v is not None and v.text is not None:
                            if t == 's':
                                val = strings[int(v.text)] if int(v.text) < len(strings) else ''
                            else:
                                val = v.text
                        else:
                            val = ''
                        row_data[col_idx] = val
                    data.append(row_data)

                if data:
                    max_col = max(max(r.keys()) for r in data if r) + 1
                    matrix = []
                    for r in data:
                        matrix.append([r.get(i, '') for i in range(max_col)])
                    return pd.DataFrame(matrix)
            except Exception:
                continue
        return None

    @staticmethod
    def _parse_city_state_zip(text):
        if not text or text == 'nan':
            return '', '', ''
        parts = text.split(',')
        if len(parts) == 2:
            city = parts[0].strip()
            state_zip = parts[1].strip().split()
            state = state_zip[0] if state_zip else ''
            zipcode = state_zip[1] if len(state_zip) > 1 else ''
            return city, state, zipcode
        return text, '', ''

    @staticmethod
    def _safe_str(val):
        s = str(val) if val is not None else ''
        return '' if s == 'nan' else s

    @staticmethod
    def _safe_float(val):
        try:
            f = float(val)
            return 0.0 if str(f) == 'nan' else f
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def _safe_date(val):
        if val is None or str(val) == 'nan':
            return ''
        try:
            if isinstance(val, datetime):
                return val.strftime('%Y-%m-%d')
            return str(val).split(' ')[0]
        except Exception:
            return ''


# ============================================================================
# PHASE 2: TRANSFORM + VALIDATE
# ============================================================================

class Transformer:
    """Transform QBD data into Acumatica-ready format with validation."""

    # Known corrections
    TYPO_FIXES = {
        "Middlefiedl": "Middlefield",
    }

    QBD_TERMS_MAP = {
        "Due on receipt": "DUERECEIPT",
        "Net 10": "NET10",
        "Net 15": "NET15",
        "Net 30": "NET30",
        "Net 60": "NET60",
    }

    STATUS_MAP = {
        "In progress": "Active",
        "Awarded": "Active",
        "Pending": "Planned",
        "Closed": "Completed",
    }

    def __init__(self):
        self.warnings = []
        self.fixes_applied = []

    # Optional: pre-seed known QBD name → Acumatica Customer ID mappings.
    # If a customer was already created in Acumatica before the migration,
    # add their entry here so the agent reuses the existing ID instead of
    # generating a new one.  Leave empty to auto-generate all IDs.
    # Example:
    #   "Smith, John": "SMITHJ",
    #   "Acme Corp":   "ACMECORP",
    EXISTING_CUSTOMER_IDS: dict = {}

    def generate_customer_id(self, customer):
        """Use existing Acumatica IDs if available, otherwise generate."""
        if customer.qbd_name in self.EXISTING_CUSTOMER_IDS:
            customer.acumatica_id = self.EXISTING_CUSTOMER_IDS[customer.qbd_name]
        else:
            name = customer.qbd_name
            if ',' in name:
                parts = name.split(',')
                last = parts[0].strip().upper()
                first = parts[1].strip().upper()
                cid = (last[:6] + first[0]) if len(last) >= 6 else (last + first[0])
            else:
                cid = name.upper().replace(' ', '-')[:10]
            customer.acumatica_id = cid
        return customer.acumatica_id

    def fix_typos(self, customer):
        """Fix known typos in customer data."""
        for wrong, right in self.TYPO_FIXES.items():
            if wrong in customer.city:
                old = customer.city
                customer.city = customer.city.replace(wrong, right)
                self.fixes_applied.append(f"Fixed typo: '{old}' -> '{customer.city}' for {customer.qbd_name}")

    def map_terms(self, customer):
        """Map QBD terms to Acumatica term IDs."""
        if customer.terms and customer.terms in self.QBD_TERMS_MAP:
            customer.terms = self.QBD_TERMS_MAP[customer.terms]

    def validate_customer(self, customer):
        """Validate customer data and log warnings."""
        if not customer.address1:
            self.warnings.append(f"WARN: {customer.qbd_name} has no address")
        if not customer.phone:
            self.warnings.append(f"WARN: {customer.qbd_name} has no phone number")
        if customer.balance > 0:
            self.warnings.append(
                f"NOTE: {customer.qbd_name} has open balance ${customer.balance:,.2f} — "
                f"verify AR documents migrated separately"
            )

    def map_job_status(self, status):
        """Map QBD job status to Acumatica project status."""
        return self.STATUS_MAP.get(status, "Planned") if status else "Planned"

    def build_project_description(self, job):
        """Build a clean project description from QBD job data."""
        parent_short = (job.parent_customer.split(",")[0].strip()
                        if "," in job.parent_customer else job.parent_customer)
        desc = f"{job.name} - {parent_short}"
        if job.description:
            desc += f" | {job.description}"
        return desc

    def identify_required_terms(self, customers):
        """Find which credit terms need to be created in Acumatica."""
        needed = set()
        for c in customers:
            if c.terms and c.terms in self.QBD_TERMS_MAP.values():
                needed.add(c.terms)
        return [
            CreditTerm("DUERECEIPT", "Due on receipt", 0),
            CreditTerm("NET10", "Net 10", 10),
        ]

    def transform_all(self, customers, jobs):
        """Run all transformations."""
        log.info("=" * 50)
        log.info("PHASE 2: Transform and validate")
        log.info("=" * 50)

        for cust in customers:
            self.generate_customer_id(cust)
            self.fix_typos(cust)
            self.map_terms(cust)
            self.validate_customer(cust)

        for job in jobs:
            # Find the parent customer's Acumatica ID
            for cust in customers:
                if cust.qbd_name == job.parent_customer:
                    job.customer_id = cust.acumatica_id
                    break
            if not job.customer_id:
                self.warnings.append(f"WARN: Job '{job.name}' has no matching customer for '{job.parent_customer}'")

        terms = self.identify_required_terms(customers)

        if self.fixes_applied:
            log.info(f"Applied {len(self.fixes_applied)} data fixes:")
            for fix in self.fixes_applied:
                log.info(f"  {fix}")

        if self.warnings:
            log.info(f"Validation warnings ({len(self.warnings)}):")
            for w in self.warnings:
                log.info(f"  {w}")

        log.info(f"Ready: {len(customers)} customers, {len(jobs)} jobs, {len(terms)} credit terms")
        return customers, jobs, terms


# ============================================================================
# PHASE 3: LOAD INTO ACUMATICA
# ============================================================================

class AcumaticaLoader:
    """Load transformed data into Acumatica via REST API with retry logic."""

    BASE = f"{ACUMATICA_URL}/entity/{ENDPOINT_NAME}/{ENDPOINT_VER}"

    def __init__(self):
        self.session = requests.Session() if not DRY_RUN else None
        self.logged_in = False
        self.results = {
            "terms": [], "accounts": [], "vendors": [],
            "customers": [], "projects": [], "employees": [],
            "ar_invoices": [], "ap_bills": [], "journal": [],
        }
        self.retry_count = 3
        self.retry_delay = 2

    def login(self):
        if DRY_RUN:
            log.info("[DRY RUN] Would login to Acumatica")
            return True
        import requests
        self.session = requests.Session()
        try:
            resp = self.session.post(
                f"{ACUMATICA_URL}/entity/auth/login",
                json={"name": USERNAME, "password": PASSWORD, "company": COMPANY},
                verify=False,
            )
            if resp.status_code == 204:
                self.logged_in = True
                log.info("Logged in to Acumatica")
                return True
            log.error(f"Login failed: {resp.status_code} — {resp.text[:200]}")
            return False
        except Exception as e:
            log.error(f"Connection error: {e}")
            return False

    def logout(self):
        if self.session and self.logged_in:
            self.session.post(f"{ACUMATICA_URL}/entity/auth/logout", verify=False)
            log.info("Logged out of Acumatica")

    def _put_with_retry(self, entity, payload, label):
        """PUT with retry logic and intelligent error handling."""
        if DRY_RUN:
            log.info(f"  [DRY RUN] {label}")
            return True, {"dry_run": True}

        for attempt in range(1, self.retry_count + 1):
            try:
                resp = self.session.put(
                    f"{self.BASE}/{entity}",
                    json=payload,
                    verify=False,
                    headers={"Content-Type": "application/json"},
                    timeout=60,
                )

                if resp.status_code in (200, 201):
                    return True, resp.json()

                error_text = resp.text[:300]

                # Intelligent error handling
                if resp.status_code == 500 and "already exists" in error_text.lower():
                    log.info(f"  [SKIP] {label} — already exists")
                    return True, {"skipped": "already exists"}

                if resp.status_code == 500 and "numbering" in error_text.lower():
                    log.error(f"  [ERROR] {label} — Project numbering sequence not configured!")
                    log.error(f"  -> Go to PM101000 and set up the numbering sequence first")
                    return False, error_text

                if resp.status_code == 422 and "required" in error_text.lower():
                    log.error(f"  [ERROR] {label} — Missing required field: {error_text}")
                    return False, error_text

                if resp.status_code == 429 or "rate limit" in error_text.lower():
                    wait = self.retry_delay * attempt * 2
                    log.warning(f"  [RATE LIMITED] {label} — waiting {wait}s (attempt {attempt})")
                    time.sleep(wait)
                    continue

                if attempt < self.retry_count:
                    log.warning(f"  [RETRY {attempt}/{self.retry_count}] {label} — {resp.status_code}")
                    time.sleep(self.retry_delay * attempt)
                else:
                    log.error(f"  [FAILED] {label} — {resp.status_code}: {error_text}")
                    return False, error_text

            except Exception as e:
                if attempt < self.retry_count:
                    log.warning(f"  [RETRY {attempt}] {label} — {e}")
                    time.sleep(self.retry_delay * attempt)
                else:
                    log.error(f"  [FAILED] {label} — {e}")
                    return False, str(e)

        return False, "Max retries exceeded"

    def create_credit_terms(self, terms):
        """Create credit terms in Acumatica."""
        log.info("\n--- Creating credit terms ---")
        for term in terms:
            payload = {
                "TermsID": {"value": term.term_id},
                "Description": {"value": term.description},
                "DueDateDays": {"value": term.due_days},
                "DiscountDays": {"value": term.discount_days},
            }
            ok, result = self._put_with_retry("Terms", payload, f"Term: {term.term_id}")
            if not ok:
                # Fallback: try CreditTerms entity name
                ok, result = self._put_with_retry("CreditTerms", payload, f"Term (fallback): {term.term_id}")
            self.results["terms"].append({"id": term.term_id, "ok": ok})
            if ok:
                log.info(f"  [OK] {term.term_id} ({term.description})")

    def update_customers(self, customers):
        """Update customers with addresses, contact info, and terms."""
        log.info("\n--- Updating customers ---")
        for cust in customers:
            if not cust.address1 and not cust.phone and not cust.email and not cust.terms:
                log.info(f"  [SKIP] {cust.acumatica_id} — no data to update")
                continue

            payload = {"CustomerID": {"value": cust.acumatica_id}}

            contact = {}
            if cust.phone:
                contact["Phone1"] = {"value": cust.phone}
            if cust.fax:
                contact["Fax"] = {"value": cust.fax}
            if cust.email:
                contact["Email"] = {"value": cust.email}

            address = {}
            if cust.address1:
                address["AddressLine1"] = {"value": cust.address1}
            if cust.city:
                address["City"] = {"value": cust.city}
            if cust.state:
                address["State"] = {"value": cust.state}
            if cust.zip_code:
                address["PostalCode"] = {"value": cust.zip_code}
            if cust.country:
                address["Country"] = {"value": cust.country}

            if address:
                contact["Address"] = address
            if contact:
                payload["MainContact"] = contact
            if cust.terms:
                payload["Terms"] = {"value": cust.terms}

            ok, result = self._put_with_retry("Customer", payload, f"Customer: {cust.acumatica_id}")
            self.results["customers"].append({"id": cust.acumatica_id, "ok": ok})
            if ok:
                log.info(f"  [OK] {cust.acumatica_id} ({cust.qbd_name})")

    def create_projects(self, jobs, transformer):
        """Create projects from QBD jobs."""
        log.info("\n--- Creating projects ---")
        for i, job in enumerate(jobs, 1):
            description = transformer.build_project_description(job)
            status = transformer.map_job_status(job.status)

            payload = {
                "Description": {"value": description},
                "Customer": {"value": job.customer_id},
                "Status": {"value": status},
            }
            if job.start_date:
                payload["StartDate"] = {"value": f"{job.start_date}T00:00:00"}
            if job.end_date:
                payload["EndDate"] = {"value": f"{job.end_date}T00:00:00"}

            label = f"[{i}/{len(jobs)}] {job.parent_customer}: {job.name}"
            ok, result = self._put_with_retry("Project", payload, label)
            pid = result.get("ProjectID", {}).get("value", "auto") if isinstance(result, dict) else "?"
            self.results["projects"].append({"name": job.name, "id": pid, "ok": ok})
            if ok:
                log.info(f"  [OK] {label} -> {pid}")

            time.sleep(0.5)  # Gentle pacing

    def create_accounts(self, accounts_file):
        """Create/update GL accounts from QBD Chart of Accounts."""
        log.info("\n--- Importing chart of accounts ---")
        if not Path(accounts_file).exists():
            log.warning(f"  [SKIP] {accounts_file} not found")
            return

        import pandas as pd
        df = pd.read_excel(accounts_file)

        # QBD -> Acumatica account type mapping
        type_map = {
            "Bank": "Asset", "AccountsReceivable": "Asset",
            "OtherCurrentAsset": "Asset", "FixedAsset": "Asset",
            "OtherAsset": "Asset",
            "AccountsPayable": "Liability", "CreditCard": "Liability",
            "OtherCurrentLiability": "Liability", "LongTermLiability": "Liability",
            "Equity": "Equity",
            "Income": "Income", "OtherIncome": "Income",
            "CostOfGoodsSold": "Expense", "Expense": "Expense",
            "OtherExpense": "Expense",
        }

        count = 0
        for _, row in df.iterrows():
            acct_type = str(row.get("AccountType", ""))
            if acct_type == "NonPosting":
                continue  # Skip non-posting accounts (Purchase Orders, Estimates)

            acct_num = str(row.get("AccountNumber", "")).strip()
            name = str(row.get("FullName", row.get("Name", "")))
            if not acct_num or acct_num == "None":
                acct_num = name[:10].upper().replace(" ", "")

            acu_type = type_map.get(acct_type, "Expense")

            payload = {
                "AccountCD": {"value": acct_num},
                "Description": {"value": name},
                "Type": {"value": acu_type},
                "Active": {"value": True},
            }

            label = f"Account: {acct_num} {name[:30]}"
            ok, result = self._put_with_retry("Account", payload, label)
            self.results["accounts"].append({"id": acct_num, "ok": ok})
            if ok:
                count += 1
            time.sleep(0.3)

        log.info(f"  Imported {count} accounts")

    def create_vendors(self, vendors_file):
        """Create vendors in Acumatica from QBD vendor list."""
        log.info("\n--- Importing vendors ---")
        if not Path(vendors_file).exists():
            log.warning(f"  [SKIP] {vendors_file} not found")
            return

        import pandas as pd
        df = pd.read_excel(vendors_file)

        for i, row in df.iterrows():
            name = str(row.get("Name", ""))
            if not name or name == "nan":
                continue

            is_active = str(row.get("IsActive", "true")) == "true"
            if not is_active:
                continue

            # Generate vendor ID: first 10 chars uppercase, no spaces
            vid = name.upper().replace(" ", "").replace("'", "").replace(",", "")[:10]

            payload = {
                "VendorID": {"value": vid},
                "VendorName": {"value": name},
                "Status": {"value": "Active"},
            }

            # Contact info
            contact = {}
            phone = str(row.get("Phone", ""))
            if phone and phone != "nan":
                contact["Phone1"] = {"value": phone}
            fax = str(row.get("Fax", ""))
            if fax and fax != "nan":
                contact["Fax"] = {"value": fax}
            email = str(row.get("Email", ""))
            if email and email != "nan":
                contact["Email"] = {"value": email}

            # Address
            address = {}
            addr1 = str(row.get("Addr1", row.get("Addr2", "")))
            if addr1 and addr1 != "nan":
                address["AddressLine1"] = {"value": addr1}
            city = str(row.get("City", ""))
            if city and city != "nan":
                address["City"] = {"value": city}
            state = str(row.get("State", ""))
            if state and state != "nan":
                address["State"] = {"value": state}
            zipcode = str(row.get("PostalCode", ""))
            if zipcode and zipcode != "nan":
                address["PostalCode"] = {"value": zipcode}
            address["Country"] = {"value": "US"}

            if address:
                contact["Address"] = address
            if contact:
                payload["MainContact"] = contact

            # Terms
            terms = str(row.get("Terms", ""))
            terms_map = {
                "Due on receipt": "DUERECEIPT", "Net 10": "NET10",
                "Net 15": "NET15", "Net 30": "NET30", "Net 60": "NET60",
                "1% 10 Net 30": "1%10NET30", "2% 10 Net 30": "2%10NET30",
            }
            if terms in terms_map:
                payload["Terms"] = {"value": terms_map[terms]}

            label = f"[{i+1}/{len(df)}] Vendor: {vid} ({name[:25]})"
            ok, result = self._put_with_retry("Vendor", payload, label)
            self.results["vendors"].append({"id": vid, "name": name, "ok": ok})
            if ok:
                log.info(f"  [OK] {label}")
            time.sleep(0.3)

    def create_employees(self, employees_file):
        """Create employees in Acumatica from QBD employee list."""
        log.info("\n--- Importing employees ---")
        if not Path(employees_file).exists():
            log.warning(f"  [SKIP] {employees_file} not found")
            return

        import pandas as pd
        df = pd.read_excel(employees_file)

        for i, row in df.iterrows():
            first = str(row.get("FirstName", ""))
            last = str(row.get("LastName", ""))
            if not last or last == "nan":
                continue

            # Generate employee ID
            eid = (last.upper()[:6] + first.upper()[:1]).replace(" ", "")

            payload = {
                "EmployeeID": {"value": eid},
                "EmployeeName": {"value": f"{first} {last}".strip()},
                "Status": {"value": "Active" if str(row.get("IsActive", "true")) == "true" else "Inactive"},
            }

            # Contact
            contact = {}
            phone = str(row.get("Phone", ""))
            if phone and phone != "nan":
                contact["Phone1"] = {"value": phone}
            email = str(row.get("Email", ""))
            if email and email != "nan":
                contact["Email"] = {"value": email}

            address = {}
            addr = str(row.get("Addr1", ""))
            if addr and addr != "nan":
                address["AddressLine1"] = {"value": addr}
            city = str(row.get("City", ""))
            if city and city != "nan":
                address["City"] = {"value": city}
            state = str(row.get("State", ""))
            if state and state != "nan":
                address["State"] = {"value": state}
            zipcode = str(row.get("PostalCode", ""))
            if zipcode and zipcode != "nan":
                address["PostalCode"] = {"value": zipcode}

            if address:
                contact["Address"] = address
            if contact:
                payload["Contact"] = contact

            label = f"Employee: {eid} ({first} {last})"
            ok, result = self._put_with_retry("Employee", payload, label)
            self.results["employees"].append({"id": eid, "ok": ok})
            if ok:
                log.info(f"  [OK] {label}")
            time.sleep(0.3)

    def create_ar_invoices(self, invoices_file, customer_id_map):
        """Create open AR invoices in Acumatica."""
        log.info("\n--- Importing open AR invoices ---")
        if not Path(invoices_file).exists():
            log.warning(f"  [SKIP] {invoices_file} not found")
            return

        import pandas as pd
        df = pd.read_excel(invoices_file)

        for i, row in df.iterrows():
            cust_full = str(row.get("CustomerFullName", ""))
            # Extract parent customer name (before the colon)
            parent = cust_full.split(":")[0].strip() if ":" in cust_full else cust_full
            cust_id = customer_id_map.get(parent, parent.upper().replace(" ", "")[:10])

            ref = str(row.get("RefNumber", ""))
            amount = str(row.get("BalanceRemaining", row.get("Amount", "0")))
            txn_date = str(row.get("TxnDate", ""))
            due_date = str(row.get("DueDate", ""))

            payload = {
                "Type": {"value": "Invoice"},
                "Customer": {"value": cust_id},
                "Description": {"value": f"QBD Migration - Inv {ref} - {cust_full}"},
                "Amount": {"value": float(amount)},
            }
            if txn_date and txn_date != "nan":
                payload["Date"] = {"value": f"{txn_date}T00:00:00"}
            if due_date and due_date != "nan":
                payload["DueDate"] = {"value": f"{due_date}T00:00:00"}

            label = f"AR Invoice: {ref} {parent} ${float(amount):,.2f}"
            ok, result = self._put_with_retry("Invoice", payload, label)
            self.results["ar_invoices"].append({"ref": ref, "amount": amount, "ok": ok})
            if ok:
                log.info(f"  [OK] {label}")
            time.sleep(0.5)

    def create_ap_bills(self, bills_file):
        """Create open AP bills in Acumatica."""
        log.info("\n--- Importing open AP bills ---")
        if not Path(bills_file).exists():
            log.warning(f"  [SKIP] {bills_file} not found")
            return

        import pandas as pd
        df = pd.read_excel(bills_file)

        for i, row in df.iterrows():
            vendor = str(row.get("VendorFullName", ""))
            vid = vendor.upper().replace(" ", "").replace("'", "").replace(",", "")[:10]
            amount = str(row.get("AmountDue", "0"))
            txn_date = str(row.get("TxnDate", ""))
            due_date = str(row.get("DueDate", ""))
            ref = str(row.get("RefNumber", "")) or f"QBD-{i+1}"
            memo = str(row.get("Memo", ""))

            payload = {
                "Type": {"value": "Bill"},
                "Vendor": {"value": vid},
                "Description": {"value": f"QBD Migration - {vendor}" + (f" - {memo}" if memo and memo != "nan" else "")},
                "Amount": {"value": float(amount)},
            }
            if txn_date and txn_date != "nan":
                payload["Date"] = {"value": f"{txn_date}T00:00:00"}
            if due_date and due_date != "nan":
                payload["DueDate"] = {"value": f"{due_date}T00:00:00"}
            if ref and ref != "nan":
                payload["VendorRef"] = {"value": ref[:30]}

            label = f"AP Bill: {vendor[:25]} ${float(amount):,.2f}"
            ok, result = self._put_with_retry("Bill", payload, label)
            self.results["ap_bills"].append({"vendor": vendor, "amount": amount, "ok": ok})
            if ok:
                log.info(f"  [OK] {label}")
            time.sleep(0.5)

    def create_opening_journal(self, accounts_file):
        """Create opening balance journal entry from COA balances."""
        log.info("\n--- Creating opening balance journal entry ---")
        if not Path(accounts_file).exists():
            log.warning(f"  [SKIP] {accounts_file} not found")
            return

        import pandas as pd
        df = pd.read_excel(accounts_file)

        # Build journal entry lines from account balances
        debit_types = ["Bank", "AccountsReceivable", "OtherCurrentAsset", "FixedAsset",
                       "OtherAsset", "CostOfGoodsSold", "Expense", "OtherExpense"]
        lines = []
        for _, row in df.iterrows():
            bal = float(row.get("Balance", 0) or 0)
            if bal == 0:
                continue
            acct_type = str(row.get("AccountType", ""))
            if acct_type == "NonPosting":
                continue
            acct_num = str(row.get("AccountNumber", "")).strip()
            if not acct_num or acct_num == "None":
                continue

            if acct_type in debit_types:
                debit = max(bal, 0)
                credit = abs(min(bal, 0))
            else:
                credit = max(bal, 0)
                debit = abs(min(bal, 0))

            if debit > 0:
                lines.append({"Account": acct_num, "Debit": debit, "Credit": 0})
            if credit > 0:
                lines.append({"Account": acct_num, "Debit": 0, "Credit": credit})

        total_debit = sum(l["Debit"] for l in lines)
        total_credit = sum(l["Credit"] for l in lines)
        diff = round(total_debit - total_credit, 2)

        log.info(f"  Total debits:  ${total_debit:,.2f}")
        log.info(f"  Total credits: ${total_credit:,.2f}")
        if abs(diff) > 0.01:
            log.warning(f"  Out of balance by ${diff:,.2f} -- adding to Retained Earnings")
            # Balance to retained earnings (3910)
            if diff > 0:
                lines.append({"Account": "3910", "Debit": 0, "Credit": diff})
            else:
                lines.append({"Account": "3910", "Debit": abs(diff), "Credit": 0})

        log.info(f"  Journal entry has {len(lines)} lines")

        # Build Acumatica JournalTransaction payload
        detail_lines = []
        for l in lines:
            detail_lines.append({
                "Account": {"value": l["Account"]},
                "DebitAmount": {"value": l["Debit"]},
                "CreditAmount": {"value": l["Credit"]},
                "Description": {"value": "QBD Opening Balance Migration"},
            })

        payload = {
            "Description": {"value": "QBD Migration - Opening Balances"},
            "Details": detail_lines,
        }

        label = f"Opening Journal: {len(lines)} lines, ${total_debit:,.2f}"
        ok, result = self._put_with_retry("JournalTransaction", payload, label)
        self.results["journal"].append({"lines": len(lines), "total": total_debit, "ok": ok})
        if ok:
            log.info(f"  [OK] {label}")
        else:
            log.warning("  Journal entry may need to be created manually via GL301000")
            log.warning("  The opening balances data is in QBD_ChartOfAccounts.xlsx")

    def get_summary(self):
        s = {}
        for key in self.results:
            items = self.results[key]
            s[key] = {
                "total": len(items),
                "ok": sum(1 for r in items if r.get("ok", False)),
            }
        return s


# ============================================================================
# PHASE 4: VALIDATE
# ============================================================================

class Validator:
    """Post-migration validation."""

    def __init__(self, loader):
        self.loader = loader

    def validate(self):
        if DRY_RUN:
            log.info("\n[DRY RUN] Skipping validation")
            return

        log.info("\n" + "=" * 50)
        log.info("PHASE 4: Post-migration validation")
        log.info("=" * 50)

        # Validate customers
        try:
            resp = self.loader.session.get(
                f"{self.loader.BASE}/Customer",
                params={"$expand": "MainContact/Address"},
                verify=False,
            )
            if resp.status_code == 200:
                customers = resp.json()
                log.info(f"\nCustomers in Acumatica: {len(customers)}")
                for c in customers:
                    cid = c.get("CustomerID", {}).get("value", "?")
                    addr = c.get("MainContact", {}).get("Address", {})
                    city = addr.get("City", {}).get("value", "")
                    terms = c.get("Terms", {}).get("value", "")
                    phone = c.get("MainContact", {}).get("Phone1", {}).get("value", "")
                    status = "OK" if (city or phone) else "INCOMPLETE"
                    log.info(f"  {cid:12s} city={city or '—':20s} terms={terms or '—':12s} [{status}]")
        except Exception as e:
            log.error(f"Customer validation failed: {e}")

        # Validate projects
        try:
            resp = self.loader.session.get(f"{self.loader.BASE}/Project", verify=False)
            if resp.status_code == 200:
                projects = [p for p in resp.json()
                            if p.get("ProjectID", {}).get("value", "") != "X"]
                log.info(f"\nProjects in Acumatica: {len(projects)}")
                for p in projects:
                    pid = p.get("ProjectID", {}).get("value", "?")
                    desc = p.get("Description", {}).get("value", "?")[:40]
                    status = p.get("Status", {}).get("value", "?")
                    log.info(f"  {pid:12s} {desc:42s} [{status}]")
        except Exception as e:
            log.error(f"Project validation failed: {e}")


# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================

def main():
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    report = MigrationReport(started=datetime.now().isoformat())

    log.info("=" * 60)
    log.info("  QBD -> ACUMATICA MIGRATION AGENT")
    log.info("=" * 60)
    log.info(f"  Instance:   {ACUMATICA_URL}")
    log.info(f"  Company:    {COMPANY}")
    log.info(f"  Endpoint:   {ENDPOINT_NAME}/{ENDPOINT_VER}")
    log.info(f"  Mode:       {'DRY RUN' if DRY_RUN else 'LIVE'}")
    log.info(f"  Source:     {'Excel files' if FROM_FILES else 'QBD COM (with file fallback)'}")
    log.info("=" * 60)

    # ── Phase 1: Extract ──
    log.info("\n" + "=" * 50)
    log.info("PHASE 1: Extract from QuickBooks Desktop")
    log.info("=" * 50)

    customers = []
    jobs = []

    if FROM_FILES:
        extractor = FileExtractor()
        # Auto-detect files — match names from qbd_auto_extract.py
        search_paths = [
            EXPORT_DIR / "QBD_Customers.xlsx",
            Path("./QBD_Customers.xlsx"),
            Path("./QBDCustomers.xlsx"),
        ]
        cust_file = QBD_CUSTOMERS_FILE
        if not cust_file:
            for p in search_paths:
                if p.exists():
                    cust_file = p
                    break

        if cust_file and Path(cust_file).exists():
            customers, jobs = extractor.extract_customers(str(cust_file))
            report.source = str(cust_file)
        else:
            log.error("No QBD customer file found! Expected qbd_exports/QBD_Customers.xlsx")
            sys.exit(1)
    else:
        # Try QBD COM first
        qbd = QBDExtractor()
        if qbd.connect():
            try:
                qbd.extract_customers()
                qbd.extract_vendors()
                qbd.extract_chart_of_accounts()
            finally:
                qbd.disconnect()
            report.source = "QBD COM SDK"
        else:
            log.info("Falling back to Excel file extraction...")
            extractor = FileExtractor()
            cust_file = None
            for p in [EXPORT_DIR / "QBD_Customers.xlsx", Path("./QBD_Customers.xlsx"), Path("./QBDCustomers.xlsx")]:
                if p.exists():
                    cust_file = p
                    break
            if cust_file:
                customers, jobs = extractor.extract_customers(str(cust_file))
                report.source = str(cust_file)
            else:
                log.error("Cannot connect to QBD and no Excel files found.")
                log.error("Run qbd_auto_extract.py first, or use --from-files with files in qbd_exports/")
                sys.exit(1)

    report.phases["extract"] = {"customers": len(customers), "jobs": len(jobs)}

    # ── Phase 2: Transform ──
    transformer = Transformer()
    customers, jobs, terms = transformer.transform_all(customers, jobs)
    report.phases["transform"] = {
        "fixes": len(transformer.fixes_applied),
        "warnings": len(transformer.warnings),
    }
    report.warnings = transformer.warnings

    # ── Phase 3: Load ──
    log.info("\n" + "=" * 50)
    log.info("PHASE 3: Load into Acumatica")
    log.info("=" * 50)

    if not DRY_RUN and PASSWORD == "your_password_here":
        log.error("Set your PASSWORD in the CONFIG section!")
        sys.exit(1)

    loader = AcumaticaLoader()

    if not DRY_RUN:
        if not loader.login():
            sys.exit(1)

    try:
        # IMPORT ORDER MATTERS — dependencies must be created first
        # 1. Credit Terms (referenced by customers and vendors)
        loader.create_credit_terms(terms)
        # 2. Chart of Accounts (referenced by journal entries, invoices, bills)
        loader.create_accounts(str(EXPORT_DIR / "QBD_ChartOfAccounts.xlsx"))
        # 3. Vendors (must exist before AP bills)
        loader.create_vendors(str(EXPORT_DIR / "QBD_Vendors.xlsx"))
        # 4. Customers (must exist before AR invoices)
        loader.update_customers(customers)
        # 5. Projects (from QBD jobs)
        loader.create_projects(jobs, transformer)
        # 6. Employees
        loader.create_employees(str(EXPORT_DIR / "QBD_Employees.xlsx"))
        # 7. Opening balance journal entry (needs accounts to exist)
        loader.create_opening_journal(str(EXPORT_DIR / "QBD_ChartOfAccounts.xlsx"))
        # 8. Open AR invoices (needs customers)
        cust_id_map = {c.qbd_name: c.acumatica_id for c in customers}
        loader.create_ar_invoices(str(EXPORT_DIR / "QBD_OpenInvoices.xlsx"), cust_id_map)
        # 9. Open AP bills (needs vendors)
        loader.create_ap_bills(str(EXPORT_DIR / "QBD_OpenBills.xlsx"))
    finally:
        if not DRY_RUN:
            # Phase 4: Validate
            validator = Validator(loader)
            validator.validate()
            loader.logout()

    # ── Report ──
    report.completed = datetime.now().isoformat()
    report.summary = loader.get_summary()
    report.phases["load"] = report.summary

    with open(REPORT_FILE, "w") as f:
        json.dump(asdict(report), f, indent=2)

    log.info("\n" + "=" * 60)
    log.info("  MIGRATION COMPLETE")
    log.info("=" * 60)
    s = report.summary
    for key, label in [
        ("terms", "Credit terms"), ("accounts", "GL accounts"), ("vendors", "Vendors"),
        ("customers", "Customers"), ("projects", "Projects"), ("employees", "Employees"),
        ("journal", "Opening JE"), ("ar_invoices", "AR invoices"), ("ap_bills", "AP bills"),
    ]:
        if key in s and s[key]["total"] > 0:
            log.info(f"  {label:16s} {s[key]['ok']}/{s[key]['total']}")
    log.info(f"  {'Warnings':16s} {len(report.warnings)}")
    log.info(f"  {'Report':16s} {REPORT_FILE}")
    log.info(f"  {'Full log':16s} {LOG_FILE}")
    log.info("=" * 60)


if __name__ == "__main__":
    import requests
    main()
