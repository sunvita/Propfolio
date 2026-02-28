"""
PDF parsers for property P&L:
  1. Rental / Ownership statements  (e.g. property management disbursements)
  2. Bank transaction statements
  3. Utility bills  (electricity, water, gas, internet)
  4. Invoices / Notices  (council rates, land tax, strata, insurance, trades, etc.)

Parser version: 2025-02-28-v3
"""

import re
import pdfplumber
from io import BytesIO


# ── P&L category mapping: keyword → (section, category) ─────────────────────
# Used by bank statement categoriser
BANK_CATEGORIES = {
    # ── Income ───────────────────────────────────────────────────────────────
    'rental income':        ('income',    'Rental Income'),
    'rent received':        ('income',    'Rental Income'),
    'eft to owner':         ('income',    'Rental Income'),
    'disbursement':         ('income',    'Rental Income'),

    # ── OpEx — Management ─────────────────────────────────────────────────────
    'management fee':       ('opex',      'Management Fees'),
    'property management':  ('opex',      'Management Fees'),
    'admin fee':            ('opex',      'Management Fees'),
    'administration fee':   ('opex',      'Management Fees'),
    'letting fee':          ('opex',      'Letting Fees'),
    'leasing fee':          ('opex',      'Letting Fees'),
    'lease renewal':        ('opex',      'Letting Fees'),
    'tenant placement':     ('opex',      'Letting Fees'),
    'reletting':            ('opex',      'Letting Fees'),

    # ── OpEx — Maintenance & Repairs ─────────────────────────────────────────
    'maintenance':          ('opex',      'Maintenance & Repairs'),
    'repair':               ('opex',      'Maintenance & Repairs'),
    'handyman':             ('opex',      'Maintenance & Repairs'),
    'trade service':        ('opex',      'Maintenance & Repairs'),
    # Plumbing
    'plumber':              ('opex',      'Maintenance & Repairs'),
    'plumbing':             ('opex',      'Maintenance & Repairs'),
    'drain':                ('opex',      'Maintenance & Repairs'),
    'blocked drain':        ('opex',      'Maintenance & Repairs'),
    'tap ':                 ('opex',      'Maintenance & Repairs'),
    'taps ':                ('opex',      'Maintenance & Repairs'),
    'toilet':               ('opex',      'Maintenance & Repairs'),
    'cistern':              ('opex',      'Maintenance & Repairs'),
    'pipe ':                ('opex',      'Maintenance & Repairs'),
    'pipes ':               ('opex',      'Maintenance & Repairs'),
    'hot water':            ('opex',      'Maintenance & Repairs'),
    'water heater':         ('opex',      'Maintenance & Repairs'),
    # Electrical
    'electrical':           ('opex',      'Maintenance & Repairs'),
    'electrician':          ('opex',      'Maintenance & Repairs'),
    'wiring':               ('opex',      'Maintenance & Repairs'),
    'switchboard':          ('opex',      'Maintenance & Repairs'),
    'light fitting':        ('opex',      'Maintenance & Repairs'),
    'smoke alarm':          ('opex',      'Maintenance & Repairs'),
    'safety switch':        ('opex',      'Maintenance & Repairs'),
    'power point':          ('opex',      'Maintenance & Repairs'),
    # Locks & access
    'locksmith':            ('opex',      'Maintenance & Repairs'),
    'lock ':                ('opex',      'Maintenance & Repairs'),
    'locks ':               ('opex',      'Maintenance & Repairs'),
    'keys ':                ('opex',      'Maintenance & Repairs'),
    'key cutting':          ('opex',      'Maintenance & Repairs'),
    'access card':          ('opex',      'Maintenance & Repairs'),
    'deadbolt':             ('opex',      'Maintenance & Repairs'),
    'door handle':          ('opex',      'Maintenance & Repairs'),
    # Pest & vermin
    'pest control':         ('opex',      'Maintenance & Repairs'),
    'termite':              ('opex',      'Maintenance & Repairs'),
    'vermin':               ('opex',      'Maintenance & Repairs'),
    'rodent':               ('opex',      'Maintenance & Repairs'),
    'cockroach':            ('opex',      'Maintenance & Repairs'),
    # Roofing & structure
    'roofing':              ('opex',      'Maintenance & Repairs'),
    'roof repair':          ('opex',      'Maintenance & Repairs'),
    'gutters':              ('opex',      'Maintenance & Repairs'),
    'gutter clean':         ('opex',      'Maintenance & Repairs'),
    'downpipe':             ('opex',      'Maintenance & Repairs'),
    'fascia':               ('opex',      'Maintenance & Repairs'),
    'ceiling':              ('opex',      'Maintenance & Repairs'),
    'wall repair':          ('opex',      'Maintenance & Repairs'),
    'plaster':              ('opex',      'Maintenance & Repairs'),
    'rendering':            ('opex',      'Maintenance & Repairs'),
    'waterproofing':        ('opex',      'Maintenance & Repairs'),
    # Flooring
    'carpet':               ('opex',      'Maintenance & Repairs'),
    'flooring':             ('opex',      'Maintenance & Repairs'),
    'tiling':               ('opex',      'Maintenance & Repairs'),
    'tile ':                ('opex',      'Maintenance & Repairs'),
    'tiles ':               ('opex',      'Maintenance & Repairs'),
    'grout':                ('opex',      'Maintenance & Repairs'),
    'floorboard':           ('opex',      'Maintenance & Repairs'),
    'vinyl':                ('opex',      'Maintenance & Repairs'),
    # Painting & cosmetic
    'painting':             ('opex',      'Maintenance & Repairs'),
    'painter':              ('opex',      'Maintenance & Repairs'),
    'touch up':             ('opex',      'Maintenance & Repairs'),
    'patching':             ('opex',      'Maintenance & Repairs'),
    # Doors, windows & glazing
    'glazier':              ('opex',      'Maintenance & Repairs'),
    'window repair':        ('opex',      'Maintenance & Repairs'),
    'glass repair':         ('opex',      'Maintenance & Repairs'),
    'screen repair':        ('opex',      'Maintenance & Repairs'),
    'door repair':          ('opex',      'Maintenance & Repairs'),
    'roller door':          ('opex',      'Maintenance & Repairs'),
    'garage door':          ('opex',      'Maintenance & Repairs'),
    # HVAC & appliances
    'air conditioning':     ('opex',      'Maintenance & Repairs'),
    'air con':              ('opex',      'Maintenance & Repairs'),
    'aircon':               ('opex',      'Maintenance & Repairs'),
    'hvac':                 ('opex',      'Maintenance & Repairs'),
    'split system':         ('opex',      'Maintenance & Repairs'),
    'ducted':               ('opex',      'Maintenance & Repairs'),
    'appliance':            ('opex',      'Maintenance & Repairs'),
    'oven repair':          ('opex',      'Maintenance & Repairs'),
    'dishwasher repair':    ('opex',      'Maintenance & Repairs'),
    'washing machine':      ('opex',      'Maintenance & Repairs'),
    'rangehood':            ('opex',      'Maintenance & Repairs'),
    # Fencing, gates & external
    'fencing':              ('opex',      'Maintenance & Repairs'),
    'fence repair':         ('opex',      'Maintenance & Repairs'),
    'gate repair':          ('opex',      'Maintenance & Repairs'),
    'concreting':           ('opex',      'Maintenance & Repairs'),
    'driveway':             ('opex',      'Maintenance & Repairs'),
    'paving':               ('opex',      'Maintenance & Repairs'),
    'retaining wall':       ('opex',      'Maintenance & Repairs'),
    'carpentry':            ('opex',      'Maintenance & Repairs'),
    'carpenter':            ('opex',      'Maintenance & Repairs'),
    'joinery':              ('opex',      'Maintenance & Repairs'),
    'cabinet':              ('opex',      'Maintenance & Repairs'),
    # Pool & spa
    'pool service':         ('opex',      'Maintenance & Repairs'),
    'pool repair':          ('opex',      'Maintenance & Repairs'),
    'spa repair':           ('opex',      'Maintenance & Repairs'),
    'pool chemical':        ('opex',      'Maintenance & Repairs'),

    # ── OpEx — Cleaning ───────────────────────────────────────────────────────
    'cleaning':             ('opex',      'Cleaning'),
    'clean ':               ('opex',      'Cleaning'),
    'cleaner':              ('opex',      'Cleaning'),
    'bond clean':           ('opex',      'Cleaning'),
    'end of lease':         ('opex',      'Cleaning'),
    'vacate clean':         ('opex',      'Cleaning'),
    'exit clean':           ('opex',      'Cleaning'),
    'move out clean':       ('opex',      'Cleaning'),
    'pressure wash':        ('opex',      'Cleaning'),
    'window clean':         ('opex',      'Cleaning'),
    'carpet clean':         ('opex',      'Cleaning'),
    'steam clean':          ('opex',      'Cleaning'),
    'rubbish removal':      ('opex',      'Cleaning'),
    'waste removal':        ('opex',      'Cleaning'),
    'junk removal':         ('opex',      'Cleaning'),
    'skip bin':             ('opex',      'Cleaning'),
    # Gardening & grounds
    'lawn':                 ('opex',      'Cleaning'),
    'mowing':               ('opex',      'Cleaning'),
    'mow ':                 ('opex',      'Cleaning'),
    'garden':               ('opex',      'Cleaning'),
    'garden maintenance':   ('opex',      'Cleaning'),
    'gardening':            ('opex',      'Cleaning'),
    'gardener':             ('opex',      'Cleaning'),
    'landscaping':          ('opex',      'Cleaning'),
    'landscaper':           ('opex',      'Cleaning'),
    'hedging':              ('opex',      'Cleaning'),
    'pruning':              ('opex',      'Cleaning'),
    'tree lopping':         ('opex',      'Cleaning'),
    'tree removal':         ('opex',      'Cleaning'),
    'arborist':             ('opex',      'Cleaning'),
    'weeding':              ('opex',      'Cleaning'),
    'mulching':             ('opex',      'Cleaning'),
    'irrigation':           ('opex',      'Cleaning'),
    'turf':                 ('opex',      'Cleaning'),
    'council rates':        ('opex',      'Council Rates'),
    'municipal rates':      ('opex',      'Council Rates'),
    'rates notice':         ('opex',      'Council Rates'),
    'land tax':             ('opex',      'Land Tax'),
    'state revenue':        ('opex',      'Land Tax'),
    'revenue nsw':          ('opex',      'Land Tax'),
    'osr ':                 ('opex',      'Land Tax'),
    'strata levy':          ('opex',      'Strata / Body Corporate'),
    'body corporate':       ('opex',      'Strata / Body Corporate'),
    'owners corporation':   ('opex',      'Strata / Body Corporate'),
    'strata management':    ('opex',      'Strata / Body Corporate'),
    'building insurance':   ('opex',      'Building Insurance'),
    'landlord insurance':   ('opex',      'Building Insurance'),
    'property insurance':   ('opex',      'Building Insurance'),
    'insurance premium':    ('opex',      'Building Insurance'),
    'advertising':          ('opex',      'Advertising'),
    'photography':          ('opex',      'Advertising'),

    # ── Utilities ─────────────────────────────────────────────────────────────
    'electricity':          ('utilities', 'Electricity'),
    'energy':               ('utilities', 'Electricity'),
    'ausgrid':              ('utilities', 'Electricity'),
    'agl':                  ('utilities', 'Electricity'),
    'origin energy':        ('utilities', 'Electricity'),
    'simply energy':        ('utilities', 'Electricity'),
    'alinta':               ('utilities', 'Electricity'),
    'water':                ('utilities', 'Water'),
    'sydney water':         ('utilities', 'Water'),
    'icon water':           ('utilities', 'Water'),
    'gas':                  ('utilities', 'Gas'),
    'jemena':               ('utilities', 'Gas'),
    'internet':             ('utilities', 'Internet'),
    'broadband':            ('utilities', 'Internet'),
    'nbn':                  ('utilities', 'Internet'),
    'telstra':              ('utilities', 'Internet'),
    'optus':                ('utilities', 'Internet'),
    'iinet':                ('utilities', 'Internet'),
    'aussie broadband':     ('utilities', 'Internet'),

    # ── Financing ────────────────────────────────────────────────────────────
    'mortgage':             ('financing', 'Mortgage Interest'),
    'home loan':            ('financing', 'Mortgage Interest'),
    'loan interest':        ('financing', 'Mortgage Interest'),
    'loan repayment':       ('financing', 'Mortgage Repayment'),
    'principal':            ('financing', 'Mortgage Repayment'),
}

# ── Invoice type detection: keywords → P&L category ─────────────────────────
INVOICE_CATEGORY_MAP = [
    # (keyword list, section, pl_category)
    (['council rates', 'rates notice', 'rate notice',
      'municipal rates', 'local council', 'council levy', 'quarterly rates',
      'local government rates', 'government rates and charges',
      'rates and charges', 'general grv', 'grv valuation',
      'rubbish/recycling service', 'rubbish recycling service',
      'emergency services levy'],
     'opex', 'Council Rates'),

    (['land tax', 'land value tax', 'state revenue office', 'revenue nsw',
      'notice of assessment', 'land tax assessment'],
     'opex', 'Land Tax'),

    (['strata levy', 'body corporate', 'owners corporation',
      'strata management', 'strata plan', 'administrative fund',
      'sinking fund', 'capital works fund'],
     'opex', 'Strata / Body Corporate'),

    (['landlord insurance', 'building insurance', 'property insurance',
      'insurance premium', 'policy renewal', 'certificate of insurance'],
     'opex', 'Building Insurance'),

    (['handyman', 'trade services', 'pest control', 'termite',
      'plumber', 'plumbing', 'electrician', 'electrical',
      'locksmith', 'painter', 'carpentry', 'carpenter',
      'roofing', 'roofer', 'gutters', 'gutter', 'air conditioning',
      'hvac', 'hot water system', 'carpet', 'tiling', 'tile',
      'concreting', 'fencing', 'fence'],
     'opex', 'Maintenance & Repairs'),

    (['cleaning service', 'bond clean', 'end of lease clean',
      'lawn mowing', 'garden maintenance', 'landscaping',
      'rubbish removal', 'window cleaning'],
     'opex', 'Cleaning'),

    (['property management', 'management fee', 'management agreement'],
     'opex', 'Management Fees'),

    (['letting fee', 'leasing fee', 'tenant placement'],
     'opex', 'Letting Fees'),

    (['real estate photography', 'advertising', 'domain listing',
      'realestate.com', 'marketing'],
     'opex', 'Advertising'),

    (['electricity', 'energy usage', 'kwh', 'power bill',
      'electricity charge'],
     'utilities', 'Electricity'),

    (['water use', 'water service', 'sewerage', 'water usage',
      'water consumption'],
     'utilities', 'Water'),

    (['natural gas', 'gas usage', 'gas service charge'],
     'utilities', 'Gas'),

    (['internet service', 'broadband', 'nbn service', 'data usage'],
     'utilities', 'Internet'),
]

MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4,  'may': 5,  'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10,
    'november': 11, 'december': 12,
}


# ── Low-level helpers ─────────────────────────────────────────────────────────
def _extract_address(text: str) -> str:
    """
    Extract a property / service address from PDF text.
    Tries labeled fields first, then falls back to Australian address regex.
    Returns empty string if nothing found.
    """
    # 1. Labeled field patterns (order = specificity)
    labeled = [
        r'property\s+address[:\s]+([^\n]+)',
        r'service\s+address[:\s]+([^\n]+)',
        r'supply\s+address[:\s]+([^\n]+)',
        r'installation\s+address[:\s]+([^\n]+)',
        r'premises[:\s]+([^\n]+)',
        r'rental\s+property[:\s]+([^\n]+)',
        r'site\s+address[:\s]+([^\n]+)',
        r'property\s+location[:\s]+([^\n]+)',
        r'delivery\s+address[:\s]+([^\n]+)',
        r'property\s+details[:\s]+([^\n]+)',
        r'address\s+of\s+supply[:\s]+([^\n]+)',
    ]
    for pat in labeled:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            addr = re.sub(r'\s+', ' ', m.group(1).strip())
            if 6 < len(addr) < 200:
                return addr

    # 2. Australian address regex fallback
    aus_pat = re.compile(
        r'\d+[A-Za-z]?\s+[\w\'\-]+(?:\s+[\w\'\-]+){0,3}\s+'
        r'(?:Street|St|Avenue|Ave|Av|Road|Rd|Drive|Dr|Place|Pl|Court|Ct|'
        r'Crescent|Cres|Cr|Boulevard|Blvd|Lane|Ln|Lne|Way|Wy|Close|Cl|'
        r'Circuit|Cct|Cir|Parade|Pde|Terrace|Tce|Highway|Hwy|'
        r'Grove|Gr|Gve|Parkway|Pkwy|Park|Pk|Square|Sq|'
        r'Freeway|Fwy|Rise|Green|Grn|Gate|Gte|Gardens|Gts|Mews|Loop)\b'
        r'(?:[,\s]+[\w\s]+?)?[,\s]+'
        r'(?:NSW|VIC|QLD|WA|SA|TAS|ACT|NT)\s+\d{4}',
        re.IGNORECASE
    )
    m = aus_pat.search(text)
    if m:
        return re.sub(r'\s+', ' ', m.group(0).strip())

    return ''


def _extract_text(file_bytes: bytes) -> str:
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        return '\n'.join(p.extract_text() or '' for p in pdf.pages)


def _extract_tables(file_bytes: bytes) -> list:
    tables = []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for t in page.extract_tables():
                tables.append(t)
    return tables


def _extract_rental_from_tables(file_bytes: bytes) -> dict:
    """
    Tier B — table-based extraction.
    Scans pdfplumber tables for label/value rows found in property management
    software that renders proper grid tables (PropertyMe, Console, Palace, etc.).
    Returns a partial dict with any fields found; may be empty.
    """
    found: dict = {}

    # Maps label substrings → result field (most specific listed first)
    LABEL_MAP = [
        # ── money_in ───────────────────────────────────────────────────────
        (['money in', 'total receipts', 'gross income', 'total income',
          'total rent', 'rental income', 'income received',
          'total trust receipts'], 'money_in'),
        # ── money_out ──────────────────────────────────────────────────────
        (['money out', 'total paid in agency', 'agency fee',
          'management fee', 'total fees', 'total disbursements',
          'total deductions', 'total charges', 'total expenses',
          'total trust disbursements'], 'money_out'),
        # ── eft ────────────────────────────────────────────────────────────
        (['you received', 'eft to owner', 'withdrawal by eft',
          'disbursement to owner', 'net amount', 'total forwarded',
          'total remitted', 'net proceeds', 'owner payout',
          'amount paid to owner', 'balance remaining',
          'net owner payment', 'owner disbursement'], 'eft'),
    ]

    tables = _extract_tables(file_bytes)
    for table in tables:
        for row in table:
            if not row or len(row) < 2:
                continue
            # Use first non-empty cell as the label
            label = next((str(c).strip().lower() for c in row if c and str(c).strip()), '')
            if not label:
                continue
            # Collect all positive numeric values beyond the label cell
            amounts = [
                v for v in (_parse_amount(c) for c in row[1:])
                if v is not None and v > 0
            ]
            if not amounts:
                continue
            for keywords, field in LABEL_MAP:
                if field in found:
                    continue
                if any(k in label for k in keywords):
                    found[field] = amounts[0]
                    break

    return found


def _llm_extract_rental(text: str) -> dict:
    """
    Tier C — LLM fallback using Claude API (Haiku).
    Only called when both regex and table extraction return no figures.
    Requires ANTHROPIC_API_KEY as an environment variable or Streamlit secret.
    Returns a partial dict or {} on any error (silent degradation).
    Cost: ~$0.0004 per call (negligible).
    """
    import os
    api_key = os.environ.get('ANTHROPIC_API_KEY', '')
    if not api_key:
        try:
            import streamlit as st          # only available at runtime
            api_key = st.secrets.get('ANTHROPIC_API_KEY', '')
        except Exception:
            pass
    if not api_key:
        return {}

    try:
        import anthropic
    except ImportError:
        return {}  # package not installed — skip silently

    try:
        client = anthropic.Anthropic(api_key=api_key)
        prompt = (
            "Extract the following fields from this rental/ownership statement.\n"
            "Return ONLY a JSON object — no explanation, no markdown fences.\n"
            "Keys:\n"
            "  money_in  – total rental income received (number)\n"
            "  money_out – total management/agency fees charged (number)\n"
            "  eft       – net amount disbursed/transferred to the owner (number)\n"
            "  year      – statement year (integer)\n"
            "  month     – statement month, 1–12 (integer)\n"
            "  address   – rental property street address (string)\n"
            "Use null for any field you cannot confidently identify.\n\n"
            f"Statement text:\n{text[:3000]}"
        )
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        # Strip any accidental markdown code fences
        raw = re.sub(r'^```(?:json)?\s*|\s*```$', '', raw, flags=re.MULTILINE).strip()

        import json
        data = json.loads(raw)

        result: dict = {}
        for key in ('money_in', 'money_out', 'eft'):
            v = data.get(key)
            if v is not None:
                try:
                    result[key] = float(v)
                except (TypeError, ValueError):
                    pass
        for key in ('year', 'month'):
            v = data.get(key)
            if v is not None:
                try:
                    result[key] = int(v)
                except (TypeError, ValueError):
                    pass
        if data.get('address'):
            result['address'] = str(data['address'])
        return result

    except Exception:
        return {}  # any API / parse error — degrade silently


def _parse_amount(s) -> float | None:
    if not s:
        return None
    s = str(s).strip().replace(',', '').replace('$', '').replace(' ', '')
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None


def _detect_year_month(text: str) -> tuple[int, int] | None:
    """Try to pull a statement month/year from text."""
    patterns = [
        # High-priority labeled dates (rate notices, invoices)
        r'issue\s*date[:\s]+(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        r'date\s+of\s+issue[:\s]+(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        r'date\s+of\s+payment[:\s]+(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        r'invoice\s+date[:\s]+(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        r'tax\s+invoice[^:]*:\s*(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        r'billing\s+date[:\s]+(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        r'statement\s+date[:\s]+(\d{1,2})[/.](\d{1,2})[/.](\d{4})',
        # General patterns
        r'(january|february|march|april|may|june|july|august|'
        r'september|october|november|december)\s+(\d{4})',
        r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[- ](\d{2,4})',
        r'(\d{1,2})[/-](\d{4})\b',
        r'period[:\s]+\d{1,2}[/-]\d{1,2}[/-](\d{4})',
        r'\b(\d{1,2})\s+(january|february|march|april|may|june|july|august|'
        r'september|october|november|december)\s+(\d{4})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            g = m.groups()
            try:
                if len(g) == 2:
                    month_str, year_str = g
                    month = MONTH_MAP.get(month_str.lower()[:3])
                    if not month:
                        month = int(month_str)
                    year = int(year_str) if len(year_str) == 4 else 2000 + int(year_str)
                    if month and 1 <= month <= 12 and 2000 <= year <= 2050:
                        return year, month
                elif len(g) == 3:
                    day, month_str, year_str = g
                    month = MONTH_MAP.get(month_str.lower()[:3]) or int(month_str)
                    year = int(year_str)
                    if month and 1 <= month <= 12 and 2000 <= year <= 2050:
                        return year, month
            except (ValueError, TypeError):
                continue
    return None


def _categorize_by_keywords(description: str) -> tuple[str, str]:
    """Return (section, category) by matching BANK_CATEGORIES keywords.
    Longer (more specific) keywords take priority over shorter ones so that
    e.g. 'carpet clean' beats 'carpet', and 'garden maintenance' beats 'maintenance'.
    """
    desc_lower = description.lower()
    for kw in sorted(BANK_CATEGORIES, key=len, reverse=True):
        if kw in desc_lower:
            return BANK_CATEGORIES[kw]
    return ('opex', 'Miscellaneous')


def _extract_invoice_amount(text: str) -> float:
    """Try common invoice total patterns to extract the payable amount.

    Handles both AUD and $ currency prefixes.
    Currency token: (?:AUD\s*|\$\s*)? — optional, matches either format.
    """
    _CUR = r'(?:AUD\s*|\$\s*)?'   # optional currency prefix: "AUD " or "$"

    patterns = [
        # RevenueWA BPay receipt: "Amount paid today AUD 823.10"
        # Use *before* surcharge line to get base tax amount
        rf'amount\s+paid\s+today\s+{_CUR}([\d,]+\.?\d*)',
        # Generic total / due patterns ($ or AUD)
        rf'total\s+amount\s+due[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'amount\s+(?:due|payable)[:\s]+{_CUR}([\d,]+\.?\d*)',
        # Rate notice: "Amount Due by 1 September 2023 $2,503.83"
        rf'amount\s+due\s+by[^$\n]{{0,40}}{_CUR}([\d,]+\.?\d*)',
        # Rate notice: "Full Payment Due 01/09/2023 $2,503.83"
        rf'full\s*payment\s*due[^$\n]{{0,30}}{_CUR}([\d,]+\.?\d*)',
        # Rate notice: "Payment Option 1 Full Payment ... $X"
        rf'payment\s+option\s*1[^$\n]{{0,50}}{_CUR}([\d,]+\.?\d*)',
        rf'invoice\s+total[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'total\s+incl(?:\.|\s+)?\s*gst[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'total\s+inc(?:\.|\s+)?\s*gst[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'total\s+including\s+gst[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'balance\s+due[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'please\s+pay[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'payment\s+required[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'amount\s+to\s+pay[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'net\s+amount[:\s]+{_CUR}([\d,]+\.?\d*)',
        rf'\btotal[:\s]+{_CUR}([\d,]+\.?\d*)',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = _parse_amount(m.group(1))
            if val and val > 0:
                return val
    return 0.0


# ── 1. RENTAL / OWNERSHIP STATEMENT ──────────────────────────────────────────
def parse_rental_statement(file_bytes: bytes, filename: str = '') -> dict:
    text = _extract_text(file_bytes)
    result = {
        'type': 'rental', 'filename': filename,
        'year': None, 'month': None,
        'money_in': 0.0, 'money_out': 0.0, 'eft': 0.0,
        'rooms': {}, 'pl_items': {}, 'raw_text': text[:2000],
        'parse_source': 'regex',   # updated to 'table' or 'llm' if fallback used
    }

    # ── Date detection ──────────────────────────────────────────────────────
    # Priority 1: Ailo/platform "Ownership statement [Month] [Year]"
    _own_m = re.search(
        r'ownership\s+statement\s+'
        r'(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|'
        r'jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|'
        r'nov(?:ember)?|dec(?:ember)?)\s+(\d{4})',
        text, re.IGNORECASE
    )
    if _own_m:
        _month = MONTH_MAP.get(_own_m.group(1).lower()[:3])
        _year  = int(_own_m.group(2))
        if _month:
            result['year'], result['month'] = _year, _month
    else:
        # Priority 2: "Statement period  1 [Month] [Year] — 30 [Month] [Year]"
        # Use the END date of the period (closing month)
        _period_m = re.search(
            r'statement\s+period[:\s]*\d{1,2}\s+\w+\s+\d{4}\s*[—\-–]+\s*\d{1,2}\s+'
            r'(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|'
            r'jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|'
            r'nov(?:ember)?|dec(?:ember)?)\s+(\d{4})',
            text, re.IGNORECASE
        )
        if _period_m:
            _month = MONTH_MAP.get(_period_m.group(1).lower()[:3])
            _year  = int(_period_m.group(2))
            if _month:
                result['year'], result['month'] = _year, _month
        else:
            ym = _detect_year_month(text)
            if ym:
                result['year'], result['month'] = ym

    # ── Address detection ────────────────────────────────────────────────────
    # Priority 1 (Ailo): extract property address from "Room N, [address] Net income:"
    _addr_m = re.search(r'Room\s+\d+,\s+(.+?)\s+Net income:', text, re.IGNORECASE)
    if _addr_m:
        result['extracted_address'] = re.sub(r'\s+', ' ', _addr_m.group(1).strip())
    else:
        result['extracted_address'] = _extract_address(text)

    # ── Financial figures ────────────────────────────────────────────────────
    # Detect Ailo platform format ("Ownership statement" header present)
    _is_ailo = bool(re.search(r'ownership\s+statement\s+\w+\s+\d{4}', text, re.IGNORECASE))

    if _is_ailo:
        # money_in: "Income   $780.00  $0.00  $780.00"  → first $ value (In column)
        _m = re.search(r'^\s*Income\s+\$([\d,]+\.?\d*)', text, re.IGNORECASE | re.MULTILINE)
        if _m:
            result['money_in'] = _parse_amount(_m.group(1)) or 0.0

        # money_out: "Total paid in agency fees  $85.80"
        _m = re.search(r'Total\s+paid\s+in\s+agency\s+fees\s+\$([\d,]+\.?\d*)',
                       text, re.IGNORECASE)
        if _m:
            result['money_out'] = _parse_amount(_m.group(1)) or 0.0

        # eft: sum of all per-room "Net income: $X" values
        # This equals money_in minus ALL expenses (fees + bills) = true net to owner
        _net_vals = re.findall(r'Net income:\s+\$([\d,]+\.?\d*)', text, re.IGNORECASE)
        if _net_vals:
            result['eft'] = round(sum(_parse_amount(v) or 0.0 for v in _net_vals), 2)
        elif result['money_in'] > 0:
            # Fallback: Income_in minus Expenses_out column
            _exp_m = re.search(
                r'^\s*Expenses\s+\$[\d,]+\.?\d*\s+\$([\d,]+\.?\d*)',
                text, re.IGNORECASE | re.MULTILINE
            )
            if _exp_m:
                _exp_out = _parse_amount(_exp_m.group(1)) or 0.0
                result['eft'] = round(result['money_in'] - _exp_out, 2)

    else:
        # Generic patterns for other management platforms
        # (e.g. "Money In / Money Out / You Received / EFT to owner")
        for _label, _key in [
            (r'money\s+in[:\s]+\$?([\d,]+\.?\d*)',                     'money_in'),
            (r'money\s+out[:\s]+\$?([\d,]+\.?\d*)',                    'money_out'),
            (r'you\s+received[:\s]+\$?([\d,]+\.?\d*)',                 'eft'),
            (r'withdrawal\s+by\s+eft[^$\n]{0,60}\$?([\d,]+\.?\d*)',   'eft'),
            (r'eft\s+to\s+owner[^$\n]{0,30}\$?([\d,]+\.?\d*)',        'eft'),
            (r'eft[^$\d\n]{0,20}\$?([\d,]+\.?\d*)',                    'eft'),
            (r'net\s+amount[:\s]+\$?([\d,]+\.?\d*)',                   'eft'),
            (r'disbursement\s+to\s+owner[:\s]+\$?([\d,]+\.?\d*)',     'eft'),
        ]:
            _m = re.search(_label, text, re.IGNORECASE)
            if _m:
                _val = _parse_amount(_m.group(1))
                if _val is not None and (result[_key] == 0.0 or _key == 'eft'):
                    result[_key] = _val

    # ── Ailo itemised bill extraction ────────────────────────────────────────
    # Extract individual expense lines ("Category · description $amount") so
    # each bill appears as its own P&L line item.
    # Structure:  Income → Mgmt Fees → [bill items by category] → Net = Transfer
    if _is_ailo:
        _BILL_SKIP = re.compile(
            r'^(rent\s+payment|management\s+fees?|paid\s+on|contributions?|'
            r'failed|transfer\s+to|withdrawal|total|gst|overview|income|expenses)',
            re.IGNORECASE
        )
        # Match:  "[Category] · [details...] $amount" on a single line
        _bill_pattern = re.compile(
            r'^([A-Za-z][^\n·]{1,80}?)\s+·\s+[^\n$]*\$([\d,]+\.?\d*)\s*$',
            re.MULTILINE
        )
        _bill_totals: dict[str, float] = {}
        for _bm in _bill_pattern.finditer(text):
            _cat_text = _bm.group(1).strip()
            _amt      = _parse_amount(_bm.group(2)) or 0.0
            if _amt <= 0 or _BILL_SKIP.match(_cat_text):
                continue
            _section, _pl_cat = _categorize_by_keywords(_cat_text)
            # Only capture opex/utilities — skip income-side matches
            if _section in ('opex', 'utilities'):
                _bill_totals[_pl_cat] = round(
                    _bill_totals.get(_pl_cat, 0.0) + _amt, 2
                )
        # Add each bill category to pl_items (skip Mgmt Fees — already in money_out)
        for _pl_cat, _amt in _bill_totals.items():
            if _pl_cat != 'Management Fees':
                result['pl_items'][_pl_cat] = _amt

    # ── Room breakdown ───────────────────────────────────────────────────────
    if _is_ailo:
        # Ailo: "Room N, [address]  Net income: $X" — reliable, address numbers don't interfere
        for _rm in re.finditer(
            r'(Room\s+\d+),\s+[^\n]+?Net income:\s+\$([\d,]+\.?\d*)',
            text, re.IGNORECASE
        ):
            _rname = _rm.group(1).strip().title()
            _rnet  = _parse_amount(_rm.group(2)) or 0.0
            result['rooms'][_rname] = {'rent': _rnet, 'mgmt': 0.0, 'net': _rnet}
    else:
        # Generic: locate each room/unit heading then find its "Total $out $in" summary
        # row within the next 600 chars.  This avoids matching address numbers or
        # per-week rent figures that appear earlier in the room block.
        for _rm_m in re.finditer(
            r'(room\s*\d+\b|unit\s*\w+\b)(?!\s*/)', text, re.IGNORECASE
        ):
            _rname   = _rm_m.group(1).strip().title()
            _segment = text[_rm_m.start(): _rm_m.start() + 600]
            _tot = re.search(
                r'Total\s+\$([\d,]+\.?\d*)\s+\$([\d,]+\.?\d*)',
                _segment, re.IGNORECASE
            )
            if _tot:
                # Column order in PropertyMe-style PDFs is Out (fees) then In (rent)
                _mgmt = _parse_amount(_tot.group(1)) or 0.0
                _rent = _parse_amount(_tot.group(2)) or 0.0
                result['rooms'][_rname] = {
                    'rent': _rent, 'mgmt': _mgmt,
                    'net':  round(_rent - _mgmt, 2),
                }

    # ── Tier B: table fallback ─────────────────────────────────────────────
    # Trigger when regex extracted nothing meaningful
    if result['money_in'] == 0.0 and result['eft'] == 0.0:
        _tbl = _extract_rental_from_tables(file_bytes)
        _any_table = False
        for _k in ('money_in', 'money_out', 'eft'):
            if _tbl.get(_k, 0.0) > 0.0:
                result[_k] = _tbl[_k]
                _any_table = True
        if _any_table:
            result['parse_source'] = 'table'

    # ── Tier C: LLM fallback ───────────────────────────────────────────────
    # Trigger only if table extraction also found nothing
    if result['money_in'] == 0.0 and result['eft'] == 0.0:
        _llm = _llm_extract_rental(text)
        _any_llm = False
        for _k in ('money_in', 'money_out', 'eft'):
            if _llm.get(_k, 0.0) > 0.0:
                result[_k] = _llm[_k]
                _any_llm = True
        # LLM can also fill in missing date / address
        if _llm.get('year') and not result['year']:
            result['year'] = _llm['year']
        if _llm.get('month') and not result['month']:
            result['month'] = _llm['month']
        if _llm.get('address') and not result.get('extracted_address'):
            result['extracted_address'] = _llm['address']
        if _any_llm:
            result['parse_source'] = 'llm'
        else:
            result['parse_source'] = 'failed'

    # Set base items — bill items extracted earlier in the Ailo branch are preserved
    result['pl_items']['Rental Income']   = result['money_in']
    result['pl_items']['Management Fees'] = result['money_out']
    return result


# ── 2. BANK TRANSACTION STATEMENT ────────────────────────────────────────────
def parse_bank_statement(file_bytes: bytes, filename: str = '') -> dict:
    text = _extract_text(file_bytes)
    tables = _extract_tables(file_bytes)

    result = {
        'type': 'bank', 'filename': filename,
        'year': None, 'month': None,
        'transactions': [], 'categorized': {}, 'raw_text': text[:2000],
    }

    ym = _detect_year_month(text)
    if ym:
        result['year'], result['month'] = ym

    transactions = []

    for table in tables:
        if not table or len(table) < 2:
            continue
        header = [str(h).lower() if h else '' for h in table[0]]

        date_col   = next((i for i, h in enumerate(header) if 'date' in h), None)
        desc_col   = next((i for i, h in enumerate(header)
                           if any(k in h for k in ['desc', 'detail', 'narr',
                                                    'particular', 'transaction'])), None)
        debit_col  = next((i for i, h in enumerate(header)
                           if 'debit' in h or 'withdraw' in h), None)
        credit_col = next((i for i, h in enumerate(header)
                           if 'credit' in h or 'deposit' in h), None)
        amount_col = next((i for i, h in enumerate(header) if 'amount' in h), None)

        if desc_col is None:
            continue

        for row in table[1:]:
            if not row or all(c is None or str(c).strip() == '' for c in row):
                continue
            desc = str(row[desc_col]).strip() if desc_col < len(row) else ''
            if not desc or desc.lower() in ('', 'none'):
                continue

            amount, t_type = 0.0, 'debit'

            if debit_col is not None and credit_col is not None:
                d = _parse_amount(row[debit_col]  if debit_col  < len(row) else None)
                c = _parse_amount(row[credit_col] if credit_col < len(row) else None)
                if c and c > 0:
                    amount, t_type = c, 'credit'
                elif d and d > 0:
                    amount, t_type = d, 'debit'
            elif amount_col is not None:
                v = _parse_amount(row[amount_col] if amount_col < len(row) else None)
                if v is not None:
                    amount  = abs(v)
                    t_type  = 'credit' if v >= 0 else 'debit'

            date_str = str(row[date_col]).strip() if date_col is not None and date_col < len(row) else ''
            section, category = _categorize_by_keywords(desc)

            transactions.append({
                'date': date_str, 'description': desc,
                'amount': round(amount, 2), 'type': t_type,
                'section': section, 'category': category,
            })

    # Fallback: line-by-line regex
    if not transactions:
        line_pat = re.compile(
            r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+(.+?)\s+\$?([\d,]+\.\d{2})\s*$'
        )
        for line in text.splitlines():
            m = line_pat.search(line)
            if m:
                date_str, desc, amt_str = m.groups()
                amount = _parse_amount(amt_str) or 0.0
                t_type = 'credit' if any(k in line.lower()
                                         for k in ['cr', 'credit', 'deposit']) else 'debit'
                section, category = _categorize_by_keywords(desc)
                transactions.append({
                    'date': date_str, 'description': desc.strip(),
                    'amount': round(amount, 2), 'type': t_type,
                    'section': section, 'category': category,
                })

    result['transactions'] = transactions

    cat_totals: dict[str, dict[str, float]] = {}
    for tx in transactions:
        sec, cat = tx['section'], tx['category']
        amt = tx['amount'] if tx['type'] == 'credit' else -tx['amount']
        cat_totals.setdefault(sec, {}).setdefault(cat, 0.0)
        cat_totals[sec][cat] = round(cat_totals[sec][cat] + amt, 2)

    result['categorized'] = cat_totals
    return result


# ── 3. UTILITY BILL ───────────────────────────────────────────────────────────
def parse_utility_bill(file_bytes: bytes, filename: str = '') -> dict:
    text = _extract_text(file_bytes)
    text_lower = text.lower()

    result = {
        'type': 'utility', 'filename': filename,
        'year': None, 'month': None,
        'utility_type': 'Miscellaneous',
        'amount': 0.0, 'raw_text': text[:1000],
    }

    ym = _detect_year_month(text)
    if ym:
        result['year'], result['month'] = ym

    # Detect utility type (order matters — more specific first)
    if any(k in text_lower for k in ['kwh', 'electricity charge', 'energy charge',
                                      'ausgrid', 'agl', 'origin energy',
                                      'simply energy', 'alinta energy']):
        result['utility_type'] = 'Electricity'
    elif any(k in text_lower for k in ['water use', 'water service', 'sewerage charge',
                                        'sydney water', 'icon water', 'water consumption',
                                        'water usage']):
        result['utility_type'] = 'Water'
    elif any(k in text_lower for k in ['natural gas', 'gas usage', 'gas service charge',
                                        'jemena', 'gas meter']):
        result['utility_type'] = 'Gas'
    elif any(k in text_lower for k in ['internet', 'broadband', 'nbn service',
                                        'data usage', 'telstra', 'optus',
                                        'iinet', 'aussie broadband']):
        result['utility_type'] = 'Internet'
    elif any(k in text_lower for k in ['water', 'gas', 'energy']):
        # Broader fallback
        if 'water' in text_lower:
            result['utility_type'] = 'Water'
        elif 'gas' in text_lower:
            result['utility_type'] = 'Gas'

    result['amount'] = _extract_invoice_amount(text)
    result['extracted_address'] = _extract_address(text)
    return result


# ── 4. INVOICE / NOTICE ───────────────────────────────────────────────────────
def parse_invoice(file_bytes: bytes, filename: str = '') -> dict:
    """
    Parse tax invoices and government notices:
    council rates, land tax, strata levies, insurance, trades, etc.
    """
    text = _extract_text(file_bytes)
    text_lower = text.lower()

    result = {
        'type': 'invoice', 'filename': filename,
        'year': None, 'month': None,
        'category': 'Miscellaneous',
        'section': 'opex',
        'amount': 0.0,
        'gst': 0.0,
        'vendor': '',
        'raw_text': text[:1000],
    }

    ym = _detect_year_month(text)
    if ym:
        result['year'], result['month'] = ym

    # Detect category
    for keywords, section, category in INVOICE_CATEGORY_MAP:
        if any(k in text_lower for k in keywords):
            result['section']  = section
            result['category'] = category
            break

    result['amount'] = _extract_invoice_amount(text)

    # GST extraction
    gm = re.search(r'gst[:\s]+\$?([\d,]+\.?\d*)', text, re.IGNORECASE)
    if gm:
        result['gst'] = _parse_amount(gm.group(1)) or 0.0

    # Vendor / issuer (first line or "from:" pattern)
    for line in text.splitlines()[:10]:
        line = line.strip()
        if line and len(line) > 3 and not line[0].isdigit():
            result['vendor'] = line
            break

    result['pl_items'] = {result['category']: result['amount']}
    result['extracted_address'] = _extract_address(text)
    return result


# ── AUTO-DETECT & DISPATCH ────────────────────────────────────────────────────
def parse_pdf(file_bytes: bytes, filename: str = '', doc_type: str = 'auto') -> dict:
    """
    Parse a PDF, auto-detecting or using explicit doc_type.
    doc_type: 'auto' | 'rental' | 'bank' | 'utility' | 'invoice'
    """
    if doc_type == 'rental':
        return parse_rental_statement(file_bytes, filename)
    if doc_type == 'bank':
        return parse_bank_statement(file_bytes, filename)
    if doc_type == 'utility':
        return parse_utility_bill(file_bytes, filename)
    if doc_type == 'invoice':
        return parse_invoice(file_bytes, filename)

    text = _extract_text(file_bytes).lower()

    # ── 1. Rental / ownership statement ──────────────────────────────────────
    if any(k in text for k in ['money in', 'money out', 'ownership statement',
                                'eft to owner', 'disbursement to owner',
                                'landlord statement']):
        return parse_rental_statement(file_bytes, filename)

    # ── 2. Government notices & rates ────────────────────────────────────────
    if any(k in text for k in ['council rates', 'rates notice', 'rate notice',
                                'municipal rates', 'local government rates',
                                'government rates and charges', 'general grv',
                                'grv valuation', 'land tax assessment',
                                'notice of assessment', 'revenue nsw', 'state revenue']):
        return parse_invoice(file_bytes, filename)

    # ── 3. Strata / body corporate ────────────────────────────────────────────
    if any(k in text for k in ['strata levy', 'body corporate levy',
                                'owners corporation', 'administrative fund levy',
                                'sinking fund levy', 'capital works levy']):
        return parse_invoice(file_bytes, filename)

    # ── 4. Trade / service invoices ───────────────────────────────────────────
    if any(k in text for k in ['tax invoice', 'invoice no', 'invoice number',
                                'abn:', 'australian business number']) and \
       any(k in text for k in ['total', 'amount due', 'amount payable',
                                'balance due', 'please pay']):
        return parse_invoice(file_bytes, filename)

    # ── 5. Utility bills (before bank to avoid false match on "account number") ─
    utility_keywords = [
        'amount due', 'amount payable', 'kwh', 'usage charge',
        'bill amount', 'water use', 'service charge',
        'electricity charge', 'energy charge', 'gas charge',
        'broadband', 'nbn service', 'data usage',
        'total charges', 'please pay by',
    ]
    if any(k in text for k in utility_keywords):
        return parse_utility_bill(file_bytes, filename)

    # ── 6. Bank transaction statement ─────────────────────────────────────────
    if any(k in text for k in ['account number', 'bsb', 'opening balance',
                                'closing balance', 'available balance',
                                'statement of account']):
        return parse_bank_statement(file_bytes, filename)

    # ── Default ────────────────────────────────────────────────────────────────
    return parse_bank_statement(file_bytes, filename)
