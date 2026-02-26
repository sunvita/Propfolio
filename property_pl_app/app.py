"""
Propfolio — Property Portfolio P&L
Streamlit app: PDF Upload → Parse → Review → Generate Excel
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import json
import re
import datetime
from difflib import SequenceMatcher

from parsers.pdf_parser import parse_pdf
from generators.excel_gen import build_workbook

# ── Helper functions (defined before UI blocks) ────────────────────────────────
def _add_or_update(data_dict, key, value):
    data_dict[key] = round(data_dict.get(key, 0) + value, 2)


def _normalize_address(s: str) -> str:
    """Lowercase, expand common abbreviations, remove punctuation noise."""
    s = s.lower().strip()
    # Full Australian street-type abbreviation map (long-form → canonical)
    abbr = {
        r'\bst\b':    'street',    r'\bave\b':  'avenue',   r'\bav\b':   'avenue',
        r'\brd\b':    'road',      r'\bdr\b':   'drive',    r'\bpl\b':   'place',
        r'\bct\b':    'court',     r'\bcres\b': 'crescent', r'\bcr\b':   'crescent',
        r'\bln\b':    'lane',      r'\blne\b':  'lane',     r'\bblvd\b': 'boulevard',
        r'\bcl\b':    'close',     r'\bcct\b':  'circuit',  r'\bcir\b':  'circuit',
        r'\bhwy\b':   'highway',   r'\bpde\b':  'parade',   r'\btce\b':  'terrace',
        r'\bgr\b':    'grove',     r'\bgve\b':  'grove',    r'\bpkwy\b': 'parkway',
        r'\bpk\b':    'park',      r'\bsq\b':   'square',   r'\bwy\b':   'way',
        r'\bfwy\b':   'freeway',   r'\bvw\b':   'view',     r'\bvws\b':  'views',
        r'\brise\b':  'rise',      r'\bgrn\b':  'green',    r'\bgte\b':  'gate',
        r'\bgts\b':   'gardens',   r'\bmws\b':  'mews',     r'\bloop\b': 'loop',
    }
    for pat, repl in abbr.items():
        s = re.sub(pat, repl, s)
    s = re.sub(r'[,\.]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _match_address(extracted: str, reference: str) -> tuple[str, float, str]:
    """
    Compare extracted PDF address with user-configured reference address.
    Returns (status_label, similarity_ratio, css_color).

    Priority logic:
    1. If Australian postcodes (4-digit) exist in both → compare them first.
       - Different postcodes  → ❌  (definitive mismatch)
       - Same postcode, same leading number → ✅
       - Same postcode, different number → ⚠️  (same suburb, verify unit/number)
    2. No postcode in one/both → fall back to token overlap + fuzzy ratio.
    """
    if not extracted:
        return ('⚪ Address not found in PDF', 0.0, '#9E9E9E')
    if not reference:
        return ('⚪ No reference address configured', 0.0, '#9E9E9E')

    ext_n = _normalize_address(extracted)
    ref_n = _normalize_address(reference)

    # ── 1. Postcode-anchored comparison ───────────────────────────────────────
    ext_pc = re.search(r'\b(\d{4})\b', ext_n)
    ref_pc = re.search(r'\b(\d{4})\b', ref_n)

    if ext_pc and ref_pc:
        if ext_pc.group(1) != ref_pc.group(1):
            return ('❌ Address mismatch — different postcode', 0.0, '#C62828')

        # Same postcode: compare leading street number
        ext_num = re.match(r'^([\w/]+)', ext_n)
        ref_num = re.match(r'^([\w/]+)', ref_n)
        if ext_num and ref_num and ext_num.group(1) == ref_num.group(1):
            return ('✅ Address matched', 1.0, '#2E7D32')

        # Same suburb, different number (unit variation, mis-typed, or wrong property)
        ratio = SequenceMatcher(None, ext_n, ref_n).ratio()
        if ratio >= 0.70:
            return ('⚠️ Same suburb — verify street number/unit', ratio, '#F57F17')
        return ('⚠️ Partial match — please verify', ratio, '#F57F17')

    # ── 2. Fallback: token overlap + sequence ratio ────────────────────────────
    ratio = SequenceMatcher(None, ext_n, ref_n).ratio()
    ext_tokens = set(ext_n.split())
    ref_tokens  = set(ref_n.split())
    state_tokens = {'nsw','vic','qld','wa','sa','tas','act','nt'}
    key_match = bool(ext_tokens & ref_tokens & state_tokens)
    common_frac = len(ext_tokens & ref_tokens) / max(len(ref_tokens), 1)

    if ratio >= 0.80 or (key_match and common_frac >= 0.5):
        return ('✅ Address matched', ratio, '#2E7D32')
    elif ratio >= 0.50 or common_frac >= 0.35 or key_match:
        return ('⚠️ Partial match — please verify', ratio, '#F57F17')
    return ('❌ Address mismatch', ratio, '#C62828')


def _session_to_json() -> str:
    """Serialize current session (property configs + all P&L data) to JSON string."""
    props_out = []
    for p in st.session_state.get('properties', []):
        # Convert (year, month) tuple keys → "YYYY-MM" strings for JSON
        data_str = {
            f"{k[0]}-{k[1]:02d}": v
            for k, v in p.get('data', {}).items()
        }
        props_out.append({'name': p['name'], 'tab': p['tab'], 'data': data_str})

    payload = {
        'version':        '1.1',
        'saved_at':       datetime.datetime.now().strftime('%Y-%m-%d'),
        'fy_start_month': st.session_state.get('fy_start_month', 7),
        'fy_labels':      st.session_state.get('fy_labels', []),
        'prop_configs':   st.session_state.get('prop_configs', []),
        'purchase_info':  st.session_state.get('purchase_info', {}),
        'properties':     props_out,
    }
    return json.dumps(payload, indent=2, ensure_ascii=False)


def _session_from_json(raw: dict) -> tuple[bool, str]:
    """
    Restore session from a parsed JSON dict.
    Returns (success, message).
    """
    try:
        st.session_state.fy_start_month = raw['fy_start_month']
        st.session_state.fy_labels      = raw['fy_labels']
        st.session_state.prop_configs   = raw['prop_configs']
        st.session_state.purchase_info  = raw.get('purchase_info', {})

        props = []
        for p in raw['properties']:
            # Convert "YYYY-MM" string keys back to (int, int) tuples
            data_tuples = {}
            for k, v in p.get('data', {}).items():
                yr, mo = map(int, k.split('-'))
                data_tuples[(yr, mo)] = v
            props.append({'name': p['name'], 'tab': p['tab'], 'data': data_tuples})

        st.session_state.properties     = props
        st.session_state.session_loaded = True
        st.session_state.merge_change_log = []

        n_props  = len(props)
        n_months = sum(len(p['data']) for p in props)
        return True, (f"Session restored: {n_props} properties, "
                      f"{n_months} months of data (saved {raw.get('saved_at','?')})")
    except Exception as e:
        return False, f"Failed to load session: {e}"


def _merge_parsed_to_properties():
    """
    Merge parsed PDF results into st.session_state.properties[].data.

    Merge strategy (auto-selected):
      • Fresh session  → additive (accumulate amounts into the month bucket)
      • Loaded session → smart update:
          - New month          → add all items
          - Existing month, item value changed  → overwrite + log '🔄 Updated'
          - Existing month, item value unchanged → skip  + log '➖ Unchanged'
          - Existing month, new item             → add   + log '➕ Added'

    Change log stored in st.session_state.merge_change_log for Step 3 display.
    """
    update_mode = st.session_state.get('session_loaded', False)
    change_log  = []

    for result in st.session_state.parsed_results:
        tab = result.get('_prop_tab')
        yr  = result.get('year')
        mo  = result.get('month')
        if not (tab and yr and mo):
            continue

        include_key = f"include_{tab}_{result.get('filename', '')}"
        if not st.session_state.get(include_key, True):
            continue

        prop = next((p for p in st.session_state.properties if p['tab'] == tab), None)
        if not prop:
            continue

        key          = (yr, mo)
        is_new_month = key not in prop['data']
        prop['data'].setdefault(key, {})

        # ── Build new_items dict from parse result ──────────────────────────
        new_items: dict[str, float] = {}
        if result['type'] == 'rental':
            new_items = {
                'Rental Income':       result.get('money_in',  0),
                'Management Fees':     result.get('money_out', 0),
                'Cash Received (EFT)': result.get('eft',       0),
            }
        elif result['type'] == 'bank':
            for _sec, cats in result.get('categorized', {}).items():
                for cat, amt in cats.items():
                    new_items[cat] = new_items.get(cat, 0) + amt
        elif result['type'] == 'utility':
            utype = result.get('utility_type', 'Miscellaneous')
            new_items = {utype: result.get('amount', 0)}
        elif result['type'] == 'invoice':
            pl_cat = result.get('category', 'Miscellaneous')
            new_items = {pl_cat: result.get('amount', 0)}

        period_str = f"{yr}-{mo:02d}"

        # ── Apply merge ─────────────────────────────────────────────────────
        if update_mode and not is_new_month:
            # Smart update: compare item-by-item
            for label, new_val in new_items.items():
                new_val = round(new_val, 2)
                old_val = prop['data'][key].get(label)

                if old_val is None:
                    prop['data'][key][label] = new_val
                    change_log.append({
                        'status': '➕ Added', 'property': tab,
                        'period': period_str, 'item': label,
                        'old': '—', 'new': f'${new_val:,.2f}',
                    })
                elif abs(old_val - new_val) > 0.005:
                    prop['data'][key][label] = new_val
                    change_log.append({
                        'status': '🔄 Updated', 'property': tab,
                        'period': period_str, 'item': label,
                        'old': f'${old_val:,.2f}', 'new': f'${new_val:,.2f}',
                    })
                else:
                    change_log.append({
                        'status': '➖ Unchanged', 'property': tab,
                        'period': period_str, 'item': label,
                        'old': f'${old_val:,.2f}', 'new': f'${new_val:,.2f}',
                    })
        else:
            # Fresh session or new month — additive merge
            for label, new_val in new_items.items():
                _add_or_update(prop['data'][key], label, new_val)

            if update_mode and is_new_month:
                change_log.append({
                    'status': '🆕 New month', 'property': tab,
                    'period': period_str, 'item': f'{len(new_items)} items',
                    'old': '—', 'new': '—',
                })

    st.session_state.merge_change_log = change_log

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Propfolio — Property Portfolio P&L",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2F5496 100%);
        color: white; padding: 20px 24px; border-radius: 8px;
        margin-bottom: 24px;
    }
    .step-badge {
        background: #FFC000; color: #1F3864; font-weight: bold;
        padding: 2px 10px; border-radius: 12px; font-size: 13px;
        display: inline-block; margin-bottom: 8px;
    }
    .info-box {
        background: #EBF3FB; border-left: 4px solid #2F5496;
        padding: 12px 16px; border-radius: 4px; margin: 8px 0;
        color: #1A237E !important;
    }
    .success-box {
        background: #E8F5E9; border-left: 4px solid #00B050;
        padding: 12px 16px; border-radius: 4px; margin: 8px 0;
        color: #1B5E20 !important;
    }
    .warn-box {
        background: #FFF8E1; border-left: 4px solid #FFC000;
        padding: 12px 16px; border-radius: 4px; margin: 8px 0;
        color: #5D4037 !important;
    }
    div[data-testid="stExpander"] { border: 1px solid #DEE2E6; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

# ── Session state init ──────────────────────────────────────────────────────────
for key, default in {
    'step':                 0,      # 0 = guide page (landing)
    'properties':           [],
    'parsed_results':       [],
    'fy_start_month':       7,
    'fy_labels':            ['2029-30','2028-29','2027-28','2026-27','2025-26','2024-25'],
    'session_loaded':       False,
    'merge_change_log':     [],
    'parse_done':           False,  # Phase A→B flag for Step 2
    'uploaded_files_meta':  [],     # Cached file bytes for Phase B
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

MONTH_NAMES = {1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
               7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}

def make_fy_labels(first_year: int, last_year: int) -> list[str]:
    """Generate FY label list newest→oldest. e.g. first=2024, last=2029 → ['2029-30',...,'2024-25']"""
    return [f'{y}-{str(y+1)[2:]}' for y in range(last_year, first_year - 1, -1)]

# ── Sidebar: steps ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style="padding:12px 4px 4px;">
<svg width="168" height="56" viewBox="0 0 220 74" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect x="10" y="40" width="9"  height="6"  rx="2" fill="#B0BEC5"/>
  <rect x="21" y="30" width="9"  height="16" rx="2" fill="#4F6FB5"/>
  <rect x="32" y="18" width="9"  height="28" rx="2" fill="#1A237E"/>
  <polyline points="2,46 35,6 66,46"
    stroke="#F57F17" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <polyline points="57,6 67,6 67,16"
    stroke="#F57F17" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
  <text x="84" y="36"
    font-family="-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"
    font-size="32" font-weight="800" fill="#1A237E" letter-spacing="-1">Prop</text>
  <text x="84" y="68"
    font-family="-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"
    font-size="32" font-weight="300" fill="#F57F17" letter-spacing="2">folio</text>
</svg>
</div>
""", unsafe_allow_html=True)
    st.markdown("---")

    # ⓪ Getting Started (guide page) — always clickable except when already there
    if st.session_state.step == 0:
        st.markdown("**▶ ⓪ Getting Started**")
    else:
        if st.button("⓪ Getting Started", use_container_width=True,
                     key="sidebar_guide"):
            st.session_state.step = 0
            st.rerun()

    # Steps 1–4
    steps = ["① Setup", "② Upload PDFs", "③ Review & Edit", "④ Generate Excel"]
    for i, s in enumerate(steps, 1):
        if st.session_state.step == i:
            st.markdown(f"**▶ {s}**")
        elif i == 1 and st.session_state.step != 1:
            # ① Setup is always navigable (forward from guide, back from later steps)
            saved = st.session_state.get('prop_configs')
            icon  = '✅' if st.session_state.step > 1 else '○'
            label = f"{icon} {s}" + (' 💾' if saved else '')
            if st.button(label, use_container_width=True, key="sidebar_setup"):
                st.session_state.step = 1
                st.rerun()
        elif st.session_state.step > i:
            st.markdown(f"✅ {s}")
        else:
            st.markdown(f"○ {s}")

    st.markdown("---")
    if st.button("🔄 Start Over", use_container_width=True):
        for k in ['step', 'properties', 'parsed_results',
                  'session_loaded', 'merge_change_log',
                  'parse_done', 'uploaded_files_meta']:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 0: Getting Started — User Guide
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.step == 0:
    st.markdown(
        '<div class="main-header">'
        '<svg width="260" height="86" viewBox="0 0 220 74" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<rect x="10" y="40" width="9"  height="6"  rx="2" fill="#37474F"/>'
        '<rect x="21" y="30" width="9"  height="16" rx="2" fill="#5C6BC0"/>'
        '<rect x="32" y="18" width="9"  height="28" rx="2" fill="#7986CB"/>'
        '<polyline points="2,46 35,6 66,46" stroke="#FFA726" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>'
        '<polyline points="57,6 67,6 67,16" stroke="#FFA726" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>'
        '<text x="84" y="36" font-family="-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif" font-size="32" font-weight="800" fill="#FFFFFF" letter-spacing="-1">Prop</text>'
        '<text x="84" y="68" font-family="-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif" font-size="32" font-weight="300" fill="#FFA726" letter-spacing="2">folio</text>'
        '</svg>'
        '<p style="margin:10px 0 0;opacity:0.85;font-size:15px;">Your Property Portfolio P&amp;L, Simplified.</p>'
        '<p style="margin:6px 0 0;opacity:0.75;font-size:13px;">'
        'Upload your property PDFs — rental statements, bank records, utility bills, and invoices — '
        'and Propfolio automatically builds a fully formatted P&amp;L report for each property '
        'and your entire portfolio. No spreadsheets, no manual entry.</p>'
        '</div>',
        unsafe_allow_html=True
    )

    # ── Overview cards ────────────────────────────────────────────────────────
    st.markdown("### What does this app do?")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            '<div style="background:#EBF3FB;border-radius:10px;padding:20px 18px 18px;color:#1A237E;">'
            '<div style="font-size:15px;font-weight:700;margin-bottom:8px;">📄 Upload PDFs</div>'
            '<div style="font-size:13px;line-height:1.6;color:#2C3E7A;">'
            'Drop in your rental statements, bank records, utility bills, '
            'and tax invoices — the app reads them automatically.'
            '</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(
            '<div style="background:#E8F5E9;border-radius:10px;padding:20px 18px 18px;color:#1B5E20;">'
            '<div style="font-size:15px;font-weight:700;margin-bottom:8px;">🔍 Auto-Parse & Validate</div>'
            '<div style="font-size:13px;line-height:1.6;color:#2E5E35;">'
            'Extracts amounts, dates, and P&amp;L categories. '
            'Cross-checks addresses against your property to flag wrong files.'
            '</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(
            '<div style="background:#FFF8E1;border-radius:10px;padding:20px 18px 18px;color:#5D4037;">'
            '<div style="font-size:15px;font-weight:700;margin-bottom:8px;">📊 Generate Excel</div>'
            '<div style="font-size:13px;line-height:1.6;color:#6D4C41;">'
            'Outputs a color-coded, formula-linked Excel workbook '
            'with a P&amp;L tab per property plus a portfolio summary dashboard.'
            '</div></div>', unsafe_allow_html=True)

    st.markdown('<div style="margin-bottom:32px;"></div>', unsafe_allow_html=True)

    # ── Two-column: step flow + supported PDFs ────────────────────────────────
    left, right = st.columns([3, 2])

    with left:
        st.markdown("### How to use it")

        st.markdown(
            '<div class="info-box">'
            '<b>🆕 First-time setup (4 steps)</b>'
            '</div>', unsafe_allow_html=True)
        st.markdown("""
① **Setup** — Enter the number of properties, financial year range, and property details (address, purchase price, etc.).

② **Upload PDFs** — Upload documents for each property. The app auto-detects the document type and validates the address. You can include or exclude any file manually.

③ **Review & Edit** — Check the parsed data in editable tables. Add or correct any missing months manually.

④ **Generate Excel** — Download your formatted portfolio Excel. Also download the **Session JSON** to save your work.
""")

        st.markdown(
            '<div class="success-box">'
            '<b>🔄 Monthly update (returning users)</b>'
            '</div>', unsafe_allow_html=True)
        st.markdown("""
On the **Setup** page, open **"Load previous session"** and upload your saved JSON file.
Then go straight to **Upload PDFs** and upload only the new month's documents.

The app will automatically:
- **Add** data for new months
- **Update** values that have changed in existing months
- **Keep** unchanged values as-is

After generating, download the updated Session JSON to replace your previous one.
""")

    with right:
        st.markdown("### Supported PDF types")
        st.markdown(
            '<div style="background:#F8F9FA;border-radius:8px;padding:16px;color:#333;">'

            '<p><b>📋 Rental / Ownership Statement</b><br>'
            '<span style="color:#555;font-size:13px;">Property management disbursement reports — '
            'extracts rental income, management fees, and EFT amount.</span></p>'

            '<p><b>🏦 Bank Statement</b><br>'
            '<span style="color:#555;font-size:13px;">Transaction history — '
            'auto-categorises debits/credits into P&amp;L items '
            '(mortgage, repairs, insurance, etc.).</span></p>'

            '<p><b>💡 Utility Bill</b><br>'
            '<span style="color:#555;font-size:13px;">Electricity, water, gas, or internet bills — '
            'detects provider and maps to the correct utility line.</span></p>'

            '<p><b>🧾 Tax Invoice / Notice</b><br>'
            '<span style="color:#555;font-size:13px;">Council rates, land tax assessments, '
            'strata levies, building insurance, and trade invoices '
            '(plumber, electrician, painter, etc.).</span></p>'

            '</div>', unsafe_allow_html=True)

        st.markdown("### Address validation")
        st.markdown(
            '<div style="background:#F8F9FA;border-radius:8px;padding:16px;">'
            '<p style="font-size:13px;color:#333;">'
            'For every non-bank PDF, the app extracts the property address and '
            'compares it to the address you entered in Setup.</p>'
            '<p style="font-size:13px;">'
            '✅ <b>Matched</b> — included automatically<br>'
            '⚠️ <b>Partial</b> — included, please verify<br>'
            '❌ <b>Mismatch</b> — excluded by default<br>'
            '⚪ <b>Not found</b> — included, your call</p>'
            '<p style="font-size:13px;color:#555;">'
            'You can override any decision with the <b>Include in P&amp;L</b> checkbox.</p>'
            '</div>', unsafe_allow_html=True)

        st.markdown("### Output Excel contains")
        st.markdown(
            '<div style="background:#F8F9FA;border-radius:8px;padding:16px;'
            'font-size:13px;color:#333;">'
            '• One <b>P&amp;L tab</b> per property<br>'
            '• Monthly columns (newest → oldest) with FY &amp; CY totals<br>'
            '• <b>KPI table</b> (NOI, net yield, DSCR) per property<br>'
            '• <b>Summary tab</b> with portfolio-level dashboard<br>'
            '• Color-coded: blue = input, black = formula, yellow = FY total'
            '</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")

    if st.button("→ Start Setup", type="primary", use_container_width=True):
        st.session_state.step = 1
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: Property Setup
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 1:
    st.markdown(
        '<div class="main-header">'
        '<svg width="220" height="73" viewBox="0 0 220 74" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<rect x="10" y="40" width="9"  height="6"  rx="2" fill="#37474F"/>'
        '<rect x="21" y="30" width="9"  height="16" rx="2" fill="#5C6BC0"/>'
        '<rect x="32" y="18" width="9"  height="28" rx="2" fill="#7986CB"/>'
        '<polyline points="2,46 35,6 66,46" stroke="#FFA726" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>'
        '<polyline points="57,6 67,6 67,16" stroke="#FFA726" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>'
        '<text x="84" y="36" font-family="-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif" font-size="32" font-weight="800" fill="#FFFFFF" letter-spacing="-1">Prop</text>'
        '<text x="84" y="68" font-family="-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif" font-size="32" font-weight="300" fill="#FFA726" letter-spacing="2">folio</text>'
        '</svg>'
        '<p style="margin:8px 0 0;opacity:0.8;font-size:13px;">Upload PDFs (rental statements, bank transactions, utility bills) '
        '→ Get a fully formatted Excel P&amp;L instantly.</p>'
        '</div>',
        unsafe_allow_html=True
    )

    # ── Session loader (monthly update mode) ──────────────────────────────────
    with st.expander("📂 Load previous session for monthly update", expanded=False):
        st.markdown(
            '<div class="info-box">💡 <b>Monthly Update Mode</b>: Load your saved session JSON, '
            'then upload only the new month\'s PDFs. The app will automatically detect '
            'new months (add) vs existing months (update if changed / keep if unchanged).'
            '</div>', unsafe_allow_html=True
        )
        session_file = st.file_uploader(
            "Upload session JSON", type=['json'], key='session_uploader',
            label_visibility='collapsed',
        )
        if session_file:
            try:
                raw = json.load(session_file)
                ok, msg = _session_from_json(raw)
                if ok:
                    st.success(f"✅ {msg}")
                    if st.button("▶ Continue with this session →",
                                 type="primary", use_container_width=True):
                        st.session_state.step = 2
                        st.rerun()
                else:
                    st.error(msg)
            except Exception as e:
                st.error(f"Could not read JSON file: {e}")

    st.markdown('<div class="step-badge">STEP 1 of 4</div>', unsafe_allow_html=True)
    st.markdown("### Property Setup")

    col1, col2 = st.columns(2)
    with col1:
        n_props = st.selectbox("Number of properties", list(range(1, 11)), index=0,
                               help="Up to 10 properties. Each gets its own tab.")
        fy_start = st.selectbox(
            "Financial Year start month",
            list(MONTH_NAMES.keys()),
            format_func=lambda x: MONTH_NAMES[x],
            index=6,   # July
            help="Australian FY = July. Change if your FY starts in a different month."
        )

    with col2:
        st.markdown("**FY Period Range**")
        cy_now = 2025   # current default
        fy_c1, fy_c2 = st.columns(2)
        with fy_c1:
            fy_first = st.number_input(
                "Oldest FY start year",
                min_value=2010, max_value=2040, value=2024, step=1,
                help="Oldest year with actual data. e.g. 2024 → FY 2024-25"
            )
        with fy_c2:
            fy_last = st.number_input(
                "Template extends to FY",
                min_value=2010, max_value=2050, value=2029, step=1,
                help="Last template year. e.g. 2029 → FY 2029-30"
            )

        if fy_last < fy_first:
            st.warning("⚠️ 'Template extends to' must be ≥ 'Oldest FY start year'.")
            fy_last = fy_first

        preview_labels = make_fy_labels(int(fy_first), int(fy_last))
        st.markdown(
            f'<div class="info-box">📅 <b>FY columns</b> ({len(preview_labels)}): '
            f'{preview_labels[-1]} → {preview_labels[0]}</div>',
            unsafe_allow_html=True
        )
        st.markdown('<div class="info-box">📄 <b>Supported PDFs</b>:<br>'
                    '• Property management / ownership statements<br>'
                    '• Bank transaction statements<br>'
                    '• Utility bills (electricity, water, gas, internet)<br>'
                    '• Tax invoices (council rates, land tax, strata, insurance, trades)</div>',
                    unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Property Details")

    prop_configs = []
    for i in range(n_props):
        with st.expander(f"Property {i+1}", expanded=(i == 0)):
            c1, c2 = st.columns(2)
            with c1:
                name = st.text_input(f"Property name", key=f"name_{i}",
                                     value=f"IP#{i+1} — Property Name",
                                     placeholder="e.g. IP#1 — 3A Montfort St")
                address = st.text_input(f"Address", key=f"addr_{i}",
                                        placeholder="e.g. 3A Montfort St, Quakers Hill NSW")
                postcode = st.text_input(f"Post Code", key=f"pc_{i}",
                                         placeholder="e.g. 2763",
                                         help="4-digit Australian postcode — used as primary key for address matching")
            with c2:
                purchase_price = st.number_input(f"Purchase Price ($)", key=f"pp_{i}",
                                                 min_value=0.0, value=0.0, step=1000.0,
                                                 help="Used for yield calculations in Summary tab")
                current_value  = st.number_input(f"Current Value ($)", key=f"cv_{i}",
                                                 min_value=0.0, value=0.0, step=1000.0)
                mortgage       = st.number_input(f"Mortgage Balance ($)", key=f"mg_{i}",
                                                 min_value=0.0, value=0.0, step=1000.0)

            tab_name = f"IP#{i+1}"
            prop_configs.append({
                'name':    name or tab_name,
                'tab':     tab_name,
                'address': address,
                'postcode': postcode.strip(),
                'purchase_price': purchase_price or None,
                'current_value':  current_value  or None,
                'mortgage':       mortgage        or None,
            })

    if st.button("Next: Upload PDFs →", type="primary", use_container_width=True):
        if fy_last < fy_first:
            st.error("FY range error: 'Template extends to' must be ≥ 'Oldest FY start year'.")
            st.stop()
        st.session_state.fy_start_month = fy_start
        st.session_state.fy_labels      = make_fy_labels(int(fy_first), int(fy_last))
        st.session_state.prop_configs   = prop_configs
        # Initialize empty data containers
        st.session_state.properties = [
            {'name': p['name'], 'tab': p['tab'], 'data': {}}
            for p in prop_configs
        ]
        st.session_state.purchase_info = {
            p['tab']: {
                'address':        p['address'],
                'postcode':       p['postcode'],
                'purchase_price': p['purchase_price'],
                'current_value':  p['current_value'],
                'mortgage':       p['mortgage'],
                'purchase_date':  None,
            }
            for p in prop_configs
        }
        st.session_state.step = 2
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: Upload PDFs
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 2:
    st.markdown('<div class="step-badge">STEP 2 of 4</div>', unsafe_allow_html=True)
    st.markdown("### Upload PDFs")

    if st.session_state.get('session_loaded'):
        n_months = sum(len(p['data']) for p in st.session_state.properties)
        st.markdown(
            f'<div class="success-box">🔄 <b>Monthly Update Mode</b> — '
            f'Session loaded with {n_months} existing months of data. '
            f'Upload new PDFs and the app will <b>add new months</b> or '
            f'<b>update changed values</b> in existing months automatically.'
            f'</div>', unsafe_allow_html=True
        )

    props = st.session_state.prop_configs
    DOC_ICON = {'rental': '📋', 'bank': '🏦', 'utility': '💡', 'invoice': '🧾'}

    # ── PHASE A: Upload files ─────────────────────────────────────────────────
    if not st.session_state.get('parse_done', False):

        already_parsed = st.session_state.get('parsed_results', [])
        if already_parsed:
            st.markdown(
                f'<div class="info-box">📎 <b>{len(already_parsed)} file(s) already parsed.</b> '
                f'Upload replacements or additions below — files with the <b>same name</b> will '
                f'replace the existing result; new filenames will be appended.</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div class="info-box">📎 <b>Phase 1 of 2 — Upload</b><br>'
                'Select PDFs for each property. You can upload multiple files at once.<br>'
                'When all files are ready, click <b>Parse All PDFs</b> to process them in one go.</div>',
                unsafe_allow_html=True
            )

        total_files = 0
        for prop in props:
            tab = prop['tab']
            st.markdown(f"#### 🏠 {prop['name']}")

            st.radio(
                "Document type detection",
                ['auto', 'rental', 'bank', 'utility', 'invoice'],
                horizontal=True,
                key=f"dtype_{tab}",
                help="'auto' detects the type automatically. Use 'invoice' for council rates, land tax, strata, insurance, trades."
            )

            uploaded = st.file_uploader(
                f"Upload PDFs for {prop['name']}",
                type=['pdf'],
                accept_multiple_files=True,
                key=f"upload_{tab}",
            )
            if uploaded:
                total_files += len(uploaded)
                st.markdown(
                    f'<div style="background:#E8F5E9;border-left:3px solid #4CAF50;'
                    f'padding:6px 12px;border-radius:4px;font-size:13px;">'
                    f'📎 <b>{len(uploaded)}</b> file(s) selected</div>',
                    unsafe_allow_html=True
                )
            st.markdown("---")

        if total_files > 0:
            st.markdown(
                f'<div class="success-box">✅ <b>{total_files} file(s)</b> ready across all properties. '
                f'Click <b>Parse All PDFs</b> to process.</div>',
                unsafe_allow_html=True
            )

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("← Back", use_container_width=True):
                st.session_state.step = 1
                st.rerun()
        with col2:
            if st.button("Parse All PDFs →", type="primary",
                         use_container_width=True, disabled=(total_files == 0)):
                # Gather all files from widget state, parse in one pass
                all_to_parse = []
                for prop in props:
                    tab = prop['tab']
                    doc_type_val = st.session_state.get(f"dtype_{tab}", 'auto')
                    uploads = st.session_state.get(f"upload_{tab}") or []
                    for uf in uploads:
                        all_to_parse.append((tab, uf.name, uf.read(), doc_type_val))

                newly_parsed = []
                prog_bar = st.progress(0)
                status_txt = st.empty()
                total = len(all_to_parse)
                for idx, (tab, fname, fbytes, dt) in enumerate(all_to_parse):
                    status_txt.text(f"Parsing {fname}  ({idx + 1}/{total})…")
                    result = parse_pdf(fbytes, filename=fname, doc_type=dt)
                    result['_prop_tab'] = tab
                    result['_filename'] = fname
                    newly_parsed.append(result)
                    prog_bar.progress((idx + 1) / total)

                prog_bar.empty()
                status_txt.empty()

                # Merge: same (tab, filename) → replace; new → append
                existing_map = {
                    (r.get('_prop_tab'), r.get('_filename')): r
                    for r in st.session_state.get('parsed_results', [])
                }
                for r in newly_parsed:
                    existing_map[(r['_prop_tab'], r['_filename'])] = r
                st.session_state.parsed_results = list(existing_map.values())
                st.session_state.parse_done = True
                st.rerun()

    # ── PHASE B: Review parsed results ────────────────────────────────────────
    else:
        parsed_all = st.session_state.parsed_results
        prop_map = {p['tab']: p for p in props}

        st.markdown(
            f'<div class="success-box">✅ <b>{len(parsed_all)} file(s) parsed.</b> '
            f'Review results below, confirm address matches, then click <b>Next</b>.</div>',
            unsafe_allow_html=True
        )
        st.markdown('<div class="info-box">📋 <b>Phase 2 of 2 — Review</b><br>'
                    'Check address validation and include/exclude each file. '
                    'Files flagged ❌ are excluded by default — tick the checkbox to override.</div>',
                    unsafe_allow_html=True)

        for idx, result in enumerate(parsed_all):
            tab      = result.get('_prop_tab', '')
            fname    = result.get('_filename', f'file_{idx}.pdf')
            prop     = prop_map.get(tab, {})
            doc_icon = DOC_ICON.get(result['type'], '📄')
            yr, mo   = result.get('year'), result.get('month')
            period_str = f"{MONTH_NAMES.get(mo, '?')} {yr}" if yr and mo else "Period: Unknown"

            # Build reference address (combine address + postcode for stronger match)
            ref_addr_base = prop.get('address', '')
            postcode_val  = prop.get('postcode', '').strip()
            ref_addr = f"{ref_addr_base} {postcode_val}".strip() if postcode_val else ref_addr_base

            include_key = f"include_{tab}_{fname}"

            if result['type'] != 'bank':
                ext_addr = result.get('extracted_address', '')
                status_label, similarity, status_color = _match_address(ext_addr, ref_addr)
                if include_key not in st.session_state:
                    st.session_state[include_key] = (status_color != '#C62828')
            else:
                ext_addr, status_label, status_color = '', '', '#9E9E9E'

            with st.expander(
                f"{doc_icon} [{prop.get('name', tab)}]  {fname}  —  {period_str}",
                expanded=True
            ):
                # ── Remove button ──────────────────────────────────────────
                if st.button("🗑 Remove this file", key=f"remove_{tab}_{fname}_{idx}",
                             help="Remove this file from the parsed results"):
                    st.session_state.parsed_results = [
                        r for r in st.session_state.parsed_results
                        if not (r.get('_prop_tab') == tab and r.get('_filename') == fname)
                    ]
                    st.rerun()

                st.markdown("---")
                # ── Parse summary ──────────────────────────────────────────
                if result['type'] == 'rental':
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Money In",  f"${result.get('money_in', 0):,.2f}")
                    c2.metric("Money Out", f"${result.get('money_out', 0):,.2f}")
                    c3.metric("EFT",       f"${result.get('eft', 0):,.2f}")
                    if result.get('rooms'):
                        st.dataframe(
                            pd.DataFrame(result['rooms']).T.rename(
                                columns={'rent': 'Rent', 'mgmt': 'Mgmt Fee', 'net': 'Net'}),
                            use_container_width=True
                        )

                elif result['type'] == 'bank':
                    txns = result.get('transactions', [])
                    if txns:
                        df = pd.DataFrame(txns)[['date', 'description', 'amount', 'type', 'category']]
                        st.dataframe(df, use_container_width=True, height=200)
                        st.markdown(f"**{len(txns)} transactions** extracted")
                    else:
                        st.warning("No transactions extracted. The PDF format may need manual review.")

                elif result['type'] == 'utility':
                    c1, c2 = st.columns(2)
                    c1.metric("Utility Type", result.get('utility_type', 'Unknown'))
                    c2.metric("Amount",       f"${result.get('amount', 0):,.2f}")

                elif result['type'] == 'invoice':
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Section",      result.get('section', 'opex').upper())
                    c2.metric("P&L Category", result.get('category', 'Miscellaneous'))
                    c3.metric("Amount",       f"${result.get('amount', 0):,.2f}")

                # ── Manual period override ─────────────────────────────────
                if not (yr and mo):
                    st.warning("⚠️ Period not detected. Please confirm below.")
                    col1, col2 = st.columns(2)
                    manual_yr = col1.number_input("Year", 2020, 2035, 2025,
                                                  key=f"yr_{tab}_{fname}")
                    manual_mo = col2.selectbox("Month", list(MONTH_NAMES.keys()),
                                               format_func=lambda x: MONTH_NAMES[x],
                                               key=f"mo_{tab}_{fname}")
                    parsed_all[idx]['year']  = manual_yr
                    parsed_all[idx]['month'] = manual_mo

                # ── Address cross-validation (non-bank) ────────────────────
                if result['type'] != 'bank':
                    st.markdown("---")
                    st.markdown("**📍 Address Validation**")
                    av1, av2 = st.columns([3, 1])
                    with av1:
                        st.markdown(
                            f'<div style="border-left:4px solid {status_color};'
                            f'padding:8px 12px;border-radius:4px;background:#FAFAFA;'
                            f'margin-bottom:4px;">'
                            f'<span style="color:{status_color};font-weight:bold;">'
                            f'{status_label}</span>'
                            f'<br><span style="font-size:12px;color:#555;">'
                            f'<b>PDF:</b> {ext_addr or "—"}</span>'
                            f'<br><span style="font-size:12px;color:#555;">'
                            f'<b>Configured:</b> {ref_addr or "—"}</span>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
                    with av2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.checkbox(
                            "Include in P&L",
                            key=include_key,
                            help="Uncheck to exclude this document from the Excel output"
                        )
                else:
                    # Bank statements: always include by default
                    if include_key not in st.session_state:
                        st.session_state[include_key] = True
                    st.checkbox("Include in P&L", key=include_key)

        st.session_state.parsed_results = parsed_all

        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("← Back to Setup", use_container_width=True):
                st.session_state.step = 1
                st.rerun()
        with col2:
            if st.button("🔄 Add / Replace Files", use_container_width=True):
                # Keep existing parsed results — Phase A will merge new uploads in
                st.session_state.parse_done = False
                st.rerun()
        with col3:
            next_label = "Next: Review Data →" if parsed_all else "Skip → Review"
            if st.button(next_label, type="primary", use_container_width=True):
                _merge_parsed_to_properties()
                st.session_state.step = 3
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: Review & Edit
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 3:
    st.markdown('<div class="step-badge">STEP 3 of 4</div>', unsafe_allow_html=True)
    st.markdown("### Review & Edit Data")
    st.markdown('<div class="info-box">Review the parsed data below. '
                'Edit any values directly in the tables, or add missing months manually.</div>',
                unsafe_allow_html=True)

    # ── Merge change log (update mode only) ───────────────────────────────────
    change_log = st.session_state.get('merge_change_log', [])
    if change_log:
        n_new       = sum(1 for c in change_log if c['status'] == '🆕 New month')
        n_updated   = sum(1 for c in change_log if c['status'] == '🔄 Updated')
        n_added     = sum(1 for c in change_log if c['status'] == '➕ Added')
        n_unchanged = sum(1 for c in change_log if c['status'] == '➖ Unchanged')

        summary_color = '#2E7D32' if (n_new + n_updated + n_added) > 0 else '#555'
        st.markdown(
            f'<div style="background:#F1F8E9;border-left:4px solid {summary_color};'
            f'padding:10px 14px;border-radius:4px;margin-bottom:12px;">'
            f'<b>📋 Merge Result</b> &nbsp;'
            f'<span style="margin-right:12px;">🆕 New months: <b>{n_new}</b></span>'
            f'<span style="margin-right:12px;">🔄 Updated: <b>{n_updated}</b></span>'
            f'<span style="margin-right:12px;">➕ Added items: <b>{n_added}</b></span>'
            f'<span style="color:#888;">➖ Unchanged: <b>{n_unchanged}</b></span>'
            f'</div>', unsafe_allow_html=True
        )

        # Detailed log (excluding unchanged to keep it readable)
        detail = [c for c in change_log if c['status'] != '➖ Unchanged']
        if detail:
            with st.expander(f"🔍 View change details ({len(detail)} changes)", expanded=False):
                df_log = pd.DataFrame(detail,
                    columns=['status','property','period','item','old','new'])
                df_log.columns = ['Status','Property','Period','Item','Old Value','New Value']
                st.dataframe(df_log, use_container_width=True, hide_index=True)

    PL_ITEMS = [
        'Rental Income', 'Other Income', 'Excess Bill Shares',
        'Management Fees', 'Letting Fees',
        'Council Rates', 'Land Tax', 'Strata / Body Corporate', 'Building Insurance',
        'Maintenance & Repairs', 'Cleaning', 'Advertising', 'Miscellaneous',
        'Electricity', 'Water', 'Gas', 'Internet',
        'Mortgage Interest',
        'Cash Received (EFT)', 'Less: Utilities Paid',
        'Less: Mortgage Repayment', 'Principal Repaid',
    ]

    for prop in st.session_state.properties:
        st.markdown(f"#### 🏠 {prop['name']}")

        prop_data = prop['data']
        if not prop_data:
            st.markdown('<div class="warn-box">No data uploaded for this property. '
                        'You can add data manually below or proceed with a blank template.</div>',
                        unsafe_allow_html=True)

        # Build editable table
        all_periods = sorted(prop_data.keys())

        if all_periods:
            period_cols = [f"{MONTH_NAMES[m][:3]}-{str(y)[2:]}" for y, m in all_periods]
            table_data = {}
            for item in PL_ITEMS:
                row = [prop_data.get(p, {}).get(item, None) for p in all_periods]
                if any(v is not None and v != 0 for v in row):
                    table_data[item] = row

            if table_data:
                df = pd.DataFrame(table_data, index=period_cols).T
                df = df.fillna(0)

                # Auto-size height: show ALL rows without internal truncation.
                # Row ≈ 35px, header ≈ 38px, 16px buffer; clamp 200–800px.
                _row_px    = 35
                _header_px = 38
                _editor_h  = _header_px + len(df) * _row_px + 16
                _editor_h  = max(200, min(_editor_h, 800))

                st.caption(
                    f"📋 {len(df)} line items × {len(period_cols)} periods — "
                    "scroll left/right to see all months, up/down for all rows"
                )
                edited = st.data_editor(
                    df,
                    use_container_width=True,
                    height=_editor_h,
                    key=f"editor_{prop['tab']}",
                    num_rows="dynamic",
                )

                # Write back edits
                if edited is not None:
                    for item_name, row_vals in edited.iterrows():
                        for period_str, val in zip(all_periods, row_vals):
                            prop_data.setdefault(period_str, {})[str(item_name)] = float(val) if val else 0.0

        # Manual entry for a new month
        with st.expander(f"➕ Add / edit a month manually for {prop['name']}"):
            c1, c2 = st.columns(2)
            m_yr = c1.number_input("Year", 2020, 2035, 2025, key=f"myr_{prop['tab']}")
            m_mo = c2.selectbox("Month", list(MONTH_NAMES.keys()),
                                format_func=lambda x: MONTH_NAMES[x],
                                key=f"mmo_{prop['tab']}")
            st.markdown("**Income**")
            ri  = st.number_input("Rental Income",      0.0, key=f"ri_{prop['tab']}")
            oi  = st.number_input("Other Income",       0.0, key=f"oi_{prop['tab']}")
            eb  = st.number_input("Excess Bill Shares", 0.0, key=f"eb_{prop['tab']}")
            st.markdown("**Operating Expenses**")
            mf  = st.number_input("Management Fees",      0.0, key=f"mf_{prop['tab']}")
            lf  = st.number_input("Letting Fees",         0.0, key=f"lf_{prop['tab']}")
            cr  = st.number_input("Council Rates",        0.0, key=f"cr_{prop['tab']}")
            lt  = st.number_input("Land Tax",             0.0, key=f"lt_{prop['tab']}")
            sc  = st.number_input("Strata / Body Corp",   0.0, key=f"sc_{prop['tab']}")
            bi  = st.number_input("Building Insurance",   0.0, key=f"bi_{prop['tab']}")
            mr  = st.number_input("Maintenance & Repairs",0.0, key=f"mr_{prop['tab']}")
            st.markdown("**Utilities**")
            el  = st.number_input("Electricity",        0.0, key=f"el_{prop['tab']}")
            wa  = st.number_input("Water",              0.0, key=f"wa_{prop['tab']}")
            ga  = st.number_input("Gas",                0.0, key=f"ga_{prop['tab']}")
            inet= st.number_input("Internet",           0.0, key=f"in_{prop['tab']}")
            st.markdown("**Financing**")
            mi  = st.number_input("Mortgage Interest",  0.0, key=f"mi_{prop['tab']}")
            prin= st.number_input("Principal Repaid",   0.0, key=f"pr_{prop['tab']}")

            if st.button(f"Add {MONTH_NAMES[m_mo]} {m_yr}", key=f"add_{prop['tab']}"):
                key = (m_yr, m_mo)
                prop['data'].setdefault(key, {})
                for label, val in [
                    ('Rental Income', ri), ('Other Income', oi),
                    ('Excess Bill Shares', eb), ('Management Fees', mf),
                    ('Letting Fees', lf), ('Council Rates', cr), ('Land Tax', lt),
                    ('Strata / Body Corporate', sc), ('Building Insurance', bi),
                    ('Maintenance & Repairs', mr),
                    ('Electricity', el), ('Water', wa),
                    ('Gas', ga), ('Internet', inet),
                    ('Mortgage Interest', mi), ('Principal Repaid', prin),
                ]:
                    if val:
                        prop['data'][key][label] = val
                st.success(f"Added data for {MONTH_NAMES[m_mo]} {m_yr}")
                st.rerun()

        st.markdown("---")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    with col2:
        if st.button("Next: Generate Excel →", type="primary", use_container_width=True):
            st.session_state.step = 4
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: Generate Excel
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 4:
    st.markdown('<div class="step-badge">STEP 4 of 4</div>', unsafe_allow_html=True)
    st.markdown("### Generate Excel")

    # KPI preview
    st.markdown("#### 📊 Data Summary")
    summary_rows = []
    for prop in st.session_state.properties:
        total_months = len(prop['data'])
        total_income = sum(
            v.get('Rental Income', 0) + v.get('Other Income', 0)
            for v in prop['data'].values()
        )
        total_opex = sum(
            v.get('Management Fees', 0) + v.get('Letting Fees', 0) +
            v.get('Maintenance & Repairs', 0)
            for v in prop['data'].values()
        )
        noi = total_income - total_opex
        summary_rows.append({
            'Property':       prop['name'],
            'Months of Data': total_months,
            'Total Income':   f"${total_income:,.0f}",
            'Total OpEx':     f"${total_opex:,.0f}",
            'NOI (Est.)':     f"${noi:,.0f}",
        })

    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### ⚙️ Generation Settings")
    c1, c2 = st.columns(2)
    with c1:
        output_name = st.text_input("Output filename", value="Property_PL_Portfolio.xlsx")
    with c2:
        st.markdown('<div class="info-box">Tabs: One per property + Summary tab<br>'
                    'Columns: FY 2024-25 → FY 2029-30 + CY 2024–2030<br>'
                    'Format: Color-coded, formula-linked, outline grouped</div>',
                    unsafe_allow_html=True)

    if st.button("🚀 Generate Excel", type="primary", use_container_width=True):
        with st.spinner("Building Excel… (this may take 10–30 seconds)"):
            try:
                xlsx_bytes = build_workbook(
                    properties    = st.session_state.properties,
                    fy_start_month= st.session_state.fy_start_month,
                    fy_labels     = st.session_state.fy_labels,
                    purchase_info = st.session_state.get('purchase_info', {}),
                )
                st.session_state['xlsx_bytes'] = xlsx_bytes
                st.session_state['output_name'] = output_name
            except Exception as e:
                st.error(f"Error generating Excel: {e}")
                import traceback
                st.code(traceback.format_exc())

    if 'xlsx_bytes' in st.session_state:
        st.markdown('<div class="success-box">✅ Excel generated successfully!</div>',
                    unsafe_allow_html=True)

        dl_col, save_col = st.columns(2)
        with dl_col:
            st.download_button(
                label="⬇️  Download Excel",
                data=st.session_state['xlsx_bytes'],
                file_name=st.session_state.get('output_name', 'Property_PL.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        with save_col:
            session_json = _session_to_json()
            save_filename = (
                f"property_pl_session_"
                f"{datetime.datetime.now().strftime('%Y%m')}.json"
            )
            st.download_button(
                label="💾 Save Session (for next month)",
                data=session_json,
                file_name=save_filename,
                mime="application/json",
                use_container_width=True,
                help="Save all property configs + data as JSON. "
                     "Load it next month to add new PDFs on top of existing data.",
            )

        st.markdown(
            '<div class="info-box">💡 <b>Next month workflow</b>: '
            'Click "Save Session" → next month, go to Step 1 → "Load previous session" '
            '→ upload only the new month\'s PDFs → the app handles add/update automatically.'
            '</div>', unsafe_allow_html=True
        )

        st.markdown("---")
        st.markdown("#### What's in the Excel?")
        for prop in st.session_state.properties:
            st.markdown(f"- **{prop['tab']}**: {prop['name']} — P&L + KPI Table A")
        st.markdown("- **Summary**: Table B (Asset Info & Yields) + Table A (Portfolio Aggregates)")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Back to Review", use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    with col2:
        if st.button("🏠 Start New", use_container_width=True):
            for k in ['step','properties','parsed_results','xlsx_bytes']:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()
