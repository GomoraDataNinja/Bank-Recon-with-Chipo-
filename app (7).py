import streamlit as st
import pandas as pd
import numpy as np
import re
import tempfile
from datetime import datetime, date, timezone
from dateutil.parser import parse as dt_parse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from copy import copy as pycopy
import hashlib
import time
import os
import warnings

warnings.filterwarnings("ignore")

# =========================
# App config & security
# =========================
APP_VERSION = "2.0.0"
APP_NAME = "Bank Reconciliation"
DEPLOYMENT_MODE = os.environ.get("DEPLOYMENT_MODE", "production")
SESSION_TIMEOUT_MINUTES = 60

st.set_page_config(
    page_title=f"{APP_NAME} v{APP_VERSION}",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

def safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

def get_org_password():
    env_pw = os.environ.get("APP_PASSWORD", "").strip()
    if env_pw:
        return env_pw
    try:
        sec_pw = str(st.secrets.get("app_password", "")).strip()
        if sec_pw:
            return sec_pw
    except Exception:
        pass
    return "recon2024"   # change before production

ORG_PASSWORD = get_org_password()

# =========================
# Theme (Wells Fargo red accent)
# =========================
THEME = {
    "bg": "#ffffff",
    "panel": "#ffffff",
    "panel2": "#f7f7f7",
    "text": "#111111",
    "muted": "#5b5b5b",
    "border": "rgba(0,0,0,0.10)",
    "border2": "rgba(0,0,0,0.14)",
    "accent": "#D71E28",
    "accent2": "#b5161f",
    "good": "#168a45",
    "bad": "#d11a2a",
    "neutral": "#6b7280",
}

def apply_style():
    st.markdown(
        f"""
        <style>
        :root {{
            --bg: {THEME['bg']};
            --panel: {THEME['panel']};
            --panel2: {THEME['panel2']};
            --text: {THEME['text']};
            --muted: {THEME['muted']};
            --border: {THEME['border']};
            --border2: {THEME['border2']};
            --accent: {THEME['accent']};
            --accent2: {THEME['accent2']};
            --good: {THEME['good']};
            --bad: {THEME['bad']};
            --neutral: {THEME['neutral']};
        }}
        html {{
            color-scheme: light !important;
        }}
        html, body, [data-testid="stAppViewContainer"], .stApp {{
            background: var(--bg) !important;
            color: var(--text) !important;
        }}
        [data-testid="stHeader"], [data-testid="stToolbar"], #MainMenu, footer {{
            display: none !important;
        }}
        .block-container {{
            max-width: 1400px;
            padding-top: 2.6rem !important;
            padding-bottom: 2.2rem !important;
        }}
        /* File uploader styling */
        div[data-testid="stFileUploader"] {{
            background: var(--panel) !important;
            border: 1px dashed var(--border2) !important;
            border-radius: 16px !important;
            padding: 10px !important;
            transition: border 0.2s ease;
        }}
        div[data-testid="stFileUploader"]:hover {{
            border: 1px dashed var(--accent) !important;
        }}
        div[data-testid="stFileUploader"] label {{
            color: var(--text) !important;
            font-weight: 800 !important;
            font-size: 14px !important;
        }}
        div[data-testid="stFileUploader"] small {{
            color: var(--muted) !important;
            font-size: 12px !important;
        }}
        /* Cards and containers */
        .card {{
            background: #ffffff !important;
            border: 1px solid var(--border) !important;
            border-radius: 18px !important;
            padding: 18px 18px !important;
        }}
        .hero {{
            border: 1px solid var(--border) !important;
            border-radius: 22px !important;
            padding: 26px 22px !important;
            background: radial-gradient(900px 260px at 50% -10%, rgba(215,30,40,0.10), transparent 60%),
                        linear-gradient(180deg, #ffffff, #ffffff) !important;
        }}
        .title {{
            font-size: 30px !important;
            font-weight: 800 !important;
            letter-spacing: 0.2px !important;
            margin: 0 !important;
        }}
        .subtitle {{
            margin-top: 8px !important;
            color: var(--muted) !important;
            font-size: 14px !important;
        }}
        .chip {{
            display: inline-flex !important;
            align-items: center !important;
            gap: 8px !important;
            padding: 6px 12px !important;
            border-radius: 999px !important;
            border: 1px solid var(--border) !important;
            background: #ffffff !important;
            font-size: 12px !important;
            font-weight: 650 !important;
            color: var(--muted) !important;
        }}
        .chip-dot {{
            width: 8px !important;
            height: 8px !important;
            border-radius: 999px !important;
            display: inline-block !important;
            background: var(--accent) !important;
        }}
        .metric {{
            border: 1px solid var(--border) !important;
            border-radius: 18px !important;
            padding: 14px 14px !important;
            background: #ffffff !important;
        }}
        .metric-k {{
            font-size: 12px !important;
            color: var(--muted) !important;
            font-weight: 700 !important;
            text-transform: uppercase !important;
        }}
        .metric-v {{
            font-size: 26px !important;
            font-weight: 850 !important;
            margin-top: 6px !important;
            color: var(--accent) !important;
        }}
        /* Buttons */
        div.stButton > button {{
            background: var(--accent) !important;
            border: 1px solid var(--accent) !important;
            border-radius: 14px !important;
            padding: 0.7rem 1rem !important;
            font-weight: 750 !important;
            color: #ffffff !important;
        }}
        div.stButton > button:hover {{
            background: var(--accent2) !important;
            border-color: var(--accent2) !important;
        }}
        div[data-baseweb="base-input"] > div,
        div[data-baseweb="input"] > div {{
            background: #ffffff !important;
            border: 1px solid var(--border2) !important;
            border-radius: 14px !important;
        }}
        [data-testid="stDataFrame"] {{
            background: #ffffff !important;
            border: 1px solid var(--border) !important;
            border-radius: 16px !important;
            overflow: hidden !important;
        }}
        a, a:visited {{
            color: var(--accent) !important;
            font-weight: 750 !important;
        }}
        .topbar {{
            background: var(--accent) !important;
            color: white;
            padding: 16px 18px;
            border-radius: 14px;
            margin-bottom: 16px;
            position: relative;
        }}
        .brand-center {{
            text-align: center;
            font-size: 30px;
            font-weight: 900;
        }}
        .sub-center {{
            text-align: center;
            font-size: 12px;
            opacity: 0.98;
            margin-top: 2px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

apply_style()

# =========================
# Session management
# =========================
def touch():
    st.session_state.last_activity = datetime.now()

def is_timed_out():
    last = st.session_state.get("last_activity")
    if not last:
        return False
    return (datetime.now() - last).total_seconds() > SESSION_TIMEOUT_MINUTES * 60

def logout():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    safe_rerun()

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "session_id" not in st.session_state:
    st.session_state.session_id = hashlib.sha256(str(time.time()).encode()).hexdigest()[:16]
if "last_activity" not in st.session_state:
    st.session_state.last_activity = datetime.now()

def login_screen():
    st.markdown('<div style="height: 1.8rem;"></div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.25, 1])
    with c2:
        st.markdown(
            f"""
            <div class="card" style="margin-top: 10vh;">
                <div class="title" style="text-align:center;">{APP_NAME}</div>
                <div class="subtitle" style="text-align:center;">Sign in to continue.</div>
                <div style="height: 14px;"></div>
                <div style="display:flex; justify-content:center;">
                    <div class="chip"><span class="chip-dot"></span> Version {APP_VERSION} • {DEPLOYMENT_MODE.title()}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form", clear_on_submit=True):
            pw = st.text_input("Password", type="password", placeholder="Organisation password")
            ok = st.form_submit_button("Sign in", use_container_width=True)
        if ok:
            if pw == ORG_PASSWORD:
                st.session_state.authenticated = True
                touch()
                safe_rerun()
            else:
                st.error("Wrong password.")

if st.session_state.authenticated and is_timed_out():
    st.session_state.authenticated = False
    st.warning("Session timed out. Sign in again.")
    login_screen()
    st.stop()

if not st.session_state.authenticated:
    login_screen()
    st.stop()

touch()

# =========================
# Main app header (after login)
# =========================
st.markdown(
    f"""
    <div class="hero" style="text-align:center;">
        <div class="title">{APP_NAME}</div>
        <div class="subtitle">Upload Bank Statement and Ledger, match transactions, download reconciliation report.</div>
        <div style="height: 12px;"></div>
        <div style="display:flex; justify-content:center; gap:10px; flex-wrap:wrap;">
            <div class="chip"><span class="chip-dot"></span> Secure session</div>
            <div class="chip">Session {st.session_state.session_id}</div>
            <div class="chip">Mode {DEPLOYMENT_MODE.title()}</div>
            <div class="chip">Version {APP_VERSION}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("")

# =========================
# Helper functions (reconciliation logic)
# =========================
def to_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def to_num(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = to_str(x).replace(",", "")
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '').strip()
    s = re.sub(r"[^\d\.\-]", "", s)
    if not s or s == '-':
        return np.nan
    try:
        return float(s)
    except Exception:
        return np.nan

def to_date(x):
    if isinstance(x, (pd.Timestamp, np.datetime64)):
        return pd.to_datetime(x, errors='coerce')
    s = to_str(x)
    if not s:
        return pd.NaT
    try:
        for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%y', '%d-%m-%y']:
            try:
                return pd.to_datetime(s, format=fmt)
            except:
                continue
        return pd.to_datetime(dt_parse(s, fuzzy=True))
    except Exception:
        return pd.NaT

def format_date(date_val):
    if date_val is None:
        return ''
    if pd.isna(date_val):
        return ''
    if isinstance(date_val, str):
        if date_val == '':
            return ''
        try:
            parsed = pd.to_datetime(date_val, errors='coerce')
            if pd.notna(parsed):
                return parsed.strftime('%d/%m/%y')
            else:
                return date_val[:10] if len(date_val) > 10 else date_val
        except:
            return date_val[:10] if len(date_val) > 10 else date_val
    if isinstance(date_val, (pd.Timestamp, datetime, date)):
        if pd.isna(date_val):
            return ''
        try:
            return date_val.strftime('%d/%m/%y')
        except:
            return ''
    return str(date_val)

def detect_table(df_raw, max_rows=100):
    if df_raw is None or df_raw.empty:
        return None, 0
    best_row = 0
    best_score = 0
    for r in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[r]
        non_empty = row.notna().sum()
        if non_empty < 2:
            continue
        header_score = 0
        for cell in row:
            if pd.isna(cell):
                continue
            s = str(cell).lower()
            if any(kw in s for kw in ['date', 'amount', 'credit', 'debit', 'ref', 'description', 'posting', 'document', 'narration']):
                header_score += 2
        score = non_empty + header_score
        if score > best_score:
            best_score = score
            best_row = r
    if best_row >= 0 and best_row < len(df_raw):
        headers = df_raw.iloc[best_row].fillna('').astype(str).tolist()
        headers = [h.strip() if h.strip() else f"Column_{i}" for i, h in enumerate(headers)]
        data = df_raw.iloc[best_row + 1:].reset_index(drop=True)
        data = data.iloc[:, :len(headers)]
        data.columns = headers
        data = data.dropna(how='all')
        return data, best_row
    return df_raw, 0

def load_bank_statement(file):
    try:
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        visible_sheets = [s for s in sheet_names if not s.startswith('_') and 'hidden' not in s.lower()]
        if not visible_sheets:
            visible_sheets = sheet_names
        df_bank = None
        header_row_used = 0
        for sheet in visible_sheets:
            try:
                df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
                if df_raw is not None and not df_raw.empty:
                    df, header_row = detect_table(df_raw)
                    if df is not None and len(df) > 3:
                        df_bank = df
                        header_row_used = header_row
                        break
            except Exception:
                continue
        if df_bank is None or df_bank.empty:
            df_bank = pd.read_excel(file)
            if df_bank.empty:
                raise ValueError("Could not read bank statement file")
    except Exception:
        df_bank = pd.read_excel(file)

    date_col = None; credit_col = None; debit_col = None; ref_col = None; desc_col = None; balance_col = None
    for col in df_bank.columns:
        cl = str(col).lower()
        if 'date' in cl and not date_col: date_col = col
        if 'credit' in cl and not credit_col: credit_col = col
        if 'debit' in cl and not debit_col: debit_col = col
        if 'ref' in cl or 'reference' in cl: ref_col = col
        if 'narrative' in cl or 'description' in cl or 'particulars' in cl: desc_col = col
        if 'balance' in cl and not balance_col: balance_col = col
    if not credit_col and not debit_col:
        for col in df_bank.columns:
            if 'amount' in str(col).lower() or 'value' in str(col).lower():
                credit_col = col; debit_col = col; break

    opening_balance = None
    for idx, row in df_bank.iterrows():
        row_str = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
        if 'balance at period start' in row_str or 'opening balance' in row_str or 'balance brought forward' in row_str:
            for col in df_bank.columns:
                val = to_num(row[col])
                if not pd.isna(val) and val != 0:
                    opening_balance = val
                    break
            break

    transactions = []
    for idx, row in df_bank.iterrows():
        credit = 0; debit = 0
        if credit_col and credit_col in df_bank.columns:
            credit = to_num(row[credit_col]) if pd.notna(row[credit_col]) else 0
        if debit_col and debit_col in df_bank.columns:
            debit = to_num(row[debit_col]) if pd.notna(row[debit_col]) else 0
        if credit == 0 and debit == 0 and credit_col == debit_col and credit_col:
            amount = to_num(row[credit_col]) if pd.notna(row[credit_col]) else 0
            if amount == 0: continue
            credit = amount if amount > 0 else 0
            debit = -amount if amount < 0 else 0
        if credit == 0 and debit == 0: continue
        trans_date = to_date(row[date_col]) if date_col and date_col in df_bank.columns else pd.NaT
        if pd.isna(trans_date): continue
        desc = ''
        if desc_col and desc_col in df_bank.columns: desc = to_str(row[desc_col])
        if not desc and ref_col and ref_col in df_bank.columns: desc = to_str(row[ref_col])
        transactions.append({
            'date': trans_date,
            'reference': to_str(row[ref_col]) if ref_col and ref_col in df_bank.columns else '',
            'description': desc,
            'credit': credit,
            'debit': debit,
            'amount': credit - debit,
            'abs_amount': abs(credit - debit),
            'source': 'BANK'
        })
    df_bank_norm = pd.DataFrame(transactions)
    if df_bank_norm.empty:
        st.error("No valid bank transactions.")
        st.stop()
    closing_balance = None
    if balance_col and balance_col in df_bank.columns:
        last_valid = df_bank[balance_col].dropna().iloc[-1] if not df_bank[balance_col].dropna().empty else None
        closing_balance = to_num(last_valid) if last_valid else None
    return df_bank_norm, opening_balance, closing_balance, header_row_used

def load_ledger(file):
    try:
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        visible_sheets = [s for s in sheet_names if not s.startswith('_') and 'hidden' not in s.lower()]
        if not visible_sheets:
            visible_sheets = sheet_names
        df_ledger = None
        header_row_used = 0
        for sheet in visible_sheets:
            try:
                df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
                if df_raw is not None and not df_raw.empty:
                    df, header_row = detect_table(df_raw)
                    if df is not None and len(df) > 3:
                        df_ledger = df
                        header_row_used = header_row
                        break
            except Exception:
                continue
        if df_ledger is None or df_ledger.empty:
            df_ledger = pd.read_excel(file)
            if df_ledger.empty:
                raise ValueError("Could not read ledger file")
    except Exception:
        df_ledger = pd.read_excel(file)

    date_col = None; amount_col = None; desc_col = None; ref_col = None
    for col in df_ledger.columns:
        cl = str(col).lower()
        if 'date' in cl and not date_col: date_col = col
        if 'amount' in cl and not amount_col: amount_col = col
        if 'description' in cl or 'desc' in cl or 'particulars' in cl: desc_col = col
        if 'document' in cl or 'ref' in cl or 'external' in cl: ref_col = col

    transactions = []
    for idx, row in df_ledger.iterrows():
        amount = to_num(row[amount_col]) if amount_col and amount_col in df_ledger.columns else np.nan
        if pd.isna(amount): continue
        trans_date = to_date(row[date_col]) if date_col and date_col in df_ledger.columns else pd.NaT
        if pd.isna(trans_date): continue
        transactions.append({
            'date': trans_date,
            'reference': to_str(row[ref_col]) if ref_col and ref_col in df_ledger.columns else '',
            'description': to_str(row[desc_col]) if desc_col and desc_col in df_ledger.columns else '',
            'amount': amount,
            'abs_amount': abs(amount),
            'type': 'CREDIT' if amount > 0 else 'DEBIT',
            'source': 'LEDGER'
        })
    df_ledger_norm = pd.DataFrame(transactions)
    if df_ledger_norm.empty:
        st.error("No valid ledger transactions.")
        st.stop()
    closing_balance = df_ledger_norm['amount'].sum() if not df_ledger_norm.empty else 0
    return df_ledger_norm, closing_balance, header_row_used

def match_transactions(bank_df, ledger_df):
    bank_copy = bank_df.copy()
    ledger_copy = ledger_df.copy()
    bank_credits = bank_copy[bank_copy['credit'] > 0].copy() if 'credit' in bank_copy.columns else bank_copy[bank_copy['amount'] > 0].copy()
    bank_debits = bank_copy[bank_copy['debit'] > 0].copy() if 'debit' in bank_copy.columns else bank_copy[bank_copy['amount'] < 0].copy()
    if 'credit' not in bank_copy.columns:
        bank_credits = bank_copy[bank_copy['amount'] > 0].copy()
        bank_debits = bank_copy[bank_copy['amount'] < 0].copy()
    ledger_credits = ledger_copy[ledger_copy['amount'] > 0].copy()
    ledger_debits = ledger_copy[ledger_copy['amount'] < 0].copy()
    bank_credits['id'] = [f'B_C_{i}' for i in range(len(bank_credits))]
    bank_debits['id'] = [f'B_D_{i}' for i in range(len(bank_debits))]
    ledger_credits['id'] = [f'L_C_{i}' for i in range(len(ledger_credits))]
    ledger_debits['id'] = [f'L_D_{i}' for i in range(len(ledger_debits))]
    matches = []
    # Credits
    matched_ledger_credit_ids = set(); matched_bank_credit_ids = set()
    all_amounts = set()
    if not ledger_credits.empty: all_amounts.update(ledger_credits['abs_amount'].unique())
    if not bank_credits.empty: all_amounts.update(bank_credits['abs_amount'].unique())
    for amount in all_amounts:
        ledger_items = ledger_credits[ledger_credits['abs_amount'] == amount] if not ledger_credits.empty else pd.DataFrame()
        bank_items = bank_credits[bank_credits['abs_amount'] == amount] if not bank_credits.empty else pd.DataFrame()
        if len(ledger_items) == 1 and len(bank_items) == 1:
            matches.append({
                'ledger_id': ledger_items.iloc[0]['id'], 'bank_id': bank_items.iloc[0]['id'],
                'amount': amount, 'type': 'CREDIT', 'match_method': 'amount_only',
                'ledger_date': ledger_items.iloc[0]['date'], 'bank_date': bank_items.iloc[0]['date'],
                'date_match': ledger_items.iloc[0]['date'] == bank_items.iloc[0]['date'],
                'ledger_desc': ledger_items.iloc[0]['description'], 'ledger_ref': ledger_items.iloc[0]['reference'],
                'bank_desc': bank_items.iloc[0]['description'], 'bank_ref': bank_items.iloc[0]['reference']
            })
            matched_ledger_credit_ids.add(ledger_items.iloc[0]['id']); matched_bank_credit_ids.add(bank_items.iloc[0]['id'])
        elif len(ledger_items) > 0 and len(bank_items) > 0:
            ledger_list = ledger_items.to_dict('records'); bank_list = bank_items.to_dict('records')
            for l_item in ledger_list:
                if l_item['id'] in matched_ledger_credit_ids: continue
                matched_bank = None
                for b_item in bank_list:
                    if b_item['id'] in matched_bank_credit_ids: continue
                    if l_item['date'].date() == b_item['date'].date():
                        matched_bank = b_item; break
                if matched_bank:
                    matches.append({
                        'ledger_id': l_item['id'], 'bank_id': matched_bank['id'],
                        'amount': amount, 'type': 'CREDIT', 'match_method': 'amount_and_date',
                        'ledger_date': l_item['date'], 'bank_date': matched_bank['date'],
                        'date_match': True, 'ledger_desc': l_item['description'], 'ledger_ref': l_item['reference'],
                        'bank_desc': matched_bank['description'], 'bank_ref': matched_bank['reference']
                    })
                    matched_ledger_credit_ids.add(l_item['id']); matched_bank_credit_ids.add(matched_bank['id'])
    unmatched_ledger_credits = [item.to_dict() for _, item in ledger_credits.iterrows() if item['id'] not in matched_ledger_credit_ids]
    unmatched_bank_credits = [item.to_dict() for _, item in bank_credits.iterrows() if item['id'] not in matched_bank_credit_ids]
    # Debits
    matched_ledger_debit_ids = set(); matched_bank_debit_ids = set()
    all_debit_amounts = set()
    if not ledger_debits.empty: all_debit_amounts.update(ledger_debits['abs_amount'].unique())
    if not bank_debits.empty: all_debit_amounts.update(bank_debits['abs_amount'].unique())
    for amount in all_debit_amounts:
        ledger_items = ledger_debits[ledger_debits['abs_amount'] == amount] if not ledger_debits.empty else pd.DataFrame()
        bank_items = bank_debits[bank_debits['abs_amount'] == amount] if not bank_debits.empty else pd.DataFrame()
        if len(ledger_items) == 1 and len(bank_items) == 1:
            matches.append({
                'ledger_id': ledger_items.iloc[0]['id'], 'bank_id': bank_items.iloc[0]['id'],
                'amount': amount, 'type': 'DEBIT', 'match_method': 'amount_only',
                'ledger_date': ledger_items.iloc[0]['date'], 'bank_date': bank_items.iloc[0]['date'],
                'date_match': ledger_items.iloc[0]['date'] == bank_items.iloc[0]['date'],
                'ledger_desc': ledger_items.iloc[0]['description'], 'ledger_ref': ledger_items.iloc[0]['reference'],
                'bank_desc': bank_items.iloc[0]['description'], 'bank_ref': bank_items.iloc[0]['reference']
            })
            matched_ledger_debit_ids.add(ledger_items.iloc[0]['id']); matched_bank_debit_ids.add(bank_items.iloc[0]['id'])
        elif len(ledger_items) > 0 and len(bank_items) > 0:
            ledger_list = ledger_items.to_dict('records'); bank_list = bank_items.to_dict('records')
            for l_item in ledger_list:
                if l_item['id'] in matched_ledger_debit_ids: continue
                matched_bank = None
                for b_item in bank_list:
                    if b_item['id'] in matched_bank_debit_ids: continue
                    if l_item['date'].date() == b_item['date'].date():
                        matched_bank = b_item; break
                if matched_bank:
                    matches.append({
                        'ledger_id': l_item['id'], 'bank_id': matched_bank['id'],
                        'amount': amount, 'type': 'DEBIT', 'match_method': 'amount_and_date',
                        'ledger_date': l_item['date'], 'bank_date': matched_bank['date'],
                        'date_match': True, 'ledger_desc': l_item['description'], 'ledger_ref': l_item['reference'],
                        'bank_desc': matched_bank['description'], 'bank_ref': matched_bank['reference']
                    })
                    matched_ledger_debit_ids.add(l_item['id']); matched_bank_debit_ids.add(matched_bank['id'])
    unmatched_ledger_debits = [item.to_dict() for _, item in ledger_debits.iterrows() if item['id'] not in matched_ledger_debit_ids]
    unmatched_bank_debits = [item.to_dict() for _, item in bank_debits.iterrows() if item['id'] not in matched_bank_debit_ids]
    return {
        'matches': matches,
        'unmatched_ledger_credits': unmatched_ledger_credits,
        'unmatched_ledger_debits': unmatched_ledger_debits,
        'unmatched_bank_credits': unmatched_bank_credits,
        'unmatched_bank_debits': unmatched_bank_debits
    }

def build_working_paper(match_results):
    rows = []
    for match in match_results['matches']:
        rows.append({
            'SECTION': 'MATCHED',
            'LEDGER_DATE': match['ledger_date'],
            'LEDGER_DESC': match.get('ledger_desc',''),
            'LEDGER_REF': match.get('ledger_ref',''),
            'LEDGER_AMOUNT': match['amount'] if match['type']=='CREDIT' else -match['amount'],
            'MATCH_STATUS': f"MATCHED - {match['match_method'].replace('_',' ').upper()}",
            'BANK_AMOUNT': match['amount'] if match['type']=='CREDIT' else -match['amount'],
            'BANK_DATE': match['bank_date'],
            'BANK_REF': match.get('bank_ref',''),
            'BANK_DESC': match.get('bank_desc','')
        })
    for item in match_results['unmatched_ledger_credits']:
        rows.append({
            'SECTION': 'UNMATCHED - LEDGER ONLY',
            'LEDGER_DATE': item['date'], 'LEDGER_DESC': item['description'], 'LEDGER_REF': item['reference'],
            'LEDGER_AMOUNT': item['amount'], 'MATCH_STATUS': 'NO BANK MATCH',
            'BANK_AMOUNT': '', 'BANK_DATE': '', 'BANK_REF': '', 'BANK_DESC': ''
        })
    for item in match_results['unmatched_ledger_debits']:
        rows.append({
            'SECTION': 'UNMATCHED - LEDGER ONLY',
            'LEDGER_DATE': item['date'], 'LEDGER_DESC': item['description'], 'LEDGER_REF': item['reference'],
            'LEDGER_AMOUNT': item['amount'], 'MATCH_STATUS': 'NO BANK MATCH',
            'BANK_AMOUNT': '', 'BANK_DATE': '', 'BANK_REF': '', 'BANK_DESC': ''
        })
    for item in match_results['unmatched_bank_credits']:
        rows.append({
            'SECTION': 'UNMATCHED - BANK ONLY',
            'LEDGER_DATE': '', 'LEDGER_DESC': '', 'LEDGER_REF': '', 'LEDGER_AMOUNT': '',
            'MATCH_STATUS': 'NO LEDGER MATCH',
            'BANK_AMOUNT': item['amount'], 'BANK_DATE': item['date'], 'BANK_REF': item['reference'], 'BANK_DESC': item['description']
        })
    for item in match_results['unmatched_bank_debits']:
        rows.append({
            'SECTION': 'UNMATCHED - BANK ONLY',
            'LEDGER_DATE': '', 'LEDGER_DESC': '', 'LEDGER_REF': '', 'LEDGER_AMOUNT': '',
            'MATCH_STATUS': 'NO LEDGER MATCH',
            'BANK_AMOUNT': item['amount'], 'BANK_DATE': item['date'], 'BANK_REF': item['reference'], 'BANK_DESC': item['description']
        })
    return pd.DataFrame(rows)

def build_recon_statement(opening_balance, ledger_closing_balance, match_results):
    recon_items = []
    total_adjustment = 0
    for item in match_results['unmatched_bank_credits']:
        recon_items.append({'date': item['date'], 'description': f"BANK ONLY: {item['description']} ({item['reference']})", 'adjustment': item['amount']})
        total_adjustment += item['amount']
    for item in match_results['unmatched_bank_debits']:
        recon_items.append({'date': item['date'], 'description': f"BANK ONLY: {item['description']} ({item['reference']})", 'adjustment': item['amount']})
        total_adjustment += item['amount']
    for item in match_results['unmatched_ledger_credits']:
        recon_items.append({'date': item['date'], 'description': f"LEDGER ONLY - DEPOSIT IN TRANSIT: {item['description']} ({item['reference']})", 'adjustment': item['amount']})
        total_adjustment += item['amount']
    for item in match_results['unmatched_ledger_debits']:
        recon_items.append({'date': item['date'], 'description': f"LEDGER ONLY - UNPRESENTED CHEQUE: {item['description']} ({item['reference']})", 'adjustment': item['amount']})
        total_adjustment += item['amount']
    adjusted_balance = (opening_balance or 0) + total_adjustment
    difference = adjusted_balance - (ledger_closing_balance or 0)
    return {
        'opening_balance': opening_balance or 0,
        'recon_items': pd.DataFrame(recon_items),
        'total_adjustment': total_adjustment,
        'adjusted_balance': adjusted_balance,
        'ledger_balance': ledger_closing_balance or 0,
        'difference': difference
    }

def export_to_excel(working_paper_df, recon_statement, bank_df, ledger_df, match_results):
    output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        wp_display = working_paper_df.copy()
        wp_display['LEDGER_DATE'] = wp_display['LEDGER_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_display['BANK_DATE'] = wp_display['BANK_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_display['LEDGER_AMOUNT'] = wp_display['LEDGER_AMOUNT'].apply(lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' else '')
        wp_display['BANK_AMOUNT'] = wp_display['BANK_AMOUNT'].apply(lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' else '')
        wp_display.to_excel(writer, sheet_name='WORKING_PAPER', index=False)
        recon_data = []
        recon_data.append(['Bank Reconciliation Statement', ''])
        recon_data.append(['', ''])
        recon_data.append(['Statement Balance (from bank)', '', f"{recon_statement['opening_balance']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['RECONCILIATION ITEMS:', '', ''])
        recon_data.append(['Date', 'Description', 'Amount'])
        recon_data.append(['', '', ''])
        for _, item in recon_statement['recon_items'].iterrows():
            date_str = format_date(item['date']) if pd.notna(item['date']) else ''
            recon_data.append([date_str, item['description'][:80], f"{item['adjustment']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Total Reconciling Items', f"{recon_statement['total_adjustment']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Total Cash (Adjusted Balance)', f"{recon_statement['adjusted_balance']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Ledger Balance (from Yellowcob)', f"{recon_statement['ledger_balance']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Difference', f"{recon_statement['difference']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', '', ''])
        recon_data.append(['PREPARED BY:', '', ''])
        recon_data.append(['Name:', '', ''])
        recon_data.append(['Signature:', '', ''])
        recon_data.append(['Date:', '', datetime.now().strftime('%d/%m/%y')])
        recon_df = pd.DataFrame(recon_data)
        recon_df.to_excel(writer, sheet_name='RECON_STATEMENT', index=False, header=False)

        if match_results['matches']:
            matched_detail = []
            matched_detail.append(['Matched Transactions - Detailed View', '', '', '', ''])
            matched_detail.append(['', '', '', '', ''])
            matched_detail.append(['Ledger Date', 'Ledger Description', 'Ledger Ref', 'Ledger Amount', 'Bank Date', 'Bank Ref', 'Bank Description', 'Bank Amount', 'Match Method'])
            for match in match_results['matches']:
                matched_detail.append([
                    format_date(match['ledger_date']),
                    (match.get('ledger_desc', '')[:50] if match.get('ledger_desc') else ''),
                    (match.get('ledger_ref', '')[:30] if match.get('ledger_ref') else ''),
                    f"{match['amount']:,.2f}",
                    format_date(match['bank_date']),
                    (match.get('bank_ref', '')[:30] if match.get('bank_ref') else ''),
                    (match.get('bank_desc', '')[:50] if match.get('bank_desc') else ''),
                    f"{match['amount']:,.2f}",
                    match['match_method']
                ])
            pd.DataFrame(matched_detail).to_excel(writer, sheet_name='MATCHED_DETAIL', index=False, header=False)

        if match_results['unmatched_ledger_credits'] or match_results['unmatched_ledger_debits']:
            uml = []
            uml.append(['Ledger Transactions with No Bank Match', '', '', '']); uml.append(['', '', '', '']); uml.append(['Date', 'Description', 'Reference', 'Amount', 'Type'])
            for item in match_results['unmatched_ledger_credits']:
                uml.append([format_date(item['date']), item['description'][:60], item['reference'][:30], f"{item['amount']:,.2f}", 'CREDIT (Deposit)'])
            for item in match_results['unmatched_ledger_debits']:
                uml.append([format_date(item['date']), item['description'][:60], item['reference'][:30], f"{item['amount']:,.2f}", 'DEBIT (Payment)'])
            pd.DataFrame(uml).to_excel(writer, sheet_name='UNMATCHED_LEDGER', index=False, header=False)

        if match_results['unmatched_bank_credits'] or match_results['unmatched_bank_debits']:
            umb = []
            umb.append(['Bank Transactions with No Ledger Match', '', '', '']); umb.append(['', '', '', '']); umb.append(['Date', 'Description', 'Reference', 'Amount', 'Type'])
            for item in match_results['unmatched_bank_credits']:
                umb.append([format_date(item['date']), item['description'][:60], item['reference'][:30], f"{item['amount']:,.2f}", 'CREDIT (Deposit)'])
            for item in match_results['unmatched_bank_debits']:
                umb.append([format_date(item['date']), item['description'][:60], item['reference'][:30], f"{item['amount']:,.2f}", 'DEBIT (Payment)'])
            pd.DataFrame(umb).to_excel(writer, sheet_name='UNMATCHED_BANK', index=False, header=False)

        summary_data = [
            ['RECONCILIATION SUMMARY', ''],
            ['', ''],
            ['Total Bank Transactions', f"{len(bank_df)}"],
            ['Total Ledger Transactions', f"{len(ledger_df)}"],
            ['Matched Transactions', f"{len(match_results['matches'])}"],
            ['', ''],
            ['Unmatched - Ledger Only', f"{len(match_results['unmatched_ledger_credits']) + len(match_results['unmatched_ledger_debits'])}"],
            ['Unmatched - Bank Only', f"{len(match_results['unmatched_bank_credits']) + len(match_results['unmatched_bank_debits'])}"],
            ['', ''],
            ['Opening Bank Balance', f"{recon_statement['opening_balance']:,.2f}"],
            ['Total Reconciling Items', f"{recon_statement['total_adjustment']:,.2f}"],
            ['Adjusted Balance', f"{recon_statement['adjusted_balance']:,.2f}"],
            ['Ledger Balance', f"{recon_statement['ledger_balance']:,.2f}"],
            ['Difference', f"{recon_statement['difference']:,.2f}"],
        ]
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='SUMMARY', index=False, header=False)

        for sheetname in writer.sheets:
            ws = writer.sheets[sheetname]
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value and len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    return output_path

# =========================
# Main UI (reconciliation)
# =========================
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🏦 Bank Statement")
    bank_file = st.file_uploader(
        "Bank Statement (Excel file)",
        type=["xlsx", "xls"],
        help="Upload your bank statement Excel file. Required columns: Date, Credit, Debit, Description (or similar).",
        key="bank"
    )
    st.caption("Accepts .xlsx or .xls, max 200MB per file")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📒 Cashbook / Ledger")
    ledger_file = st.file_uploader(
        "Ledger / Cashbook (Excel file)",
        type=["xlsx", "xls"],
        help="Upload your Yellowcob ledger Excel file. Required columns: Date, Amount, Description (or similar).",
        key="ledger"
    )
    st.caption("Accepts .xlsx or .xls, max 200MB per file")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("⚙️ Reconciliation Settings")
tolerance_col1, tolerance_col2 = st.columns(2)
with tolerance_col1:
    opening_balance_manual = st.number_input("Opening Bank Balance (if not auto-detected)", value=0.00, step=100.00, format="%.2f")
with tolerance_col2:
    use_manual_opening = st.checkbox("Use manual opening balance", value=False)
st.markdown('</div>', unsafe_allow_html=True)

run_recon = st.button("🔄 RUN RECONCILIATION", use_container_width=True)

if run_recon:
    if not bank_file or not ledger_file:
        st.error("Please upload both Bank Statement and Ledger files")
        st.stop()
    try:
        with st.spinner("Loading bank statement..."):
            bank_df, bank_opening, _, _ = load_bank_statement(bank_file)
        with st.spinner("Loading ledger..."):
            ledger_df, ledger_closing, _ = load_ledger(ledger_file)
        opening_balance = opening_balance_manual if use_manual_opening else (bank_opening or 0)
        with st.spinner("Matching transactions..."):
            match_results = match_transactions(bank_df, ledger_df)
        working_paper = build_working_paper(match_results)
        recon_statement = build_recon_statement(opening_balance, ledger_closing, match_results)
        output_file = export_to_excel(working_paper, recon_statement, bank_df, ledger_df, match_results)
        st.success(f"✅ Reconciliation complete! {len(match_results['matches'])} matched.")
        
        st.markdown("---")
        st.subheader("📋 Working Paper Preview")
        wp_preview = working_paper[['SECTION','LEDGER_DATE','LEDGER_DESC','LEDGER_REF','LEDGER_AMOUNT','MATCH_STATUS','BANK_AMOUNT','BANK_DATE','BANK_REF','BANK_DESC']].head(30).copy()
        wp_preview['LEDGER_DATE'] = wp_preview['LEDGER_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_preview['BANK_DATE'] = wp_preview['BANK_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_preview['LEDGER_AMOUNT'] = wp_preview['LEDGER_AMOUNT'].apply(lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' else '')
        wp_preview['BANK_AMOUNT'] = wp_preview['BANK_AMOUNT'].apply(lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' else '')
        st.dataframe(wp_preview, use_container_width=True)
        
        with open(output_file, 'rb') as f:
            st.download_button(
                "📥 Download Excel Report",
                data=f,
                file_name=f"bank_reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())

# =========================
# Footer with logout
# =========================
st.markdown("")
logout_c1, logout_c2, logout_c3 = st.columns([1, 1, 1])
with logout_c2:
    if st.button("Logout", use_container_width=True):
        logout()
