import pandas as pd
import numpy as np
import re
from datetime import datetime
from dateutil.parser import parse as dt_parse
import warnings
warnings.filterwarnings("ignore")

# -------------------------------------------------------------------
# 1. Smart table detection
# -------------------------------------------------------------------
def detect_header_row(df_raw, max_scan=80):
    """
    Finds the most likely header row by scoring rows on:
    - number of non‑empty cells
    - presence of date/amount/reference keywords
    - ratio of text vs numeric
    """
    best_score = -1
    best_row = 0
    for r in range(min(max_scan, len(df_raw))):
        row = df_raw.iloc[r]
        non_empty = row.notna().sum()
        if non_empty < 2:
            continue
        # Count how many cells look like they could be column headers
        # (i.e. contain letters and are not purely numeric)
        header_like = 0
        for v in row:
            if pd.isna(v):
                continue
            s = str(v).strip()
            if len(s) > 1 and re.search(r'[A-Za-z]', s):
                header_like += 1
        # Check for keywords in any cell
        keyword_score = 0
        row_lower = ' '.join([str(v).lower() for v in row if pd.notna(v)])
        keywords = ['date', 'amount', 'description', 'reference', 'invoice', 'document', 'posting', 'balance', 'due', 'entry']
        for kw in keywords:
            if kw in row_lower:
                keyword_score += 3
        # Score: non_empty + header_like*2 + keyword_score
        score = non_empty + header_like * 2 + keyword_score
        if score > best_score:
            best_score = score
            best_row = r
    return best_row

def clean_headers(headers):
    """Normalise header names: strip, replace slashes/spaces, ensure uniqueness."""
    clean = []
    seen = {}
    for h in headers:
        h = str(h).strip()
        if not h:
            h = "Column"
        # Replace problematic characters
        h = re.sub(r'[^a-zA-Z0-9_ ]', '_', h)
        h = re.sub(r'\s+', '_', h)
        # Ensure uniqueness
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        clean.append(h)
    return clean

def extract_table(df_raw, header_row):
    """Extract the data rows below the header row."""
    headers = df_raw.iloc[header_row].fillna('').astype(str).tolist()
    headers = clean_headers(headers)
    data = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    # Ensure data has same number of columns
    if data.shape[1] < len(headers):
        for _ in range(len(headers) - data.shape[1]):
            data[f"extra_{_}"] = None
    elif data.shape[1] > len(headers):
        data = data.iloc[:, :len(headers)]
    data.columns = headers
    data = data.dropna(how='all')
    return data, headers

# -------------------------------------------------------------------
# 2. Column role inference
# -------------------------------------------------------------------
def infer_roles(df):
    """
    Assign roles: 'date', 'amount', 'reference', 'description'
    Returns a dict {role: column_name}
    """
    cols = df.columns.tolist()
    roles = {}
    # Helper: score a column for a given role
    def score_col(role, keywords, sample_size=100):
        best_col = None
        best_score = -1
        for col in cols:
            # Sample non‑null values
            sample = df[col].dropna().head(sample_size).astype(str)
            if len(sample) == 0:
                continue
            # Keyword match on column name
            col_lower = col.lower()
            name_score = 2 if any(kw in col_lower for kw in keywords) else 0
            # Data characteristics
            if role == 'date':
                date_like = sum(1 for v in sample if is_date_like(v)) / len(sample)
                data_score = date_like * 10
            elif role == 'amount':
                num_like = sum(1 for v in sample if is_number_like(v)) / len(sample)
                data_score = num_like * 10
            elif role == 'reference':
                # Reference columns tend to have alphanumeric, slashes, dashes, not pure numbers
                mixed_ratio = sum(1 for v in sample if re.search(r'[A-Za-z/-]', v)) / len(sample)
                data_score = mixed_ratio * 8
            elif role == 'description':
                # Description columns tend to have longer text, more spaces, no obvious pattern
                avg_len = sample.str.len().mean()
                data_score = min(avg_len / 20, 2) * 5
            else:
                data_score = 0
            score = name_score + data_score
            if score > best_score:
                best_score = score
                best_col = col
        return best_col

    roles['date'] = score_col('date', ['date', 'posting', 'doc date', 'entry date', 'transaction date'])
    roles['amount'] = score_col('amount', ['amount', 'amt', 'value', 'total', 'net', 'lcy', 'balance'])
    roles['reference'] = score_col('reference', ['reference', 'doc', 'document', 'invoice', 'inv', 'order', 'po', 'external', 'number'])
    roles['description'] = score_col('description', ['description', 'desc', 'particulars', 'details', 'narration', 'text'])

    # If amount not found, try to find a column with many numeric values
    if not roles['amount']:
        for col in cols:
            if df[col].apply(is_number_like).mean() > 0.5:
                roles['amount'] = col
                break

    # If date not found, try to find a column with many date‑like values
    if not roles['date']:
        for col in cols:
            if df[col].apply(is_date_like).mean() > 0.3:
                roles['date'] = col
                break

    return roles

def is_date_like(s):
    try:
        dt_parse(str(s), fuzzy=True)
        return True
    except:
        return False

def is_number_like(s):
    try:
        float(re.sub(r'[^\d.\-]', '', str(s)))
        return True
    except:
        return False

# -------------------------------------------------------------------
# 3. Data normalisation and cleaning
# -------------------------------------------------------------------
def normalize_transactions(df, roles):
    """
    Extract transactions from the DataFrame using the inferred roles.
    Returns a clean DataFrame with columns: date, amount, reference, description
    """
    date_col = roles.get('date')
    amount_col = roles.get('amount')
    ref_col = roles.get('reference')
    desc_col = roles.get('description')

    if not date_col or not amount_col:
        return pd.DataFrame()  # insufficient info

    # Copy and convert types
    df_clean = df.copy()
    # Convert amount to numeric
    df_clean['_amount'] = df_clean[amount_col].apply(parse_amount)
    # Convert date
    df_clean['_date'] = df_clean[date_col].apply(parse_date)
    # Remove rows with missing date or amount
    df_clean = df_clean[df_clean['_date'].notna() & df_clean['_amount'].notna()]

    if df_clean.empty:
        return pd.DataFrame()

    # Extract reference and description
    if ref_col:
        df_clean['_ref'] = df_clean[ref_col].astype(str).str.strip()
    else:
        df_clean['_ref'] = ''

    if desc_col:
        df_clean['_desc'] = df_clean[desc_col].astype(str).str.strip()
    else:
        df_clean['_desc'] = ''

    # Filter out summary rows (heuristic)
    # Remove rows where reference is empty and description contains 'total' or 'balance' or 'opening'
    summary_keywords = ['total', 'balance', 'opening', 'closing', 'summary', 'subtotal']
    mask = ~(df_clean['_ref'].str.len() == 0) | ~(df_clean['_desc'].str.lower().str.contains('|'.join(summary_keywords)))
    df_clean = df_clean[mask]

    # Keep only relevant columns
    result = pd.DataFrame({
        'date': df_clean['_date'],
        'amount': df_clean['_amount'],
        'reference': df_clean['_ref'],
        'description': df_clean['_desc'],
    })
    return result

def parse_amount(x):
    try:
        s = str(x).replace(',', '').strip()
        # handle parentheses
        if '(' in s and ')' in s:
            s = '-' + re.sub(r'[()]', '', s)
        # remove any non-numeric except minus and dot
        s = re.sub(r'[^\d.\-]', '', s)
        if s == '' or s == '-':
            return np.nan
        return float(s)
    except:
        return np.nan

def parse_date(x):
    try:
        return pd.to_datetime(x, errors='coerce')
    except:
        return pd.NaT

# -------------------------------------------------------------------
# 4. Matching logic (amount first, date tie-breaker)
# -------------------------------------------------------------------
def match_transactions(supplier_df, ledger_df, tolerance=0.01):
    """
    Match transactions by absolute amount and date.
    Returns:
      - matches: list of dicts
      - unmatched_supplier: list of dicts
      - unmatched_ledger: list of dicts
    """
    sup = supplier_df.copy()
    led = ledger_df.copy()
    sup['abs_amount'] = sup['amount'].abs().round(2)
    led['abs_amount'] = led['amount'].abs().round(2)
    sup['date'] = pd.to_datetime(sup['date'])
    led['date'] = pd.to_datetime(led['date'])
    sup['row_id'] = [f'S_{i}' for i in range(len(sup))]
    led['row_id'] = [f'L_{i}' for i in range(len(led))]

    matches = []
    used_sup = set()
    used_led = set()

    # Group by rounded amount
    all_amounts = set(sup['abs_amount'].unique()).union(set(led['abs_amount'].unique()))
    for amt in sorted(all_amounts):
        sup_group = sup[sup['abs_amount'] == amt]
        led_group = led[led['abs_amount'] == amt]
        if sup_group.empty or led_group.empty:
            continue

        # Try to match by exact date first
        sup_list = sup_group.to_dict('records')
        led_list = led_group.to_dict('records')
        for s in sup_list:
            if s['row_id'] in used_sup:
                continue
            # Find a ledger with same date
            matched = None
            for l in led_list:
                if l['row_id'] in used_led:
                    continue
                if s['date'].date() == l['date'].date():
                    matched = l
                    break
            if matched:
                matches.append({
                    'supplier_id': s['row_id'],
                    'ledger_id': matched['row_id'],
                    'amount': amt,
                    'supplier_date': s['date'],
                    'ledger_date': matched['date'],
                    'supplier_ref': s['reference'],
                    'ledger_ref': matched['reference'],
                    'supplier_desc': s['description'],
                    'ledger_desc': matched['description'],
                    'match_method': 'amount_and_date',
                })
                used_sup.add(s['row_id'])
                used_led.add(matched['row_id'])
            else:
                # No date match, will try amount-only later
                pass

        # Now match remaining by amount only
        sup_remaining = sup_group[~sup_group['row_id'].isin(used_sup)]
        led_remaining = led_group[~led_group['row_id'].isin(used_led)]
        for s in sup_remaining.to_dict('records'):
            if s['row_id'] in used_sup:
                continue
            for l in led_remaining.to_dict('records'):
                if l['row_id'] in used_led:
                    continue
                matches.append({
                    'supplier_id': s['row_id'],
                    'ledger_id': l['row_id'],
                    'amount': amt,
                    'supplier_date': s['date'],
                    'ledger_date': l['date'],
                    'supplier_ref': s['reference'],
                    'ledger_ref': l['reference'],
                    'supplier_desc': s['description'],
                    'ledger_desc': l['description'],
                    'match_method': 'amount_only',
                })
                used_sup.add(s['row_id'])
                used_led.add(l['row_id'])
                break

    # Unmatched
    unmatched_sup = sup[~sup['row_id'].isin(used_sup)].to_dict('records')
    unmatched_led = led[~led['row_id'].isin(used_led)].to_dict('records')

    return {
        'matches': matches,
        'unmatched_supplier': unmatched_sup,
        'unmatched_ledger': unmatched_led,
    }

# -------------------------------------------------------------------
# 5. Build working paper (side by side)
# -------------------------------------------------------------------
def build_working_paper(match_results):
    rows = []
    # Matched items
    for m in match_results['matches']:
        rows.append({
            'SECTION': 'MATCHED',
            'SUPPLIER_DATE': format_date(m['supplier_date']),
            'SUPPLIER_REF': m['supplier_ref'],
            'SUPPLIER_DESC': m['supplier_desc'],
            'SUPPLIER_AMOUNT': m['amount'],
            'MATCH_METHOD': m['match_method'],
            'LEDGER_DATE': format_date(m['ledger_date']),
            'LEDGER_REF': m['ledger_ref'],
            'LEDGER_DESC': m['ledger_desc'],
            'LEDGER_AMOUNT': m['amount'],
        })
    # Unmatched supplier
    for item in match_results['unmatched_supplier']:
        rows.append({
            'SECTION': 'UNMATCHED - SUPPLIER ONLY',
            'SUPPLIER_DATE': format_date(item['date']),
            'SUPPLIER_REF': item['reference'],
            'SUPPLIER_DESC': item['description'],
            'SUPPLIER_AMOUNT': item['amount'],
            'MATCH_METHOD': 'NO LEDGER MATCH',
            'LEDGER_DATE': '',
            'LEDGER_REF': '',
            'LEDGER_DESC': '',
            'LEDGER_AMOUNT': '',
        })
    # Unmatched ledger
    for item in match_results['unmatched_ledger']:
        rows.append({
            'SECTION': 'UNMATCHED - LEDGER ONLY',
            'SUPPLIER_DATE': '',
            'SUPPLIER_REF': '',
            'SUPPLIER_DESC': '',
            'SUPPLIER_AMOUNT': '',
            'MATCH_METHOD': 'NO SUPPLIER MATCH',
            'LEDGER_DATE': format_date(item['date']),
            'LEDGER_REF': item['reference'],
            'LEDGER_DESC': item['description'],
            'LEDGER_AMOUNT': item['amount'],
        })
    return pd.DataFrame(rows)

def format_date(dt):
    if pd.isna(dt):
        return ''
    return dt.strftime('%d/%m/%y')

# -------------------------------------------------------------------
# 6. Full reconciliation pipeline
# -------------------------------------------------------------------
def reconcile_supplier(supplier_file, ledger_file, tolerance=0.01):
    """
    Main entry point: loads files, detects tables, infers columns,
    normalises, matches, and returns results.
    """
    # Load raw files
    try:
        sup_raw = pd.read_excel(supplier_file, header=None)
        led_raw = pd.read_excel(ledger_file, header=None)
    except Exception as e:
        raise ValueError(f"Could not read files: {e}")

    # Detect header rows
    sup_header = detect_header_row(sup_raw)
    led_header = detect_header_row(led_raw)

    # Extract tables
    sup_df, _ = extract_table(sup_raw, sup_header)
    led_df, _ = extract_table(led_raw, led_header)

    if sup_df.empty or led_df.empty:
        return None, "No data found after table extraction."

    # Infer roles
    sup_roles = infer_roles(sup_df)
    led_roles = infer_roles(led_df)

    # Normalise
    sup_norm = normalize_transactions(sup_df, sup_roles)
    led_norm = normalize_transactions(led_df, led_roles)

    if sup_norm.empty or led_norm.empty:
        return None, "No valid transactions found after cleaning."

    # Match
    match_results = match_transactions(sup_norm, led_norm, tolerance)

    # Build working paper
    working_paper = build_working_paper(match_results)

    # Summary
    summary = {
        'total_supplier': len(sup_norm),
        'total_ledger': len(led_norm),
        'matched': len(match_results['matches']),
        'unmatched_supplier': len(match_results['unmatched_supplier']),
        'unmatched_ledger': len(match_results['unmatched_ledger']),
    }

    return {
        'working_paper': working_paper,
        'match_results': match_results,
        'summary': summary,
        'used_columns': {'supplier': sup_roles, 'ledger': led_roles},
    }, None