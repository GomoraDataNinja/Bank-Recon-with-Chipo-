# Bank Reconciliation App With Chipo

A Streamlit-based bank reconciliation tool that matches transactions by amount and date, identifies reconciling items, and produces a downloadable Excel workbook with working paper, reconciliation statement, and diagnostics.

## Features

- 🔐 Password-protected login (set via `APP_PASSWORD` env or Streamlit secrets)
- 📂 Upload bank statement and ledger Excel files
- 🔄 Automatic matching by amount (rounded) and exact date
- 📋 Side-by-side working paper (matched items first, then unmatched)
- 📊 Clean reconciliation statement with opening balance, bank closing, reconciling items, adjusted balance, ledger balance, and unreconciled difference
- 🔎 Diagnostic analysis to explain why a difference exists
- 📥 Excel export with 6 sheets: WORKING_PAPER, RECON_STATEMENT, MATCHED_DETAIL, UNMATCHED_LEDGER, UNMATCHED_BANK, SUMMARY
- 💾 Session persistence – results survive page reloads

## Deploy on Streamlit Cloud

1. Push this repository to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in.
3. Click "New app" and select your repo.
4. Set the main file to `bank.py` (or whatever you named it).
5. In the "Advanced settings", add your `APP_PASSWORD` as a secret, or set it in the secrets management.
6. Click "Deploy".

## Local Development

```bash
pip install -r requirements.txt
streamlit run bank.py
