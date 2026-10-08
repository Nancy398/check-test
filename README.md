# FreeCheckPrint + existing Streamlit Check Generator (Lakeview test)

This package preserves the existing Single and Bulk workflows and adds a third sidebar mode: **FreeCheckPrint (Lakeview Test)**.

## Setup

```bash
pip install -r requirements.txt
python -m playwright install chromium
streamlit run app.py
```

Keep your existing `.streamlit/secrets.toml` for Google Sheets; it is **not included** in this ZIP. The Project sheet must have `Project_Name`, `Company`, `Account`, `Routing_Number`, `Account_Number`, `Company_Name`, `Bank_Name`. Bank number columns should be plain text.

## Important limitations

**This is a test integration, not a verified end-to-end download.** This execution environment cannot access the live website. `selectors.json` contains plausible selector candidates, **not verified selectors**. If the site uses different names, inspect the website and update that file.

The app fills form fields, selects **Check Stock**, checks Terms **only after you explicitly consent**, and tries to capture a real PDF download or PDF tab. It does not silently fall back to Plain Paper, or fabricate a PDF using browser `page.pdf()`. Some sites invoke the print dialog rather than providing a PDF download; that path needs further site-specific integration.

The website may have anti-bot restrictions, changed form fields, or a different multi-step flow. Streamlit Cloud may require Playwright browser OS dependencies; local Windows/Linux execution is easiest for the first trial.

Do not enter production bank information until you have verified the website's privacy/security practices and reviewed the automation. Generated checks must be reviewed before printing or issuance. No Google Sheets history is written in test mode.
