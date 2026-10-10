"""Plain white, logo-free check layout proof with single and bulk modes."""
import io
import json
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Company Check Management", layout="wide")
st.title("Company Check Management")
st.caption("MICR layout proof • Test only — not a negotiable check. Verify with your bank before production printing.")

FONT_PATH = Path(__file__).resolve().parent / "micr-e13b.ttf"
FONT_OK = False
if FONT_PATH.is_file():
    try:
        pdfmetrics.registerFont(TTFont("MICR_E13B", str(FONT_PATH)))
        FONT_OK = True
    except Exception as exc:
        st.error(f"MICR font could not be loaded: {exc}")
else:
    st.warning("micr-e13b.ttf not found next to streamlit_app.py. PDF will show a MICR placeholder.")

DEFAULT_COMPANY_ADDRESS = "3250 Wilshire Blvd, STE1502, Los Angeles, CA 90010"
MAP_COLUMNS = ["Property", "Company", "Company Address", "Bank Name", "Routing Number", "Account Number", "Starting Check Number"]
DEFAULT_MAP = pd.DataFrame([
    {"Property": "Example Property A", "Company": "Example Development LLC", "Company Address": DEFAULT_COMPANY_ADDRESS, "Bank Name": "Example Bank", "Routing Number": "000000000", "Account Number": "0001234567", "Starting Check Number": 1001},
    {"Property": "Example Property B", "Company": "Example Housing LLC", "Company Address": DEFAULT_COMPANY_ADDRESS, "Bank Name": "Example Bank", "Routing Number": "000000000", "Account Number": "0009876543", "Starting Check Number": 2001},
])
if "property_map" not in st.session_state:
    st.session_state.property_map = DEFAULT_MAP.copy()


def _sheet_client():
    """Read-only access to existing Google Sheets service account secrets."""
    if "GOOGLE_APPLICATION_CREDENTIALS" not in st.secrets:
        raise ValueError("Streamlit Secrets 缺少 GOOGLE_APPLICATION_CREDENTIALS")
    raw = st.secrets["GOOGLE_APPLICATION_CREDENTIALS"]
    if isinstance(raw, str):
        info = json.loads(raw)
    else:
        info = dict(raw)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    return gspread.authorize(creds)


def load_google_sheet_mapping(spreadsheet_name="Check Issuance History", worksheet_name="Project"):
    """Convert the prior Project worksheet into this app's property mapping."""
    worksheet = _sheet_client().open(spreadsheet_name).worksheet(worksheet_name)
    records = worksheet.get_all_records(numericise_ignore=["all"])
    result = []
    for idx, record in enumerate(records, start=2):
        r = {str(k).strip().lower().replace(" ", "_"): str(v).strip() for k, v in record.items()}
        def pick(*names):
            return next((r[n] for n in names if r.get(n)), "")
        project = pick("project_name", "property", "project")
        if not project:
            continue
        company = pick("company_name", "company")
        bank = pick("bank_name", "bank")
        routing = pick("routing_number", "routing")
        account = pick("account_number", "bank_account_number")
        # 'Account' in the prior sheet is an internal account ID (e.g. ACC-8652).
        # Never mistake it for the actual bank account number.
        start_text = pick("starting_check_number", "start_check_number") or "1001"
        try:
            start_no = int(start_text)
        except ValueError:
            raise ValueError(f"Project sheet 第 {idx} 行 Starting Check Number 无效")
        result.append({
            "Property": project,
            "Company": company,
            "Company Address": pick("company_address", "address") or DEFAULT_COMPANY_ADDRESS,
            "Bank Name": bank,
            "Routing Number": routing,
            "Account Number": account,
            "Starting Check Number": start_no,
        })
    if not result:
        raise ValueError("Project worksheet 没有可用项目")
    return validate_map(pd.DataFrame(result))


def money(value):
    try:
        d = Decimal(str(value).replace("$", "").replace(",", "").strip())
        if not d.is_finite():
            raise InvalidOperation
        d = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if d <= 0:
            raise ValueError("Amount must be greater than zero")
        return d
    except (InvalidOperation, ValueError, TypeError):
        raise ValueError(f"Invalid positive amount: {value!r}")


def amount_words(amount):
    ones = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    def under_1000(n):
        parts = []
        if n >= 100:
            parts.extend([ones[n // 100], "Hundred"])
            n %= 100
        if n >= 20:
            parts.append(tens[n // 10])
            if n % 10:
                parts.append(ones[n % 10])
        elif n:
            parts.append(ones[n])
        return " ".join(parts)
    n = int(amount)
    if n == 0:
        result = "Zero"
    else:
        chunks = []
        for factor, label in [(10**9, "Billion"), (10**6, "Million"), (1000, "Thousand"), (1, "")]:
            v, n = divmod(n, factor)
            if v:
                chunks.append(under_1000(v) + (" " + label if label else ""))
        result = " ".join(chunks)
    return f"{result} and {int((amount % 1) * 100):02d}/100 Dollars"


def validate_map(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for optional in ("Company Address",):
        if optional not in df.columns:
            df[optional] = DEFAULT_COMPANY_ADDRESS
        else:
            df[optional] = df[optional].fillna("").astype(str).apply(lambda x: x if x.strip() else DEFAULT_COMPANY_ADDRESS)
    missing = set(MAP_COLUMNS) - set(df.columns)
    if missing:
        raise ValueError(f"Missing mapping columns: {', '.join(sorted(missing))}")
    for col in MAP_COLUMNS[:-1]:
        df[col] = df[col].fillna("").astype(str).str.replace("\r", "", regex=False).str.strip()
    if df.empty or df["Property"].eq("").any() or df["Property"].duplicated().any():
        raise ValueError("Property must be nonempty and unique")
    for _, item in df.iterrows():
        if not re.fullmatch(r"\d{9}", item["Routing Number"]):
            raise ValueError(f"{item['Property']}: Routing Number must be exactly 9 digits")
        if not re.fullmatch(r"\d{1,20}", item["Account Number"]):
            raise ValueError(f"{item['Property']}: Account Number must be 1–20 digits")
        if not item["Company"] or not item["Bank Name"]:
            raise ValueError(f"{item['Property']}: Company and Bank Name are required")
    df["Starting Check Number"] = pd.to_numeric(df["Starting Check Number"], errors="raise").astype(int)
    if (df["Starting Check Number"] < 1).any():
        raise ValueError("Starting Check Number must be positive")
    return df[MAP_COLUMNS]


def validate_payments(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    required = ["Payee", "Property", "Amount", "Check Date"]
    missing = set(required) - set(df.columns)
    if missing:
        raise ValueError(f"Missing payment columns: {', '.join(sorted(missing))}")
    for col in ["Payee", "Property"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
        if df[col].eq("").any():
            raise ValueError(f"{col} cannot be blank")
    df["Memo"] = df["Memo"].fillna("").astype(str) if "Memo" in df else ""
    df["Amount"] = df["Amount"].map(money)
    dates = pd.to_datetime(df["Check Date"], errors="coerce")
    if dates.isna().any():
        raise ValueError("Invalid Check Date")
    df["Check Date"] = dates.dt.date
    return df


def micr_string(check_no, routing, account_no, transit="A", on_us="C"):
    # Character mapping must be visually checked against the specific TTF glyphs.
    # This is an unverified proof string, not bank-approved MICR encoding.
    return f"{on_us}{check_no}{on_us}  {transit}{routing}{transit}  {account_no}{on_us}"


def draw_check(c, payment, account, check_no, transit="A", on_us="C"):
    """FreeCheckPrint-inspired US Letter layout; all printed fields are generated anew.

    Top portion: check; lower portion: record stub. Coordinates use top-left
    origin to match the supplied reference. This is a watermarked layout proof.
    """
    W, H = letter
    def txt(x, top, text, font="Helvetica", size=8, align="left", max_width=None):
        text = str(text or "")
        if max_width is not None:
            while size > 5.5 and pdfmetrics.stringWidth(text, font, size) > max_width:
                size -= 0.3
        c.setFont(font, size)
        y = H - top
        if align == "right": c.drawRightString(x, y, text)
        elif align == "center": c.drawCentredString(x, y, text)
        else: c.drawString(x, y, text)

    def address_lines(raw):
        raw = str(raw or "").replace("\r", "")
        return [v.strip() for v in raw.split("\n") if v.strip()][:3]

    company = str(account["Company"])
    bank = str(account["Bank Name"])
    address = address_lines(account.get("Company Address", ""))
    date_text = payment["Check Date"].strftime("%m/%d/%Y")
    payee = str(payment["Payee"])
    amount = payment["Amount"]
    amount_text = f"{amount:,.2f}"
    memo = str(payment.get("Memo", "") or "")
    words = amount_words(amount)

    # Header positions and typography closely follow the uploaded 2001 sample.
    txt(144, 40, company, "Helvetica-Bold", 9.2, "center", 210)
    for i, line in enumerate(address):
        txt(144, 54 + i*10, line, size=7.7, align="center", max_width=240)
    txt(370, 35, bank, "Helvetica-Bold", 7.4, "center", 188)
    txt(598, 39, check_no, "Courier", 11, "right")
    txt(555, 73, date_text, "Helvetica", 10, "right")

    txt(13, 99, "PAY TO THE", "Helvetica", 5.4)
    txt(13, 106, "ORDER OF", "Helvetica", 5.4)
    txt(72, 105, payee, "Helvetica", 11, max_width=360)
    txt(476, 105, "$", "Helvetica-Bold", 11)
    txt(548, 105, f"****{amount_text}", "Helvetica", 10.5, "right", 90)
    txt(13, 128, "Pay", "Helvetica-Bold", 7.5)
    # Clip to available width without truncating numeric amount.
    txt(36, 128, words + "*"*15, "Helvetica", 9.2, max_width=465)
    txt(602, 128, "DOLLARS", "Helvetica-Bold", 7, "right")
    txt(602, 148, "VOID 90 DAYS AFTER ISSUE", "Helvetica-Bold", 6.5, "right")
    # Removed redundant payee name in the middle of the check.
    c.setLineWidth(.6)
    c.line(350, H-204, 600, H-204)
    txt(492, 216, "AUTHORIZED SIGNATURE", "Helvetica", 5.5, "center")
    txt(13, 221, "MEMO", "Helvetica-Bold", 7)
    txt(46, 221, memo, "Helvetica", 7.2, max_width=290)

    # MICR line: mapping and alignment must be validated by the bank/printer.
    line = micr_string(check_no, account["Routing Number"], account["Account Number"], transit, on_us)
    if FONT_OK:
        size = 12
        while size > 7 and pdfmetrics.stringWidth(line, "MICR_E13B", size) > 400:
            size -= .25
        txt(125, 249, line, "MICR_E13B", size)
    else:
        txt(125, 249, "MICR FONT MISSING — TEST ONLY", "Helvetica-Bold", 8)

    # Payment stub below the check, aligned like the reference.
    txt(7, 298, "CHECK NUMBER:", "Helvetica-Bold", 9)
    txt(97, 298, check_no, "Helvetica", 9)
    txt(7, 316, "DATE:", "Helvetica-Bold", 9)
    txt(97, 316, date_text, "Helvetica", 9)
    txt(7, 334, "PAYEE:", "Helvetica-Bold", 9)
    txt(97, 334, payee, "Helvetica", 9, max_width=330)
    txt(7, 352, "AMOUNT:", "Helvetica-Bold", 9)
    txt(97, 352, "$"+amount_text, "Helvetica", 9)
    txt(7, 370, "MEMO:", "Helvetica-Bold", 9)
    txt(97, 370, memo, "Helvetica", 9, max_width=420)

    txt(306, 402, "Print settings (layout proof):", "Helvetica-Bold", 10, "center")
    for i, tip in enumerate([
        "- US Letter (8.5 x 11 in), portrait orientation",
        "- Print at 100% / Actual size (not Fit to page)",
        "- Layout proof only: MICR and stock require bank verification",
        "- No logo or background; verify details before issuance",
    ]):
        txt(55, 424 + i*20, tip, "Helvetica", 8.5, max_width=510)

    txt(306, 555, "Keep this copy for your records:", "Helvetica-Bold", 11, "center")
    # Miniature duplicate; plain black on white, no decorative branding.
    c.saveState()
    c.translate(126, H-610)
    c.scale(.52, .52)
    # Coordinates in miniature's local space, y goes up.
    c.setFont("Helvetica-Bold", 8)
    c.drawString(0, 0, company[:42])
    c.setFont("Helvetica", 7)
    for i, line in enumerate(address[:2]):
        c.drawString(0, -12 - i*10, line[:50])
    c.setFont("Helvetica-Bold", 8)
    c.drawString(290, 0, bank[:37])
    c.drawRightString(820, 0, str(check_no))
    c.setFont("Helvetica", 8)
    c.drawString(0, -50, payee[:45])
    c.drawRightString(820, -50, "$"+amount_text)
    c.drawString(0, -75, words[:95])
    c.drawString(0, -112, "MEMO  " + memo[:65])
    c.line(560, -112, 820, -112)
    c.setFont("Helvetica-Bold", 10)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(410, -155, "COPY — VOID")
    c.restoreState()

    # No logo, colored background, or diagonal watermark.
    # A clear non-negotiable notice remains outside the check area.
    txt(306, 525, "LAYOUT TEST ONLY — NOT NEGOTIABLE", "Helvetica-Bold", 8, "center")
    c.showPage()

def make_pdf(payments, mapping, transit="A", on_us="C"):
    mapping_by_property = mapping.set_index("Property").to_dict("index")
    counters = {prop: int(record["Starting Check Number"]) for prop, record in mapping_by_property.items()}
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=letter)
    c.setTitle("MICR Layout Test - Not Negotiable")
    numbers = []
    for _, row in payments.iterrows():
        prop = row["Property"]
        if prop not in mapping_by_property:
            raise ValueError(f"Unmapped property: {prop}")
        number = counters[prop]
        counters[prop] += 1
        numbers.append(number)
        draw_check(c, row, mapping_by_property[prop], number, transit, on_us)
    c.save()
    return out.getvalue(), numbers


st.sidebar.header("Google Sheets")
st.sidebar.caption("可直接读取之前 Check Issuance History → Project 的项目、公司和银行信息。只读取，不会修改 Google Sheets。")
spreadsheet_name = st.sidebar.text_input("Spreadsheet name", value="Check Issuance History")
worksheet_name = st.sidebar.text_input("Project worksheet", value="Project")
if st.sidebar.button("Load / Refresh from Google Sheets"):
    try:
        loaded = load_google_sheet_mapping(spreadsheet_name, worksheet_name)
        st.session_state.property_map = loaded
        # Force data_editor to initialize with freshly loaded rows.
        st.session_state.map_editor_version = st.session_state.get("map_editor_version", 0) + 1
        st.sidebar.success(f"Loaded {len(loaded)} properties")
    except Exception as exc:
        st.sidebar.error(f"Google Sheets 读取失败：{exc}")

st.sidebar.header("Property & Bank Mapping")
st.sidebar.caption("Company Address 默认为 3250 Wilshire Blvd, STE1502, Los Angeles, CA 90010，显示在公司名称下面；可单独修改。Routing/Account 请在 Sheets 里设为纯文本以保留前导零。")
map_upload = st.sidebar.file_uploader("Import mapping CSV", type="csv", key="map_upload")
if map_upload is not None:
    token = (map_upload.name, map_upload.size)
    if st.session_state.get("last_map_upload") != token:
        try:
            st.session_state.property_map = validate_map(pd.read_csv(map_upload, dtype=str))
            st.session_state.last_map_upload = token
            st.sidebar.success("Mapping imported")
        except Exception as exc:
            st.sidebar.error(str(exc))

edited_map = st.sidebar.data_editor(
    st.session_state.property_map,
    num_rows="dynamic",
    use_container_width=True,
    key=f"property_editor_{st.session_state.get('map_editor_version', 0)}",
    column_config={
        "Routing Number": st.column_config.TextColumn("Routing Number"),
        "Account Number": st.column_config.TextColumn("Account Number"),
        "Company Address": st.column_config.TextColumn("Company Address"),
        "Starting Check Number": st.column_config.NumberColumn("Starting Check Number", min_value=1, step=1),
    },
)
st.session_state.property_map = edited_map.drop(columns=["Bank Address"], errors="ignore")
st.sidebar.download_button("Export mapping CSV (sensitive)", data=st.session_state.property_map.to_csv(index=False).encode(), file_name="property_bank_mapping.csv", mime="text/csv")
st.sidebar.caption("Mapping is kept in session state only; export to retain it. Protect the exported file.")

try:
    mapping = validate_map(edited_map)
except Exception as exc:
    st.error(f"Please fix Property & Bank Mapping: {exc}")
    st.stop()

with st.expander("MICR font and symbol proof", expanded=False):
    st.write("Font loaded:", FONT_OK)
    st.write("The letters A/B/C/D may be mapped to special MICR symbols by your font. Confirm visually before relying on any mapping.")
    transit = st.selectbox("Transit glyph character (test)", ["A", "B", "C", "D"], index=0)
    on_us = st.selectbox("On-Us glyph character (test)", ["A", "B", "C", "D"], index=2)
    st.code(f"{on_us}2000{on_us}  {transit}123456789{transit}  0012345678{on_us}")

mode = st.radio("Mode", ["Single Check", "Bulk Checks (Excel)"], horizontal=True)
if mode == "Single Check":
    prop = st.selectbox("Property", mapping["Property"].tolist())
    acct = mapping.set_index("Property").loc[prop]
    st.write(f"**Company:** {acct['Company']} | **Bank:** {acct['Bank Name']}")
    st.caption(f"Company Address: {acct.get('Company Address', '') or 'Not entered'}")
    st.caption(f"Routing: *****{acct['Routing Number'][-4:]} | Account: ****{acct['Account Number'][-4:]}")
    with st.form("single_check"):
        payee = st.text_input("Payee")
        amount = st.text_input("Amount", "100.00")
        check_date = st.date_input("Check Date", value=date.today())
        memo = st.text_input("Memo")
        check_no = st.number_input("Check Number (test)", min_value=1, step=1, value=int(acct["Starting Check Number"]))
        confirmed = st.checkbox("I understand this is a test layout, not a payable check")
        submit = st.form_submit_button("Generate test PDF")
    if submit:
        try:
            if not confirmed:
                raise ValueError("Please confirm the test-only acknowledgement")
            if not payee.strip():
                raise ValueError("Payee is required")
            row = {"Payee": payee.strip(), "Property": prop, "Amount": money(amount), "Check Date": check_date, "Memo": memo}
            local_map = mapping.copy()
            local_map.loc[local_map["Property"] == prop, "Starting Check Number"] = int(check_no)
            pdf, numbers = make_pdf(pd.DataFrame([row]), local_map, transit, on_us)
            st.session_state.single_pdf = pdf
            st.session_state.single_filename = f"check_test_{numbers[0]}.pdf"
            st.success("Test PDF generated; no check number was reserved")
        except Exception as exc:
            st.error(str(exc))
    if st.session_state.get("single_pdf"):
        st.download_button("Download single check test PDF", st.session_state.single_pdf, st.session_state.single_filename, mime="application/pdf")
else:
    st.write("Excel columns: **Payee, Property, Amount, Check Date**; optional **Memo**.")
    example = pd.DataFrame([{"Payee": "Example Vendor", "Property": mapping.iloc[0]["Property"], "Amount": 250.00, "Check Date": date.today(), "Memo": "Test"}])
    sample = io.BytesIO()
    with pd.ExcelWriter(sample, engine="openpyxl") as writer:
        example.to_excel(writer, index=False, sheet_name="Payments")
    st.download_button("Download sample Excel", sample.getvalue(), "payment_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    upload = st.file_uploader("Upload payments Excel", type="xlsx")
    if upload:
        try:
            payments = validate_payments(pd.read_excel(upload, dtype={"Payee": str, "Property": str, "Memo": str}))
            unmapped = set(payments["Property"]) - set(mapping["Property"])
            if unmapped:
                raise ValueError("Unmapped properties: " + ", ".join(sorted(unmapped)))
            st.dataframe(payments, hide_index=True, use_container_width=True)
            st.write(f"**{len(payments)} checks** • Total ${sum(payments['Amount'], Decimal('0.00')):,.2f}")
            confirmed = st.checkbox("I reviewed all payments and understand the PDF is a test-only layout proof")
            if st.button("Generate bulk test PDF", disabled=not confirmed):
                pdf, numbers = make_pdf(payments, mapping, transit, on_us)
                st.session_state.bulk_pdf = pdf
                st.success(f"Generated {len(numbers)} test pages. No check numbers reserved.")
        except Exception as exc:
            st.error(str(exc))
    if st.session_state.get("bulk_pdf"):
        st.download_button("Download bulk test PDF", st.session_state.bulk_pdf, "bulk_checks_test.pdf", mime="application/pdf")

st.divider()
st.caption("This app does not persist check issuance history or provide bank-approved MICR output. Test with fictional data first; use bank-approved stock, equipment and verification for actual checks.")
