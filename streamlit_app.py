"""Company Check Management — single/bulk MICR layout proof (not negotiable)."""
import io
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

MAP_COLUMNS = ["Property", "Company", "Bank Name", "Routing Number", "Account Number", "Starting Check Number"]
DEFAULT_MAP = pd.DataFrame([
    {"Property": "Example Property A", "Company": "Example Development LLC", "Bank Name": "Example Bank", "Routing Number": "000000000", "Account Number": "0001234567", "Starting Check Number": 1001},
    {"Property": "Example Property B", "Company": "Example Housing LLC", "Bank Name": "Example Bank", "Routing Number": "000000000", "Account Number": "0009876543", "Starting Check Number": 2001},
])
if "property_map" not in st.session_state:
    st.session_state.property_map = DEFAULT_MAP.copy()


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
    missing = set(MAP_COLUMNS) - set(df.columns)
    if missing:
        raise ValueError(f"Missing mapping columns: {', '.join(sorted(missing))}")
    for col in MAP_COLUMNS[:-1]:
        df[col] = df[col].fillna("").astype(str).str.strip()
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
    page_w, page_h = letter
    x, y, w, h = .45*inch, 4.25*inch, 7.6*inch, 3.0*inch
    c.setStrokeColorRGB(.35, .35, .35)
    c.rect(x, y, w, h, stroke=1, fill=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x+.2*inch, y+h-.4*inch, str(account["Company"])[:52])
    c.setFont("Helvetica", 8)
    c.drawString(x+.2*inch, y+h-.61*inch, str(account["Bank Name"])[:64])
    c.drawRightString(x+w-.2*inch, y+h-.4*inch, f"CHECK NO. {check_no}")
    c.drawRightString(x+w-.2*inch, y+h-.61*inch, payment["Check Date"].strftime("%m/%d/%Y"))
    c.setFont("Helvetica", 8)
    c.drawString(x+.2*inch, y+h-1.06*inch, "PAY TO THE ORDER OF")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x+1.48*inch, y+h-1.06*inch, str(payment["Payee"])[:48])
    c.drawRightString(x+w-.2*inch, y+h-1.06*inch, f"${payment['Amount']:,.2f}")
    c.setFont("Helvetica", 8)
    c.drawString(x+.2*inch, y+h-1.45*inch, amount_words(payment["Amount"])[:110])
    c.line(x+.2*inch, y+h-1.53*inch, x+w-.2*inch, y+h-1.53*inch)
    c.drawString(x+.2*inch, y+h-1.9*inch, "MEMO: " + str(payment.get("Memo", ""))[:75])
    c.drawRightString(x+w-.2*inch, y+.68*inch, "AUTHORIZED SIGNATURE: __________________________")
    line = micr_string(check_no, account["Routing Number"], account["Account Number"], transit, on_us)
    if FONT_OK:
        font_size = 12
        width = pdfmetrics.stringWidth(line, "MICR_E13B", font_size)
        if width > w-.6*inch:
            font_size *= (w-.6*inch)/width
        c.setFont("MICR_E13B", font_size)
        c.drawString(x+.3*inch, y+.24*inch, line)
    else:
        c.setFont("Helvetica", 8)
        c.drawString(x+.3*inch, y+.24*inch, "MICR FONT MISSING — NOT VALID")
    c.saveState()
    c.setFillColorRGB(.73, .73, .73)
    c.setFont("Helvetica-Bold", 27)
    c.translate(page_w/2, page_h/2)
    c.rotate(30)
    c.drawCentredString(0, 0, "TEST PREVIEW — NOT NEGOTIABLE")
    c.restoreState()
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


st.sidebar.header("Property & Bank Mapping")
st.sidebar.caption("Edit account details here. Bank numbers are text to preserve leading zeros.")
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
    key="property_editor",
    column_config={
        "Routing Number": st.column_config.TextColumn("Routing Number"),
        "Account Number": st.column_config.TextColumn("Account Number"),
        "Starting Check Number": st.column_config.NumberColumn("Starting Check Number", min_value=1, step=1),
    },
)
st.session_state.property_map = edited_map
st.sidebar.download_button("Export mapping CSV (sensitive)", data=edited_map.to_csv(index=False).encode(), file_name="property_bank_mapping.csv", mime="text/csv")
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
    st.caption(f"Routing: *****{acct['Routing Number'][-4:]} | Account: ****{acct['Account Number'][-4:]}")
    with st.form("single_check"):
        payee = st.text_input("Payee")
        amount = st.text_input("Amount", "100.00")
        check_date = st.date_input("Check Date", value=date.today())
        memo = st.text_input("Memo")
        check_no = st.number_input("Check Number (test)", min_value=1, step=1, value=int(acct["Starting Check Number"]))
        confirmed = st.checkbox("I understand this is a watermarked layout proof, not a payable check")
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
