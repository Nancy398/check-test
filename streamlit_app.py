import io
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import date
import pandas as pd
import streamlit as st
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path

from fontTools.ttLib import TTFont as FontToolsTTFont
from pathlib import Path

font_path = Path(__file__).resolve().parent / "micr-e13b.ttf"

font = FontToolsTTFont(str(font_path))
cmap = font.getBestCmap() or {}

st.write("MICR special character mapping:")

for char in ["A", "B", "C", "D", "⑆", "⑇", "⑈", "⑉"]:
    st.write(
        repr(char),
        "→",
        cmap.get(ord(char), "NOT SUPPORTED")
    )

MICR_FONT_PATH = Path(__file__).resolve().parent / "micr-e13b.ttf"

if MICR_FONT_PATH.exists():
    pdfmetrics.registerFont(
        TTFont("MICR_E13B", str(MICR_FONT_PATH))
    )
else:
    raise FileNotFoundError(
        f"MICR font not found: {MICR_FONT_PATH}"
    )

st.set_page_config(page_title="Check Management - Prototype", layout="wide")

st.title("Company Check Management")
st.caption("Prototype only — test with fictional data and plain paper. Not approved for issuing negotiable checks.")

# -----------------------------
# Session state
# -----------------------------
if "property_map" not in st.session_state:
    st.session_state.property_map = pd.DataFrame([
        {
            "Property": "1438 W 37th Dr",
            "Company": "1438 Development LLC",
            "Bank Name": "Test Bank",
            "Routing Number": "000000000",
            "Account Number": "TEST-ACCOUNT-001",
            "Starting Check Number": 1001,
        },
        {
            "Property": "1252 W 37th St",
            "Company": "1252 Development LLC",
            "Bank Name": "Test Bank",
            "Routing Number": "000000000",
            "Account Number": "TEST-ACCOUNT-002",
            "Starting Check Number": 2001,
        },
    ])
if "confirmed_batch" not in st.session_state:
    st.session_state.confirmed_batch = None

# -----------------------------
# Helpers
# -----------------------------
REQUIRED_COLUMNS = ["Payee", "Property", "Amount", "Check Date"]
OPTIONAL_COLUMNS = ["Memo", "Reference"]

def money(value):
    try:
        d = Decimal(str(value).replace("$", "").replace(",", "").strip())
        if not d.is_finite():
            raise InvalidOperation
        d = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if d <= 0:
            raise ValueError("Amount must be greater than zero.")
        return d
    except (InvalidOperation, ValueError, AttributeError):
        raise ValueError(f"Invalid positive amount: {value!r}")

def amount_words(amount: Decimal) -> str:
    """Convert USD amount to words for display."""
    ones = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven",
            "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen",
            "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def under_thousand(n):
        parts = []
        if n >= 100:
            parts += [ones[n // 100], "Hundred"]
            n %= 100
        if n >= 20:
            parts.append(tens[n // 10])
            if n % 10:
                parts.append(ones[n % 10])
        elif n > 0:
            parts.append(ones[n])
        return " ".join(parts) if parts else "Zero"

    def integer_words(n):
        if n == 0:
            return "Zero"
        groups = [(1_000_000_000, "Billion"), (1_000_000, "Million"), (1000, "Thousand"), (1, "")]
        parts = []
        for value, label in groups:
            group = n // value
            if group:
                parts.append(under_thousand(group) + (f" {label}" if label else ""))
                n %= value
        return " ".join(parts)

    cents = int((amount * 100) % 100)
    dollars = int(amount)
    return f"{integer_words(dollars)} Dollars and {cents:02d}/100"

def validate_upload(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype={"Payee": str, "Property": str})
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))

    for col in ["Payee", "Property"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "Memo" not in df.columns:
        df["Memo"] = ""
    if "Reference" not in df.columns:
        df["Reference"] = ""

    df["Memo"] = df["Memo"].fillna("").astype(str).str.strip()
    df["Reference"] = df["Reference"].fillna("").astype(str).str.strip()

    if (df["Payee"] == "").any():
        raise ValueError("Payee cannot be blank.")
    if (df["Property"] == "").any():
        raise ValueError("Property cannot be blank.")

    df["Amount"] = df["Amount"].apply(money)
    df["Check Date"] = pd.to_datetime(df["Check Date"], errors="coerce")
    if df["Check Date"].isna().any():
        raise ValueError("One or more Check Date values are invalid.")
    df["Check Date"] = df["Check Date"].dt.date

    # Add a visible row number for troubleshooting
    df.insert(0, "Excel Row", range(2, len(df) + 2))
    return df

def safe_text(value, max_len=70):
    text = str(value or "")
    return text if len(text) <= max_len else text[:max_len - 3] + "..."

def draw_preview_check(c, row, check_number, account, preview_only=True):
    # US Letter page; sample layout is for layout testing only.
    page_w, page_h = letter
    x = 0.45 * inch
    y = 4.25 * inch
    w = 7.6 * inch
    h = 3.0 * inch

    c.setLineWidth(0.8)
    c.rect(x, y, w, h)

    # Bank/company header
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x + 0.22 * inch, y + h - 0.38 * inch, safe_text(account["Company"], 55))
    c.setFont("Helvetica", 8)
    c.drawString(x + 0.22 * inch, y + h - 0.58 * inch, safe_text(account["Bank Name"], 55))
    c.drawRightString(x + w - 0.22 * inch, y + h - 0.38 * inch, f"CHECK NO. {check_number}")
    c.drawRightString(x + w - 0.22 * inch, y + h - 0.58 * inch, row["Check Date"].strftime("%m/%d/%Y"))

    c.setFont("Helvetica", 9)
    c.drawString(x + 0.22 * inch, y + h - 1.05 * inch, "PAY TO THE ORDER OF")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x + 1.45 * inch, y + h - 1.05 * inch, safe_text(row["Payee"], 52))
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(x + w - 0.22 * inch, y + h - 1.05 * inch, f"${row['Amount']:,.2f}")

    c.setFont("Helvetica", 9)
    c.drawString(x + 0.22 * inch, y + h - 1.48 * inch, safe_text(amount_words(row["Amount"]), 92))
    c.line(x + 0.22 * inch, y + h - 1.56 * inch, x + w - 0.22 * inch, y + h - 1.56 * inch)

    c.setFont("Helvetica", 8)
    memo = safe_text(row.get("Memo", ""), 80)
    c.drawString(x + 0.22 * inch, y + h - 1.92 * inch, f"Memo: {memo}")
    c.drawString(x + 0.22 * inch, y + 0.58 * inch, "AUTHORIZED SIGNATURE: __________________________________")
    c.drawRightString(x + w - 0.22 * inch, y + 0.58 * inch, "VOID IF NOT SIGNED")

    c.setFont("MICR_E13B", 12)
    c.setFillColorRGB(0, 0, 0)
    
    micr_test = f"{check_number}  123456789  0001234567"
    
    c.drawString(
        x + 0.35 * inch,
        y + 0.22 * inch,
        micr_test
    )

    if preview_only:
        c.setFillGray(0.75)
        c.setFont("Helvetica-Bold", 30)
        c.saveState()
        c.translate(page_w / 2, page_h / 2)
        c.rotate(32)
        c.drawCentredString(0, 0, "TEST PREVIEW — NOT NEGOTIABLE")
        c.restoreState()
        c.setFillGray(0)
    c.showPage()

def build_pdf(df, property_map, start_override, preview_only=True):
    mapping = property_map.set_index("Property").to_dict("index")
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=letter)
    c.setTitle("Check Management Test PDF")

    for i, (_, row) in enumerate(df.iterrows()):
        account = mapping[row["Property"]]
        check_number = int(start_override) + i
        draw_preview_check(c, row, check_number, account, preview_only=preview_only)

    c.save()
    output.seek(0)
    return output.getvalue()

# -----------------------------
# Sidebar configuration
# -----------------------------
st.sidebar.header("1. Property & Bank Mapping")
st.sidebar.caption("Use fictional/test details for this prototype.")
edited_map = st.sidebar.data_editor(
    st.session_state.property_map,
    num_rows="dynamic",
    use_container_width=True,
    key="property_mapping_editor",
    column_config={
        "Starting Check Number": st.column_config.NumberColumn(
            min_value=1, step=1, format="%d"
        ),
        "Routing Number": st.column_config.TextColumn(),
        "Account Number": st.column_config.TextColumn(),
    },
)
st.session_state.property_map = edited_map

st.sidebar.header("2. Check Number")
start_number = st.sidebar.number_input(
    "Starting check number for this test PDF",
    min_value=1,
    max_value=999999999,
    value=1001,
    step=1,
)
st.sidebar.warning(
    "Prototype only: the starting number is manually entered and is not reserved in a database."
)

# -----------------------------
# Upload and validate
# -----------------------------
st.header("Upload payment file")
st.write("Required columns: `Payee`, `Property`, `Amount`, `Check Date`. Optional: `Memo`, `Reference`.")

sample = pd.DataFrame([
    {"Payee": "ABC Framing Inc.", "Property": "1438 W 37th Dr", "Amount": 2850.00, "Check Date": date.today(), "Memo": "Framing labor", "Reference": "TEST-001"},
    {"Payee": "John Smith", "Property": "1438 W 37th Dr", "Amount": 950.00, "Check Date": date.today(), "Memo": "Weekly labor", "Reference": "TEST-002"},
])
sample_buffer = io.BytesIO()
with pd.ExcelWriter(sample_buffer, engine="openpyxl") as writer:
    sample.to_excel(writer, index=False, sheet_name="Payments")
sample_buffer.seek(0)
st.download_button(
    "Download sample Excel template",
    data=sample_buffer.getvalue(),
    file_name="check_upload_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

if uploaded:
    try:
        df = validate_upload(uploaded)
        st.success(f"Imported {len(df)} payment row(s).")
        st.subheader("Imported payments")
        display_df = df.copy()
        display_df["Amount"] = display_df["Amount"].map(lambda x: f"${x:,.2f}")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        mapping_df = st.session_state.property_map.copy()
        if mapping_df.empty or "Property" not in mapping_df.columns:
            st.error("Property mapping table is empty or invalid.")
            st.stop()

        mapping_df["Property"] = mapping_df["Property"].fillna("").astype(str).str.strip()
        if mapping_df["Property"].duplicated().any():
            st.error("Property mapping has duplicate Property names. Please fix them in the sidebar.")
            st.stop()

        map_records = mapping_df.set_index("Property").to_dict("index")
        missing_properties = sorted(set(df["Property"]) - set(map_records.keys()))
        if missing_properties:
            st.error("These Properties are not mapped: " + ", ".join(missing_properties))
            st.stop()

        problems = []
        for prop in sorted(set(df["Property"])):
            acct = map_records[prop]
            for field in ["Company", "Bank Name", "Routing Number", "Account Number"]:
                val = str(acct.get(field, "") or "").strip()
                if not val:
                    problems.append(f"{prop}: missing {field}")
            routing = str(acct.get("Routing Number", "") or "").strip()
            if routing and (not routing.isdigit() or len(routing) != 9):
                problems.append(f"{prop}: Routing Number should be 9 digits")
        if problems:
            st.error("Please fix mapping issues:\n\n- " + "\n- ".join(problems))
            st.stop()

        # Show matched company/account summary with masked account number
        st.subheader("Matched company and bank account")
        summary_rows = []
        for prop in sorted(set(df["Property"])):
            acct = map_records[prop]
            account = str(acct["Account Number"])
            masked = ("*" * max(0, len(account) - 4)) + account[-4:]
            summary_rows.append({
                "Property": prop,
                "Company": acct["Company"],
                "Bank": acct["Bank Name"],
                "Routing Number": acct["Routing Number"],
                "Account (masked)": masked,
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        st.subheader("Review individual checks")
        selected_row = st.selectbox(
            "Select a payment to inspect",
            options=list(range(len(df))),
            format_func=lambda i: f"{i + 1}. {df.iloc[i]['Payee']} — ${df.iloc[i]['Amount']:,.2f}",
        )
        row = df.iloc[selected_row]
        acct = map_records[row["Property"]]
        c1, c2, c3 = st.columns(3)
        c1.metric("Payee", safe_text(row["Payee"], 40))
        c2.metric("Amount", f"${row['Amount']:,.2f}")
        c3.metric("Check Number (test)", str(int(start_number) + selected_row))
        st.write(f"**Property:** {row['Property']}  |  **Company:** {acct['Company']}")
        st.write(f"**Bank:** {acct['Bank Name']}  |  **Date:** {row['Check Date']}")
        st.write(f"**Amount in words:** {amount_words(row['Amount'])}")
        st.write(f"**Memo:** {row['Memo'] or '—'}")

        total = sum(df["Amount"], Decimal("0.00"))
        a, b, ccol = st.columns(3)
        a.metric("Number of checks", len(df))
        b.metric("Batch total", f"${total:,.2f}")
        ccol.metric("Properties", df["Property"].nunique())

        st.info(
            "The PDF produced by this prototype is watermarked and includes a MICR placeholder. "
            "It is for layout review only and is not suitable for payment."
        )

        confirm = st.checkbox(
            "I reviewed the payees, properties, amounts, dates, and company/account mapping. "
            "I understand this prototype PDF is not a valid payment instrument."
        )

        if st.button("Generate test PDF", type="primary", disabled=not confirm):
            pdf = build_pdf(df, mapping_df, int(start_number), preview_only=True)
            st.session_state["generated_pdf"] = pdf
            st.session_state["generated_filename"] = "check_test_preview.pdf"
            st.session_state["generated_count"] = len(df)
            st.success("Test PDF generated. No check numbers were reserved and no payment was made.")

        if st.session_state.get("generated_pdf"):
            st.download_button(
                "Download test PDF",
                data=st.session_state["generated_pdf"],
                file_name=st.session_state.get("generated_filename", "check_test_preview.pdf"),
                mime="application/pdf",
            )

    except Exception as exc:
        st.error(f"Could not process file: {exc}")

else:
    st.info("Upload an Excel file or download the sample template to get started.")

with st.expander("Important limitations"):
    st.markdown("""
- This is a local prototype and does not connect to MySQL or QuickBooks.
- The PDF is watermarked and contains a MICR placeholder, not a real MICR line.
- It does not issue, clear, transmit, or verify payments.
- Do not enter real bank account numbers in a shared or unprotected test environment.
- Before issuing checks, confirm your bank's check-stock, MICR font, magnetic toner, layout, and testing requirements.
- The prototype does not implement authentication, approval workflows, persistent audit logs, transactional check-number reservation, or secure secrets management.
""")
