import base64
from datetime import date
import io
import os
import zipfile
import re
from decimal import Decimal
import fitz
from num2words import num2words
import pandas as pd
import streamlit as st
import gspread
from PyPDF2 import PdfMerger
from google.oauth2.service_account import Credentials

# ----------------- 页面配置 -----------------
st.set_page_config(
    page_title="Check Generator System", page_icon="🧾", layout="wide"
)

# ----------------- 配置文件与路径定义 -----------------

GS_SPREADSHEET_NAME = "Check Issuance History"  # Google 表格的名字
GS_WORKSHEET_NAME = "Sheet1"                  # 历史记录工作表的名字

# ----------------- 1. 获取 Authorization 客户端 -----------------
def get_gc_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["GOOGLE_APPLICATION_CREDENTIALS"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(credentials)

# ----------------- 2. 读取 Google Sheets -----------------
def read_file(name, sheet):
    try:
        gc = get_gc_client()
        worksheet = gc.open(name).worksheet(sheet)
        rows = worksheet.get_all_values()
        if not rows or len(rows) <= 1:
            return pd.DataFrame()
        df = pd.DataFrame.from_records(rows)
        df = pd.DataFrame(df.values[1:], columns=df.iloc[0])
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"❌ 读取 Google Sheets ({sheet}) 失败: {e}")
        return pd.DataFrame()

# ----------------- 3. 保存写入 Google Sheets -----------------
def save_to_history(records):
    """把生成的支票记录直接追加保存到 Google Sheets 中"""
    try:
        gc = get_gc_client()
        worksheet = gc.open(GS_SPREADSHEET_NAME).worksheet(GS_WORKSHEET_NAME)
        
        rows_to_append = []
        for r in records:
            rows_to_append.append([
                r["Check Number"],
                r["Issue Date"],
                r["Company"],
                r["Account"],
                r["Project"],
                r.get("Stage", ""),
                r["Payee Name"],
                r["Amount"],
                r["Memo"]
            ])
        worksheet.append_rows(rows_to_append)
        return True
    except Exception as e:
        st.error(f"❌ 数据保存至 Google Sheets 失败: {e}")
        return False

# ----------------- 4. 计算最新可用支票号 -----------------
def fetch_next_check_numbers_from_gs(df_p_list, default_start_number=1001):
    next_numbers = {}
    try:
        df_gs = read_file(GS_SPREADSHEET_NAME, GS_WORKSHEET_NAME)
        if not df_gs.empty and "Project" in df_gs.columns and "Check Number" in df_gs.columns:
            df_gs["Check Number"] = pd.to_numeric(df_gs["Check Number"], errors="coerce")
            max_checks = df_gs.groupby("Project")["Check Number"].max().to_dict()
            
            for p_name in df_p_list:
                if p_name in max_checks and pd.notnull(max_checks[p_name]):
                    next_numbers[p_name] = int(max_checks[p_name]) + 1
    except Exception as e:
        st.sidebar.warning(f"⚠ 读取云端历史推算支票号失败，将使用默认起始号。({e})")

    for p_name in df_p_list:
        if p_name not in next_numbers:
            next_numbers[p_name] = default_start_number
            
    return next_numbers

# ----------------- 5. 自动匹配 Stage 的拦截逻辑 -----------------
def auto_match_stage(worker_name, project_name, default_stg, default_stg_name, default_sub):
    """根据特殊的 Worker 或 Project 自动覆盖默认的 Stage"""
    if worker_name and "Valente Herrera" in worker_name:
        return "Valente Salary", "Valente Salary", "Payroll"
    
    if project_name and "83 Patrician" in project_name:
        return "83 Patrician Way", "83 Patrician Way", "General"
    
    if project_name and "365 San Gabrial" in project_name:
        return "365 San Gabrial", "365 San Gabrial", "General"
    
    return default_stg, default_stg_name, default_sub

# ----------------- 6. 数据预设加载函数 -----------------
@st.cache_data(ttl=60)
def load_project_presets():
    df_p = read_file(GS_SPREADSHEET_NAME, "Project")
    
    extra_projects = pd.DataFrame([
        {"Project_Name": "83 Patrician Way", "Company": "Development Company", "Account": "ACC-5027"},
        {"Project_Name": "365 San Gabrial", "Company": "Development Company", "Account": "ACC-5027"},
    ])

    if df_p.empty or "Project_Name" not in df_p.columns:
        return extra_projects

    df_combined = pd.concat([df_p, extra_projects], ignore_index=True)
    df_combined = df_combined.drop_duplicates(subset=["Project_Name"], keep="first")
    return df_combined

@st.cache_data(ttl=60)
def load_worker_presets():
    df_w = read_file(GS_SPREADSHEET_NAME, "Worker")
    required_cols = {"Worker_Name", "Stage", "Stage_Name", "Sub_Stage"}
    if df_w.empty or not required_cols.issubset(df_w.columns):
        st.warning("⚠ Worker 表格为空或缺失列！")
        return pd.DataFrame(columns=["Worker_Name", "Stage", "Stage_Name", "Sub_Stage"])
    return df_w

@st.cache_data(ttl=60)
def load_stage_presets():
    df_s = read_file(GS_SPREADSHEET_NAME, "Stage")
    
    extra_stages = pd.DataFrame([
        {"Stage": "83 Patrician Way", "Stage_Name": "83 Patrician Way", "Sub_Stage": "General"},
        {"Stage": "365 San Gabrial", "Stage_Name": "365 San Gabrial", "Sub_Stage": "General"},
        {"Stage": "Valente Salary", "Stage_Name": "Valente Salary", "Sub_Stage": "Payroll"},
    ])

    required_cols = {"Stage", "Stage_Name", "Sub_Stage"}
    if df_s.empty or not required_cols.issubset(df_s.columns):
        return extra_stages

    df_combined = pd.concat([df_s, extra_stages], ignore_index=True)
    df_combined = df_combined.drop_duplicates(subset=["Stage", "Sub_Stage"], keep="first")

    return df_combined
    
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from pathlib import Path

DEFAULT_COMPANY_ADDRESS = "3250 Wilshire Blvd, STE1502\nLos Angeles, CA 90010"
FONT_PATH = Path(__file__).resolve().parent / "micr-e13b.ttf"
FONT_OK = False
if FONT_PATH.exists():
    try:
        pdfmetrics.registerFont(TTFont("MICR_E13B", str(FONT_PATH)))
        FONT_OK = True
    except Exception as e:
        st.warning(f"MICR font error: {e}")

# Optional secure account registry in Streamlit Secrets, not GitHub.
# [BANK_ACCOUNTS."ACC-8652"] Bank_Name = "..." etc.
def bank_registry():
    registry = {}
    try:
        secret = st.secrets.get("BANK_ACCOUNTS", {})
        for alias, item in dict(secret).items():
            registry[str(alias)] = {
                "Bank Name": str(item.get("Bank_Name", item.get("Bank Name", ""))),
                "Routing Number": str(item.get("Routing_Number", item.get("Routing Number", ""))),
                "Account Number": str(item.get("Account_Number", item.get("Account Number", ""))),
            }
    except Exception:
        pass
    return registry

def bank_from_project(alias, project=""):
    result = bank_registry().get(str(alias), {}).copy()
    if not df_projects.empty and project in df_projects["Project_Name"].values:
        p = df_projects[df_projects["Project_Name"] == project].iloc[0]
        for dest, candidates in {
            "Bank Name": ["Bank_Name", "Bank Name"],
            "Routing Number": ["Routing_Number", "Routing Number"],
            "Account Number": ["Account_Number", "Account Number", "Bank_Account_Number"],
        }.items():
            if not result.get(dest):
                for col in candidates:
                    if str(p.get(col, "")).strip():
                        result[dest] = str(p[col]).strip()
                        break
    return result

def resolve_bank(alias, project=""):
    result = bank_from_project(alias, project)
    overrides = st.session_state.get("bank_overrides", {})
    result.update({k:v for k,v in overrides.get(str(alias), {}).items() if str(v).strip()})
    return result

def ensure_bank(alias, project=""):
    bank = resolve_bank(alias, project)
    missing = [k for k in ["Bank Name", "Routing Number", "Account Number"] if not bank.get(k)]
    if missing:
        raise ValueError(f"Account {alias}: missing {', '.join(missing)}. Configure in sidebar or Streamlit Secrets.")
    routing = str(bank["Routing Number"]).strip()
    acct = str(bank["Account Number"]).strip()
    if not re.fullmatch(r"\d{9}", routing):
        raise ValueError(f"Account {alias}: Routing Number must be exactly 9 digits")
    if not re.fullmatch(r"\d+", acct):
        raise ValueError(f"Account {alias}: Account Number must contain digits only")
    return bank

def generate_check_pdf(company, account_alias, project, check_no, pay_date, payee, amount, memo):
    bank = ensure_bank(account_alias, project)
    info = {
        "Company": str(company),
        "Company Address": DEFAULT_COMPANY_ADDRESS,
        "Bank Name": bank["Bank Name"],
        "Routing Number": bank["Routing Number"],
        "Account Number": bank["Account Number"],
    }
    payment = {"Check Date": pay_date, "Payee": payee, "Amount": Decimal(str(amount)), "Memo": memo}
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=letter)
    c.setTitle("Check layout proof - not negotiable")
    draw_check(c, payment, info, int(check_no))
    c.save()
    return out.getvalue()

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
    txt(13, 211, "MEMO", "Helvetica-Bold", 9)
    txt(46, 211, memo, "Helvetica", 10, max_width=290)

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


# 加载云端预设数据
df_projects = load_project_presets()
df_workers = load_worker_presets()
df_stages = load_stage_presets()

preset_worker_list = df_workers["Worker_Name"].dropna().tolist() if not df_workers.empty else []
preset_project_list = df_projects["Project_Name"].dropna().tolist() if not df_projects.empty else []

latest_check_map = fetch_next_check_numbers_from_gs(preset_project_list)

# ----------------- 通用 Stage 联动选择组件 (升级支持强制控制 UI) -----------------
def render_stage_selector(key_prefix="default", default_stage="", default_stage_name="", default_sub_stage=""):
    """
    支持根据 Session State 强制更新 UI 选择索引的三级 Stage 联动组件
    """
    if df_stages.empty:
        st.warning("Stage 配置数据为空")
        return "", "", "", None

    df_stages_temp = df_stages.copy()
    df_stages_temp["Display_Stage"] = df_stages_temp["Stage"] + ": " + df_stages_temp["Stage_Name"]
    unique_stages = df_stages_temp["Display_Stage"].unique().tolist()

    # 如果系统强制设置了该 Key 的 Session State，则使用它；否则使用传入的 default
    main_stage_key = f"{key_prefix}_main_stage"
    sub_stage_1_key = f"{key_prefix}_sub_stage_1"
    sub_stage_2_key = f"{key_prefix}_sub_stage_2"

    target_display = f"{default_stage}: {default_stage_name}"
    
    # 检测是否在 Session State 中强行改写了值
    if main_stage_key in st.session_state and st.session_state[main_stage_key] in unique_stages:
        stage_idx = unique_stages.index(st.session_state[main_stage_key])
    else:
        stage_idx = unique_stages.index(target_display) if target_display in unique_stages else 0

    col_s1, col_s2, col_s3 = st.columns([2, 2, 2])

    with col_s1:
        selected_stage_display = st.selectbox(
            "Stage (主阶段)",
            options=unique_stages,
            index=stage_idx,
            key=main_stage_key
        )

    filtered_df = df_stages_temp[df_stages_temp["Display_Stage"] == selected_stage_display]
    sub_stage_options = filtered_df["Sub_Stage"].dropna().tolist()

    if sub_stage_1_key in st.session_state and st.session_state[sub_stage_1_key] in sub_stage_options:
        sub1_idx = sub_stage_options.index(st.session_state[sub_stage_1_key])
    else:
        sub1_idx = sub_stage_options.index(default_sub_stage) if default_sub_stage in sub_stage_options else 0

    with col_s2:
        sub_stage_1 = st.selectbox(
            "Sub-Stage 1 (必选)",
            options=sub_stage_options,
            index=sub1_idx,
            key=sub_stage_1_key
        )

    with col_s3:
        remaining_options = ["None (无)"] + [item for item in sub_stage_options if item != sub_stage_1]
        sub_stage_2 = st.selectbox(
            "Sub-Stage 2 (可选)",
            options=remaining_options,
            index=0,
            key=sub_stage_2_key
        )

    selected_stage_code = filtered_df["Stage"].iloc[0]
    selected_stage_name = filtered_df["Stage_Name"].iloc[0]
    opt_sub_stage_2 = sub_stage_2 if sub_stage_2 != "None (无)" else None

    return selected_stage_code, selected_stage_name, sub_stage_1, opt_sub_stage_2

# ----------------- 核心工具函数 -----------------
def number_to_words_usd(amount):
    """金额转英文大写"""
    try:
        dollars = int(amount)
        cents = int(round((amount - dollars) * 100))
        words = num2words(dollars, lang="en").title()
        return f"{words} and {cents:02d}/100 Dollars"
    except Exception:
        return ""


# Real bank routing/account data must be supplied separately from internal aliases.
st.sidebar.markdown("### Bank account details")
st.sidebar.caption("ACC-8652 / Chase-1185 are internal aliases, not MICR account numbers.")
known_aliases = {"ACC-8652", "ACC-3738", "ACC-5027", "Chase-1185"}
if not df_projects.empty and "Account" in df_projects.columns:
    known_aliases.update(str(x).strip() for x in df_projects["Account"].dropna() if str(x).strip())
chosen_alias = st.sidebar.selectbox("Account alias", sorted(known_aliases))
known = resolve_bank(chosen_alias)
with st.sidebar.expander("Edit bank information", expanded=False):
    bank_name_edit = st.text_input("Bank Name", value=known.get("Bank Name", ""), key=f"bankname_{chosen_alias}")
    routing_edit = st.text_input("Routing Number (9 digits)", value=known.get("Routing Number", ""), key=f"routing_{chosen_alias}")
    account_edit = st.text_input("Real Account Number", value=known.get("Account Number", ""), key=f"account_{chosen_alias}")
    if st.button("Apply for this session"):
        st.session_state.setdefault("bank_overrides", {})[chosen_alias] = {
            "Bank Name": bank_name_edit.strip(),
            "Routing Number": routing_edit.strip(),
            "Account Number": account_edit.strip(),
        }
        st.success("Applied for current session")
st.sidebar.caption("For persistent storage, use BANK_ACCOUNTS in Streamlit Secrets. Do not commit real bank details to GitHub.")

# ----------------- 页面架构 -----------------
st.sidebar.title("⚙ 系统导航")
mode = st.sidebar.radio(
    "请选择业务场景：",
    [
        "📝 Single Mannul Check",
        "👷 Construction Bulk Checks",
    ],
)

# ==============================================================================
# 场景 1：单张手动生成支票
# ==============================================================================
if mode == "📝 Single Mannul Check":
    st.title("📝 Single Mannul Check")
    st.caption("Please enter the information")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Information")

        main_category = st.radio(
            "Business Category：",
            ["🏗 Construction", "🏠 Moo Housing"],
            index=0,
            horizontal=True,
        )

        st.markdown("---")

        def_stg, def_stg_name, def_sub_stg = "", "", ""

        # 分支 1：Construction 逻辑
        if main_category == "🏗 Construction":
            project_options = preset_project_list + ["+ New Project"]
            
            # 定义回调函数：当选择 Project 或 Payee 发生变化时，强行改变 Stage 下拉框的值
            def update_single_check_stage():
                cur_payee = st.session_state.get("single_payee_select", "")
                cur_proj = st.session_state.get("single_proj_select", "")
                
                # 特殊拦截：如果是 Valente Herrera，默认固定为 Salary
                if cur_payee == "Valente Herrera":
                    st.session_state["single_check_main_stage"] = "Salary"  # 根据你的实际下拉选项调整，如 "99: Salary" 或 "Salary"
                    st.session_state["single_check_sub_stage_1"] = "Payroll" # 根据实际子阶段调整
                    return
        
                # 获取原默认值
                orig_stg, orig_sname, orig_sub = "", "", ""
                if not df_workers.empty and cur_payee in df_workers["Worker_Name"].values:
                    w_info = df_workers[df_workers["Worker_Name"] == cur_payee].iloc[0]
                    orig_stg = w_info.get("Stage", "")
                    orig_sname = w_info.get("Stage_Name", "")
                    orig_sub = w_info.get("Sub_Stage", "")
                
                # 计算拦截联动值
                auto_stg, auto_sname, auto_sub = auto_match_stage(cur_payee, cur_proj, orig_stg, orig_sname, orig_sub)
                
                # 强行重置界面 Session State 中的下拉框属性
                st.session_state["single_check_main_stage"] = f"{auto_stg}: {auto_sname}"
                st.session_state["single_check_sub_stage_1"] = auto_sub
        
            # 收款人选择模式 (提到前面先渲染/选择，以便 Project 根据 Payee 动态调整，也可以保持原顺序)
            payee_mode = st.radio(
                "Payee Input Mode",
                ["List Selection", "Custom Input"],
                index=0,
                horizontal=True
            )
        
            if payee_mode == "List Selection" and preset_worker_list:
                payee_name = st.selectbox(
                    "Payee Name", 
                    options=preset_worker_list,
                    key="single_payee_select",
                    on_change=update_single_check_stage
                )
                if not df_workers.empty and payee_name in df_workers["Worker_Name"].values:
                    w_info = df_workers[df_workers["Worker_Name"] == payee_name].iloc[0]
                    def_stg = w_info.get("Stage", "")
                    def_stg_name = w_info.get("Stage_Name", "")
                    def_sub_stg = w_info.get("Sub_Stage", "")
            else:
                payee_name = st.text_input("Payee Name", value="", placeholder="Enter payee full name")
        
            # 特殊处理：判断是否为固定薪资工人
            is_valente = (payee_name == "Valente Herrera")
        
            # 根据 Payee 判断 Project 输入逻辑
            if is_valente:
                # 如果是 Valente Herrera，无需选择 Project
                selected_proj = "None (Salary)"
                project_site = "N/A"
                default_chk_val = 1001
                project_account = "ACC-5027"
                project_company = "Development Company"
                st.warning("⚠ **Valente Herrera** 默认为内部固定薪资发薪（Salary），不关联任何具体施工项目。")
            else:
                selected_proj = st.selectbox(
                    "Project", 
                    project_options, 
                    key="single_proj_select",
                    on_change=update_single_check_stage
                )
        
                if selected_proj != "+ New Project":
                    project_site = selected_proj
                    default_chk_val = latest_check_map.get(selected_proj, 1001)
                    p_info = df_projects[df_projects["Project_Name"] == selected_proj].iloc[0]
                    project_account = str(p_info.get("Account", "ACC-5027")).strip()
                    project_company = str(p_info.get("Company", "Development Company")).strip()
                else:
                    project_site = st.text_input("Project Name", value="New Site")
                    default_chk_val = 1001
                    project_account = "ACC-5027"
                    project_company = "Development Company"
        
            payer_entity = st.selectbox(
                "Payer Entity",
                options=["Development Company", "Moo Construction"],
                index=0,
                help="选择付款主体"
            )
        
            if payer_entity == "Moo Construction":
                company_name = "Moo Construction"
                account_num = "Chase-1185"
                st.info("💡 **Moo Construction** 付款账户已自动固定为: `Chase-1185`")
            else:
                company_name = project_company
                account_num = project_account
                st.info(f"💡 **Development Company** 已自动使用项目对应的公司 (`{company_name}`) 与账户 (`{account_num}`)")

        # 分支 2：Moo Housing 逻辑
        else:
            payer_entity = "Moo Housing"
            company_name = "Moo Housing Inc"
            project_site = "Moo Housing"
            default_chk_val = latest_check_map.get("Moo Housing", 1001)

            moo_acc_options = ["ACC-8652", "ACC-3738", "Other"]
            acc_choice = st.selectbox("Bank Account Choice", moo_acc_options, index=0)
            if acc_choice == "Other":
                account_num = st.text_input("Enter Custom Bank Account", value="ACC-8652")
            else:
                account_num = acc_choice

            payee_name = st.text_input("Payee Name", value="", placeholder="Enter payee full name")

        company_display = st.text_input("Company Display Name", value=company_name)

        st.caption("PDF is generated directly; no check template files required.")

        st.markdown("---")

        if main_category == "🏠 Moo Housing":
            default_memo_text = "Deposit Refund"
            selected_stage_str = ""
        else:
            def_stg, def_stg_name, def_sub_stg = auto_match_stage(
                payee_name, project_site, def_stg, def_stg_name, def_sub_stg
            )

            st.markdown("##### 🏗 工程阶段 (Stage Selection)")
            st_code, st_name, sub1, sub2 = render_stage_selector(
                key_prefix="single_check",
                default_stage=def_stg,
                default_stage_name=def_stg_name,
                default_sub_stage=def_sub_stg
            )
            sub_str = f"{sub1}, {sub2}" if sub2 else sub1
            selected_stage_str = f"{st_code} ({sub_str})" if st_code else ""

            default_memo_text = f"{st_name} - {sub1}" if sub1 else st_name

        user_memo = st.text_input("Memo Detail", value=default_memo_text)

        pay_amount = st.number_input("Amount", min_value=0.01, value=1500.00, step=100.0)

        c_a, c_b = st.columns(2)
        with c_a:
            pay_date = st.date_input("Date", value=date.today())
        with c_b:
            check_num = st.number_input(
                "Check Number",
                min_value=1,
                value=int(default_chk_val),
                step=1,
                help="Automatically generated by System"
            )

        if main_category == "🏠 Moo Housing":
            memo_text = user_memo.strip()
        else:
            if project_site and user_memo.strip():
                memo_text = f"{project_site} - {user_memo.strip()}"
            elif project_site:
                memo_text = project_site
            else:
                memo_text = user_memo.strip()

        amount_words = number_to_words_usd(pay_amount)

    try:
        filled_pdf = generate_check_pdf(company_display, account_num, project_site, check_num, pay_date, payee_name, pay_amount, memo_text)
        pdf_error = None
    except ValueError as e:
        filled_pdf = None
        pdf_error = str(e)

    with col2:
        st.subheader("👁 Check Preview")
        st.markdown(f"""
        > **Company**: {company_display}  
        > **Bank Account**: `{account_num}`  
        > **PDF**: Directly generated  
        > **Project / Stage**: {project_site} | **`{selected_stage_str}`**  
        > **Check Number**: `#{check_num}`  
        > **Date**: {pay_date.strftime("%Y-%m-%d")}  
        > **Payee**: **{payee_name}**  
        > **Amount**: **${pay_amount:,.2f}**  
        > **Amount Words**: *{amount_words}*  
        > **Memo**: {memo_text}
        """)

        def handle_sync_and_download():
            record = [
                {
                    "Check Number": check_num,
                    "Issue Date": pay_date.strftime("%Y-%m-%d"),
                    "Company": company_display,
                    "Account": account_num,
                    "Project": project_site,
                    "Stage": selected_stage_str,
                    "Payee Name": payee_name,
                    "Amount": pay_amount,
                    "Memo": memo_text,
                }
            ]
            if save_to_history(record):
                st.session_state["sync_success_msg"] = f"🎉 Check #{check_num} Successfully saved & transferred to Google Sheets!"
            else:
                st.session_state["sync_error_msg"] = f"⚠ Check #{check_num} PDF downloaded, but failed to sync to Google Sheets."

        if pdf_error:
            st.error(pdf_error)
        st.download_button(
            label=f"🚀 Save to Sheets & Download PDF (#{check_num})",
            data=filled_pdf or b"",
            disabled=filled_pdf is None,
            file_name=f"Check_{check_num}_{payee_name}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
            on_click=handle_sync_and_download
        )

        if "sync_success_msg" in st.session_state:
            st.balloons()
            st.success(st.session_state.pop("sync_success_msg"))
            
        if "sync_error_msg" in st.session_state:
            st.error(st.session_state.pop("sync_error_msg"))

# ==============================================================================
# 场景 2：多项目/施工队周薪批量开单
# ==============================================================================
elif mode == "👷 Construction Bulk Checks":
    st.title("👷 Construction Bulk Checks")


    pay_date = st.date_input("Date", value=date.today())

    st.markdown("---")

    # ----------------- 1. 确认各项目起始支票号 -----------------
    st.subheader("1. Confirm the start check number")

    proj_start_nums = {}
    cols = st.columns(min(max(len(df_projects), 1), 4))
    for idx, p_row in df_projects.iterrows():
        p_name = p_row["Project_Name"]
        default_num = latest_check_map.get(p_name, 1001)
        with cols[idx % 4]:
            proj_start_nums[p_name] = st.number_input(
                f"🏗 {p_name}",
                min_value=1,
                value=int(default_num),
                key=f"start_num_{p_name}"
            )

    st.markdown("---")

    # ----------------- 2. 快捷添加面板 -----------------
    st.subheader("2. Information")

    if "payroll_list" not in st.session_state:
        st.session_state.payroll_list = []

    def update_bulk_stage_and_memo():
        selected_w = st.session_state.get("input_w", "")
        selected_p = st.session_state.get("input_p", "")
        
        if selected_w == "Valente Herrera":
            st.session_state.input_p = ""
            selected_p = ""
        
        w_stg, w_stg_name, w_sub_stg = "", "", ""
        if not df_workers.empty and selected_w in df_workers["Worker_Name"].values:
            w_info = df_workers[df_workers["Worker_Name"] == selected_w].iloc[0]
            w_stg = w_info.get("Stage", "")
            w_stg_name = w_info.get("Stage_Name", "")
            w_sub_stg = w_info.get("Sub_Stage", "")
            
        auto_stg, auto_sname, auto_sub = auto_match_stage(selected_w, selected_p, w_stg, w_stg_name, w_sub_stg)
        
        st.session_state["bulk_check_main_stage"] = f"{auto_stg}: {auto_sname}"
        st.session_state["bulk_check_sub_stage_1"] = auto_sub
        
        st.session_state.input_m = f"{auto_sname} - {auto_sub}" if auto_sub else auto_sname

    def update_account_and_company():
        c_type = st.session_state.get("input_company_type", "Development Company")
        selected_p = st.session_state.get("input_p", "")

        if c_type == "Moo Construction":
            st.session_state.input_acc = "Chase-1185"
        elif c_type == "Moo Housing":
            st.session_state.input_acc = "ACC-8652"
        else:
            if not df_projects.empty and selected_p in df_projects["Project_Name"].values:
                p_info = df_projects[df_projects["Project_Name"] == selected_p].iloc[0]
                st.session_state.input_acc = str(p_info.get("Account", "ACC-8652")).strip()
        
        update_bulk_stage_and_memo()

    def calculate_amount_from_days():
        days = st.session_state.get("input_days", 0.0)
        rate = st.session_state.get("input_rate", 0.0)
        if days > 0 and rate > 0:
            st.session_state.input_a = round(days * rate, 2)

    if "input_company_type" not in st.session_state:
        st.session_state.input_company_type = "Development Company"

    if "input_acc" not in st.session_state and preset_project_list:
        first_p = preset_project_list[0]
        p_info = df_projects[df_projects["Project_Name"] == first_p].iloc[0]
        st.session_state.input_acc = str(p_info.get("Account", "ACC-5027")).strip()

    if "input_m" not in st.session_state and preset_worker_list:
        first_w = preset_worker_list[0]
        first_p = preset_project_list[0] if preset_project_list else ""
        if not df_workers.empty and first_w in df_workers["Worker_Name"].values:
            w_info = df_workers[df_workers["Worker_Name"] == first_w].iloc[0]
            _, auto_sname, auto_sub = auto_match_stage(first_w, first_p, w_info.get("Stage", ""), w_info.get("Stage_Name", ""), w_info.get("Sub_Stage", ""))
            st.session_state.input_m = f"{auto_sname} - {auto_sub}" if auto_sub else auto_sname

    st.markdown("##### ➕ New Check")
    
    r1_c1, r1_c2, r1_c3, r1_c4 = st.columns([1.2, 1, 1, 1])
    with r1_c1:
        add_comp_type = st.selectbox(
            "Paying Entity",
            options=["Development Company", "Moo Housing", "Moo Construction"],
            index=0,
            key="input_company_type",
            on_change=update_account_and_company
        )
    with r1_c2:
        add_worker = st.selectbox(
            "Payee Name", 
            preset_worker_list, 
            key="input_w",
            on_change=update_bulk_stage_and_memo
        )

    is_valente = (add_worker == "Valente Herrera")
    project_options = [""] if is_valente else preset_project_list

    with r1_c3:
        add_proj = st.selectbox(
            "Project", 
            project_options, 
            key="input_p",
            disabled=is_valente,
            on_change=update_account_and_company
        )
    with r1_c4:
        add_account = st.text_input("Bank Account", key="input_acc")

    w_stg, w_stg_name, w_sub_stg = "", "", ""
    if not df_workers.empty and add_worker in df_workers["Worker_Name"].values:
        w_info = df_workers[df_workers["Worker_Name"] == add_worker].iloc[0]
        w_stg = w_info.get("Stage", "")
        w_stg_name = w_info.get("Stage_Name", "")
        w_sub_stg = w_info.get("Sub_Stage", "")

    w_stg, w_stg_name, w_sub_stg = auto_match_stage(
        add_worker, add_proj, w_stg, w_stg_name, w_sub_stg
    )

    st_code, st_name, sub1, sub2 = render_stage_selector(
        key_prefix="bulk_check",
        default_stage=w_stg,
        default_stage_name=w_stg_name,
        default_sub_stage=w_sub_stg
    )

    r2_c1, r2_c2, r2_c3, r2_c4, r2_c5 = st.columns([1.5, 1.5, 2.0, 3.0, 1.2])
    with r2_c1:
        add_days = st.number_input(
            "Days (Opt)", 
            min_value=0.0, 
            value=0.0, 
            step=0.5, 
            key="input_days",
            on_change=calculate_amount_from_days
        )
    with r2_c2:
        add_rate = st.number_input(
            "Rate/Day (Opt)", 
            min_value=0.0, 
            value=0.0, 
            step=10.0, 
            key="input_rate",
            on_change=calculate_amount_from_days
        )
    with r2_c3:
        add_amt = st.number_input("Total Amount", min_value=0.01, value=1200.00, step=50.0, key="input_a")
    with r2_c4:
        add_memo = st.text_input("Memo Detail", key="input_m")
    with r2_c5:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Add", type="primary", use_container_width=True):
            final_project_val = "" if add_worker == "Valente Herrera" else add_proj

            existing_checks = [
                row["Check #"] 
                for row in st.session_state.payroll_list 
                if row["Project"] == final_project_val
            ]
            
            if existing_checks:
                next_chk = max(existing_checks) + 1
            else:
                next_chk = proj_start_nums.get(final_project_val, 1001)

            stage_val_str = f"{st_code} ({f'{sub1}, {sub2}' if sub2 else sub1})" if st_code else ""

            if add_comp_type == "Development Company":
                if not df_projects.empty and final_project_val in df_projects["Project_Name"].values:
                    p_row_info = df_projects[df_projects["Project_Name"] == final_project_val].iloc[0]
                    final_company_val = str(p_row_info.get("Company", "Development Company")).strip()
                else:
                    final_company_val = "Development Company"
            else:
                final_company_val = add_comp_type

            st.session_state.payroll_list.append({
                "Company": final_company_val,
                "Payee": add_worker,
                "Project": final_project_val,
                "Account": add_account,
                "Stage": stage_val_str,
                "Days": add_days if add_days > 0 else None,
                "Rate": add_rate if add_rate > 0 else None,
                "Check #": int(next_chk),
                "Amount": add_amt,
                "Memo": add_memo
            })
            st.rerun()

    st.markdown("---")

    # ----------------- 3. 待打款列表预览与编辑 -----------------
    st.subheader("3. Review & Edit")

    col_title, col_clear = st.columns([8, 2])
    with col_clear:
        if st.session_state.payroll_list:
            if st.button("🗑 Clear List", type="secondary", use_container_width=True):
                st.session_state.payroll_list = []
                st.rerun()

    df_payroll_input = pd.DataFrame(st.session_state.payroll_list)

    if not df_payroll_input.empty:
        edited_df = st.data_editor(
            df_payroll_input,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Company": st.column_config.SelectboxColumn("Company", options=["Development Company", "Moo Housing", "Moo Construction"], required=True),
                "Payee": st.column_config.SelectboxColumn("Payee", options=preset_worker_list, required=True),
                "Project": st.column_config.SelectboxColumn("Project", options=[""] + preset_project_list),
                "Check #": st.column_config.NumberColumn("Check #", format="%d"),
                "Amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
                "Days": st.column_config.NumberColumn("Days", format="%.1f"),
                "Rate": st.column_config.NumberColumn("Rate", format="$%.2f")
            },
            key="payroll_data_editor"
        )
        df_payroll_input = edited_df
    else:
        st.info("💡 暂无待处理支票，请在上方添加数据。")

    st.markdown("---")

    # ----------------- 4. 批量生成与导出 -----------------
    if not df_payroll_input.empty:
        if st.button(f"🚀 Confirm & Batch Generate {len(df_payroll_input)} Checks", type="primary", use_container_width=True):
            account_pdf_dict = {}
            records_log = []
    
            proj_map = df_projects.set_index("Project_Name").to_dict(orient="index") if not df_projects.empty else {}
    
            for idx, row in df_payroll_input.iterrows():
                worker_name = str(row.get("Payee", "")).strip()
                project_name = str(row.get("Project", "")).strip()
                stage_name = str(row.get("Stage", "")).strip()
                
                try:
                    cur_check = int(row.get("Check #", 0))
                    amt = float(row.get("Amount", 0.0))
                except (ValueError, TypeError):
                    cur_check = 0
                    amt = 0.0
    
                detail_memo = str(row.get("Memo", "")).strip()
    
                if amt <= 0 or not worker_name or cur_check <= 0:
                    continue
    
                p_info = proj_map.get(project_name, {"Account": "ACC-8652", "Company": "Development Company"})
                
                company_name = str(row.get("Company", "")).strip()
                if not company_name or company_name == "Development Company":
                    company_name = str(p_info.get("Company", "Development Company")).strip()
                
                if company_name == "Moo Construction":
                    account_num = "Chase-1185"
                elif company_name == "Moo Housing":
                    account_num = str(row.get("Account", "")).strip() or "ACC-8652"
                else:
                    account_num = str(row.get("Account", "")).strip() or str(p_info.get("Account", "ACC-8652")).strip()
    
                if project_name and detail_memo:
                    full_memo = f"{project_name} - {detail_memo}"
                elif project_name:
                    full_memo = project_name
                else:
                    full_memo = detail_memo
    
                replacements = {
                    "date": pay_date.strftime("%m/%d/%Y"),
                    "name": worker_name,
                    "amount": f"{amt:,.2f}",
                    "amount_words": number_to_words_usd(amt),
                    "memo": full_memo,
                    "number": str(cur_check),
                    "account": account_num
                }
    
                try:
                    pdf_res = generate_check_pdf(company_name, account_num, project_name, cur_check, pay_date, worker_name, amt, full_memo)
                except ValueError as e:
                    st.error(f"Row {idx+1}: {e}")
                    st.stop()
                
                acc_key = (company_name, account_num)
                if acc_key not in account_pdf_dict:
                    account_pdf_dict[acc_key] = []
                account_pdf_dict[acc_key].append((cur_check, project_name, worker_name, pdf_res))
    
                records_log.append({
                    "Check Number": cur_check,
                    "Issue Date": pay_date.strftime("%Y-%m-%d"),
                    "Company": company_name,
                    "Account": account_num,
                    "Project": project_name,
                    "Stage": stage_name,
                    "Payee Name": worker_name,
                    "Amount": amt,
                    "Memo": full_memo
                })
    
            if records_log:
                # ----------------- 新增：按账户合并 PDF -----------------
                merged_account_pdfs = {}
    
                for (company, acc_num), items in account_pdf_dict.items():
                    # 按照 Check Number (cur_check) 升序排序，保证连号输出
                    sorted_items = sorted(items, key=lambda x: x[0])
                    
                    merger = PdfMerger()
                    for item in sorted_items:
                        pdf_bytes = item[3]  # 单张支票的 PDF Bytes
                        merger.append(io.BytesIO(pdf_bytes))
                    
                    merged_output = io.BytesIO()
                    merger.write(merged_output)
                    merger.close()
                    merged_output.seek(0)
                    
                    # 存储合并后的 PDF 二进制数据
                    merged_account_pdfs[(company, acc_num)] = merged_output.getvalue()
    
                # 将合并后的 PDF 结果存入 session_state
                st.session_state.last_generated_pdfs = merged_account_pdfs
                # ------------------------------------------------------
    
                if save_to_history(records_log):
                    st.session_state.payroll_list = []
    
                    st.balloons()
                    st.success(f"🎉 Successfully generated {len(records_log)} check(s)! Data synced to Google Sheets.")
    
                    st.markdown("### 📊 Current Period Disbursement Summary")
                    df_batch = pd.DataFrame(records_log)
                    col_sum1, col_sum2 = st.columns(2)
    
                    with col_sum1:
                        st.markdown("#### 🏢 Summary by Company / Account")
                        summary_company = df_batch.groupby(["Company", "Account"]).agg(
                            **{
                                "Total Amount": ("Amount", "sum"),
                                "Total Number": ("Check Number", "count")
                            }
                        ).reset_index()
                        st.dataframe(summary_company.style.format({"Total Amount": "${:,.2f}"}), use_container_width=True, hide_index=True)
    
                    with col_sum2:
                        st.markdown("#### 🏗 Summary by Project")
                        summary_project = df_batch.groupby(["Project", "Company"]).agg(
                            **{
                                "Total Labor Cost": ("Amount", "sum"),
                                "Worker Count": ("Check Number", "count")
                            }
                        ).reset_index()
                        st.dataframe(summary_project.style.format({"Total Labor Cost": "${:,.2f}"}), use_container_width=True, hide_index=True)
    
                    st.markdown("### 📥 Download Merged PDF Checks by Account")
                    for (company, acc_num), pdf_data in merged_account_pdfs.items():
                        file_name = f"Checks_{company}_{acc_num}_{pay_date.strftime('%Y%m%d')}.pdf"
                        st.download_button(
                            label=f"📄 Download Combined PDF for {company} ({acc_num})",
                            data=pdf_data,
                            file_name=file_name,
                            mime="application/pdf"
                        )
