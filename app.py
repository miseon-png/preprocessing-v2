import base64
import copy
import io
import os
import re
import sqlite3
import traceback
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="샐러드 생산일보 및 원료코드 System", layout="wide")

DB_FILE = "stock_master.db"

# ---------------------------------------------------------
# 구글 시트 연동 함수
# ---------------------------------------------------------
@st.cache_resource
def init_gspread():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    gcp_secrets = st.secrets["gcp_service_account"]
    private_key = gcp_secrets.get("private_key", "")
    if "\\n" in private_key:
        private_key = private_key.replace("\\n", "\n")
    
    creds_dict = {
        "type": "service_account",
        "project_id": gcp_secrets.get("project_id", ""),
        "private_key_id": gcp_secrets.get("private_key_id", ""),
        "private_key": private_key,
        "client_email": gcp_secrets.get("client_email", ""),
        "client_id": gcp_secrets.get("client_id", ""),
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": gcp_secrets.get("client_x509_cert_url", "")
    }

    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    url = "https://docs.google.com/spreadsheets/d/1WM472gdj_7HBUvJdLRgwIU0t7GAMnZ47GpsCJfB5pBo/edit"
    return client.open_by_url(url)

# ---------------------------------------------------------
# DB 관리 함수
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS closed_lots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_name TEXT,
            yymmdd TEXT,
            avail_g REAL,
            vendor TEXT,
            closed_ym TEXT
        )
    """)
    conn.commit()
    conn.close()

def clear_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM closed_lots")
    conn.commit()
    conn.close()

def get_latest_closed_lots():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(closed_ym) FROM closed_lots")
    max_ym = cursor.fetchone()[0]
    
    lots = {"양상추": [], "양배추": [], "적채": []}
    if max_ym:
        cursor.execute(
            "SELECT item_name, yymmdd, avail_g, vendor FROM closed_lots WHERE closed_ym = ?",
            (max_ym,)
        )
        rows = cursor.fetchall()
        for item_name, yymmdd, avail_g, vendor in rows:
            if item_name in lots:
                lots[item_name].append({
                    "yymmdd": yymmdd,
                    "avail_g": avail_g,
                    "avail": avail_g,
                    "vendor": vendor,
                    "label": f"{yymmdd[2:4]}/{yymmdd[4:]} 이월",
                    "is_from_db": True
                })
    conn.close()
    return max_ym, lots

def save_closed_lots(closed_ym, item_lots):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM closed_lots WHERE closed_ym = ?", (closed_ym,))
    
    for item_name, lot_list in item_lots.items():
        for lot in lot_list:
            avail_val = lot.get("avail", lot.get("avail_g", 0.0))
            if avail_val > 0:
                cursor.execute(
                    "INSERT INTO closed_lots (item_name, yymmdd, avail_g, vendor, closed_ym) VALUES (?, ?, ?, ?, ?)",
                    (item_name, lot["yymmdd"], avail_val, lot.get("vendor", "K01"), closed_ym)
                )
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# 출력 렌더링 함수
# ---------------------------------------------------------
def render_direct_print_button(html_payload, button_label="🖨️ 바로 출력 / PDF"):
    b64 = base64.b64encode(html_payload.encode('utf-8')).decode('utf-8')
    button_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{ margin: 0; padding: 0; font-family: "Source Sans Pro", sans-serif; overflow: hidden; }}
        .btn {{
            display: inline-flex; align-items: center; justify-content: center;
            font-weight: 400; padding: 0.25rem 0.75rem; border-radius: 0.5rem;
            min-height: 38.4px; margin: 0px; line-height: 1.6;
            width: 100%; cursor: pointer; font-size: 16px; box-sizing: border-box;
            transition: all 0.2s ease;
        }}
        @media (prefers-color-scheme: dark) {{
            body {{ background-color: transparent; }}
            .btn {{ background-color: #262730; color: #FAFAFA; border: 1px solid rgba(250, 250, 250, 0.2); }}
            .btn:hover {{ border-color: #FF4B4B; color: #FF4B4B; }}
        }}
        @media (prefers-color-scheme: light) {{
            body {{ background-color: transparent; }}
            .btn {{ background-color: #FFFFFF; color: #31333F; border: 1px solid rgba(49, 51, 63, 0.2); }}
            .btn:hover {{ border-color: #FF4B4B; color: #FF4B4B; }}
        }}
    </style>
    <script>
        function directPrint() {{
            try {{
                var b64 = "{b64}";
                var decoded = decodeURIComponent(escape(window.atob(b64)));
                var iframe = document.createElement('iframe');
                iframe.style.position = 'absolute';
                iframe.style.width = '0px';
                iframe.style.height = '0px';
                iframe.style.border = 'none';
                document.body.appendChild(iframe);
                var doc = iframe.contentWindow.document;
                doc.open();
                doc.write(decoded);
                doc.close();
                iframe.contentWindow.focus();
                setTimeout(function() {{ iframe.contentWindow.print(); }}, 500);
            }} catch (e) {{
                alert("인쇄 화면을 여는 중 오류가 발생했습니다.");
            }}
        }}
    </script>
    </head>
    <body>
        <button class="btn" onclick="directPrint()">{button_label}</button>
    </body>
    </html>
    """
    components.html(button_html, height=45)

def render_printable_document(title, df_data):
    html_table = df_data.to_html(index=False, classes="report-table")
    # HTML 내 줄바꿈(\n)을 <br>로 변환하여 엑셀 렌더링 시에도 예쁘게 보이도록 처리
    html_table = html_table.replace('\\n', '<br>')
    
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <title>{title}</title>
    <style>
        body {{ font-family: '맑은 고딕', sans-serif; padding: 20px; color: #000; background-color: #fff; }}
        .header-container {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }}
        .report-title {{ font-size: 20pt; font-weight: bold; color: #1F4E79; margin: 0; }}
        .approval-table {{ border-collapse: collapse; text-align: center; font-size: 9pt; }}
        .approval-table th, .approval-table td {{ border: 1px solid #333; padding: 4px 8px; }}
        .approval-table th {{ background-color: #f2f2f2; width: 50px; }}
        .approval-table td {{ height: 40px; width: 60px; }}
        .report-table {{ width: 100%; border-collapse: collapse; font-size: 9pt; margin-top: 10px; }}
        .report-table th, .report-table td {{ border: 1px solid #333; padding: 6px 4px; text-align: center; }}
        .report-table th {{ background-color: #D9E1F2; font-weight: bold; color: #1F4E79; }}
        .report-table tr:nth-child(even) {{ background-color: #F9FAFB; }}
        @media print {{
            body {{ padding: 0; margin: 0; }}
            @page {{ size: A4 portrait; margin: 15mm; }}
        }}
    </style>
    </head>
    <body>
        <div class="header-container">
            <h1 class="report-title">{title}</h1>
            <table class="approval-table">
                <tr><th rowspan="2">결<br>재</th><th>담당</th><th>검토</th><th>승인</th></tr>
                <tr><td></td><td></td><td></td></tr>
            </table>
        </div>
        {html_table}
    </body>
    </html>
    """
    return full_html

def render_process_log_html(date_str, df_target, is_completed=False):
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        yyyy, mm, dd = dt.strftime("%Y"), dt.strftime("%m"), dt.strftime("%d")
    except:
        yyyy, mm, dd = "202 ", "  ", "  "
    
    row_count = len(df_target)
    rows_html = ""
    
    if row_count == 0:
        rows_html = "<tr><td colspan='9' style='text-align:center; padding:15px;'>해당 일자의 데이터가 없습니다.</td></tr>"
    else:
        current_product = None
        for i, row in enumerate(df_target.to_dict('records')):
            item = row.get("품목", "")
            prod_name = row.get("제품명", "")
            
            # 절단 규격 분류
            if "컷레터스" in prod_name:
                cut_spec = "절단 6X6cm"
            elif "양상추" in item or "프릴" in item or "로메인" in item or "카이" in item:
                cut_spec = "절단 4X4cm"
            elif "양배추" in item or "적채" in item or "당근" in item:
                cut_spec = "슬라이스 4mm"
            else:
                cut_spec = "절단 4X4cm" # 기본값

            display_code = row.get("통합_생성코드", "-")

            # 다중 제품 rowspan 동적 그룹핑
            if prod_name != current_product:
                current_product = prod_name
                prod_row_count = len(df_target[df_target['제품명'] == prod_name])
                
                # ★ 데이터프레임의 줄바꿈(\n)을 HTML 태그(<br>)로 변환 ★
                html_prod_name = str(prod_name).replace('\n', '<br>')
                
                prod_cell = f'<td rowspan="{prod_row_count}" style="line-height: 1.4;"><b>{html_prod_name}</b></td>'
                check_cell = f'<td rowspan="{prod_row_count}" style="font-size:8pt; text-align:left; line-height:1.5;">부패,변질,잔잎 등의 비가식부위 제거<br>원재료 상태 및 이물제거 확인</td>'
            else:
                prod_cell = ""
                check_cell = ""

            if is_completed:
                input_after_td = f'<td style="text-align:right;">{row.get("준비양(후)", 0)} kg</td>'
            else:
                input_after_td = '<td style="text-align:right;">kg</td>'

            rows_html += f"""
            <tr>
                {prod_cell}
                <td>{item}</td>
                <td>{display_code}</td>
                <td>적 / 부</td>
                {check_cell}
                <td>{cut_spec}</td>
                <td style="text-align:right;">{row.get("실투입량(전)", 0)} kg</td>
                {input_after_td}
                <td></td>
            </tr>
            """
    
    for _ in range(3):
        rows_html += "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <title>전처리 공정일지</title>
    <style>
        body {{ font-family: '맑은 고딕', sans-serif; padding: 20px; color: #000; background-color: #fff; margin: 0 auto; max-width: 1200px; }}
        @media print {{
            body {{ padding: 0; margin: 0; max-width: none; }}
            @page {{ size: A4 landscape; margin: 15mm; }}
        }}
        .header-table {{ width: 100%; margin-bottom: 10px; }}
        .header-table td {{ vertical-align: top; }}
        .title {{ font-size: 26pt; font-weight: bold; letter-spacing: 2px; }}
        .approval-table {{ border-collapse: collapse; text-align: center; font-size: 10pt; float: right; margin-bottom: 5px; }}
        .approval-table th, .approval-table td {{ border: 1px solid #000; padding: 4px 8px; }}
        .approval-table th {{ width: 50px; font-weight: normal; }}
        .approval-table td {{ height: 40px; width: 60px; }}
        .main-table {{ width: 100%; border-collapse: collapse; font-size: 9pt; border: 1px solid #000; text-align: center; }}
        .main-table th, .main-table td {{ border: 1px solid #000; padding: 6px; }}
        .bg-gray {{ background-color: #f2f2f2; font-weight: bold; }}
    </style>
    </head>
    <body>
        <table class="header-table">
            <tr>
                <td class="title">전처리 공정일지</td>
                <td>
                    <table class="approval-table">
                        <tr><th rowspan="2">결<br>재</th><th>작 성</th><th>검 토</th><th>승 인</th></tr>
                        <tr><td></td><td></td><td></td></tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="main-table">
            <tr class="bg-gray">
                <td colspan="2">작성일자</td>
                <td colspan="2" style="background-color: #fff;">{yyyy} 년 &nbsp;&nbsp;&nbsp; {mm} 월 &nbsp;&nbsp;&nbsp; {dd} 일</td>
                <td colspan="2" style="background-color: #fff;">&nbsp;&nbsp; : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ~ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : &nbsp;&nbsp;</td>
                <td>작성자</td>
                <td colspan="2" style="background-color: #fff;"></td>
            </tr>
            <tr>
                <td colspan="2" rowspan="2" class="bg-gray">원료코드<br>부여방법</td>
                <td colspan="2" rowspan="2">원료 구분+입고일자 - 공급업체 구분</td>
                <td colspan="5" style="text-align: left; font-size: 8pt; padding: 4px 8px;">
                    카이피라(C), 프릴아이스(F), 버터헤드(B), 레드오크(O), 로메인(R), 케일(K), 적근대(S)<br>
                    치커리(E), 적채(A), 양배추(D), 양상추(X), 당근(G)
                </td>
            </tr>
            <tr>
                <td colspan="5" style="text-align: left; font-size: 8pt; padding: 4px 8px; border-top:1px dashed #999;">
                    공급업체 구분: 구름 (G01) , 에상스팜(H02), 승승장구(S01), 한스(K01)
                </td>
            </tr>
            <tr class="bg-gray">
                <th rowspan="2">제품명</th>
                <th rowspan="2">원료명</th>
                <th colspan="7">작업원물 정보</th>
            </tr>
            <tr class="bg-gray">
                <th>원료코드</th>
                <th>상태</th>
                <th>점검 항목</th>
                <th>절단규격</th>
                <th>생산량(전)</th>
                <th>투입량(후)</th>
                <th>비고</th>
            </tr>
            {rows_html}
        </table>
    </body>
    </html>
    """
    return html

# ---------------------------------------------------------
# 데이터 추출 및 처리 핵심 로직
# ---------------------------------------------------------
def extract_yymmdd_and_mmdd(val):
    if pd.isna(val):
        today = datetime.now()
        return today.strftime("%y"), today.strftime("%y%m%d"), today.strftime("%m%d")
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%y"), val.strftime("%y%m%d"), val.strftime("%m%d")
    val_str = str(val).strip()
    m_dot_digit = re.search(r"(\d{2})\.(\d{2})(\d{2})", val_str)
    if m_dot_digit:
        yy, mm, dd = m_dot_digit.group(1), m_dot_digit.group(2), m_dot_digit.group(3)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"
    m_full = re.search(r"(\d{2,4})[-./](\d{1,2})[-./](\d{1,2})", val_str)
    if m_full:
        p1 = m_full.group(1)
        yy = p1[2:] if len(p1) == 4 else p1.zfill(2)
        mm = m_full.group(2).zfill(2)
        dd = m_full.group(3).zfill(2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"
    m_digits = re.search(r"\b(\d{6})\b", val_str)
    if m_digits:
        d = m_digits.group(1)
        return d[:2], d, d[2:]
    current_yy = datetime.now().strftime("%y")
    m_short = re.search(r"(\d{1,2})\s*[-./월]\s*(\d{1,2})", val_str)
    if m_short:
        mm = m_short.group(1).zfill(2)
        dd = m_short.group(2).zfill(2)
        return current_yy, f"{current_yy}{mm}{dd}", f"{mm}{dd}"
    try:
        dt = pd.to_datetime(val_str, errors="coerce", yearfirst=True)
        if pd.notna(dt):
            return dt.strftime("%y"), dt.strftime("%y%m%d"), dt.strftime("%m%d")
    except: pass
    today = datetime.now()
    return current_yy, today.strftime("%y%m%d"), today.strftime("%m%d")

def find_stock_sheet_name(xls):
    for sheet_name in xls.sheet_names:
        if sheet_name.strip().lower() == "sheet1": return sheet_name
    keywords = ["양상추 매입", "양배추 매입", "적채 매입"]
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=10)
            for idx, row in df.iterrows():
                row_str = " ".join(row.dropna().astype(str).tolist())
                if all(k in row_str for k in keywords): return sheet_name
        except: continue
    for sheet_name in xls.sheet_names:
        try:
            df_col_a = pd.read_excel(xls, sheet_name=sheet_name, header=None, usecols=[0])
            col_a = df_col_a[0].dropna()
            if len(col_a) < 5: continue
            valid_dates = pd.to_datetime(col_a, errors='coerce').dropna()
            if len(valid_dates) >= 5 and valid_dates.is_monotonic_increasing: return sheet_name
        except: continue
    return xls.sheet_names[0]

def parse_csv_purchase(uploaded_file):
    if uploaded_file is None: return pd.DataFrame()
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        try:
            df_dict = pd.read_excel(uploaded_file, sheet_name=None)
            return pd.concat(df_dict.values(), ignore_index=True)
        except: return pd.DataFrame()
    for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except: continue
    return pd.DataFrame()

# [핵심 수정] 거래처 추출 및 차단벽이 적용된 블록 단위 파싱 로직
def get_production_records(xls, exclude_sheets):
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets and not s.startswith("재고") and not re.match(r"^\d{2}\.\d{2}$", s.strip())]
    records = []
    
    for sheet in daily_sheets:
        df_s = pd.read_excel(xls, sheet_name=sheet, header=None)
        
        for idx, row in df_s.iterrows():
            row_vals = [str(x).strip() for x in row.tolist()]
            if '발주' in row_vals and '수량' in row_vals:
                # 1. 기준 날짜 및 거래처명 찾기 (B열 = 인덱스 1)
                date_val = row_vals[0]
                _, prod_yymmdd, _ = extract_yymmdd_and_mmdd(date_val)
                std_date = f"20{prod_yymmdd[:2]}-{prod_yymmdd[2:4]}-{prod_yymmdd[4:]}"
                
                vendor_name = row_vals[1] if len(row_vals) > 1 else ""
                if vendor_name == 'nan': vendor_name = ""
                
                # 2. 원료 기입열(Column) 파악
                ingredients = []
                for c_idx in range(4, len(row_vals)):
                    val = row_vals[c_idx]
                    if val != 'nan' and val and val not in ['롯트', '총중량', '담당자'] and not val.replace('.', '').isdigit():
                        ingredients.append({'name': val, 'col': c_idx})
                
                # 3. 제품명 추출 (기준행 바로 아랫줄)
                if idx + 1 >= len(df_s): continue
                product_name = str(df_s.iloc[idx + 1, 0]).strip()
                if product_name == 'nan' or not product_name: continue
                
                # 4. 준비량 및 실투입량 행 찾기 (차단벽 로직 적용)
                prep_row, act_row = None, None
                for scan_idx in range(idx + 2, min(idx + 20, len(df_s))):
                    scan_vals_raw = [str(x) for x in df_s.iloc[scan_idx].tolist()]
                    scan_vals = [x.replace(" ", "") for x in scan_vals_raw]
                    
                    if '발주' in scan_vals and '수량' in scan_vals:
                        break
                    
                    if any('전처리준비' in x for x in scan_vals) and prep_row is None: 
                        prep_row = scan_idx
                    if any('실투입' in x for x in scan_vals) and act_row is None: 
                        act_row = scan_idx
                
                # 5. 각 원료별 데이터 추출
                for ing in ingredients:
                    c_idx = ing['col']
                    ing_name = ing['name']
                    
                    lot_str = str(df_s.iloc[idx + 1, c_idx + 1]).strip() if c_idx + 1 < len(df_s.columns) else ""
                    if lot_str == 'nan' or lot_str.lower() == 'none': lot_str = ""
                    
                    prep_val, act_val = 0.0, 0.0
                    if prep_row is not None:
                        try: prep_val = float(df_s.iloc[prep_row, c_idx])
                        except: pass
                    if act_row is not None:
                        try: act_val = float(df_s.iloc[act_row, c_idx])
                        except: pass
                        
                    if prep_val > 0 or act_val > 0:
                        records.append({
                            'std_date': std_date,
                            'prod_yymmdd': prod_yymmdd,
                            'vendor': vendor_name,       # ★ 거래처명 추가
                            'product': product_name,
                            'ingredient': ing_name,
                            'lot': lot_str,
                            'has_lot': bool(lot_str),
                            'prep_g': prep_val,
                            'act_g': act_val
                        })
    return sorted(records, key=lambda x: x['std_date'])

def process_ledger_and_fifo(xls, stock_sheet, raw_records):
    df_raw = pd.read_excel(xls, sheet_name=stock_sheet, header=None)
    
    def parse_cell(val):
        if pd.isna(val): return 0.0
        try:
            c = re.sub(r"[^\d.]", "", str(val))
            return float(c) if c else 0.0
        except: return 0.0

    grouped_by_date = {}
    date_order = []

    # 1. 수불부 원본에서 입고 및 재고 읽어오기
    for idx, row in df_raw.iterrows():
        if len(row) == 0 or pd.isna(row.iloc[0]): continue
        a_val = str(row.iloc[0]).strip()
        if not a_val or "날짜" in a_val or "입고" in a_val or "매입" in a_val: continue

        _, yymmdd, _ = extract_yymmdd_and_mmdd(a_val)
        dt = f"20{yymmdd[:2]}-{yymmdd[2:4]}-{yymmdd[4:]}"
        
        y_in, y_stk = (parse_cell(row.iloc[1]) if len(row)>1 else 0.0), (parse_cell(row.iloc[2]) if len(row)>2 else 0.0)
        b_in, b_stk = (parse_cell(row.iloc[4]) if len(row)>4 else 0.0), (parse_cell(row.iloc[5]) if len(row)>5 else 0.0)
        j_in, j_stk = (parse_cell(row.iloc[7]) if len(row)>7 else 0.0), (parse_cell(row.iloc[8]) if len(row)>8 else 0.0)

        if dt not in grouped_by_date:
            grouped_by_date[dt] = {
                "date": dt, "yymmdd": yymmdd, "y_in": 0.0, "y_use": 0.0, "y_stk": 0.0,
                "b_in": 0.0, "b_use": 0.0, "b_stk": 0.0, "j_in": 0.0, "j_use": 0.0, "j_stk": 0.0
            }
            date_order.append(dt)

        grouped_by_date[dt]["y_in"] += y_in
        grouped_by_date[dt]["b_in"] += b_in
        grouped_by_date[dt]["j_in"] += j_in
        if y_stk > 0: grouped_by_date[dt]["y_stk"] = y_stk
        if b_stk > 0: grouped_by_date[dt]["b_stk"] = b_stk
        if j_stk > 0: grouped_by_date[dt]["j_stk"] = j_stk

    # 2. 파싱된 생산 내역에서 실투입량 누적 반영 (Y, B, J 한정)
    for rec in raw_records:
        dt = rec['std_date']
        ing = rec['ingredient']
        act_g = rec['act_g']
        if dt not in grouped_by_date:
            grouped_by_date[dt] = {
                "date": dt, "yymmdd": dt.replace("-", "")[2:], "y_in": 0.0, "y_use": 0.0, "y_stk": 0.0,
                "b_in": 0.0, "b_use": 0.0, "b_stk": 0.0, "j_in": 0.0, "j_use": 0.0, "j_stk": 0.0
            }
            if dt not in date_order: date_order.append(dt)
            
        if "양상추" in ing: grouped_by_date[dt]["y_use"] += act_g
        elif "양배추" in ing: grouped_by_date[dt]["b_use"] += act_g
        elif "적채" in ing: grouped_by_date[dt]["j_use"] += act_g

    date_order.sort()

    # 3. FIFO 선입선출 로직 적용
    _, db_lots = get_latest_closed_lots()
    item_queues = {
        "양상추": copy.deepcopy(db_lots["양상추"]),
        "양배추": copy.deepcopy(db_lots["양배추"]),
        "적채": copy.deepcopy(db_lots["적채"])
    }

    data_rows = []
    fifo_map = {}

    for dt in date_order:
        g = grouped_by_date[dt]
        yymmdd = g["yymmdd"]
        mmdd_str = f"{yymmdd[2:4]}/{yymmdd[4:]}"

        if g["y_in"] > 0: item_queues["양상추"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["y_in"], "vendor": "K01"})
        if g["b_in"] > 0: item_queues["양배추"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["b_in"], "vendor": "K01"})
        if g["j_in"] > 0: item_queues["적채"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["j_in"], "vendor": "K01"})

        def trace_usage(item_key, item_prefix, use_amount):
            if use_amount <= 0: return "-", "-"
            rem = use_amount
            used_sources, used_codes = [], []
            for lot in item_queues[item_key]:
                if rem <= 0: break
                avail_val = lot.get("avail", 0.0)
                if avail_val <= 0: continue
                u = min(avail_val, rem)
                lot["avail"] = avail_val - u
                rem -= u
                vendor = lot.get("vendor", "K01")
                used_sources.append(f"{lot['label']}({round(u/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{lot['yymmdd']}-{vendor}")
            if rem > 0:
                used_sources.append(f"부족({round(rem/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{yymmdd}-K01[부족]")
            return " / ".join(used_sources), " / ".join(list(dict.fromkeys(used_codes)))

        y_src, y_code = trace_usage("양상추", "X", g["y_use"])
        b_src, b_code = trace_usage("양배추", "D", g["b_use"])
        j_src, j_code = trace_usage("적채", "A", g["j_use"])

        fifo_map[(dt, "양상추")] = {"code": y_code, "source": y_src}
        fifo_map[(dt, "양배추")] = {"code": b_code, "source": b_src}
        fifo_map[(dt, "적채")] = {"code": j_code, "source": j_src}

        note = "휴무" if (sum([g["y_in"], g["y_use"], g["b_in"], g["b_use"], g["j_in"], g["j_use"]]) == 0) else ""
        
        data_rows.append({
            "date": dt, "yymmdd": yymmdd,
            "yangsangchu_in": round(g["y_in"] / 1000.0, 1), "yangsangchu_use": round(g["y_use"] / 1000.0, 1), 
            "yangsangchu_source": y_src, "yangsangchu_code": y_code, "yangsangchu_stock": round(g["y_stk"] / 1000.0, 1),
            "yangbaechu_in": round(g["b_in"] / 1000.0, 1), "yangbaechu_use": round(g["b_use"] / 1000.0, 1), 
            "yangbaechu_source": b_src, "yangbaechu_code": b_code, "yangbaechu_stock": round(g["b_stk"] / 1000.0, 1),
            "jeokchae_in": round(g["j_in"] / 1000.0, 1), "jeokchae_use": round(g["j_use"] / 1000.0, 1), 
            "jeokchae_source": j_src, "jeokchae_code": j_code, "jeokchae_stock": round(g["j_stk"] / 1000.0, 1),
            "note": note
        })

    # 4. 엑셀 수불부 파일 생성
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "원재료 수불부"
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells("A1:Q1")
    ws["A1"] = "원재료 수불부 (Raw Material Inventory Ledger - kg 단위)"
    ws["A1"].font = Font(name="맑은 고딕", size=16, bold=True, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    header_fill_main = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    regular_font = Font(name="맑은 고딕", size=10)
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    headers_main = [("A2:A3", "일자"), ("B2:F2", "양상추(kg)"), ("G2:K2", "양배추(kg)"), ("L2:P2", "적채(kg)"), ("Q2:Q3", "비고")]
    for cell_range, title in headers_main:
        ws.merge_cells(cell_range)
        ws[cell_range.split(":")[0]] = title

    for col_idx in range(1, 18):
        c = ws.cell(row=2, column=col_idx)
        c.fill, c.font, c.alignment = header_fill_main, Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center")

    sub_headers = ["기초재고", "입고량", "투입량", "소진 매입일자(상세)", "기말재고"]
    for start_col in [2, 7, 12]:
        for idx, sub in enumerate(sub_headers):
            c = ws.cell(row=3, column=start_col + idx, value=sub)
            c.fill, c.font, c.alignment = header_fill_sub, Font(name="맑은 고딕", size=10, bold=True, color="1F4E79"), Alignment(horizontal="center", vertical="center")

    start_row = 4
    for r_idx, item in enumerate(data_rows, start=start_row):
        ws.cell(row=r_idx, column=1, value=item["date"]).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_idx, column=17, value=item.get("note", "")).alignment = Alignment(horizontal="left", vertical="center")

        items_map = [
            (2, item.get("yangsangchu_in", 0), item.get("yangsangchu_use", 0), item.get("yangsangchu_source", "-"), item.get("yangsangchu_stock", 0)),
            (7, item.get("yangbaechu_in", 0), item.get("yangbaechu_use", 0), item.get("yangbaechu_source", "-"), item.get("yangbaechu_stock", 0)),
            (12, item.get("jeokchae_in", 0), item.get("jeokchae_use", 0), item.get("jeokchae_source", "-"), item.get("jeokchae_stock", 0))
        ]

        for start_col, in_val, use_val, source_str, stock_val in items_map:
            end_col_let = get_column_letter(start_col + 4)
            if r_idx == start_row:
                ws.cell(row=r_idx, column=start_col, value=round(stock_val + use_val - in_val, 1))
            else:
                ws.cell(row=r_idx, column=start_col, value=f"={end_col_let}{r_idx-1}")

            ws.cell(row=r_idx, column=start_col + 1, value=in_val)
            ws.cell(row=r_idx, column=start_col + 2, value=use_val)
            ws.cell(row=r_idx, column=start_col + 3, value=source_str).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r_idx, column=start_col + 4, value=stock_val)

        for c_idx in range(1, 18):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font, cell.border = regular_font, thin_border
            if c_idx in [2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16]:
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")

    buffer = io.BytesIO()
    wb.save(buffer)

    # 잔여 롯트 이월 계산
    last_row_data = data_rows[-1] if data_rows else {}
    target_stocks = {
        "양상추": last_row_data.get("yangsangchu_stock", 0.0) * 1000.0,
        "양배추": last_row_data.get("yangbaechu_stock", 0.0) * 1000.0,
        "적채": last_row_data.get("jeokchae_stock", 0.0) * 1000.0
    }
    final_remaining_lots = {"양상추": [], "양배추": [], "적채": []}
    for item_k in ["양상추", "양배추", "적채"]:
        target_g = target_stocks[item_k]
        valid_q = [q for q in item_queues[item_k] if q.get("avail", 0.0) > 0]
        rem_target = target_g
        for q in reversed(valid_q):
            if rem_target <= 0: break
            assign_g = min(q["avail"], rem_target)
            final_remaining_lots[item_k].insert(0, {
                "yymmdd": q["yymmdd"], "avail_g": assign_g, "avail": assign_g, "vendor": q.get("vendor", "K01")
            })
            rem_target -= assign_g

    return buffer.getvalue(), data_rows, fifo_map, final_remaining_lots

def build_final_reports(raw_records, fifo_map, df_purchase):
    prefix_map = {"양상추": "X", "양배추": "D", "적채": "A", "프릴": "F", "로메인": "R", "케일": "K", "적근대": "S", "치커리": "E", "당근": "G", "카이": "C", "버터헤드": "B", "레드오크": "O"}
    vendor_code_map = {"한스": "K01", "승승장구": "S01", "에상스팜": "H02", "넥스토팜": "N01"}

    purchase_list = []
    if not df_purchase.empty:
        date_col = next((c for c in df_purchase.columns if "기간" in str(c) or "일자" in str(c) or "날짜" in str(c)), None)
        item_col = next((c for c in df_purchase.columns if "품명" in str(c) or "품목" in str(c) or "원재료" in str(c)), None)
        vendor_col = next((c for c in df_purchase.columns if "거래처" in str(c) or "매입처" in str(c) or "공급처" in str(c)), None)
        
        if date_col and item_col:
            for _, r in df_purchase.iterrows():
                p_dt_raw, p_item_raw = r[date_col], str(r[item_col]).strip()
                p_ven_raw = str(r[vendor_col]).strip() if vendor_col and pd.notna(r[vendor_col]) else ""
                if pd.isna(p_dt_raw) or not p_item_raw or p_ven_raw == "기타": continue
                
                _, p_yymmdd, _ = extract_yymmdd_and_mmdd(p_dt_raw)
                std_dt = f"20{p_yymmdd[:2]}-{p_yymmdd[2:4]}-{p_yymmdd[4:]}"
                
                matched_item = p_item_raw
                for key in prefix_map.keys():
                    if key in p_item_raw: 
                        matched_item = key; break
                
                pref = prefix_map.get(matched_item, "M")
                v_code = vendor_code_map.get(p_ven_raw, "K01")
                purchase_list.append({"std_date": std_dt, "yymmdd": p_yymmdd, "item": matched_item, "code": f"{pref}{p_yymmdd}-{v_code}", "vendor": p_ven_raw})
    
    final_rows = []
    summary_rows = []
    summary_tracker = set() # To prevent duplicate rows in code summary
    
    for rec in raw_records:
        std_date = rec['std_date']
        vendor = rec.get('vendor', '')
        p_name_raw = rec['product']
        
        # ★ 엑셀 데이터상 줄바꿈(\n) 형태로 거래처명과 제품명 합치기 ★
        combined_product_name = f"{vendor}\n{p_name_raw}" if vendor else p_name_raw
        
        ing_raw = rec['ingredient']
        lot = rec['lot']
        act_g = rec['act_g']
        prep_g = rec['prep_g']
        
        ing_lookup = ing_raw
        for key in prefix_map.keys():
            if key in ing_raw:
                ing_lookup = key; break
        
        pref = prefix_map.get(ing_lookup, "M")
        
        # Determine Codes
        code_subul, code_purchase, src_detail = "-", "-", "-"
        
        if ing_lookup in ["양상추", "양배추", "적채"]:
            f_data = fifo_map.get((std_date, ing_lookup), {})
            code_subul = f_data.get("code", "-")
            src_detail = f_data.get("source", "-")
        
        # Sub-ingredients (or getting purchase code for reference)
        valid_p = [p for p in purchase_list if p["item"] == ing_lookup and p["std_date"] <= std_date]
        if valid_p:
            valid_p.sort(key=lambda x: x["std_date"], reverse=True)
            code_purchase = valid_p[0]["code"]
            if ing_lookup not in ["양상추", "양배추", "적채"]:
                src_detail = f"{valid_p[0]['yymmdd'][2:4]}/{valid_p[0]['yymmdd'][4:]} 매입 참조"

        # Final unified code generator
        unified_code = code_subul if ing_lookup in ["양상추", "양배추", "적채"] else code_purchase
        if unified_code == "-": unified_code = f"{pref}{std_date.replace('-','')[2:]}-K01(임시)"
        
        # User input code
        code_worklog = "없음"
        if rec['has_lot']:
            _, yymmdd_worklog, _ = extract_yymmdd_and_mmdd(lot)
            code_worklog = f"{pref}{yymmdd_worklog}-K01" if yymmdd_worklog else "없음"
            
        # Match Status
        if not rec['has_lot']:
            is_match = "자동부여(작업일지 미기재)"
        else:
            if ing_lookup in ["양상추", "양배추", "적채"]:
                is_match = "일치" if (code_worklog != "없음" and code_subul != "-" and code_worklog in code_subul) else "불일치"
            else:
                is_match = "일치" if (code_worklog != "없음" and code_purchase != "-" and code_worklog.split("-")[0] in code_purchase) else "불일치(최근매입 상이)"

        final_rows.append({
            "생산일자": std_date,
            "제품명": combined_product_name, # 거래처+제품명 
            "품목": ing_lookup,
            "작업일지 기준 코드": code_worklog,
            "수불부 기준 코드": code_subul,
            "매입자료 기준 코드": code_purchase,
            "통합_생성코드": unified_code,
            "일치여부": is_match,
            "실투입량(전)": round(act_g / 1000.0, 1),
            "준비양(후)": round(prep_g / 1000.0, 1)
        })
        
        # Populate Tab 3 Summary
        summary_key = (std_date, ing_lookup, unified_code, src_detail)
        if summary_key not in summary_tracker and act_g > 0:
            summary_tracker.add(summary_key)
            summary_rows.append({
                "날짜": std_date, "원재료": ing_lookup, "생성코드": unified_code,
                "투입량(kg)": round(act_g / 1000.0, 1), "준비량(kg)": round(prep_g / 1000.0, 1),
                "소진매입일(상세)": src_detail, "비고": ""
            })

    df_result = pd.DataFrame(final_rows)
    df_summary = pd.DataFrame(summary_rows)
    return df_result, df_summary


st.title("🥗 샐러드 생산일보 원료코드 대조 및 수불부 연동 System")

with st.sidebar:
    st.header("🗄️ 마스터 수불 DB 관리")
    last_ym, db_lots = get_latest_closed_lots()
    if last_ym:
        st.success(f"📌 현재 로드된 마감월: **{last_ym}**")
        with st.expander("이월 잔여 롯트 상세"):
            for k, v in db_lots.items():
                if v:
                    st.caption(f"**{k}**")
                    for l in v: st.write(f"- 입고일: {l['yymmdd']} | 잔여: {round(l['avail_g']/1000.0, 1)}kg")
    else:
        st.info("💡 등록된 전월 마감 DB가 없습니다.\n당월 업로드 엑셀의 입고분부터 순서대로 소진됩니다.")
    st.markdown("---")
    if st.button("🗑️ 과거 이월 DB 완전 초기화", type="secondary"):
        clear_db()
        st.success("✅ DB가 초기화되었습니다!")
        st.rerun()

col1, col2 = st.columns(2)
with col1: excel_file = st.file_uploader("1. 샐러드 생산일지 (Excel) 업로드", type=["xlsx", "xls"])
with col2: csv_file = st.file_uploader("2. 원료 매입 입고자료 / 수불부 (CSV/Excel) 업로드", type=["csv", "xlsx", "xls"])

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🔍 원료코드 대조표 (자동 연동)",
    "📊 표준 원재료 수불부 정리",
    "🏷️ 원료코드 생성 및 일자별 정리",
    "🔒 당월 수불 마감 확정",
    "📋 전처리 공정일지 출력",
    "📋 전처리 공정일지 (완성본)"
])

if excel_file is not None:
    try:
        # 1. Load Data
        df_purchase = parse_csv_purchase(csv_file)
        xls = pd.ExcelFile(excel_file)
        stock_sheet_name = find_stock_sheet_name(xls)
        exclude_sheets = ["테스트", "원가", stock_sheet_name, "Sheet1", "sheet1"]
        
        # 2. Block-based Parsing
        raw_records = get_production_records(xls, exclude_sheets)
        
        if not raw_records:
            st.error("엑셀 파일에서 생산 내역을 찾을 수 없습니다. 양식을 확인해주세요.")
            st.stop()
            
        # 3. Ledger & Reports
        excel_bytes, standard_rows, fifo_map, final_lots = process_ledger_and_fifo(xls, stock_sheet_name, raw_records)
        df_result, df_code_summary = build_final_reports(raw_records, fifo_map, df_purchase)

        # --- Tab 1: 원료코드 대조표 ---
        with tab1:
            st.success("✅ 블록형 생산일보 파싱 및 원료코드 매칭이 성공적으로 완료되었습니다!")
            st.markdown("##### 📌 (단위: kg)")
            display_cols = ["생산일자", "제품명", "품목", "작업일지 기준 코드", "수불부 기준 코드", "매입자료 기준 코드", "통합_생성코드", "일치여부", "실투입량(전)", "준비양(후)"]
            df_display = df_result[display_cols]
            
            col_d, col_p = st.columns([3, 1])
            with col_d:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_display.to_excel(writer, index=False, sheet_name="원료코드_대조표")
                st.download_button("📥 엑셀 다운로드 (.xlsx)", data=buffer.getvalue(), file_name="원료코드_대조표.xlsx", use_container_width=True)
            with col_p:
                html_content = render_printable_document("샐러드 생산일보 원료코드 대조표 (단위: kg)", df_display)
                render_direct_print_button(html_content, "🖨️ 바로 출력 / PDF 저장")
            st.dataframe(df_display, use_container_width=True)

        # --- Tab 2: 표준 원재료 수불부 ---
        with tab2:
            st.subheader("📋 표준 원재료 수불부 (메인 3종 FIFO 적용)")
            df_preview = pd.DataFrame(standard_rows)
            ordered_cols = ["date", "yangsangchu_in", "yangsangchu_use", "yangsangchu_source", "yangsangchu_stock",
                            "yangbaechu_in", "yangbaechu_use", "yangbaechu_source", "yangbaechu_stock",
                            "jeokchae_in", "jeokchae_use", "jeokchae_source", "jeokchae_stock", "note"]
            df_preview = df_preview[[c for c in ordered_cols if c in df_preview.columns]]
            df_preview = df_preview.rename(columns={"date": "일자", "yangsangchu_in": "양상추_입고", "yangsangchu_use": "양상추_투입", "yangsangchu_source":"양상추_소진출처", "yangsangchu_stock":"양상추_재고",
                                                    "yangbaechu_in": "양배추_입고", "yangbaechu_use": "양배추_투입", "yangbaechu_source":"양배추_소진출처", "yangbaechu_stock":"양배추_재고",
                                                    "jeokchae_in": "적채_입고", "jeokchae_use": "적채_투입", "jeokchae_source":"적채_소진출처", "jeokchae_stock":"적채_재고", "note": "비고"})
            
            col_d2, col_p2 = st.columns([3, 1])
            with col_d2:
                st.download_button("📥 엑셀 다운로드 (.xlsx)", data=excel_bytes, file_name="표준_원재료_수불부(kg).xlsx", use_container_width=True)
            with col_p2:
                html_content = render_printable_document("표준 원재료 수불부", df_preview)
                render_direct_print_button(html_content, "🖨️ 바로 출력 / PDF 저장")
            st.dataframe(df_preview, use_container_width=True)
            
            # 구글 시트 연동
            st.markdown("---")
            st.subheader("☁️ 수불부 구글 시트로 출고량 연동")
            col_sync_s, col_sync_e = st.columns(2)
            with col_sync_s: sync_start_date = st.date_input("전송 시작일", value=datetime.today().replace(day=1))
            with col_sync_e: sync_end_date = st.date_input("전송 종료일", value=datetime.today())
            
            if st.button("🚀 선택한 기간 데이터 구글 시트 업데이트", type="primary", use_container_width=True):
                with st.spinner("구글 시트 연동 중..."):
                    try:
                        filtered_rows = [r for r in standard_rows if sync_start_date <= datetime.strptime(r["date"], "%Y-%m-%d").date() <= sync_end_date]
                        doc = init_gspread()
                        sheet = doc.worksheet("시트1")
                        all_values = sheet.get_all_values()
                        rows_to_delete = []
                        for i, r_data in enumerate(all_values):
                            if i == 0: continue
                            try:
                                sheet_date = datetime.strptime(r_data[0].strip(), "%Y-%m-%d").date()
                                if sync_start_date <= sheet_date <= sync_end_date and r_data[2].strip() == "시스템자동":
                                    rows_to_delete.append(i + 1)
                            except: pass
                        for row_idx in sorted(rows_to_delete, reverse=True): sheet.delete_row(row_idx)
                        
                        rows_to_upload = []
                        for r in filtered_rows:
                            d_str = str(r["date"])
                            if r.get("yangsangchu_use", 0) > 0: rows_to_upload.append([d_str, "출고", "시스템자동", "양상추", r["yangsangchu_use"], 0, 0, r.get("yangsangchu_code", "")])
                            if r.get("yangbaechu_use", 0) > 0: rows_to_upload.append([d_str, "출고", "시스템자동", "양배추", r["yangbaechu_use"], 0, 0, r.get("yangbaechu_code", "")])
                            if r.get("jeokchae_use", 0) > 0: rows_to_upload.append([d_str, "출고", "시스템자동", "적채", r["jeokchae_use"], 0, 0, r.get("jeokchae_code", "")])
                        if rows_to_upload:
                            sheet.append_rows(rows_to_upload)
                            st.success(f"✅ 기존 데이터 {len(rows_to_delete)}건 삭제 및 최신 데이터 {len(rows_to_upload)}건 업데이트 완료!")
                        else:
                            st.success("✅ 기존 데이터 삭제 완료. 새로 전송할 데이터는 없습니다.")
                    except Exception as e: st.error(f"오류 발생: {e}")

        # --- Tab 3: 원료코드 생성 및 일자별 정리 ---
        with tab3:
            st.subheader("🏷️ 전체 원료(서브 포함) 일자별 원료코드 컨닝페이퍼")
            st.info("💡 메인 3종은 선입선출, 나머지 서브원료(로메인, 케일 등)는 가장 최근 매입일 기준으로 자동 생성됩니다.")
            if not df_code_summary.empty:
                item_priority = {"양상추": 1, "양배추": 2, "적채": 3, "프릴": 4}
                df_code_summary["_sort_key"] = df_code_summary["원재료"].map(lambda x: item_priority.get(x, 99))
                df_code_summary = df_code_summary.sort_values(by=["날짜", "_sort_key"]).drop(columns=["_sort_key"])

                col_d3, col_p3 = st.columns([3, 1])
                with col_d3:
                    code_buffer = io.BytesIO()
                    with pd.ExcelWriter(code_buffer, engine="openpyxl") as writer:
                        df_code_summary.to_excel(writer, index=False, sheet_name="일자별_원료코드_생성")
                    st.download_button("📥 엑셀 다운로드 (.xlsx)", data=code_buffer.getvalue(), file_name="일자별_전체_원료코드_목록.xlsx", use_container_width=True)
                with col_p3:
                    html_content = render_printable_document("일자별 전체 생성 원료코드 목록", df_code_summary)
                    render_direct_print_button(html_content, "🖨️ 바로 출력 / PDF 저장")
                st.dataframe(df_code_summary, use_container_width=True)

        # --- Tab 4: 당월 수불 마감 ---
        with tab4:
            st.subheader("🔒 당월 수불 마감 및 전월 DB 확정")
            last_date_str = standard_rows[-1]["date"] if standard_rows else datetime.now().strftime("%Y-%m-%d")
            default_ym = last_date_str[:7]
            closed_ym_input = st.text_input("마감 확정 월 (YYYY-MM)", value=default_ym)

            st.write("▼ **마감 후 다음 달로 이월될 메인 3종 잔여 롯트 미리보기**")
            for item_k, l_list in final_lots.items():
                tot_kg = sum(l['avail_g'] for l in l_list) / 1000.0
                st.write(f"• **{item_k}**: 총 {round(tot_kg, 1)}kg (롯트 {len(l_list)}개 잔여)")
                for l in l_list: st.caption(f"   - 입고일: {l['yymmdd']} | 잔여량: {round(l['avail_g']/1000.0, 1)}kg")

            if st.button("🔒 당월 수불 마감 및 DB 확정 저장", type="primary"):
                save_closed_lots(closed_ym_input, final_lots)
                st.balloons()
                st.success(f"✅ **{closed_ym_input}** 수불 마감이 성공적으로 저장되었습니다!")

        # --- Tab 5/6: 전처리 공정일지 HTML ---
        for tab, is_comp in zip([tab5, tab6], [False, True]):
            with tab:
                st.subheader(f"📋 전처리 공정일지 자동 생성 ({'완성본' if is_comp else '빈칸용'})")
                if not df_result.empty:
                    unique_dates = sorted(df_result["생산일자"].unique())
                    col_dt, _ = st.columns([1, 2])
                    with col_dt:
                        sel_date = st.selectbox(f"📅 출력할 생산일자 ({'완성본' if is_comp else '빈칸용'})", unique_dates, key=f"date_sel_{is_comp}")
                    
                    df_target = df_result[df_result["생산일자"] == sel_date].copy()
                    item_priority = {"양상추": 1, "양배추": 2, "적채": 3, "프릴": 4}
                    df_target["_item_sort"] = df_target["품목"].map(lambda x: item_priority.get(x, 99))
                    df_target = df_target.sort_values(by=["제품명", "_item_sort"])
                    
                    html_log = render_process_log_html(sel_date, df_target, is_completed=is_comp)
                    
                    col_pb, _ = st.columns([1, 3])
                    with col_pb: render_direct_print_button(html_log, f"🖨️ {'완성본 ' if is_comp else ''}바로 출력 / PDF")
                    components.html(html_log, height=800, scrolling=True)

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())
else:
    for tab_idx, tab_obj in enumerate([tab1, tab2, tab3, tab4, tab5, tab6]):
        with tab_obj: st.info("👆 상단에서 1. 샐러드 생산일지와 2. 원료 매입 수불부를 업로드해 주세요.")
