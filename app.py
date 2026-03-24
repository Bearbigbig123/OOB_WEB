import streamlit as st
import requests
import time
import pandas as pd
import base64
import os
import uuid
import io
import xlsxwriter
import hashlib
from PIL import Image
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# --- 配置 ---
API_BASE_URL = "http://localhost:8000"

# --- 帳號設定（sha256 雜湊；可自行新增帳號）---
# 產生方式：import hashlib; hashlib.sha256(b"your_password").hexdigest()
_ACCOUNTS = {
    "admin": hashlib.sha256(b"admin123").hexdigest(),
    "osat":  hashlib.sha256(b"osat2026").hexdigest(),
}
POLLING_INTERVAL = 2.0  # 輪詢頻率(秒)

st.set_page_config(page_title="OSAT SPC System", layout="wide")

# --- CSS：優化進度條顏色、隱藏預設 Header、調整版面 ---
st.markdown("""
    <style>
    /* 1. 這裡保留你的進度條漸層色 */
    .stProgress > div > div > div > div { 
        background-image: linear-gradient(to right, #344CB7, #577BC1); 
    }
    
    /* 2. 💡 這裡移除了隱藏 Header 的代碼，小人圖示會回來 */
    
    /* 3. 調整頁面間距，既然 Header 回來了，頂部 padding 縮小一點才不會留白太多 */
    .block-container { padding-top: 3rem; }
    
    /* 4. 讓所有 st.subheader (h3) 小一號 */
    h3 {
        font-size: 1.3rem !important;
        font-weight: 600 !important;
        color: #262730;
        margin-bottom: 0.5rem !important;
        padding-top: 0.5rem !important;
    }
    
    /* 5. 增加 st.popover (Settings 彈出視窗) 的高度 */
    div[data-testid="stPopoverBody"] {
        max-height: 95vh !important;
        height: auto !important;
        overflow-y: auto !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 狀態管理 ---
if 'task_id' not in st.session_state: st.session_state.task_id = None
if 'last_task_id' not in st.session_state: st.session_state.last_task_id = None
if 'status' not in st.session_state: st.session_state.status = "idle"
if 'progress' not in st.session_state: st.session_state.progress = 0
if 'results' not in st.session_state: st.session_state.results = None
if 'current_mode' not in st.session_state: st.session_state.current_mode = None
if 'full_excel_data' not in st.session_state: st.session_state.full_excel_data = None
if 'trigger_analysis' not in st.session_state: st.session_state.trigger_analysis = False
if 'pending_endpoint' not in st.session_state: st.session_state.pending_endpoint = None
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'login_user' not in st.session_state: st.session_state.login_user = ""

# ==========================================
# LOGIN GATE
# ==========================================
if not st.session_state.logged_in:
    st.markdown("""
            <style>
            /* 1. 隱藏原本突兀的白色卡片，改用透明毛玻璃質感 */
            .stApp {
                background-color: #0E1117;
            }

            /* 2. 登入容器：徹底拿掉白色背景 */
            .login-box {
                max-width: 420px;
                margin: 12vh auto 0 auto;
                padding: 2rem;
                text-align: center;
                background: rgba(255, 255, 255, 0.03); /* 極微弱的白色透明 */
                border: 1px solid rgba(255, 255, 255, 0.1); /* 纖細邊框 */
                border-radius: 15px 15px 0 0;
                backdrop-filter: blur(10px); /* 毛玻璃效果 */
            }

            .login-title {
                color: #FFFFFF;
                font-size: 2rem;
                font-weight: 800;
                letter-spacing: 1px;
                text-shadow: 0 0 10px rgba(52, 76, 183, 0.5); /* 標題微光 */
                margin-bottom: 0.5rem;
            }

            .login-sub {
                color: #888;
                font-size: 0.9rem;
                margin-bottom: 0;
            }

            /* 3. 處理下方 Form：無縫接軌上方的透明框 */
            div[data-testid="stForm"] {
                max-width: 420px;
                margin: -1px auto 0 auto; /* 讓上下框完美接合 */
                background: rgba(255, 255, 255, 0.05) !important; 
                border: 1px solid rgba(255, 255, 255, 0.1) !important;
                border-top: none !important; /* 拿掉銜接處的線 */
                border-radius: 0 0 15px 15px !important;
                padding: 2rem !important;
                box-shadow: 0 20px 40px rgba(0,0,0,0.4);
            }

            /* 4. 讓 Input 框更帥：深黑底+淡藍邊 */
            input {
                background-color: #0E1117 !important;
                color: white !important;
                border: 1px solid #344CB7 !important;
            }

            /* 5. 登入按鈕：改用藍色漸層光，更符合 QC 戰情室的科技感 */
            button[kind="primaryFormSubmit"] {
                width: 100% !important;
                background: linear-gradient(90deg, #344CB7, #577BC1) !important;
                border: none !important;
                color: white !important;
                font-weight: 700 !important;
                height: 3rem !important;
                transition: 0.3s !important;
                box-shadow: 0 4px 15px rgba(52, 76, 183, 0.3) !important;
            }
            
            button[kind="primaryFormSubmit"]:hover {
                box-shadow: 0 0 20px rgba(52, 76, 183, 0.6) !important;
                transform: translateY(-2px);
            }
            </style>

            <div class='login-box'>
                <div class='login-title'>OSAT SPC SYSTEM</div>
                <div class='login-sub'>PRECISION QUALITY CONTROL</div>
            </div>
        """, unsafe_allow_html=True)

    _, login_col, _ = st.columns([1, 1.2, 1])
    with login_col:
        username = st.text_input("帳號", placeholder="Username", label_visibility="collapsed", key="_lu")
        password = st.text_input("密碼", placeholder="Password", type="password", label_visibility="collapsed", key="_lp")
        if st.button("登入", type="primary", use_container_width=True):
            hashed = hashlib.sha256(password.encode()).hexdigest()
            if username in _ACCOUNTS and _ACCOUNTS[username] == hashed:
                st.session_state.logged_in = True
                st.session_state.login_user = username
                st.rerun()
            else:
                st.error("帳號或密碼錯誤")
    st.stop()
if 'pending_payload' not in st.session_state: st.session_state.pending_payload = None
if 'pending_mode' not in st.session_state: st.session_state.pending_mode = None
if 'auto_split_info' not in st.session_state: st.session_state.auto_split_info = None

# ==========================================
# 0. 輔助函數：產生包含圖片的完整 Excel
# ==========================================
def generate_full_excel_with_images(data_list, mode):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Analysis_Results')

    df = pd.DataFrame(data_list)

    # 處理 Metrics 展開 (用於 CPK 模式)
    if 'metrics' in df.columns:
        metrics_df = pd.json_normalize(df['metrics'])
        df = pd.concat([df.drop(columns=['metrics']), metrics_df], axis=1)

    # 定義 Excel 樣式
    header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11, 'bold': True, 'bg_color': '#344CB7', 'font_color': 'white'})
    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10})

    # 決定要排除的欄位 (圖片路徑不需要顯示為文字)
    exclude_cols = ['chart_path', 'weekly_chart_path', 'by_tool_color_path', 'by_tool_group_path', 'qq_plot_path', 'chart_image', 'spc_chart_path', 'boxplot_chart_path', 'timeline_chart_path']
    data_cols = [c for c in df.columns if c not in exclude_cols]

    is_oob_mode = mode == "OOB/SPC"
    is_cpk_mode = mode == "CPK Dashboard"
    is_tm_mode  = mode == "Tool Matching"

    # OOB: 5 張圖表欄位
    img_fields = [
        ('chart_path', 'Total SPC'),
        ('weekly_chart_path', 'Weekly SPC'),
        ('by_tool_color_path', 'By Tool (Color)'),
        ('by_tool_group_path', 'By Tool (Group)'),
        ('qq_plot_path', 'Q-Q Plot')
    ]
    # CPK: 1 張圖表欄位 (base64)
    cpk_img_fields = [('chart_image', 'SPC Chart')]
    # TM: 3 張圖表欄位 (file path)
    tm_img_fields = [
        ('spc_chart_path', 'SPC Chart'),
        ('boxplot_chart_path', 'Boxplot'),
        ('timeline_chart_path', 'Timeline'),
    ]

    if is_oob_mode:
        num_img_cols = len(img_fields)
    elif is_cpk_mode:
        num_img_cols = len(cpk_img_fields)
    elif is_tm_mode:
        num_img_cols = len(tm_img_fields)
    else:
        num_img_cols = 0
    start_col = num_img_cols

    # 寫入圖片的表頭
    if is_oob_mode:
        for i, (key, title) in enumerate(img_fields):
            worksheet.write(0, i, title, header_format)
    elif is_cpk_mode:
        for i, (key, title) in enumerate(cpk_img_fields):
            worksheet.write(0, i, title, header_format)
    elif is_tm_mode:
        for i, (key, title) in enumerate(tm_img_fields):
            worksheet.write(0, i, title, header_format)

    # 寫入文字數據的表頭
    col_widths = {}
    for i, col_name in enumerate(data_cols):
        worksheet.write(0, start_col + i, col_name.replace('_', ' ').title(), header_format)
        col_widths[start_col + i] = max(len(str(col_name)), 12)

    scale_factor = 0.28 if is_oob_mode else (0.40 if is_tm_mode else 0.55)
    max_image_height = 0
    max_image_width = 0

    # CPK 模式：建立暫存目錄用於 base64 圖片解碼
    import tempfile
    _cpk_tmp_dir = tempfile.mkdtemp() if is_cpk_mode else None
    _tm_scale_factor = 0.40

    # 迭代資料列
    for row_idx, row in df.iterrows():
        excel_row = row_idx + 1

        # 插入圖片 (OOB: 5 張路徑圖片；CPK: 1 張 base64 圖片)
        if is_oob_mode:
            for col_offset, (key, _) in enumerate(img_fields):
                img_path = row.get(key)
                if pd.notna(img_path) and isinstance(img_path, str) and os.path.exists(img_path):
                    try:
                        options = {'x_scale': scale_factor, 'y_scale': scale_factor, 'x_offset': 5, 'y_offset': 5, 'object_position': 1}
                        worksheet.insert_image(excel_row, col_offset, img_path, options)
                        with Image.open(img_path) as img:
                            scaled_h = img.height * scale_factor
                            scaled_w = img.width * scale_factor
                            max_image_height = max(max_image_height, scaled_h)
                            max_image_width = max(max_image_width, scaled_w)
                    except Exception:
                        worksheet.write(excel_row, col_offset, "Image Error", cell_format)
        elif is_cpk_mode:
            for col_offset, (key, _) in enumerate(cpk_img_fields):
                img_b64 = row.get(key)
                if pd.notna(img_b64) and isinstance(img_b64, str) and img_b64:
                    try:
                        img_bytes = base64.b64decode(img_b64)
                        tmp_path = os.path.join(_cpk_tmp_dir, f"cpk_{row_idx}_{col_offset}.png")
                        with open(tmp_path, 'wb') as f:
                            f.write(img_bytes)
                        options = {'x_scale': scale_factor, 'y_scale': scale_factor, 'x_offset': 5, 'y_offset': 5, 'object_position': 1}
                        worksheet.insert_image(excel_row, col_offset, tmp_path, options)
                        with Image.open(tmp_path) as img:
                            scaled_h = img.height * scale_factor
                            scaled_w = img.width * scale_factor
                            max_image_height = max(max_image_height, scaled_h)
                            max_image_width = max(max_image_width, scaled_w)
                    except Exception:
                        worksheet.write(excel_row, col_offset, "Image Error", cell_format)
        elif is_tm_mode:
            for col_offset, (key, _) in enumerate(tm_img_fields):
                img_path = row.get(key)
                if pd.notna(img_path) and isinstance(img_path, str) and os.path.exists(img_path):
                    try:
                        options = {'x_scale': _tm_scale_factor, 'y_scale': _tm_scale_factor, 'x_offset': 5, 'y_offset': 5, 'object_position': 1}
                        worksheet.insert_image(excel_row, col_offset, img_path, options)
                        with Image.open(img_path) as img:
                            scaled_h = img.height * _tm_scale_factor
                            scaled_w = img.width * _tm_scale_factor
                            max_image_height = max(max_image_height, scaled_h)
                            max_image_width = max(max_image_width, scaled_w)
                    except Exception:
                        worksheet.write(excel_row, col_offset, "Image Error", cell_format)

        # 寫入文字數據
        for i, col_name in enumerate(data_cols):
            val = row.get(col_name)
            if pd.isna(val):
                val = ""
            elif isinstance(val, (list, dict)):
                val = str(val)
            worksheet.write(excel_row, start_col + i, val, cell_format)
            col_widths[start_col + i] = max(col_widths.get(start_col + i, 12), len(str(val)) + 2)

    # 設定寬度與高度
    if (is_oob_mode or is_cpk_mode or is_tm_mode) and max_image_height > 0:
        if is_oob_mode:
            col_w_cap, row_h_cap = 55, 150
        elif is_tm_mode:
            col_w_cap, row_h_cap = 75, 200
        else:
            col_w_cap, row_h_cap = 80, 200
        for i in range(num_img_cols):
            worksheet.set_column(i, i, min((max_image_width / 7) + 2, col_w_cap))
        for row_idx in range(1, len(df) + 1):
            worksheet.set_row(row_idx, min((max_image_height * 0.75) + 10, row_h_cap))
    else:
        for row_idx in range(1, len(df) + 1):
            worksheet.set_row(row_idx, 20)

    for col_idx, width in col_widths.items():
        worksheet.set_column(col_idx, col_idx, min(width, 40))

    # 凍結首列與圖片欄位
    if is_oob_mode or is_cpk_mode or is_tm_mode:
        worksheet.freeze_panes(1, num_img_cols)
    else:
        worksheet.freeze_panes(1, 0)

    workbook.close()
    if _cpk_tmp_dir and os.path.exists(_cpk_tmp_dir):
        import shutil
        shutil.rmtree(_cpk_tmp_dir, ignore_errors=True)
    return output.getvalue()


# ==========================================
# 1. 頂部導覽橫條 (Header)
# ==========================================
col1, col2, col3, col4 = st.columns([1.5, 3.8, 0.8, 0.8], gap="medium")

with col1:
    with st.popover("⚙️ Settings & Run", use_container_width=True):
        st.markdown("##### 分析設定")
        mode = st.radio("選擇功能", ["OOB/SPC", "Tool Matching", "CPK Dashboard"])
        base_date = st.date_input("分析基準日", value=datetime.now())
        
       
        # --- 檔案上傳區塊 (水平排列) ---
        st.markdown("###### 📁 上傳自訂檔案 (若不傳則使用預設)")
        up_col_left, up_col_right = st.columns(2)
        
        with up_col_left:
            excel_file = st.file_uploader("1️⃣ All Charts (Excel)", type=["xlsx"], help="上傳定義控制線的 Excel 檔")
            
        with up_col_right:
            csv_files = st.file_uploader("2️⃣ Raw Data (CSV, 多選)", type=["csv"], accept_multiple_files=True, help="上傳產線原始資料 CSV 檔")
            
        st.divider()
        
        if st.button("🚀 Start Analysis", type="primary", use_container_width=True):
            st.session_state.current_mode = mode
            st.session_state.results = None
            
            # 處理使用者上傳的檔案
            custom_excel_path = None
            custom_raw_dir = None
            
            if excel_file or csv_files:
                upload_session_id = str(uuid.uuid4())
                base_upload_dir = os.path.abspath(os.path.join("temp_uploads", "ui_uploads", upload_session_id))
                os.makedirs(base_upload_dir, exist_ok=True)
                
                if excel_file:
                    custom_excel_path = os.path.join(base_upload_dir, excel_file.name)
                    with open(custom_excel_path, "wb") as f:
                        f.write(excel_file.getbuffer())
                
                if csv_files:
                    custom_raw_dir = os.path.join(base_upload_dir, "raw_charts")
                    os.makedirs(custom_raw_dir, exist_ok=True)
                    for csv in csv_files:
                        csv_path = os.path.join(custom_raw_dir, csv.name)
                        with open(csv_path, "wb") as f:
                            f.write(csv.getbuffer())

            # 自動偵測並拆分待拆格式（一次上傳整份 CSV + AllChart 即可）
            auto_split_id      = None
            auto_split_raw_dir = None
            st.session_state.auto_split_info = None
            
            VENDOR_COLS = {"Part ID", "Item Name", "Report Time", "Lot Mean", "Vendor Site"}
            TEST_H_COLS = {"Part ID", "FT Test End Time", "Test Site"}
            TYPE2_COLS  = {"GroupName", "ChartName", "point_time", "point_val"}

            if csv_files and len(csv_files) == 1 and custom_raw_dir:
                first_saved = os.path.join(custom_raw_dir, csv_files[0].name)
                try:
                    # 1. 先讀取標準表頭來判斷 Vendor, Test_H, Type2
                    peek = pd.read_csv(first_saved, nrows=0)
                    detected_cols = set(peek.columns)
                    detected_split_mode = None
                    
                    if VENDOR_COLS.issubset(detected_cols):
                        detected_split_mode = "Vendor_Vertical"
                    elif TEST_H_COLS.issubset(detected_cols):
                        detected_split_mode = "Test_Horizontal"
                    elif TYPE2_COLS.issubset(detected_cols):
                        detected_split_mode = "Type2_Vertical"
                    else:
                        # 2. 如果都不是，嘗試無表頭讀取前 3 列，偵測是否為 Type3_Horizontal (多層表頭)
                        peek_no_header = pd.read_csv(first_saved, nrows=3, header=None)
                        flat_vals = peek_no_header.iloc[0:2].fillna("").astype(str).values.flatten().tolist()
                        if any("GroupName" in val for val in flat_vals) and any("ChartName" in val for val in flat_vals):
                            detected_split_mode = "Type3_Horizontal"

                    # 若有成功配對到任何一種模式，就呼叫後端進行自動拆分
                    if detected_split_mode:
                        split_resp = requests.post(
                            f"{API_BASE_URL}/split",
                            json={"mode": detected_split_mode, "input_files": [os.path.abspath(first_saved)]},
                            timeout=120,
                        )
                        if split_resp.status_code == 200:
                            split_result = split_resp.json()
                            auto_split_id      = split_result.get("split_id")
                            auto_split_raw_dir = split_result.get("raw_data_directory")
                            st.session_state.auto_split_info = (
                                f"🔀 自動偵測到 **{detected_split_mode}** 格式，"
                                f"已完成拆分（{split_result.get('processed', 0)} 個 chart）"
                            )
                        else:
                            st.session_state.auto_split_info = (
                                f"⚠️ 自動拆分失敗（HTTP {split_resp.status_code}），將以原始目錄繼續分析"
                            )
                except Exception:
                    pass  # 偵測失敗就視為一般上傳，不影響原有流程

            # 組裝 API Payload
            payload = {}
            if mode == "OOB/SPC":
                endpoint = "/process"
                if custom_excel_path: payload["filepath"] = custom_excel_path
                if auto_split_id:
                    payload["split_id"] = auto_split_id
                elif custom_raw_dir:
                    payload["raw_data_directory"] = custom_raw_dir
            elif mode == "Tool Matching":
                endpoint = "/tool-matching"
                payload = {"base_date": base_date.strftime("%Y-%m-%d"), "filter_mode": "specified_date"}
                if custom_excel_path: payload["chart_excel_path"] = custom_excel_path
                if auto_split_raw_dir:
                    payload["raw_data_directory"] = auto_split_raw_dir
                elif custom_raw_dir:
                    payload["raw_data_directory"] = custom_raw_dir
            else: # CPK Dashboard
                endpoint = "/spc-cpk"
                payload = {"end_date": base_date.strftime("%Y-%m-%d")}
                if custom_excel_path: payload["chart_excel_path"] = custom_excel_path
                if auto_split_raw_dir:
                    payload["raw_data_directory"] = auto_split_raw_dir
                elif custom_raw_dir:
                    payload["raw_data_directory"] = custom_raw_dir
            
            # 先儲存參數並關閉彈窗，重繪後再送出 API
            st.session_state.pending_mode = mode
            st.session_state.pending_endpoint = endpoint
            st.session_state.pending_payload = payload
            st.session_state.trigger_analysis = True
            st.rerun()

with col2:
    if st.session_state.status == "processing":
        st.progress(st.session_state.progress / 100.0)
        st.caption(f"⏳ {st.session_state.current_mode} 執行中... ({st.session_state.progress}%)")
    elif st.session_state.status == "completed":
        st.progress(1.0)
        st.caption("✅ 任務已完成，請點擊左下方表格任意一列檢視圖表")
    elif st.session_state.status == "failed":
        st.progress(0)
        err_msg = ""
        if st.session_state.task_id and st.session_state.task_id in st.session_state.get('_last_errors', {}):
            err_msg = st.session_state['_last_errors'][st.session_state.task_id]
        st.caption(f"❌ 任務失敗{f': {err_msg}' if err_msg else ''}")
    if st.session_state.auto_split_info:
        st.caption(st.session_state.auto_split_info)

with col3:
    if st.button("🔄 Reset", use_container_width=True):
        st.session_state.task_id = None
        st.session_state.last_task_id = None
        st.session_state.status = "idle"
        st.session_state.results = None
        st.session_state.progress = 0
        st.session_state.full_excel_data = None
        st.rerun()

with col4:
    st.caption(f"👤 {st.session_state.login_user}")

# ==========================================
# 1.5 延遲執行分析 (彈窗關閉後才送出 API)
# ==========================================
if st.session_state.trigger_analysis:
    st.session_state.trigger_analysis = False
    st.session_state.current_mode = st.session_state.pending_mode
    st.session_state.results = None
    try:
        resp = requests.post(f"{API_BASE_URL}{st.session_state.pending_endpoint}", json=st.session_state.pending_payload)
        if resp.status_code == 200:
            st.session_state.task_id = resp.json().get("task_id")
            st.session_state.status = "processing"
            st.session_state.progress = 0
            st.rerun()
        else:
            st.error(f"啟動失敗: {resp.text}")
    except Exception as e:
        st.error(f"API 連線失敗: {e}")

# ==========================================
# 2. 輪詢進度 (Polling Logic)
# ==========================================
if st.session_state.status == "processing" and st.session_state.task_id:
    try:
        resp = requests.get(f"{API_BASE_URL}/process/status/{st.session_state.task_id}").json()
        st.session_state.progress = resp.get("progress", 0)
        status = resp.get("status")
        if status == "completed":
            st.session_state.results = resp.get("result")
            st.session_state.status = "completed"
            st.rerun()
        elif status == "failed":
            st.session_state.status = "failed"
            err = resp.get('error', '')
            if err:
                if '_last_errors' not in st.session_state:
                    st.session_state['_last_errors'] = {}
                st.session_state['_last_errors'][st.session_state.task_id] = err
                st.error(f"後端出錯: {err}")
            st.rerun()
        else:
            time.sleep(POLLING_INTERVAL)
            st.rerun()
    except Exception as e:
        st.error(f"輪詢失敗: {e}")

st.divider()

# ==========================================
# 3. 報表與圖表顯示區 (Master-Detail with AgGrid)
# ==========================================
if st.session_state.results:
    res = st.session_state.results
    data_list = []
    if st.session_state.current_mode in ["OOB/SPC", "Tool Matching"]:
        data_list = res.get("results", [])
    elif st.session_state.current_mode == "CPK Dashboard":
        data_list = res.get("charts", [])
        
    if data_list:
        df_summary = pd.DataFrame(data_list)
        
        if 'metrics' in df_summary.columns:
            metrics_df = pd.json_normalize(df_summary['metrics'])
            df_processed = pd.concat([df_summary.drop(columns=['metrics']), metrics_df], axis=1)
        else:
            df_processed = df_summary.copy()

        c_top_left, c_top_right = st.columns([1.6, 2.4], gap="small")
        
        with c_top_left:
            # === 💡 在標題旁邊加入「下載完整 Excel 報告」按鈕 ===
            title_col, btn_col = st.columns([1, 1])
            with title_col:
                st.markdown("##### Summary Table")
            
            with btn_col:
                # 如果是新的任務，則在背景產生一份 Excel 暫存在 Session 裡
                if st.session_state.last_task_id != st.session_state.task_id:
                    st.session_state.full_excel_data = generate_full_excel_with_images(data_list, st.session_state.current_mode)
                    st.session_state.last_task_id = st.session_state.task_id
                
                st.download_button(
                    label="📥 下載完整 Excel 報告 (含圖片)",
                    data=st.session_state.full_excel_data,
                    file_name=f"SPC_Full_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # --- 💡 資料清洗與自動置頂邏輯 ---
            keep_list = ['gname', 'cname', 'group', 'group_name', 'chart_name', 'Characteristics', 'characteristics', 'WE_Rule', 'OOB_Rule', 'abnormal_type', 'cpk', 'cpk_l1', 'cpk_l2', 'r1', 'r2', 'k_value', 'mean_index', 'sigma_index', 'data_cnt']
            existing_cols = [c for c in keep_list if c in df_processed.columns]
            display_df = df_processed[existing_cols].copy()
            
            normal_keywords = ["", "NA", "N/A", "NONE", "NAN", "NORMAL", "PASS", "NO_HIGHLIGHT", "FALSE", "<NA>", "0", "OK", "-"]
            
            for rule_col in ['WE_Rule', 'OOB_Rule']:
                if rule_col in display_df.columns:
                    clean_str = display_df[rule_col].astype(str).str.strip().str.upper()
                    display_df.loc[clean_str.isin(normal_keywords), rule_col] = ""
            
            has_we = (display_df['WE_Rule'] != "") if 'WE_Rule' in display_df.columns else pd.Series(False, index=display_df.index)
            has_oob = (display_df['OOB_Rule'] != "") if 'OOB_Rule' in display_df.columns else pd.Series(False, index=display_df.index)
            has_tm = (display_df['abnormal_type'] != "") if 'abnormal_type' in display_df.columns else pd.Series(False, index=display_df.index)
            display_df['has_issue'] = has_we | has_oob | has_tm
            
            _sort_candidates = ['has_issue', 'group_name', 'gname', 'chart_name', 'cname']
            _sort_cols = [c for c in _sort_candidates if c in display_df.columns]
            _ascending = [False if c == 'has_issue' else True for c in _sort_cols]
            display_df = display_df.sort_values(by=_sort_cols, ascending=_ascending)          
            display_df = display_df.drop(columns=['has_issue'])
            
            if 'WE_Rule' in display_df.columns:
                display_df.loc[display_df['WE_Rule'] == "", 'WE_Rule'] = "-"
            if 'OOB_Rule' in display_df.columns:
                display_df.loc[display_df['OOB_Rule'] == "", 'OOB_Rule'] = "-"
            
            _na_vals = ["N/A", "None", "nan", "NaN", "none", "<NA>"]
            for _col in display_df.columns:
                if _col not in ['WE_Rule', 'OOB_Rule']: 
                    if display_df[_col].dtype == object:
                        display_df[_col] = display_df[_col].fillna("-").replace(_na_vals, "-")
                    else:
                        display_df[_col] = display_df[_col].where(display_df[_col].notna(), other="-")

            # --- 建立 AgGrid ---
            gb = GridOptionsBuilder.from_dataframe(display_df)
            gb.configure_selection(selection_mode="single", use_checkbox=False)
            gb.configure_default_column(resizable=True)
            
            col_settings = {
                "gname": {"header_name": "Group", "width": 90},
                "cname": {"header_name": "Chart", "width": 140},
                "group": {"header_name": "Group", "width": 90},
                "group_name": {"header_name": "Group", "width": 90},
                "chart_name": {"header_name": "Chart", "width": 140},
                "Characteristics": {"header_name": "Char.", "width": 75},
                "characteristics": {"header_name": "Char.", "width": 75},
                "WE_Rule": {"header_name": "WE Rule", "width": 200},
                "OOB_Rule": {"header_name": "OOB Rule", "width": 200},
                "abnormal_type": {"header_name": "Abnormal", "width": 90},
                "cpk": {"header_name": "Cpk", "width": 70},
                "cpk_l1": {"header_name": "Cpk L1", "width": 70},
                "cpk_l2": {"header_name": "Cpk L2", "width": 70},
                "r1": {"header_name": "R1", "width": 65},
                "r2": {"header_name": "R2", "width": 65},
                "k_value": {"header_name": "K Value", "width": 80},
                "mean_index": {"header_name": "Mean Idx", "width": 90},
                "sigma_index": {"header_name": "Sigma Idx", "width": 90},
                "data_cnt": {"header_name": "N", "width": 70}
            }
            
            for col in display_df.columns:
                if col in col_settings:
                    gb.configure_column(col, **col_settings[col])
            
            gridOptions = gb.build()
            
            grid_response = AgGrid(
                display_df,
                gridOptions=gridOptions,
                update_mode=GridUpdateMode.SELECTION_CHANGED,
                fit_columns_on_grid_load=False,
                height=350,
                theme='streamlit' 
            )
            
            selected_rows = grid_response.get('selected_rows')

        with c_top_right:
            
            item = None
            if selected_rows is not None and len(selected_rows) > 0:
                if isinstance(selected_rows, pd.DataFrame):
                    sel_dict = selected_rows.iloc[0].to_dict()
                else:
                    sel_dict = selected_rows[0]
                    
                # 同時支援 OOB (group_name/chart_name) 與 Tool Matching (gname/cname) 欄位
                g_name = sel_dict.get('group_name') or sel_dict.get('gname') or sel_dict.get('group')
                c_name = sel_dict.get('chart_name') or sel_dict.get('cname')
                item = next(
                    (x for x in data_list if
                     (str(x.get('group_name', '')) == str(g_name) or str(x.get('gname', '')) == str(g_name) or str(x.get('group', '')) == str(g_name)) and
                     (str(x.get('chart_name', '')) == str(c_name) or str(x.get('cname', '')) == str(c_name))),
                    None
                )
            
            if item:
                if st.session_state.current_mode == "OOB/SPC":
                    if item.get('chart_path'):
                        st.markdown(f"**All Data SPC - {item.get('chart_name')}**")
                        st.image(item['chart_path'], use_container_width=True)
                    else:
                        st.info("無主圖資料")
                elif st.session_state.current_mode == "CPK Dashboard" and item.get('chart_image'):
                    img_bytes = base64.b64decode(item['chart_image'])
                    st.image(img_bytes, caption=f"{item.get('group_name')} - {item.get('chart_name')}", use_container_width=True)
                elif st.session_state.current_mode == "Tool Matching":
                    spc_path = item.get('spc_chart_path')
                    if spc_path and os.path.exists(spc_path):
                        st.image(spc_path, use_container_width=True)
                    else:
                        st.info("此項目無 SPC 圖表。")
            else:
                st.markdown(
                    """
                    <div style='background-color: #f9f9f9; border: 2px dashed #ccc; border-radius: 10px; padding: 50px; text-align: center; margin-top: 20px;'>
                        <h2 style='color: #555;'>👈 請點擊左側表格</h2>
                        <p style='color: #777; font-size: 1.1em;'>滑鼠點擊表格中的<strong style='color: #344CB7;'>任意一列</strong>，右側將自動載入對應圖表。</p>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )

        # ==========================================
        # 下方區塊：2x2 網格放置剩餘圖表
        # ==========================================
        if item and st.session_state.current_mode == "OOB/SPC":
            
            rest_charts = [
                ("**Weekly SPC**", item.get('weekly_chart_path')),
                ("**By Tool SPC**", item.get('by_tool_color_path')),
                ("**By Tool Group SPC**", item.get('by_tool_group_path')),
                ("**Q-Q Plot**", item.get('qq_plot_path'))
            ]
            valid_rest_charts = [(title, path) for title, path in rest_charts if path]
            
            if valid_rest_charts:
                st.divider()
                
                for i in range(0, len(valid_rest_charts), 2):
                    bottom_cols = st.columns(2)
                    with bottom_cols[0]:
                        st.markdown(valid_rest_charts[i][0])
                        st.image(valid_rest_charts[i][1], use_container_width=True)
                    if i + 1 < len(valid_rest_charts):
                        with bottom_cols[1]:
                            st.markdown(valid_rest_charts[i+1][0])
                            st.image(valid_rest_charts[i+1][1], use_container_width=True)

        elif item and st.session_state.current_mode == "Tool Matching":
            tl_path = item.get('timeline_chart_path')
            box_path = item.get('boxplot_chart_path')
            charts_to_show = []
            if tl_path and os.path.exists(tl_path):
                charts_to_show.append(("**Timeline (All Tools)**", tl_path))
            if box_path and os.path.exists(box_path):
                charts_to_show.append(("**Boxplot**", box_path))
            if charts_to_show:
                st.divider()
                bottom_cols = st.columns(len(charts_to_show))
                for col, (title, path) in zip(bottom_cols, charts_to_show):
                    with col:
                        st.markdown(title)
                        st.image(path, use_container_width=True)
else:
    if st.session_state.status == "idle":
        st.markdown("<h3 style='text-align: center; color: #888; padding-top: 100px;'>點擊左上角 Settings 開始分析</h3>", unsafe_allow_html=True)