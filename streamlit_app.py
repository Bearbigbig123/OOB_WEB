import streamlit as st
import requests
import pandas as pd
import json
from typing import Optional, Dict, Any
import base64
from datetime import datetime
import io
import zipfile
import os
from PIL import Image

# 嘗試載入 AG Grid，如果無法載入則使用標準 dataframe
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
    AGGRID_AVAILABLE = True
except ImportError:
    AGGRID_AVAILABLE = False
    st.warning("⚠️ AG Grid 未安裝，將使用標準表格。安裝指令：pip install streamlit-aggrid")

# 配置頁面
st.set_page_config(
    page_title="OOB & Tool Matching 分析系統", 
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# API 基礎 URL
API_BASE_URL = "http://localhost:8000"

class APIClient:
    """API 客戶端類別，處理與後台的通訊"""
    
    @staticmethod
    def check_health() -> bool:
        """檢查後台 API 是否運行"""
        try:
            response = requests.get(f"{API_BASE_URL}/health", timeout=5)
            return response.status_code == 200
        except:
            return False
    
    @staticmethod
    def process_oob(request_data: dict) -> Optional[dict]:
        """呼叫 OOB 處理 API"""
        try:
            response = requests.post(
                f"{API_BASE_URL}/process",
                json=request_data,
                timeout=300  # 5 分鐘超時
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            st.error(f"API 呼叫失敗: {str(e)}")
            return None
    
    @staticmethod
    def analyze_tool_matching(request_data: dict) -> Optional[dict]:
        """呼叫 Tool Matching 分析 API"""
        try:
            response = requests.post(
                f"{API_BASE_URL}/tool-matching",
                json=request_data,
                timeout=300
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            st.error(f"API 呼叫失敗: {str(e)}")
            return None
    
    @staticmethod
    def analyze_spc_cpk(request_data: dict) -> Optional[dict]:
        """呼叫 SPC CPK Dashboard 分析 API"""
        try:
            response = requests.post(
                f"{API_BASE_URL}/spc-cpk",
                json=request_data,
                timeout=300
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            st.error(f"API 呼叫失敗: {str(e)}")
            return None

def init_session_state():
    """初始化 session state"""
    if 'api_connected' not in st.session_state:
        st.session_state.api_connected = False
    if 'oob_results' not in st.session_state:
        st.session_state.oob_results = None
    if 'tool_matching_results' not in st.session_state:
        st.session_state.tool_matching_results = None
    if 'spc_cpk_results' not in st.session_state:
        st.session_state.spc_cpk_results = None

def check_api_connection():
    """檢查並顯示 API 連線狀態"""
    st.session_state.api_connected = APIClient.check_health()
    
    if st.session_state.api_connected:
        st.sidebar.success("🟢 後台 API 連線正常")
    else:
        st.sidebar.error("🔴 後台 API 連線失敗")
        st.sidebar.info("請確保後台服務正在運行：`uvicorn main:app --host localhost --port 8000`")

def display_chart_images_vertical(chart_result: dict, location: str = "vertical"):
    """上下排列顯示兩張圖片，用於表格旁的互動區域
    
    Args:
        chart_result: 圖表結果數據
        location: 顯示位置標識
    """
    chart_path = chart_result.get('chart_path')
    weekly_chart_path = chart_result.get('weekly_chart_path')
    
    # 設定目標尺寸（橫向長方形：寬度大於高度，提高解析度）
    target_width = 1200  # 像素寬度（提高解析度）
    target_height = 600  # 像素高度（提高解析度）
    
    # SPC 圖表（上方）
    if chart_path and os.path.exists(chart_path):
        try:
            image = Image.open(chart_path)
            # 使用高品質重採樣算法調整大小
            resized_image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)
            
            st.image(resized_image, caption="SPC Chart", width=600)
        except Exception as e:
            st.error(f"無法載入 SPC 圖表: {e}")
    else:
        st.info("SPC 圖表未生成")
    
    # 小間距
    st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
    
    # Weekly SPC 圖表（下方）
    if weekly_chart_path and os.path.exists(weekly_chart_path):
        try:
            image = Image.open(weekly_chart_path)
            # 使用高品質重採樣算法調整大小
            resized_image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)
            
            st.image(resized_image, caption="Weekly SPC Chart", width=600)
        except Exception as e:
            st.error(f"無法載入 Weekly SPC 圖表: {e}")
    else:
        st.info("Weekly SPC 圖表未生成")

def display_chart_images_fullwidth(chart_result: dict, index: int = 0, location: str = "main"):
    """全寬度顯示單個圖表結果的兩張圖片和違規規則詳情
    
    Args:
        chart_result: 圖表結果數據
        index: 圖表索引
        location: 顯示位置標識（main, table_selection, etc.）
    """
    chart_path = chart_result.get('chart_path')
    weekly_chart_path = chart_result.get('weekly_chart_path')
    
    if chart_path or weekly_chart_path:
        # 圖表資訊標題
        chart_name = f"{chart_result.get('group_name', 'Unknown')} - {chart_result.get('chart_name', 'Unknown')}"
        st.markdown(f"### 📊 {chart_name}")
        
        # 使用三欄佈局：左邊兩欄顯示圖表，右邊一欄顯示違規詳情
        col1, col2, col3 = st.columns([2, 2, 1], gap="medium")
        
        with col1:
            if chart_path and os.path.exists(chart_path):
                st.markdown("#### 📈 SPC 圖表")
                try:
                    image = Image.open(chart_path)
                    # 使用高解析度處理，保持原始品質並適度縮放
                    if image.width > 1200 or image.height > 600:
                        # 保持比例縮放到適合尺寸
                        ratio = min(1200/image.width, 600/image.height)
                        new_width = int(image.width * ratio)
                        new_height = int(image.height * ratio)
                        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    st.image(image, caption="SPC Chart", width=700)
                    
                    # 提供單張圖片下載
                    with open(chart_path, "rb") as f:
                        st.download_button(
                            label="📥 下載 SPC 圖表",
                            data=f.read(),
                            file_name=os.path.basename(chart_path),
                            mime="image/png",
                            key=f"download_spc_{location}_{chart_result.get('group_name', 'unknown')}_{chart_result.get('chart_name', 'unknown')}"
                        )
                except Exception as e:
                    st.error(f"無法載入 SPC 圖表: {e}")
            else:
                st.info("SPC 圖表未生成")
        
        with col2:
            if weekly_chart_path and os.path.exists(weekly_chart_path):
                st.markdown("#### 📅 Weekly SPC 圖表")
                try:
                    image = Image.open(weekly_chart_path)
                    # 使用高解析度處理，保持原始品質並適度縮放
                    if image.width > 1200 or image.height > 600:
                        # 保持比例縮放到適合尺寸
                        ratio = min(1200/image.width, 600/image.height)
                        new_width = int(image.width * ratio)
                        new_height = int(image.height * ratio)
                        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    st.image(image, caption="Weekly SPC Chart", width=700)
                    
                    # 提供單張圖片下載
                    with open(weekly_chart_path, "rb") as f:
                        st.download_button(
                            label="📥 下載 Weekly SPC 圖表",
                            data=f.read(),
                            file_name=os.path.basename(weekly_chart_path),
                            mime="image/png",
                            key=f"download_weekly_{location}_{chart_result.get('group_name', 'unknown')}_{chart_result.get('chart_name', 'unknown')}"
                        )
                except Exception as e:
                    st.error(f"無法載入 Weekly SPC 圖表: {e}")
            else:
                st.info("Weekly SPC 圖表未生成")
        
        with col3:
            # 顯示違規規則詳情
            st.markdown("#### 🚨 違規分析")
            
            # 創建規則狀態表格
            rule_data = []
            
            # Western Electric Rules
            we_rule = chart_result.get('WE_Rule', 'N/A')
            if we_rule and we_rule != 'N/A' and we_rule.strip():
                we_rules = we_rule.split(',') if ',' in str(we_rule) else [str(we_rule)]
                for rule in we_rules:
                    rule = rule.strip()
                    if rule:
                        rule_data.append({
                            '類型': 'WE Rule',
                            '規則': rule,
                            '狀態': '❌ 違反'
                        })
            
            # OOB Rules  
            oob_rule = chart_result.get('OOB_Rule', 'N/A')
            if oob_rule and oob_rule != 'N/A' and oob_rule.strip():
                oob_rules = oob_rule.split(',') if ',' in str(oob_rule) else [str(oob_rule)]
                for rule in oob_rules:
                    rule = rule.strip()
                    if rule:
                        rule_data.append({
                            '類型': 'OOB Rule',
                            '規則': rule,
                            '狀態': '❌ 違反'
                        })
            
            if rule_data:
                # 顯示違規規則表格
                rule_df = pd.DataFrame(rule_data).copy()
                st.dataframe(rule_df, width='stretch', hide_index=True)
            else:
                # 沒有違規的情況
                st.success("✅ 無違規規則")
        
        # 添加分隔線
        st.markdown("---")

def display_chart_images(chart_result: dict, index: int = 0):
    """顯示單個圖表結果的兩張圖片"""
    chart_path = chart_result.get('chart_path')
    weekly_chart_path = chart_result.get('weekly_chart_path')
    
    if chart_path or weekly_chart_path:
        col1, col2 = st.columns(2)
        
        with col1:
            if chart_path and os.path.exists(chart_path):
                st.write("**SPC 圖表**")
                try:
                    image = Image.open(chart_path)
                    # 使用高解析度處理，保持原始品質並適度縮放
                    if image.width > 1200 or image.height > 600:
                        # 保持比例縮放到適合尺寸
                        ratio = min(1200/image.width, 600/image.height)
                        new_width = int(image.width * ratio)
                        new_height = int(image.height * ratio)
                        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    st.image(image, caption="SPC Chart", width=700)
                    
                    # 提供單張圖片下載
                    with open(chart_path, "rb") as f:
                        st.download_button(
                            label="📥 下載 SPC 圖表",
                            data=f.read(),
                            file_name=os.path.basename(chart_path),
                            mime="image/png",
                            key=f"download_spc_{index}_{chart_result.get('group_name', 'unknown')}_{chart_result.get('chart_name', 'unknown')}"
                        )
                except Exception as e:
                    st.error(f"無法載入 SPC 圖表: {e}")
            else:
                st.info("SPC 圖表未生成")
        
        with col2:
            if weekly_chart_path and os.path.exists(weekly_chart_path):
                st.write("**Weekly SPC 圖表**")
                try:
                    image = Image.open(weekly_chart_path)
                    # 使用高解析度處理，保持原始品質並適度縮放
                    if image.width > 1200 or image.height > 600:
                        # 保持比例縮放到適合尺寸
                        ratio = min(1200/image.width, 600/image.height)
                        new_width = int(image.width * ratio)
                        new_height = int(image.height * ratio)
                        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    st.image(image, caption="Weekly SPC Chart", width=700)
                    
                    # 提供單張圖片下載
                    with open(weekly_chart_path, "rb") as f:
                        st.download_button(
                            label="📥 下載 Weekly SPC 圖表",
                            data=f.read(),
                            file_name=os.path.basename(weekly_chart_path),
                            mime="image/png",
                            key=f"download_weekly_{index}_{chart_result.get('group_name', 'unknown')}_{chart_result.get('chart_name', 'unknown')}"
                        )
                except Exception as e:
                    st.error(f"無法載入 Weekly SPC 圖表: {e}")
            else:
                st.info("Weekly SPC 圖表未生成")

def render_file_uploader_with_filter(key: str, accept_multiple_files: bool = False, file_types: list = None, title: str = "選擇檔案"):
    """帶篩選功能的檔案上傳元件"""
    if file_types is None:
        file_types = ['csv', 'xlsx', 'xls']
    
    # 檔案上傳
    uploaded_files = st.file_uploader(
        title,
        type=file_types,
        accept_multiple_files=accept_multiple_files,
        key=key
    )
    
    # 如果有多個檔案，顯示篩選功能
    if uploaded_files and accept_multiple_files and len(uploaded_files) > 3:  # 降低閾值讓篩選功能更容易觸發
        st.markdown("**🔍 檔案篩選**")
        # 篩選控制 - 添加下拉選擇
        filter_col1, filter_col2 = st.columns([3, 2])
        with filter_col1:
            search_term = st.text_input("篩選檔案名稱", placeholder="輸入關鍵字...", key=f"{key}_filter")
        with filter_col2:
            # 準備下拉選項
            file_names = [f.name for f in uploaded_files]
            file_names.insert(0, "全部檔案")  # 添加「全部」選項
            
            selected_file = st.selectbox("選擇特定檔案", file_names, key=f"{key}_selector")
        
        # 篩選邏輯
        if selected_file and selected_file != "全部檔案":
            # 使用下拉選擇的結果
            filtered_files = [f for f in uploaded_files if f.name == selected_file]
        elif search_term:
            # 使用文字篩選
            filtered_files = [f for f in uploaded_files if search_term.lower() in f.name.lower()]
        else:
            # 預設顯示所有檔案
            filtered_files = uploaded_files
        
        # 只在有篩選時顯示計數
        if len(filtered_files) < len(uploaded_files):
            st.caption(f"📁 顯示 {len(filtered_files)} / {len(uploaded_files)} 個檔案")
        
        # 顯示篩選後的檔案列表
        for i, file in enumerate(filtered_files):
            col_file, col_info = st.columns([3, 1])
            with col_file:
                st.write(f"✅ {file.name}")
            with col_info:
                # 顯示檔案大小
                file_size = len(file.getvalue()) / 1024  # KB
                if file_size < 1024:
                    st.caption(f"{file_size:.1f} KB")
                else:
                    st.caption(f"{file_size/1024:.1f} MB")
        
        return uploaded_files  # 返回所有檔案，但界面只顯示篩選後的
    elif uploaded_files and accept_multiple_files:
        # 檔案數量少時直接顯示
        st.caption(f" {len(uploaded_files)} 個檔案")
        for file in uploaded_files:
            col_file, col_info = st.columns([3, 1])
            with col_file:
                st.write(f"✅ {file.name}")
            with col_info:
                # 顯示檔案大小
                file_size = len(file.getvalue()) / 1024  # KB
                if file_size < 1024:
                    st.caption(f"{file_size:.1f} KB")
                else:
                    st.caption(f"{file_size/1024:.1f} MB")
    
    return uploaded_files

def render_file_uploader(key: str, accept_multiple_files: bool = False, file_types: list = None):
    """檔案上傳元件（保持向後相容）"""
    return render_file_uploader_with_filter(key, accept_multiple_files, file_types)

def save_uploaded_file(uploaded_file, directory: str) -> str:
    """儲存上傳的檔案到指定目錄"""
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    file_path = os.path.join(directory, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    return file_path

def render_oob_page():
    """OOB 分析分頁"""
    
    if not st.session_state.api_connected:
        st.warning("⚠️ 後台 API 未連線，無法進行分析")
        return
    
    # 頂部控制欄 - 使用彈窗設定
    col_header1, col_header2, col_header3 = st.columns([1, 2, 1])
    
    with col_header1:
        # 檔案設定彈窗按鈕
        st.markdown("<br>", unsafe_allow_html=True)  # 添加一些間距
        with st.popover("📁 檔案設定"):
            # 標題列帶關閉提示
            col_title, col_close = st.columns([4, 1])
            with col_title:
                st.markdown("**📋 檔案上傳設定**")
            with col_close:
                st.markdown("*點擊外部關閉*", help="點擊彈窗外的任何地方即可關閉此設定窗口")
            
            st.divider()
            
            # 使用橫向排版
            col_upload1, col_upload2 = st.columns(2)
            
            with col_upload1:
                st.write("**📊 Chart Information 檔案**")
                chart_info_file = render_file_uploader_with_filter("chart_info", file_types=['xlsx'], title="上傳 Excel 檔案")
                
                # 檔案狀態檢查
                if chart_info_file:
                    st.success(f"✅ {chart_info_file.name}")
                else:
                    st.error("❌ 未上傳 Chart Info 檔案")
            
            with col_upload2:
                st.write("**📁 原始資料檔案 (CSV)**")
                raw_data_files = render_file_uploader_with_filter("raw_data", accept_multiple_files=True, file_types=['csv'], title="上傳多個 CSV 檔案")
                
                # 檔案狀態檢查
                if raw_data_files:
                    # 如果檔案數量少於等於5個，或者沒有篩選功能啟用，顯示完整清單
                    if len(raw_data_files) <= 5:
                        st.success(f"✅ {len(raw_data_files)} 個檔案")
                        for file in raw_data_files:
                            st.write(f"✅ {file.name}")
                    else:
                        st.success(f"✅ 已上傳 {len(raw_data_files)} 個檔案")
                        # 篩選功能已在 render_file_uploader_with_filter 中處理
                else:
                    st.warning("⚠️ 未上傳原始資料檔案")
            
            # 使用預設參數
            save_excel = True
            scale_factor = 0.3
            limit_charts = None
    
    with col_header2:
        # 置中的標題區
        st.markdown("<div style='text-align: center; padding: 20px;'><h3>� OOB/SPC 分析系統</h3></div>", 
                   unsafe_allow_html=True)
    
    with col_header3:
        # 執行按鈕
        st.markdown("<br>", unsafe_allow_html=True)  # 添加一些間距
        if st.button("🚀 開始分析", key="oob_analyze", type="primary"):
            if chart_info_file is None:
                st.error("❌ 請先在設定中上傳 Chart Information 檔案")
                return
            
            # 儲存檔案
            temp_dir = "temp_uploads"
            chart_info_path = save_uploaded_file(chart_info_file, temp_dir)
            
            # 處理原始資料檔案
            raw_data_dir = None
            if raw_data_files:
                raw_data_dir = os.path.join(temp_dir, "raw_charts")
                for file in raw_data_files:
                    save_uploaded_file(file, raw_data_dir)
            
            # 準備 API 請求資料
            request_data = {
                "filepath": chart_info_path,
                "raw_data_directory": raw_data_dir,
                "save_excel": save_excel,
                "scale_factor": scale_factor,
                "limit_charts": limit_charts
            }
            
            # 顯示處理中狀態
            with st.spinner("正在處理分析..."):
                result = APIClient.process_oob(request_data)
                
            if result:
                st.session_state.oob_results = result
                st.success("✅ 分析完成！")
                st.rerun()  # 刷新頁面顯示結果
                
                # 清理暫存檔案
                try:
                    import shutil
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                except:
                    pass
            else:
                st.error("❌ 分析失敗，請檢查檔案格式和後台狀態")
    
    # 分析結果區域 - 占據整個頁面寬度
    st.markdown("---")
    
    if st.session_state.oob_results:
        result = st.session_state.oob_results
        
        # 計算違規統計
        we_violations = 0
        oob_violations = 0
        if result['results']:
            for chart_result in result['results']:
                we_rule = chart_result.get('WE_Rule', 'N/A')
                oob_rule = chart_result.get('OOB_Rule', 'N/A')
                
                if we_rule and we_rule != 'N/A' and str(we_rule).strip():
                    we_violations += 1
                if oob_rule and oob_rule != 'N/A' and str(oob_rule).strip():
                    oob_violations += 1
        
        # 計算無資料統計
        no_data_count = result['summary']['skipped_charts'] if 'skipped_charts' in result['summary'] else 0
        
        # 顯示摘要 - 使用全寬度的指標卡片
        st.subheader("📊 分析摘要")
        col_metrics = st.columns(5)
        with col_metrics[0]:
            st.metric("總圖表數", result['summary']['total_charts'])
        with col_metrics[1]:
            st.metric("已處理", result['summary']['processed_charts'])
        with col_metrics[2]:
            st.metric("無資料", no_data_count)
        with col_metrics[3]:
            st.metric("WE 違規", we_violations)
        with col_metrics[4]:
            st.metric("OOB 違規", oob_violations)
        
        # 下載功能區
        st.subheader("📥 下載選項")
        
        if result['results']:
            # Excel 報告下載
            if result['summary'].get('excel_output') and os.path.exists(result['summary']['excel_output']):
                with open(result['summary']['excel_output'], "rb") as f:
                    excel_data = f.read()
                
                col_download = st.columns([1, 2, 1])
                with col_download[1]:
                    st.download_button(
                        label="📊 下載 Excel 報告",
                        data=excel_data,
                        file_name=f"oob_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.warning("⚠️ 無 Excel 報告可供下載")
        
        # 顯示圖表（如果有）- 全寬度顯示
        chart_results = [r for r in result['results'] if r.get('chart_path') or r.get('weekly_chart_path')]
        if chart_results:
            st.subheader("📊 生成的圖表")
            
            # 添加圖表篩選功能
            if len(chart_results) > 5:
                st.markdown("##### 🔍 圖表篩選")
                filter_col1, filter_col2 = st.columns([3, 2])
                
                with filter_col1:
                    chart_filter = st.text_input("篩選圖表名稱", placeholder="輸入關鍵字篩選圖表...", key="chart_result_filter")
                with filter_col2:
                    # 準備下拉選項
                    chart_names = [f"{r.get('group_name', 'Unknown')} - {r.get('chart_name', 'Unknown')}" 
                                 for r in chart_results]
                    chart_names.insert(0, "全部圖表")  # 添加「全部」選項
                    
                    selected_chart = st.selectbox("選擇特定圖表", chart_names, key="chart_selector")
                
                # 篩選圖表結果
                if selected_chart and selected_chart != "全部圖表":
                    # 使用下拉選擇的結果
                    selected_index = chart_names.index(selected_chart) - 1  # 減1因為第0個是「全部」
                    filtered_chart_results = [chart_results[selected_index]]
                elif chart_filter:
                    # 使用文字篩選
                    filtered_chart_results = [
                        r for r in chart_results 
                        if chart_filter.lower() in r.get('group_name', '').lower() 
                        or chart_filter.lower() in r.get('chart_name', '').lower()
                    ]
                else:
                    # 預設顯示所有圖表
                    filtered_chart_results = chart_results
            else:
                filtered_chart_results = chart_results
            
            # 使用 tabs 來組織多個圖表，讓圖表占據更多空間
            if len(filtered_chart_results) > 1:
                # 為每個圖表組合創建一個 tab
                tab_names = [f"{r.get('group_name', 'Unknown')}_{r.get('chart_name', 'Unknown')}" 
                           for r in filtered_chart_results]
                tabs = st.tabs(tab_names)
                
                for i, (tab, chart_result) in enumerate(zip(tabs, filtered_chart_results)):
                    with tab:
                        display_chart_images_fullwidth(chart_result, i, "main_tabs")
            else:
                # 只有一組圖表時直接顯示
                if filtered_chart_results:
                    display_chart_images_fullwidth(filtered_chart_results[0], 0, "main_single")
        
        # 詳細結果表格
        if result['results']:
            # 使用兩欄佈局：左邊表格標題，右邊互動圖表標題
            title_col1, title_col2 = st.columns([3, 2])  # 3:2 的比例
            
            with title_col1:
                st.subheader("📋 詳細結果表格")
            with title_col2:
                st.subheader("🎯 互動圖表區域")
            
            # 轉換為 DataFrame（使用 .copy() 避免警告）
            df_results = pd.DataFrame(result['results']).copy()
            
            # 使用兩欄佈局：左邊表格，右邊互動圖表
            table_col, chart_col = st.columns([3, 2])  # 3:2 的比例
            
            with table_col:
                # 添加表格篩選功能
                if len(df_results) > 10:
                    st.markdown("##### 🔍 表格篩選")
                    table_filter_col1, table_filter_col2 = st.columns([3, 2])
                    
                    with table_filter_col1:
                        table_filter = st.text_input("篩選結果", placeholder="輸入關鍵字篩選結果...", key="table_result_filter")
                    with table_filter_col2:
                        # 準備下拉選項 - 使用組合名稱
                        unique_combinations = df_results[['group_name', 'chart_name']].drop_duplicates()
                        result_names = [f"{row['group_name']} - {row['chart_name']}" 
                                      for _, row in unique_combinations.iterrows()]
                        result_names.insert(0, "全部結果")  # 添加「全部」選項
                        
                        selected_result = st.selectbox("選擇特定結果", result_names, key="result_selector")
                    
                    # 篩選表格結果
                    if selected_result and selected_result != "全部結果":
                        # 使用下拉選擇的結果
                        group_name, chart_name = selected_result.split(" - ", 1)
                        filtered_df = df_results[
                            (df_results['group_name'] == group_name) & 
                            (df_results['chart_name'] == chart_name)
                        ]
                    elif table_filter:
                        # 使用文字篩選
                        mask = df_results.astype(str).apply(
                            lambda x: x.str.contains(table_filter, case=False, na=False)
                        ).any(axis=1)
                        filtered_df = df_results[mask]
                    else:
                        # 預設顯示所有結果
                        filtered_df = df_results
                else:
                    filtered_df = df_results
                
                # 過濾顯示欄位
                display_columns = ['group_name', 'chart_name', 'data_type', 'data_cnt', 
                                 'ooc_cnt', 'Cpk', 'WE_Rule', 'OOB_Rule']
                available_columns = [col for col in display_columns if col in filtered_df.columns]
                
                if available_columns:
                    # 根據是否有 AG Grid 來選擇表格實現
                    if AGGRID_AVAILABLE:
                        # 使用 AG Grid 來實現可點擊的表格，無需頁面刷新（使用副本避免警告）
                        df_subset = filtered_df[available_columns].copy()
                        gb = GridOptionsBuilder.from_dataframe(df_subset)
                        gb.configure_pagination(paginationAutoPageSize=True)
                        gb.configure_side_bar()
                        gb.configure_default_column(groupable=True, value=True, enableRowGroup=True, editable=False)
                        gb.configure_selection('single')  # 單選模式
                        gb.configure_grid_options(domLayout='normal')
                        grid_options = gb.build()
                        
                        # 顯示 AG Grid（使用 DataFrame 副本避免警告）
                        grid_response = AgGrid(
                            df_subset,
                            gridOptions=grid_options,
                            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                            update_mode=GridUpdateMode.SELECTION_CHANGED,
                            fit_columns_on_grid_load=True,
                            height=470,  # 增加高度以配合新佈局
                            allow_unsafe_jscode=True,
                            enable_enterprise_modules=False
                        )
                        
                        # 儲存選擇的行數據用於右側顯示
                        selected_rows = grid_response['selected_rows']
                        if selected_rows is not None and len(selected_rows) > 0:
                            st.session_state['selected_chart_data'] = selected_rows.iloc[0].to_dict()
                        else:
                            st.session_state['selected_chart_data'] = None
                    else:
                        # 回退到標準 dataframe
                        st.info("💡 提示：點擊表格行可查看對應圖表")
                        
                        event = st.dataframe(
                            filtered_df[available_columns],
                            width='stretch',
                            hide_index=True,
                            height=500,
                            on_select="rerun",
                            selection_mode="single-row"
                        )
                        
                        # 儲存選擇的行數據
                        if event.selection.rows:
                            selected_row = event.selection.rows[0]
                            if selected_row < len(filtered_df):
                                st.session_state['selected_chart_data'] = filtered_df.iloc[selected_row].to_dict()
                            else:
                                st.session_state['selected_chart_data'] = None
                        else:
                            st.session_state['selected_chart_data'] = None
            
            with chart_col:
                # 檢查是否有選擇的項目
                selected_chart_data = st.session_state.get('selected_chart_data')
                if selected_chart_data:
                    # 找到對應的圖表結果
                    matching_chart = None
                    for chart_result in chart_results:
                        if (chart_result.get('group_name') == selected_chart_data.get('group_name') and
                            chart_result.get('chart_name') == selected_chart_data.get('chart_name')):
                            matching_chart = chart_result
                            break
                    
                    if matching_chart:
                        # 使用新的上下排列顯示函數
                        display_chart_images_vertical(matching_chart, "table_interactive")
                    else:
                        st.info("💡 選中的項目沒有對應的圖表")
                else:
                    st.info("� 點擊表格中的任意行來查看對應圖表")
    else:
        # 空狀態顯示
        st.markdown("""
        <div style='text-align: center; padding: 100px; color: #888;'>
            <h2>🎯 準備開始分析</h2>
            <p style='font-size: 18px;'>請點擊「⚙️ 檔案設定與參數」上傳檔案並設定參數</p>
            <p style='color: #666;'>然後點擊「🚀 開始分析」按鈕執行分析</p>
        </div>
        """, unsafe_allow_html=True)

def render_tool_matching_page():
    """Tool Matching 分析分頁"""
    
    if not st.session_state.api_connected:
        st.warning("⚠️ 後台 API 未連線，無法進行分析")
        return
    
    # 頂部控制欄 - 使用彈窗設定
    col_header1, col_header2, col_header3 = st.columns([1, 2, 1])
    
    with col_header1:
        # 檔案設定彈窗按鈕
        st.markdown("<br>", unsafe_allow_html=True)  # 添加一些間距
        with st.popover("📁 檔案設定"):
            # 標題列帶關閉提示
            col_title, col_close = st.columns([4, 1])
            with col_title:
                st.markdown("**📋 檔案上傳設定**")
            with col_close:
                st.markdown("*點擊外部關閉*", help="點擊彈窗外的任何地方即可關閉此設定窗口")
            
            st.divider()
            
            # 使用橫向排版
            col_upload1, col_upload2 = st.columns([1, 1])
            
            with col_upload1:
                st.write("**📊 Tool Matching 檔案**")
                tool_matching_file = render_file_uploader_with_filter("tool_matching", file_types=['csv'], title="上傳 CSV 檔案")
                
                # 檔案狀態檢查
                if tool_matching_file:
                    st.success(f"✅ {tool_matching_file.name}")
                else:
                    st.error("❌ 未上傳 Tool Matching 檔案")
            
            with col_upload2:
                st.write("**分析方法**")
                # 分析方法選擇
                analysis_method = st.radio(
                    "選擇分析方法",
                    ["指標分析", "統計檢定"],
                    help="選擇使用指標分析或統計檢定進行 Tool Matching",
                    horizontal=True
                )
            
            st.divider()
            
            # 根據分析方法顯示相應的參數設定
            if analysis_method == "指標分析":
                st.write("**📊 指標分析參數**")
                
                # 使用兩欄布局
                param_col1, param_col2 = st.columns([1, 1])
                
                with param_col1:
                    mean_threshold = st.number_input(
                        "Mean Index 門檻", 
                        min_value=0.1, 
                        max_value=10.0, 
                        value=1.0, 
                        step=0.1,
                        help="Mean Index 超過此門檻視為異常"
                    )
                
                with param_col2:
                    use_custom_sigma = st.checkbox(
                        "自訂 Sigma Index 門檻", 
                        value=False,
                        help="勾選後可自訂固定門檻，否則使用各項目的K值（基於樣本數）"
                    )
                    
                    if use_custom_sigma:
                        sigma_threshold = st.number_input(
                            "Sigma Index 門檻", 
                            min_value=0.1, 
                            max_value=10.0, 
                            value=2.0, 
                            step=0.1,
                            help="Sigma Index 超過此門檻視為異常"
                        )
                    else:
                        sigma_threshold = 2.0  # 預設值，實際會使用K值
                        st.info("📊 將使用各項目的K值作為 Sigma Index 門檻（基於樣本數）")
                
                use_statistical_test = False
                statistical_method = "unpaired"
                alpha_level = 0.05
                
            else:
                st.write("**🧮 統計檢定參數**")
                
                # 統計檢定參數
                stat_col1, stat_col2 = st.columns([1, 1])
                
                with stat_col1:
                    statistical_method = st.selectbox(
                        "檢定方法",
                        ["unpaired", "paired"],
                        index=0,
                        help="選擇統計檢定方法"
                    )
                
                with stat_col2:
                    alpha_level = st.number_input(
                        "顯著水準",
                        min_value=0.001,
                        max_value=0.1,
                        value=0.05,
                        step=0.01,
                        format="%.3f",
                        help="統計檢定的顯著水準"
                    )
                
                use_statistical_test = True
                mean_threshold = 1.0  # 統計檢定模式下的預設值
                sigma_threshold = 2.0
            
            # 使用預設的其他參數
            fill_sample_size = 5
            filter_mode = "all_data"
            base_date = None
    
    with col_header2:
        # 置中的標題區
        st.markdown("<div style='text-align: center; padding: 20px;'><h3>🔧 Tool Matching 分析系統</h3></div>", 
                   unsafe_allow_html=True)
    
    with col_header3:
        # 執行按鈕
        st.markdown("<br>", unsafe_allow_html=True)  # 添加一些間距
        if st.button("🚀 開始分析", key="tool_matching_analyze", type="primary"):
            if tool_matching_file is None:
                st.error("❌ 請先在設定中上傳 Tool Matching 檔案")
                return
            
            # 儲存檔案
            temp_dir = "temp_uploads"
            file_path = save_uploaded_file(tool_matching_file, temp_dir)
            
            # 準備 API 請求資料
            request_data = {
                "filepath": file_path,
                "mean_index_threshold": mean_threshold,
                "sigma_index_threshold": sigma_threshold,
                "use_statistical_test": use_statistical_test,
                "statistical_method": statistical_method,
                "alpha_level": alpha_level,
                "fill_sample_size": fill_sample_size,
                "filter_mode": filter_mode,
                "base_date": base_date
            }
            
            # 顯示處理中狀態
            with st.spinner("正在進行 Tool Matching 分析..."):
                result = APIClient.analyze_tool_matching(request_data)
            
            if result:
                # 保存分析結果和參數到 session state
                st.session_state.tool_matching_results = result
                st.session_state.analysis_params = {
                    "mean_threshold": mean_threshold,
                    "sigma_threshold": sigma_threshold,
                    "use_statistical_test": use_statistical_test,
                    "statistical_method": statistical_method,
                    "alpha_level": alpha_level,
                    "use_custom_sigma": use_custom_sigma if 'use_custom_sigma' in locals() else False
                }
                st.success("✅ 分析完成！")
                st.rerun()  # 刷新頁面顯示結果
                
                # 清理暫存檔案
                try:
                    import shutil
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                except:
                    pass
            else:
                st.error("❌ 分析失敗，請檢查檔案格式和後台狀態")
    
    # 分析結果區域 - 占據整個頁面寬度
    st.markdown("---")
    
    if st.session_state.tool_matching_results:
        result = st.session_state.tool_matching_results
        
        # 顯示摘要資訊
        if 'summary' in result:
            st.subheader("📊 分析摘要")
            summary = result['summary']
            
            col_metrics = st.columns(3)
            with col_metrics[0]:
                total_groups = summary.get('total_groups', 0)
                st.metric("總分析項目數", total_groups)
            with col_metrics[1]:
                abnormal_groups = summary.get('abnormal_groups', 0)
                st.metric("異常項目數", abnormal_groups)
            with col_metrics[2]:
                normal_groups = total_groups - abnormal_groups
                st.metric("正常項目數", normal_groups)
        
        # 下載功能區
        if 'results' in result and result['results']:
            st.subheader("📥 下載選項")
            col_download = st.columns(2)
            
            # 轉換為 DataFrame（使用 .copy() 避免警告）
            df_results = pd.DataFrame(result['results']).copy()
            
            with col_download[0]:
                # 下載 Excel 報告（如果有）
                if result.get('excel_output'):
                    try:
                        with open(result['excel_output'], "rb") as f:
                            excel_data = f.read()
                        
                        st.download_button(
                            label="� 下載 Excel 報告",
                            data=excel_data,
                            file_name=f"tool_matching_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"❌ Excel 檔案讀取失敗: {e}")
                else:
                    st.info("⚠️ 無 Excel 報告可供下載")
            
            with col_download[1]:
                st.empty()  # 預留空間，移除 CSV 下載
        
        # 顯示詳細結果
        if 'results' in result and result['results']:
            st.subheader("📋 詳細結果表格")
            
            # 轉換結果數據並計算異常狀態（使用保存的參數）
            df_results = pd.DataFrame(result['results']).copy()  # 使用 .copy() 避免 SettingWithCopyWarning
            processed_data = []
            
            # 獲取分析參數（如果有的話）
            analysis_params = st.session_state.get('analysis_params', {})
            user_mean_threshold = analysis_params.get('mean_threshold', 1.0)
            user_sigma_threshold = analysis_params.get('sigma_threshold', 2.0)
            user_use_statistical_test = analysis_params.get('use_statistical_test', False)
            user_use_custom_sigma = analysis_params.get('use_custom_sigma', False)
            
            for _, row in df_results.iterrows():
                # 獲取基本數據
                gname = str(row['gname'])
                cname = str(row['cname'])
                group_id = str(row['group'])
                mean_index = row['mean_index']
                sigma_index = row['sigma_index']
                k_value = row['k_value']
                
                # 判斷異常狀態（使用用戶設定的參數）
                is_abnormal = False
                abnormal_type = ""
                is_data_insufficient = (
                    mean_index == 'Insufficient Data' or 
                    sigma_index == 'Insufficient Data' or 
                    k_value == 'No Compare'
                )
                
                if not is_data_insufficient:
                    # 檢查統計檢定顯著性
                    is_statistical_significant = False
                    if isinstance(mean_index, str) and ("Significant" in str(mean_index) or "ANOVA" in str(mean_index)):
                        if "No Significant" not in str(mean_index):
                            is_statistical_significant = True
                    
                    if is_statistical_significant:
                        is_abnormal = True
                        # 檢查 sigma 是否也異常
                        try:
                            if isinstance(sigma_index, (int, float)) and isinstance(k_value, (int, float)):
                                sigma_abn = float(sigma_index) >= float(k_value)
                                abnormal_type = "Mean, Sigma" if sigma_abn else "Mean"
                            else:
                                abnormal_type = "Mean"
                        except (ValueError, TypeError):
                            abnormal_type = "Mean"
                    else:
                        # 使用門檻判斷
                        try:
                            mean_threshold = user_mean_threshold
                            sigma_threshold = user_sigma_threshold
                            
                            # 如果沒有自訂 Sigma 門檻（預設值 2.0），則使用 K 值（與 PyQt 版本邏輯一致）
                            if not user_use_custom_sigma and k_value not in [None, '', 'No Compare']:
                                try:
                                    sigma_threshold = float(k_value)
                                except (ValueError, TypeError):
                                    pass
                            
                            mean_abn = False
                            sigma_abn = False
                            
                            # 檢查 Mean Index 異常
                            if str(mean_index).lower() in ['infinite', 'inf', '-inf'] or mean_index == float('inf') or mean_index == float('-inf'):
                                mean_abn = True  # Infinite 值視為異常
                            elif isinstance(mean_index, (int, float)) and not (isinstance(mean_index, float) and (mean_index != mean_index)):  # 排除 NaN
                                mean_abn = float(mean_index) >= mean_threshold
                            
                            # 檢查 Sigma Index 異常
                            if str(sigma_index).lower() in ['infinite', 'inf', '-inf'] or sigma_index == float('inf') or sigma_index == float('-inf'):
                                sigma_abn = True  # Infinite 值視為異常
                            elif isinstance(sigma_index, (int, float)) and not (isinstance(sigma_index, float) and (sigma_index != sigma_index)):  # 排除 NaN
                                sigma_abn = float(sigma_index) >= sigma_threshold
                            
                            if mean_abn or sigma_abn:
                                is_abnormal = True
                                if mean_abn and sigma_abn:
                                    abnormal_type = "Mean, Sigma"
                                elif mean_abn:
                                    abnormal_type = "Mean"
                                elif sigma_abn:
                                    abnormal_type = "Sigma"
                        except (ValueError, TypeError):
                            pass
                
                # 格式化數值（與 PyQt 版本一致）
                def format_value(val, is_numeric=True):
                    if val in ['Insufficient Data', 'No Compare', '', None]:
                        return str(val)
                    if str(val).lower() in ['infinite', 'inf', '-inf']:
                        return "Infinite"  # 統一顯示 Infinite
                    if val == float('inf'):
                        return "Infinite"
                    if val == float('-inf'):
                        return "-Infinite"
                    if is_numeric:
                        try:
                            return f"{float(val):.2f}"
                        except (ValueError, TypeError):
                            return str(val)
                    return str(val)
                
                # 樣本數處理
                try:
                    samplesize_val = int(float(row['n'])) if pd.notna(row['n']) else 0
                except Exception:
                    samplesize_val = 0
                
                # 組裝顯示數據
                display_row = {
                    "🔍": "👁️" if is_abnormal or is_data_insufficient else "ℹ️",  # 視覺指示器
                    "異常類型": abnormal_type,
                    "群組名稱": gname,
                    "圖表名稱": cname,
                    "匹配群組": group_id,
                    "Mean Index": format_value(mean_index),
                    "Sigma Index": format_value(sigma_index),
                    "K值": format_value(k_value),
                    "均值": format_value(row['mean']),
                    "標準差": format_value(row['std']),
                    "均值中位數": format_value(row['mean_median']),
                    "標準差中位數": format_value(row['sigma_median']),
                    "樣本數": samplesize_val,
                    "特性": str(row['characteristic']),
                    "_is_abnormal": is_abnormal,
                    "_full_data": row.to_dict()
                }
                processed_data.append(display_row)
            
            # 顯示所有項目（不再僅顯示異常項目）
            all_data = processed_data  # 顯示所有處理過的數據
            
            if all_data:
                # 統計異常項目
                abnormal_count = sum(1 for item in all_data if item['_is_abnormal'])
                insufficient_count = sum(1 for item in all_data if 
                                       any(x in str(item['Mean Index']) + str(item['Sigma Index']) + str(item['K值']) 
                                           for x in ['Insufficient Data', 'No Compare']))
                
                status_msg = f"📊 **顯示所有 {len(all_data)} 個項目**"
                if abnormal_count > 0:
                    status_msg += f" - {abnormal_count} 個異常項目"
                if insufficient_count > 0:
                    status_msg += f" - {insufficient_count} 個資料不足項目"
                
                st.write(status_msg)
                
                # 創建顯示用的 DataFrame（避免 SettingWithCopyWarning）
                display_df = pd.DataFrame([{k: v for k, v in item.items() if not k.startswith('_')} 
                                         for item in all_data]).copy()
                
                # 使用標準 dataframe 顯示表格
                st.dataframe(
                    display_df,
                    width='stretch',
                    hide_index=True,
                    height=400
                )
            
            else:
                st.info("📊 **沒有任何分析結果**")

    else:
        # 空狀態顯示
        st.markdown("""
        <div style='text-align: center; padding: 100px; color: #888;'>
            <h2>🔧 準備開始 Tool Matching 分析</h2>
            <p style='font-size: 18px;'>請點擊「⚙️ 檔案設定與參數」上傳檔案並設定參數</p>
            <p style='color: #666;'>然後點擊「🚀 開始分析」按鈕執行分析</p>
        </div>
        """, unsafe_allow_html=True)


def render_spc_cpk_page():
    """渲染 SPC CPK Dashboard 頁面"""
    
    if not st.session_state.api_connected:
        st.warning("⚠️ 後台 API 未連線，無法進行分析")
        return
    
    # 頂部控制欄 - 使用彈窗設定
    col_header1, col_header2, col_header3 = st.columns([1, 2, 1])
    
    with col_header1:
        # 檔案設定彈窗按鈕
        st.markdown("<br>", unsafe_allow_html=True)  # 添加一些間距
        with st.popover("⚙️ 分析設定"):
            # 標題列帶關閉提示
            col_title, col_close = st.columns([4, 1])
            with col_title:
                st.markdown("**📋 SPC CPK 分析設定**")
            with col_close:
                st.markdown("*點擊外部關閉*", help="點擊彈窗外的任何地方即可關閉此設定窗口")
            
            st.divider()
            
            # 使用橫向排版
            col_upload1, col_upload2 = st.columns(2)
            
            with col_upload1:
                st.write("**� Chart Information 檔案**")
                chart_info_file = render_file_uploader_with_filter("spc_chart_info", file_types=['xlsx'], title="上傳 Excel 檔案")
                
                # 檔案狀態檢查
                if chart_info_file:
                    st.success(f"✅ {chart_info_file.name}")
                else:
                    st.error("❌ 未上傳 Chart Info 檔案")
            
            with col_upload2:
                st.write("**📁 原始資料檔案 (CSV)**")
                raw_data_files = render_file_uploader_with_filter("spc_raw_data", accept_multiple_files=True, file_types=['csv'], title="上傳多個 CSV 檔案")
                
                # 檔案狀態檢查
                if raw_data_files:
                    # 如果檔案數量少於等於5個，或者沒有篩選功能啟用，顯示完整清單
                    if len(raw_data_files) <= 5:
                        st.success(f"✅ {len(raw_data_files)} 個檔案")
                        for file in raw_data_files:
                            st.write(f"✅ {file.name}")
                    else:
                        st.success(f"✅ 已上傳 {len(raw_data_files)} 個檔案")
                        # 篩選功能已在 render_file_uploader_with_filter 中處理
                else:
                    st.warning("⚠️ 未上傳原始資料檔案")
            
            st.divider()
            
            # 時間範圍設定
            st.write("**📅 時間範圍設定**")
            custom_mode = st.checkbox(
                "自訂時間模式",
                help="勾選以啟用自訂時間範圍，否則使用標準三個月窗口分析"
            )
            
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input(
                    "開始日期",
                    value=pd.Timestamp.now() - pd.DateOffset(months=3),
                    help="分析的開始日期"
                )
            with col2:
                end_date = st.date_input(
                    "結束日期", 
                    value=pd.Timestamp.now(),
                    help="分析的結束日期"
                )
            
            st.divider()
            
            # 圖表選擇
            st.write("**🎯 圖表選擇設定**")
            selected_chart = st.text_input(
                "指定圖表 (可選)",
                value="",
                placeholder="例如: GroupName - ChartName",
                help="留空則分析所有圖表，或指定特定圖表進行分析"
            )
    
    with col_header2:
        # 置中的標題區
        st.markdown("<div style='text-align: center; padding: 20px;'><h3>📈 SPC CPK Dashboard</h3></div>", 
                   unsafe_allow_html=True)
    
    with col_header3:
        # 執行按鈕
        st.markdown("<br>", unsafe_allow_html=True)  # 添加一些間距
        if st.button("🚀 開始分析", key="spc_cpk_analyze", type="primary"):
            # 檢查檔案上傳
            if chart_info_file is None:
                st.error("❌ 請先在設定中上傳 Chart Information 檔案")
                return
            
            # 儲存檔案
            temp_dir = "temp_uploads"
            chart_excel_path = save_uploaded_file(chart_info_file, temp_dir)
            
            # 處理原始資料檔案
            raw_data_directory = None
            if raw_data_files:
                raw_data_directory = os.path.join(temp_dir, "raw_charts")
                for file in raw_data_files:
                    save_uploaded_file(file, raw_data_directory)
            
            # 準備請求資料
            request_data = {
                "chart_excel_path": chart_excel_path,
                "raw_data_directory": raw_data_directory,
                "start_date": start_date.isoformat() if start_date else None,
                "end_date": end_date.isoformat() if end_date else None,
                "custom_mode": custom_mode,
                "selected_chart": selected_chart if selected_chart.strip() else None
            }
            
            # 顯示分析進度
            with st.spinner("🔄 正在執行 SPC CPK 分析..."):
                result = APIClient.analyze_spc_cpk(request_data)
                
                if result:
                    st.session_state.spc_cpk_results = result
                    st.success("✅ SPC CPK 分析完成！")
                    st.rerun()
                else:
                    st.error("❌ SPC CPK 分析失敗")
    
    # 主要內容區域
    if st.session_state.spc_cpk_results:
        render_spc_cpk_results(st.session_state.spc_cpk_results)
    else:
        # 空狀態顯示
        st.markdown("""
        <div style='text-align: center; padding: 100px; color: #888;'>
            <h2>📈 準備開始 SPC CPK 分析</h2>
            <p style='font-size: 18px;'>請點擊左上角的「⚙️ 分析設定」上傳檔案並設定參數</p>
            <p style='color: #666;'>設定完成後點擊「🚀 開始分析」按鈕</p>
        </div>
        """, unsafe_allow_html=True)


def render_spc_cpk_results(results: dict):
    """渲染 SPC CPK 分析結果"""
    charts_data = results.get('charts', [])
    summary = results.get('summary', {})
    excel_path = results.get('excel_path')
    
    # 顯示分析摘要
    st.subheader("📊 分析摘要")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("圖表總數", summary.get('total_charts', 0))
    with col2:
        st.metric("分析日期範圍", f"{summary.get('date_range', 'N/A')}")
    with col3:
        mode_text = "自訂模式" if summary.get('custom_mode', False) else "標準模式"
        st.metric("分析模式", mode_text)
    with col4:
        st.metric("分析時間", summary.get('analysis_date', 'N/A'))
    
    # Excel 下載按鈕
    if excel_path:
        try:
            with open(excel_path, "rb") as file:
                st.download_button(
                    label="📥 下載 Excel 詳細報告",
                    data=file.read(),
                    file_name=f"spc_cpk_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )
        except FileNotFoundError:
            st.warning("Excel 檔案不存在或已被移動")
    
    if not charts_data:
        st.warning("沒有找到圖表資料")
        return
    
    st.markdown("---")
    
    # 圖表選擇器
    st.subheader("🎯 圖表選擇")
    chart_options = [f"{chart['group_name']} - {chart['chart_name']}" for chart in charts_data]
    
    if chart_options:
        selected_chart_idx = st.selectbox(
            "選擇要檢視的圖表",
            range(len(chart_options)),
            format_func=lambda x: chart_options[x],
            key="spc_chart_selector"
        )
        
        if selected_chart_idx is not None:
            selected_chart_data = charts_data[selected_chart_idx]
            render_single_spc_chart(selected_chart_data)
    
    st.markdown("---")
    
    # 所有圖表的指標摘要表
    st.subheader("📋 所有圖表 CPK 指標摘要")
    
    # 準備表格資料
    table_data = []
    for chart in charts_data:
        metrics = chart.get('metrics', {})
        table_data.append({
            '圖表': f"{chart['group_name']} - {chart['chart_name']}",
            '特性': chart.get('characteristics', ''),
            'USL': chart.get('usl', ''),
            'LSL': chart.get('lsl', ''),
            'Target': chart.get('target', ''),
            'Cpk (當月)': _format_metric_value(metrics.get('cpk')),
            'Cpk L1': _format_metric_value(metrics.get('cpk_l1')),
            'Cpk L2': _format_metric_value(metrics.get('cpk_l2')),
            'Long-Term Cpk': _format_metric_value(metrics.get('custom_cpk')),
            'R1 (%)': _format_metric_value(metrics.get('r1'), is_percent=True),
            'R2 (%)': _format_metric_value(metrics.get('r2'), is_percent=True),
            'K 值': _format_metric_value(metrics.get('k_value'))
        })
    
    if table_data:
        df_summary = pd.DataFrame(table_data)
        
        # 使用 AG Grid 顯示（如果可用）
        if AGGRID_AVAILABLE:
            gb = GridOptionsBuilder.from_dataframe(df_summary)
            gb.configure_pagination(paginationAutoPageSize=True)
            gb.configure_side_bar()
            gb.configure_default_column(groupable=True, value=True, enableRowGroup=True, aggFunc='sum', editable=False)
            
            # 設定欄位寬度
            gb.configure_column("圖表", width=250, pinned="left")
            gb.configure_column("特性", width=120)
            
            gridOptions = gb.build()
            
            AgGrid(
                df_summary,
                gridOptions=gridOptions,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                update_mode=GridUpdateMode.MODEL_CHANGED,
                fit_columns_on_grid_load=True,
                enable_enterprise_modules=False,
                height=400,
                width='100%'
            )
        else:
            # 使用標準 dataframe
            st.dataframe(df_summary, width='stretch', height=400)


def render_single_spc_chart(chart_data: dict):
    """渲染單一 SPC 圖表的詳細資訊"""
    st.subheader(f"📊 {chart_data['group_name']} - {chart_data['chart_name']}")
    
    # 基本資訊
    col1, col2, col3 = st.columns(3)
    with col1:
        st.write(f"**特性**: {chart_data.get('characteristics', 'N/A')}")
        st.write(f"**USL**: {chart_data.get('usl', 'N/A')}")
    with col2:
        st.write(f"**LSL**: {chart_data.get('lsl', 'N/A')}")
        st.write(f"**Target**: {chart_data.get('target', 'N/A')}")
    with col3:
        st.write(f"**K 值**: {_format_metric_value(chart_data.get('metrics', {}).get('k_value'))}")
    
    # CPK 指標卡片
    st.markdown("#### CPK 指標")
    metrics = chart_data.get('metrics', {})
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        cpk_val = _format_metric_value(metrics.get('cpk'))
        st.metric("Cpk (當月)", cpk_val)
    with col2:
        l1_val = _format_metric_value(metrics.get('cpk_l1'))
        st.metric("L1 Cpk", l1_val)
    with col3:
        l2_val = _format_metric_value(metrics.get('cpk_l2'))
        st.metric("L2 Cpk", l2_val)
    with col4:
        custom_val = _format_metric_value(metrics.get('custom_cpk'))
        st.metric("Long-Term Cpk", custom_val)
    with col5:
        r1_val = _format_metric_value(metrics.get('r1'), is_percent=True)
        st.metric("R1", r1_val)
    with col6:
        r2_val = _format_metric_value(metrics.get('r2'), is_percent=True)
        st.metric("R2", r2_val)
    
    # 顯示 SPC 圖表
    chart_image = chart_data.get('chart_image')
    if chart_image:
        st.markdown("#### SPC 控制圖")
        try:
            # 解碼 base64 圖片
            image_data = base64.b64decode(chart_image)
            image = Image.open(io.BytesIO(image_data))
            # 調整圖片大小為高解析度橫向長方形
            target_width = 1600  # 提高解析度
            target_height = 800
            image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)
            st.image(image, width='stretch')
        except Exception as e:
            st.error(f"無法顯示圖表: {e}")
    
    # 統計資訊表
    st.markdown("#### 統計資訊")
    
    stats_data = {
        '時段': ['L0', 'L1', 'L2', 'Long-Term'],
        'Mean': [
            _format_metric_value(chart_data.get('mean_current')),
            _format_metric_value(chart_data.get('mean_last_month')),
            _format_metric_value(chart_data.get('mean_last2_month')),
            _format_metric_value(chart_data.get('mean_all'))
        ],
        'Sigma': [
            _format_metric_value(chart_data.get('sigma_current')),
            _format_metric_value(chart_data.get('sigma_last_month')),
            _format_metric_value(chart_data.get('sigma_last2_month')),
            _format_metric_value(chart_data.get('sigma_all'))
        ]
    }
    
    df_stats = pd.DataFrame(stats_data)
    st.dataframe(df_stats, width='stretch')


def _format_metric_value(value, is_percent: bool = False) -> str:
    """格式化指標值顯示"""
    if value is None:
        return "-"
    
    try:
        if isinstance(value, (int, float)):
            if is_percent:
                return f"{value:.1f}%"
            else:
                return f"{value:.3f}"
        else:
            return str(value)
    except:
        return str(value) if value is not None else "-"


def main():
    """主函數"""
    # 初始化
    init_session_state()
    
    # 標題和導航
    st.title("📊 OOB & Tool Matching 分析系統")
    st.markdown("---")
    
    # 檢查 API 連線狀態
    check_api_connection()
    
    # 建立分頁
    tab1, tab2, tab3 = st.tabs(["OOB/SPC 分析", "Tool Matching", "SPC CPK Dashboard"])
    
    with tab1:
        render_oob_page()
    
    with tab2:
        render_tool_matching_page()
    
    with tab3:
        render_spc_cpk_page()
    
    # 側邊欄資訊
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📋 系統資訊")
        st.markdown(f"**API 位址**: {API_BASE_URL}")
        st.markdown(f"**更新時間**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        if st.button("🔄 重新檢查連線", key="refresh_connection"):
            check_api_connection()
            st.rerun()

if __name__ == "__main__":
    main()