import os
import sys
import math
import pandas as pd
import numpy as np
import matplotlib.dates as mdates
import os
import sys
import math
import pandas as pd
import numpy as np
import tempfile
import xlsxwriter
import matplotlib.dates as mdates
import matplotlib.transforms as mtransforms
import importlib.util
import scipy.stats as stats
from PyQt6 import QtWidgets, QtCore, QtGui
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

def calculate_cpk(raw_df, chart_info):
    print(raw_df)
    print(f"[DEBUG] chart_info type: {type(chart_info)}")
    print(f"[DEBUG] chart_info keys: {list(chart_info.keys()) if hasattr(chart_info, 'keys') else chart_info}")
    print(f"[DEBUG] chart_info: {chart_info}")
    mean = raw_df['point_val'].mean()
    print(f"[DEBUG] mean: {mean}")
    std = raw_df['point_val'].std()
    print(f"[DEBUG] std: {std}")
    characteristic = chart_info['Characteristics']
    usl = chart_info.get('USL', None)
    print(f"[DEBUG] usl: {usl}")
    lsl = chart_info.get('LSL', None)
    print(f"[DEBUG] lsl: {lsl}")
    print(f"[DEBUG] usl: {usl}, lsl: {lsl}, characteristic: {characteristic}")
    cpk = None
    if std > 0:
        if characteristic == 'Nominal':
            if usl is not None and lsl is not None:
                cpu = (usl - mean) / (3 * std)
                cpl = (mean - lsl) / (3 * std)
                cpk = min(cpu, cpl)
        elif characteristic in ['Smaller', 'Sigma']:  # 支援 Sigma 特性，邏輯與 Smaller 相同
            if usl is not None:
                cpk = (usl - mean) / (3 * std)
        elif characteristic == 'Bigger':
            if lsl is not None:
                cpk = (mean - lsl) / (3 * std)
    if cpk is not None:
        cpk = round(cpk, 3)
    return {'Cpk': cpk}

class SPCCpkDashboard(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("SPC Cpk Dashboard")
        self.resize(1400, 900)  # 增加視窗寬度以容納新圖表
        # 資料結構初始化
        self.all_charts_info = None
        self.raw_charts_dict = {}
        self.cpk_results = {}  # {(group_name, chart_name): {'Cpk': value}}
        self.chart_date_states = {}  # 每張圖的日期狀態：{'custom': bool, 'start': date, 'end': date}
        self.axis_mode = 'index'  # 'index' (等距) 或 'time'
        self.init_ui()


    def init_ui(self):
        # 重新打造為 Dashboard 版型
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(14)
        # ===== Top Filter / Action Bar =====
        top_bar = QtWidgets.QHBoxLayout()
        top_bar.setSpacing(12)
        self.chart_combo = QtWidgets.QComboBox()
        self.chart_combo.addItem("")
        self.chart_combo.setMinimumWidth(280)
        self.start_date = QtWidgets.QDateEdit(QtCore.QDate.currentDate().addMonths(-3))
        self.start_date.setCalendarPopup(True)
        self.end_date = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        # 只保留執行分析按鈕，並重新設計
        self.recalc_btn = QtWidgets.QPushButton("Run Analysis")
        self.recalc_btn.setMinimumHeight(38)
        self.recalc_btn.setMinimumWidth(120)
        self.recalc_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #1d4fd8);
                color: #fff;
                border: none;
                border-radius: 18px;
                font-size: 16px;
                font-weight: bold;
                padding: 8px 24px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4fd8, stop:1 #2563eb);
            }
            QPushButton:pressed {
                background: #163fae;
            }
        """)
        # 新增下載 Excel 按鈕
        self.export_excel_btn = QtWidgets.QPushButton("Download Cpk Detail")
        self.export_excel_btn.setMinimumHeight(38)
        self.export_excel_btn.setMinimumWidth(180)
        self.export_excel_btn.setStyleSheet(self.recalc_btn.styleSheet())
        lbl_chart = QtWidgets.QLabel("Chart:")
        lbl_chart.setObjectName("plainLabel")
        lbl_start = QtWidgets.QLabel("Start:")
        lbl_start.setObjectName("plainLabel")
        lbl_end = QtWidgets.QLabel("End:")
        lbl_end.setObjectName("plainLabel")
        
        # 新增自訂時間模式按鈕
        self.custom_range_btn = QtWidgets.QPushButton("Custom Time Mode")
        self.custom_range_btn.setMinimumHeight(38)
        self.custom_range_btn.setMinimumWidth(140)
        self.custom_range_btn.setCheckable(True)  # 可切換狀態
        self.custom_range_btn.setStyleSheet("""
            QPushButton {
                background: #6b7280;
                color: #fff;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 600;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background: #4b5563;
            }
            QPushButton:checked {
                background: #059669;
                color: #fff;
            }
            QPushButton:checked:hover {
                background: #047857;
            }
        """)
        
        top_bar.addWidget(lbl_chart)
        top_bar.addWidget(self.chart_combo)
        top_bar.addSpacing(6)
        top_bar.addWidget(lbl_start)
        top_bar.addWidget(self.start_date)
        top_bar.addWidget(lbl_end)
        top_bar.addWidget(self.end_date)
        top_bar.addWidget(self.custom_range_btn)
        top_bar.addStretch(1)
        top_bar.addWidget(self.recalc_btn)
        top_bar.addWidget(self.export_excel_btn)
        root.addLayout(top_bar)
        # ===== Metric Cards Row =====
        self.metric_cards = {}
        cards_layout = QtWidgets.QGridLayout()
        cards_layout.setHorizontalSpacing(16)
        cards_layout.setVerticalSpacing(14)
        def create_metric_card(key, title, col, row=0):
            frame = QtWidgets.QFrame()
            frame.setObjectName("metricCard")
            frame.setProperty("status", "neutral")
            pal = frame.palette()
            pal.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#ffffff"))
            frame.setAutoFillBackground(True)
            frame.setPalette(pal)
            layout = QtWidgets.QVBoxLayout(frame)
            layout.setContentsMargins(16, 12, 16, 12)
            layout.setSpacing(4)
            title_label = QtWidgets.QLabel(title)
            title_label.setObjectName("metricTitle")
            title_label.setAutoFillBackground(True)
            tpal = title_label.palette()
            tpal.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#ffffff"))
            title_label.setPalette(tpal)
            value_label = QtWidgets.QLabel("-")
            value_label.setObjectName("metricValue")
            value_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter)
            value_label.setAutoFillBackground(True)
            vpal = value_label.palette()
            vpal.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#ffffff"))
            value_label.setPalette(vpal)
            layout.addWidget(title_label)
            layout.addWidget(value_label)
            layout.addStretch(1)
            cards_layout.addWidget(frame, row, col)
            self.metric_cards[key] = {"frame": frame, "value_label": value_label, "title_label": title_label}
        create_metric_card("cpk", "Cpk", 0)
        create_metric_card("l1", "L1 Cpk", 1)
        create_metric_card("l2", "L2 Cpk", 2)
        create_metric_card("custom", "Long-Term Cpk", 3)
        create_metric_card("r1", "R1", 4)
        create_metric_card("r2", "R2", 5)
        create_metric_card("kval", "K", 6)
        root.addLayout(cards_layout)
        # ===== Chart Area =====
        self.chart_frame = QtWidgets.QFrame()
        self.chart_frame.setObjectName("chartFrame")
        chart_layout = QtWidgets.QVBoxLayout(self.chart_frame)
        chart_layout.setContentsMargins(18, 16, 18, 16)
        chart_layout.setSpacing(8)
        header = QtWidgets.QHBoxLayout()
        title_lbl = QtWidgets.QLabel("SPC Chart Analysis")
        title_lbl.setObjectName("sectionTitle")
        header.addWidget(title_lbl)
        header.addStretch(1)
        
        # 加上切換按鈕
        self.prev_chart_btn = QtWidgets.QPushButton("◀ Prev")
        self.prev_chart_btn.setMinimumHeight(32)
        self.prev_chart_btn.setMinimumWidth(80)
        self.prev_chart_btn.setStyleSheet("""
            QPushButton {
                background: #6b7280;
                color: #fff;
                border: none;
                border-radius: 6px;
                font-size: 12px;
                font-weight: 500;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background: #4b5563;
            }
            QPushButton:pressed {
                background: #374151;
            }
        """)
        
        self.next_chart_btn = QtWidgets.QPushButton("Next ▶")
        self.next_chart_btn.setMinimumHeight(32)
        self.next_chart_btn.setMinimumWidth(80)
        self.next_chart_btn.setStyleSheet(self.prev_chart_btn.styleSheet())
        
        header.addWidget(self.prev_chart_btn)
        header.addWidget(self.next_chart_btn)
        chart_layout.addLayout(header)
        
        # 修改：使用 Figure 但配置成 GridSpec 布局（主圖大，右邊兩個小圖）
        self.figure = Figure(figsize=(12, 6))
        self.canvas = FigureCanvas(self.figure)
        chart_layout.addWidget(self.canvas, 1)
        root.addWidget(self.chart_frame, 1)
        # 事件連接
        self.recalc_btn.clicked.connect(self.recalculate)
        self.chart_combo.currentIndexChanged.connect(self.update_cpk_labels)
        self.start_date.dateChanged.connect(self.on_date_changed)
        self.end_date.dateChanged.connect(self.on_date_changed)
        self.export_excel_btn.clicked.connect(self.export_chart_info_excel)
        self.prev_chart_btn.clicked.connect(self.prev_chart)
        self.next_chart_btn.clicked.connect(self.next_chart)
        self.custom_range_btn.clicked.connect(self.on_custom_range_toggle)
        self.apply_theme()
    def export_chart_info_excel(self):
        # 匯出所有 chart 的 group_name@chart_name@characteristics 及 Cpk 指標到 Excel，並加上 debug log
        if self.all_charts_info is None:
            QtWidgets.QMessageBox.warning(self, "No Data", "Chart information not loaded yet!")
            return
        rows = []
        chart_images = []
        for _, chart_info in self.all_charts_info.iterrows():
            group_name = str(chart_info.get('GroupName', ''))
            chart_name = str(chart_info.get('ChartName', ''))
            characteristics = str(chart_info.get('Characteristics', ''))
            usl = chart_info.get('USL', None)
            lsl = chart_info.get('LSL', None)
            target = None
            for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                if key_t in chart_info and pd.notna(chart_info[key_t]):
                    target = chart_info[key_t]
                    break
            key = (group_name, chart_name)
            cpk = None
            cpk_last_month = None
            cpk_last2_month = None
            custom_cpk = None
            r1 = None
            r2 = None
            mean_month = sigma_month = mean_last_month = sigma_last_month = mean_last2_month = sigma_last2_month = mean_all = sigma_all = None
            # 讓匯出時 end_date 也抓 chart 最新資料日期
            export_end_date = self.end_date.date().toPyDate() if hasattr(self, 'end_date') else None
            if key in self.raw_charts_dict:
                raw_df = self.raw_charts_dict[key]
                if raw_df is not None and not raw_df.empty and 'point_time' in raw_df.columns:
                    raw_df_local = raw_df.copy()
                    raw_df_local['point_time'] = pd.to_datetime(raw_df_local['point_time'])
                    latest = raw_df_local['point_time'].max().date()
                    export_end_date = latest
                    start1 = pd.to_datetime(export_end_date) - pd.DateOffset(months=1)
                    print(f"[DEBUG][Excel] {group_name}@{chart_name} Cpk區間: {start1.date()} ~ {export_end_date}")
                if raw_df is not None:
                    print(f"[DEBUG] raw_df shape: {raw_df.shape}")
                else:
                    print(f"[DEBUG] raw_df is None")
                cpk_res = self._recompute_cpk_for_chart(chart_info, export_end_date)
                print(f"[DEBUG] cpk_res: {cpk_res}")
                cpk = cpk_res.get('Cpk')
                cpk_last_month = cpk_res.get('Cpk_last_month')
                cpk_last2_month = cpk_res.get('Cpk_last2_month')
                if raw_df is not None:
                    custom_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
                if cpk is not None and cpk_last_month is not None and cpk_last_month != 0 and cpk <= cpk_last_month:
                    r1 = (1 - (cpk / cpk_last_month)) * 100
                if cpk is not None and cpk_last_month is not None and cpk_last2_month is not None and cpk_last2_month != 0 and cpk <= cpk_last_month <= cpk_last2_month:
                    r2 = (1 - (cpk / cpk_last2_month)) * 100
                # 計算四個區間的 mean, sigma 並印出
                def print_mean_sigma(df, label, group_name, chart_name):
                    if df is not None and not df.empty:
                        mean = df['point_val'].mean()
                        sigma = df['point_val'].std()
                        print(f"[MEAN_SIGMA][{group_name}@{chart_name}][{label}] mean: {mean:.4f}, sigma: {sigma:.4f}")
                    else:
                        print(f"[MEAN_SIGMA][{group_name}@{chart_name}][{label}] No Data")
                # 取得各區間資料
                if raw_df is not None and not raw_df.empty and 'point_time' in raw_df.columns:
                    raw_df_local = raw_df.copy()
                    raw_df_local['point_time'] = pd.to_datetime(raw_df_local['point_time'])
                    end_time = pd.to_datetime(export_end_date)
                    start1 = end_time - pd.DateOffset(months=1)
                    start2 = end_time - pd.DateOffset(months=2)
                    start3 = end_time - pd.DateOffset(months=3)
                    df_all = raw_df_local[raw_df_local['point_time'] <= end_time]
                    df_month = raw_df_local[(raw_df_local['point_time'] > start1) & (raw_df_local['point_time'] <= end_time)]
                    df_last_month = raw_df_local[(raw_df_local['point_time'] > start2) & (raw_df_local['point_time'] <= start1)]
                    df_last2_month = raw_df_local[(raw_df_local['point_time'] > start3) & (raw_df_local['point_time'] <= start2)]
                    mean_month = df_month['point_val'].mean() if not df_month.empty else None
                    sigma_month = df_month['point_val'].std() if not df_month.empty else None
                    mean_last_month = df_last_month['point_val'].mean() if not df_last_month.empty else None
                    sigma_last_month = df_last_month['point_val'].std() if not df_last_month.empty else None
                    mean_last2_month = df_last2_month['point_val'].mean() if not df_last2_month.empty else None
                    sigma_last2_month = df_last2_month['point_val'].std() if not df_last2_month.empty else None
                    mean_all = df_all['point_val'].mean() if not df_all.empty else None
                    sigma_all = df_all['point_val'].std() if not df_all.empty else None
                    print_mean_sigma(df_month, '當月', group_name, chart_name)
                    print_mean_sigma(df_last_month, '上月', group_name, chart_name)
                    print_mean_sigma(df_last2_month, '上上月', group_name, chart_name)
                    print_mean_sigma(df_all, '全部', group_name, chart_name)
            # --- 新增：用與 UI 完全一致的方式繪製圖表並存成圖片 ---
            # 計算 K 參數
            kval = None
            try:
                if target is not None and usl is not None and lsl is not None:
                    mean_val = None
                    plot_df = self.raw_charts_dict.get(key)
                    if plot_df is not None and not plot_df.empty:
                        mean_val = plot_df['point_val'].mean()
                    rng = (usl - lsl) / 2 if (usl is not None and lsl is not None and (usl-lsl)!=0) else None
                    if mean_val is not None and rng:
                        kval = abs(mean_val - target) / rng
            except Exception:
                kval = None
            import tempfile
            tmp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            tmp_img.close()
            # 用 draw_spc_chart 的繪圖邏輯（x軸刻度與UI一致，點與點等距）
            from matplotlib.figure import Figure
            fig = Figure(figsize=(10, 4))
            ax = fig.add_subplot(111)
            characteristics = chart_info.get('Characteristics', '')
            ax.set_title(f"{group_name}@{chart_name}@{characteristics}", pad=18)
            ax.set_xlabel("")
            ax.set_ylabel("")
            plot_df = self.raw_charts_dict.get(key)
            if plot_df is None or plot_df.empty:
                ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
            else:
                plot_df2 = plot_df.copy()
                # 日期過濾 (若有 point_time 欄位)
                if 'point_time' in plot_df2.columns:
                    try:
                        plot_df2['point_time'] = pd.to_datetime(plot_df2['point_time'])
                        plot_df2 = plot_df2.sort_values('point_time').reset_index(drop=True)
                    except Exception:
                        pass
                y = plot_df2['point_val'].values
                # x軸等距模式（與UI一致）
                x = range(1, len(y) + 1)
                # === 在圖上標示「當月/上月/上上月」區間 ===
                if 'point_time' in plot_df2.columns and not plot_df2.empty:
                    try:
                        import matplotlib.transforms as mtransforms
                        times = pd.to_datetime(plot_df2['point_time']).to_numpy()
                        tmin, tmax = times.min(), times.max()
                        end_sel = pd.Timestamp(tmax)
                        start1 = end_sel - pd.DateOffset(months=1)
                        start2 = end_sel - pd.DateOffset(months=2)
                        start3 = end_sel - pd.DateOffset(months=3)
                        windows = [
                            (start1, end_sel,  'L0',   '#dbeafe'),
                            (start2, start1,   'L1',   '#fef9c3'),
                            (start3, start2,   'L2',   '#ede9fe'),
                        ]
                        text_trans = mtransforms.blended_transform_factory(ax.transData, ax.transAxes)
                        import numpy as np
                        n = len(times)
                        def t2ix_left(t):
                            return float(np.searchsorted(times, np.datetime64(t), side='left')) + 0.5
                        def t2ix_right(t):
                            return float(np.searchsorted(times, np.datetime64(t), side='right')) + 0.5
                        x_min, x_max = 0.5, n + 0.5
                        for s, e, lab, col in windows:
                            s_clip = max(pd.Timestamp(s), pd.Timestamp(tmin))
                            e_clip = min(pd.Timestamp(e), pd.Timestamp(tmax))  # 右邊界只到資料最後一天
                            if e_clip <= s_clip:
                                continue
                            xl = max(x_min, t2ix_left(s_clip))
                            xr = min(x_max, t2ix_right(e_clip))
                            if xr <= xl:
                                continue
                            ax.axvspan(xl, xr, color=col, alpha=0.25, zorder=0)
                            x_center = (xl + xr) / 2.0
                            ax.text(x_center, 1.04, lab, transform=text_trans, ha='center', va='top', fontsize=9, color='#374151', alpha=0.9)
                    except Exception as _:
                        pass
                # 主數據線
                ax.plot(x, y, linestyle='-', marker='o', color='#2563eb', markersize=5, linewidth=1.2, label='_nolegend_')
                usl = chart_info.get('USL', None)
                lsl = chart_info.get('LSL', None)
                target = None
                for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                    if key_t in chart_info and pd.notna(chart_info[key_t]):
                        target = chart_info[key_t]
                        break
                import numpy as np
                mean_val = float(np.mean(y)) if len(y) else None
                # 超規點
                if usl is not None:
                    ax.scatter([xi for xi, yi in zip(x, y) if yi > usl], [yi for yi in y if yi > usl], color='#dc2626', s=36, zorder=5, label='_nolegend_')
                if lsl is not None:
                    ax.scatter([xi for xi, yi in zip(x, y) if yi < lsl], [yi for yi in y if yi < lsl], color='#dc2626', marker='s', s=36, zorder=5, label='_nolegend_')
                # Y軸範圍（納入 USL/LSL/Target/Mean）
                extra_vals = [v for v in [usl, lsl, target, mean_val]
                              if v is not None and not (isinstance(v, float) and np.isnan(v))]
                if len(y) > 0:
                    ymin_sel = float(np.min(y))
                    ymax_sel = float(np.max(y))
                else:
                    ymin_sel, ymax_sel = (0.0, 1.0)
                if extra_vals:
                    ymin_sel = min(ymin_sel, min(extra_vals))
                    ymax_sel = max(ymax_sel, max(extra_vals))
                rng = ymax_sel - ymin_sel
                margin = 0.05 * rng if rng > 0 else 1.0
                ax.set_ylim(ymin_sel - margin, ymax_sel + margin)
                # 畫短水平線，文字貼右
                from matplotlib import transforms as mtransforms
                trans = mtransforms.blended_transform_factory(ax.transAxes, ax.transData)
                def segment_with_label(val, name, color, va='center'):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return
                    x0, x1 = 0.0, 0.965
                    ax.plot([x0, x1], [val, val], transform=trans, color=color, linestyle='--', linewidth=1.1)
                    ax.text(x1, val, name, transform=trans, color=color, va=va, ha='left', fontsize=9)
                segment_with_label(usl, 'USL', '#ef4444', va='center')
                segment_with_label(lsl, 'LSL', '#ef4444', va='center')
                segment_with_label(target, 'Target', '#f59e0b', va='center')
                segment_with_label(mean_val, 'Mean', '#16a34a', va='center')
                # x軸刻度（等距模式顯示日期）
                if 'point_time' in plot_df2.columns and not plot_df2.empty:
                    times = plot_df2['point_time'].tolist()
                    total = len(times)
                    if total <= 12:
                        tick_idx = list(range(1, total + 1))
                    else:
                        step = max(1, total // 8)
                        tick_idx = list(range(1, total + 1, step))
                        if tick_idx[-1] != total:
                            tick_idx.append(total)
                    labels = [times[i-1].strftime('%Y-%m-%d') for i in tick_idx]
                    ax.set_xticks(tick_idx)
                    ax.set_xticklabels(labels, rotation=45, ha='center', fontsize=8)
                ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.5)
            fig.tight_layout()
            fig.savefig(tmp_img.name)
            chart_images.append(tmp_img.name)
            rows.append({
                'ChartImage': '',  # 佔位，稍後插入圖片
                'ChartKey': f"{group_name}@{chart_name}@{characteristics}",
                'GroupName': group_name,
                'ChartName': chart_name,
                'Characteristics': characteristics,
                'USL': usl,
                'LSL': lsl,
                'Target': target,
                'K': kval,
                'Cpk_Curr': cpk,
                'Cpk_L1': cpk_last_month,
                'Cpk_L2': cpk_last2_month,
                'Custom_Cpk': custom_cpk,
                'R1(%)': r1,
                'R2(%)': r2,
                'Mean_Curr': mean_month,
                'Sigma_CurrentMonth': sigma_month,
                'Mean_LastMonth': mean_last_month,
                'Sigma_LastMonth': sigma_last_month,
                'Mean_Last2Month': mean_last2_month,
                'Sigma_Last2Month': sigma_last2_month,
                'Mean_All': mean_all,
                'Sigma_All': sigma_all
            })
        df = pd.DataFrame(rows)
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Download Chart Information Excel", "chart_info.xlsx", "Excel Files (*.xlsx)")
        if path:
            try:
                import xlsxwriter
                # 欄位順序：圖片 + 其他欄位
                columns = ['ChartImage'] + [c for c in df.columns if c != 'ChartImage']
                workbook = xlsxwriter.Workbook(path)
                worksheet = workbook.add_worksheet()
                # 設定欄寬（只調整第一欄圖片欄位）
                worksheet.set_column(0, 0, 120)  # 圖片欄位寬度設為 120
                for i in range(1, len(columns)):
                    worksheet.set_column(i, i, 15)  # 其他欄位保持預設寬度 15
                # 標題粗體
                bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
                cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                for col_idx, col_name in enumerate(columns):
                    worksheet.write(0, col_idx, col_name, bold)
                
                # 添加自動篩選功能
                last_row = len(df)
                last_col = len(columns) - 1
                worksheet.autofilter(0, 0, last_row, last_col)
                
                # 寫入資料 & 插入圖片
                x_scale_factor = 1.0  # 水平縮放保持原始大小
                y_scale_factor = 0.6  # 垂直縮放調整為 0.6，讓圖片變扁
                for row_idx, row in enumerate(df.to_dict('records')):
                    # 設定圖片插入選項
                    x_offset = 0
                    y_offset = 15
                    options = {
                        'x_scale': x_scale_factor,
                        'y_scale': y_scale_factor,
                        'x_offset': x_offset,
                        'y_offset': y_offset,
                        'object_position': 1
                    }
                    
                    # 插入圖片
                    img_path = chart_images[row_idx]
                    worksheet.set_row(row_idx+1, 200)  # 行高調整為 200，配合扁平圖片
                    worksheet.insert_image(row_idx+1, 0, img_path, options)
                    
                    # 其他欄位（置中）
                    for col_idx, col_name in enumerate(columns[1:], 1):
                        val = row.get(col_name, '')
                        # 修正 NaN/Inf 問題
                        import math
                        if val is None:
                            val = ''
                        elif isinstance(val, float):
                            if math.isnan(val) or math.isinf(val):
                                val = 'N/A'
                        worksheet.write(row_idx+1, col_idx, val, cell_format)
                workbook.close()
                QtWidgets.QMessageBox.information(self, "Export Success", f"Excel exported to: {path}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Export Failed", f"Excel export failed: {e}")

    # === 檔案載入 ===
    def load_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv)")
        if path:
            try:
                self.data = pd.read_csv(path)
                self.file_label.setText(f"Loaded: {os.path.basename(path)}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Load failed: {e}")
                self.file_label.setText("Load failed")
            self.recalculate()

    def load_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv)")
        if path:
            try:
                self.data = pd.read_csv(path)
                self.file_label.setText(f"Loaded: {os.path.basename(path)}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Load failed: {e}")
                self.file_label.setText("Load failed")
            self.recalculate()

    # === 重新計算 ===
    def recalculate(self):
        print("[DEBUG] recalculate called")
        # 重新載入 chart 資訊
        def get_app_dir():
            import sys, os
            if getattr(sys, 'frozen', False):
                return os.path.dirname(sys.executable)
            else:
                return os.path.dirname(__file__)
        
        chart_excel_path = os.path.join(get_app_dir(), 'input', 'All_Chart_Information.xlsx')
        try:
            # 優先使用靜態 import（打包時 PyInstaller 會自動包含 oob_module）
            try:
                import oob_module
                self.oob_module = oob_module
            except ImportError:
                # 開發環境 fallback：若尚未改名，嘗試從原始 0621.py 檔案載入
                import importlib.util
                import sys
                
                if getattr(sys, 'frozen', False):
                    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
                else:
                    base_path = os.path.dirname(__file__)
                
                legacy_path = os.path.join(base_path, "0621.py")
                if os.path.exists(legacy_path):
                    spec = importlib.util.spec_from_file_location("oob_module", legacy_path)
                    oob_module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(oob_module)
                    self.oob_module = oob_module
                else:
                    raise ImportError("Cannot load oob_module; please ensure oob_module.py exists and can be imported")
            
            self.all_charts_info = self.oob_module.load_chart_information(chart_excel_path)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load chart information: {e}")
            self.all_charts_info = None
            return

        # 更新下拉選單
        if self.all_charts_info is not None:
            self.chart_combo.clear()
            self.chart_combo.addItem("Select Chart")
            for _, chart_info in self.all_charts_info.iterrows():
                self.chart_combo.addItem(f"{chart_info['GroupName']} - {chart_info['ChartName']}")

        # 重置資料結構
        self.raw_charts_dict = {}
        self.cpk_results = {}
        self.chart_date_states = {}

        if self.all_charts_info is not None:
            raw_data_dir = os.path.join(get_app_dir(), 'input', 'raw_charts')
            for _, chart_info in self.all_charts_info.iterrows():
                if not isinstance(chart_info, pd.Series):
                    continue
                group_name = str(chart_info['GroupName'])
                chart_name = str(chart_info['ChartName'])
                raw_path = self.oob_module.find_matching_file(raw_data_dir, group_name, chart_name)
                if raw_path and os.path.exists(raw_path):
                    try:
                        raw_df = pd.read_csv(raw_path)
                        
                        # Debug: 檢查欄位
                        print(f"[DEBUG] {group_name}/{chart_name} 欄位: {list(raw_df.columns)}")
                        if 'EQP_id' in raw_df.columns:
                            unique_eqp = raw_df['EQP_id'].unique()
                            print(f"[DEBUG] EQP_id 發現: {unique_eqp}")
                        
                        usl = chart_info.get('USL', None)
                        lsl = chart_info.get('LSL', None)
                        if usl is not None and lsl is not None:
                            raw_df = raw_df[(raw_df['point_val'] <= usl) & (raw_df['point_val'] >= lsl)]
                        elif usl is not None:
                            raw_df = raw_df[raw_df['point_val'] <= usl]
                        elif lsl is not None:
                            raw_df = raw_df[raw_df['point_val'] >= lsl]
                        self.raw_charts_dict[(group_name, chart_name)] = raw_df
                        quick_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
                        self.cpk_results[(group_name, chart_name)] = {'Cpk': quick_cpk}
                        self.chart_date_states[(group_name, chart_name)] = {'custom': False, 'start': None, 'end': None}
                    except Exception as e:
                        self.raw_charts_dict[(group_name, chart_name)] = None
                        self.cpk_results[(group_name, chart_name)] = {'Cpk': None}
                        print(f"[ERROR] raw chart 載入失敗 {group_name}/{chart_name}: {e}")
                else:
                    self.raw_charts_dict[(group_name, chart_name)] = None
                    self.cpk_results[(group_name, chart_name)] = {'Cpk': None}
        # 清空圖表並等待選擇
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_title("SPC Control Chart (Not Plotted Yet)")
        # ax.set_xlabel("日期")
        # ax.set_ylabel("值")
        self.canvas.draw()
        self.update_cpk_labels()
        
        # 顯示執行完成提示
        QtWidgets.QMessageBox.information(self, "Analysis Complete", 
                                         f"Analysis completed successfully!\n"
                                         f"Charts loaded: {len(self.raw_charts_dict) if self.raw_charts_dict else 0}\n"
                                         f"Please select a chart to view the results.")

    def apply_theme(self, mode: str = "light"):
        if mode == "light":
            self.setStyleSheet("""
            QWidget { background:#eef1f5; color:#222; font-family:'Microsoft YaHei'; font-size:13px; }
            QComboBox, QDateEdit { background:#ffffff; border:1px solid #c5ccd4; padding:4px 8px; border-radius:7px; }
            QComboBox:hover, QDateEdit:hover { border:1px solid #98a3af; }
            QPushButton { background:#2563eb; color:#fff; border:none; padding:7px 18px; border-radius:8px; font-weight:600; }
            QPushButton:hover { background:#1d4fd8; }
            QPushButton:pressed { background:#163fae; }
            QFrame#metricCard, QFrame#metricCard * { background:#ffffff !important; }
            QFrame#metricCard { border:1px solid #d8dde2; border-radius:16px; }
            QLabel#metricTitle { font-size:11px; font-weight:600; color:#6c7681; letter-spacing:1px; }
            QLabel#metricValue { font-size:30px; font-weight:700; color:#111827; }
            QFrame#metricCard:hover { border:1px solid #aeb5bb; }
            QFrame#chartFrame { background:#ffffff; border:1px solid #d2d7dc; border-radius:22px; }
            QLabel#sectionTitle { font-size:15px; font-weight:600; color:#1f2937; background:transparent; }
            QLabel#plainLabel { font-size:13px; font-weight:600; color:#1f2937; background:transparent; }
            """)
        for meta in self.metric_cards.values():
            if meta["frame"].graphicsEffect() is None:
                eff = QtWidgets.QGraphicsDropShadowEffect(self)
                eff.setBlurRadius(18)
                eff.setOffset(0, 4)
                eff.setColor(QtGui.QColor(0, 0, 0, 26))
                meta["frame"].setGraphicsEffect(eff)
        if self.chart_frame.graphicsEffect() is None:
            eff2 = QtWidgets.QGraphicsDropShadowEffect(self)
            eff2.setBlurRadius(28)
            eff2.setOffset(0, 5)
            eff2.setColor(QtGui.QColor(0, 0, 0, 30))
            self.chart_frame.setGraphicsEffect(eff2)

    # ==== 重複定義刪除 (上方已有 recalculate) ====

    # (duplicate apply_theme & recalculate removed)

    def _apply_card_status(self, key: str, status: str):
        # 不再改變邊框顏色，保持固定樣式
        return

    def update_cpk_labels(self):
        """選擇 chart 時：若該 chart 尚未自訂日期 -> 自動用最新往回三個月，之後使用者調整不再被覆蓋。"""
        idx = self.chart_combo.currentIndex() - 1
        for key, comp in self.metric_cards.items():
            comp["value_label"].setText("-")
        if idx < 0 or self.all_charts_info is None:
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.set_title("SPC Control Chart (Not Selected)")
            self.canvas.draw()
            return
        chart_info = self.all_charts_info.iloc[idx]
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        key = (group_name, chart_name)
        raw_df = self.raw_charts_dict.get(key)
        state = self.chart_date_states.get(key)
        if state is None:
            state = {'custom': False, 'start': None, 'end': None}
            self.chart_date_states[key] = state
        # 第一次自動套日期
        if (not state['custom']) and raw_df is not None and not raw_df.empty and 'point_time' in raw_df.columns:
            try:
                tmp = raw_df.copy()
                tmp['point_time'] = pd.to_datetime(tmp['point_time'])
                latest = tmp['point_time'].max()
                start_candidate = latest - pd.DateOffset(months=3)
                earliest = tmp['point_time'].min()
                if start_candidate < earliest:
                    start_candidate = earliest
                blocker1 = QtCore.QSignalBlocker(self.start_date)
                blocker2 = QtCore.QSignalBlocker(self.end_date)
                self.end_date.setDate(QtCore.QDate(latest.year, latest.month, latest.day))
                self.start_date.setDate(QtCore.QDate(start_candidate.year, start_candidate.month, start_candidate.day))
                del blocker1, blocker2
                state['start'] = start_candidate.date()
                state['end'] = latest.date()
            except Exception as e:
                print(f"[WARN] 自動日期設定失敗: {e}")
        # 動態計算 + 繪圖
        self._update_current_chart_dynamic(chart_info)

    # === 使用者調整日期事件 ===
    def on_date_changed(self, *_):
        idx = self.chart_combo.currentIndex() - 1
        if idx < 0 or self.all_charts_info is None:
            return
        chart_info = self.all_charts_info.iloc[idx]
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        key = (group_name, chart_name)
        state = self.chart_date_states.get(key)
        if state is None:
            state = {'custom': False, 'start': None, 'end': None}
            self.chart_date_states[key] = state
        state['custom'] = True
        state['start'] = self.start_date.date().toPyDate()
        state['end'] = self.end_date.date().toPyDate()
        self._update_current_chart_dynamic(chart_info)

    # === Cpk 動態計算 ===
    def _compute_cpk_windows(self, raw_df: pd.DataFrame, chart_info: pd.Series, end_time: pd.Timestamp):
        """以 end_time 為基準計算最近三個連續月度窗口的 Cpk。"""
        result = {'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None}
        if raw_df is None or raw_df.empty:
            return result
        if 'point_time' not in raw_df.columns:
            result['Cpk'] = calculate_cpk(raw_df, chart_info)['Cpk']
            return result
        df = raw_df.copy()
        df['point_time'] = pd.to_datetime(df['point_time'])
        df = df[df['point_time'] <= end_time]
        if df.empty:
            return result
        start1 = end_time - pd.DateOffset(months=1)
        start2 = end_time - pd.DateOffset(months=2)
        start3 = end_time - pd.DateOffset(months=3)
        mask1 = (df['point_time'] > start1) & (df['point_time'] <= end_time)
        mask2 = (df['point_time'] > start2) & (df['point_time'] <= start1)
        mask3 = (df['point_time'] > start3) & (df['point_time'] <= start2)
        if mask1.any():
            result['Cpk'] = calculate_cpk(df[mask1], chart_info)['Cpk']
        if mask2.any():
            result['Cpk_last_month'] = calculate_cpk(df[mask2], chart_info)['Cpk']
        if mask3.any():
            result['Cpk_last2_month'] = calculate_cpk(df[mask3], chart_info)['Cpk']
        return result

    def _recompute_cpk_for_chart(self, chart_info: pd.Series, end_date):
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        raw_df = self.raw_charts_dict.get((group_name, chart_name))
        if raw_df is None or raw_df.empty:
            return {'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None}
        if 'point_time' not in raw_df.columns:
            return {'Cpk': calculate_cpk(raw_df, chart_info)['Cpk'], 'Cpk_last_month': None, 'Cpk_last2_month': None}
        raw_df_local = raw_df.copy()
        raw_df_local['point_time'] = pd.to_datetime(raw_df_local['point_time'])
        latest = raw_df_local['point_time'].max()
        end_time = pd.to_datetime(end_date)
        if end_time > latest:
            end_time = latest
        return self._compute_cpk_windows(raw_df_local, chart_info, end_time)

    def _update_current_chart_dynamic(self, chart_info: pd.Series):
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        
        # 判斷是否使用自訂時間模式
        if self.custom_range_btn.isChecked():
            # 自訂時間模式：根據 UI 上的起始和結束時間計算
            start_d = self.start_date.date().toPyDate()
            end_d = self.end_date.date().toPyDate()
            start_time = pd.to_datetime(start_d)
            end_time = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
            
            # 取得原始資料
            raw_df = self.raw_charts_dict.get((group_name, chart_name))
            
            if raw_df is not None and 'point_time' in raw_df.columns:
                print(f"[DEBUG][UI] {group_name}@{chart_name} 自訂Cpk區間: {start_d} ~ {end_d}")
            
            # 使用自訂範圍計算方法
            cpk_res = self._compute_cpk_custom_range(raw_df, chart_info, start_time, end_time)
            
            # 計算全部資料 Cpk
            all_data_cpk = None
            if raw_df is not None and not raw_df.empty:
                all_data_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
            
            # 更新卡片顯示
            def set_card(key, value, is_percent=False):
                comp = self.metric_cards[key]
                if value is None:
                    comp['value_label'].setText('-')
                else:
                    comp['value_label'].setText(f"{value:.1f}%" if is_percent else f"{value:.3f}")
            
            # 計算 K 參數（用自訂範圍內的資料）
            kval = None
            try:
                usl = chart_info.get('USL', None)
                lsl = chart_info.get('LSL', None)
                target = None
                for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                    if key_t in chart_info and pd.notna(chart_info[key_t]):
                        target = chart_info[key_t]
                        break
                
                # 用自訂範圍內的資料計算 mean
                mean_val = None
                if raw_df is not None and not raw_df.empty and 'point_time' in raw_df.columns:
                    filtered_df = raw_df[(pd.to_datetime(raw_df['point_time']) >= start_time) & 
                                        (pd.to_datetime(raw_df['point_time']) <= end_time)]
                    if not filtered_df.empty:
                        mean_val = filtered_df['point_val'].mean()
                elif raw_df is not None and not raw_df.empty:
                    mean_val = raw_df['point_val'].mean()
                    
                rng = (usl - lsl) / 2 if (usl is not None and lsl is not None and (usl-lsl)!=0) else None
                if mean_val is not None and target is not None and rng:
                    kval = abs(mean_val - target) / rng
            except Exception:
                kval = None
            
            # 在自訂模式下，只顯示當月 Cpk，L1 和 L2 設為 "-"
            set_card('kval', kval)
            set_card('cpk', cpk_res.get('Cpk'))
            set_card('l1', None)  # 自訂模式不顯示 L1
            set_card('l2', None)  # 自訂模式不顯示 L2
            set_card('custom', all_data_cpk)
            set_card('r1', None)  # 自訂模式不顯示 R1
            set_card('r2', None)  # 自訂模式不顯示 R2
            
        else:
            # 原本的邏輯：從結束時間回推三個月
            end_d = self.end_date.date().toPyDate()
            raw_df = self.raw_charts_dict.get((group_name, chart_name))
            
            if raw_df is not None and 'point_time' in raw_df.columns:
                raw_df_local = raw_df.copy()
                raw_df_local['point_time'] = pd.to_datetime(raw_df_local['point_time'])
                latest = raw_df_local['point_time'].max()
                start1 = pd.to_datetime(end_d) - pd.DateOffset(months=1)
                print(f"[DEBUG][UI] {group_name}@{chart_name} Cpk區間: {start1.date()} ~ {end_d}")
            
            cpk_res = self._recompute_cpk_for_chart(chart_info, end_d)
            
            # 改為全部資料 Cpk
            all_data_cpk = None
            if raw_df is not None and not raw_df.empty:
                all_data_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
            
            def set_card(key, value, is_percent=False):
                comp = self.metric_cards[key]
                if value is None:
                    comp['value_label'].setText('-')
                else:
                    comp['value_label'].setText(f"{value:.1f}%" if is_percent else f"{value:.3f}")

            # 計算 K 參數
            kval = None
            try:
                usl = chart_info.get('USL', None)
                lsl = chart_info.get('LSL', None)
                target = None
                for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                    if key_t in chart_info and pd.notna(chart_info[key_t]):
                        target = chart_info[key_t]
                        break
                mean_val = None
                raw_df2 = self.raw_charts_dict.get((group_name, chart_name))
                if raw_df2 is not None and not raw_df2.empty:
                    mean_val = raw_df2['point_val'].mean()
                rng = (usl - lsl) / 2 if (usl is not None and lsl is not None and (usl-lsl)!=0) else None
                if mean_val is not None and target is not None and rng:
                    kval = abs(mean_val - target) / rng
            except Exception:
                kval = None
            
            set_card('kval', kval)
            set_card('cpk', cpk_res.get('Cpk'))
            set_card('l1', cpk_res.get('Cpk_last_month'))
            set_card('l2', cpk_res.get('Cpk_last2_month'))
            set_card('custom', all_data_cpk)
            
            cpk = cpk_res.get('Cpk')
            l1 = cpk_res.get('Cpk_last_month')
            l2 = cpk_res.get('Cpk_last2_month')
            
            r1 = r2 = None
            if cpk is not None and l1 is not None and l1 != 0 and cpk <= l1:
                r1 = (1 - (cpk / l1)) * 100
            if cpk is not None and l1 is not None and l2 is not None and l2 != 0 and cpk <= l1 <= l2:
                r2 = (1 - (cpk / l2)) * 100
                
            set_card('r1', r1, is_percent=True)
            set_card('r2', r2, is_percent=True)
        
        # 依目前日期範圍重畫圖
        self.draw_spc_chart(group_name, chart_name, chart_info)

    # === X 軸模式切換 ===
    def toggle_axis_mode(self):
        self.axis_mode = 'time' if self.axis_mode == 'index' else 'index'
        # 更新按鈕文字
        self.axis_mode_btn.setText('Equal Axis' if self.axis_mode == 'time' else 'Time Axis')
        # 重新繪圖（若已選 chart）
        idx = self.chart_combo.currentIndex() - 1
        if idx >= 0 and self.all_charts_info is not None:
            chart_info = self.all_charts_info.iloc[idx]
            group_name = str(chart_info['GroupName'])
            chart_name = str(chart_info['ChartName'])
            self.draw_spc_chart(group_name, chart_name, chart_info)

    def draw_spc_chart(self, group_name: str, chart_name: str, chart_info):
        raw_df = self.raw_charts_dict.get((group_name, chart_name))
        self.figure.clear()
        
        # 創建 GridSpec 布局：左邊一個大圖，右邊上下兩個小圖
        import matplotlib.gridspec as gridspec
        gs = gridspec.GridSpec(2, 2, width_ratios=[3, 1], height_ratios=[1, 1], 
                              hspace=0.3, wspace=0.25)
        
        # 主 SPC 圖（左側，跨兩行）
        ax_main = self.figure.add_subplot(gs[:, 0])
        # Box Plot（右上）
        ax_box = self.figure.add_subplot(gs[0, 1])
        # QQ Plot（右下）
        ax_qq = self.figure.add_subplot(gs[1, 1])
        
        # 標題格式: [GroupName@ChartName@Characteristics]
        characteristics = chart_info.get('Characteristics', '')
        ax_main.set_title(f"{group_name}@{chart_name}@{characteristics}", pad=18, fontsize=12)
        
        # Box Plot 標題根據是否有 EQP_id 欄位動態調整
        if raw_df is not None and 'EQP_id' in raw_df.columns and not raw_df['EQP_id'].isna().all():
            ax_box.set_title("Box Plot (by EQP_id)", fontsize=10)
        else:
            ax_box.set_title("Box Plot", fontsize=10)
            
        ax_qq.set_title("Q-Q Plot", fontsize=10)
        
        if raw_df is None or raw_df.empty:
            ax_main.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_main.transAxes)
            ax_box.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_box.transAxes)
            ax_qq.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_qq.transAxes)
            self.canvas.draw()
            return
        
        plot_df = raw_df.copy()
        
        # 日期過濾 (若有 point_time 欄位)
        if 'point_time' in plot_df.columns:
            try:
                plot_df['point_time'] = pd.to_datetime(plot_df['point_time'])
                start_ts = pd.to_datetime(self.start_date.date().toString('yyyy-MM-dd'))
                end_ts = pd.to_datetime(self.end_date.date().toString('yyyy-MM-dd')) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                filtered = plot_df[(plot_df['point_time'] >= start_ts) & (plot_df['point_time'] <= end_ts)]
                if not filtered.empty:
                    plot_df = filtered
            except Exception:
                pass
                
        if plot_df.empty:
            ax_main.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_main.transAxes)
            ax_box.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_box.transAxes)
            ax_qq.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_qq.transAxes)
            self.canvas.draw()
            return

        # === 繪製主 SPC 圖 ===
        self._draw_main_spc_chart(ax_main, plot_df, chart_info)
        
        # === 繪製 Box Plot ===
        self._draw_box_plot(ax_box, plot_df, chart_info)
        
        # === 繪製 QQ Plot ===
        self._draw_qq_plot(ax_qq, plot_df, chart_info)
        
        # 使用安全的佈局調整
        try:
            self.figure.tight_layout(pad=1.5)
        except:
            # 如果 tight_layout 失敗，使用手動調整
            self.figure.subplots_adjust(left=0.08, right=0.98, top=0.92, bottom=0.12, 
                                      wspace=0.25, hspace=0.35)
        self.canvas.draw()

    def _draw_main_spc_chart(self, ax, plot_df, chart_info):
        """繪製主要的 SPC 控制圖"""
        ax.set_xlabel("" if self.axis_mode == 'index' else "")
   
        # X 軸處理：
        y = plot_df['point_val'].values
        use_time_axis = False
        if self.axis_mode == 'time' and 'point_time' in plot_df.columns:
            try:
                plot_df_sorted = plot_df.sort_values('point_time')
                x = plot_df_sorted['point_time'].values
                y = plot_df_sorted['point_val'].values
                use_time_axis = True
            except Exception:
                x = range(1, len(y) + 1)
        else:
            # 等距模式：保持所有點等距，避免同時間戳疊在一起
            if 'point_time' in plot_df.columns:
                try:
                    plot_df = plot_df.sort_values('point_time').reset_index(drop=True)
                    y = plot_df['point_val'].values
                except Exception:
                    pass
            x = range(1, len(y) + 1)

        # === 在圖上標示「當月/上月/上上月」區間 ===
        if 'point_time' in plot_df.columns and not plot_df.empty:
            try:
                times = pd.to_datetime(plot_df['point_time']).to_numpy()
                tmin, tmax = times.min(), times.max()
                
                # 判斷是否為自訂時間模式
                if self.custom_range_btn.isChecked():
                    # 自訂模式：只顯示一個自訂範圍區塊
                    start_time = pd.to_datetime(self.start_date.date().toString('yyyy-MM-dd'))
                    end_time = pd.to_datetime(self.end_date.date().toString('yyyy-MM-dd')) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                    
                    windows = [
                        (start_time, end_time, 'Custom', '#dbeafe'),
                    ]
                else:
                    # 原本邏輯：從結束時間回推三個月
                    end_sel = pd.to_datetime(self.end_date.date().toString('yyyy-MM-dd')) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                    if end_sel > pd.Timestamp(tmax):
                        end_sel = pd.Timestamp(tmax)
                    start1 = end_sel - pd.DateOffset(months=1)
                    start2 = end_sel - pd.DateOffset(months=2)
                    start3 = end_sel - pd.DateOffset(months=3)
                    windows = [
                        (start1, end_sel,  'L0',   '#dbeafe'),
                        (start2, start1,   'L1',   '#fef9c3'),
                        (start3, start2,   'L2',   '#ede9fe'),
                    ]
                
                text_trans = mtransforms.blended_transform_factory(ax.transData, ax.transAxes)
                if use_time_axis:
                    for s, e, lab, col in windows:
                        s_clip = max(pd.Timestamp(s), pd.Timestamp(tmin))
                        e_clip = min(pd.Timestamp(e), pd.Timestamp(tmax))
                        if e_clip <= s_clip:
                            continue
                        ax.axvspan(s_clip, e_clip, color=col, alpha=0.25, zorder=0)
                        x_center = s_clip + (e_clip - s_clip) / 2
                        ax.text(x_center, 1.04, lab, transform=text_trans, ha='center', va='top', fontsize=8, color='#374151', alpha=0.9)
                else:
                    # 對應到等距索引（x: 1..N），區塊邊界用格線中間: 0.5 ~ N+0.5
                    n = len(times)
                    def t2ix_left(t):
                        return float(np.searchsorted(times, np.datetime64(t), side='left')) + 0.5
                    def t2ix_right(t):
                        return float(np.searchsorted(times, np.datetime64(t), side='right')) + 0.5
                    x_min, x_max = 0.5, n + 0.5
                    for s, e, lab, col in windows:
                        s_clip = max(pd.Timestamp(s), pd.Timestamp(tmin))
                        e_clip = min(pd.Timestamp(e), pd.Timestamp(tmax))
                        if e_clip <= s_clip:
                            continue
                        xl = max(x_min, t2ix_left(s_clip))
                        xr = min(x_max, t2ix_right(e_clip))
                        if xr <= xl:
                            continue
                        ax.axvspan(xl, xr, color=col, alpha=0.25, zorder=0)
                        x_center = (xl + xr) / 2.0
                        ax.text(x_center, 1.04, lab, transform=text_trans, ha='center', va='top', fontsize=8, color='#374151', alpha=0.9)
            except Exception as _:
                pass
                
        # 計算統計線
        usl = chart_info.get('USL', None)
        lsl = chart_info.get('LSL', None)
        target = None
        for key in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
            if key in chart_info and pd.notna(chart_info[key]):
                target = chart_info[key]
                break
        mean_val = float(np.mean(y)) if len(y) else None
        
        # 繪製點與線 (主數據線與超規點不加入 legend)
        ax.plot(x, y, linestyle='-', marker='o', color='#2563eb', markersize=4, linewidth=1.0, label='_nolegend_')
        if usl is not None:
            ax.scatter([xi for xi, yi in zip(x, y) if yi > usl], [yi for yi in y if yi > usl], color='#dc2626', s=25, zorder=5, label='_nolegend_')
        if lsl is not None:
            ax.scatter([xi for xi, yi in zip(x, y) if yi < lsl], [yi for yi in y if yi < lsl], color='#dc2626', marker='s', s=25, zorder=5, label='_nolegend_')
            
        # 計算 y 範圍（納入 USL/LSL/Target/Mean）避免被裁切
        extra_vals = [v for v in [usl, lsl, target, mean_val]
                      if v is not None and not (isinstance(v, float) and np.isnan(v))]
        if len(y) > 0:
            ymin_sel = float(np.min(y))
            ymax_sel = float(np.max(y))
        else:
            ymin_sel, ymax_sel = (0.0, 1.0)
        if extra_vals:
            ymin_sel = min(ymin_sel, min(extra_vals))
            ymax_sel = max(ymax_sel, max(extra_vals))
        rng = ymax_sel - ymin_sel
        margin = 0.05 * rng if rng > 0 else 1.0
        ax.set_ylim(ymin_sel - margin, ymax_sel + margin)

        # 畫短水平線，並讓文字直接接在線的末端
        trans = mtransforms.blended_transform_factory(ax.transAxes, ax.transData)
        def segment_with_label(val, name, color, va='center'):
            if val is None or (isinstance(val, float) and np.isnan(val)):
                return
            x0, x1 = 0.0, 0.96  # 線更長，文字更貼近右邊界
            ax.plot([x0, x1], [val, val], transform=trans, color=color, linestyle='--', linewidth=1.0)
            ax.text(x1, val, name, transform=trans, color=color, va=va, ha='left', fontsize=8)

        segment_with_label(usl, 'USL', '#ef4444', va='center')
        segment_with_label(lsl, 'LSL', '#ef4444', va='center')
        segment_with_label(target, 'Target', '#f59e0b', va='center')
        segment_with_label(mean_val, 'Mean', '#16a34a', va='center')
        
        # 時間軸格式化
        if use_time_axis:
            try:
                locator = mdates.AutoDateLocator(minticks=3, maxticks=6)
                formatter = mdates.ConciseDateFormatter(locator)
                ax.xaxis.set_major_locator(locator)
                ax.xaxis.set_major_formatter(formatter)
                for label in ax.get_xticklabels():
                    label.set_rotation(90)
                    label.set_ha('center')
                    label.set_fontsize(8)
            except Exception:
                pass
        else:
            # 等距模式若有時間欄位，挑選部分刻度顯示對應日期字串
            if 'point_time' in plot_df.columns and not plot_df.empty:
                times = plot_df['point_time'].tolist()
                total = len(times)
                if total <= 8:
                    tick_idx = list(range(1, total + 1))
                else:
                    # 確保最後一點和其他點保持相對等距
                    step = max(1, total // 6)
                    tick_idx = list(range(1, total + 1, step))
                    # 如果最後一點距離倒數第二個標記點太近，移除倒數第二個
                    if len(tick_idx) > 1 and total - tick_idx[-1] < step // 2:
                        tick_idx[-1] = total
                    elif tick_idx[-1] != total:
                        tick_idx.append(total)
                labels = [times[i-1].strftime('%Y-%m-%d') for i in tick_idx]
                ax.set_xticks(tick_idx)
                ax.set_xticklabels(labels, rotation=30, ha='right', fontsize=8)
                
        ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.5)

    def _draw_box_plot(self, ax, plot_df, chart_info):
        """繪製箱型圖，支援按 EQP_id 分組"""
        if plot_df.empty:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes,
                   fontsize=12, color='gray', style='italic')
            return
            
        # 檢查是否有 EQP_id 欄位
        if 'EQP_id' in plot_df.columns and not plot_df['EQP_id'].isna().all():
            # 按 EQP_id 分組繪製多個 box plot
            grouped = plot_df.groupby('EQP_id')
            eqp_ids = list(grouped.groups.keys())
            
            # 準備數據
            box_data = []
            labels = []
            
            for eqp_id in sorted(eqp_ids):
                group_data = grouped.get_group(eqp_id)['point_val'].values
                if len(group_data) > 0:
                    box_data.append(group_data)
                    labels.append(str(eqp_id))
            
            if len(box_data) == 0:
                ax.text(0.5, 0.5, "No Valid Data", ha='center', va='center', transform=ax.transAxes,
                       fontsize=12, color='gray', style='italic')
                return
                
            # 創建多組箱型圖 - 更漂亮的樣式
            box_plot = ax.boxplot(box_data, patch_artist=True, notch=False,
                                 boxprops=dict(facecolor='lightblue', alpha=0.8, linewidth=1.5),
                                 medianprops=dict(color='darkred', linewidth=2.5),
                                 whiskerprops=dict(linewidth=1.5, color='darkblue'),
                                 capprops=dict(linewidth=1.5, color='darkblue'),
                                 flierprops=dict(marker='o', markerfacecolor='red', markeredgecolor='darkred',
                                               markersize=5, alpha=0.8, markeredgewidth=0.5))
            
            # 設定不同顏色 - 更優雅的配色方案
            colors = ['#87CEEB', '#98FB98', '#FFB6C1', '#F0E68C', '#DDA0DD', '#F5DEB3', '#B0E0E6']
            for i, patch in enumerate(box_plot['boxes']):
                patch.set_facecolor(colors[i % len(colors)])
                patch.set_edgecolor('darkblue')
                patch.set_alpha(0.8)
                
            # 只顯示機台名稱，不要下方的 X 軸標籤
            ax.set_xticklabels(labels, rotation=0, ha='right', fontsize=9, 
                              fontweight='bold', color='darkblue')
            # 移除 X 軸標題
            ax.set_xlabel('')
            
        else:
            # 原始邏輯：單一 box plot - 更漂亮的樣式
            y = plot_df['point_val'].values
            if len(y) == 0:
                ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes,
                       fontsize=12, color='gray', style='italic')
                return
                
            # 創建箱型圖 - 單一但精美的樣式
            box_plot = ax.boxplot(y, patch_artist=True, notch=False,
                                 boxprops=dict(facecolor='#87CEEB', alpha=0.8, linewidth=2,
                                              edgecolor='darkblue'),
                                 medianprops=dict(color='darkred', linewidth=3),
                                 whiskerprops=dict(linewidth=2, color='darkblue'),
                                 capprops=dict(linewidth=2, color='darkblue'),
                                 flierprops=dict(marker='o', markerfacecolor='red', markeredgecolor='darkred',
                                               markersize=6, alpha=0.8, markeredgewidth=1))
            
            # 沒有 EQP_id 時，隱藏 X 軸刻度和標籤
            ax.set_xticks([])
            ax.set_xlabel('')
        
        # 設定 y 軸範圍 - 移除 USL/LSL 相關程式碼
        all_values = plot_df['point_val'].values
        if len(all_values) > 0:
            ymin_sel = float(np.min(all_values))
            ymax_sel = float(np.max(all_values))
            rng = ymax_sel - ymin_sel
            margin = 0.1 * rng if rng > 0 else 1.0  # 增加邊距讓圖更好看
            ax.set_ylim(ymin_sel - margin, ymax_sel + margin)
        
        # 美化網格和軸
        # ax.grid(True, linestyle='--', linewidth=0.8, alpha=0.6, color='lightgray')
        # ax.set_facecolor('#FAFAFA')  # 淺灰背景
        # ax.tick_params(axis='both', which='major', labelsize=9, colors='darkblue')
        # ax.set_ylabel('Values', fontsize=10, fontweight='bold', color='darkblue')
        
        # # 加入邊框美化
        # for spine in ax.spines.values():
        #     spine.set_edgecolor('darkblue')
        #     spine.set_linewidth(1.2)

    def _draw_qq_plot(self, ax, plot_df, chart_info):
        """繪製 Q-Q 圖（與常態分布比較）"""
        y = plot_df['point_val'].values
        if len(y) == 0:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
            return
            
        # 計算 Q-Q plot 資料
        try:
            # 使用 scipy.stats.probplot 生成 Q-Q plot
            (osm, osr), (slope, intercept, r) = stats.probplot(y, dist="norm", plot=None)
            
            # 繪製資料點
            ax.scatter(osm, osr, alpha=0.7, color='blue', s=20)
            
            # 繪製理論線（完美常態分布線）
            line_x = np.array([osm.min(), osm.max()])
            line_y = slope * line_x + intercept
            ax.plot(line_x, line_y, 'r-', linewidth=1.5, alpha=0.8, label=f'R²={r**2:.3f}')
            
            ax.set_xlabel('Theoretical Quantiles', fontsize=8)
            ax.set_ylabel('Sample Quantiles', fontsize=8)
            ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.3)
            ax.tick_params(axis='both', which='major', labelsize=8)
            
            # 顯示 R² 值
            ax.legend(fontsize=7, loc='lower right')
            
        except Exception as e:
            ax.text(0.5, 0.5, f"Calculation Error:\n{str(e)}", ha='center', va='center', 
                   transform=ax.transAxes, fontsize=8)

    def prev_chart(self):
        """切換到上一張圖表"""
        current_idx = self.chart_combo.currentIndex()
        if current_idx > 1:  # 第0項是"請選擇Chart"
            self.chart_combo.setCurrentIndex(current_idx - 1)
    
    def next_chart(self):
        """切換到下一張圖表"""
        current_idx = self.chart_combo.currentIndex()
        if current_idx < self.chart_combo.count() - 1:
            self.chart_combo.setCurrentIndex(current_idx + 1)

    def on_custom_range_toggle(self):
        """切換自訂時間模式"""
        if self.custom_range_btn.isChecked():
            self.custom_range_btn.setText("✓ Custom Time Mode")
        else:
            self.custom_range_btn.setText("Custom Time Mode")
        
        # 重新計算當前選中的圖表
        idx = self.chart_combo.currentIndex() - 1
        if idx >= 0 and self.all_charts_info is not None:
            chart_info = self.all_charts_info.iloc[idx]
            self._update_current_chart_dynamic(chart_info)

    def _compute_cpk_custom_range(self, raw_df: pd.DataFrame, chart_info: pd.Series, start_time: pd.Timestamp, end_time: pd.Timestamp):
        """根據自訂的起始和結束時間計算 Cpk（只算當月）"""
        result = {'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None}
        
        if raw_df is None or raw_df.empty:
            return result
        
        if 'point_time' not in raw_df.columns:
            result['Cpk'] = calculate_cpk(raw_df, chart_info)['Cpk']
            return result
        
        df = raw_df.copy()
        df['point_time'] = pd.to_datetime(df['point_time'])
        
        # 篩選自訂範圍的資料
        filtered_df = df[(df['point_time'] >= start_time) & (df['point_time'] <= end_time)]
        
        if filtered_df.empty:
            return result
        
        # 只計算自訂範圍的 Cpk（當月）
        result['Cpk'] = calculate_cpk(filtered_df, chart_info)['Cpk']
        
        return result

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    dashboard = SPCCpkDashboard()
    dashboard.show()
    sys.exit(app.exec())