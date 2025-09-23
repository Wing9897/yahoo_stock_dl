import sys
import os
import sqlite3
import threading
from datetime import datetime, timedelta

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QGridLayout, QLabel, QTextEdit,
                            QComboBox, QPushButton, QProgressBar, QGroupBox,
                            QDateEdit, QFileDialog, QMessageBox, QRadioButton,
                            QButtonGroup, QSplitter, QFrame, QScrollArea,
                            QTabWidget, QTableWidget, QTableWidgetItem,
                            QMenuBar, QMenu, QSizePolicy)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QTimer, QSettings
from PyQt6.QtGui import QFont, QIcon, QPixmap, QAction

import yfinance as yf
import pandas as pd

try:
    import darkdetect
    DARKDETECT_AVAILABLE = True
except ImportError:
    DARKDETECT_AVAILABLE = False


class DownloadWorker(QThread):
    progress_updated = pyqtSignal(int, int)  # current, total
    log_message = pyqtSignal(str)
    status_updated = pyqtSignal(str)
    finished = pyqtSignal(bool)  # success
    download_result = pyqtSignal(str, bool, str)  # symbol, success, message

    def __init__(self, symbols, market_info, interval, start_date, end_date,
                 output_format, output_dir):
        super().__init__()
        self.symbols = symbols
        self.market_info = market_info
        self.interval = interval
        self.start_date = start_date
        self.end_date = end_date
        self.output_format = output_format
        self.output_dir = output_dir
        self.stop_flag = False

    def stop(self):
        self.stop_flag = True

    def run(self):
        try:
            # 添加市場後綴
            suffix = self.market_info['suffix']
            if suffix:
                processed_symbols = []
                for s in self.symbols:
                    if not any(s.endswith(suf) for suf in ['.TW', '.HK', '.T', '.DE', '.L']):
                        processed_symbols.append(s + suffix)
                    else:
                        processed_symbols.append(s)
            else:
                processed_symbols = self.symbols

            total_stocks = len(processed_symbols)

            # 準備SQLite連接（如果需要）
            db_connection = None
            if self.output_format == "SQLite":
                db_path = os.path.join(self.output_dir, "stock_data.db")
                db_connection = sqlite3.connect(db_path)
                self.log_message.emit(f"SQLite數據庫: {db_path}")

            # 下載每隻股票
            for i, symbol in enumerate(processed_symbols):
                if self.stop_flag:
                    break

                self.status_updated.emit(f"正在下載 {symbol} ({i+1}/{total_stocks})")
                self.log_message.emit(f"開始下載 {symbol}")

                try:
                    # 下載數據
                    ticker = yf.Ticker(symbol)
                    data = ticker.history(
                        start=self.start_date,
                        end=self.end_date,
                        interval=self.interval
                    )

                    if data.empty:
                        self.log_message.emit(f"警告: {symbol} 沒有數據")
                        self.download_result.emit(symbol, False, "沒有數據")
                        continue

                    # 保存數據
                    self.save_data(symbol, data, db_connection)
                    self.log_message.emit(f"完成 {symbol} - {len(data)} 條記錄")
                    self.download_result.emit(symbol, True, f"{len(data)} 條記錄")

                except Exception as e:
                    error_msg = str(e)
                    self.log_message.emit(f"錯誤 {symbol}: {error_msg}")
                    self.download_result.emit(symbol, False, error_msg)

                # 更新進度
                self.progress_updated.emit(i + 1, total_stocks)

            # 清理
            if db_connection:
                db_connection.close()

            if not self.stop_flag:
                self.status_updated.emit("下載完成")
                self.log_message.emit("所有股票數據下載完成")
                self.finished.emit(True)
            else:
                self.status_updated.emit("下載已停止")
                self.log_message.emit("下載已被用戶停止")
                self.finished.emit(False)

        except Exception as e:
            self.log_message.emit(f"下載過程發生錯誤: {str(e)}")
            self.finished.emit(False)

    def save_data(self, symbol, data, db_connection=None):
        # 清理檔案名稱中的特殊字符
        safe_symbol = symbol.replace('.', '_').replace(':', '_')

        if self.output_format == "CSV":
            filename = os.path.join(self.output_dir, f"{safe_symbol}_data.csv")
            data.to_csv(filename)

        elif self.output_format == "Excel":
            filename = os.path.join(self.output_dir, f"{safe_symbol}_data.xlsx")
            data.to_excel(filename)

        elif self.output_format == "SQLite" and db_connection:
            table_name = f"stock_{safe_symbol}"
            data.to_sql(table_name, db_connection, if_exists='replace', index=True)


class EnhancedStockGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("股票數據爬取器 - 增強版")

        # 設置最小窗口大小，但允許完全縮放
        self.setMinimumSize(750, 500)
        self.setGeometry(100, 100, 900, 700)  # 減小默認窗口大小

        # 設置窗口可以完全縮放
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # 設置和主題管理
        self.settings = QSettings("StockDataApp", "EnhancedGUI")
        self.current_theme = "auto"  # auto, light, dark

        # 股票國家/市場映射
        self.markets = {
            "美國": {"suffix": "", "currency": "USD", "examples": "AAPL, MSFT, GOOGL"},
            "台灣": {"suffix": ".TW", "currency": "TWD", "examples": "2330.TW, 2317.TW"},
            "香港": {"suffix": ".HK", "currency": "HKD", "examples": "0700.HK, 0005.HK"},
            "日本": {"suffix": ".T", "currency": "JPY", "examples": "7203.T, 6758.T"},
            "德國": {"suffix": ".DE", "currency": "EUR", "examples": "SAP.DE, VOW3.DE"},
            "英國": {"suffix": ".L", "currency": "GBP", "examples": "BARC.L, LLOY.L"}
        }

        self.download_worker = None
        self.download_results = {"success": [], "failed": [], "total": 0}
        self.stop_flag = False
        self.validation_timer = QTimer()
        self.validation_timer.setSingleShot(True)
        self.validation_timer.timeout.connect(self.perform_validation)

        # 設置菜單和UI
        self.setup_menu()
        self.setup_ui()

        # 載入保存的設置
        self.load_settings()

        # 應用主題
        self.apply_theme()

    def setup_menu(self):
        menubar = self.menuBar()

        # 視圖菜單
        view_menu = menubar.addMenu('視圖(&V)')

        # 主題子菜單
        theme_menu = view_menu.addMenu('主題')

        # 自動主題（跟隨系統）
        self.auto_action = QAction('自動（跟隨系統）', self)
        self.auto_action.setCheckable(True)
        self.auto_action.triggered.connect(lambda: self.set_theme("auto"))
        theme_menu.addAction(self.auto_action)

        # 淺色主題
        self.light_action = QAction('淺色主題', self)
        self.light_action.setCheckable(True)
        self.light_action.triggered.connect(lambda: self.set_theme("light"))
        theme_menu.addAction(self.light_action)

        # 深色主題
        self.dark_action = QAction('深色主題', self)
        self.dark_action.setCheckable(True)
        self.dark_action.triggered.connect(lambda: self.set_theme("dark"))
        theme_menu.addAction(self.dark_action)

        # 設置初始選中
        self.auto_action.setChecked(True)

        # 工具菜單
        tools_menu = menubar.addMenu('工具(&T)')

        refresh_action = QAction('重新整理', self)
        refresh_action.triggered.connect(self.refresh_data)
        tools_menu.addAction(refresh_action)

        tools_menu.addSeparator()

        settings_action = QAction('設定...', self)
        settings_action.triggered.connect(self.show_settings)
        tools_menu.addAction(settings_action)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局 - 使用可縮放的布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 使用分割器創建可調整的上下布局
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        main_layout.addWidget(splitter)

        # 上半部分：設置區域
        settings_widget = QWidget()
        settings_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        splitter.addWidget(settings_widget)

        # 下半部分：日誌和進度
        log_widget = QWidget()
        log_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        splitter.addWidget(log_widget)

        # 設置初始比例，設置區域更緊湊
        splitter.setSizes([400, 300])
        splitter.setStretchFactor(0, 0)  # 設置區域固定大小
        splitter.setStretchFactor(1, 1)  # 日誌區域可伸縮

        self.setup_settings_area(settings_widget)
        self.setup_log_area(log_widget)

    def setup_settings_area(self, parent):
        layout = QVBoxLayout(parent)
        layout.setSpacing(15)

        # 使用Tab來組織界面
        tab_widget = QTabWidget()
        tab_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(tab_widget)

        # 基本設置Tab
        basic_tab = QWidget()
        basic_tab.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        tab_widget.addTab(basic_tab, "基本設置")
        self.setup_basic_tab(basic_tab)

        # 高級設置Tab
        advanced_tab = QWidget()
        advanced_tab.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        tab_widget.addTab(advanced_tab, "高級設置")
        self.setup_advanced_tab(advanced_tab)

        # 控制按鈕
        self.setup_control_buttons(layout)

    def setup_basic_tab(self, parent):
        layout = QVBoxLayout(parent)
        layout.setSpacing(10)  # 減少間距

        # 股票輸入區域（並列布局）
        stock_group = QGroupBox("股票代號設置")
        stock_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        stock_main_layout = QHBoxLayout(stock_group)
        stock_main_layout.setSpacing(15)

        # 左側：股票代號輸入（佔據更大空間）
        stock_input_area = QVBoxLayout()

        # 股票代號輸入
        stock_label = QLabel("股票代號:")
        stock_label.setMinimumWidth(70)
        self.stock_text = QTextEdit()
        self.stock_text.setMaximumHeight(80)  # 增加高度給更大的textarea
        self.stock_text.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.stock_text.setPlaceholderText("輸入股票代號，多個用逗號分隔\n例如: AAPL, MSFT, GOOGL\n      2330.TW, 2317.TW")
        self.stock_text.textChanged.connect(self.validate_stock_symbols)

        stock_input_layout = QHBoxLayout()
        stock_input_layout.addWidget(stock_label)
        stock_input_layout.addWidget(self.stock_text, 1)
        stock_input_area.addLayout(stock_input_layout)

        # 股票驗證狀態
        self.validation_status = QLabel("✅ 準備輸入股票代號")
        self.validation_status.setStyleSheet("color: #4CAF50; font-size: 10px; padding: 2px; font-weight: bold;")
        self.validation_status.setWordWrap(True)
        stock_input_area.addWidget(self.validation_status)

        stock_main_layout.addLayout(stock_input_area, 2)  # 給股票輸入區域更多空間

        # 右側：設置控制區域
        settings_area = QVBoxLayout()

        # 市場選擇和間隔
        market_interval_layout = QVBoxLayout()

        # 市場選擇
        market_row = QHBoxLayout()
        market_label = QLabel("市場:")
        market_label.setMinimumWidth(40)
        self.market_combo = QComboBox()
        self.market_combo.addItems(list(self.markets.keys()))
        self.market_combo.setCurrentText("美國")
        self.market_combo.currentTextChanged.connect(self.update_market_example)
        self.market_combo.setMinimumWidth(100)
        market_row.addWidget(market_label)
        market_row.addWidget(self.market_combo)
        market_row.addStretch()
        market_interval_layout.addLayout(market_row)

        # 時間間隔
        interval_row = QHBoxLayout()
        interval_label = QLabel("間隔:")
        interval_label.setMinimumWidth(40)
        self.interval_combo = QComboBox()
        intervals = ["1m", "2m", "5m", "15m", "30m", "60m", "1h", "1d", "1wk", "1mo"]
        self.interval_combo.addItems(intervals)
        self.interval_combo.setCurrentText("1d")
        self.interval_combo.setMinimumWidth(100)
        self.interval_combo.currentTextChanged.connect(self.update_interval_warning)
        interval_row.addWidget(interval_label)
        interval_row.addWidget(self.interval_combo)
        interval_row.addStretch()
        market_interval_layout.addLayout(interval_row)

        settings_area.addLayout(market_interval_layout)

        # 市場示例
        self.market_example = QLabel(f"範例: {self.markets['美國']['examples']}")
        self.market_example.setStyleSheet("font-style: italic; padding: 2px; font-size: 10px;")
        self.market_example.setWordWrap(True)
        settings_area.addWidget(self.market_example)

        # yfinance限制警告
        self.interval_warning = QLabel("💡 提醒: 1分鐘數據僅限7天，分鐘級數據僅限60天，小時數據僅限730天")
        self.interval_warning.setStyleSheet("color: #FF9800; font-size: 10px; padding: 2px; font-weight: bold;")
        self.interval_warning.setWordWrap(True)
        settings_area.addWidget(self.interval_warning)

        stock_main_layout.addLayout(settings_area, 1)  # 設置區域佔較少空間

        layout.addWidget(stock_group)

        # 時間設置和輸出設置在同一行
        settings_row_layout = QHBoxLayout()
        settings_row_layout.setSpacing(15)

        # 時間設置區域
        time_group = QGroupBox("日期範圍")
        time_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        time_layout = QGridLayout(time_group)
        time_layout.setSpacing(8)

        # 開始日期
        start_label = QLabel("開始:")
        start_label.setMinimumWidth(40)
        time_layout.addWidget(start_label, 0, 0)

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addYears(-1))
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setMaximumWidth(120)
        self.start_date.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        time_layout.addWidget(self.start_date, 0, 1)

        # 結束日期
        end_label = QLabel("結束:")
        end_label.setMinimumWidth(40)
        time_layout.addWidget(end_label, 1, 0)

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setMaximumWidth(120)
        self.end_date.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        time_layout.addWidget(self.end_date, 1, 1)

        settings_row_layout.addWidget(time_group)

        # 輸出設置區域
        output_group = QGroupBox("輸出設置")
        output_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        output_layout = QVBoxLayout(output_group)
        output_layout.setSpacing(8)

        # 輸出格式選擇
        format_label = QLabel("格式:")
        format_label.setMinimumWidth(40)

        self.format_group = QButtonGroup()
        self.csv_radio = QRadioButton("CSV")
        self.excel_radio = QRadioButton("Excel")
        self.sqlite_radio = QRadioButton("SQLite")
        self.csv_radio.setChecked(True)

        self.format_group.addButton(self.csv_radio, 0)
        self.format_group.addButton(self.excel_radio, 1)
        self.format_group.addButton(self.sqlite_radio, 2)

        format_layout = QHBoxLayout()
        format_layout.addWidget(format_label)
        format_layout.addWidget(self.csv_radio)
        format_layout.addWidget(self.excel_radio)
        format_layout.addWidget(self.sqlite_radio)
        format_layout.addStretch()
        output_layout.addLayout(format_layout)

        # 輸出目錄選擇
        dir_label = QLabel("目錄:")
        dir_label.setMinimumWidth(40)
        self.output_dir_label = QLabel(os.getcwd())
        self.output_dir_label.setStyleSheet("border: 1px solid; padding: 6px; border-radius: 4px;")
        self.output_dir_label.setMinimumHeight(25)
        self.output_dir_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        browse_btn = QPushButton("瀏覽")
        browse_btn.setMaximumWidth(60)
        browse_btn.clicked.connect(self.browse_directory)

        dir_layout = QHBoxLayout()
        dir_layout.addWidget(dir_label)
        dir_layout.addWidget(self.output_dir_label, 1)
        dir_layout.addWidget(browse_btn)
        output_layout.addLayout(dir_layout)

        settings_row_layout.addWidget(output_group)
        layout.addLayout(settings_row_layout)

        layout.addStretch()

    def setup_advanced_tab(self, parent):
        layout = QVBoxLayout(parent)

        # 高級選項區域
        advanced_group = QGroupBox("高級選項")
        advanced_layout = QVBoxLayout(advanced_group)

        # 主題快速切換
        theme_quick_group = QGroupBox("主題快速切換")
        theme_quick_layout = QHBoxLayout(theme_quick_group)

        auto_theme_btn = QPushButton("🌓 自動")
        light_theme_btn = QPushButton("☀️ 淺色")
        dark_theme_btn = QPushButton("🌙 深色")

        auto_theme_btn.clicked.connect(lambda: self.set_theme("auto"))
        light_theme_btn.clicked.connect(lambda: self.set_theme("light"))
        dark_theme_btn.clicked.connect(lambda: self.set_theme("dark"))

        theme_quick_layout.addWidget(auto_theme_btn)
        theme_quick_layout.addWidget(light_theme_btn)
        theme_quick_layout.addWidget(dark_theme_btn)
        theme_quick_layout.addStretch()

        advanced_layout.addWidget(theme_quick_group)

        # yfinance數據限制說明
        yf_limits_group = QGroupBox("yfinance 數據限制說明")
        yf_limits_layout = QVBoxLayout(yf_limits_group)

        limits_text = QLabel(
            "📊 時間間隔限制：\n"
            "• 1分鐘數據(1m)：僅限最近7天\n"
            "• 分鐘級數據(2m~30m)：僅限最近60天\n"
            "• 小時數據(1h)：僅限最近730天(約2年)\n"
            "• 日線及以上：可獲取長期歷史數據\n\n"
            "⚠️ 重要提醒：\n"
            "• 數據成功率約98%（偶有失敗）\n"
            "• 僅限個人研究和教育用途\n"
            "• 數據可能延遲或缺失\n"
            "• 不建議用於實際交易決策"
        )
        limits_text.setStyleSheet("padding: 10px; font-size: 11px; line-height: 1.3;")
        limits_text.setWordWrap(True)
        yf_limits_layout.addWidget(limits_text)

        advanced_layout.addWidget(yf_limits_group)

        # 其他高級選項
        info_label = QLabel("其他功能：\n"
                           "• 智能日期範圍驗證\n"
                           "• 自動數據格式檢測\n"
                           "• 完全可縮放界面\n"
                           "• 系統主題跟隨")
        info_label.setStyleSheet("padding: 20px;")
        advanced_layout.addWidget(info_label)

        layout.addWidget(advanced_group)
        layout.addStretch()

    def setup_control_buttons(self, parent_layout):
        # 控制按鈕區域
        button_layout = QHBoxLayout()

        self.download_btn = QPushButton("🚀 開始下載")
        self.download_btn.setMinimumHeight(40)
        self.download_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.download_btn.clicked.connect(self.start_download)

        self.stop_btn = QPushButton("⏹️ 停止")
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_download)

        self.retry_btn = QPushButton("🔄 重試失敗")
        self.retry_btn.setMinimumHeight(40)
        self.retry_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.retry_btn.setEnabled(False)
        self.retry_btn.clicked.connect(self.retry_failed)

        clear_btn = QPushButton("🧹 清空")
        clear_btn.setMinimumHeight(40)
        clear_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        clear_btn.clicked.connect(self.clear_form)

        button_layout.addWidget(self.download_btn)
        button_layout.addWidget(self.stop_btn)
        button_layout.addWidget(self.retry_btn)
        button_layout.addWidget(clear_btn)

        parent_layout.addLayout(button_layout)

    def setup_log_area(self, parent):
        layout = QVBoxLayout(parent)

        # 進度區域
        progress_group = QGroupBox("下載進度")
        progress_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        progress_layout = QVBoxLayout(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        progress_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("準備就緒")
        self.status_label.setStyleSheet("font-weight: bold;")
        progress_layout.addWidget(self.status_label)

        layout.addWidget(progress_group)

        # 日誌區域
        log_group = QGroupBox("下載日誌")
        log_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        log_layout = QVBoxLayout(log_group)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        log_layout.addWidget(self.log_text)

        layout.addWidget(log_group)

    def get_light_theme(self):
        return """
            QMainWindow {
                background-color: #f0f0f0;
                color: #000000;
            }
            QWidget {
                background-color: #f0f0f0;
                color: #000000;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                border: 2px solid #999999;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: #ffffff;
                color: #000000;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                color: #1976D2;
                font-size: 13px;
                font-weight: bold;
                background-color: #f0f0f0;
            }
            QLabel {
                color: #000000;
                font-size: 12px;
                font-weight: normal;
                background-color: transparent;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 11px;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
            QComboBox, QDateEdit {
                border: 2px solid #999999;
                border-radius: 4px;
                padding: 6px;
                background-color: #ffffff;
                color: #000000;
                font-size: 12px;
                min-height: 20px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
                background-color: #e0e0e0;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #999999;
            }
            QTextEdit {
                border: 2px solid #999999;
                border-radius: 4px;
                background-color: #ffffff;
                color: #000000;
                font-size: 12px;
                padding: 5px;
            }
            QRadioButton {
                font-size: 12px;
                spacing: 5px;
                color: #000000;
                background-color: transparent;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QProgressBar {
                border: 2px solid #999999;
                border-radius: 4px;
                text-align: center;
                font-size: 12px;
                background-color: #ffffff;
                color: #000000;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
            QTabWidget::pane {
                border: 2px solid #999999;
                background-color: #ffffff;
                color: #000000;
                top: -1px;
            }
            QTabBar::tab {
                background-color: #e0e0e0;
                color: #000000;
                padding: 8px 16px;
                margin-right: 2px;
                border: 1px solid #999999;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-size: 12px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #000000;
                border-bottom: 2px solid #2196F3;
            }
            QTabBar::tab:hover {
                background-color: #d0d0d0;
                color: #000000;
            }
            QMenuBar {
                background-color: #f0f0f0;
                color: #000000;
            }
            QMenuBar::item:selected {
                background-color: #d0d0d0;
            }
            QMenu {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #999999;
            }
            QMenu::item:selected {
                background-color: #e3f2fd;
            }
        """

    def get_dark_theme(self):
        return """
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                border: 2px solid #555555;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                color: #64B5F6;
                font-size: 13px;
                font-weight: bold;
                background-color: #2b2b2b;
            }
            QLabel {
                color: #ffffff;
                font-size: 12px;
                font-weight: normal;
                background-color: transparent;
            }
            QPushButton {
                background-color: #1976D2;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 11px;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #1565C0;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QPushButton:disabled {
                background-color: #555555;
            }
            QComboBox, QDateEdit {
                border: 2px solid #555555;
                border-radius: 4px;
                padding: 6px;
                background-color: #3c3c3c;
                color: #ffffff;
                font-size: 12px;
                min-height: 20px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
                background-color: #555555;
            }
            QComboBox QAbstractItemView {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555555;
            }
            QTextEdit {
                border: 2px solid #555555;
                border-radius: 4px;
                background-color: #3c3c3c;
                color: #ffffff;
                font-size: 12px;
                padding: 5px;
            }
            QRadioButton {
                font-size: 12px;
                spacing: 5px;
                color: #ffffff;
                background-color: transparent;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QProgressBar {
                border: 2px solid #555555;
                border-radius: 4px;
                text-align: center;
                font-size: 12px;
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QProgressBar::chunk {
                background-color: #66BB6A;
                border-radius: 3px;
            }
            QTabWidget::pane {
                border: 2px solid #555555;
                background-color: #3c3c3c;
                color: #ffffff;
                top: -1px;
            }
            QTabBar::tab {
                background-color: #555555;
                color: #ffffff;
                padding: 8px 16px;
                margin-right: 2px;
                border: 1px solid #555555;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-size: 12px;
            }
            QTabBar::tab:selected {
                background-color: #3c3c3c;
                color: #ffffff;
                border-bottom: 2px solid #64B5F6;
            }
            QTabBar::tab:hover {
                background-color: #666666;
                color: #ffffff;
            }
            QMenuBar {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QMenuBar::item:selected {
                background-color: #555555;
            }
            QMenu {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555555;
            }
            QMenu::item:selected {
                background-color: #1976D2;
            }
        """

    def apply_theme(self):
        if self.current_theme == "auto":
            # 檢測系統主題
            if DARKDETECT_AVAILABLE and darkdetect.isDark():
                theme_style = self.get_dark_theme()
                self.status_label.setStyleSheet("font-weight: bold; color: #64B5F6;")
            else:
                theme_style = self.get_light_theme()
                self.status_label.setStyleSheet("font-weight: bold; color: #2196F3;")
        elif self.current_theme == "dark":
            theme_style = self.get_dark_theme()
            self.status_label.setStyleSheet("font-weight: bold; color: #64B5F6;")
        else:  # light
            theme_style = self.get_light_theme()
            self.status_label.setStyleSheet("font-weight: bold; color: #2196F3;")

        self.setStyleSheet(theme_style)

    def set_theme(self, theme):
        self.current_theme = theme

        # 手動管理菜單項的選中狀態
        self.auto_action.setChecked(theme == "auto")
        self.light_action.setChecked(theme == "light")
        self.dark_action.setChecked(theme == "dark")

        self.apply_theme()
        self.save_settings()

    def load_settings(self):
        # 載入窗口大小和位置
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)

        # 載入主題設置
        self.current_theme = self.settings.value("theme", "auto")

        # 設置菜單選中狀態
        self.auto_action.setChecked(self.current_theme == "auto")
        self.light_action.setChecked(self.current_theme == "light")
        self.dark_action.setChecked(self.current_theme == "dark")

        # 載入其他設置
        output_dir = self.settings.value("output_dir", os.getcwd())
        self.output_dir_label.setText(output_dir)

    def save_settings(self):
        # 保存窗口大小和位置
        self.settings.setValue("geometry", self.saveGeometry())

        # 保存主題設置
        self.settings.setValue("theme", self.current_theme)

        # 保存其他設置
        self.settings.setValue("output_dir", self.output_dir_label.text())

    def closeEvent(self, event):
        # 應用關閉時保存設置
        self.save_settings()
        event.accept()

    def refresh_data(self):
        # 重新整理數據（重新檢測系統主題等）
        if self.current_theme == "auto":
            self.apply_theme()

    def show_settings(self):
        # 顯示設置對話框（可以在此添加更多設置選項）
        QMessageBox.information(self, "設定", "設定功能將在後續版本中添加")

    def update_market_example(self, market):
        example = self.markets[market]['examples']
        self.market_example.setText(f"範例: {example}")

    def update_interval_warning(self, interval):
        """根據選中的時間間隔更新警告信息"""
        warnings = {
            "1m": "⚠️ 1分鐘數據：僅限最近7天",
            "2m": "⚠️ 2分鐘數據：僅限最近60天",
            "5m": "⚠️ 5分鐘數據：僅限最近60天",
            "15m": "⚠️ 15分鐘數據：僅限最近60天",
            "30m": "⚠️ 30分鐘數據：僅限最近60天",
            "60m": "⚠️ 1小時數據：僅限最近730天(約2年)",
            "1h": "⚠️ 1小時數據：僅限最近730天(約2年)",
            "1d": "✅ 日線數據：可獲取長期歷史數據",
            "1wk": "✅ 週線數據：可獲取長期歷史數據",
            "1mo": "✅ 月線數據：可獲取長期歷史數據"
        }

        warning_text = warnings.get(interval, "💡 請檢查時間間隔設置")
        # 小時數據用橙色表示中等限制，其他用紅色表示嚴格限制
        if interval in ["60m", "1h"]:
            color = "#FF9800"  # 橙色 - 中等限制
        elif interval in ["1m", "2m", "5m", "15m", "30m"]:
            color = "#F44336"  # 紅色 - 嚴格限制
        else:
            color = "#4CAF50"  # 綠色 - 無限制

        self.interval_warning.setText(warning_text)
        self.interval_warning.setStyleSheet(f"color: {color}; font-size: 10px; padding: 2px; font-weight: bold;")

    def validate_date_range(self, interval, start_date, end_date):
        """驗證日期範圍是否符合yfinance限制"""
        from datetime import datetime, timedelta

        today = datetime.now().date()
        days_diff = (end_date - start_date).days
        days_from_today = (today - end_date).days

        # 檢查不同間隔的限制
        if interval == "1m":
            if days_diff > 7:
                return False, "1分鐘數據的日期範圍不能超過7天"
            if days_from_today > 7:
                return False, "1分鐘數據只能獲取最近7天的數據"

        elif interval in ["2m", "5m", "15m", "30m"]:
            if days_diff > 60:
                return False, f"{interval}數據的日期範圍不能超過60天"
            if days_from_today > 60:
                return False, f"{interval}數據只能獲取最近60天的數據"

        elif interval in ["60m", "1h"]:
            if days_diff > 730:
                return False, f"1小時數據的日期範圍不能超過730天(約2年)"
            if days_from_today > 730:
                return False, f"1小時數據只能獲取最近730天(約2年)的數據"

        return True, ""

    def validate_stock_symbols(self):
        """延遲驗證股票代號以避免頻繁請求"""
        self.validation_timer.stop()
        self.validation_timer.start(1000)  # 1秒後執行驗證

    def perform_validation(self):
        """執行股票代號驗證"""
        stocks_text = self.stock_text.toPlainText().strip()
        if not stocks_text:
            self.validation_status.setText("✅ 準備輸入股票代號")
            self.validation_status.setStyleSheet("color: #4CAF50; font-size: 10px; padding: 2px; font-weight: bold;")
            return

        # 解析股票代號
        try:
            symbols = [s.strip() for s in stocks_text.replace('\n', ',').split(',') if s.strip()]
            if not symbols:
                self.validation_status.setText("❌ 請輸入有效的股票代號")
                self.validation_status.setStyleSheet("color: #F44336; font-size: 10px; padding: 2px; font-weight: bold;")
                return

            # 添加市場後綴
            market = self.market_combo.currentText()
            suffix = self.markets[market]['suffix']
            if suffix:
                processed_symbols = []
                for s in symbols:
                    if not any(s.endswith(suf) for suf in ['.TW', '.HK', '.T', '.DE', '.L']):
                        processed_symbols.append(s + suffix)
                    else:
                        processed_symbols.append(s)
            else:
                processed_symbols = symbols

            # 快速驗證前3個股票代號（避免過多請求）
            validation_symbols = processed_symbols[:3]
            valid_count = 0
            total_symbols = len(processed_symbols)

            for symbol in validation_symbols:
                try:
                    ticker = yf.Ticker(symbol)
                    info = ticker.info
                    if info and 'symbol' in info:
                        valid_count += 1
                except:
                    pass

            if len(validation_symbols) == valid_count:
                if total_symbols <= 3:
                    self.validation_status.setText(f"✅ 已驗證 {total_symbols} 個股票代號")
                    self.validation_status.setStyleSheet("color: #4CAF50; font-size: 10px; padding: 2px; font-weight: bold;")
                else:
                    self.validation_status.setText(f"✅ 前3個代號有效，共 {total_symbols} 個股票")
                    self.validation_status.setStyleSheet("color: #4CAF50; font-size: 10px; padding: 2px; font-weight: bold;")
            elif valid_count > 0:
                self.validation_status.setText(f"⚠️ {valid_count}/{len(validation_symbols)} 個代號有效，請檢查其他代號")
                self.validation_status.setStyleSheet("color: #FF9800; font-size: 10px; padding: 2px; font-weight: bold;")
            else:
                self.validation_status.setText("❌ 股票代號可能無效，請檢查格式和市場選擇")
                self.validation_status.setStyleSheet("color: #F44336; font-size: 10px; padding: 2px; font-weight: bold;")

        except Exception as e:
            self.validation_status.setText("⚠️ 驗證時發生錯誤，但仍可嘗試下載")
            self.validation_status.setStyleSheet("color: #FF9800; font-size: 10px; padding: 2px; font-weight: bold;")

    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(
            self,
            "選擇輸出目錄",
            self.output_dir_label.text()
        )
        if directory:
            self.output_dir_label.setText(directory)

    def clear_form(self):
        self.stock_text.clear()
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.status_label.setText("準備就緒")
        self.download_results = {"success": [], "failed": [], "total": 0}
        self.retry_btn.setEnabled(False)
        self.validation_status.setText("✅ 準備輸入股票代號")
        self.validation_status.setStyleSheet("color: #4CAF50; font-size: 10px; padding: 2px; font-weight: bold;")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

    def start_download(self):
        # 驗證輸入
        stocks_text = self.stock_text.toPlainText().strip()
        if not stocks_text:
            QMessageBox.warning(self, "錯誤", "請輸入股票代號")
            return

        start_date = self.start_date.date().toPyDate()
        end_date = self.end_date.date().toPyDate()
        if start_date >= end_date:
            QMessageBox.warning(self, "錯誤", "開始日期必須早於結束日期")
            return

        # 驗證時間間隔和日期範圍
        interval = self.interval_combo.currentText()
        is_valid, error_msg = self.validate_date_range(interval, start_date, end_date)
        if not is_valid:
            reply = QMessageBox.question(
                self,
                "日期範圍警告",
                f"{error_msg}\n\n是否仍要繼續下載？（可能無法獲取數據）",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        output_dir = self.output_dir_label.text()
        if not os.path.exists(output_dir):
            QMessageBox.warning(self, "錯誤", "輸出目錄不存在")
            return

        # 解析股票代號
        symbols = [s.strip() for s in stocks_text.replace('\n', ',').split(',') if s.strip()]

        # 獲取選中的輸出格式
        output_format = "CSV"
        if self.excel_radio.isChecked():
            output_format = "Excel"
        elif self.sqlite_radio.isChecked():
            output_format = "SQLite"

        # 獲取市場信息
        market = self.market_combo.currentText()
        market_info = self.markets[market]

        # 準備下載
        self.stop_flag = False  # 重置停止標誌
        self.download_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)

        # 創建並啟動工作線程
        self.download_worker = DownloadWorker(
            symbols=symbols,
            market_info=market_info,
            interval=self.interval_combo.currentText(),
            start_date=start_date.strftime("%Y-%m-%d"),
            end_date=end_date.strftime("%Y-%m-%d"),
            output_format=output_format,
            output_dir=output_dir
        )

        # 重置結果追蹤
        self.download_results = {"success": [], "failed": [], "total": len(symbols)}

        # 連接信號
        self.download_worker.progress_updated.connect(self.update_progress)
        self.download_worker.log_message.connect(self.log_message)
        self.download_worker.status_updated.connect(self.status_label.setText)
        self.download_worker.finished.connect(self.download_finished)
        self.download_worker.download_result.connect(self.track_download_result)

        self.download_worker.start()

    def stop_download(self):
        if self.download_worker:
            self.stop_flag = True
            self.download_worker.stop()
            self.log_message("正在停止下載...")

    def update_progress(self, current, total):
        if total > 0:
            progress = int((current / total) * 100)
            self.progress_bar.setValue(progress)

    def track_download_result(self, symbol, success, message):
        """追蹤每個股票的下載結果"""
        if success:
            self.download_results["success"].append({"symbol": symbol, "message": message})
        else:
            self.download_results["failed"].append({"symbol": symbol, "message": message})

    def download_finished(self, success):
        self.download_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        # 顯示下載結果摘要
        total = self.download_results["total"]
        success_count = len(self.download_results["success"])
        failed_count = len(self.download_results["failed"])

        if success and not self.stop_flag:
            result_msg = f"下載完成！\n\n"
            result_msg += f"📊 總計: {total} 個股票\n"
            result_msg += f"✅ 成功: {success_count} 個\n"
            result_msg += f"❌ 失敗: {failed_count} 個\n\n"

            if failed_count > 0:
                result_msg += "失敗的股票:\n"
                for item in self.download_results["failed"][:5]:  # 只顯示前5個
                    result_msg += f"• {item['symbol']}: {item['message'][:50]}...\n"
                if len(self.download_results["failed"]) > 5:
                    result_msg += f"... 還有 {len(self.download_results['failed']) - 5} 個失敗項目\n"

                self.retry_btn.setEnabled(True)
                result_msg += "\n💡 可以使用「重試失敗」按鈕重新下載失敗的項目"
            else:
                result_msg += "🎉 所有股票都下載成功！"

            QMessageBox.information(self, "下載結果", result_msg)
        else:
            # 啟用重試按鈕（如果有失敗項目）
            if len(self.download_results["failed"]) > 0:
                self.retry_btn.setEnabled(True)

        self.download_worker = None

    def retry_failed(self):
        """重試失敗的股票"""
        if not self.download_results["failed"]:
            QMessageBox.information(self, "提示", "沒有需要重試的項目")
            return

        # 準備重試
        failed_symbols = [item["symbol"] for item in self.download_results["failed"]]

        reply = QMessageBox.question(
            self,
            "重試確認",
            f"即將重試 {len(failed_symbols)} 個失敗的股票:\n\n" +
            "\n".join(failed_symbols[:10]) +
            ("\n..." if len(failed_symbols) > 10 else "") +
            "\n\n確定要重試嗎？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )

        if reply == QMessageBox.StandardButton.Yes:
            # 清空失敗記錄，準備重試
            self.download_results["failed"] = []

            # 使用相同的設置重新下載失敗的項目
            self.retry_btn.setEnabled(False)
            self.download_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.progress_bar.setValue(0)

            # 獲取當前設置
            start_date = self.start_date.date().toPyDate()
            end_date = self.end_date.date().toPyDate()
            interval = self.interval_combo.currentText()
            market = self.market_combo.currentText()
            market_info = self.markets[market]

            # 獲取輸出格式和目錄
            output_format = "CSV"
            if self.excel_radio.isChecked():
                output_format = "Excel"
            elif self.sqlite_radio.isChecked():
                output_format = "SQLite"

            output_dir = self.output_dir_label.text()

            # 創建工作線程重試失敗項目
            self.download_worker = DownloadWorker(
                symbols=failed_symbols,
                market_info=market_info,
                interval=interval,
                start_date=start_date.strftime("%Y-%m-%d"),
                end_date=end_date.strftime("%Y-%m-%d"),
                output_format=output_format,
                output_dir=output_dir
            )

            # 連接信號
            self.download_worker.progress_updated.connect(self.update_progress)
            self.download_worker.log_message.connect(self.log_message)
            self.download_worker.status_updated.connect(self.status_label.setText)
            self.download_worker.finished.connect(self.download_finished)
            self.download_worker.download_result.connect(self.track_download_result)

            self.download_worker.start()
            self.log_message(f"開始重試 {len(failed_symbols)} 個失敗的股票...")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("股票數據爬取器 - 增強版")
    app.setApplicationDisplayName("股票數據爬取器 - 增強版")

    # 設置應用程序圖標（如果有的話）
    # app.setWindowIcon(QIcon("icon.png"))

    window = EnhancedStockGUI()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()