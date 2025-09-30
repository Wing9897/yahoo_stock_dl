#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""股票定期投資分析器 - PyQt6 GUI版本"""

import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QLineEdit,
                             QTextEdit, QFileDialog, QGroupBox, QRadioButton,
                             QDateEdit, QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont
import pandas as pd
from stock_analyzer import StockInvestmentAnalyzer

class StockAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.analyzer = StockInvestmentAnalyzer()
        self.file_path = None
        self.data = None
        self.data_info = None
        self.init_ui()

    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle('股票定期投資分析器')
        self.setGeometry(100, 100, 1000, 700)

        # 設置全局樣式 - 高對比度
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 13px;
                border: 2px solid #2c3e50;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                color: #2c3e50;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QLabel {
                color: #000000;
                font-size: 12px;
            }
            QPushButton {
                font-size: 13px;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 8px;
                min-height: 25px;
            }
            QPushButton:hover {
                opacity: 0.9;
            }
            QTableWidget {
                border: 2px solid #34495e;
                gridline-color: #bdc3c7;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #34495e;
                color: white;
                padding: 8px;
                border: 1px solid #2c3e50;
                font-weight: bold;
                font-size: 12px;
            }
            QRadioButton {
                font-size: 12px;
                color: #000000;
                spacing: 5px;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
            }
            QDateEdit {
                border: 2px solid #7f8c8d;
                border-radius: 3px;
                padding: 5px;
                font-size: 12px;
                background-color: white;
            }
            QDateEdit:disabled {
                background-color: #ecf0f1;
                color: #7f8c8d;
            }
        """)

        # 主視窗
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 文件選擇區
        file_group = self.create_file_selection()
        main_layout.addWidget(file_group)

        # 數據概要區
        self.info_label = QLabel('請先選擇CSV文件')
        self.info_label.setStyleSheet("""
            padding: 15px;
            background-color: #ecf0f1;
            border: 2px solid #95a5a6;
            border-radius: 5px;
            color: #2c3e50;
            font-size: 13px;
            font-weight: bold;
        """)
        main_layout.addWidget(self.info_label)

        # 時間範圍選擇區
        time_group = self.create_time_selection()
        main_layout.addWidget(time_group)

        # 執行按鈕
        self.run_button = QPushButton('▶ 執行分析')
        self.run_button.setEnabled(False)
        self.run_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 12px;
                font-size: 15px;
                font-weight: bold;
                border: 2px solid #229954;
                min-height: 35px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
                border: 2px solid #7f8c8d;
                color: #ecf0f1;
            }
        """)
        self.run_button.clicked.connect(self.run_analysis)
        main_layout.addWidget(self.run_button)

        # 結果表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels([
            '策略', '投資次數', '總投資', '手續費', '累積股數', '平均成本', '最終價值', '報酬率'
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.result_table)

        # 結論區
        self.conclusion_label = QLabel('')
        self.conclusion_label.setStyleSheet("""
            padding: 15px;
            background-color: #d5f4e6;
            border: 3px solid #27ae60;
            border-radius: 5px;
            font-weight: bold;
            font-size: 13px;
            color: #0e3b24;
        """)
        main_layout.addWidget(self.conclusion_label)

    def create_file_selection(self):
        """創建文件選擇區"""
        group = QGroupBox('1. 選擇股票數據文件')
        layout = QHBoxLayout()

        self.file_label = QLabel('未選擇文件')
        self.file_label.setStyleSheet("""
            padding: 8px;
            background-color: white;
            border: 2px solid #bdc3c7;
            border-radius: 3px;
            color: #34495e;
            font-size: 12px;
        """)
        layout.addWidget(self.file_label)

        select_button = QPushButton('📁 瀏覽...')
        select_button.setStyleSheet("""
            background-color: #3498db;
            color: white;
            padding: 8px 20px;
            font-weight: bold;
            border: 2px solid #2980b9;
        """)
        select_button.clicked.connect(self.select_file)
        layout.addWidget(select_button)

        group.setLayout(layout)
        return group

    def create_time_selection(self):
        """創建時間範圍選擇區"""
        group = QGroupBox('2. 選擇時間範圍')
        layout = QVBoxLayout()

        # 單選按鈕
        self.use_all_radio = QRadioButton('使用全部數據範圍')
        self.use_all_radio.setChecked(True)
        self.use_all_radio.toggled.connect(self.toggle_date_inputs)
        layout.addWidget(self.use_all_radio)

        self.use_custom_radio = QRadioButton('自訂時間範圍')
        layout.addWidget(self.use_custom_radio)

        # 日期選擇
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel('開始日期:'))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate(2020, 1, 1))
        self.start_date.setEnabled(False)
        date_layout.addWidget(self.start_date)

        date_layout.addWidget(QLabel('結束日期:'))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate(2023, 12, 31))
        self.end_date.setEnabled(False)
        date_layout.addWidget(self.end_date)

        layout.addLayout(date_layout)

        group.setLayout(layout)
        return group

    def toggle_date_inputs(self):
        """切換日期輸入框啟用狀態"""
        enabled = self.use_custom_radio.isChecked()
        self.start_date.setEnabled(enabled)
        self.end_date.setEnabled(enabled)

    def select_file(self):
        """選擇CSV文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '選擇股票數據文件', '', 'CSV Files (*.csv)'
        )

        if file_path:
            # 驗證文件
            is_valid, message = self.analyzer.validate_csv(file_path)

            if is_valid:
                self.file_path = file_path
                self.file_label.setText(file_path)

                # 載入並分析數據
                try:
                    self.data = self.analyzer.load_data(file_path)
                    self.data_info = self.analyzer.analyze_data(self.data)

                    # 更新信息顯示
                    stock_symbol = self.data['Symbol'].iloc[0] if 'Symbol' in self.data.columns else 'N/A'
                    granularity = self.data_info.get('granularity', 'unknown')
                    info_text = f"""
                    <div style='color: #2c3e50;'>
                    <b style='font-size: 14px; color: #000000;'>📊 數據概要</b><br>
                    <span style='color: #34495e;'>股票代號: <b style='color: #2980b9;'>{stock_symbol}</b></span><br>
                    <span style='color: #34495e;'>檔案名稱: <b>{file_path.split('/')[-1]}</b></span><br>
                    <span style='color: #34495e;'>數據範圍: <b>{self.data_info['start_date'].strftime('%Y-%m-%d')}</b> 到 <b>{self.data_info['end_date'].strftime('%Y-%m-%d')}</b></span><br>
                    <span style='color: #34495e;'>總天數: <b style='color: #e74c3c;'>{self.data_info['total_days']}</b> 天</span><br>
                    <span style='color: #34495e;'>數據粒度: <b style='color: #8e44ad;'>{granularity}</b></span><br>
                    <span style='color: #34495e;'>價格範圍: <b>${self.data_info['min_price']:.2f}</b> - <b>${self.data_info['max_price']:.2f}</b></span><br>
                    <span style='color: #34495e;'>最新價格: <b style='color: #2980b9; font-size: 13px;'>${self.data_info['latest_price']:.2f}</b></span><br>
                    <span style='color: #34495e;'>整體報酬: <b style='color: #27ae60;'>{self.data_info['total_return']:+.2f}%</b> (年化: <b style='color: #27ae60;'>{self.data_info['annual_return']:+.2f}%</b>)</span>
                    </div>
                    """
                    self.info_label.setText(info_text)

                    # 設置日期範圍
                    start_qdate = QDate(
                        self.data_info['start_date'].year,
                        self.data_info['start_date'].month,
                        self.data_info['start_date'].day
                    )
                    end_qdate = QDate(
                        self.data_info['end_date'].year,
                        self.data_info['end_date'].month,
                        self.data_info['end_date'].day
                    )
                    self.start_date.setDateRange(start_qdate, end_qdate)
                    self.end_date.setDateRange(start_qdate, end_qdate)
                    self.start_date.setDate(start_qdate)
                    self.end_date.setDate(end_qdate)

                    # 啟用執行按鈕
                    self.run_button.setEnabled(True)

                except Exception as e:
                    self.info_label.setText(f'錯誤: {str(e)}')
                    self.run_button.setEnabled(False)
            else:
                self.file_label.setText(f'驗證失敗: {message}')
                self.run_button.setEnabled(False)

    def run_analysis(self):
        """執行分析"""
        if not self.data_info:
            return

        # 確定時間範圍
        if self.use_all_radio.isChecked():
            start = self.data_info['start_date'].strftime('%Y-%m-%d')
            end = self.data_info['end_date'].strftime('%Y-%m-%d')
        else:
            start = self.start_date.date().toString('yyyy-MM-dd')
            end = self.end_date.date().toString('yyyy-MM-dd')

        # 根據數據粒度選擇合適的策略
        granularity = self.data_info.get('granularity', 'unknown')
        base_strategies = self.analyzer.get_strategies_for_granularity(granularity)

        # 選擇最多8個策略進行分析
        strategies = []
        if len(base_strategies) <= 8:
            strategies = [(name, amount, freq) for name, amount, freq in base_strategies]
        else:
            # 如果策略太多,均勻選取8個
            step = len(base_strategies) / 8
            for i in range(8):
                idx = int(i * step)
                strategies.append(base_strategies[idx])

        # 執行分析
        results = []
        for name, amount, freq in strategies:
            result = self.analyzer.simulate_investment(self.data, start, end, amount, freq)
            results.append((name, result))

        # 顯示結果到表格
        self.result_table.setRowCount(len(results))

        for i, (name, result) in enumerate(results):
            # 策略名稱 - 粗體
            name_item = QTableWidgetItem(name)
            name_item.setFont(QFont('Arial', 11, QFont.Weight.Bold))
            self.result_table.setItem(i, 0, name_item)

            self.result_table.setItem(i, 1, QTableWidgetItem(str(result['count'])))
            self.result_table.setItem(i, 2, QTableWidgetItem(f"${result['invested']:,.2f}"))
            self.result_table.setItem(i, 3, QTableWidgetItem(f"${result['commission']:.2f}"))
            self.result_table.setItem(i, 4, QTableWidgetItem(f"{result['shares']:.4f}"))
            self.result_table.setItem(i, 5, QTableWidgetItem(f"${result['invested'] / result['shares']:.2f}"))
            self.result_table.setItem(i, 6, QTableWidgetItem(f"${result['value']:,.2f}"))

            # 報酬率用顏色標示 - 更高對比度
            return_item = QTableWidgetItem(f"{result['return']:.2f}%")
            return_font = QFont('Arial', 11, QFont.Weight.Bold)
            return_item.setFont(return_font)
            if result['return'] > 0:
                # 深綠色，更高對比度
                return_item.setForeground(Qt.GlobalColor.darkGreen)
                return_item.setBackground(Qt.GlobalColor.lightGreen)
            else:
                # 深紅色
                return_item.setForeground(Qt.GlobalColor.darkRed)
                return_item.setBackground(Qt.GlobalColor.red)
            self.result_table.setItem(i, 7, return_item)

        # 找出最佳和最差策略
        best = max(results, key=lambda x: x[1]['return'])
        worst = min(results, key=lambda x: x[1]['return'])

        conclusion_text = f"""
        <div style='color: #0e3b24; font-size: 13px;'>
        <b style='font-size: 15px; color: #000000;'>🏆 結論</b><br><br>
        <span style='font-size: 13px;'><b>✓ 最佳策略:</b> <b style='color: #27ae60; font-size: 14px;'>{best[0]}</b></span><br>
        <span style='margin-left: 20px; color: #34495e;'>報酬率: <b style='color: #27ae60; font-size: 14px;'>{best[1]['return']:.2f}%</b></span><br><br>
        <span style='font-size: 13px;'><b>✗ 最差策略:</b> <b style='color: #c0392b; font-size: 14px;'>{worst[0]}</b></span><br>
        <span style='margin-left: 20px; color: #34495e;'>報酬率: <b style='color: #c0392b;'>{worst[1]['return']:.2f}%</b></span><br><br>
        <span style='color: #34495e;'><b>差異:</b> <b style='color: #e74c3c; font-size: 14px;'>{best[1]['return'] - worst[1]['return']:.2f}%</b></span><br>
        <span style='color: #34495e;'><b>最終股價:</b> <b style='color: #2980b9; font-size: 14px;'>${results[0][1]['final_price']:.2f}</b></span>
        </div>
        """
        self.conclusion_label.setText(conclusion_text)

def main():
    app = QApplication(sys.argv)

    # 設置應用風格
    app.setStyle('Fusion')

    window = StockAnalyzerGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()