import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import yfinance as yf
import pandas as pd
import sqlite3
import threading
from datetime import datetime, timedelta
import os

class StockDataGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("股票數據爬取器")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        # 股票國家/市場映射
        self.markets = {
            "美國": {"suffix": "", "currency": "USD", "examples": "AAPL, MSFT, GOOGL"},
            "台灣": {"suffix": ".TW", "currency": "TWD", "examples": "2330.TW, 2317.TW"},
            "香港": {"suffix": ".HK", "currency": "HKD", "examples": "0700.HK, 0005.HK"},
            "日本": {"suffix": ".T", "currency": "JPY", "examples": "7203.T, 6758.T"},
            "德國": {"suffix": ".DE", "currency": "EUR", "examples": "SAP.DE, VOW3.DE"},
            "英國": {"suffix": ".L", "currency": "GBP", "examples": "BARC.L, LLOY.L"}
        }

        self.setup_ui()

    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 股票代號輸入
        ttk.Label(main_frame, text="股票代號 (用逗號分隔多個股票):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.stock_entry = tk.Text(main_frame, height=3, width=60)
        self.stock_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # 股票市場選擇
        ttk.Label(main_frame, text="股票市場:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.market_var = tk.StringVar(value="美國")
        market_combo = ttk.Combobox(main_frame, textvariable=self.market_var,
                                   values=list(self.markets.keys()), state="readonly")
        market_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        market_combo.bind('<<ComboboxSelected>>', self.update_market_example)

        # 市場示例標籤
        self.market_example = ttk.Label(main_frame, text=f"範例: {self.markets['美國']['examples']}")
        self.market_example.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=2)

        # 時間間隔選擇
        ttk.Label(main_frame, text="時間間隔:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.interval_var = tk.StringVar(value="1d")
        interval_combo = ttk.Combobox(main_frame, textvariable=self.interval_var,
                                     values=["1m", "2m", "5m", "15m", "30m", "60m", "1h", "1d", "1wk", "1mo"],
                                     state="readonly")
        interval_combo.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)

        # 時間間隔限制提示
        interval_hint = ttk.Label(main_frame,
                                 text="⚠️ 限制: 1m(7天), 2m-30m(59天/60天內), 60m/1h(729天), 1d+(無限)",
                                 foreground="#FF6600", font=("Arial", 8))
        interval_hint.grid(row=4, column=2, sticky=tk.W, padx=(10, 0))

        # 日期範圍
        date_frame = ttk.LabelFrame(main_frame, text="日期範圍", padding="5")
        date_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(date_frame, text="開始日期:").grid(row=0, column=0, sticky=tk.W)
        self.start_date = tk.StringVar(value=(datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d"))
        start_entry = ttk.Entry(date_frame, textvariable=self.start_date, width=12)
        start_entry.grid(row=0, column=1, padx=5)

        ttk.Label(date_frame, text="結束日期:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.end_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        end_entry = ttk.Entry(date_frame, textvariable=self.end_date, width=12)
        end_entry.grid(row=0, column=3, padx=5)

        # 輸出設置
        output_frame = ttk.LabelFrame(main_frame, text="輸出設置", padding="5")
        output_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # 輸出格式選擇
        ttk.Label(output_frame, text="輸出格式:").grid(row=0, column=0, sticky=tk.W)
        self.output_format = tk.StringVar(value="CSV")
        format_frame = ttk.Frame(output_frame)
        format_frame.grid(row=0, column=1, sticky=tk.W)

        ttk.Radiobutton(format_frame, text="CSV", variable=self.output_format, value="CSV").pack(side=tk.LEFT)
        ttk.Radiobutton(format_frame, text="Excel", variable=self.output_format, value="Excel").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(format_frame, text="SQLite", variable=self.output_format, value="SQLite").pack(side=tk.LEFT)

        # 輸出目錄選擇
        ttk.Label(output_frame, text="輸出目錄:").grid(row=1, column=0, sticky=tk.W, pady=5)
        dir_frame = ttk.Frame(output_frame)
        dir_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        self.output_dir = tk.StringVar(value=os.getcwd())
        ttk.Entry(dir_frame, textvariable=self.output_dir, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(dir_frame, text="瀏覽", command=self.browse_directory).pack(side=tk.RIGHT, padx=(5, 0))

        # 控制按鈕
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=20)

        self.download_btn = ttk.Button(button_frame, text="開始下載", command=self.start_download)
        self.download_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(button_frame, text="停止", command=self.stop_download, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="清空", command=self.clear_form).pack(side=tk.LEFT, padx=5)

        # 進度條
        progress_frame = ttk.LabelFrame(main_frame, text="進度", padding="5")
        progress_frame.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)

        self.status_label = ttk.Label(progress_frame, text="準備就緒")
        self.status_label.grid(row=1, column=0, sticky=tk.W)

        # 日誌區域
        log_frame = ttk.LabelFrame(main_frame, text="下載日誌", padding="5")
        log_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        self.log_text = tk.Text(log_frame, height=8, width=70)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # 配置行列權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(9, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        output_frame.columnconfigure(1, weight=1)
        dir_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(0, weight=1)

        self.stop_flag = False

    def update_market_example(self, event=None):
        market = self.market_var.get()
        example = self.markets[market]['examples']
        self.market_example.config(text=f"範例: {example}")

    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.output_dir.get())
        if directory:
            self.output_dir.set(directory)

    def clear_form(self):
        self.stock_entry.delete('1.0', tk.END)
        self.log_text.delete('1.0', tk.END)
        self.progress['value'] = 0
        self.status_label.config(text="準備就緒")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_download(self):
        # 驗證輸入
        stocks_text = self.stock_entry.get('1.0', tk.END).strip()
        if not stocks_text:
            messagebox.showerror("錯誤", "請輸入股票代號")
            return

        try:
            start_date = datetime.strptime(self.start_date.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.end_date.get(), "%Y-%m-%d")
            today = datetime.now()

            if start_date >= end_date:
                messagebox.showerror("錯誤", "開始日期必須早於結束日期")
                return

            if end_date > today:
                messagebox.showerror("錯誤", f"結束日期不能是未來日期！今天是 {today.strftime('%Y-%m-%d')}")
                return

            # 檢查時間間隔限制
            interval = self.interval_var.get()
            days_diff = (end_date - start_date).days
            days_from_today = (today - start_date).days

            if interval == "1m" and (days_diff > 7 or days_from_today > 30):
                result = messagebox.askquestion("日期範圍警告",
                                               "1分鐘數據：範圍最多7天，須在最近30天內\n\n是否繼續下載？",
                                               icon='warning')
                if result != 'yes':
                    return

            elif interval in ["2m", "5m", "15m", "30m"]:
                if days_diff > 59:
                    messagebox.showerror("錯誤",
                                        f"{interval}數據的日期範圍不能超過59天（Yahoo Finance限制）\n當前選擇了{days_diff}天")
                    return
                if days_from_today > 60:
                    messagebox.showerror("錯誤",
                                        f"{interval}數據只能獲取最近60天內的數據\n開始日期距今{days_from_today}天")
                    return

            elif interval in ["60m", "1h"] and days_diff > 729:
                messagebox.showerror("錯誤",
                                    f"小時數據的日期範圍不能超過729天（Yahoo Finance限制）\n當前選擇了{days_diff}天")
                return

        except ValueError:
            messagebox.showerror("錯誤", "日期格式不正確，請使用 YYYY-MM-DD 格式")
            return

        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("錯誤", "輸出目錄不存在")
            return

        # 準備下載
        self.stop_flag = False
        self.download_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        # 在新線程中執行下載
        download_thread = threading.Thread(target=self.download_stocks)
        download_thread.daemon = True
        download_thread.start()

    def stop_download(self):
        self.stop_flag = True
        self.log_message("正在停止下載...")

    def download_stocks(self):
        try:
            # 解析股票代號
            stocks_text = self.stock_entry.get('1.0', tk.END).strip()
            stock_symbols = [s.strip() for s in stocks_text.replace('\n', ',').split(',') if s.strip()]

            # 添加市場後綴
            market = self.market_var.get()
            suffix = self.markets[market]['suffix']
            if suffix:
                stock_symbols = [s + suffix if not any(s.endswith(suf) for suf in ['.TW', '.HK', '.T', '.DE', '.L']) else s
                               for s in stock_symbols]

            total_stocks = len(stock_symbols)
            self.progress['maximum'] = total_stocks

            # 準備SQLite連接（如果需要）
            db_connection = None
            if self.output_format.get() == "SQLite":
                db_path = os.path.join(self.output_dir.get(), "stock_data.db")
                db_connection = sqlite3.connect(db_path)
                self.log_message(f"SQLite數據庫: {db_path}")

            # 下載每隻股票
            for i, symbol in enumerate(stock_symbols):
                if self.stop_flag:
                    break

                self.status_label.config(text=f"正在下載 {symbol} ({i+1}/{total_stocks})")
                self.log_message(f"開始下載 {symbol}")

                try:
                    # 下載數據
                    ticker = yf.Ticker(symbol)
                    data = ticker.history(
                        start=self.start_date.get(),
                        end=self.end_date.get(),
                        interval=self.interval_var.get()
                    )

                    if data.empty:
                        self.log_message(f"警告: {symbol} 沒有數據")
                        continue

                    # 保存數據
                    self.save_data(symbol, data, db_connection)
                    self.log_message(f"完成 {symbol} - {len(data)} 條記錄")

                except Exception as e:
                    self.log_message(f"錯誤 {symbol}: {str(e)}")

                # 更新進度
                self.progress['value'] = i + 1
                self.root.update_idletasks()

            # 清理
            if db_connection:
                db_connection.close()

            if not self.stop_flag:
                self.status_label.config(text="下載完成")
                self.log_message("所有股票數據下載完成")
                messagebox.showinfo("完成", "股票數據下載完成")
            else:
                self.status_label.config(text="下載已停止")
                self.log_message("下載已被用戶停止")

        except Exception as e:
            self.log_message(f"下載過程發生錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"下載過程發生錯誤: {str(e)}")
        finally:
            self.download_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    def save_data(self, symbol, data, db_connection=None):
        output_format = self.output_format.get()
        output_dir = self.output_dir.get()
        interval = self.interval_var.get()

        # 清理檔案名稱中的特殊字符
        safe_symbol = symbol.replace('.', '_').replace(':', '_')

        # 添加Symbol欄位並統一時間格式
        data_copy = data.copy()
        data_copy.insert(0, 'Symbol', symbol)  # 在第一列插入Symbol

        # 統一使用完整日期時間格式 YYYY-MM-DD HH:MM:SS（所有時間間隔）
        data_copy.index = data_copy.index.strftime('%Y-%m-%d %H:%M:%S')

        if output_format == "CSV":
            filename = os.path.join(output_dir, f"{safe_symbol}_data.csv")
            data_copy.to_csv(filename, index_label='Date')

        elif output_format == "Excel":
            filename = os.path.join(output_dir, f"{safe_symbol}_data.xlsx")
            data_copy.to_excel(filename, index_label='Date')

        elif output_format == "SQLite" and db_connection:
            table_name = f"stock_{safe_symbol}"
            data_copy.to_sql(table_name, db_connection, if_exists='replace', index_label='Date')

def main():
    root = tk.Tk()
    app = StockDataGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()