#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""股票定期投資分析器 - 比較每週50元 vs 每月200元"""

import pandas as pd
from datetime import timedelta
import os

class StockInvestmentAnalyzer:
    def __init__(self, commission_rate=0.001):
        self.commission_rate = commission_rate

    def validate_csv(self, file_path):
        """驗證CSV格式"""
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"

        try:
            df = pd.read_csv(file_path, nrows=5)
            required = ['Symbol', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']
            missing = [col for col in required if col not in df.columns]

            if missing:
                return False, f"缺少欄位: {missing}"

            pd.to_datetime(df['Date'].iloc[0])
            return True, "CSV格式驗證通過"
        except Exception as e:
            return False, f"驗證錯誤: {e}"

    def load_data(self, file_path):
        """載入並處理股票數據"""
        df = pd.read_csv(file_path)

        # 處理 Symbol 欄位 (移除引號)
        if 'Symbol' in df.columns and df['Symbol'].dtype == 'object':
            df['Symbol'] = df['Symbol'].str.replace('"', '')

        # 處理 Date 欄位 (移除引號並轉換為日期時間)
        if df['Date'].dtype == 'object':
            df['Date'] = df['Date'].str.replace('"', '')
        df['Date'] = pd.to_datetime(df['Date'])

        # 處理數值欄位 (移除引號並轉換為浮點數)
        for col in ['Open', 'High', 'Low', 'Close', 'Volume']:
            if df[col].dtype == 'object':
                df[col] = df[col].str.replace('"', '').astype(float)

        return df.sort_values('Date').reset_index(drop=True)

    def detect_data_granularity(self, df):
        """檢測數據時間粒度"""
        if len(df) < 2:
            return 'unknown'

        # 計算前100個數據點的平均時間間隔
        sample_size = min(100, len(df))
        time_diffs = []
        for i in range(1, sample_size):
            diff = (df['Date'].iloc[i] - df['Date'].iloc[i-1]).total_seconds() / 60  # 轉換為分鐘
            if diff > 0:  # 忽略重複時間戳
                time_diffs.append(diff)

        if not time_diffs:
            return 'unknown'

        avg_diff = sum(time_diffs) / len(time_diffs)

        # 判斷粒度 (允許一些誤差)
        if avg_diff < 45:  # < 45分鐘
            return 'intraday_minute'  # 分鐘級 (支持半小時、小時、3小時等)
        elif avg_diff < 90:  # < 1.5小時
            return 'intraday_hourly'  # 小時級 (支持小時、3小時、6小時等)
        elif avg_diff < 18 * 60:  # < 18小時
            return 'intraday_half_day'  # 半天級 (支持12小時、日等)
        elif avg_diff < 36 * 60:  # < 36小時
            return 'daily'  # 日級 (支持日、3日、週等)
        else:
            return 'daily_plus'  # 日級以上 (支持週、月、季等)

    def get_strategies_for_granularity(self, granularity):
        """根據數據粒度返回合適的投資策略"""
        if granularity == 'intraday_minute':
            # 分鐘級數據: 半小時到日
            return [
                ("每半小時", 0.5, 'halfhourly'),
                ("每小時", 1, 'hourly'),
                ("每3小時", 3, '3hourly'),
                ("每6小時", 6, '6hourly'),
                ("每12小時", 12, '12hourly'),
                ("每日", 24, 'daily'),
                ("每3日", 72, '3daily'),
                ("每週", 168, 'weekly')
            ]
        elif granularity == 'intraday_hourly':
            # 小時級數據: 小時到週
            return [
                ("每小時", 1, 'hourly'),
                ("每3小時", 3, '3hourly'),
                ("每6小時", 6, '6hourly'),
                ("每12小時", 12, '12hourly'),
                ("每日", 24, 'daily'),
                ("每3日", 72, '3daily'),
                ("每週", 168, 'weekly')
            ]
        elif granularity == 'intraday_half_day':
            # 半天級數據: 12小時到月
            return [
                ("每12小時", 12, '12hourly'),
                ("每日", 24, 'daily'),
                ("每3日", 72, '3daily'),
                ("每週", 168, 'weekly'),
                ("每月", 720, 'monthly')
            ]
        elif granularity == 'daily':
            # 日級數據: 日到半年
            return [
                ("每日", 24, 'daily'),
                ("每3日", 72, '3daily'),
                ("每週", 168, 'weekly'),
                ("每月", 720, 'monthly'),
                ("每季", 2160, 'quarterly'),
                ("每半年", 4320, 'semiannually')
            ]
        else:  # daily_plus or unknown
            # 日級以上: 週到年
            return [
                ("每週", 168, 'weekly'),
                ("每月", 720, 'monthly'),
                ("每季", 2160, 'quarterly'),
                ("每半年", 4320, 'semiannually'),
                ("每年", 8760, 'yearly')
            ]

    def analyze_data(self, df):
        """分析數據範圍"""
        first_price = df['Close'].iloc[0]
        latest_price = df['Close'].iloc[-1]
        years = (df['Date'].max() - df['Date'].min()).days / 365.25

        # 檢測數據粒度
        granularity = self.detect_data_granularity(df)

        return {
            'start_date': df['Date'].min(),
            'end_date': df['Date'].max(),
            'total_days': len(df),
            'min_price': df['Low'].min(),
            'max_price': df['High'].max(),
            'latest_price': latest_price,
            'total_return': ((latest_price - first_price) / first_price) * 100,
            'annual_return': ((latest_price / first_price) ** (1/years) - 1) * 100 if years > 0 else 0,
            'granularity': granularity
        }

    def simulate_investment(self, data, start, end, amount, frequency='weekly'):
        """分析定期投資"""
        investments = []
        current = pd.to_datetime(start)
        end_dt = pd.to_datetime(end)

        while current <= end_dt:
            available = data[data['Date'] >= current]
            if len(available) == 0 or available.iloc[0]['Date'] > end_dt:
                break

            price = available.iloc[0]['Close']
            net = amount * (1 - self.commission_rate)
            shares = net / price

            investments.append({
                'shares': shares,
                'commission': amount * self.commission_rate
            })

            # 移動到下一個投資日
            if frequency == 'halfhourly':  # 每半小時
                current += timedelta(minutes=30)
            elif frequency == 'hourly':  # 每小時
                current += timedelta(hours=1)
            elif frequency == '3hourly':  # 每3小時
                current += timedelta(hours=3)
            elif frequency == '6hourly':  # 每6小時
                current += timedelta(hours=6)
            elif frequency == '12hourly':  # 每12小時
                current += timedelta(hours=12)
            elif frequency == 'daily':  # 每日
                current += timedelta(days=1)
            elif frequency == '3daily':  # 每3日
                current += timedelta(days=3)
            elif frequency == 'weekly':  # 每週
                current += timedelta(days=7)
            elif frequency == 'monthly':  # 每月
                try:
                    current = current.replace(month=current.month + 1) if current.month < 12 else current.replace(year=current.year + 1, month=1)
                except ValueError:
                    current = current.replace(month=current.month + 1, day=1) if current.month < 12 else current.replace(year=current.year + 1, month=1, day=1)
            elif frequency == 'quarterly':  # 每3個月
                try:
                    new_month = current.month + 3
                    if new_month > 12:
                        current = current.replace(year=current.year + 1, month=new_month - 12)
                    else:
                        current = current.replace(month=new_month)
                except ValueError:
                    new_month = current.month + 3
                    if new_month > 12:
                        current = current.replace(year=current.year + 1, month=new_month - 12, day=1)
                    else:
                        current = current.replace(month=new_month, day=1)
            elif frequency == 'semiannually':  # 每半年
                try:
                    new_month = current.month + 6
                    if new_month > 12:
                        current = current.replace(year=current.year + 1, month=new_month - 12)
                    else:
                        current = current.replace(month=new_month)
                except ValueError:
                    new_month = current.month + 6
                    if new_month > 12:
                        current = current.replace(year=current.year + 1, month=new_month - 12, day=1)
                    else:
                        current = current.replace(month=new_month, day=1)
            elif frequency == 'yearly':  # 每年
                try:
                    current = current.replace(year=current.year + 1)
                except ValueError:
                    current = current.replace(year=current.year + 1, day=1)

        total_invested = len(investments) * amount
        total_shares = sum(inv['shares'] for inv in investments)
        final_price = data[data['Date'] <= end_dt].iloc[-1]['Close']
        final_value = total_shares * final_price

        return {
            'count': len(investments),
            'invested': total_invested,
            'commission': sum(inv['commission'] for inv in investments),
            'shares': total_shares,
            'value': final_value,
            'return': ((final_value - total_invested) / total_invested) * 100,
            'final_price': final_price
        }

    def run(self, file_path, start=None, end=None):
        """執行分析"""
        data = self.load_data(file_path)
        # 從數據中提取股票代號,如果沒有則使用文件名
        stock_name = data['Symbol'].iloc[0] if 'Symbol' in data.columns else os.path.basename(file_path).replace('.csv', '')
        info = self.analyze_data(data)

        print(f"\n=== {stock_name} 投資分析結果 ===")
        print(f"數據範圍: {info['start_date'].strftime('%Y-%m-%d')} 到 {info['end_date'].strftime('%Y-%m-%d')}")
        print(f"總天數: {info['total_days']} 天")
        print(f"數據粒度: {info['granularity']}")
        print(f"價格範圍: ${info['min_price']:.2f} - ${info['max_price']:.2f}")
        print(f"最新價格: ${info['latest_price']:.2f}")
        print(f"整體報酬: {info['total_return']:+.2f}% (年化: {info['annual_return']:+.2f}%)")

        start = start or info['start_date'].strftime('%Y-%m-%d')
        end = end or info['end_date'].strftime('%Y-%m-%d')
        print(f"\n分析期間: {start} 到 {end}")

        # 根據數據粒度選擇合適的策略
        base_strategies = self.get_strategies_for_granularity(info['granularity'])

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

        results = []
        for name, amount, freq in strategies:
            result = self.simulate_investment(data, start, end, amount, freq)
            results.append((name, result))

        # 顯示結果
        for name, result in results:
            print(f"\n{name}投資策略:")
            print(f"  投資次數: {result['count']}")
            print(f"  總投資: ${result['invested']:,.2f}")
            print(f"  手續費: ${result['commission']:.2f}")
            print(f"  累積股數: {result['shares']:.4f}")
            print(f"  平均成本: ${result['invested'] / result['shares']:.2f}")
            print(f"  最終價值: ${result['value']:,.2f}")
            print(f"  總報酬: ${result['value'] - result['invested']:,.2f}")
            print(f"  報酬率: {result['return']:.2f}%")

        # 比較結論 - 找出最佳策略
        best = max(results, key=lambda x: x[1]['return'])
        worst = min(results, key=lambda x: x[1]['return'])

        print(f"\n結論:")
        print(f"  最佳策略: {best[0]} (報酬率: {best[1]['return']:.2f}%)")
        print(f"  最差策略: {worst[0]} (報酬率: {worst[1]['return']:.2f}%)")
        print(f"  差異: {best[1]['return'] - worst[1]['return']:.2f}%")
        print(f"  最終股價: ${results[0][1]['final_price']:.2f}")

def main():
    """主程式"""
    sim = StockInvestmentAnalyzer()

    print("股票定期投資分析器")
    print("=" * 50)
    print("CSV格式要求: Symbol, Date, Open, High, Low, Close, Volume")

    # 輸入文件路徑
    while True:
        file_path = input("\n請輸入CSV文件路徑: ").strip().replace('"', '')
        is_valid, msg = sim.validate_csv(file_path)

        if is_valid:
            print(msg)
            break
        else:
            print(f"錯誤: {msg}")
            if input("重新輸入? (y/n): ").lower() != 'y':
                return

    try:
        # 顯示數據概要
        data = sim.load_data(file_path)
        info = sim.analyze_data(data)

        print(f"\n數據概要:")
        print(f"  範圍: {info['start_date'].strftime('%Y-%m-%d')} 到 {info['end_date'].strftime('%Y-%m-%d')}")
        print(f"  天數: {info['total_days']}")
        print(f"  價格: ${info['min_price']:.2f} - ${info['max_price']:.2f}")

        # 選擇時間範圍
        print("\n1. 使用全部數據")
        print("2. 自訂時間範圍")
        choice = input("請選擇 (1/2): ").strip()

        if choice == '1':
            sim.run(file_path)
        else:
            print("\n請輸入日期 (格式: YYYY-MM-DD)")
            start = input("開始日期: ")
            end = input("結束日期: ")
            sim.run(file_path, start, end)

    except Exception as e:
        print(f"錯誤: {e}")

if __name__ == "__main__":
    main()