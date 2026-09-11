# -*- coding: utf-8 -*-
"""
历史数据抓取脚本
功能：抓取股票池里每只股票的历史日线数据（开高低收、成交量、涨跌幅等）
用法：
    1. 先安装依赖（只需一次）：
       pip install akshare pandas
    2. 直接运行本脚本：
       python fetch_history.py
    3. 结果会保存在 history_data 文件夹下，每只股票一个csv文件

后续想更新数据，直接重新运行一次即可，会自动覆盖成最新数据。
"""

import akshare as ak
import pandas as pd
import os
import time
from stock_list import STOCK_POOL

# ============ 可以自己修改的参数 ============
START_DATE = "20230101"   # 历史数据起始日期，格式 YYYYMMDD，改成你想要的起点
END_DATE = "20500101"     # 结束日期，写得很后可以理解成"到今天为止"
OUTPUT_DIR = "history_data"  # 保存数据的文件夹名
# ==========================================


def fetch_one_stock(code: str, name: str, max_retries: int = 3) -> pd.DataFrame:
    """
    抓取单只股票的历史日线数据
    直接用腾讯接口（东财接口在你的网络环境下被拦截，已确认不通，故不再尝试）
    """
    last_error = None
    prefix = "sh" if code.startswith(("6", "9")) else "sz"
    symbol_tx = f"{prefix}{code}"

    for attempt in range(1, max_retries + 1):
        try:
            df = ak.stock_zh_a_hist_tx(
                symbol=symbol_tx,
                start_date=START_DATE,
                end_date=END_DATE,
                adjust="qfq",
            )
            # 统一列名，方便后续分析
            df = df.rename(columns={
                "date": "日期", "open": "开盘", "close": "收盘",
                "high": "最高", "low": "最低", "amount": "成交额",
                "volume": "成交量", "turnover": "换手率",
            })
            df.insert(0, "股票代码", code)
            df.insert(1, "股票名称", name)
            return df
        except Exception as e:
            last_error = e
            if attempt < max_retries:
                wait_seconds = 3 * attempt
                print(f"    第{attempt}次失败，{wait_seconds}秒后重试...")
                time.sleep(wait_seconds)

    raise last_error


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_data = []
    failed_stocks = []
    total = len(STOCK_POOL)

    for i, stock in enumerate(STOCK_POOL, start=1):
        code = stock["code"]
        name = stock["name"]
        print(f"[{i}/{total}] 正在抓取 {name}({code}) ...")

        try:
            df = fetch_one_stock(code, name)
            # 每只股票单独存一个文件，方便查看和调用
            save_path = os.path.join(OUTPUT_DIR, f"{code}_{name}.csv")
            df.to_csv(save_path, index=False, encoding="utf-8-sig")
            all_data.append(df)
        except Exception as e:
            print(f"  !! {name}({code}) 抓取失败: {e}")
            failed_stocks.append(stock)

        # 稍微停顿一下，避免请求过于频繁被限制
        time.sleep(1.5)

    # 对第一轮失败的股票，等待更长时间后做一轮补抓
    if failed_stocks:
        print(f"\n第一轮有 {len(failed_stocks)} 只失败，等待10秒后进行补抓...")
        time.sleep(10)
        still_failed = []
        for stock in failed_stocks:
            code = stock["code"]
            name = stock["name"]
            print(f"补抓 {name}({code}) ...")
            try:
                df = fetch_one_stock(code, name)
                save_path = os.path.join(OUTPUT_DIR, f"{code}_{name}.csv")
                df.to_csv(save_path, index=False, encoding="utf-8-sig")
                all_data.append(df)
            except Exception as e:
                print(f"  !! {name}({code}) 补抓仍然失败: {e}")
                still_failed.append(stock)
            time.sleep(2)
        failed_stocks = still_failed

    # 同时也生成一个合并的总表，方便整体分析
    if all_data:
        merged = pd.concat(all_data, ignore_index=True)
        merged_path = os.path.join(OUTPUT_DIR, "_合并历史数据.csv")
        merged.to_csv(merged_path, index=False, encoding="utf-8-sig")
        print(f"\n完成！共成功抓取 {len(all_data)} 只股票，合并数据已保存到 {merged_path}")
        if failed_stocks:
            names = "、".join(f"{s['name']}({s['code']})" for s in failed_stocks)
            print(f"仍有 {len(failed_stocks)} 只失败，可以隔一会儿单独重跑一次本脚本：{names}")
    else:
        print("\n没有成功抓取到任何数据，请检查网络或akshare是否安装正确。")

    # 额外抓取沪深300指数，作为后续"相对强弱"分析的比较基准
    print("\n正在抓取沪深300指数（用于相对强弱对比）...")
    try:
        benchmark_df = ak.stock_zh_a_hist_tx(
            symbol="sh000300",
            start_date=START_DATE,
            end_date=END_DATE,
            adjust="",
        )
        benchmark_df = benchmark_df.rename(columns={
            "date": "日期", "open": "开盘", "close": "收盘",
            "high": "最高", "low": "最低", "amount": "成交额",
            "volume": "成交量", "turnover": "换手率",
        })
        benchmark_df.insert(0, "股票代码", "000300")
        benchmark_df.insert(1, "股票名称", "沪深300")
        benchmark_path = os.path.join(OUTPUT_DIR, "000300_沪深300.csv")
        benchmark_df.to_csv(benchmark_path, index=False, encoding="utf-8-sig")
        print(f"沪深300指数抓取成功，已保存到 {benchmark_path}")
    except Exception as e:
        print(f"沪深300指数抓取失败: {e}（不影响个股数据，只是相对强弱这项会缺失）")


if __name__ == "__main__":
    main()
