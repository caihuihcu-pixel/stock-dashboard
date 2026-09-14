# -*- coding: utf-8 -*-
"""
早盘竞价/实时快照抓取脚本
功能：抓取股票池里每只股票【当前时刻】的行情快照
      在 9:15-9:25 集合竞价期间运行，抓到的就是竞价价格
      在盘中任意时间运行，抓到的就是实时行情

【重要提醒】数据源(新浪)官方注明：短时间内反复调用这个接口会被临时封IP。
      所以不要在早盘短时间内跑好几次，建议一个交易日只跑1次，选在集合竞价
      快结束的时候(比如9:24-9:25之间)跑，数据最有参考价值。

用法：
    python fetch_auction.py
    每次运行会在 auction_data 文件夹下：
    - 追加一行到当天的CSV记录里(竞价快照_20260101.csv这种命名)
    - 同时生成/更新一份 竞价数据.json，供 auction_dashboard.html 网页查看
"""

import akshare as ak
import pandas as pd
import os
from datetime import datetime, timezone, timedelta
from stock_list import STOCK_POOL

def beijing_now():
    return datetime.now(timezone(timedelta(hours=8)))

OUTPUT_DIR = "auction_data"
JSON_FILE = os.path.join(OUTPUT_DIR, "竞价数据.json")
# 把股票代码整理成一个 集合，方便快速筛选
CODE_SET = {s["code"] for s in STOCK_POOL}
CODE_NAME_MAP = {s["code"]: s["name"] for s in STOCK_POOL}


def fetch_snapshot() -> pd.DataFrame:
    """
    抓取全市场实时快照，然后筛选出自选股池里的股票
    直接用新浪接口（东财接口在你的网络环境下被拦截，已确认不通，故不再尝试）
    """
    full_df = ak.stock_zh_a_spot()
    source = "新浪"

    # 新浪接口代码列可能带 sh/sz 前缀，统一提取出6位数字再匹配
    full_df = full_df.copy()
    full_df["代码_纯数字"] = full_df["代码"].astype(str).str.extract(r"(\d{6})")

    filtered = full_df[full_df["代码_纯数字"].isin(CODE_SET)].copy()
    filtered.insert(0, "数据源", source)
    filtered.insert(0, "抓取时间", beijing_now().strftime("%Y-%m-%d %H:%M:%S"))
    return filtered


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("正在抓取实时快照（自选股池）...")
    try:
        df = fetch_snapshot()
    except Exception as e:
        print(f"抓取失败: {e}")
        return

    if df.empty:
        print("没有匹配到自选股池里的股票，请检查 stock_list.py 里的代码是否正确。")
        return

    today_str = beijing_now().strftime("%Y%m%d")
    save_path = os.path.join(OUTPUT_DIR, f"竞价快照_{today_str}.csv")

    # 如果今天已经跑过，就追加一行；没有就新建文件
    if os.path.exists(save_path):
        df.to_csv(save_path, mode="a", header=False, index=False, encoding="utf-8-sig")
    else:
        df.to_csv(save_path, mode="w", header=True, index=False, encoding="utf-8-sig")

    print(f"完成！本次抓取 {len(df)} 只股票，已追加到 {save_path}")
    # 不同数据源返回的列名可能不完全一样，这里只挑实际存在的列展示，避免报错
    preferred_cols = ["抓取时间", "数据源", "代码", "名称", "最新价", "涨跌幅", "成交量", "换手率"]
    show_cols = [c for c in preferred_cols if c in df.columns]
    print(df[show_cols])

    export_json_for_dashboard(save_path)


def export_json_for_dashboard(csv_path: str):
    """把当天完整的CSV记录(可能有多次抓取)转成JSON，方便网页展示"""
    import json

    full_day_df = pd.read_csv(csv_path, encoding="utf-8-sig")
    records = json.loads(full_day_df.to_json(orient="records", force_ascii=False))
    payload = {
        "交易日期": beijing_now().strftime("%Y-%m-%d"),
        "生成时间": beijing_now().strftime("%Y-%m-%d %H:%M:%S"),
        "股票数量": len(CODE_SET),
        "数据": records,
    }
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"网页版竞价数据已生成：{JSON_FILE}（打开 auction_dashboard.html 查看）")


if __name__ == "__main__":
    main()
