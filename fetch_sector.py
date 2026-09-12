# -*- coding: utf-8 -*-
"""
板块数据抓取脚本
功能：
    1. 获取每只自选股所属的申万行业分类
    2. 获取该行业指数的历史每日涨跌数据
    这样就能在分析里算出"个股所属板块最近涨得怎么样"这个维度

数据来源：申万宏源研究(swsresearch.com)，跟你股票池的历史数据(腾讯接口)是不同的数据源，
互不影响，即使其中一个访问不了，也不影响另一个。

用法：
    python fetch_sector.py
    跑完会生成 sector_data 文件夹，里面有：
    - stock_industry_map.csv：每只股票对应的申万行业代码(和名称，如果抓得到的话)
    - {行业代码}.csv：该行业指数的历史每日行情

建议：板块归属不常变化，这个脚本不用每天跑，一两周跑一次更新一下就够了。
"""

import akshare as ak
import pandas as pd
import os
import re
import time
import io
import requests
import urllib3
from stock_list import STOCK_POOL

# 申万网站的SSL证书在某些云端Linux环境下会触发证书校验失败(跟akshare本身无关，
# 是对方网站证书链的兼容性问题)，这里关闭校验规避一下，同时关掉多余的警告输出
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

OUTPUT_DIR = "sector_data"
CODE_SET = {s["code"] for s in STOCK_POOL}
SW_CLASSIFY_URL = "https://www.swsresearch.com/swindex/pdf/SwClass2021/StockClassifyUse_stock.xls"


def fetch_sw_classification_raw() -> pd.DataFrame:
    """
    直接请求申万行业分类原始文件，跳过证书校验(verify=False)
    效果等同于 ak.stock_industry_clf_hist_sw()，但绕开了它在部分云端环境下的SSL报错
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    }
    r = requests.get(SW_CLASSIFY_URL, headers=headers, verify=False, timeout=30)
    r.raise_for_status()
    raw = pd.read_excel(io.BytesIO(r.content), dtype={"股票代码": "str", "行业代码": "str"})
    raw = raw.rename(columns={
        "股票代码": "symbol", "计入日期": "start_date",
        "行业代码": "industry_code", "更新日期": "update_time",
    })
    return raw


def build_stock_industry_map() -> pd.DataFrame:
    """获取全市场股票的申万行业分类，筛选出我们股票池里的，每只股票只保留最新一条分类"""
    print("正在获取申万行业分类数据(全市场，文件较大，可能需要几秒到几十秒)...")
    raw = fetch_sw_classification_raw()
    print(f"  获取到 {len(raw)} 条分类记录，字段包括：{list(raw.columns)}")

    # 股票代码格式不确定(可能带.SZ/.SH后缀)，统一提取6位数字再匹配
    raw = raw.copy()
    raw["代码_纯数字"] = raw["symbol"].astype(str).str.extract(r"(\d{6})")
    filtered = raw[raw["代码_纯数字"].isin(CODE_SET)].copy()

    if filtered.empty:
        print("  !! 没有匹配到股票池里的任何股票，请检查数据源是否正常")
        return pd.DataFrame()

    # 每只股票可能有多条历史分类记录(行业分类会变动)，只保留最新的一条
    sort_col = "update_time" if "update_time" in filtered.columns else "start_date"
    filtered = filtered.sort_values(sort_col).groupby("代码_纯数字", as_index=False).last()

    # 尝试自动找一个"行业名称"相关的列，方便展示；找不到就用代码代替
    name_col = None
    for col in filtered.columns:
        if "名称" in col or "行业" in col and col != "industry_code":
            if col not in ("代码_纯数字",):
                name_col = col
                break

    result = pd.DataFrame({
        "股票代码": filtered["代码_纯数字"],
        "行业代码": filtered["industry_code"],
    })
    result["行业名称"] = filtered[name_col] if name_col else filtered["industry_code"]

    return result


def fetch_sector_history(industry_code: str, max_retries: int = 3) -> pd.DataFrame:
    """抓取某个申万行业指数的历史日线数据"""
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            df = ak.index_hist_sw(symbol=industry_code, period="day")
            return df
        except Exception as e:
            last_error = e
            if attempt < max_retries:
                time.sleep(2 * attempt)
    raise last_error


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    mapping = build_stock_industry_map()
    if mapping.empty:
        print("未能获取行业分类映射，终止。")
        return

    mapping_path = os.path.join(OUTPUT_DIR, "stock_industry_map.csv")
    mapping.to_csv(mapping_path, index=False, encoding="utf-8-sig")
    print(f"股票-行业映射已保存到 {mapping_path}")

    unique_codes = mapping["行业代码"].dropna().unique().tolist()
    print(f"\n共涉及 {len(unique_codes)} 个申万行业，开始抓取各行业历史指数...")

    for i, code in enumerate(unique_codes, start=1):
        name_row = mapping[mapping["行业代码"] == code].iloc[0]
        print(f"[{i}/{len(unique_codes)}] 正在抓取行业指数 {name_row['行业名称']}({code}) ...")
        try:
            df = fetch_sector_history(code)
            save_path = os.path.join(OUTPUT_DIR, f"{code}.csv")
            df.to_csv(save_path, index=False, encoding="utf-8-sig")
        except Exception as e:
            print(f"  !! 抓取失败: {e}")
        time.sleep(1)

    print("\n完成！可以运行 analyze_history.py 生成包含板块因子的分析报表了。")


if __name__ == "__main__":
    main()
