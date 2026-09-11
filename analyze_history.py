# -*- coding: utf-8 -*-
"""
历史数据分析脚本（增强版）
功能：读取 history_data 文件夹里的历史数据，计算：
      基础表现：5/20/60日涨跌幅、量能变化
      趋势判断：均线排列(多头/空头/震荡)、MACD金叉死叉、相对沪深300强弱
      建仓参考：RSI(14)、乖离率(距20日均线偏离度)
      风险参考：ATR止损参考幅度、综合趋势评分(0-5分)
      生成一份带颜色标注的 Excel 报表

重要说明（务必先看）：
      本脚本所有指标都基于公开、通用的技术分析公式计算，是对历史价格规律的
      客观描述，不构成任何投资建议，也不保证未来表现。是否建仓、仓位多少、
      止损设在哪，最终都需要你自己判断和承担风险。脚本作者/生成者不是持牌
      投资顾问。

用法：
    1. 先运行 fetch_history.py 更新数据（确保里面已经生成了 000300_沪深300.csv）
    2. 运行本脚本：
       python analyze_history.py
    3. 生成 分析报表.xlsx，双击打开即可
"""

import pandas as pd
import numpy as np
import os
import glob
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

HISTORY_DIR = "history_data"
OUTPUT_FILE = "分析报表.xlsx"
JSON_FILE = "选股数据.json"
BENCHMARK_CODE = "000300"


def load_all_files() -> dict:
    """读取 history_data 文件夹里所有csv，返回 {股票代码: DataFrame}（含基准指数）"""
    result = {}
    files = glob.glob(os.path.join(HISTORY_DIR, "*.csv"))
    for f in files:
        if "_合并历史数据" in f:
            continue
        df = pd.read_csv(f, encoding="utf-8-sig")
        if df.empty or "股票代码" not in df.columns:
            continue
        df["日期"] = pd.to_datetime(df["日期"])
        df = df.sort_values("日期").reset_index(drop=True)
        code = str(df["股票代码"].iloc[0]).zfill(6)
        result[code] = df
    return result


# ---------------- 技术指标计算函数 ----------------

def calc_ma(df, window):
    return df["收盘"].rolling(window).mean()


def calc_macd(close: pd.Series):
    """返回 DIF, DEA, MACD柱"""
    ema12 = close.ewm(span=12, adjust=False).mean()
    ema26 = close.ewm(span=26, adjust=False).mean()
    dif = ema12 - ema26
    dea = dif.ewm(span=9, adjust=False).mean()
    hist = (dif - dea) * 2
    return dif, dea, hist


def calc_rsi(close: pd.Series, period: int = 14):
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(period).mean()
    avg_loss = loss.rolling(period).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))
    return rsi


def calc_atr(df: pd.DataFrame, period: int = 14):
    prev_close = df["收盘"].shift(1)
    tr1 = df["最高"] - df["最低"]
    tr2 = (df["最高"] - prev_close).abs()
    tr3 = (df["最低"] - prev_close).abs()
    tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    return tr.rolling(period).mean()


# ---------------- 单只股票指标汇总 ----------------

def calc_metrics(df: pd.DataFrame, benchmark_df: pd.DataFrame) -> dict:
    if len(df) < 30:
        return None  # 数据不足30天，均线/MACD/RSI等指标意义不大，跳过

    latest = df.iloc[-1]
    name = df["股票名称"].iloc[-1]
    code = str(df["股票代码"].iloc[-1]).zfill(6)

    def pct_change(n):
        if len(df) <= n:
            return None
        old_close = df["收盘"].iloc[-(n + 1)]
        new_close = df["收盘"].iloc[-1]
        if not old_close or pd.isna(old_close):
            return None
        return round((new_close - old_close) / old_close * 100, 2)

    chg_5 = pct_change(5)
    chg_20 = pct_change(20)
    chg_60 = pct_change(60)

    # 量能变化
    vol_ratio = None
    if "成交量" in df.columns and len(df) >= 25:
        recent_5_vol = df["成交量"].iloc[-5:].mean()
        prior_20_vol = df["成交量"].iloc[-25:-5].mean()
        if prior_20_vol and prior_20_vol > 0:
            vol_ratio = round(recent_5_vol / prior_20_vol, 2)

    # 均线 & 排列判断
    ma5 = calc_ma(df, 5).iloc[-1]
    ma20 = calc_ma(df, 20).iloc[-1]
    ma60 = calc_ma(df, 60).iloc[-1] if len(df) >= 60 else None

    if ma60 is not None and not pd.isna(ma60):
        if ma5 > ma20 > ma60:
            ma_trend = "多头排列"
        elif ma5 < ma20 < ma60:
            ma_trend = "空头排列"
        else:
            ma_trend = "震荡交织"
    else:
        ma_trend = "数据不足(需60日+)"

    above_ma20 = "是" if latest["收盘"] > ma20 else "否"

    # 乖离率（距20日均线偏离度）
    bias20 = round((latest["收盘"] - ma20) / ma20 * 100, 2) if ma20 else None

    # MACD
    dif, dea, hist = calc_macd(df["收盘"])
    macd_status = "数据不足"
    if len(df) >= 35:
        if hist.iloc[-2] < 0 and hist.iloc[-1] > 0:
            macd_status = "金叉(转多)"
        elif hist.iloc[-2] > 0 and hist.iloc[-1] < 0:
            macd_status = "死叉(转空)"
        elif dif.iloc[-1] > dea.iloc[-1]:
            macd_status = "多头持续"
        else:
            macd_status = "空头持续"

    # RSI
    rsi = calc_rsi(df["收盘"]).iloc[-1]
    rsi = round(rsi, 1) if not pd.isna(rsi) else None
    if rsi is not None:
        if rsi >= 70:
            rsi_flag = "超买区间"
        elif rsi <= 30:
            rsi_flag = "超卖区间"
        else:
            rsi_flag = "中性"
    else:
        rsi_flag = "数据不足"

    # ATR 止损参考幅度（2倍ATR法，是常见的技术止损参考方法之一）
    atr = calc_atr(df).iloc[-1]
    stop_loss_pct = None
    stop_loss_price = None
    if atr and not pd.isna(atr) and latest["收盘"] > 0:
        stop_loss_pct = round((2 * atr) / latest["收盘"] * 100, 2)
        stop_loss_price = round(latest["收盘"] - 2 * atr, 2)

    # 60日区间高低点，给网页版画"价格区间位置条"用
    lookback60 = min(60, len(df))
    high_60 = round(df["最高"].iloc[-lookback60:].max(), 2)
    low_60 = round(df["最低"].iloc[-lookback60:].min(), 2)

    # 相对强弱（跟沪深300比，20日区间涨跌幅之差）
    rel_strength = None
    if benchmark_df is not None and len(benchmark_df) > 20 and chg_20 is not None:
        bench_old = benchmark_df["收盘"].iloc[-21]
        bench_new = benchmark_df["收盘"].iloc[-1]
        if bench_old and bench_old > 0:
            bench_chg_20 = (bench_new - bench_old) / bench_old * 100
            rel_strength = round(chg_20 - bench_chg_20, 2)

    # 综合趋势评分（0-5分，越高说明多个信号同时偏多，仅供参考，不是买卖信号）
    score = 0
    if ma_trend == "多头排列":
        score += 1
    if macd_status in ("金叉(转多)", "多头持续"):
        score += 1
    if rsi_flag == "中性":
        score += 1
    if above_ma20 == "是":
        score += 1
    if rel_strength is not None and rel_strength > 0:
        score += 1

    return {
        "股票代码": code,
        "股票名称": name,
        "最新收盘": latest["收盘"],
        "5日涨跌幅%": chg_5,
        "20日涨跌幅%": chg_20,
        "60日涨跌幅%": chg_60,
        "量能比(5日/前20日)": vol_ratio,
        "均线排列": ma_trend,
        "是否站上20日线": above_ma20,
        "乖离率%(距20日线)": bias20,
        "MACD状态": macd_status,
        "RSI(14)": rsi,
        "RSI区间": rsi_flag,
        "相对沪深300强弱%(20日)": rel_strength,
        "ATR止损参考幅度%": stop_loss_pct,
        "止损参考价": stop_loss_price,
        "20日均线值": round(ma20, 2) if ma20 and not pd.isna(ma20) else None,
        "60日最高": high_60,
        "60日最低": low_60,
        "综合趋势评分(0-5)": score,
        "最新日期": latest["日期"].strftime("%Y-%m-%d"),
    }


def prepare_ranked(result_df: pd.DataFrame) -> pd.DataFrame:
    """按综合评分+5日涨幅排序，并加上排名列，Excel和网页版JSON共用这份结果"""
    df = result_df.sort_values(
        ["综合趋势评分(0-5)", "5日涨跌幅%"], ascending=[False, False], na_position="last"
    ).reset_index(drop=True)
    df.insert(0, "综合排名", range(1, len(df) + 1))
    return df


def export_json(result_df: pd.DataFrame, json_path: str):
    """导出给网页版仪表盘用的JSON数据文件"""
    import json
    from datetime import datetime

    records = json.loads(result_df.to_json(orient="records", force_ascii=False))
    payload = {
        "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "股票数量": len(records),
        "数据": records,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def style_and_save(result_df: pd.DataFrame):
    result_df.to_excel(OUTPUT_FILE, index=False, sheet_name="选股分析")

    from openpyxl import load_workbook
    wb = load_workbook(OUTPUT_FILE)
    ws = wb["选股分析"]

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    header_font = Font(bold=True)

    for cell in ws[1]:
        cell.font = header_font

    col_index_map = {cell.value: cell.column for cell in ws[1]}

    # 数值型涨跌类列：正数红、负数绿
    numeric_color_cols = [
        "5日涨跌幅%", "20日涨跌幅%", "60日涨跌幅%",
        "乖离率%(距20日线)", "相对沪深300强弱%(20日)",
    ]
    for col_name in numeric_color_cols:
        if col_name not in col_index_map:
            continue
        col_letter = get_column_letter(col_index_map[col_name])
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row}"]
            if isinstance(cell.value, (int, float)):
                cell.fill = red_fill if cell.value > 0 else (green_fill if cell.value < 0 else cell.fill)

    # 文字状态列：多头/金叉标红，空头/死叉标绿
    text_color_cols = ["均线排列", "MACD状态"]
    bullish_words = ["多头", "金叉"]
    bearish_words = ["空头", "死叉"]
    for col_name in text_color_cols:
        if col_name not in col_index_map:
            continue
        col_letter = get_column_letter(col_index_map[col_name])
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row}"]
            val = str(cell.value)
            if any(w in val for w in bullish_words):
                cell.fill = red_fill
            elif any(w in val for w in bearish_words):
                cell.fill = green_fill

    # RSI区间：超买/超卖标黄提醒关注
    if "RSI区间" in col_index_map:
        col_letter = get_column_letter(col_index_map["RSI区间"])
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row}"]
            if cell.value in ("超买区间", "超卖区间"):
                cell.fill = yellow_fill

    # 综合评分列：4-5分标红，0-1分标绿
    if "综合趋势评分(0-5)" in col_index_map:
        col_letter = get_column_letter(col_index_map["综合趋势评分(0-5)"])
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row}"]
            if isinstance(cell.value, (int, float)):
                if cell.value >= 4:
                    cell.fill = red_fill
                elif cell.value <= 1:
                    cell.fill = green_fill

    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    # 加一个说明sheet，把免责声明和指标含义写清楚
    ws2 = wb.create_sheet("指标说明")
    notes = [
        ["重要说明", "本报表所有指标基于公开技术分析公式计算，是历史价格的客观统计描述，"
                     "不构成投资建议，不保证未来表现。是否建仓、仓位大小、止损位设置，"
                     "最终需要你自己判断和承担风险。"],
        ["", ""],
        ["指标", "含义"],
        ["5/20/60日涨跌幅", "近N个交易日收盘价的涨跌百分比"],
        ["量能比", "近5日平均成交量 相比 之前20日平均成交量 的倍数，大于1说明近期在放量"],
        ["均线排列", "5/20/60日均线的排列方向，多头排列=短期在上长期在下且趋势向上"],
        ["是否站上20日线", "最新收盘价是否在20日均线上方"],
        ["乖离率", "最新收盘价偏离20日均线的百分比，绝对值越大代表离均线越远，可能有回归需求"],
        ["MACD状态", "基于DIF/DEA判断的金叉/死叉/多空持续状态"],
        ["RSI(14)", "14日相对强弱指标，>=70通常视为超买，<=30通常视为超卖"],
        ["相对沪深300强弱", "该股近20日涨跌幅 减去 沪深300同期涨跌幅，正数代表跑赢大盘"],
        ["ATR止损参考幅度", "基于近14日真实波动幅度(ATR)的2倍，计算出的止损参考百分比，"
                          "是常见的技术止损方法之一，不是必须遵循的规则"],
        ["综合趋势评分", "0-5分，统计多头排列/MACD偏多/RSI中性/站上20日线/跑赢大盘 这5项"
                       "各是否成立，分数越高代表当前技术面信号偏多，仅供参考"],
    ]
    for row in notes:
        ws2.append(row)
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 90
    for cell in ws2["A"]:
        cell.font = Font(bold=True)
    ws2["A1"].fill = yellow_fill
    ws2["B1"].fill = yellow_fill

    wb.save(OUTPUT_FILE)


def main():
    print("正在读取 history_data 里的数据...")
    all_files = load_all_files()

    benchmark_df = all_files.pop(BENCHMARK_CODE, None)
    if benchmark_df is None:
        print("提示：没有找到沪深300数据(000300_沪深300.csv)，相对强弱这项会缺失。")
        print("      可以重新运行 fetch_history.py 来补充抓取。")

    if not all_files:
        print("没有找到任何个股历史数据，请先运行 fetch_history.py")
        return

    print(f"共读取到 {len(all_files)} 只股票，正在计算技术指标...")
    rows = []
    for code, df in all_files.items():
        metrics = calc_metrics(df, benchmark_df)
        if metrics:
            rows.append(metrics)
        else:
            print(f"  {code} 数据不足30天，跳过（建议数据积累更长时间后再分析）")

    if not rows:
        print("数据不足，无法计算指标")
        return

    result_df = pd.DataFrame(rows)
    ranked_df = prepare_ranked(result_df)
    style_and_save(ranked_df)
    export_json(ranked_df, JSON_FILE)

    print(f"\n完成！报表已生成：{OUTPUT_FILE}")
    print(f"网页版数据已生成：{JSON_FILE}（打开 dashboard.html 后选择这个文件即可查看）")
    print("已按 综合趋势评分 + 5日涨跌幅 排序，红色=偏多信号，绿色=偏空信号，黄色=超买超卖提醒")
    print("详细指标含义见 Excel 里的\"指标说明\"sheet")
    print("\n预览评分前5名：")
    cols = ["股票名称", "综合趋势评分(0-5)", "均线排列", "MACD状态", "RSI区间", "相对沪深300强弱%(20日)"]
    print(ranked_df.head(5)[cols])


if __name__ == "__main__":
    main()
