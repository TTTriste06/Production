import pandas as pd
import streamlit as st
from datetime import timedelta

def convert_excel_date(val):
    if pd.isnull(val):
        return pd.NaT
    try:
        if isinstance(val, str) and val.strip().isdigit():
            val = float(val)
        if isinstance(val, (int, float)):
            return pd.to_datetime("1899-12-30") + pd.to_timedelta(val, unit="D")
        return pd.to_datetime(val, errors="coerce")
    except:
        return pd.NaT

def schedule_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if df["分配产能"].isnull().any():
        raise ValueError("部分产品缺少“分配产能”字段，请检查原始数据！")

    # 转换日期字段，强制要求实际开始测试日期非空
    df["waferin"] = pd.to_datetime(df["waferin"], errors='coerce')
    df["实际开始测试日期"] = df["实际开始测试日期"].apply(convert_excel_date)
    st.write(df["实际开始测试日期"])

    def compute_start_date(row):
        standard_start = row["waferin"] + timedelta(days=int(row["排产周期"]) + int(row["磨划周期"]) + int(row["封装周期"]))
        if pd.notnull(row["实际开始测试日期"]) and row["实际开始测试日期"] < standard_start:
            return row["实际开始测试日期"]
        return standard_start

    df["排产起始日"] = df.apply(compute_start_date, axis=1)

    # 排序处理：按排产起始日从早到晚
    df.sort_values("排产起始日", inplace=True)

    # 模拟产能分配过程
    records = []
    capacity_tracker = {}  # {(封装厂, 封装形式, 日期): 已使用产能}

    for _, row in df.iterrows():
        order_id = row["订单号"]
        total_qty = int(row["需排产"])
        group = (row["封装厂"], row["封装形式"])
        max_daily = int(row["分配产能"])

        date = row["排产起始日"]
        remain = total_qty
        daily_output = {}

        while remain > 0:
            key = (group[0], group[1], date)
            used = capacity_tracker.get(key, 0)
            available = max_daily - used

            if available > 0:
                assign = min(available, remain)
                daily_output[date] = assign
                remain -= assign
                capacity_tracker[key] = used + assign

            date += timedelta(days=1)

        # 记录产出
        out_row = row.to_dict()
        for d, v in daily_output.items():
            out_row[d.strftime("%Y-%m-%d")] = v
        out_row["建议启动日期"] = min(daily_output.keys())

        # 将"建议启动日期"插入到"排产起始日"之后
        reordered = {}
        for k in out_row:
            reordered[k] = out_row[k]
            if k == "排产起始日":
                reordered["建议启动日期"] = out_row["建议启动日期"]
                break
        for k in out_row:
            if k not in reordered:
                reordered[k] = out_row[k]

        records.append(reordered)

    return pd.DataFrame(records)
