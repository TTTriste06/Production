import pandas as pd
import streamlit as st
from datetime import timedelta

from datetime import timedelta
import pandas as pd

def schedule_production(df_info: pd.DataFrame, date_columns: list) -> pd.DataFrame:
    """
    按照封装厂 + 封装形式的分配产能，安排需排产数量到各天的日期列。

    参数:
    - df_info: 原始订单数据，必须包含 ['封装厂', '封装形式', '需排产', '分配产能', '实际开始测试日期']
    - date_columns: 排产日历列名（格式为 yyyy/mm/dd 的连续日期）

    返回:
    - 更新后的 df_info，日期列写入每日排产数量
    """
    st.write(df_info)
    
    df = df_info.copy()
    df[date_columns] = 0  # 初始化所有日期列为 0

    # 封装组合的日产能
    group_capacity = df.groupby(["封装厂", "封装形式"])["分配产能"].first().to_dict()

    # 每个封装组合每天的已使用产能
    used_capacity = {}

    # 按是否指定“实际开始测试日期”拆分并排序
    df_fixed = df[df["实际开始测试日期"].notna()].copy()
    df_fixed["实际开始测试日期"] = pd.to_datetime(df_fixed["实际开始测试日期"])

    df_unspecified = df[df["实际开始测试日期"].isna()].copy()
    df_unspecified["实际开始测试日期"] = pd.NaT  # 显式设为缺失时间格式，便于统一处理

    df_sorted = pd.concat([df_fixed, df_unspecified], ignore_index=True)

    for idx, row in df_sorted.iterrows():
        key = (row["封装厂"], row["封装形式"])
        total_qty = int(row["需排产"])
        daily_cap = int(group_capacity.get(key, 0))
        used_capacity.setdefault(key, {})

        # 起始排产日期
        if pd.notna(row["实际开始测试日期"]):
            current_date = row["实际开始测试日期"]
        else:
            current_date = pd.to_datetime(date_columns[1])  # 默认从首日开始

        produced = 0

        while produced < total_qty:
            date_str = current_date.strftime("%Y/%m/%d")
            if date_str not in date_columns:
                break  # 超出排产日历范围

            used_today = used_capacity[key].get(date_str, 0)
            available = daily_cap - used_today
            assign_qty = min(available, total_qty - produced)

            if assign_qty > 0:
                # 找到原始 df 中的对应行进行写入
                df.loc[(df["封装厂"] == row["封装厂"]) &
                       (df["封装形式"] == row["封装形式"]) &
                       (df["需排产"] == row["需排产"]) &
                       (df["实际开始测试日期"] == row["实际开始测试日期"]), date_str] += assign_qty

                used_capacity[key][date_str] = used_today + assign_qty
                produced += assign_qty

            current_date += timedelta(days=1)

    return df
