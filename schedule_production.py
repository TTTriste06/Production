import pandas as pd
from datetime import timedelta

def schedule_production(df_info: pd.DataFrame, date_columns: list) -> pd.DataFrame:
    """
    按照封装厂 + 封装形式的分配产能，安排需排产数量到各天的日期列。

    参数:
    - df_info: 原始订单数据，必须包含 ['封装厂', '封装形式', '需排产', '分配产能', '实际开始测试日期']
    - date_columns: 已在 write_calendar_headers 中写入的列名，顺序为连续日期字符串 yyyy/mm/dd

    返回：
    - 包含新增日期列的 df_info，每个订单的排产情况写入对应日期列
    """
    df = df_info.copy()
    df[date_columns] = 0  # 初始化每天排产量为0

    # 分组资源日容量：封装厂 + 封装形式 -> 日产能
    group_capacity = (
        df.groupby(["封装厂", "封装形式"])["分配产能"].first().to_dict()
    )

    # 每个组合每天的已使用产能：key = (封装厂, 封装形式) -> {日期: 已用产能}
    used_capacity = {}

    # 拆成有“实际开始测试日”和无的两部分，确保优先排定日期的
    df_fixed = df[df["实际开始测试日期"].notna()].copy()
    df_fixed["实际开始测试日期"] = pd.to_datetime(df_fixed["实际开始测试日期"])

    df_unspecified = df[df["实际开始测试日期"].isna()].copy()

    # 合并顺序列表
    df_sorted = pd.concat([df_fixed, df_unspecified], ignore_index=True)

    for idx, row in df_sorted.iterrows():
        key = (row["封装厂"], row["封装形式"])
        total_to_produce = int(row["需排产"])
        daily_capacity = int(group_capacity.get(key, 0))
        used_capacity.setdefault(key, {})

        # 起始排产日
        if pd.notna(row["实际开始测试日期"]):
            current_date = row["实际开始测试日期"]
        else:
            current_date = pd.to_datetime(date_columns[0])  # 用排产起始日期

        produced = 0

        while produced < total_to_produce:
            date_str = current_date.strftime("%Y/%m/%d")

            # 如果日期超出了生成的日期列，就中断
            if date_str not in date_columns:
                break

            used_today = used_capacity[key].get(date_str, 0)
            remaining_today = daily_capacity - used_today
            need_today = min(remaining_today, total_to_produce - produced)

            if need_today > 0:
                # 在原始 df 中写入排产数量
                df.loc[idx, date_str] = need_today
                # 更新已用产能
                used_capacity[key][date_str] = used_today + need_today
                produced += need_today

            # 如果当天无空间，或还需生产，继续下一天
            current_date += timedelta(days=1)

    return df
