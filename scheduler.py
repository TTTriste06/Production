import pandas as pd
from datetime import timedelta

def schedule_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # 检查产能字段
    if df["分配产能"].isnull().any():
        raise ValueError("部分产品缺少“分配产能”字段，请检查原始数据！")

    # 填充日期列
    df["wafer in"] = pd.to_datetime(df["wafer in"])
    df["实际开始测试日期"] = pd.to_datetime(df["实际开始测试日期"], errors='coerce')

    def compute_start_date(row):
        standard_start = row["wafer in"] + timedelta(days=int(row["排产周期"]) + int(row["磨划周期"]) + int(row["封装周期"]))
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
        product, total_qty = row["产品"], int(row["订单数"])
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
        out_row["11月产出合计"] = total_qty  # 可改成动态字段
        out_row["建议启动日期"] = min(daily_output.keys())
        records.append(out_row)

    return pd.DataFrame(records)
