import re
import pandas as pd

def extract_all_year_months(df_forecast, df_order, df_sales):
    # 1. 从 forecast header 提取 x月预测 列中的月份
    month_pattern = re.compile(r"(\d{1,2})月预测")
    forecast_months = []
    for col in df_forecast.columns:
        match = month_pattern.match(str(col))
        if match:
            month = match.group(1).zfill(2)
            forecast_months.append(f"2025-{month}")  # ✅ 根据需要调整年份

    # 2. 从 order 文件第 B 列（假设是“订单日期”）
    order_date_col = df_order.columns[1]
    df_order[order_date_col] = pd.to_datetime(df_order[order_date_col], errors="coerce")
    order_months = (
        df_order[order_date_col]
        .dropna()
        .dt.to_period("M")
        .astype(str)
        .loc[lambda x: x != "NaT"]
        .unique()
        .tolist()
    )

    # 3. 从 sales 文件第 F 列（假设是“交易日期”）
    sales_date_col = df_sales.columns[5]
    df_sales[sales_date_col] = pd.to_datetime(df_sales[sales_date_col], errors="coerce")
    sales_months = (
        df_sales[sales_date_col]
        .dropna()
        .dt.to_period("M")
        .astype(str)
        .loc[lambda x: x != "NaT"]
        .unique()
        .tolist()
    )

    # 合并并去重
    all_months = sorted(set(forecast_months + order_months + sales_months))

    # 生成从最小到最大之间的所有月份
    if all_months:
        min_month = pd.Period(min(all_months), freq="M")
        max_month = pd.Period(max(all_months), freq="M")
        full_months = [str(p) for p in pd.period_range(min_month, max_month, freq="M")]
    else:
        full_months = []

    return full_months

def fill_forecast_data(main_df, df_forecast, forecast_months):
    """
    从 forecast_file 填入预测数据，按生产料号对应品名，支持月份列为“6月预测”格式。
    """
    # 清洗生产料号 → 品名
    df_forecast["生产料号"] = df_forecast["生产料号"].astype(str).str.strip()
    df_forecast["品名"] = df_forecast["生产料号"]

    # 正则提取“x月预测”字段
    month_pattern = re.compile(r"(\d{1,2})月预测")
    forecast_cols = {
        f"2025-{match.group(1).zfill(2)}": col
        for col in df_forecast.columns
        if (match := month_pattern.match(str(col)))
    }

    for ym in forecast_months:
        colname = f"{ym}-预测"
        if colname in main_df.columns and ym in forecast_cols:
            month_col = forecast_cols[ym]

            # ✅ 按品名汇总（避免重复索引）
            forecast_series = (
                df_forecast.groupby("品名")[month_col]
                .sum(min_count=1)
            )

            main_df[colname] = main_df["品名"].map(forecast_series).fillna(0)

    return main_df
