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


def fill_order_data(main_df, df_order, forecast_months):
    """
    将订单数据按“订单日期”和“品名”聚合并填入 main_df 中每月的“订单”列。
    
    参数：
    - main_df: 主计划 DataFrame，需包含“品名”列
    - df_order: 上传的未交订单 DataFrame，包含“订单日期”和“品名”
    - forecast_months: 所有涉及的 yyyy-mm 字符串列表
    """
    df_order = df_order.copy()

    # 确保日期字段为 datetime 类型
    df_order["订单日期"] = pd.to_datetime(df_order["订单日期"], errors="coerce")
    df_order["年月"] = df_order["订单日期"].dt.to_period("M").astype(str)

    # 数值字段清洗
    df_order["未交订单数量"] = pd.to_numeric(df_order["未交订单数量"], errors="coerce").fillna(0)

    # 聚合出每品名每月的订单量
    grouped = (
        df_order.groupby(["品名", "年月"])["未交订单数量"]
        .sum()
        .unstack()
        .fillna(0)
    )

    for ym in forecast_months:
        colname = f"{ym}-订单"
        if colname in main_df.columns and ym in grouped.columns:
            main_df[colname] = main_df["品名"].map(grouped[ym]).fillna(0)

    return main_df

def fill_sales_data(main_df, df_sales, forecast_months):
    """
    将出货数据按“交易日期”和“品名”聚合并填入 main_df 中每月的“出货”列。
    
    参数：
    - main_df: 主计划 DataFrame，需包含“品名”列
    - df_sales: 出货明细 DataFrame，包含“交易日期”和“品名”
    - forecast_months: 所有涉及的 yyyy-mm 字符串列表
    """
    df_sales = df_sales.copy()

    # 确保交易日期为 datetime
    df_sales["交易日期"] = pd.to_datetime(df_sales["交易日期"], errors="coerce")
    df_sales["年月"] = df_sales["交易日期"].dt.to_period("M").astype(str)

    # 数值字段清洗
    df_sales["数量"] = pd.to_numeric(df_sales["数量"], errors="coerce").fillna(0)

    # 聚合：每品名每月出货数量
    grouped = (
        df_sales.groupby(["品名", "年月"])["数量"]
        .sum()
        .unstack()
        .fillna(0)
    )

    for ym in forecast_months:
        colname = f"{ym}-出货"
        if colname in main_df.columns and ym in grouped.columns:
            main_df[colname] = main_df["品名"].map(grouped[ym]).fillna(0)

    return main_df

def highlight_forecast_without_order(ws, forecast_months: list[str]):
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # 第二行是表头
    header = [cell.value for cell in ws[2]]

    for month in forecast_months:
        try:
            m = int(month.split("-")[1])
        except:
            continue
        forecast_col = f"{m}月预测"
        order_col = f"{m}月订单"

        try:
            col_f = header.index(forecast_col) + 1
            col_o = header.index(order_col) + 1
        except ValueError:
            continue

        for row in range(3, ws.max_row + 1):
            cell_f = ws.cell(row=row, column=col_f)
            cell_o = ws.cell(row=row, column=col_o)
            try:
                val_f = float(cell_f.value or 0)
                val_o = float(cell_o.value or 0)
            except:
                continue

            if val_f > 0 and val_o == 0:
                cell_f.fill = red_fill
                cell_o.fill = red_fill
