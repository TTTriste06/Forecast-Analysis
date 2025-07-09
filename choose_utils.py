import streamlit as st
import pandas as pd
import io
from datetime import datetime

def show_raw_data_filter_ui(forecast_file, order_file, sales_file, forecast_months):
    st.markdown("### 🔍 按品名 + 月份查看原始数据")

    # 品名集合：来自三张表中字段
    all_names = sorted(
        set(forecast_file["生产料号"].dropna().astype(str)) |
        set(order_file["品名"].dropna().astype(str)) |
        set(sales_file["品名"].dropna().astype(str))
    )

    selected_name = st.selectbox("📦 选择品名", all_names)
    selected_month = st.selectbox("📅 选择月份", forecast_months)

    # 时间范围
    month_start = pd.to_datetime(f"{selected_month}-01")
    month_end = month_start + pd.offsets.MonthEnd(0)

    # 过滤预测（仅按品名）
    df_forecast_filtered = forecast_file[forecast_file["生产料号"] == selected_name]

    # 过滤订单（品名 + 客户要求交期）
    df_order_filtered = order_file.copy()
    if "客户要求交期" in df_order_filtered.columns:
        df_order_filtered["客户要求交期"] = pd.to_datetime(df_order_filtered["客户要求交期"], errors="coerce")
        df_order_filtered = df_order_filtered[
            (df_order_filtered["品名"] == selected_name) &
            (df_order_filtered["客户要求交期"].between(month_start, month_end))
        ]
    else:
        df_order_filtered = pd.DataFrame(columns=order_file.columns)  # 空表防报错

    # 过滤出货（品名 + 交易日期）
    df_sales_filtered = sales_file.copy()
    if "交易日期" in df_sales_filtered.columns:
        df_sales_filtered["交易日期"] = pd.to_datetime(df_sales_filtered["交易日期"], errors="coerce")
        df_sales_filtered = df_sales_filtered[
            (df_sales_filtered["品名"] == selected_name) &
            (df_sales_filtered["交易日期"].between(month_start, month_end))
        ]
    else:
        df_sales_filtered = pd.DataFrame(columns=sales_file.columns)

    # 显示结果
    st.subheader("📈 预测数据")
    st.dataframe(df_forecast_filtered, use_container_width=True)

    st.subheader("📦 订单数据")
    st.dataframe(df_order_filtered, use_container_width=True)

    st.subheader("🚚 出货数据")
    st.dataframe(df_sales_filtered, use_container_width=True)

    # 下载按钮
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_forecast_filtered.to_excel(writer, index=False, sheet_name="预测")
        df_order_filtered.to_excel(writer, index=False, sheet_name="订单")
        df_sales_filtered.to_excel(writer, index=False, sheet_name="出货")
    buffer.seek(0)

    st.download_button(
        "📥 下载该品名原始记录",
        data=buffer,
        file_name=f"{selected_name}_{selected_month}_原始记录.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
