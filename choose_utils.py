import streamlit as st
import pandas as pd
import io
from datetime import datetime

def show_raw_data_filter_ui(forecast_df, order_df, sales_df, forecast_months):
    st.markdown("### 🔍 原始数据查看")

    # 品名选项来自预测、订单、出货三表的合集
    all_names = sorted(set(forecast_df["生产料号"]) | set(order_df["品名"]) | set(sales_df["品名"]))
    selected_name = st.selectbox("📦 选择品名", all_names)
    selected_month = st.selectbox("📅 选择月份", forecast_months)

    # 转换月份为起止日期
    month_start = pd.to_datetime(f"{selected_month}-01")
    month_end = (month_start + pd.offsets.MonthEnd(0))

    # ====== 过滤每张表 ======

    df_forecast_filtered = forecast_df[forecast_df["生产料号"] == selected_name]

    df_order_filtered = order_df.copy()
    df_order_filtered["客户要求交期"] = pd.to_datetime(df_order_filtered["客户要求交期"], errors="coerce")
    df_order_filtered = df_order_filtered[
        (df_order_filtered["品名"] == selected_name) &
        (df_order_filtered["客户要求交期"].between(month_start, month_end))
    ]

    df_sales_filtered = sales_df.copy()
    df_sales_filtered["交易日期"] = pd.to_datetime(df_sales_filtered["交易日期"], errors="coerce")
    df_sales_filtered = df_sales_filtered[
        (df_sales_filtered["品名"] == selected_name) &
        (df_sales_filtered["交易日期"].between(month_start, month_end))
    ]

    # ====== 显示结果 ======
    st.subheader("🔹 预测数据")
    st.dataframe(df_forecast_filtered)

    st.subheader("🔹 订单数据")
    st.dataframe(df_order_filtered)

    st.subheader("🔹 出货数据")
    st.dataframe(df_sales_filtered)

    # ====== 下载按钮 ======
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_forecast_filtered.to_excel(writer, index=False, sheet_name="预测")
        df_order_filtered.to_excel(writer, index=False, sheet_name="订单")
        df_sales_filtered.to_excel(writer, index=False, sheet_name="出货")
    output_buffer.seek(0)

    st.download_button(
        label="📥 下载该品名原始记录",
        data=output_buffer,
        file_name=f"{selected_name}_{selected_month}_原始记录.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
