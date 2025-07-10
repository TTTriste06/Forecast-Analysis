import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from ui import get_uploaded_files
from pivot_processor import PivotProcessor
from github_utils import load_file_with_github_fallback

def main():
    st.set_page_config(page_title="预测分析主计划工具", layout="wide")
    st.title("📊 预测分析主计划生成器")

    # 获取上传的文件
    template_file, forecast_file, order_file, sales_file, mapping_file, start = get_uploaded_files()

    if start:
        # 加载数据
        template_df = load_file_with_github_fallback("template", template_file, sheet_name=0, header=1)
        forecast_df = load_file_with_github_fallback("forecast", forecast_file, sheet_name="预测")
        order_df = load_file_with_github_fallback("order", order_file, sheet_name="Sheet")
        sales_df = load_file_with_github_fallback("sales", sales_file, sheet_name="原表")
        mapping_df = load_file_with_github_fallback("mapping", mapping_file, sheet_name=0)

        # 生成主计划
        processor = PivotProcessor()
        df_result, excel_output = processor.process(template_df, forecast_df, order_df, sales_df, mapping_df)

        st.success("✅ 主计划生成成功！")

        # 显示汇总表
        st.markdown("## 📊 主计划汇总表")
        st.dataframe(df_result)

        # 设置点击跳转逻辑
        st.markdown("---")
        st.markdown("## 🔍 查看对应原始数据")

        if "selected_item" not in st.session_state:
            st.session_state.selected_item = None

        items = df_result["品名"].dropna().unique().tolist()
        selected = st.selectbox("请选择品名查看其原始数据：", [""] + items)

        if selected:
            st.session_state.selected_item = selected

        if st.session_state.selected_item:
            st.markdown(f"### 🔎 原始数据：{st.session_state.selected_item}")

            # 过滤原始数据
            filtered_forecast = forecast_df[forecast_df["生产料号"] == st.session_state.selected_item]
            filtered_order = order_df[order_df["品名"] == st.session_state.selected_item]
            filtered_sales = sales_df[sales_df["品名"] == st.session_state.selected_item]

            if not filtered_forecast.empty:
                st.markdown("#### 📈 预测数据")
                st.dataframe(filtered_forecast)

            if not filtered_order.empty:
                st.markdown("#### 📦 未交订单")
                st.dataframe(filtered_order)

            if not filtered_sales.empty:
                st.markdown("#### 💰 销售明细")
                st.dataframe(filtered_sales)

if __name__ == "__main__":
    main()
