import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from ui import setup_sidebar, get_uploaded_files
from pivot_processor import PivotProcessor

st.set_page_config(page_title="预测分析主计划工具", layout="wide")
st.title("📊 预测分析主计划生成器")

setup_sidebar()
template_file, forecast_file, order_file, sales_file, start = get_uploaded_files()

if start:
    if not all([template_file, forecast_file, order_file, sales_file]):
        st.error("❌ 请上传所有所需文件")
        st.stop()

    processor = PivotProcessor()
    df_result, excel_output = processor.process(template_file, forecast_file, order_file, sales_file)

    st.success("✅ 主计划生成成功！")
    st.dataframe(df_result, use_container_width=True)

    st.download_button(
        label="📥 下载主计划 Excel 文件",
        data=excel_output.getvalue(),
        file_name=f"预测分析主计划_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
