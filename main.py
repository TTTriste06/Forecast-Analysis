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
    
    template_file, forecast_file, order_file, sales_file, mapping_file, start = get_uploaded_files()
    
    if start:
        template_df = load_file_with_github_fallback("template", template_file, sheet_name=0, header=1)
        forecast_df = load_file_with_github_fallback("forecast", forecast_file, sheet_name="预测")
        order_df = load_file_with_github_fallback("order", order_file, sheet_name="Sheet")
        sales_df = load_file_with_github_fallback("sales", sales_file, sheet_name="原表")
        mapping_df = load_file_with_github_fallback("mapping", mapping_file, sheet_name=0)

    
        processor = PivotProcessor()
        df_result, excel_output = processor.process(template_file, forecast_file, order_file, sales_file, mapping_file)
    
        st.success("✅ 主计划生成成功！")
        st.dataframe(df_result, use_container_width=True)
    
        st.download_button(
            label="📥 下载主计划 Excel 文件",
            data=excel_output.getvalue(),
            file_name=f"预测分析主计划_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("❌ Streamlit app crashed:", e)
        traceback.print_exc()
