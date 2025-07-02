import streamlit as st
import pandas as pd
from dateutil.relativedelta import relativedelta
from datetime import date
from datetime import datetime

def setup_sidebar():
    with st.sidebar:
        st.title("功能简介")
        st.markdown("---")
        st.markdown("- 预测、未交订单和出货")
        
def get_uploaded_files():
    st.header("📤 Excel 预测分析")

    st.subheader("📈 上传预测文件")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    
    st.subheader("🔐 上传未交文件")
    order_file = st.file_uploader("🔐 上传未交文件", type="xlsx", key="order")
    
    st.subheader("🔁 上传出货文件")
    sales_file = st.file_uploader("🔁 上传出货文件", type="xlsx", key="sales")

    # 🚀 生成按钮
    start = st.button("🚀 生成汇总 Excel")

    return forecast_file, order_file, sales_file, start
