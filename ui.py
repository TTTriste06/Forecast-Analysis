import streamlit as st

def setup_sidebar():
    st.sidebar.header("📤 上传文件")
    st.sidebar.markdown("上传以下文件以生成主计划：")

def get_uploaded_files():
    template_file = st.sidebar.file_uploader("📁 上传主计划模板", type="xlsx", key="template")
    forecast_file = st.sidebar.file_uploader("📈 上传预测数据", type="xlsx", key="forecast")
    order_file = st.sidebar.file_uploader("📦 上传未交订单", type="xlsx", key="order")
    sales_file = st.sidebar.file_uploader("🚚 上传出货明细", type="xlsx", key="sales")
    start = st.sidebar.button("🚀 生成主计划")
    return template_file, forecast_file, order_file, sales_file, start
