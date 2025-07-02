import pandas as pd
import re
from io import BytesIO
import streamlit as st
from openpyxl.styles import Alignment, Font
from urllib.parse import quote

class PivotProcessor:
    def process(self, template_file, forecast_file, order_file, sales_file):
        
        # 🔗 构建 raw URL，确保路径中文被编码
        raw_mapping_url = (
            "https://raw.githubusercontent.com/TTTriste06/operation_planning-/main/"
            + quote("新旧料号.xlsx")
        )

        # 📥 尝试加载
        try:
            mapping_df = pd.read_excel(raw_mapping_url)
        except Exception as e:
            raise ValueError(f"❌ 加载新旧料号映射表失败：{e}")

        st.write(mapping_df)
        # 创建新的 mapping_semi：仅保留“半成品”字段非空的行
        mapping_semi1 = mapping_df[
            ["新晶圆", "新规格", "新品名", "半成品"]
        ]
        mapping_semi1 = mapping_semi1[~mapping_df["半成品"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi1 = mapping_semi1[~mapping_df["新品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi2 = mapping_df[
            ["新晶圆", "新规格", "新品名", "旧晶圆", "旧规格", "旧品名", "半成品"]
        ]
        mapping_semi2 = mapping_semi2[mapping_semi2["新品名"].astype(str).str.strip().replace("nan", "") == ""].copy()
        mapping_semi2 = mapping_semi2[~mapping_semi2["半成品"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi2 = mapping_semi2[~mapping_semi2["旧品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi2 = mapping_semi2.drop(columns=["新晶圆", "新规格", "新品名"])
        mapping_semi2.columns = ["新晶圆", "新规格", "新品名", "半成品"]
        
        mapping_semi = pd.concat([mapping_semi1, mapping_semi2], ignore_index=True)
       
        # 去除“品名”为空的行
        mapping_new = mapping_df[
            ["旧晶圆", "旧规格", "旧品名", "新晶圆", "新规格", "新品名"]
        ]
        mapping_new = mapping_new[~mapping_df["新品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_new = mapping_new[~mapping_new["旧品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        
        
        # 去除“替代品名”为空的行，并保留指定字段
        mapping_sub = pd.DataFrame()
        for i in range(1, 5):
            sub_cols = ["新晶圆", "新规格", "新品名", f"替代晶圆{i}", f"替代规格{i}", f"替代品名{i}"]
            sub_df = mapping_df[sub_cols].copy()
            
            # 去除“替代品名”为空或为 nan 的行
            valid_mask = ~sub_df[f"替代品名{i}"].astype(str).str.strip().replace("nan", "").eq("")
            sub_df = sub_df[valid_mask].copy()
        
            # 统一列名
            sub_df.columns = ["新晶圆", "新规格", "新品名", "替代晶圆", "替代规格", "替代品名"]
            mapping_sub = pd.concat([mapping_sub, sub_df], ignore_index=True)


        # Step 1: 读取主计划模板
        main_df = pd.read_excel(template_file, sheet_name=0, header=1)
        main_df = main_df[["晶圆", "规格", "品名"]].copy()
        st.write(main_df)

        # Step 2: 加载数据
        df_forecast = pd.read_excel(forecast_file)
        df_order = pd.read_excel(order_file, sheet_name="Sheet")
        df_sales = pd.read_excel(sales_file, sheet_name="原表")

        # === 进行新旧料号替换 ===
        all_replaced_names = set()
        for df, key in zip([df_forecast, df_order, df_sales], ["forecast", "order", "sales"]):
            df, replaced_main = apply_mapping_and_merge(df, mapping_df, FIELD_MAPPINGS[key])
            all_replaced_names.update(replaced_main)
            df, replaced_sub = apply_extended_substitute_mapping(df, mapping_df, FIELD_MAPPINGS[key])
            all_replaced_names.update(replaced_sub)
        
        

        # Step 3: 提取月份列
        month_pattern = re.compile(r"(\d{1,2})月预测")
        month_cols = [col for col in df_forecast.columns if month_pattern.match(col)]
        forecast_months = [f"2025-{month_pattern.match(col).group(1).zfill(2)}" for col in month_cols]

        # Step 4: 初始化列
        for ym in forecast_months:
            main_df[f"{ym}-预测"] = 0
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        # Step 5: 填入预测数据
        df_forecast["品名"] = df_forecast["生产料号"].astype(str).str.strip()
        for col in month_cols:
            month_num = month_pattern.match(col).group(1).zfill(2)
            ym = f"2025-{month_num}"
            summary = df_forecast.groupby("品名")[col].sum(min_count=1)
            main_df[f"{ym}-预测"] = main_df["品名"].map(summary).fillna(0)

        # Step 6: 填入订单数据
        df_order["回复客户交期"] = pd.to_datetime(df_order["回复客户交期"], errors="coerce")
        df_order["年月"] = df_order["回复客户交期"].dt.to_period("M").astype(str)
        grouped_order = df_order.groupby(["品名", "年月"])["未交订单数量"].sum().unstack().fillna(0)
        for ym in grouped_order.columns:
            colname = f"{ym}-订单"
            if colname in main_df.columns:
                main_df[colname] = main_df["品名"].map(grouped_order[ym]).fillna(0)

        # Step 7: 填入出货数据
        df_sales["交易日期"] = pd.to_datetime(df_sales["交易日期"], errors="coerce")
        df_sales["年月"] = df_sales["交易日期"].dt.to_period("M").astype(str)
        grouped_sales = df_sales.groupby(["品名", "年月"])["数量"].sum().unstack().fillna(0)
        for ym in grouped_sales.columns:
            colname = f"{ym}-出货"
            if colname in main_df.columns:
                main_df[colname] = main_df["品名"].map(grouped_sales[ym]).fillna(0)

        # Step 8: 输出 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]

            # 合并前3列标题
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

            # 合并月份字段
            col = 4
            for ym in forecast_months:
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
                top_cell = ws.cell(row=1, column=col)
                top_cell.value = ym
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)

                ws.cell(row=2, column=col).value = "预测"
                ws.cell(row=2, column=col + 1).value = "订单"
                ws.cell(row=2, column=col + 2).value = "出货"
                col += 3

        output.seek(0)
        return main_df, output
