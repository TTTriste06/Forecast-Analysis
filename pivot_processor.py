import re
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime, date
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

from mapping_utils import (
    clean_mapping_headers, 
    replace_all_names_with_mapping, 
    apply_mapping_and_merge, 
    apply_extended_substitute_mapping,
    apply_all_name_replacements
)

class PivotProcessor:
    def process(self, output_buffer, df_forecast, df_order, df_sales):
        """
        替换品名、新建主计划表，并直接写入 Excel 文件（含列宽调整、标题行）。
        """
        # === 读取文件 ===
        raw_mapping_url = "https://raw.githubusercontent.com/TTTriste06/operation_planning-/main/新旧料号.xlsx"
        raw_template_url = "https://raw.githubusercontent.com/TTTriste06/Forecast-Analysis/main/预测分析.xlsx"
        mapping_df = pd.read_excel(raw_mapping_url)
        main_df = pd.read_excel(raw_template_url)

        st.write(df_forecast)
        st.write(df_order)
        st.write(df_sales)
        st.write(mapping_df)
        st.write(main_df)

        # 创建新的 mapping_semi：仅保留“半成品”字段非空的行
        mapping_semi1 = mapping_df[
            ["新晶圆品名", "新规格", "新品名", "半成品"]
        ]
        mapping_semi1 = mapping_semi1[~mapping_df["半成品"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi1 = mapping_semi1[~mapping_df["新品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi2 = mapping_df[
            ["新晶圆品名", "新规格", "新品名", "旧晶圆品名", "旧规格", "旧品名", "半成品"]
        ]
        mapping_semi2 = mapping_semi2[mapping_semi2["新品名"].astype(str).str.strip().replace("nan", "") == ""].copy()
        mapping_semi2 = mapping_semi2[~mapping_semi2["半成品"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi2 = mapping_semi2[~mapping_semi2["旧品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_semi2 = mapping_semi2.drop(columns=["新晶圆品名", "新规格", "新品名"])
        mapping_semi2.columns = ["新晶圆品名", "新规格", "新品名", "半成品"]
        
        mapping_semi = pd.concat([mapping_semi1, mapping_semi2], ignore_index=True)
       
        # 去除“品名”为空的行
        mapping_new = mapping_df[
            ["旧晶圆品名", "旧规格", "旧品名", "新晶圆品名", "新规格", "新品名"]
        ]
        mapping_new = mapping_new[~mapping_df["新品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        mapping_new = mapping_new[~mapping_new["旧品名"].astype(str).str.strip().replace("nan", "").eq("")].copy()
        
        
        # 去除“替代品名”为空的行，并保留指定字段
        mapping_sub = pd.DataFrame()
        for i in range(1, 5):
            sub_cols = ["新晶圆品名", "新规格", "新品名", f"替代晶圆{i}", f"替代规格{i}", f"替代品名{i}"]
            sub_df = mapping_df[sub_cols].copy()
            
            # 去除“替代品名”为空或为 nan 的行
            valid_mask = ~sub_df[f"替代品名{i}"].astype(str).str.strip().replace("nan", "").eq("")
            sub_df = sub_df[valid_mask].copy()
        
            # 统一列名
            sub_df.columns = ["新晶圆品名", "新规格", "新品名", "替代晶圆", "替代规格", "替代品名"]
            mapping_sub = pd.concat([mapping_sub, sub_df], ignore_index=True)
        
        # === 构建主计划 ===
        main_plan_df = pd.DataFrame(columns=headers)

        ## == 品名 ==
        df_unfulfilled = self.dataframes.get("赛卓-未交订单")
        df_forecast = self.additional_sheets.get("赛卓-预测")

        name_unfulfilled = []
        name_forecast = []

        if df_unfulfilled is not None and not df_unfulfilled.empty:
            col_name = FIELD_MAPPINGS["赛卓-未交订单"]["品名"]
            name_unfulfilled = df_unfulfilled[col_name].astype(str).str.strip().tolist()

        if df_forecast is not None and not df_forecast.empty:
            col_name = FIELD_MAPPINGS["赛卓-预测"]["品名"]
            name_forecast = df_forecast[col_name].astype(str).str.strip().tolist()

        all_names = pd.Series(name_unfulfilled + name_forecast)
        all_names = replace_all_names_with_mapping(all_names, mapping_new, mapping_df)
        main_plan_df = main_plan_df.reindex(index=range(len(all_names)))
        if not all_names.empty:
            main_plan_df["品名"] = all_names.values



        ## == 替换新旧料号、替代料号 ==
        target_sheets = [
            ("赛卓-安全库存", self.additional_sheets),
            ("赛卓-预测", self.additional_sheets),
            ("赛卓-未交订单", self.dataframes),
            ("赛卓-成品库存", self.dataframes),
            ("赛卓-成品在制", self.dataframes),
            ("赛卓-CP在制", self.dataframes),
            ("赛卓-晶圆库存", self.dataframes),
            ("赛卓-到货明细", self.dataframes),
            ("赛卓-下单明细", self.dataframes),
            ("赛卓-销货明细", self.dataframes),
        ]
        
        all_replaced_names = set()
        
        # 执行替换逻辑
        for sheet_name, container in target_sheets:
            df_new = container[sheet_name]
        
            # 主映射替换
            df_new, replaced_main = apply_mapping_and_merge(df_new, mapping_new, FIELD_MAPPINGS[sheet_name])
            all_replaced_names.update(replaced_main)
        
            # 替代映射替换（1~4）
            df_new, replaced_sub = apply_extended_substitute_mapping(df_new, mapping_sub, FIELD_MAPPINGS[sheet_name])
            all_replaced_names.update(replaced_sub)
        
            # 更新回字典
            container[sheet_name] = df_new
        

    
         
        # === 写入 Excel 文件（主计划）===
        timestamp = datetime.now().strftime("%Y%m%d")
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            # 写入Summary
            summary_data = [
                ["", "超链接", "备注"],
                ["数据汇总", "主计划", ""],
                ["赛卓-未交订单-汇总", "赛卓-未交订单-汇总", ""],
                ["赛卓-成品库存-汇总", "赛卓-成品库存-汇总", "关注“hold仓”“成品仓”"],
                ["赛卓-晶圆库存-汇总", "赛卓-晶圆库存-汇总", "晶圆片数已转换为对应的Die数量"],
                ["赛卓-CP在制-汇总", "赛卓-CP在制-汇总", ""],
                ["赛卓-成品在制-汇总", "赛卓-成品在制-汇总", ""],
                ["赛卓-预测", "赛卓-预测", ""],
                ["赛卓-安全库存", "赛卓-安全库存", ""],
                ["赛卓-新旧料号", "赛卓-新旧料号", ""]
            ]
            df_summary = pd.DataFrame(summary_data[1:], columns=summary_data[0])
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
                    
            # 写入主计划表
            main_plan_df = clean_df(main_plan_df)
            main_plan_df.to_excel(writer, sheet_name="主计划", index=False, startrow=1)
        
            # 获取 workbook 和 worksheet
            wb = writer.book
            ws = wb["主计划"]
        
            # 写时间戳和说明
            ws.cell(row=1, column=1, value=f"主计划生成时间：{timestamp}")            
            legend_cell = ws.cell(row=1, column=3)
            legend_cell.value = (
                "Red < 0    "
                "Yellow < 安全库存    "
                "Orange > 2 × 安全库存"
            )
            legend_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            fill = PatternFill(start_color="FFCCE6FF", end_color="FFCCE6FF", fill_type="solid")
            legend_cell.fill = fill

            # 合并单元格
            merge_safety_header(ws, main_plan_df)
            merge_unfulfilled_order_header(ws)
            merge_forecast_header(ws)
            merge_inventory_header(ws)
            merge_product_in_progress_header(ws)
            merge_order_delivery_amount(ws)
            merge_forecast_accuracy(ws)

            # 高亮显示
            format_monthly_grouped_headers(ws)
            highlight_production_plan_cells(ws, main_plan_df)
            highlight_replaced_names_in_main_sheet(ws, all_replaced_names)

            # 格式调整
            adjust_column_width(ws)
            format_currency_columns_rmb(ws)
            format_thousands_separator(ws)

            # 设置字体加粗，行高也调高一点
            bold_font = Font(bold=True)
            ws.row_dimensions[2].height = 35
    
            # 遍历这一行所有已用到的列，对单元格字体加粗、居中、垂直居中
            max_col = ws.max_column
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=2, column=col_idx)
                cell.font = bold_font
                # 垂直水平居中
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # 自动筛选
            last_col_letter = get_column_letter(ws.max_column)
            ws.auto_filter.ref = f"A2:{last_col_letter}2"
        
            # 冻结
            ws.freeze_panes = "D3"
            append_all_standardized_sheets(writer, uploaded_files, self.additional_sheets)
            
            # 透视表
            standardized_files = standardize_uploaded_keys(uploaded_files, RENAME_MAP)
            parsed_dataframes = {
                filename: pd.read_excel(file)  # 或提前 parse 完成的 DataFrame dict
                for filename, file in standardized_files.items()
            }
            pivot_tables = generate_monthly_pivots(parsed_dataframes, pivot_config)
            for sheet_name, df in pivot_tables.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
            # 写完后手动调整所有透视表 sheet 的列宽
            for sheet_name, df in pivot_tables.items():
                ws = writer.book[sheet_name]
                for col_cells in ws.columns:
                    max_length = 0
                    col_letter = col_cells[0].column_letter
                    for cell in col_cells:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = max_length * 1.2 + 10
                    ws.column_dimensions[col_letter].width = min(adjusted_width, 50)

            # 获取 workbook 和 worksheet
            ws_summary = wb["Summary"]
            add_sheet_hyperlinks(ws_summary, wb.sheetnames)
            
            for col_idx in range(1, ws_summary.max_column + 1):
                col_letter = get_column_letter(col_idx)
                ws_summary.column_dimensions[col_letter].width = 25

        output_buffer.seek(0)
       
