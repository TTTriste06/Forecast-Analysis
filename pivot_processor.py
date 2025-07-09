dimport pandas as pd
import re
from io import BytesIO
import streamlit as st
from openpyxl.styles import Alignment, Font
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from urllib.parse import quote
from mapping_utils import (
    apply_mapping_and_merge, 
    apply_extended_substitute_mapping,
    split_mapping_data
)
from info_extract import (
    extract_all_year_months, 
    fill_forecast_data, 
    fill_order_data, 
    fill_sales_data, 
    highlight_by_detecting_column_headers
)
from collections import defaultdict

class PivotProcessor:
    def __init__(self):
        self.source_map = defaultdict(list)
    def process(self, template_file, forecast_file, order_file, sales_file, mapping_file):
        if mapping_file is None:
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
        else:
            mapping_df = pd.read_excel(mapping_file)
        mapping_semi, mapping_new, mapping_sub = split_mapping_data(mapping_df)

        # Step 1: 读取主计划模板
        main_df = template_file[["晶圆", "规格", "品名"]].copy()
        main_df.columns = ["晶圆品名", "规格", "品名"]
        
        # Step 2: 进行新旧料号替换 
        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }
        
        forecast_file, replaced_main = apply_mapping_and_merge(forecast_file, mapping_new, FIELD_MAPPINGS["forecast"])
        forecast_file, replaced_sub = apply_extended_substitute_mapping(forecast_file, mapping_sub, FIELD_MAPPINGS["forecast"])
        order_file, replaced_main = apply_mapping_and_merge(order_file, mapping_new, FIELD_MAPPINGS["order"])
        order_file, replaced_sub = apply_extended_substitute_mapping(order_file, mapping_sub, FIELD_MAPPINGS["order"])
        sales_file, replaced_main = apply_mapping_and_merge(sales_file, mapping_new, FIELD_MAPPINGS["sales"])
        sales_file, replaced_sub = apply_extended_substitute_mapping(sales_file, mapping_sub, FIELD_MAPPINGS["sales"])

        # Step 3: 提取月份列
        all_months = extract_all_year_months(forecast_file, order_file, sales_file)

        # Step 4: 初始化列
        for ym in all_months:
            main_df[f"{ym}-预测"] = 0
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        # Step 5: 填入预测数据
        main_df = fill_forecast_data(main_df, forecast_file, all_months)

        # Step 6: 填入订单数据
        main_df = fill_order_data(main_df, order_file, all_months)

        # Step 7: 填入出货数据
        main_df = fill_sales_data(main_df, sales_file, all_months)
        
        # Step 8: 输出 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]

            highlight_by_detecting_column_headers(ws)

            # === 设置基本字段（三列）合并行 ===
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)
        
            # === 合并每月三列，并设置标题 ===
            fill_colors = [
                "FFF2CC",  # 浅黄色
                "D9EAD3",  # 浅绿色
                "D0E0E3",  # 浅蓝色
                "F4CCCC",  # 浅红色
                "EAD1DC",  # 浅紫色
                "CFE2F3",  # 浅青色
                "FFE599",  # 明亮黄
            ]
            
            col = 4
            for i, ym in enumerate(all_months):
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
                top_cell = ws.cell(row=1, column=col)
                top_cell.value = ym
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)
            
                # 设置底部三列
                ws.cell(row=2, column=col).value = "预测"
                ws.cell(row=2, column=col + 1).value = "订单"
                ws.cell(row=2, column=col + 2).value = "出货"
            
                # 应用颜色样式
                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)], end_color=fill_colors[i % len(fill_colors)], fill_type="solid")
                for j in range(col, col + 3):
                    ws.cell(row=1, column=j).fill = fill
                    ws.cell(row=2, column=j).fill = fill
            
                col += 3

            # === 自动列宽调整 ===
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 10

        output.seek(0)
        return main_df, output
