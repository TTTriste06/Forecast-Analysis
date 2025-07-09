import pandas as pd
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

class PivotProcessor:
    def process(self, template_file, forecast_file, order_file, sales_file, mapping_file):
        if mapping_file is None:
            raw_mapping_url = (
                "https://raw.githubusercontent.com/TTTriste06/operation_planning-/main/"
                + quote("新旧料号.xlsx")
            )
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
    
        # Step 2: 新旧料号替换
        FIELD_MAPPINGS = {
            "forecast": {"品名": "生产料号"},
            "order": {"品名": "品名"},
            "sales": {"品名": "品名"}
        }
        forecast_file, _ = apply_mapping_and_merge(forecast_file, mapping_new, FIELD_MAPPINGS["forecast"])
        forecast_file, _ = apply_extended_substitute_mapping(forecast_file, mapping_sub, FIELD_MAPPINGS["forecast"])
        order_file, _ = apply_mapping_and_merge(order_file, mapping_new, FIELD_MAPPINGS["order"])
        order_file, _ = apply_extended_substitute_mapping(order_file, mapping_sub, FIELD_MAPPINGS["order"])
        sales_file, _ = apply_mapping_and_merge(sales_file, mapping_new, FIELD_MAPPINGS["sales"])
        sales_file, _ = apply_extended_substitute_mapping(sales_file, mapping_sub, FIELD_MAPPINGS["sales"])
    
        # Step 3: 提取所有月份
        all_months = extract_all_year_months(forecast_file, order_file, sales_file)
    
        # Step 4: 初始化列
        for ym in all_months:
            main_df[f"{ym}-预测"] = 0
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0
    
        # Step 5~7: 数据填充
        main_df = fill_forecast_data(main_df, forecast_file)
        main_df = fill_order_data(main_df, order_file, all_months)
        main_df = fill_sales_data(main_df, sales_file, all_months)
    
        # Step 8: 输出 Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
    
            highlight_by_detecting_column_headers(ws)
    
            # 设置主表头样式
            for i, label in enumerate(["晶圆品名", "规格", "品名"], start=1):
                ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
                cell = ws.cell(row=1, column=i)
                cell.value = label
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)
    
            fill_colors = ["FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EAD1DC", "CFE2F3", "FFE599"]
            col = 4
            for i, ym in enumerate(all_months):
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
                top_cell = ws.cell(row=1, column=col)
                top_cell.value = ym
                top_cell.alignment = Alignment(horizontal="center", vertical="center")
                top_cell.font = Font(bold=True)
                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)], end_color=fill_colors[i % len(fill_colors)], fill_type="solid")
                for j in range(col, col + 3):
                    ws.cell(row=1, column=j).fill = fill
                    ws.cell(row=2, column=j).fill = fill
                col += 3
    
            # 添加超链接（透视跳转）
            from openpyxl.styles import Font as OpenpyxlFont
            from openpyxl.utils import quote_sheetname
    
            header = [cell.value for cell in ws[2]]
            for row in range(3, ws.max_row + 1):
                品名 = ws.cell(row=row, column=3).value
                for col in range(4, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    try:
                        val = float(cell.value or 0)
                        if val == 0:
                            continue
                    except:
                        continue
    
                    col_title = str(ws.cell(row=2, column=col).value).strip()
                    ym = str(ws.cell(row=1, column=col).value).strip()
                    source_sheet = (
                        "原始-预测" if col_title == "预测"
                        else "原始-订单" if col_title == "订单"
                        else "原始-出货" if col_title == "出货"
                        else None
                    )
                    if source_sheet:
                        link = f"#'{quote_sheetname(source_sheet)}'!A1"
                        display = f"{val:.0f}"
                        cell.value = f'=HYPERLINK("{link}", "{display}")'
                        cell.font = OpenpyxlFont(underline="single", color="0563C1")
    
            # 自动列宽
            from openpyxl.utils import get_column_letter
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 5
    
            # 添加原始数据工作表
            forecast_file.to_excel(writer, index=False, sheet_name="原始-预测")
            order_file.to_excel(writer, index=False, sheet_name="原始-订单")
            sales_file.to_excel(writer, index=False, sheet_name="原始-出货")
    
            # 每个原始表前加提示行
            for sheet_name in ["原始-预测", "原始-订单", "原始-出货"]:
                ws_src = writer.sheets[sheet_name]
                ws_src.insert_rows(1)
                ws_src["A1"] = "🔍 请使用筛选功能查看跳转品名与月份对应的原始记录"
                ws_src["A1"].font = OpenpyxlFont(bold=True, color="888888")
    
        output.seek(0)
        return main_df, output
