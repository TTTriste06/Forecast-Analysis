import pandas as pd
import re
from io import BytesIO
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from urllib.parse import quote
from collections import defaultdict

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
    def __init__(self):
        self.source_map = defaultdict(list)

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

        main_df = template_file[["晶圆", "规格", "品名"]].copy()
        main_df.columns = ["晶圆品名", "规格", "品名"]

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

        all_months = extract_all_year_months(forecast_file, order_file, sales_file)

        for ym in all_months:
            main_df[f"{ym}-预测"] = 0
            main_df[f"{ym}-订单"] = 0
            main_df[f"{ym}-出货"] = 0

        main_df = fill_forecast_data(main_df, forecast_file, all_months, self.source_map)
        main_df = fill_order_data(main_df, order_file, all_months, self.source_map)
        main_df = fill_sales_data(main_df, sales_file, all_months, self.source_map)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            main_df.to_excel(writer, index=False, sheet_name="预测分析", startrow=1)
            ws = writer.sheets["预测分析"]
            highlight_by_detecting_column_headers(ws)

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

                ws.cell(row=2, column=col).value = "预测"
                ws.cell(row=2, column=col + 1).value = "订单"
                ws.cell(row=2, column=col + 2).value = "出货"

                fill = PatternFill(start_color=fill_colors[i % len(fill_colors)], end_color=fill_colors[i % len(fill_colors)], fill_type="solid")
                for j in range(col, col + 3):
                    ws.cell(row=1, column=j).fill = fill
                    ws.cell(row=2, column=j).fill = fill
                col += 3

            for row_idx in range(3, ws.max_row + 1):
                name = ws.cell(row=row_idx, column=3).value
                col_ptr = 4
                for ym in all_months:
                    for dtype in ["预测", "订单", "出货"]:
                        val = ws.cell(row=row_idx, column=col_ptr).value
                        if val not in [None, "", 0]:
                            link_cell = f'=HYPERLINK("#\'数据来源\'!A1", "{val}")'
                            ws.cell(row=row_idx, column=col_ptr).value = link_cell
                        col_ptr += 1

            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 10

            source_rows = []
            for (name, ym, dtype), entries in self.source_map.items():
                for e in entries:
                    source_rows.append({
                        "品名": name,
                        "月份": ym,
                        "类型": dtype,
                        "来源": e["来源"],
                        "来源行号": e["来源行号"],
                        "字段值": e["字段值"]
                    })
            df_source = pd.DataFrame(source_rows)
            df_source.to_excel(writer, sheet_name="数据来源", index=False)

        output.seek(0)
        return main_df, output
