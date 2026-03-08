#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO

# ========== 样式定义（复用原有逻辑） ==========
title_font = Font(size=14, bold=True, color="FFFFFF")
title_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
title_alignment = Alignment(horizontal="center", vertical="center")
header_font = Font(size=11, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
input_font = Font(size=10, color="1565C0")
input_alignment = Alignment(horizontal="center", vertical="center")
formula_font = Font(size=10, color="000000")
formula_alignment = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
color_list = ["Negro", "Plata", "Champán", "Madera", "Blanco"]

# ========== 原有工具函数（复用） ==========
def set_column_width(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

def apply_title(ws, title, row=1, col_range="A:F"):
    """应用标题样式"""
    # 修正合并单元格范围：将 "B:G"+"1" 改为 "B1:G1"
    start_col, end_col = col_range.split(":")
    merge_range = f"{start_col}{row}:{end_col}{row}"
    ws.merge_cells(merge_range)
    # 取合并区域的第一个单元格（而非固定B列），避免列范围超出的问题
    cell = ws[f'{start_col}{row}']
    cell.value = title
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = title_alignment
    ws.row_dimensions[row].height = 30

def apply_header(ws, headers, start_row=2):
    for col_idx, header in enumerate(headers, start=2):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

def apply_input_style(ws, row, col):
    cell = ws.cell(row=row, column=col)
    cell.font = input_font
    cell.alignment = input_alignment
    cell.border = thin_border

def apply_formula_style(ws, row, col):
    cell = ws.cell(row=row, column=col)
    cell.font = formula_font
    cell.alignment = formula_alignment
    cell.border = thin_border

# ========== 生成Excel核心函数 ==========
def generate_excel():
    wb = Workbook()
    wb.remove(wb.active)

    # 1. 仪表盘
    ws_dashboard = wb.create_sheet("仪表盘")
    ws_dashboard.sheet_view.showGridLines = False
    apply_title(ws_dashboard, "库存管理系统仪表盘", col_range="B:G")
    ws_dashboard['B3'] = "日期筛选："
    ws_dashboard['B3'].font = Font(size=11, bold=True)
    ws_dashboard['C3'] = "起始日期"
    ws_dashboard['D3'] = "=DATE(2024,1,1)"
    ws_dashboard['E3'] = "终止日期"
    ws_dashboard['F3'] = "=TODAY()"
    ws_dashboard['B5'] = "财务汇总"
    ws_dashboard['B5'].font = Font(size=12, bold=True)
    headers_dash = ["指标", "数值", "单位"]
    apply_header(ws_dashboard, headers_dash, start_row=6)
    dash_data = [
        ["总入库金额", "=SUM(入库记录!H:H)", "美元"],
        ["总出库金额", "=SUM(出库记录!J:J)", "美元"],
        ["当前库存价值", "=SUM(期末库存!J:J)", "美元"],
        ["销售毛利（估算）", "=SUM(出库记录!J:J)-SUM(出库记录!I:I)", "美元"],
        ["入库次数", "=COUNTA(入库记录!B:B)-1", "次"],
        ["出库次数", "=COUNTA(出库记录!B:B)-1", "次"],
    ]
    for row_idx, data in enumerate(dash_data, start=7):
        for col_idx, value in enumerate(data, start=2):
            cell = ws_dashboard.cell(row=row_idx, column=col_idx, value=value)
            if col_idx == 2:
                cell.number_format = '#,##0.00'
                cell.font = formula_font
            else:
                cell.font = Font(size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
    set_column_width(ws_dashboard, {'B': 20, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15})

    # 2. 客户信息
    ws_customers = wb.create_sheet("客户信息")
    ws_customers.sheet_view.showGridLines = False
    apply_title(ws_customers, "客户信息", col_range="B:F")
    headers_customers = ["注册日期", "客户名称", "客户电话", "客户地区", "详细地址"]
    apply_header(ws_customers, headers_customers)
    sample_customers = [
        ["2024-01-15", "ABC公司", "1234567890", "北京", "朝阳区某某路123号"],
        ["2024-02-20", "XYZ贸易", "0987654321", "上海", "浦东新区某某街456号"],
    ]
    for row_idx, data in enumerate(sample_customers, start=3):
        for col_idx, value in enumerate(data, start=2):
            cell = ws_customers.cell(row=row_idx, column=col_idx, value=value)
            cell.font = input_font
            cell.alignment = input_alignment
            cell.border = thin_border
    set_column_width(ws_customers, {'B': 15, 'C': 20, 'D': 15, 'E': 15, 'F': 25})

    # 3. 商品信息
    ws_products = wb.create_sheet("商品信息")
    ws_products.sheet_view.showGridLines = False
    apply_title(ws_products, "商品信息", col_range="B:H")
    headers_products = ["商品名称", "商品型号", "商品颜色", "商品数量", "长度(m)", "米重(kg/m)", "支重(kg)"]
    apply_header(ws_products, headers_products)
    color_dv = DataValidation(type="list", formula1='"Negro,Plata,Champán,Madera,Blanco"', allow_blank=True)
    color_dv.error = "请从列表中选择颜色"
    color_dv.errorTitle = "无效输入"
    ws_products.add_data_validation(color_dv)
    color_dv.add('C3:C1000')
    sample_products = [
        ["钢管产品A", "MODEL-A", "Negro", 100, 6.0, 2.5, 15.0],
        ["钢管产品B", "MODEL-B", "Plata", 50, 6.0, 3.0, 18.0],
        ["钢管产品C", "MODEL-C", "Champán", 80, 6.0, 2.8, 16.8],
    ]
    for row_idx, data in enumerate(sample_products, start=3):
        for col_idx, value in enumerate(data, start=2):
            cell = ws_products.cell(row=row_idx, column=col_idx, value=value)
            cell.font = input_font
            cell.alignment = input_alignment
            cell.border = thin_border
            if col_idx in [4, 5, 6, 7]:
                cell.number_format = '#,##0.00'
    set_column_width(ws_products, {'B': 15, 'C': 12, 'D': 12, 'E': 12, 'F': 12, 'G': 12, 'H': 12})

    # 4. 入库记录
    ws_stockin = wb.create_sheet("入库记录")
    ws_stockin.sheet_view.showGridLines = False
    apply_title(ws_stockin, "入库记录", col_range="B:L")
    headers_stockin = ["入库日期", "商品名称", "商品型号", "商品颜色", "商品数量", "长度(m)", "米重(kg/m)", "FOB价格(US)", "汇率", "总成本", "备注"]
    apply_header(ws_stockin, headers_stockin)
    color_dv2 = DataValidation(type="list", formula1='"Negro,Plata,Champán,Madera,Blanco"', allow_blank=True)
    ws_stockin.add_data_validation(color_dv2)
    color_dv2.add('D3:D1000')
    for row in range(3, 103):
        cell = ws_stockin.cell(row=row, column=11)
        cell.value = f"=IF(AND(E{row}>0,F{row}>0,G{row}>0,H{row}>0,I{row}>0),E{row}*F{row}*G{row}*H{row}*I{row},\"\")"
        cell.number_format = '#,##0.00'
        cell.font = formula_font
        cell.alignment = formula_alignment
        cell.border = thin_border
    sample_stockin = [
        ["2024-01-10", "钢管产品A", "MODEL-A", "Negro", 100, 6.0, 2.5, 10.0, 7.2, "", "首批入库"],
        ["2024-01-15", "钢管产品B", "MODEL-B", "Plata", 50, 6.0, 3.0, 12.0, 7.2, "", "首批入库"],
        ["2024-02-01", "钢管产品A", "MODEL-A", "Negro", 80, 6.0, 2.5, 10.5, 7.3, "", "第二批入库"],
    ]
    for row_idx, data in enumerate(sample_stockin, start=3):
        for col_idx, value in enumerate(data, start=2):
            cell = ws_stockin.cell(row=row_idx, column=col_idx, value=value)
            if col_idx <= 5 or col_idx in [7, 8, 9]:
                cell.font = input_font
                cell.number_format = '#,##0.00' if col_idx >= 5 else ''
            elif col_idx == 10:
                cell.value = data[9]
            else:
                cell.font = input_font
            cell.alignment = input_alignment
            cell.border = thin_border
    set_column_width(ws_stockin, {'B': 12, 'C': 12, 'D': 12, 'E': 10, 'F': 10, 'G': 10, 'H': 12, 'I': 8, 'J': 12, 'K': 12, 'L': 15})

    # 5. 出库记录
    ws_stockout = wb.create_sheet("出库记录")
    ws_stockout.sheet_view.showGridLines = False
    apply_title(ws_stockout, "出库记录", col_range="B:K")
    headers_stockout = ["出库日期", "销售单号", "商品名称", "商品型号", "商品颜色", "商品数量", "长度(m)", "米重(kg/m)", "总成本", "客户", "备注"]
    apply_header(ws_stockout, headers_stockout)
    color_dv3 = DataValidation(type="list", formula1='"Negro,Plata,Champán,Madera,Blanco"', allow_blank=True)
    ws_stockout.add_data_validation(color_dv3)
    color_dv3.add('E3:E1000')
    customer_dv = DataValidation(type="list", formula1='客户信息!$C$3:$C$100', allow_blank=True)
    ws_stockout.add_data_validation(customer_dv)
    customer_dv.add('J3:J1000')
    for row in range(3, 103):
        cell = ws_stockout.cell(row=row, column=9)
        cell.value = f"=IF(AND(F{row}>0,G{row}>0,H{row}>0),F{row}*G{row}*H{row}*VLOOKUP(C{row},商品信息!$B$3:$G$100,5,FALSE)*7.2,\"\")"
        cell.number_format = '#,##0.00'
        cell.font = formula_font
        cell.alignment = formula_alignment
        cell.border = thin_border
    sample_stockout = [
        ["2024-02-15", "SO-20240215-001", "钢管产品A", "MODEL-A", "Negro", 30, 6.0, 2.5, "", "ABC公司", "首次销售"],
        ["2024-02-20", "SO-20240220-002", "钢管产品B", "MODEL-B", "Plata", 20, 6.0, 3.0, "", "XYZ贸易", "首次销售"],
    ]
    for row_idx, data in enumerate(sample_stockout, start=3):
        for col_idx, value in enumerate(data, start=2):
            cell = ws_stockout.cell(row=row_idx, column=col_idx, value=value)
            if col_idx <= 7 or col_idx == 9:
                cell.font = input_font
                cell.number_format = '#,##0.00' if col_idx >= 6 else ''
            cell.alignment = input_alignment
            cell.border = thin_border
    set_column_width(ws_stockout, {'B': 12, 'C': 18, 'D': 12, 'E': 12, 'F': 10, 'G': 10, 'H': 10, 'I': 12, 'J': 15, 'K': 15})

    # 6. 期末库存
    ws_inventory = wb.create_sheet("期末库存")
    ws_inventory.sheet_view.showGridLines = False
    apply_title(ws_inventory, "期末库存 (FIFO)", col_range="B:J")
    headers_inventory = ["商品名称", "商品型号", "商品颜色", "剩余数量", "长度(m)", "米重(kg/m)", "FOB价格(US)", "汇率", "总成本", "入库批次"]
    apply_header(ws_inventory, headers_inventory)
    for row in range(3, 103):
        cell = ws_inventory.cell(row=row, column=9)
        cell.value = f"=IF(AND(D{row}>0,E{row}>0,F{row}>0,G{row}>0,H{row}>0),D{row}*E{row}*F{row}*G{row}*H{row},\"\")"
        cell.number_format = '#,##0.00'
        cell.font = formula_font
        cell.alignment = formula_alignment
        cell.border = thin_border
    sample_inventory = [
        ["钢管产品A", "MODEL-A", "Negro", 150, 6.0, 2.5, 10.2, 7.25, "", "BATCH-20240110"],
        ["钢管产品B", "MODEL-B", "Plata", 30, 6.0, 3.0, 12.0, 7.2, "", "BATCH-20240115"],
        ["钢管产品C", "MODEL-C", "Champán", 80, 6.0, 2.8, 11.0, 7.2, "", "BATCH-20240201"],
    ]
    for row_idx, data in enumerate(sample_inventory, start=3):
        for col_idx, value in enumerate(data, start=2):
            cell = ws_inventory.cell(row=row_idx, column=col_idx, value=value)
            if col_idx >= 4:
                cell.font = input_font
                cell.number_format = '#,##0.00'
            else:
                cell.font = input_font
            cell.alignment = input_alignment
            cell.border = thin_border
    set_column_width(ws_inventory, {'B': 15, 'C': 12, 'D': 12, 'E': 10, 'F': 10, 'G': 10, 'H': 12, 'I': 8, 'J': 12, 'K': 15})

    # 7. 销售单
    ws_invoice = wb.create_sheet("销售单")
    ws_invoice.sheet_view.showGridLines = False
    ws_invoice.merge_cells('B2:K2')
    ws_invoice['B2'] = "销售单 / PEDIDO DE VENTA"
    ws_invoice['B2'].font = Font(size=16, bold=True, color="FFFFFF")
    ws_invoice['B2'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    ws_invoice['B2'].alignment = Alignment(horizontal="center", vertical="center")
    ws_invoice.row_dimensions[2].height = 35
    ws_invoice['B4'] = "销售单号 / No. de Pedido:"
    ws_invoice['C4'] = "SO-__________"
    ws_invoice['B5'] = "日期 / Fecha:"
    ws_invoice['C5'] = "____/____/________"
    ws_invoice['F4'] = "客户 / Cliente:"
    ws_invoice['G4'] = "________________________"
    ws_invoice['F5'] = "电话 / Tel:"
    ws_invoice['G5'] = "________________________"
    for row in [4, 5]:
        for col in ['B', 'C', 'F', 'G']:
            ws_invoice[f'{col}{row}'].font = Font(size=10, bold=True)
    ws_invoice.merge_cells('B7:K7')
    ws_invoice['B7'] = "商品明细 / Detalles de Mercancía"
    ws_invoice['B7'].font = Font(size=12, bold=True, color="FFFFFF")
    ws_invoice['B7'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    ws_invoice['B7'].alignment = Alignment(horizontal="center", vertical="center")
    invoice_headers = ["序号\nNo.", "商品名称\nNombre", "型号\nModelo", "颜色\nColor", "数量\nCantidad", "长度(m)\nLongitud", "米重(kg/m)\nPeso", "单价(US)\nPrecio Unit.", "汇率\nTasa", "总金额\nImporte Total"]
    apply_header(ws_invoice, invoice_headers, start_row=8)
    for row in range(9, 14):
        for col in range(2, 12):
            cell = ws_invoice.cell(row=row, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = input_font
    # 总计行（修正合并单元格赋值问题）
# 1. 先给K14设置格式和公式（合并前操作）
ws_invoice['K14'].font = Font(size=11, bold=True)
ws_invoice['K14'].number_format = '#,##0.00'
ws_invoice['K14'].value = "=SUM(K9:K13)"
ws_invoice['K14'].alignment = Alignment(horizontal="center", vertical="center")

# 2. 再合并单元格（B14:K14）
ws_invoice.merge_cells('B14:K14')

# 3. 给合并后的首单元格设置总计文字
ws_invoice['B14'] = "总计 / TOTAL:"
ws_invoice['B14'].font = Font(size=11, bold=True)
ws_invoice['B14'].alignment = Alignment(horizontal="right", vertical="center")
    ws_invoice['B16'] = "备注 / Observaciones:"
    ws_invoice['B16'].font = Font(size=10, bold=True)
    ws_invoice.merge_cells('B17:K19')
    ws_invoice['B17'].border = thin_border
    ws_invoice['B21'] = "客户签名 / Firma del Cliente:_________________"
    ws_invoice['F21'] = "日期 / Fecha:____/____/________"
    ws_invoice['B22'] = "销售员 / Vendedor:_________________"
    ws_invoice['F22'] = "日期 / Fecha:____/____/________"
    for row in [21, 22]:
        for col in ['B', 'F']:
            ws_invoice[f'{col}{row}'].font = Font(size=10)
    set_column_width(ws_invoice, {'B': 8, 'C': 15, 'D': 12, 'E': 10, 'F': 10, 'G': 10, 'H': 12, 'I': 8, 'J': 10, 'K': 14})

    # 8. 使用说明
    ws_guide = wb.create_sheet("使用说明")
    ws_guide.sheet_view.showGridLines = False
    apply_title(ws_guide, "库存管理系统使用说明", col_range="B:H")
    guide_content = [
        ["", ""],
        ["1. 客户信息", "在「客户信息」表中添加客户数据，用于出库记录选择客户。"],
        ["2. 商品信息", "在「商品信息」表中添加商品数据，包括颜色（从下拉列表选择）。"],
        ["3. 入库记录", "录入入库数据，系统自动计算总成本（数量×长度×米重×FOB价格×汇率）。"],
        ["4. 出库记录", "录入出库数据，选择对应的客户，系统自动计算总成本。"],
        ["5. 期末库存", "根据FIFO先进先出法，自动计算当前库存数量和价值。"],
        ["6. 销售单", "在「销售单」工作表填写销售信息，可打印输出。"],
        ["", ""],
        ["日期筛选:", "在「仪表盘」工作表可以设置日期范围筛选数据。"],
        ["颜色选择:", "商品颜色请从下拉列表中选择：Negro(黑)、Plata(银)、Champán(香槟)、Madera(木)、Blanco(白)。"],
        ["数据联动:", "出库记录的客户从「客户信息」表选择，商品信息自动关联。"],
    ]
    for row_idx, data in enumerate(guide_content, start=3):
        ws_guide.cell(row=row_idx, column=2, value=data[0]).font = Font(size=11, bold=True)
        ws_guide.cell(row=row_idx, column=3, value=data[1]).font = Font(size=10)
        ws_guide.merge_cells(f'C{row_idx}:H{row_idx}')
    set_column_width(ws_guide, {'B': 20, 'C': 60})

    # 调整工作表顺序
    wb._sheets = [wb['使用说明'], wb['仪表盘'], wb['客户信息'], wb['商品信息'], 
                  wb['入库记录'], wb['出库记录'], wb['期末库存'], wb['销售单']]
    
    # 保存到BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ========== Streamlit Web界面 ==========
st.set_page_config(page_title="Excel库存管理系统", layout="wide")
st.title("📊 Excel库存管理系统生成工具")
st.markdown("### 点击下方按钮生成库存管理系统Excel文件")

# 生成Excel按钮
if st.button("📥 生成Excel文件", type="primary"):
    excel_file = generate_excel()
    st.success("✅ Excel文件生成成功！")
    # 下载按钮
    st.download_button(
        label="📤 下载Excel文件",
        data=excel_file,
        file_name="Excel库存管理系统.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.markdown("### 使用说明")
st.markdown("1. 点击生成按钮后等待文件生成")
st.markdown("2. 点击下载按钮保存Excel文件到本地")
st.markdown("3. 打开Excel文件后可直接录入/编辑库存数据")
st.markdown("4. 所有公式已自动配置，支持FIFO库存计算、财务汇总等功能")
