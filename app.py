#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from datetime import datetime
import streamlit as st
from io import BytesIO

# 基础样式（极简版）
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

def generate_excel():
    wb = Workbook()
    wb.remove(wb.active)

    # ========== 1. 入库记录（核心：直接写死3行公式） ==========
    ws_stockin = wb.create_sheet("入库记录")
    # 表头
    headers = ["入库日期", "商品名称", "型号", "颜色", "数量", "长度(m)", "米重(kg/m)", "FOB价(US)", "汇率", "总成本"]
    for col, header in enumerate(headers, 1):
        cell = ws_stockin.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    # 示例数据 + 硬编码公式（100%确保列对应）
    # 第3行数据（行号2）
    ws_stockin['A2'] = "2024-01-10"
    ws_stockin['B2'] = "钢管A"
    ws_stockin['C2'] = "MODEL-A"
    ws_stockin['D2'] = "Negro"
    ws_stockin['E2'] = 100  # 数量
    ws_stockin['F2'] = 6.0  # 长度
    ws_stockin['G2'] = 2.5  # 米重
    ws_stockin['H2'] = 10.0 # FOB价
    ws_stockin['I2'] = 7.2  # 汇率
    ws_stockin['J2'] = "=E2*F2*G2*H2*I2"  # 总成本公式（数量×长度×米重×FOB×汇率）
    ws_stockin['J2'].number_format = '#,##0.00'

    # 第4行数据（行号3）
    ws_stockin['A3'] = "2024-01-15"
    ws_stockin['B3'] = "钢管B"
    ws_stockin['C3'] = "MODEL-B"
    ws_stockin['D3'] = "Plata"
    ws_stockin['E3'] = 50
    ws_stockin['F3'] = 6.0
    ws_stockin['G3'] = 3.0
    ws_stockin['H3'] = 12.0
    ws_stockin['I3'] = 7.2
    ws_stockin['J3'] = "=E3*F3*G3*H3*I3"

    # 第5行数据（行号4）
    ws_stockin['A4'] = "2024-02-01"
    ws_stockin['B4'] = "钢管A"
    ws_stockin['C4'] = "MODEL-A"
    ws_stockin['D4'] = "Negro"
    ws_stockin['E4'] = 80
    ws_stockin['F4'] = 6.0
    ws_stockin['G4'] = 2.5
    ws_stockin['H4'] = 10.5
    ws_stockin['I4'] = 7.3
    ws_stockin['J4'] = "=E4*F4*G4*H4*I4"

    # 给所有单元格加边框
    for row in range(1, 5):
        for col in range(1, 11):
            ws_stockin.cell(row=row, column=col).border = thin_border

    # ========== 2. 仪表盘（汇总入库记录的总成本） ==========
    ws_dash = wb.create_sheet("仪表盘")
    ws_dash['A1'] = "总入库金额（美元）"
    ws_dash['B1'] = "=SUM(入库记录!J2:J4)"  # 汇总入库记录J列的总成本
    ws_dash['B1'].number_format = '#,##0.00'
    ws_dash['A2'] = "入库次数"
    ws_dash['B2'] = "=COUNTA(入库记录!A2:A4)"  # 统计入库行数

    # ========== 3. 启用自动计算 ==========
    wb.calculation.calcMode = "auto"

    # 保存文件
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit界面
st.set_page_config(page_title="极简自动计算Excel", layout="wide")
st.title("📊 极简版自动计算库存表")

if st.button("生成Excel文件", type="primary"):
    excel_file = generate_excel()
    st.success("✅ 文件生成成功！")
    st.download_button(
        label="下载Excel",
        data=excel_file,
        file_name="极简自动计算库存表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("### 验证步骤：")
st.markdown("1. 下载文件后打开「入库记录」工作表")
st.markdown("2. 修改E2单元格（数量）的数值（比如改成200）")
st.markdown("3. 看J2单元格（总成本）是否自动变成 200×6×2.5×10×7.2 = 216000")
st.markdown("4. 切换到「仪表盘」，B1单元格会自动汇总J2:J4的总和")
