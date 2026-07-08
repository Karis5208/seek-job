from openpyxl import Workbook

# 创建新的工作簿
wb = Workbook()
ws = wb.active
ws.title = "秋招信息汇总"

# 设置表头
headers = [
    "公司名称",
    "岗位名称",
    "岗位类别",
    "工作地点",
    "学历要求",
    "专业要求",
    "投递方式/链接",
    "开始时间",
    "截止时间",
    "信息来源",
    "收集日期",
    "备注",
    "状态"
]

for col_num, header in enumerate(headers, 1):
    ws.cell(row=1, column=col_num, value=header)

# 调整列宽
column_widths = [15, 20, 10, 15, 10, 12, 30, 12, 12, 15, 12, 25, 10]
for i, width in enumerate(column_widths, 1):
    ws.column_dimensions[chr(64 + i)].width = width

# 保存文件
wb.save("/workspace/autumn_recruitment_2025.xlsx")
print("Excel模板创建成功: /workspace/autumn_recruitment_2025.xlsx")
