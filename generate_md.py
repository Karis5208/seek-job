from openpyxl import load_workbook

wb = load_workbook('/workspace/autumn_recruitment_2025.xlsx')
ws = wb['秋招信息汇总']

rows = list(ws.iter_rows(values_only=True))
headers = rows[0]

md_lines = []
md_lines.append('# 2026 秋招信息汇总')
md_lines.append('')
md_lines.append('> 每天早上 6:00 自动更新 | 最后更新：2026-07-08')
md_lines.append('')
md_lines.append('## 使用说明')
md_lines.append('')
md_lines.append('- 本文件自动生成，手机和网页端均可直接查看')
md_lines.append('- 同时提供 [Excel版本](./autumn_recruitment_2025.xlsx)，便于筛选排序')
md_lines.append('- 岗位类型涵盖：金融、战略、商分、运营、风险管理等泛商科岗位')
md_lines.append('')
md_lines.append('## 岗位汇总')
md_lines.append('')

if len(rows) <= 1:
    md_lines.append('暂无数据，首次运行收集后将自动填充。')
    md_lines.append('')
else:
    for row in rows[1:]:
        company, job, category, location, edu, major, apply, start, end, source, collect_date, note, status = row
        status_emoji = '🟢' if status == '开放中' else '🔴'
        md_lines.append(f'### {company} - {job}')
        md_lines.append('')
        md_lines.append(f'- **岗位类别**：{category or "—"}')
        md_lines.append(f'- **工作地点**：{location or "—"}')
        md_lines.append(f'- **学历要求**：{edu or "—"}')
        md_lines.append(f'- **专业要求**：{major or "—"}')
        md_lines.append(f'- **投递方式**：{apply or "—"}')
        md_lines.append(f'- **开始时间**：{start or "—"}')
        md_lines.append(f'- **截止时间**：{end or "—"}')
        md_lines.append(f'- **信息来源**：{source or "—"}')
        md_lines.append(f'- **收集日期**：{collect_date or "—"}')
        md_lines.append(f'- **状态**：{status_emoji} {status or "—"}')
        if note:
            md_lines.append(f'- **备注**：{note}')
        md_lines.append('')

md_lines.append('---')
md_lines.append('')
md_lines.append('*每日自动更新 · 祝秋招顺利！*')

with open('/workspace/autumn_recruitment.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(md_lines))

print('Markdown文件已生成: /workspace/autumn_recruitment.md')
