"""生成秋招信息 Markdown 文件（大公司优先，开放中优先）"""
from openpyxl import load_workbook
from datetime import datetime
import os, re

wb = load_workbook('/workspace/autumn_recruitment_2025.xlsx')
ws = wb['秋招信息汇总']
rows = list(ws.iter_rows(values_only=True))
today = datetime.now().strftime("%Y-%m-%d")

# 按类别分组（保持Excel中的排序顺序）
by_category = {}
for row in rows[1:]:
    cat = row[2] if row[2] else "其他"
    if cat not in by_category:
        by_category[cat] = []
    by_category[cat].append(row)

# 类别排序：金融 > 战略 > 商分 > 运营 > 管培 > 市场 > 产品 > 咨询 > 职能 > 供应链 > 传媒 > 零售 > 其他
cat_order = ["金融", "战略", "商分", "运营", "管培", "市场", "产品", "咨询", "职能", "供应链", "传媒", "零售"]
cats_sorted = sorted(by_category.items(), key=lambda x: (cat_order.index(x[0]) if x[0] in cat_order else 99, -len(x[1])))

total = len(rows) - 1
open_count = sum(1 for r in rows[1:] if r[12] == "开放中")
closed_count = total - open_count

# 主索引
lines = []
lines.append('# 2026 秋招信息汇总')
lines.append('')
lines.append(f'> 早上 6:00 自动更新 | 更新：{today} | 共 {total} 条（开放中 {open_count} | 已截止 {closed_count}）')
lines.append('')
lines.append('📱 **大公司优先排列**，点击分类进入详情：')
lines.append('')
lines.append('## 岗位分类')
lines.append('')
lines.append('| 分类 | 数量 | 开放中 |')
lines.append('|------|------|--------|')
for cat, items in cats_sorted:
    safe_name = cat.replace("/", "-")
    cat_open = sum(1 for r in items if r[12] == "开放中")
    lines.append(f'| [{cat}](./docs/{safe_name}.md) | {len(items)} 条 | {cat_open} |')
lines.append('')
lines.append('---')
lines.append('')
lines.append('同时提供 [Excel版本](./autumn_recruitment_2025.xlsx)，下载后可筛选排序。')
lines.append('')
lines.append('*每日自动更新 · 祝秋招顺利！*')

with open('/workspace/autumn_recruitment.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

# 各分类文件（保持Excel中的排序顺序）
os.makedirs('/workspace/docs', exist_ok=True)
for cat, items in cats_sorted:
    safe_name = cat.replace("/", "-")
    
    # Excel中已按公司评分排序，直接保持顺序即可
    # 只需确保开放中优先（Excel中已排序）
    items_sorted = sorted(items, key=lambda r: (0 if r[12] == "开放中" else 1, list(rows[1:]).index(r) if r in rows[1:] else 999))
    
    clines = []
    clines.append(f'# {cat}')
    clines.append('')
    open_count = sum(1 for r in items_sorted if r[12] == "开放中")
    closed_count = len(items_sorted) - open_count
    clines.append(f'> 共 {len(items)} 条 | 开放中 {open_count} | 已截止 {closed_count} | 更新于 {today}')
    clines.append('')
    clines.append('| 公司 | 岗位 | 地点 | 截止 | 投递 | 备注 | 状态 |')
    clines.append('|------|------|------|------|------|------|------|')
    
    for row in items_sorted:
        company, job, category, location, edu, major, apply, start, end, source, collect_date, note, status = row
        status_e = '🟢' if status == '开放中' else '🔴'
        company = (company or '—')[:20]
        job = (job or '—')[:35]
        location = (location or '—')[:12]
        end = (end or '—')[:10]
        note = (note or '—')[:18]
        apply_cell = f'[投递]({apply})' if apply and apply.startswith('http') else '—'
        clines.append(f'| {company} | {job} | {location} | {end} | {apply_cell} | {note} | {status_e} {status} |')
    
    clines.append('')
    clines.append(f'[← 返回首页](../autumn_recruitment.md)')
    
    with open(f'/workspace/docs/{safe_name}.md', 'w', encoding='utf-8') as f:
        f.write('\n'.join(clines))

print(f'已生成 {len(cats_sorted)+1} 个 Markdown 文件')