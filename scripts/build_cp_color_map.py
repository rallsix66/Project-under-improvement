# -*- coding: utf-8 -*-
"""从 COA 基准提取 CP 产品码 → 颜色/性状映射，输出到 cp_color_map.json"""
import os, re, json
import win32com.client

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASELINE = os.path.join(SKILL_DIR, "resources", "COA基准", "多款-COA-出报告信息确认表.xls")
OUTPUT = os.path.join(SKILL_DIR, "resources", "cp_color_map.json")

def extract_product_code(name_cn, name_en):
    """从品名提取产品编码: S01, P03, M08, #05, 201, C04, 01#"""
    # 尝试各种编码格式
    patterns = [
        r'\b(S\d{2})\b',      # S01-S06 双头唇釉
        r'\b(P\d{2})\b',      # P01-P09 唇泥
        r'\b(M\d{2})\b',      # M01-M08 腮红
        r'\b(C\d{2})\b',      # C01-C06 双头唇釉 C series
        r'#(\d{2})\b',        # #01-#09 按压唇冻, 液体腮红
        r'\b(\d{3})\b',       # 201-206 唇膏笔
        r'(\d{2})#',          # 01#-03# 气垫
    ]
    for pat in patterns:
        m = re.search(pat, name_cn)
        if m:
            return m.group(0)
    # 尝试英文名
    for pat in patterns:
        m = re.search(pat, name_en)
        if m:
            return m.group(0)
    return None

def extract_product_type(name_cn, name_en):
    """识别产品类型"""
    combined = (name_cn + ' ' + name_en).upper()
    if 'LIP GLAZE' in combined and 'DOUBLE TOUCH' in combined:
        return '双头唇釉'
    if 'LIP GLAZEALL' in combined or ('LIP GLAZE' in combined and 'MATTE' in combined):
        return '哑光亮光双头唇釉'
    if 'LIP MUD' in combined:
        return '唇泥'
    if 'BLUSH MUD' in combined:
        return '腮红泥'
    if 'LIQUID BLUSH' in combined:
        return '液体腮红'
    if 'LIP JELLY' in combined:
        return '按压唇冻'
    if 'LIP STICK' in combined:
        return '双头唇膏+唇线笔'
    if 'CUSHION' in combined:
        return '防晒气垫'
    if 'SETTING COMPACT' in combined or '粉饼' in name_cn:
        return '粉饼'
    return '未知'

def parse_color_shape(cs_cn):
    """拆分颜色+性状，如 '浅珊瑚粉色膏状物' → ('浅珊瑚粉色', '膏状物')"""
    for suffix in ['膏状物', '膏状体', '膏状', '液体', '液状', '乳液', '凝胶',
                   '粉状物', '粉状', '泥状', '棒状', '粘稠液体', '烟雾剂',
                   '气雾', '喷雾', '洁面', '慕斯', '泡沫', '霜状', '乳霜',
                   '手套型', '精华']:
        if cs_cn.endswith(suffix):
            return cs_cn[:-len(suffix)], suffix
    return cs_cn, ""

def main():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(BASELINE)
    ws = wb.Worksheets("CP")

    color_map = {}
    entries = []

    for r in range(7, ws.UsedRange.Rows.Count + 1):
        idx = ws.Cells(r, 1).Value
        name_cell = ws.Cells(r, 2).Value
        cs_cell = ws.Cells(r, 4).Value
        if not idx or not name_cell:
            continue

        cn_text = str(name_cell).split('\n')[0].strip()
        en_text = str(name_cell).split('\n')[1].strip() if '\n' in str(name_cell) else ""

        cs_str = str(cs_cell)
        cs_cn = cs_str.split('\n')[0].strip() if '\n' in cs_str else cs_str.strip()

        code = extract_product_code(cn_text, en_text)
        ptype = extract_product_type(cn_text, en_text)
        color, shape = parse_color_shape(cs_cn)

        # Fix parsing errors
        if 'Grayish-pink' in color:  # P05 data concatenated
            color = '灰玫粉色'

        if not code:
            # Try to identify by type only
            if ptype == '粉饼':
                code = 'SETTING_COMPACT'
            else:
                print(f"WARN: No code for: {cn_text[:40]}")
                continue

        key = f"{ptype}:{code}"
        entry = {
            'code': code,
            'type': ptype,
            'name_cn': cn_text,
            'name_en': en_text,
            'color_cn': color,
            'shape_cn': shape,
        }
        color_map[key] = entry
        entries.append(entry)

    wb.Close(False)
    excel.Quit()

    # Output
    with open(OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(color_map, f, ensure_ascii=False, indent=2)

    print(f"Extracted {len(color_map)} CP product color entries")
    print(f"Saved to: {OUTPUT}")

    # Print summary by type
    types = {}
    for e in entries:
        t = e['type']
        if t not in types:
            types[t] = []
        types[t].append(e)

    print(f"\n--- By Type ---")
    for t, items in sorted(types.items()):
        print(f"\n  {t} ({len(items)}):")
        for e in items:
            print(f"    {e['code']:<6} → {e['color_cn']:<12} | {e['shape_cn']}")

if __name__ == '__main__':
    main()
