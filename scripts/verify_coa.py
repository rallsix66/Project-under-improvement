# -*- coding: utf-8 -*-
"""COA 校对：逐行逐列对比输出文件与基准文件，报告差异"""
import os, sys, re
import win32com.client

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
BASELINE = os.path.join(SKILL_DIR, "resources", "COA基准", "多款-COA-出报告信息确认表.xls")

def read_sheet_data(ws):
    """读取整个 sheet → {(row,col): value_str}"""
    data = {}
    used = ws.UsedRange
    for r in range(1, used.Rows.Count + 1):
        for c in range(1, used.Columns.Count + 1):
            v = ws.Cells(r, c).Value
            if v is not None and str(v).strip():
                data[(r, c)] = str(v).strip()
    return data

def normalize_for_match(s):
    """标准化品名用于匹配：去空格、去标点、全大写"""
    s = re.sub(r'[\s\.\-\#\+]', '', s.upper())
    return s

def extract_products(data):
    """按 Col2 品名标准化后做索引 → {norm_name: (row, {col: value})}"""
    products = {}
    for (r, c), v in data.items():
        if r >= 7:
            if r not in products:
                products[r] = {}
            products[r][c] = v
    result = {}
    for row_num, row_data in products.items():
        name = row_data.get(2, "")
        if name:
            parts = name.split('\n', 1)
            cn = parts[0].strip() if parts else ""
            en = parts[1].strip() if len(parts) > 1 else ""
            norm = normalize_for_match(f"{cn}|{en}")
            result[norm] = (row_num, row_data, name)
    return result

def category_diff(ov, bv):
    """分类差异类型"""
    # 序号差异 → 忽略（产品顺序可能不同）
    # 大小写差异（规格 ml vs ML）
    if ov.upper() == bv.upper():
        return "case"
    # 中英文术语差异
    if ov in ('无','无 ') and bv in ('None','无 '):
        return "term"
    if bv in ('无','无 ') and ov in ('Odorless',):
        return "term"
    return "real"

def compare_delegate(ws_out, ws_base, sheet_name):
    issues = []
    for r in range(2, 6):
        for c in range(1, 10):
            ov = str(ws_out.Cells(r, c).Value or "").strip()
            bv = str(ws_base.Cells(r, c).Value or "").strip()
            if ov != bv:
                issues.append(f"  [{sheet_name}] Row{r} Col{c}: 输出=[{ov[:80]}]  基准=[{bv[:80]}]")
    return issues

def compare_products(products_out, products_base, sheet_name):
    issues = []
    real_issues = []

    out_keys = set(products_out.keys())
    base_keys = set(products_base.keys())

    matched_out = set()
    matched_base = set()

    # 1. 精确匹配
    for bk in base_keys:
        if bk in out_keys:
            matched_base.add(bk)
            matched_out.add(bk)
            _, out_data, out_name = products_out[bk]
            _, base_data, base_name = products_base[bk]
            for c in range(1, 10):
                ov = out_data.get(c, "")
                bv = base_data.get(c, "")
                if ov != bv:
                    dt = category_diff(ov, bv)
                    msg = f"  [{sheet_name}] {base_name[:40]} Col{c}: 输出=[{ov[:60]}]  基准=[{bv[:60]}]"
                    issues.append(msg)
                    if dt == "real":
                        real_issues.append(msg)

    # 2. 模糊匹配
    for bk in base_keys - matched_base:
        _, base_data, base_name = products_base[bk]
        b_cn = normalize_for_match(base_name.split('\n')[0])

        best = None
        for ok in out_keys - matched_out:
            o_cn = normalize_for_match(products_out[ok][2].split('\n')[0])
            if b_cn == o_cn:
                best = ok; break
            if len(b_cn) >= 10 and (b_cn[:15] in o_cn or o_cn[:15] in b_cn):
                best = ok; break

        if best:
            matched_base.add(bk)
            matched_out.add(best)
            _, out_data, out_name = products_out[best]
            for c in range(1, 10):
                ov = out_data.get(c, "")
                bv = base_data.get(c, "")
                if ov != bv:
                    dt = category_diff(ov, bv)
                    msg = f"  [{sheet_name}] {base_name[:40]} Col{c}: 输出=[{ov[:60]}]  基准=[{bv[:60]}]"
                    issues.append(msg)
                    if dt == "real":
                        real_issues.append(msg)
        else:
            # 用 b_cn 再去 out 里模糊匹配
            for ok in out_keys - matched_out:
                o_cn = normalize_for_match(products_out[ok][2].split('\n')[0])
                # 提取核心关键词
                if len(b_cn) >= 6 and len(o_cn) >= 6:
                    if b_cn[:8] == o_cn[:8]:
                        best = ok; break
            if best:
                matched_base.add(bk)
                matched_out.add(best)
                _, out_data, out_name = products_out[best]
                for c in range(1, 10):
                    ov = out_data.get(c, "")
                    bv = base_data.get(c, "")
                    if ov != bv:
                        dt = category_diff(ov, bv)
                        msg = f"  [{sheet_name}] {base_name[:40]} Col{c}: 输出=[{ov[:60]}]  基准=[{bv[:60]}]"
                        issues.append(msg)
                        if dt == "real":
                            real_issues.append(msg)
            else:
                issues.append(f"  [{sheet_name}] 缺失产品: {base_name[:60]}")
                real_issues.append(f"  [{sheet_name}] 缺失产品: {base_name[:60]}")

    for ok in out_keys - matched_out:
        _, _, out_name = products_out[ok]
        issues.append(f"  [{sheet_name}] 多余产品: {out_name[:60]}")
        real_issues.append(f"  [{sheet_name}] 多余产品: {out_name[:60]}")

    return issues, real_issues, len(base_keys), len(out_keys)

def main():
    output_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        SKILL_DIR, "..", "..", "..", "成分表处理脚本", "调试", "COA_test_output.xls")
    output_path = os.path.abspath(output_path)

    if not os.path.exists(BASELINE):
        print(f"[ERROR] Baseline not found: {BASELINE}")
        return
    if not os.path.exists(output_path):
        print(f"[ERROR] Output not found: {output_path}")
        return

    lines = []
    lines.append(f"Baseline: {BASELINE}")
    lines.append(f"Output:   {output_path}")
    lines.append("")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb_out = wb_base = None
    try:
        wb_out = excel.Workbooks.Open(output_path)
        wb_base = excel.Workbooks.Open(BASELINE)

        out_sheets = [ws.Name for ws in wb_out.Worksheets]
        base_sheets = [ws.Name for ws in wb_base.Worksheets]
        lines.append(f"Baseline sheets: {base_sheets}")
        lines.append(f"Output sheets:   {out_sheets}")
        lines.append("")

        total_all = 0
        total_real = 0

        for sn in base_sheets:
            if sn not in out_sheets:
                lines.append(f"=== {sn} === MISSING in output!")
                total_all += 1
                total_real += 1
                continue

            lines.append(f"{'='*70}")
            lines.append(f"=== {sn} ===")
            lines.append(f"{'='*70}")

            ws_out = wb_out.Worksheets(sn)
            ws_base = wb_base.Worksheets(sn)

            # 1. 委托单位
            d_issues = compare_delegate(ws_out, ws_base, sn)
            for iss in d_issues:
                lines.append(iss)

            # 2. 产品
            out_data = read_sheet_data(ws_out)
            base_data = read_sheet_data(ws_base)

            out_prods = extract_products(out_data)
            base_prods = extract_products(base_data)

            all_issues, real_issues, base_count, out_count = compare_products(out_prods, base_prods, sn)
            lines.append(f"  Base: {base_count} products | Output: {out_count} products")
            lines.append(f"  Total diffs: {len(all_issues)} | Real diffs: {len(real_issues)}")
            lines.append("")

            # 分类展示
            case_issues = [i for i in all_issues if i not in real_issues]
            if case_issues:
                lines.append("  --- Minor (case/format) ---")
                for iss in case_issues:
                    lines.append(iss)

            if real_issues:
                lines.append("  --- REAL DIFFERENCES ---")
                for iss in real_issues:
                    lines.append(iss)

            if not all_issues:
                lines.append("  ALL MATCH!")

            total_all += len(all_issues)
            total_real += len(real_issues)
            lines.append("")

        lines.append(f"{'='*70}")
        lines.append(f"SUMMARY: {total_all} total diffs, {total_real} real diffs")
        if total_real == 0:
            lines.append("PASSED - all real content matches!")
        else:
            lines.append("ACTION NEEDED - review real diffs above")

    finally:
        if wb_out: wb_out.Close(False)
        if wb_base: wb_base.Close(False)
        excel.Quit()

    result = '\n'.join(lines)
    print(result)

    # Also save to file
    out_file = os.path.join(SKILL_DIR, "..", "verify_coa_result.txt")
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(result)
    print(f"\nSaved to: {out_file}")

if __name__ == '__main__':
    main()
