# -*- coding: utf-8 -*-
"""SFA 报告 + 开发 markdown 一键诊断。

比对结果不对劲时（一条都没对上、指标异常、缺陷像是虚警）先跑这个：
把「工具到底读到了什么、列是按什么定位的、引用能不能落地、哪一环断了」
一次输出完整上下文，不必反复来回确认。

用法：
    python doctor.py <SFA报告.xlsx>
    python doctor.py <SFA报告.xlsx> <开发markdown>
    python doctor.py <SFA报告.xlsx> <开发markdown> -o 诊断报告.txt

输出同时打印到终端并写入 doctor_report.txt（用 -o 改路径）。
"""
import argparse
import collections
import datetime
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import openpyxl                      # noqa: E402
import pmi_core as core              # noqa: E402

CJK = lambda ch: 0x4E00 <= ord(ch) <= 0x9FFF      # noqa: E731


def known_symbols():
    """pmi_core 里已登记的字符，用于挑出「没见过的」。

    含五类：形位公差符号（_FCF_CHARS）、符号别名、字形归一化表（SYMBOL_MAP，
    含 `±`/`×` 这类通用字符）、残缺乱码定点表、SFA 排版字形。
    后四类都不是需要新增的公差符号，报出来只会产生噪音。
    """
    out = set()
    for name in ("_FCF_CHARS", "SFA_LAYOUT_GLYPHS"):
        out |= set(getattr(core, name, ()) or ())
    for name in ("_SYM_ALIAS", "SYMBOL_MAP", "_MOJIBAKE_FIXED", "_CP936_PATCH"):
        out |= set((getattr(core, name, {}) or {}).keys())
    out |= set("⌀ⓂⓁⓅⓊⓈ")
    return out


def rare_chars(t):
    """统计 SFA 文本里的非 ASCII 字符（排除 CJK），挑出符号表里没有的。"""
    cnt = collections.Counter()
    for s in t.semantic.values():
        cnt.update(ch for ch in str(s.get("text") or "") if ord(ch) > 127)
    for v in t.ta.values():
        cnt.update(ch for ch in str(v.get("text") or "") if ord(ch) > 127)
    return cnt


def main():
    ap = argparse.ArgumentParser(description="SFA 报告 + 开发 markdown 一键诊断")
    ap.add_argument("xlsx", help="SFA 导出的 xlsx 报告")
    ap.add_argument("md", nargs="?", default="", help="开发侧 markdown（可选）")
    ap.add_argument("-o", "--out", default="", help="诊断报告输出路径")
    args = ap.parse_args()

    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:               # noqa: BLE001
        pass

    if not os.path.exists(args.xlsx):
        print(f"找不到 SFA 报告：{args.xlsx}")
        return 2

    L = []

    def out(s=""):
        L.append(s)
        print(s)

    def head(title):
        out()
        out("=" * 72)
        out(title)
        out("=" * 72)

    # ---------------- 0. 环境 ----------------
    head("0. 环境")
    out(f"时间       : {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    out(f"Python     : {sys.version.split()[0]}  ({sys.executable})")
    out(f"工作目录   : {HERE}")
    for mod in ("openpyxl", "pandas"):
        try:
            m = __import__(mod)
            out(f"{mod:11s}: {getattr(m, '__version__', '?')}")
        except Exception:           # noqa: BLE001
            out(f"{mod:11s}: 未安装")
    out(f"SFA 报告   : {args.xlsx}")
    out(f"开发 markdown: {args.md or '（未提供，跳过比对相关章节）'}")

    # ---------------- 1. 原始表结构 ----------------
    wb = openpyxl.load_workbook(args.xlsx, read_only=True, data_only=True)
    sheets = list(wb.sheetnames)
    head("1. 工作表清单")
    out(f"共 {len(sheets)} 张表")
    for i, n in enumerate(sheets, 1):
        out(f"  {i:3d}. {n}")

    # ---------------- 2. 装载与列定位 ----------------
    t = core.load_sfa(args.xlsx)

    head("2. 关键表结构与列定位")
    out("说明：列位置一律按表头列名定位。来源显示 [fallback] 表示没识别出表头，")
    out("      已退回默认列号 —— 该表列序一变就会静默失配。")
    for cr in t.recipes:
        out()
        out(f"--- {cr.sheet}  (逻辑名 {cr.key})")
        out(f"    列数 {cr.n_cols} | 表头行 {'第 %d 行' % (cr.header_row + 1) if cr.header_row >= 0 else '未识别'}"
            f" | 来源 [{'header' if cr.trustworthy else 'fallback'}]")
        out(f"    数据行 {cr.n_rows} | 装载 {cr.n_loaded}")
        out(f"    用到的列：{cr.resolved}")
        if cr.mapping:
            out(f"    表头全文：{cr.mapping}")
        for n in cr.notes:
            out(f"    ! {n}")
        # 原始前几行，便于人工核对列序
        try:
            ws = wb[cr.sheet]
            out("    原始前 5 行：")
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i >= 5:
                    break
                cells = [f"[{j}]{str(c)[:26]}" for j, c in enumerate(r) if c is not None]
                out(f"      r{i}: {'  '.join(cells) if cells else '(空行)'}")
        except Exception as e:      # noqa: BLE001
            out(f"    （原始行读取失败：{e}）")

    # ---------------- 3. ID 索引规模 ----------------
    head("3. ID 索引规模")
    out(f"语义表条目        : {len(t.semantic)}")
    out(f"draughting_callout: {len(t.dc)}")
    out(f"图形标注(ta)      : {len(t.ta)}   (ta 表列数 {t.ta_cols})")
    out(f"datum             : {len(t.datum)}")
    out(f"dcr               : {len(t.dcr_by_dim)}")
    out(f"单位              : {t.units or '未识别'}")
    out(f"语义表标记行      : {t.sentinel_row}")
    if t.semantic:
        kinds = collections.Counter(v["kind"] for v in t.semantic.values())
        out(f"语义实体分类      : {dict(kinds)}")

    # ---------------- 4. 交叉校验 ----------------
    head("4. 加载期交叉校验")
    bad = t.broken_checks()
    out(f"{len(t.checks) - len(bad)}/{len(t.checks)} 通过")
    for ck in t.checks:
        out("  " + ck.line())

    # ---------------- 5. 装载告警 ----------------
    head("5. 装载告警")
    if t.warnings:
        for w in t.warnings:
            out(f"  ! {w}")
    else:
        out("  （无）")

    # ---------------- 6. 未收录字符 ----------------
    head("6. 未收录字符（可能是新的 GD&T 符号或 SFA 渲染字形）")
    known = known_symbols()
    cnt = rare_chars(t)
    unknown = [(ch, n) for ch, n in cnt.most_common() if ch not in known and not CJK(ch)]
    if unknown:
        out("以下字符出现在文本里，但不在 pmi_core 的符号表内：")
        for ch, n in unknown:
            out(f"  {ch!r}  U+{ord(ch):04X}  ×{n}")
        out()
        out("若其中某个是形位公差符号（不是 SFA 的渲染杂符），把它补进 pmi_core 的")
        out("_FCF_CHARS 与 _SYM_ALIAS，否则该符号会被漏检或误报「缺符号」。")
    else:
        out("  （无：文本里出现的字符都已在符号表内）")
    if cnt:
        out()
        out("全部非 ASCII 字符统计（含已知）：")
        out("  " + "  ".join(f"{ch}×{n}" for ch, n in cnt.most_common(30)))

    wb.close()

    # ---------------- 7. 开发 markdown ----------------
    head("7. 开发 markdown 解析")
    if not args.md:
        out("  （未提供 markdown，跳过）")
        out("  提示：提供 markdown 后可一并输出解析自检、关联路径分布与三项指标。")
    elif not os.path.exists(args.md):
        out(f"  找不到 markdown：{args.md}")
    else:
        raw = open(args.md, encoding="utf-8", errors="replace").read()
        items, diag = core.parse_dev_markdown_ex(raw)
        out(f"  文件大小 {len(raw)} 字符 / {len(raw.splitlines())} 行")
        out(f"  解析条目      : {diag.items}")
        out(f"  带 handle     : {diag.with_handle}")
        out(f"  带 name       : {diag.with_name}")
        out(f"  未解析到 detailData: {len(diag.detail_missing)}")
        out(f"  JSON 解析失败 : {len(diag.json_failed)}")
        out(f"  缺关键字段    : {len(diag.key_missing)}")
        out(f"  识别分组      : {diag.groups}")
        probs = diag.problems()
        if probs:
            out()
            out("  问题：")
            for p in probs:
                out(f"    ! {p}")
        for label, lst in (("未解析到 detailData", diag.detail_missing),
                           ("JSON 失败", diag.json_failed),
                           ("缺 handle/name", diag.key_missing),
                           ("像标题但未匹配", diag.headings_unmatched)):
            if lst:
                out()
                out(f"  {label}（{len(lst)} 条，列前 20）：")
                for x in lst[:20]:
                    out(f"    - {x}")

        # ---- 8. 比对结果 ----
        head("8. 比对结果")
        rows = core.match_items(t, items)
        m = core.compute_metrics(rows)
        out(f"比对行数     : {len(rows)}")
        out(f"状态分布     : {dict(collections.Counter(r.status for r in rows))}")
        out(f"关联路径分布 : {core.link_stats(rows)}")
        out()
        out(f"召回率 {m.recall:.2f}% | 精确率 {m.precision:.2f}% | 基准覆盖率 {m.datum_coverage:.2f}%")
        out(f"  语义: 期望 {m.sem_expected} / 提取 {m.sem_extracted} / 命中 {m.sem_hit}"
            f" / 缺失 {m.sem_miss} / 多余 {m.sem_extra} / 差异 {m.sem_diff} / 疑似 {m.sem_suspect}")
        out(f"  基准: 期望 {m.datum_expected} / 命中 {m.datum_hit}")
        out(f"  注释: 共 {m.note_total} / 图形有文本 {m.note_graphic_only} / 内部独占 {m.note_exclusive}")
        out(f"  缺陷: 条目 {m.defect_rows} / 总数 {m.defect_total}"
            f" / 占比 {m.defect_rate:.2f}% {m.defect_by_code}")

        seen = set()
        defects = []
        for r in rows:
            if not r.defects:
                continue
            base = r.key.split("@")[0]
            if base in seen:
                continue
            seen.add(base)
            defects.append((base, r.dev_title, r.defects, r.defect_detail))
        if defects:
            out()
            out(f"提取缺陷清单（{len(defects)} 条）：")
            for base, title, codes, detail in defects:
                out(f"  {base:12s} {title[:26]:28s} {' '.join(codes)} | {detail}")

        suspects = core.suspect_links(items, t, rows)
        if suspects:
            out()
            out(f"疑似对应（{len(suspects)} 条，仅供参考、不参与判定）：")
            for sl in suspects:
                out(f"  {sl.key:12s} {sl.title[:24]:26s} → {sl.text()}")

        bad_rows = [r for r in rows if r.status not in (core.ST_HIT, core.ST_NOTE)]
        if bad_rows:
            out()
            out(f"非命中条目（{len(bad_rows)} 条，列前 40）：")
            for r in bad_rows[:40]:
                out(f"  {r.status:16s} {r.key:18s} {r.dev_title[:24]:26s}"
                    f" SFA={r.sfa_text[:30]!r} path={r.path} {' '.join(r.defects)}")

    # ---------------- 输出 ----------------
    dst = args.out or os.path.join(os.getcwd(), "doctor_report.txt")
    try:
        with open(dst, "w", encoding="utf-8") as f:
            f.write("\n".join(L) + "\n")
        print()
        print(f"诊断报告已写入：{dst}")
        print("如果结果不符合预期，把这份文件发给维护者即可，无需再补充其他信息。")
    except Exception as e:          # noqa: BLE001
        print(f"写入失败：{e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
