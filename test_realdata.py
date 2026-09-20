# -*- coding: utf-8 -*-
"""真实数据回归：用开发侧导出的 markdown + SFA 报告跑端到端，锁定期望指标。

与 test_pmi_core.py 的分工：那个用构造样张覆盖规则细节，这个用**真实语料**
锁住端到端数值 —— 真实数据才会踩出列数变化、字形差异、映射缺口这类问题。

语料不进仓库（体积大、含内部数据），数据缺失时自动跳过，不阻塞。
放置方式见 samples/README.md。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import openpyxl
import pmi_core as core              # noqa: E402

SFA_DIR = os.environ.get(
    "PMI_SFA_DIR",
    r"F:\1【机械零件】\Step官方标准数据\NIST-PMI-STEP-Files",
)

CASES = [
    {
        "md": "samples/dev_ftc_10.md",
        "xlsx": "nist_ftc_10_asme1_ap242-e2-sfa.xlsx",
        "recall": 100.0, "precision": 100.0, "datum": 100.0,
        "defects": 6,
        "defect_codes": {"SYM 符号丢失": 4, "CNT 数量前缀不符": 2},
    },
    {
        # CTC 系列：实体 ID 只有 2~3 位（FTC 是 4~7 位），且 datum 与 datum_feature
        # 的 ID 不相邻（差 3 而非 1）。这两个差异曾让比对直接 0% 命中，
        # 用例放在这里是为了锁死「不写死 ID 位数、不用 ID 算术推基准」这两条规则。
        "md": "samples/dev_ctc_01.md",
        "xlsx": "nist_ctc_01_asme1_ap242-e1-sfa.xlsx",
        "recall": 100.0, "precision": 100.0, "datum": 100.0,
        "defects": 0,
        "defect_codes": {},
    },
]

FAIL = []
TOTAL = 0


# 多视图复用用例：开发侧导出粒度是「保存视图 × 标注」，一条 annotation 挂在几个
# 视图下就会导出几条（handle 相同）。NIST ctc_02 的基准目标正好挂在 MBD_A + MBD_B，
# 早期实现按 handle 一刀切判 DUP，会凭空报出 16 条「重复」缺陷。
MULTIVIEW_CASES = [
    {
        "xlsx": "nist_ctc_02_asme1_ap242-e2-sfa.xlsx",
        "annotations": ["Datum Target.1", "Datum Target.2", "Datum Target.3",
                        "Datum Target.4", "Datum Target.5", "Datum Target.6",
                        "Datum Target.7", "Datum Target.8"],
        "views": "MBD_A、MBD_B",
    },
]

# 基准目标用例：同一实体在 SFA 报告里有两种写法 —— ftc_10 写 `datum_target`，
# ctc_02 / ctc_05 / ftc_06 写 `placed_datum_target_feature`。只认一种会让整批
# 基准目标静默断链：引用完全能落到语义表（装载/引用交叉校验全绿），但判据不认
# 实体名，结果开发侧全判「多余」、SFA 侧全判「缺失」。
# 另外标识两侧常有干扰 —— `A1 (point)` 的目标形式、`⌀85\nK1` 的目标尺寸，
# 按位置取首片段会拿到尺寸而不是标识。
DATUM_TARGET_CASES = [
    {
        "xlsx": "nist_ctc_02_asme1_ap242-e2-sfa.xlsx",
        # name -> (基准字母, 目标序号)。序号取自 STP 的
        # PLACED_DATUM_TARGET_FEATURE 第 5 个参数，字母是图纸上的基准符号。
        "targets": {
            "Datum Target.1": ("A", "1"), "Datum Target.2": ("A", "2"),
            "Datum Target.3": ("A", "3"), "Datum Target.4": ("B", "1"),
            "Datum Target.5": ("B", "2"), "Datum Target.6": ("B", "3"),
            "Datum Target.7": ("B", "4"), "Datum Target.8": ("C", "1"),
            "Datum Target.10": ("K", "1"),
        },
        # 语义文本带干扰的两条：目标尺寸前置（⌀85，circle 形式）
        "noisy_text": {"Datum Target.10": "D85 K1"},
    },
]


def _multiview_probe(xlsx: str, t):
    """按 SFA 的 `Saved Views` 列反推「开发侧按视图导出」，返回 (条目, 比对行)。"""
    import re
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    sname = core._pick_sheet(wb.sheetnames, "tessellated_annotation_occurren")
    rows = core._rows(wb[sname])
    hi, hdr = next((i, r) for i, r in enumerate(rows)
                   if any(c.startswith("Saved Views") for c in r))
    i_name = hdr.index("name")
    i_view = next(i for i, c in enumerate(hdr) if c.startswith("Saved Views"))
    name2id = {v: k for k, v in t.dc.items()}

    items, seq = [], {}
    for r in rows[hi + 1:]:
        nm = r[i_name]
        if not nm or nm not in name2id:
            continue
        for g in (re.findall(r"\((MBD_[A-Z]+)\)", r[i_view]) or ["MBD_A"]):
            seq[g] = seq.get(g, 0) + 1
            items.append(core.DevItem(group=g, seq=seq[g], title=nm,
                                      handle=name2id[nm], name=nm))
    return items, core.match_items(t, items)


def check(name, cond, got=""):
    global TOTAL
    TOTAL += 1
    if cond:
        print(f"  ok   {name}")
    else:
        FAIL.append(f"{name}（实际：{got}）")
        print(f"  FAIL {name}（实际：{got}）")


def main():
    ran = 0
    for case in CASES:
        mdp = os.path.join(HERE, case["md"].replace("/", os.sep))
        xlp = os.path.join(SFA_DIR, case["xlsx"])
        tag = os.path.basename(case["md"])
        print(f"=== {tag} ===")
        if not os.path.exists(mdp):
            print(f"  跳过：语料不存在 {mdp}")
            continue
        if not os.path.exists(xlp):
            print(f"  跳过：SFA 报告不存在 {xlp}（可用 PMI_SFA_DIR 指定目录）")
            continue
        ran += 1

        raw = open(mdp, encoding="utf-8", errors="replace").read()
        t = core.load_sfa(xlp)
        items, diag = core.parse_dev_markdown_ex(raw)
        rows = core.match_items(t, items)
        m = core.compute_metrics(rows)
        paths = core.link_stats(rows)

        # 加载期：交叉校验必须全过（列定位、装载条数、引用可解释）
        broken = t.broken_checks()
        check(f"{tag} 加载期交叉校验全通过",
              not broken, [b.line() for b in broken])

        # 解析层：真实语料必须 100% 解析出 handle/name
        check(f"{tag} 解析条目 = 带 handle 数",
              diag.items and diag.items == diag.with_handle,
              f"{diag.items}/{diag.with_handle}")
        check(f"{tag} 解析条目 = 带 name 数",
              diag.items and diag.items == diag.with_name,
              f"{diag.items}/{diag.with_name}")

        # 关联层：不该有断链条目
        check(f"{tag} 无断链（none）条目", paths.get("none", 0) == 0, paths)

        # 指标：数值被锁死，改坏即报警
        check(f"{tag} 召回率 = {case['recall']}%", round(m.recall, 2) == case["recall"],
              round(m.recall, 2))
        check(f"{tag} 精确率 = {case['precision']}%", round(m.precision, 2) == case["precision"],
              round(m.precision, 2))
        check(f"{tag} 基准覆盖率 = {case['datum']}%",
              round(m.datum_coverage, 2) == case["datum"], round(m.datum_coverage, 2))

        # 缺陷层：条数与分布都要对得上（虚警与漏报都会让这里失败）
        check(f"{tag} 缺陷条数 = {case['defects']}", m.defect_total == case["defects"],
              m.defect_total)
        check(f"{tag} 缺陷分布一致", m.defect_by_code == case["defect_codes"],
              m.defect_by_code)

        print(f"  状态 {dict((k, sum(1 for r in rows if r.status == k)) for k in set(r.status for r in rows))}")
        print(f"  路径 {paths}")
        print()

    # ---- 多视图复用：跨视图重复不是缺陷，不能报 DUP ----
    for mc in MULTIVIEW_CASES:
        xlp = os.path.join(SFA_DIR, mc["xlsx"])
        tag = mc["xlsx"].replace("-sfa.xlsx", "")
        print(f"=== {tag}（多视图复用）===")
        if not os.path.exists(xlp):
            print(f"  跳过：SFA 报告不存在 {xlp}（可用 PMI_SFA_DIR 指定目录）")
            continue
        ran += 1

        t = core.load_sfa(xlp)
        items, mrows = _multiview_probe(xlp, t)
        dup = sorted({r.key for r in mrows if core.DF_DUP in r.defects})
        mv_names = sorted({r.dev_name for r in mrows if r.multiview})
        mv_views = sorted({r.multiview for r in mrows if r.multiview})

        check(f"{tag} 跨视图复用不报 DUP", not dup, dup[:8])
        check(f"{tag} 多视图标注名单一致", mv_names == mc["annotations"], mv_names)
        check(f"{tag} 多视图分组一致", mv_views == [mc["views"]], mv_views)
        check(f"{tag} 多视图记录判定一致（同一标注不因视图而变）",
              all(len({r.status for r in mrows if r.dev_name == nm}) == 1
                  for nm in mc["annotations"]),
              {nm: sorted({r.status for r in mrows if r.dev_name == nm})
               for nm in mc["annotations"]})
        check(f"{tag} 缺陷层不被多视图记录污染",
              not any(r.defects for r in mrows if r.dev_name in mc["annotations"]),
              [r.key for r in mrows if r.dev_name in mc["annotations"] and r.defects])
        print(f"  条目 {len(items)} 条 → 多视图标记 {len(mv_names)} 个标注")
        print()

    # ---- 基准目标：实体名漂移 + 标识两侧缀 ----
    for dc in DATUM_TARGET_CASES:
        xlp = os.path.join(SFA_DIR, dc["xlsx"])
        tag = dc["xlsx"].replace("-sfa.xlsx", "")
        print(f"=== {tag}（基准目标）===")
        if not os.path.exists(xlp):
            print(f"  跳过：SFA 报告不存在 {xlp}（可用 PMI_SFA_DIR 指定目录）")
            continue
        ran += 1

        t = core.load_sfa(xlp)
        items, sids = [], {}
        for i, (nm, (lt, tid)) in enumerate(dc["targets"].items(), 1):
            items.append(core.DevItem(
                group="MBD_A", seq=i, title=f"Datum Target {lt}{tid}", name=nm,
                types=["datum_target"],
                detail={"datum": lt, "target id": tid, "type": "datum_target"}))
            sids[nm] = (t.ta.get(nm) or {}).get("sem_refs", [])
        mrows = core.match_items(t, items)
        mine = [r for r in mrows if r.dev_name in dc["targets"]]
        stuck = [r.sfa_sem_id for r in mrows
                 if r.status == core.ST_MISS and r.sfa_sem_id in
                 {s for v in sids.values() for s in v}]

        check(f"{tag} 基准目标全部命中",
              all(r.status == core.ST_HIT for r in mine),
              {r.dev_name: r.status for r in mine if r.status != core.ST_HIT})
        check(f"{tag} 基准目标走 datum_target-1hop 通道",
              {r.path for r in mine} == {"datum_target-1hop"},
              {r.path for r in mine})
        check(f"{tag} 基准目标不产生反向缺失", not stuck, stuck)
        check(f"{tag} 实体名漂移不影响体检结论",
              not [c for c in t.broken_checks() if c.name == "语义表实体归类"],
              [c.line() for c in t.broken_checks()])
        for nm, want in dc.get("noisy_text", {}).items():
            got = next((r.sfa_text for r in mine if r.dev_name == nm), "")
            check(f"{tag} {nm} 语义文本原样呈现（含目标尺寸）", got == want, got)
        print(f"  基准目标 {len(mine)} 条全部命中，语义实体 "
              f"{[s for v in sids.values() for s in v]}")
        print()

    print()
    if not ran:
        print("没有可跑的真实语料，已跳过。")
        return 0
    if FAIL:
        print(f"未通过 {len(FAIL)} / {TOTAL}")
        for f in FAIL:
            print("  -", f)
        return 1
    print(f"真实数据回归全部通过（{TOTAL}/{TOTAL}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
