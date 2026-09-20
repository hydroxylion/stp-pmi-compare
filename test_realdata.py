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
