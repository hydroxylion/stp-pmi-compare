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
    {
        # 非 tessellated 导出（ctc_05-e1）：报告里**没有**
        # `tessellated_annotation_occurrence`，只有 `draughting_model_item_association`。
        # 旧代码只认前者 → t.ta 全空 → 指标 0%，而装载 / 引用 / name 三项交叉校验全绿。
        # 用例锁死这条等价通道，以及「一条 callout 多行需聚合」的装载校验口径。
        "md": "samples/dev_ctc_05.md",
        "xlsx": "nist_ctc_05_asme1_ap242-e1-sfa.xlsx",
        "recall": 100.0, "precision": 100.0, "datum": 100.0,
        "defects": 8,
        "defect_codes": {"SYM 符号丢失": 8},
    },
    {
        # ftc_08 的**正确**配对：开发侧 markdown 与 `-e2` 报告 52/52 handle 全对上。
        # 提醒：同一测试件的 `-e1-tg` 变体是「图形专用」导出（见 GRAPHIC_ONLY_CASES），
        # 拿它来比会得到满屏 0% —— 本用例同时锁住「换对文件后是什么数字」。
        "md": "samples/dev_ftc_08.md",
        "xlsx": "nist_ftc_08_asme1_ap242-e2-sfa.xlsx",
        "recall": 90.7, "precision": 90.7, "datum": 100.0,
        "defects": 8,
        "defect_codes": {"NUM 数值缺失": 3, "SYM 符号丢失": 5},
    },
    {
        # stc_09：18 个标注同时挂在 MBD_A / MBD_A(Work) 两个保存视图下，
        # 开发侧按「视图 × 标注」导出必然各出一条。
        # 旧代码对「第二次命中同一语义 ID」**无条件**判 ⚠️ 文本差异、不比对内容，
        # 于是产出 12 条「归一化后只差一个空格」的假差异 —— 三项指标明明 100%，
        # 结果表却满屏标黄，观感是「一个都对不上」。
        # 本用例锁死：复用条目照常比对内容，差异数只能剩下有实际原因的 15 条
        # （7 条 CNT 数量前缀 + 8 条复合公差被开发侧合并）。
        "md": "samples/dev_stc_09.md",
        "xlsx": "nist_stc_09_asme1_ap242-e3-sfa.xlsx",
        "recall": 100.0, "precision": 100.0, "datum": 100.0,
        "defects": 7,
        "defect_codes": {"CNT 数量前缀不符": 7},
        "diff": 15,
    },
]

# 图形专用导出：ta 表在，但表头里没有 `Associated Semantic PMI`，整份报告也没有
# `Semantic PMI Summary` —— SFA 侧零真值。此时**必须**判「⛔ 无可比真值」，
# 而不是让开发侧 52 条全部显示「⚠️ 多余、精确率 0%」（那看起来像开发侧全错提取，
# 实际是选错了报告文件）。旧代码在表头识别成功时仍给缺列套默认列号，关联键的默认值
# 10 正好落在 `Saved Views` 上，于是从相机/视图 ID 里抠出 116 条假引用。
GRAPHIC_ONLY_CASES = [
    {
        "md": "samples/dev_ftc_08.md",
        "xlsx": "nist_ftc_08_asme1_ap242-e1-tg-sfa.xlsx",
        # 同目录下应该被指出来的替代报告
        "suggest": "nist_ftc_08_asme1_ap242-e2-sfa.xlsx",
    },
]

# --------------------------------------------------------------------------
# 经确认保留的判定口径：**不改逻辑，只把现状钉死**
#
# 这两类条目从 ftc_08 单独拎出来锁，是因为它们的输出形态容易被误读成
# 「比对没对上」，但它们都是既定口径下的**正确**结果。谁要动它们，得先看懂为什么。
#
# a) Radial Dimension（handle 1231 / 1234 / 1236）：开发侧 `Ø0.237 +.005 / -0.001`，
#    SFA `⌀0.238  +0.005 -0.001` —— 这是**同一个数的不同精度呈现**：开发侧
#    detailData 原始值 `0.237500`、value format qualifier `NR2 1.3`（1 位整数
#    3 位小数）→ 渲染时截断成 `0.237`；SFA 四舍五入成 `0.238`。
#    关联链本身是通的（path=dim-2hop、语义 ID 命中 525/526/527），所以不是「对不上」；
#    是 `_num_equivalent` 按 min(位数) 比较 → 判不等 → 缺陷层记 `NUM 数值缺失`、
#    匹配层落 `⚠️ 疑似（人工复核）`。
#    **保持不改**：放宽成「数值近似即等价」会让 `0.15 vs 0.1` 这类真差异被静默吞掉，
#    损失的是检出能力，换来的只是图上几行黄。宁可人工复核。
#
# b) `⏥ .015 L1<#h>L2`（handle 1248）：开发侧 description 里的 `#h` 是**未替换的
#    占位符**（已确认为开发侧缺陷），`L1` / `L2` 是被并进标题的关联标签；
#    SFA 侧 `▱ | .015` 自然没有这些文本。当前缺陷层只覆盖
#    「编码 / 空 / 数量 / 符号 / 数值 / 映射 / 重复」七类，**「多出文本」不在其中**，
#    所以这条 defects 为空；真正把它兜住的是匹配层 → `⚠️ 疑似（人工复核）`。
#    **按此现状比对**：不为它新增「多余文本」判据 —— SFA 文本是排版拼装出来的，
#    这类判据的误报面远大于收益，且 `#h` 该由开发侧修提取，不该由比对工具兜底。
KNOWN_VERDICTS = [
    {
        "md": "samples/dev_ftc_08.md",
        "xlsx": "nist_ftc_08_asme1_ap242-e2-sfa.xlsx",
        # handle -> (期望状态, 期望缺陷码集合, 期望关联路径)
        "lock": {
            # 舍入口径差异：关联命中，判疑似 + 记数值缺失
            1231: (core.ST_SUSPECT, {core.DF_NUM}, "dim-2hop"),
            1234: (core.ST_SUSPECT, {core.DF_NUM}, "dim-2hop"),
            1236: (core.ST_SUSPECT, {core.DF_NUM}, "dim-2hop"),
            # `#h` 占位符 + 并进标题的标签：关联命中，判疑似，**不进缺陷层**
            1248: (core.ST_SUSPECT, set(), "gt-1hop"),
        },
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

        # 判定层：差异条数锁死。多视图复用若被误判成差异，数字会明显偏大
        # （stc_09 是 26 vs 15），这是「复用 ≠ 差异」这道口径的直接闸门。
        if "diff" in case:
            n_diff = sum(1 for r in rows if r.status == core.ST_HIT_DIFF)
            check(f"{tag} 文本差异条数 = {case['diff']}", n_diff == case["diff"], n_diff)

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

    # ---- 图形专用导出：无真值要判「不可比」，不能判「多余」 ----
    for gc in GRAPHIC_ONLY_CASES:
        xlp = os.path.join(SFA_DIR, gc["xlsx"])
        tag = gc["xlsx"].replace("-sfa.xlsx", "")
        print(f"=== {tag}（图形专用导出）===")
        if not os.path.exists(xlp):
            print(f"  跳过：SFA 报告不存在 {xlp}（可用 PMI_SFA_DIR 指定目录）")
            continue
        ran += 1

        t = core.load_sfa(xlp)
        raw = open(os.path.join(HERE, gc["md"].replace("/", os.sep)),
                   encoding="utf-8", errors="replace").read()
        items, _ = core.parse_dev_markdown_ex(raw)
        rows = core.match_items(t, items)
        m = core.compute_metrics(rows)
        dev = [r for r in rows if not r.key.startswith("SFA") and r.kind != core.KIND_NOTE]

        check(f"{tag} 判定为图形专用导出", t.graphic_only, True)
        check(f"{tag} 判定为无真值", t.truth_missing and m.truth_missing,
              (t.truth_missing, m.truth_missing))
        check(f"{tag} 关联通道判定为不可用", not t.link_channel, t.link_channel)
        check(f"{tag} 不按默认列号抠出 Saved Views 的假引用",
              not [r for r in t.ta.values() if r["sem_refs"]],
              [r["sem_refs"] for r in t.ta.values() if r["sem_refs"]][:3])
        check(f"{tag} 开发侧全判「⛔ 无可比真值」",
              {r.status for r in dev} == {core.ST_ABSENT},
              sorted({r.status for r in dev}))
        check(f"{tag} 不产生任何「⚠️ 多余」", m.sem_extra == 0, m.sem_extra)
        check(f"{tag} 不进语义指标分母",
              m.sem_expected == 0 and m.datum_expected == 0,
              (m.sem_expected, m.datum_expected))
        check(f"{tag} 不可比条目数 = 开发侧非注释条目数", m.no_truth == len(dev),
              (m.no_truth, len(dev)))
        check(f"{tag} 体检点名关联键列与语义真值",
              sorted(c.name for c in t.broken_checks() if not c.ok)
              == ["draughting_model_item_associati 装载", "ta 引用可解释",
                  "图形标注关联通道", "图形标注关联键列", "语义真值"],
              sorted(c.name for c in t.broken_checks() if not c.ok))
        check(f"{tag} 指出同目录含语义 PMI 的替代报告",
              any(gc["suggest"] in w for w in t.warnings), t.warnings)
        print(f"  开发侧 {len(dev)} 条 → 全部「{core.ST_ABSENT}」；"
              f"提示改用 {gc['suggest']}")
        print()

    # ---- 经确认保留的判定口径：锁住现状，防止被「顺手改好」----
    for kc in KNOWN_VERDICTS:
        xlp = os.path.join(SFA_DIR, kc["xlsx"])
        tag = kc["xlsx"].replace("-sfa.xlsx", "")
        print(f"=== {tag}（既定口径锁定）===")
        if not os.path.exists(xlp):
            print(f"  跳过：SFA 报告不存在 {xlp}（可用 PMI_SFA_DIR 指定目录）")
            continue
        ran += 1

        t = core.load_sfa(xlp)
        raw = open(os.path.join(HERE, kc["md"].replace("/", os.sep)),
                   encoding="utf-8", errors="replace").read()
        items, _ = core.parse_dev_markdown_ex(raw)
        rows = core.match_items(t, items)

        for h, (want_status, want_codes, want_path) in sorted(kc["lock"].items()):
            hit = [r for r in rows if r.handle == h]
            if not hit:
                check(f"{tag} handle {h} 存在", False, "结果表里没有这个 handle")
                continue
            got_status = sorted({r.status for r in hit})
            got_codes = {c for r in hit for c in r.defects}
            got_path = sorted({r.path for r in hit})
            check(f"{tag} handle {h} 状态 = {want_status}", got_status == [want_status],
                  got_status)
            check(f"{tag} handle {h} 缺陷码 = {sorted(want_codes) or '（无）'}",
                  got_codes == want_codes, sorted(got_codes) or "（无）")
            check(f"{tag} handle {h} 关联路径 = {want_path}", got_path == [want_path],
                  got_path)
        locked = len(kc["lock"])
        print(f"  锁定 {locked} 个 handle：舍入口径差异 3 条判「疑似」+ 记数值缺失，"
              f"`#h` 占位符 1 条判「疑似」但不进缺陷层")
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
