# -*- coding: utf-8 -*-
"""pmi_core 回归测试。

用法::

    python test_pmi_core.py                    # 只跑纯函数单测（无需外部文件）
    python test_pmi_core.py <SFA报告.xlsx>      # 追加集成测试（真实数据端到端）

纯函数单测：乱码修复 / 数值归一 / 语义指纹 / 比对判定 / markdown 解析
集成测试  ：SFA 五表索引 + 三条 ID 关联路径 + 三层指标
"""
import io, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pmi_core as C

DEFAULT_XLSX = r"F:\1【机械零件】\Step官方标准数据\NIST-PMI-STEP-Files\nist_ftc_07_asme1_ap242-e2-sfa-1.xlsx"
XLSX = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX

# (group, handle, name, title, type, extra)
ROWS = [
 ("MBD_A",13706,"Simple Datum.1","A","composite_group_shape_aspect+composite_shape_aspect+datum_feature+shape_aspect",{"datum":"A"}),
 ("MBD_A",13747,"Simple Datum.2","B","composite_shape_aspect+datum_feature+shape_aspect",{"datum":"B"}),
 ("MBD_A",13788,"Simple Datum.3","C","composite_shape_aspect+datum_feature+shape_aspect",{"datum":"C"}),
 ("MBD_A",13928,"Geometrical Tolerance.1","⌓ .06 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_A",13959,"Profile tolerance of any surface.1","⌓ .015","surface_profile_tolerance",{}),
 ("MBD_A",13999,"Perpendicularity.1","⟂ Ø.015Ⓟ.425 A","perpendicularity_tolerance",{}),
 ("MBD_A",14040,"Position.1","⌖ Ø.02Ⓟ.425 A B","geometric_tolerance+geometric_tolerance_with_datum_reference+position_tolerance",{}),
 ("MBD_A",14101,"Position.2","⌖ Ø.023Ⓟ.425 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+position_tolerance",{}),
 ("MBD_A",14233,"Position.3","⌖ Ø.023Ⓟ.15 D B C","geometric_tolerance+geometric_tolerance_with_datum_reference+position_tolerance",{}),
 ("MBD_A",15146,"Linear Size.1","Ø0.250 +.003 / -0","dimensional_size",{"size type":"diameter"}),
 ("MBD_A",15209,"Oriented Linear Dimension.1","0.30 ±.02","dimensional_location",{}),
 ("MBD_A",15254,"Linear Size.2","Ø0.250 +.003 / -0","dimensional_size",{"size type":"diameter"}),
 ("MBD_A",15315,"Oriented Linear Dimension.2","0.30 ±.02","dimensional_location",{}),
 ("MBD_A",15360,"Linear Size.3","2× Ø0.250 +.003 / -0","dimensional_size",{"size type":"diameter"}),
 ("MBD_A",15405,"Linear Size.4","4× Ø0.250 +.003 / -0.001","dimensional_size",{"size type":"diameter"}),
 ("MBD_A",15506,"Oriented Linear Dimension.3","0.30 ±.02","dimensional_location",{}),
 ("MBD_A",15578,"Oriented Linear Dimension.4","0.30 ±.02","dimensional_location",{}),
 ("MBD_A",16261,"Text.1","Text.1","property_definition",{"property":"semantic text"}),
 ("MBD_A",16288,"Text.2","F1.2","property_definition",{"property":"semantic text"}),
 ("MBD_A",16316,"Text.3","F1.1","property_definition",{"property":"semantic text"}),
 ("MBD_A",16345,"Text.4","Text.4","property_definition",{"property":"semantic text"}),
 ("MBD_B",13837,"Simple Datum.4","D","composite_group_shape_aspect+composite_shape_aspect+datum_feature+shape_aspect",{"datum":"D"}),
 ("MBD_B",14143,"Position surfacic profile.1","⌓ .04 A B C .01 A","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_B",14269,"Flatness.1","⏥ .01","flatness_tolerance+geometric_tolerance+geometric_tolerance_with_defined_area_unit+geometric_tolerance_with_defined_unit",{}),
 ("MBD_B",14302,"Parallelism.1",".03 A","parallelism_tolerance",{}),
 ("MBD_B",14335,"Position surfacic profile.2","⌓ .06 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_B",14405,"Position.4","⌖ Ø.05Ⓢ A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+geometric_tolerance_with_modifiers+position_tolerance",{"modifiers":"(.STATISTICAL_TOLERANCE.)"}),
 ("MBD_B",14498,"Position.5","⌖ Ø.06 A B C Ø.015 A","geometric_tolerance+geometric_tolerance_with_datum_reference+position_tolerance",{}),
 ("MBD_B",14558,"Position.6","⌖ Ø.04Ⓜ A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+geometric_tolerance_with_modifiers+position_tolerance",{"modifiers":"(.MAXIMUM_MATERIAL_REQUIREMENT.)"}),
 ("MBD_B",14593,"Position.7","⌖ Ø.045 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+position_tolerance",{}),
 ("MBD_B",14735,"Position.8","⌖ Ø.06Ⓜ E B C Ø.02Ⓜ E B C","geometric_tolerance+geometric_tolerance_with_datum_reference+geometric_tolerance_with_modifiers+position_tolerance",{"modifiers":"(.MAXIMUM_MATERIAL_REQUIREMENT.)"}),
 ("MBD_B",15623,"Linear Size.5","3× Ø0.875 ±.02","dimensional_size",{"size type":"diameter","dimensional note":"statistical"}),
 ("MBD_B",15669,"Linear Size.6","3× Ø0.875 ±.01","dimensional_size",{"size type":"diameter"}),
 ("MBD_B",15715,"Linear Size.7","4× Ø0.173 +.005 / -0.001","dimensional_size",{"size type":"diameter"}),
 ("MBD_B",15761,"Linear Size.8","2× Ø0.594 +.012 / -0.002","dimensional_size",{"size type":"diameter"}),
 ("MBD_B",15807,"Linear Size.9","4× Ø0.281 +.006 / -0.001","dimensional_size",{"size type":"diameter"}),
 ("MBD_B",15899,"Linear Size.11","4× Ø0.438 ±.01","dimensional_size",{"size type":"diameter"}),
 ("MBD_B",15974,"Oriented Linear Dimension.5","0.30 ±.02","dimensional_location",{}),
 ("MBD_B",16374,"Text.5","Text.5","property_definition",{"property":"semantic text"}),
 ("MBD_B",16403,"Text.6","D","property_definition",{"property":"semantic text"}),
 ("MBD_B",16431,"Text.7","飦","property_definition",{"property":"semantic text"}),
 ("MBD_B",16490,"Text.9","鈱´","property_definition",{"property":"semantic text"}),
 ("MBD_C",13887,"Simple Datum.5","E","composite_group_shape_aspect+composite_shape_aspect+datum_feature+shape_aspect",{"datum":"E"}),
 ("MBD_C",14636,"Position surfacic profile.3","⌓ .07 A B C .015 A","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_C",14795,"Position.9","⌖ Ø.04Ⓜ A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+geometric_tolerance_with_modifiers+position_tolerance",{"modifiers":"(.MAXIMUM_MATERIAL_REQUIREMENT.)"}),
 ("MBD_C",14830,"Position.10","⌖ Ø.045 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+position_tolerance",{}),
 ("MBD_C",16020,"Linear Size.12","2× Ø0.594 +.012 / -0.002","dimensional_size",{"size type":"diameter"}),
 ("MBD_C",16461,"Text.8","Text.8","property_definition",{"property":"semantic text"}),
 ("MBD_D",14866,"Position surfacic profile.5","⌓ .006 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_D",15050,"Position surfacic profile.6","⌓ .03 A B C .005 A B C","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_D",15098,"Position surfacic profile.7","⌓ .07 A B C .015 A","geometric_tolerance+geometric_tolerance_with_datum_reference+surface_profile_tolerance",{}),
 ("MBD_D",16200,"Linear Size.13","2× 0.012-0.022","dimensional_size",{"size type":"radius"}),
 ("MBD_D",16520,"Text.10","Text.10","property_definition",{"property":"semantic text"}),
]


def build_md() -> str:
    out = ["# PMI 提取结果", "", "- 视图分组：4 个", "- 标注总数：53 条", ""]
    cur = None
    for g, h, nm, title, ty, extra in ROWS:
        if g != cur:
            cur = g
            n = sum(1 for r in ROWS if r[0] == g)
            out += [f"## {g}", "", f"- 分组 Handle：`16{hash(g) % 900 + 100}`", f"- 标注数量：{n} 条", ""]
        seq = sum(1 for r in ROWS[:ROWS.index((g, h, nm, title, ty, extra))] if r[0] == g) + 1
        d = {"handle": str(h), "name": nm, "original_type_name": "draughting_callout", "type": ty}
        d.update(extra)
        out += [f"### {seq}. {title}", "", "detailData:", json.dumps(d, ensure_ascii=False, indent=2), ""]
    return "\n".join(out)


# ============================================================
# 断言工具
# ============================================================
PASS, FAIL = [], []


def check(name, got, want):
    if got == want:
        PASS.append(name)
    else:
        FAIL.append(f"{name}：期望 {want!r}，实际 {got!r}")


# ============================================================
# 一、纯函数单测
# ============================================================
# 1. 乱码修复（UTF-8 被按 GBK 解码）
check("乱码 鈱€ -> ⌀", C.fix_mojibake("鈱€.015"), "⌀.015")
check("乱码 鈱´ -> ⌴", C.fix_mojibake("鈱´"), "⌴")
check("正常中文不误伤", C.fix_mojibake("位置度"), "位置度")
check("正常中文不误伤2", C.fix_mojibake("机械零件"), "机械零件")
check("纯 ASCII 不变", C.fix_mojibake("Simple Datum.1"), "Simple Datum.1")

# 2. 数值归一（保留负号、忽略正号、消浮点噪声）
check("数值 +.003", C.norm_number("+.003"), "0.003")
check("数值 -.003", C.norm_number("-.003"), "-0.003")
check("数值 -.000", C.norm_number("-.000"), "-0")
check("数值 .02", C.norm_number(".02"), "0.02")
check("数值 浮点噪声", C.norm_number("0.2999999999"), "0.3")

# 3. 符号归一 + 分隔符保留
check("符号归一", C.normalize("Ø0.250 +.003 / -0"), "D0.25 0.003 / -0")
check("保留竖线分隔", C.normalize("A|B|C"), "A | B | C")
check("统计公差符号", C.normalize("飦"), "ST")
check("平面度符号", C.normalize("⏥ .01"), "▱ 0.01")

# 4. 语义指纹：同一公差的不同呈现方式应等价
t1, t2 = C.sem_tokens("⌖ Ø.04Ⓜ A B C"), C.sem_tokens("⌖ | ⌀0.040 Ⓜ | A | B | C")
check("指纹 数值", t1["numbers"], t2["numbers"])
check("指纹 基准", t1["letters"], t2["letters"])
check("指纹 修饰符", t1["modifiers"], t2["modifiers"])

# 5. 数量前缀（`4X`）不参与数值集合，单独比较
t3, t4 = C.sem_tokens("4× Ø0.173 +.005 / -0.001"), C.sem_tokens("4X ⌀0.173 +.005 -.001")
check("指纹 数量前缀剔除", t3["numbers"], t4["numbers"])
check("指纹 数量值", (t3["count"], t4["count"]), (4, 4))

# 6. 缺陷检测：乱码识别（含字节不全、无法程序化还原的残缺码）
check("乱码识别 鈱€", C.mojibake_chars("鈱€.015"), ["鈱", "€"])
check("乱码识别 鈱´", C.mojibake_chars("鈱´"), ["鈱", "´"])
check("乱码识别 飦（残缺映射表）", C.mojibake_chars("飦"), ["飦"])
check("乱码识别 正常中文不误伤", C.mojibake_chars("位置度"), [])
check("乱码识别 纯 ASCII", C.mojibake_chars("Simple Datum.1"), [])
check("乱码还原 鈱€ -> ⌀", C.repaired_text("鈱€.015"), "⌀.015")
check("乱码还原 飦 -> Ⓢ", C.repaired_text("飦"), "Ⓢ")

# 7. 数量前缀：SFA 图形文本里前缀不在串首（`DIM | 4X ⌀.250`），且常用 `×`
check("数量前缀 串首", C._CNT_PREFIX.search("4X ⌀0.173").group(1), "4")
check("数量前缀 竖线后", C._CNT_PREFIX.search("DIM | 4X ⌀.250").group(1), "4")
check("数量前缀 乘号写法", C._CNT_PREFIX.search("2× Ø0.438 ±.01").group(1), "2")
check("数量前缀 无", C._CNT_PREFIX.search("⌓ .06 A B C"), None)

# 8. 数值精度等价：SFA 图形通道对同一注释输出两行（全精度 + 展示精度）
check("精度等价 .4375/.438", C._num_equivalent("0.4375", "0.438"), True)
check("精度等价 .25/.250", C._num_equivalent("0.25", "0.250"), True)
check("精度等价 真差异", C._num_equivalent("0.015", "0.02"), False)
check("公差数值过滤 整数", C._tolerance_like("1"), False)
check("公差数值过滤 零", C._tolerance_like("0"), False)
check("公差数值过滤 真公差", C._tolerance_like("0.015"), True)

# ============================================================
# 二、markdown 解析
# ============================================================
from collections import Counter

md = build_md()
items = C.parse_dev_markdown(md)
check("解析条目数", len(items), 53)
cnt = Counter(i.kind for i in items)
check("几何公差条数", cnt[C.KIND_GT], 21)
check("尺寸公差条数", cnt[C.KIND_DIM], 17)
check("基准条数", cnt[C.KIND_DATUM], 5)
check("注释条数", cnt[C.KIND_NOTE], 10)
check("handle 解析", items[0].handle, 13706)
check("name 解析", items[0].name, "Simple Datum.1")
check("末条分组解析", items[-1].group, "MBD_D")

# ============================================================
# 三、集成测试（需真实 SFA 报告）
# ============================================================
if not os.path.exists(XLSX):
    print(f"[跳过集成测试] 未找到 SFA 报告：{XLSX}")
    print("   可执行： python test_pmi_core.py <你的SFA报告.xlsx>")
else:
    t = C.load_sfa(XLSX)
    check("SFA 工作表数", len(t.sheets), 113)
    check("SFA 语义项数", len(t.semantic), 49)
    check("SFA draughting_callout 行数", len(t.dc), 53)
    check("SFA 图形注释行数", len(t.ta), 53)
    check("SFA datum 行数", len(t.datum), 5)
    check("无解析警告", t.warnings, [])

    # 关联键：md.handle 必须等于 SFA draughting_callout.ID，且 name 一致
    check("全部 handle 命中 SFA", [i.handle for i in items if i.handle not in t.dc], [])
    check("handle/name 与 SFA 一致",
          [i.name for i in items if i.handle in t.dc and t.dc[i.handle] != i.name], [])

    rows = C.match_items(t, items)
    m = C.compute_metrics(rows)

    # 结果表与汇总表必须带缺陷字段
    check("结果表含缺陷列", [c for c in rows[0].as_dict() if "缺陷" in c], ["提取缺陷", "缺陷详情"])
    s = C.summarize(rows, t, items)
    check("汇总含缺陷字段",
          [k for k in s if k.startswith("缺陷_")],
          ["缺陷_条目数", "缺陷_总数", "缺陷_占比", "缺陷_分布"])

    # 三条关联路径都必须出现，且无漏网
    paths = Counter(r.path for r in rows)
    check("1 跳直连（几何公差）", paths["gt-1hop"], 27)
    check("2 跳（尺寸 dcr）", paths["dim-2hop"], 17)
    check("2 跳（基准 datum）", paths["datum-2hop"], 5)
    check("注释条目", paths["note"], 10)
    check("基准体系组合项", paths["datum_system"], 5)
    check("无未关联条目", [r.key for r in rows if r.path == "none"], [])

    # 三层指标
    check("语义 应提取", m.sem_expected, 44)
    check("语义 缺失", m.sem_miss, 0)
    check("语义 多余", m.sem_extra, 0)
    check("语义 疑似", m.sem_suspect, 0)
    check("语义 召回率", round(m.recall, 2), 100.00)
    check("语义 精确率", round(m.precision, 2), 100.00)
    check("基准 应提取", m.datum_expected, 5)
    check("基准 命中", m.datum_hit, 5)
    check("注释 总数", m.note_total, 10)
    check("注释 图形通道有文本", m.note_graphic_only, 3)
    check("注释 内部独占", m.note_exclusive, 7)

    # 人工校验不污染自动指标
    m2 = C.compute_metrics(rows, {r.key: "待定" for r in rows})
    check("初始/待定指标一致", (m.sem_hit, round(m.recall, 2)), (m2.sem_hit, round(m2.recall, 2)))

    # 已知差异应被正确识别为「文本差异」而非「疑似」
    diff_kinds = {r.key: r.status for r in rows if r.status == C.ST_HIT_DIFF}
    check("复合公差合并判定",
          diff_kinds.get("MBD_B#2@14121"), C.ST_HIT_DIFF)
    check("数量前缀差异判定",
          diff_kinds.get("MBD_D#4@16189"), C.ST_HIT_DIFF)

    # ------------------------------------------------------------
    # 四、开发侧提取缺陷层
    # 缺陷=开发侧数据本身错（乱码/丢符号/丢前缀），必须报出且不得影响判定
    # ------------------------------------------------------------
    defects = {}
    for r in rows:
        if r.defects:
            defects.setdefault(r.key.split("@")[0], (r.defects, r.defect_detail))

    check("缺陷 条目数与指标一致", m.defect_rows, len(defects))
    check("缺陷 ENC 编码损坏条目",
          sorted(k for k, (c, _) in defects.items() if C.DF_ENC in c), ["MBD_B#20", "MBD_B#21"])
    check("缺陷 ENC 明细给出还原值", "Ⓢ" in defects["MBD_B#20"][1], True)
    check("缺陷 CNT 数量前缀不符条目",
          sorted(k for k, (c, _) in defects.items() if C.DF_CNT in c),
          ["MBD_A#16", "MBD_A#17", "MBD_B#17", "MBD_D#4"])
    check("缺陷 SYM 符号丢失 (Parallelism.1 缺 ∥)",
          C.DF_SYM in defects.get("MBD_B#4", ([], ""))[0], True)
    check("缺陷 SYM 符号丢失 (Flatness.1 缺 ⌀ 定义区域)",
          C.DF_SYM in defects.get("MBD_B#3", ([], ""))[0], True)
    check("缺陷 NUM 无误报", C.DF_NUM in m.defect_by_code, False)
    check("缺陷 图形通道双行精度差不计入缺陷", "MBD_A#15" in defects, False)
    check("缺陷 无 MAP / DUP 误报",
          [k for k, (c, _) in defects.items() if C.DF_MAP in c or C.DF_DUP in c], [])
    check("缺陷 分布", m.defect_by_code, {C.DF_CNT: 4, C.DF_SYM: 3, C.DF_ENC: 2})
    check("缺陷 总数", m.defect_total, 9)
    # 缺陷与分析指标解耦：有缺陷的条目只要内容对得上，仍计入命中
    check("缺陷不拉低指标（缺陷条目仍全命中）",
          [r.key for r in rows if r.defects and r.layer == C.LAYER_SEM
           and r.status not in (C.ST_HIT, C.ST_HIT_DIFF)], [])
    check("缺陷不进 precision 分母", m.sem_extracted, 44)

    print("\n--- 开发侧提取缺陷（本工具要产出的核心结论）---")
    for k, (codes, detail) in sorted(defects.items()):
        print(f"  {k:10s} {' '.join(codes):34s} {detail}")
    print("  缺陷分布:", m.defect_by_code)
    print("  缺陷占比: %.2f%%（%d 条开发标注含缺陷 / 共 %d 条参与比对）"
          % (m.defect_rate, m.defect_rows,
             m.sem_extracted + m.datum_expected + m.note_total))

# ============================================================
# 三、解析自检（ParseDiag）—— 覆盖「一条都没对上」的各类畸形输入
# ============================================================
def _md(detail_block: str, heading: str = "### 1. 标题A") -> str:
    return f"## MBD_A\n\n- 标注数量：1 条\n\n{heading}\n\n{detail_block}\n"


_GOOD = '{\n  "handle": "13706",\n  "name": "Simple Datum.1",\n  "type": "datum_feature"\n}'

# 正常
it, dg = C.parse_dev_markdown_ex(_md("detailData:\n" + _GOOD))
check("自检 正常 条目数", dg.items, 1)
check("自检 正常 带handle", dg.with_handle, 1)
check("自检 正常 带name", dg.with_name, 1)
check("自检 正常 healthy", dg.healthy, True)
check("自检 正常 无问题", dg.problems(), [])

# 标签与 { 同行（旧实现会静默失败）
it, dg = C.parse_dev_markdown_ex(_md("detailData: " + _GOOD))
check("自检 detailData 同行 仍能解析", dg.with_handle, 1)
check("自检 detailData 同行 healthy", dg.healthy, True)

# 带 ```json 围栏
it, dg = C.parse_dev_markdown_ex(_md("detailData:\n```json\n" + _GOOD + "\n```"))
check("自检 围栏 仍能解析", dg.with_handle, 1)

# 全角冒号 + 引号标签
it, dg = C.parse_dev_markdown_ex(_md('"detailData"：\n' + _GOOD))
check("自检 全角冒号+引号 仍能解析", dg.with_handle, 1)

# 缺 detailData 块 -> 必然全判「多余」
it, dg = C.parse_dev_markdown_ex(_md(""))
check("自检 缺块 条目数", dg.items, 1)
check("自检 缺块 记为未解析", len(dg.detail_missing), 1)
check("自检 缺块 未取到 handle", dg.with_handle, 0)
check("自检 缺块 healthy", dg.healthy, False)
check("自检 缺块 给出提示", any("detailData" in p for p in dg.problems()), True)

# broken JSON
it, dg = C.parse_dev_markdown_ex(_md('detailData:\n{\n  "handle": "13706",\n')) 
check("自检 JSON 截断 记为失败", len(dg.json_failed), 1)
check("自检 JSON 截断 未取到 handle", dg.with_handle, 0)

# 字段名大小写
it, dg = C.parse_dev_markdown_ex(
    _md('detailData:\n{\n  "Handle": "13706",\n  "Name": "Simple Datum.1"\n}'))
check("自检 大写字段名 容错", (dg.with_handle, dg.with_name), (1, 1))

# JSON 在但缺 handle/name
it, dg = C.parse_dev_markdown_ex(_md('detailData:\n{\n  "type": "datum_feature"\n}'))
check("自检 缺关键字段 记为 key_missing", len(dg.key_missing), 1)

# 标题没序号
it, dg = C.parse_dev_markdown_ex(_md("detailData:\n" + _GOOD, heading="### 标题A"))
check("自检 标题无序号 条目数", dg.items, 0)
check("自检 标题无序号 记录未匹配行", len(dg.headings_unmatched), 1)
check("自检 标题无序号 给出提示",
      any("没有解析出任何标注条目" in p for p in dg.problems()), True)

# 四级标题（解析不出条目）
it, dg = C.parse_dev_markdown_ex(_md("detailData:\n" + _GOOD, heading="#### 1. 标题A"))
check("自检 四级标题 条目数", dg.items, 0)
check("自检 四级标题 记录未匹配行", len(dg.headings_unmatched), 1)

# 兼容入口仍返回纯条目
check("parse_dev_markdown 仍返回条目", len(C.parse_dev_markdown(_md("detailData:\n" + _GOOD))), 1)


# ============================================================
# 四、SFA 列数容错 / datum_target / 分段取文本 / 修饰符字形
#     覆盖「一条都没比对上」的根因：
#       - tessellated_annotation_occurrence 缺 col14 时整表被跳过 -> 关联全断
#       - datum_target 未走 1 跳引用 -> 基准目标全判多余
#       - 语义行把尺寸行与 FCF 行揉在一起 -> 内容比对大面积误报
# ============================================================
def _build_sfa(tmpdir: str) -> str:
    """构造一份最小 SFA 报告：ta 表刻意只给 12 列（缺 Equivalent Unicode String）。"""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Header")
    for r in ([f"Header row {i}", None, None] for i in range(3)):
        ws.append(r)

    # 语义表：注意 2002 号把「尺寸行 + FCF 行」揉在一行
    ws = wb.create_sheet("Semantic PMI Summary")
    ws.append(["Semantic PMI Summary (5)", None, None, None])
    ws.append(["ID", "Entity", "Expected PMI", "Similar"])
    for rid, ent, txt in [
        (1001, "datum_system", "A | B"),
        (2001, "position_tolerance", "⌖ | ⌀0.8 | A | B"),
        (2002, "cylindricity_tolerance", "2X ⌀3.50 ± 0.2\n⌭ | 0.1"),
        (2003, "datum_target", "K1 (area)"),
    ]:
        ws.append([str(rid), ent, txt, None])

    ws = wb.create_sheet("draughting_callout")
    ws.append(["draughting_callout  (4)", None])
    for rid, nm in [(1001, "Simple Datum.1"), (1101, "Position.1"),
                    (1201, "Cylindricity.1"), (1301, "Datum Target.1")]:
        ws.append([str(rid), nm])

    # ta 表只写 12 列，第 12 列（Associated Semantic PMI）在索引 10
    ws = wb.create_sheet("tessellated_annotation_occurren")
    ws.append(["tessellated_annotation_occurrence (4)"] + [None] * 11)
    ws.append(["ID", "name", "styles", "item", "name", "children",
               "presentation style", "color", "plane", "Associated Geometry",
               "Associated Semantic PMI", "Saved Views"])
    for rid, nm, ref in [(1100, "Position.1", "position_tolerance 2001"),
                         (1200, "Cylindricity.1", "cylindricity_tolerance 2002"),
                         (1000, "Simple Datum.1", "datum_feature 1002"),
                         (1300, "Datum Target.1", "datum_target 2003")]:
        ws.append([str(rid), nm] + [""] * 8 + [ref, ""])

    ws = wb.create_sheet("datum")
    ws.append(["datum  (1)", None, None, None, None, None])
    # 只放 A：`Simple Datum.1` 的 ta 引用是 1002，靠 ID-1 两跳命中 1001
    for rid, ident in [(1001, "A")]:
        ws.append([str(rid), "datum", "x", "x", "x", ident])

    path = os.path.join(tmpdir, "mini_sfa.xlsx")
    wb.save(path)
    return path


_MD_MINI = """# PMI 提取结果

## MBD_X

- 标注数量：4 条

### 1. A

detailData:
{
  "handle": "1001",
  "datum": "A",
  "name": "Simple Datum.1",
  "type": "datum_feature"
}

### 2. ⌖ Ø.8 A B

detailData:
{
  "handle": "1101",
  "name": "Position.1",
  "type": "position_tolerance"
}

### 3. .1

detailData:
{
  "handle": "1201",
  "name": "Cylindricity.1",
  "type": "cylindricity_tolerance"
}

### 4. K1

detailData:
{
  "handle": "1301",
  "datum": "K",
  "target id": "1",
  "name": "Datum Target.1",
  "type": "datum_target"
}
"""

import tempfile  # noqa: E402

with tempfile.TemporaryDirectory() as _td:
    _truth = C.load_sfa(_build_sfa(_td))

check("列数容错 ta 表 12 列仍装载", len(_truth.ta), 4)
check("列数容错 ta 列数已记录", _truth.ta_cols, 12)
check("列数容错 缺图形文本通道时给出告警", bool(_truth.warnings), True)
check("列数容错 告警指向 Unicode 列",
      any("Equivalent Unicode String" in w for w in _truth.warnings), True)

_items_mini, _ = C.parse_dev_markdown_ex(_MD_MINI)
_rows_mini = C.match_items(_truth, _items_mini)
_m_mini = C.compute_metrics(_rows_mini)
check("列数容错 全部关联成功（无缺失）", _m_mini.sem_miss + _m_mini.sem_extra, 0)
check("列数容错 召回 100", _m_mini.recall, 100.0)
check("列数容错 基准 100", _m_mini.datum_coverage, 100.0)
check("datum_target 走 1 跳引用",
      C._lookup_semantic_for_dev(_items_mini[3], _truth)[1], "datum_target-1hop")
check("datum_target 判定为命中",
      [r.status for r in _rows_mini if r.key.startswith("MBD_X#4")], [C.ST_HIT])
check("普通基准走 ID-1 两跳",
      C._lookup_semantic_for_dev(_items_mini[0], _truth)[1], "datum-2hop")

# GT 分段：语义行含「尺寸行 + FCF 行」时只取 FCF 行，且保留原文字形
check("分段 GT 只取 FCF 行",
      C.sfa_text_for("2X ⌀3.50 ± 0.2\n⌭ | 0.1", C.KIND_GT), "⌭ | 0.1")
check("分段 GT 保留原文字形",
      C.sfa_text_for("⌖ | ⌀0.8 | A | B", C.KIND_GT), "⌖ | ⌀0.8 | A | B")
check("分段 无 FCF 行时退回首行",
      C.sfa_text_for("⌀3.50 G6", C.KIND_GT), "⌀3.50 G6")
check("分段 DIM 合并折行",
      C.sfa_text_for("⌀0.250 +.003\n-.000", C.KIND_DIM), "⌀0.250 +.003 -.000")

# 数量前缀不得污染字母集（`2X` 的 X 曾被当成基准字母）
_tk_cnt = C.sem_tokens("2X ⌀5.50 ± 0.2")
check("数量前缀 已识别", _tk_cnt["count"], 2)
check("数量前缀 不泄漏进基准字母", "X" in _tk_cnt["letters"], False)

# 修饰符优先按字形识别（开发 `EⓂ-FⓂ-GⓂ` vs SFA `E Ⓜ-F Ⓜ-G Ⓜ`）
_tk_dev = C.sem_tokens("⌖ Ø.8 D EⓂ-FⓂ-GⓂ")
_tk_sfa = C.sem_tokens("⌖ | ⌀0.8 | D | E Ⓜ-F Ⓜ-G Ⓜ")
check("修饰符字形 开发侧识别 M", _tk_dev["modifiers"], ["M"])
check("修饰符字形 SFA 侧识别 M", _tk_sfa["modifiers"], ["M"])
check("修饰符字形 基准字母含 F", _tk_dev["letters"], ["D", "E", "F", "G"])
check("修饰符字形 两侧判定一致",
      C.compare_tokens(_tk_dev, _tk_sfa)[0], C.ST_HIT)

# 基准字母与修饰符同形（`L` 既是基准又是 LMC）需对称消歧
check("同形字母 对称剔除后判命中",
      C.compare_tokens(C.sem_tokens("⌖ Ø0 H K L"),
                       C.sem_tokens("⌖ | ⌀0 Ⓛ | H | K | L"))[0], C.ST_HIT)

# 符号集：直线度用 U+2212、圆柱度 ⌭ 要检出；`⭩◎` 属渲染杂符不得误报
check("符号 直线度符已收录", C.symbols_of("− | 0.2 / 15"), {"−"})
check("符号 圆柱度符已收录", C.symbols_of("⌭ | 0.1"), {"⌭"})
check("符号 杂符 ⭩◎ 不计入", C.symbols_of("⭩◎ | ⌓ | 1.5"), {"⌓"})

# SFA 把统计公差渲染成字体私有区字形 U+F055（同一份报告里另一种写法是 `<ST>`），
# 认不出就等于丢掉修饰符
check("符号 私有区 ST 字形归一化为 ST", C.normalize("\uf055"), "ST")
check("符号 指向符号 ↧ 被剔除", "↧" in C.normalize("DIM | ↧.30±.02"), False)
check("符号 注释杂符 ⌴ 被剔除", C.normalize("⌴"), "")
# ▽ / ⎹ 出现在带基准框架里、与 [A] 同行，属 SFA 版面元素；
# 混进公差符号集会让「缺符号」检查误报一片
check("符号 排版字形不混入公差符号集",
      C.symbols_of("▽ ⎹ ⌓ | 1.2 | A"), {"⌓"})
check("符号 排版字形已登记", {"▽", "⎹"} <= C.SFA_LAYOUT_GLYPHS, True)


# ============================================================
# 五、列名驱动装载 / 加载期交叉校验 / 断链候选
#     覆盖本次事故的同源根因：装载器只认列号，列数或列序一变就静默失配
#     （曾整张 tessellated 表被跳过 -> 结果全判「多余 / 缺失」）。
#     表头列名其实就在表里（带换行与 (Sec. x) 后缀），以前从来没读。
# ============================================================
def _build_sfa_shuffled(tmpdir: str) -> str:
    """列序全部打乱 + 表头带换行与 (Sec. x) 后缀的报告。

    若装载器仍按列号取，ref 会落到空列、关联全断；按列名定位则应完全正常。
    """
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Semantic PMI Summary")
    ws.append(["Semantic PMI Summary (3)", None, None, None])
    # 表头带换行与章节号；Expected PMI 被挪到第 4 列
    ws.append(["Entity\n(Sec. 2.1)", "ID", "Similar PMI / Exception", "Semantic PMI"])
    for rid, ent, txt in [(2001, "position_tolerance", "⌖ | ⌀0.8 | A | B"),
                          (3001, "flatness_tolerance", "▱ | 0.2"),
                          (4001, "dimensional_size", "⌀6.00 ± 0.15")]:
        ws.append([ent, str(rid), None, txt])

    ws = wb.create_sheet("draughting_callout")
    ws.append(["draughting_callout  (3)", None])
    ws.append(["name", "ID"])                       # 列序反转
    for rid, nm in [(1101, "Position.1"), (1201, "Flatness.1"), (1301, "Linear Size.1")]:
        ws.append([nm, str(rid)])

    ws = wb.create_sheet("tessellated_annotation_occurren")
    ws.append(["tessellated_annotation_occurrence  (3)"] + [None] * 4)
    ws.append(["列1", "列2", "列3", "列4", "列5"])    # Excel 占位行，须跳过
    ws.append(["ID", "name", "Associated Semantic PMI (Sec. 7.3)", "styles",
               "Equivalent Unicode String(s)\n(Sec. 10.1.3.3)"])
    for rid, nm, ref, uni in [
        (1100, "Position.1", "position_tolerance 2001", "⌖ | ⌀0.8 | A | B"),
        (1200, "Flatness.1", "flatness_tolerance 3001", "▱ | 0.2"),
        (1300, "Linear Size.1", "dimensional_size 5001", "⌀6.00 ± 0.15"),
    ]:
        ws.append([str(rid), nm, ref, "", uni])

    ws = wb.create_sheet("dimensional_characteristic_repr")
    ws.append(["dimensional_characteristic_representation  (1)"] + [None] * 2)
    ws.append(["ID", "dimension (Sec. 5.1.1)", "Associated Geometry"])
    ws.append(["4001", "dimensional_size 5001", "x"])

    path = os.path.join(tmpdir, "shuffled_sfa.xlsx")
    wb.save(path)
    return path


def _build_sfa_noheader(tmpdir: str) -> str:
    """ta 表没有可用表头（只有 Excel 占位行），且 ref 列不在默认位置。

    期望：退回默认列号、标记为不可信、交叉校验报出来 —— 而不是静默失配。
    """
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Semantic PMI Summary")
    ws.append(["Semantic PMI Summary (1)", None, None, None])
    ws.append(["ID", "Entity", "Semantic PMI", "Similar"])
    ws.append(["2001", "position_tolerance", "⌖ | ⌀0.8 | A | B", None])

    ws = wb.create_sheet("draughting_callout")
    ws.append(["draughting_callout  (1)", None])
    ws.append(["ID", "name"])
    ws.append(["1101", "Position.1"])

    ws = wb.create_sheet("tessellated_annotation_occurren")
    ws.append(["tessellated_annotation_occurrence  (1)"] + [None] * 4)
    ws.append(["列1", "列2", "列3", "列4", "列5"])     # 唯一「表头」就是占位行
    ws.append(["1100", "Position.1", "", "position_tolerance 2001", ""])

    path = os.path.join(tmpdir, "noheader_sfa.xlsx")
    wb.save(path)
    return path


def _build_sfa_nonumeric(tmpdir: str) -> str:
    """模拟本次事故形态：ta 表 ID 列读不出数字，整表装载为 0。

    期望：交叉校验直接报「无可识别的数据行」，而不是等到结果表全红才发现。
    """
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Semantic PMI Summary")
    ws.append(["Semantic PMI Summary (1)", None, None, None])
    ws.append(["ID", "Entity", "Semantic PMI", "Similar"])
    ws.append(["2001", "position_tolerance", "⌖ | ⌀0.8 | A | B", None])

    ws = wb.create_sheet("draughting_callout")
    ws.append(["draughting_callout  (1)", None])
    ws.append(["ID", "name"])
    ws.append(["1101", "Position.1"])

    ws = wb.create_sheet("tessellated_annotation_occurren")
    ws.append(["tessellated_annotation_occurrence  (1)"] + [None] * 11)
    ws.append(["ID", "name", "styles", "item", "name", "children", "presentation style",
               "color", "plane", "Associated Geometry", "Associated Semantic PMI",
               "Saved Views"])
    # ID 带 `#` 前缀 -> isdigit() 为假 -> 整表装载 0
    ws.append(["#1100", "Position.1", "", "", "", "", "", "", "", "",
               "position_tolerance 2001", ""])

    path = os.path.join(tmpdir, "nonumeric_sfa.xlsx")
    wb.save(path)
    return path


_MD_SHUFFLED = """# PMI 提取结果

## MBD_Y

- 标注数量：3 条

### 1. ⌖ Ø.8 A B

detailData:
{
  "handle": "1101",
  "name": "Position.1",
  "type": "position_tolerance"
}

### 2. ⏥ .2

detailData:
{
  "handle": "1201",
  "name": "Flatness.1",
  "type": "flatness_tolerance"
}

### 3. Ø6.00 ±.15

detailData:
{
  "handle": "1301",
  "name": "Linear Size.1",
  "type": "dimensional_size"
}
"""

with tempfile.TemporaryDirectory() as _td_sh:
    _sh = C.load_sfa(_build_sfa_shuffled(_td_sh))
_sh_items, _ = C.parse_dev_markdown_ex(_MD_SHUFFLED)
_sh_rows = C.match_items(_sh, _sh_items)
_sh_m = C.compute_metrics(_sh_rows)

check("列序打乱 语义表仍装载 3 项", len(_sh.semantic), 3)
check("列序打乱 dc 表 ID/name 反转仍装载", len(_sh.dc), 3)
check("列序打乱 ta 表 ref 列挪位仍装载", len(_sh.ta), 3)
check("列序打乱 dcr 表仍装载", len(_sh.dcr_by_dim), 1)
check("列序打乱 语义表定位到 Semantic PMI 列",
      _sh.recipe("semantic").resolved["semantic pmi"], 3)
check("列序打乱 dc 定位到 name=0 / id=1",
      (_sh.recipe("dc").resolved["name"], _sh.recipe("dc").resolved["id"]), (0, 1))
check("列序打乱 ta 定位到 ref 列 2",
      _sh.recipe("ta").resolved["associated semantic pmi"], 2)
check("列序打乱 ta 识别出 Unicode 列",
      _sh.recipe("ta").resolved["equivalent unicode string"], 4)
check("列序打乱 每张表都按表头识别（无 fallback）",
      sorted({r.source for r in _sh.recipes}), ["header"])
check("列序打乱 交叉校验全通过", [b.line() for b in _sh.broken_checks()], [])
check("列序打乱 召回 100", _sh_m.recall, 100.0)
check("列序打乱 精确 100", _sh_m.precision, 100.0)
check("列序打乱 无断链条目", C.link_stats(_sh_rows).get("none", 0), 0)

with tempfile.TemporaryDirectory() as _td_nh:
    _nh = C.load_sfa(_build_sfa_noheader(_td_nh))
check("无表头 ta 定位不可信", _nh.recipe("ta").trustworthy, False)
check("无表头 ta 仍装载（不静默跳过整表）", len(_nh.ta), 1)
check("无表头 ta ref 取空（关联会失效）", _nh.ta["Position.1"]["sem_refs"], [])
check("无表头 交叉校验报出列名未识别",
      any("列名识别" in b.name for b in _nh.broken_checks()), True)
check("无表头 语义表不受影响、仍可靠",
      _nh.recipe("semantic").trustworthy, True)

with tempfile.TemporaryDirectory() as _td_ne:
    _ne = C.load_sfa(_build_sfa_nonumeric(_td_ne))
check("ID 非数字 ta 装载 0 条", len(_ne.ta), 0)
check("ID 非数字 交叉校验报出无可识别数据行",
      any("无可识别的数据行" in b.value for b in _ne.broken_checks()), True)

# 断链候选：只提示、不参与判定
_MD_TYPO = _MD_SHUFFLED.replace('"name": "Position.1"', '"name": "Position.11"')
_tp_items, _ = C.parse_dev_markdown_ex(_MD_TYPO)
_tp_rows = C.match_items(_sh, _tp_items)
_tp_m = C.compute_metrics(_tp_rows)
_sus = C.suspect_links(_tp_items, _sh, _tp_rows)

check("断链候选 关联确已断裂", C.link_stats(_tp_rows).get("none", 0), 1)
check("断链候选 给出候选", len(_sus), 1)
check("断链候选 指向正确条目",
      [c["ta_name"] for c in _sus[0].candidates] if _sus else [], ["Position.1"])
check("断链候选 判定未被改写（条目仍判多余）", _tp_m.sem_extra >= 1, True)
check("断链候选 判定未被改写（对端仍判缺失）", _tp_m.sem_miss >= 1, True)
check("断链候选 正常数据下不产生噪音",
      C.suspect_links(_sh_items, _sh, _sh_rows), [])


# ============================================================
# 汇总
# ============================================================
for f in FAIL:
    print("FAIL -", f)
print(f"\n通过 {len(PASS)} / {len(PASS) + len(FAIL)}")
sys.exit(1 if FAIL else 0)

