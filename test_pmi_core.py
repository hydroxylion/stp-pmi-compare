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
# 汇总
# ============================================================
for f in FAIL:
    print("FAIL -", f)
print(f"\n通过 {len(PASS)} / {len(PASS) + len(FAIL)}")
sys.exit(1 if FAIL else 0)

