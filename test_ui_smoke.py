# -*- coding: utf-8 -*-
"""
界面冒烟测试（无头）
====================

用 Streamlit 官方 `AppTest` 直接执行 `pmi_compare_work.py`，验证：

1. 首屏能渲染（上传区 + 视图选项）
2. 注入比对结果后，表格走「精简视图」且无异常
3. 打开全部诊断字段后，表格走「完整视图」且无异常
4. 切换筛选条件后 key 重建正常，不残留异常

只验证「不炸 + 列数口径」，不做像素级断言。

用法::

    python test_ui_smoke.py [真值xlsx路径]
"""

import ast
import json
import os
import sys

PROJ = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJ)

DEFAULT_XLSX = (r"F:\1【机械零件】\Step官方标准数据\NIST-PMI-STEP-Files"
                r"\nist_ftc_07_asme1_ap242-e2-sfa-1.xlsx")
XLSX = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX

from streamlit.testing.v1 import AppTest  # noqa: E402

import pmi_core as core  # noqa: E402


def load_consts(path, names):
    """从界面脚本里静态读取模块级常量。

    直接 import 界面脚本会把它整个执行一遍（bare mode），所以走 AST。
    """
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    out = {}
    for node in tree.body:
        if (isinstance(node, ast.Assign) and len(node.targets) == 1
                and isinstance(node.targets[0], ast.Name)
                and node.targets[0].id in names):
            out[node.targets[0].id] = ast.literal_eval(node.value)
    missing = set(names) - set(out)
    if missing:
        raise AssertionError(f"界面脚本缺少常量：{sorted(missing)}")
    return out


UI_PATH = os.path.join(PROJ, "pmi_compare_work.py")
_SPEC = load_consts(UI_PATH, {"VIEW_COLS", "DIAG_COLS"})
VIEW_COLS, DIAG_COLS = _SPEC["VIEW_COLS"], _SPEC["DIAG_COLS"]

# 覆盖语义项 / 注释项 / 含缺陷项三类，够走通所有渲染分支
ROWS = [
    ("MBD_A", 13706, "Simple Datum.1", "A",
     "composite_shape_aspect+datum_feature", {"datum": "A"}),
    ("MBD_A", 13999, "Perpendicularity.1", "⟂ Ø.015Ⓟ.425 A",
     "perpendicularity_tolerance", {}),
    ("MBD_A", 15405, "Linear Size.4", "4× Ø0.250 +.003 / -0.001",
     "dimensional_size", {}),
    ("MBD_B", 16431, "Text.7", "飦",
     "property_definition", {"property": "semantic text"}),
    ("MBD_B", 16490, "Text.9", "鈱´",
     "property_definition", {"property": "semantic text"}),
]


def build_md() -> str:
    out = ["# PMI 提取结果", "", "- 标注总数：%d 条" % len(ROWS), ""]
    cur, seqs = None, {}
    for g, h, nm, title, ty, extra in ROWS:
        if g != cur:
            cur = g
            seqs[g] = 0
            out += [f"## {g}", "",
                    "- 标注数量：%d 条" % sum(1 for r in ROWS if r[0] == g), ""]
        seqs[g] += 1
        d = {"handle": str(h), "name": nm,
             "original_type_name": "draughting_callout", "type": ty}
        d.update(extra)
        out += [f"### {seqs[g]}. {title}", "",
                "detailData:", json.dumps(d, ensure_ascii=False, indent=2), ""]
    return "\n".join(out)


FAIL = []


def check(name, cond, got=""):
    if cond:
        print(f"  ok   {name}")
    else:
        FAIL.append(f"{name}（实际：{got}）")
        print(f"  FAIL {name}（实际：{got}）")


def caption_of(at, must_have):
    for c in at.caption:
        if must_have in c.value:
            return c.value
    return ""


def main():
    if not os.path.exists(XLSX):
        print(f"跳过：真值文件不存在 {XLSX}")
        return 0

    truth = core.load_sfa(XLSX)
    items = core.parse_dev_markdown(build_md())
    rows = core.match_items(truth, items)
    meta = core.summarize(rows, truth, items)
    print(f"数据：开发条目 {len(items)} 条 → 比对行 {len(rows)} 行")
    print(f"精简视图列定义：{len(VIEW_COLS)} 列 / 诊断列定义：{len(DIAG_COLS)} 列")

    # 列定义本身的口径断言（防止以后误把诊断字段塞回默认视图）
    must_show = {"匹配状态", "开发标注", "SFA语义文本", "提取缺陷", "人工校验", "备注"}
    must_hide = {"Handle", "SFA语义ID", "SFA实体类型", "关联路径"}
    check("默认视图含核心列", must_show <= set(VIEW_COLS),
          must_show - set(VIEW_COLS))
    check("检索字段默认隐藏", not (must_hide & set(VIEW_COLS)),
          must_hide & set(VIEW_COLS))
    check("视图与诊断列无重叠", not (set(VIEW_COLS) & set(DIAG_COLS)),
          set(VIEW_COLS) & set(DIAG_COLS))
    all_cols = set(core.MatchRow().as_dict()) | {"人工校验"}
    check("列定义覆盖 as_dict 全部字段",
          (set(VIEW_COLS) | set(DIAG_COLS)) == all_cols,
          all_cols - (set(VIEW_COLS) | set(DIAG_COLS)))

    at = AppTest.from_file(UI_PATH, default_timeout=60)
    at.run()
    check("首屏渲染无异常", not at.exception, [e.value for e in at.exception])

    at.session_state["rows"] = rows
    at.session_state["verdicts"] = {r.key: "待定" for r in rows}
    at.session_state["meta"] = meta
    at.run()
    check("精简视图渲染无异常", not at.exception, [e.value for e in at.exception])
    cap = caption_of(at, "行 ×")
    check(f"精简视图列数 = {len(VIEW_COLS)}", f"× {len(VIEW_COLS)} 列" in cap, cap)

    ms = [m for m in at.sidebar.multiselect]
    check("侧边栏有「附加字段」多选框", len(ms) == 1, [m.label for m in ms])
    if ms:
        ms[0].set_value(list(DIAG_COLS))
        at.run()
        total = len(VIEW_COLS) + len(DIAG_COLS)
        check("完整视图渲染无异常", not at.exception, [e.value for e in at.exception])
        cap = caption_of(at, "行 ×")
        check(f"完整视图列数 = {total}", f"× {total} 列" in cap, cap)

    at.selectbox[0].set_value("语义PMI").run()
    check("筛选切换无异常", not at.exception, [e.value for e in at.exception])

    print()
    if FAIL:
        print(f"未通过 {len(FAIL)} / {len(FAIL) + 12}")
        for f in FAIL:
            print("  -", f)
        return 1
    print("UI 冒烟全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
