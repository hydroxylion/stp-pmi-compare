# -*- coding: utf-8 -*-
"""
界面冒烟测试（无头）
====================

用 Streamlit 官方 `AppTest` 直接执行 `pmi_compare_work.py`，验证：

1. 首屏能渲染（上传区 + 视图选项）
2. 注入比对结果后，表格走「精简视图」且无异常
3. 侧边栏「显示字段」可勾选恢复隐藏列：全选走完整视图、全不选回落默认视图、
   单选 `SFA图形文本` 也能正常渲染
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
                r"\nist_ftc_07_asme1_ap242-e2-sfa.xlsx")
XLSX = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX

from streamlit.testing.v1 import AppTest  # noqa: E402

import pmi_core as core  # noqa: E402


def load_consts(path, names):
    """从界面脚本里静态读取模块级常量。

    直接 import 界面脚本会把它整个执行一遍（bare mode），所以走 AST。
    支持 `A = B + C` 这类引用其它常量的写法：按声明顺序逐个求值，
    已解析出的常量作为前序命名空间传入。
    """
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    out = {}
    for node in tree.body:
        if (isinstance(node, ast.Assign) and len(node.targets) == 1
                and isinstance(node.targets[0], ast.Name)
                and node.targets[0].id in names):
            try:
                out[node.targets[0].id] = ast.literal_eval(node.value)
            except (ValueError, TypeError):
                expr = ast.Expression(body=node.value)
                ast.fix_missing_locations(expr)
                out[node.targets[0].id] = eval(          # noqa: S307
                    compile(expr, path, "eval"), {}, dict(out))
    missing = set(names) - set(out)
    if missing:
        raise AssertionError(f"界面脚本缺少常量：{sorted(missing)}")
    return out


UI_PATH = os.path.join(PROJ, "pmi_compare_work.py")
_SPEC = load_consts(UI_PATH, {"VIEW_COLS", "DIAG_COLS", "ALL_COLS"})
VIEW_COLS, DIAG_COLS = _SPEC["VIEW_COLS"], _SPEC["DIAG_COLS"]
ALL_COLS = _SPEC["ALL_COLS"]

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
TOTAL = 0


def check(name, cond, got=""):
    global TOTAL
    TOTAL += 1
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


def _all_text(at) -> str:
    """把面板里所有可见文本（caption / markdown / 表格 / 告警）拼起来，便于整体断言。"""
    parts = [str(c.value) for c in at.caption]
    parts += [str(m.value) for m in at.markdown]
    parts += [str(x.value) for x in at.warning]
    parts += [str(x.value) for x in at.error]
    for d in at.dataframe:
        try:
            parts.append(d.value.to_csv(index=False))
        except Exception:               # noqa: BLE001
            parts.append(str(d.value))
    return "\n".join(parts)


def main():
    if not os.path.exists(XLSX):
        print(f"跳过：真值文件不存在 {XLSX}")
        return 0

    truth = core.load_sfa(XLSX)
    items = core.parse_dev_markdown(build_md())
    rows = core.match_items(truth, items)
    meta = core.summarize(rows, truth, items)
    print(f"数据：开发条目 {len(items)} 条 → 比对行 {len(rows)} 行")
    print(f"精简视图列定义：{len(VIEW_COLS)} 列 / 默认隐藏：{len(DIAG_COLS)} 列")

    # 列定义本身的口径断言（防止以后误把诊断字段塞回默认视图，或把分组列藏起来）
    must_show = {"匹配状态", "分组", "开发名称", "开发标注",
                 "SFA语义文本", "提取缺陷", "人工校验", "备注"}
    must_hide = {"SFA图形文本", "Handle", "SFA语义ID", "SFA实体类型", "关联路径"}
    check("默认视图含核心列（含分组）", must_show <= set(VIEW_COLS),
          must_show - set(VIEW_COLS))
    check("分组列默认显示", "分组" in VIEW_COLS, VIEW_COLS)
    check("默认视图不含图形文本与检索字段", not (must_hide & set(VIEW_COLS)),
          must_hide & set(VIEW_COLS))
    check("SFA图形文本默认隐藏", "SFA图形文本" in DIAG_COLS, DIAG_COLS)
    check("视图与隐藏列无重叠", not (set(VIEW_COLS) & set(DIAG_COLS)),
          set(VIEW_COLS) & set(DIAG_COLS))
    check("全字段 = 可见 + 隐藏且无重复", len(ALL_COLS) == len(set(ALL_COLS)),
          ALL_COLS)
    all_cols = set(core.MatchRow().as_dict()) | {"人工校验"}
    check("列定义覆盖 as_dict 全部字段",
          set(ALL_COLS) == all_cols,
          all_cols - set(ALL_COLS))
    check("可勾选字段含全部默认隐藏项",
          set(DIAG_COLS) <= set(ALL_COLS), set(DIAG_COLS) - set(ALL_COLS))

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
    check("侧边栏有「显示字段」多选框", len(ms) == 1, [m.label for m in ms])
    if ms:
        check("显示字段默认值 = 核心列", list(ms[0].value) == list(VIEW_COLS), ms[0].value)
        check("可勾选字段 = 全部字段", list(ms[0].options) == list(ALL_COLS), ms[0].options)
        check("默认状态有隐藏列提示", "已隐藏" in caption_of(at, "已隐藏"), "")

        at.sidebar.multiselect[0].set_value(list(ALL_COLS))
        at.run()
        total = len(ALL_COLS)
        check("完整视图渲染无异常", not at.exception, [e.value for e in at.exception])
        cap = caption_of(at, "行 ×")
        check(f"完整视图列数 = {total}", f"× {total} 列" in cap, cap)
        check("完整视图不再提示隐藏", "已隐藏" not in cap, cap)

        # 全部取消也要能渲染：回落默认视图，不炸
        at.sidebar.multiselect[0].set_value([])
        at.run()
        check("取消全部勾选不炸", not at.exception, [e.value for e in at.exception])
        cap = caption_of(at, "行 ×")
        check("取消全部回落默认视图", f"× {len(VIEW_COLS)} 列" in cap, cap)

        # 只恢复「SFA图形文本」单列叠加核心列
        at.sidebar.multiselect[0].set_value(list(VIEW_COLS) + ["SFA图形文本"])
        at.run()
        cap = caption_of(at, "行 ×")
        check("单独恢复 SFA图形文本",
              f"× {len(VIEW_COLS) + 1} 列" in cap and "SFA图形文本" not in cap, cap)

    at.selectbox[0].set_value("语义PMI").run()
    check("筛选切换无异常", not at.exception, [e.value for e in at.exception])

    # ---------------- 解析自检面板 ----------------
    at.text_area(key="md_paste").set_value(build_md())
    at.run()
    check("自检面板 正常输入 渲染无异常", not at.exception, [e.value for e in at.exception])
    labels = [m.label for m in at.metric]
    for need in ("解析条目", "带 handle", "带 name"):
        check(f"自检面板 有指标「{need}」", need in labels, labels)
    check("自检面板 正常输入 不报错", not at.error, [e.value for e in at.error])

    broken = ("## MBD_A\n\n### 1. Simple Datum.1\n\n"
              "detailData:\n{\n  \"type\": \"datum_feature\"\n}\n\n"
              "### 2. Position.1\n\n这里忘了写 detailData\n")
    at.text_area(key="md_paste").set_value(broken)
    at.run()
    check("自检面板 畸形输入 渲染无异常", not at.exception, [e.value for e in at.exception])
    errs = " ".join(e.value for e in at.error)
    check("自检面板 指出 detailData 缺失", "detailData" in errs, errs[:120])
    check("自检面板 指出缺 handle", "handle" in errs, errs[:120])

    # ---------------- 报告体检面板 ----------------
    at.session_state["sfa_stat"] = (
        "语义表 12 项 · draughting_callout 13 条 · 图形标注 13 条"
        "（tessellated_annotation_occurrence） · datum 5 项 · dcr 6 项")
    at.session_state["sfa_warn"] = []
    at.session_state["sfa_checks"] = [
        {"name": "ta 列名识别", "ok": True, "value": "表头第 4 行", "detail": ""},
        {"name": "ta 装载", "ok": True, "value": "13/13 条", "detail": ""},
    ]
    at.session_state["sfa_recipes"] = [{
        "key": "ta", "sheet": "tessellated_annotation_occurren", "header_row": 3,
        "source": "header", "n_cols": 14, "n_rows": 13, "n_loaded": 13,
        "resolved": {"id": 0, "name": 1, "associated semantic pmi": 10,
                     "equivalent unicode string": 13},
        "notes": [],
    }, {
        # DMIA 形态：行数 > 装载条数（一条 callout 多行），须显示聚合组数
        "key": "ta", "sheet": "draughting_model_item_associati", "header_row": 2,
        "source": "header", "n_cols": 6, "n_rows": 46, "n_loaded": 20,
        "n_rows_grouped": 20,
        "resolved": {"definition": 3, "identified_item": 5},
        "notes": [],
    }]
    at.session_state["link_stats"] = {"gt-1hop": 3, "none": 0}
    at.session_state["suspected"] = []
    at.run()
    check("体检面板 渲染无异常", not at.exception, [e.value for e in at.exception])
    check("体检面板 全部正常时给出正常态",
          caption_of(at, "全部正常") != "", [c.value[:40] for c in at.caption][:8])
    check("体检面板 展示列定位（含列来源）",
          "表头识别" in _all_text(at), _all_text(at)[:140])
    check("体检面板 展示交叉校验项",
          "ta 列名识别" in _all_text(at), _all_text(at)[:140])
    check("体检面板 展示关联路径分布",
          "关联路径" in _all_text(at), _all_text(at)[:140])

    # 断链 + 疑似对应：面板应转为「发现问题」并列出候选
    at.session_state["link_stats"] = {"gt-1hop": 1, "none": 2}
    at.session_state["suspected"] = [{
        "开发条目": "MBD_X#1", "标题": "A", "name": "Simple Datum.1",
        "疑似对应": "Simple Datum.1（name 完全一致，关联本应成功）"}]
    at.run()
    check("体检面板 断链时渲染无异常", not at.exception, [e.value for e in at.exception])
    check("体检面板 断链时标为需要关注",
          caption_of(at, "需要关注") != "", [c.value[:40] for c in at.caption][:8])
    check("体检面板 疑似对应入表",
          "疑似对应" in _all_text(at), _all_text(at)[:140])
    check("体检面板 疑似对应不产生 error 污染结果",
          not any("Simple Datum" in e.value for e in at.error), [e.value[:40] for e in at.error])

    _w = "tessellated_annotation_occurrence 未提供 `Equivalent Unicode String(s)`"
    at.session_state["sfa_warn"] = [_w]
    at.run()
    check("体检面板 有告警时逐条展示",
          [_w] == [x.value for x in at.warning if x.value == _w], [x.value[:40] for x in at.warning])

    # ---------------- 无真值（图形专用导出） ----------------
    # 报告不含语义 PMI 时必须显式点破，否则满屏 0% 会被读成「开发侧提取全错」。
    at.session_state["truth_missing"] = True
    at.session_state["link_stats"] = {"no-truth": 2, "note": 1}
    at.session_state["suspected"] = []
    at.session_state["sfa_warn"] = ["该 SFA 报告不含任何语义 PMI：……"]
    at.session_state["sfa_checks"] = [
        {"name": "图形标注关联键列", "ok": False, "value": "缺 `Associated Semantic PMI` 列",
         "detail": "这份导出件只有图形 PMI、没有语义关联"},
        {"name": "语义真值", "ok": False, "value": "报告不含语义 PMI", "detail": ""},
    ]
    abs_rows = [
        core.MatchRow(key="MBD_A#1", layer=core.LAYER_NOTRUTH, kind=core.KIND_GT,
                      group="MBD_A", dev_title="⏥ .03", dev_name="Flatness.1",
                      status=core.ST_ABSENT, path="no-truth",
                      remark="该 SFA 报告不含语义 PMI …… 不代表开发侧提取错误。"),
        core.MatchRow(key="MBD_A#2", layer=core.LAYER_NOTRUTH, kind=core.KIND_DATUM,
                      group="MBD_A", dev_title="B", dev_name="Simple Datum.1",
                      status=core.ST_ABSENT, path="no-truth",
                      remark="该 SFA 报告不含语义 PMI …… 不代表开发侧提取错误。"),
    ]
    at.session_state["rows"] = abs_rows
    at.session_state["verdicts"] = {r.key: "待定" for r in abs_rows}
    at.session_state["meta"] = {"开发条目": 2, "SFA语义项": 0, "SFA表数": 80, "警告": []}
    at.run()
    check("无真值 渲染无异常", not at.exception, [e.value for e in at.exception])
    check("无真值 体检面板点破「不含语义 PMI」",
          "不含任何语义 PMI" in _all_text(at), _all_text(at)[:200])
    check("无真值 结果表点破「选错报告文件」",
          "选错报告文件" in _all_text(at) or "语义版" in _all_text(at),
          _all_text(at)[:200])
    check("无真值 三项指标显示 N/A",
          sum(1 for m in at.metric if m.value == "N/A") == 3,
          [m.value for m in at.metric])
    check("无真值 关联路径标注为不可比",
          "不可比" in _all_text(at), _all_text(at)[:200])
    check("无真值 点出不可比条数",
          any("2" in c.value and "不可比" in c.value for c in at.caption)
          or "不可比" in _all_text(at), [c.value[:40] for c in at.caption][:8])

    print()
    if FAIL:
        print(f"未通过 {len(FAIL)} / {TOTAL}")
        for f in FAIL:
            print("  -", f)
        return 1
    print(f"UI 冒烟全部通过（{TOTAL}/{TOTAL}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
