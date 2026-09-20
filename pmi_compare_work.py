import streamlit as st
import pandas as pd
import json
import re
import unicodedata
from io import BytesIO
from difflib import SequenceMatcher

# ================= 直接运行兜底 =================
# 允许直接 `python pmi_compare_work.py` 启动：自动转交 Streamlit 服务，
# 避免 bare mode 运行导致无页面（PyCharm 里直接点运行也能正常打开）。
# 通过 `streamlit run` 正常启动时此处不生效（已有 ScriptRunContext）。
if __name__ == "__main__":
    from streamlit.runtime.scriptrunner import get_script_run_ctx
    if get_script_run_ctx() is None:
        import os
        import sys
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", os.path.abspath(__file__)] + sys.argv[1:]
        sys.exit(stcli.main())

st.set_page_config(page_title="PMI 数据一致性比对工具", layout="wide")
st.title("PMI 数据一致性比对工具")
st.caption(
    "以 SFA 报告的**语义 PMI 通道**为公差真值、**图形 PMI 通道**为非语义参照，"
    "与内部项目提取结果按 STEP 实体 ID 精确关联；差异按「语义 PMI / 基准 / 注释」三层归因。"
)

# ================= 文本清洗 =================
def clean_text(s):
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'\(composite\s*with\s*\d+\)', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\[.*?\]', '', s)
    s = re.sub(r'<.*?>', '', s)
    s = re.sub(r'[\s\u00a0\u3000]+', '', s)
    s = s.replace('⌀', 'Ø').replace('∅', 'Ø')
    s = s.replace('⌐', '-').replace('¬', '-').replace('˥', '-').replace('﹁', '-')
    s = s.replace('\u2212', '-').replace('\u2013', '-').replace('\u2014', '-')
    s = s.replace('∨', 'V').replace('︿', 'V')
    s = s.replace('⟂', '⊥').replace('⏊', '⊥')
    s = s.replace('⌖', '⊕').replace('⊢', '⊕').replace('⌓', '-')
    s = s.replace('⎹', '|').replace('▽', 'V')
    s = s.replace('×', 'X')
    s = s.replace('|', '')
    return s.upper()

def display_clean(s):
    if not isinstance(s, str):
        return ""
    s = re.sub(r'\(composite\s*with\s*\d+\)', '', s, flags=re.IGNORECASE)
    s = re.sub(r'<br\s*/?>', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def classify_text(s):
    if not isinstance(s, str) or not s.strip():
        return "未知"
    s_strip = s.strip()
    if re.match(r'^(Text\.\d+|F\d+\.\d+|Note\.\d+)$', s_strip, re.IGNORECASE):
        return "🔖 辅助标注"
    pmi_symbols = ['±', '⊕', '⌖', '⊢', '⊥', '⟂', '⏊', '⌓', '⌐', '⌀', 'Ø', '∅', '⌒']
    if any(sym in s_strip for sym in pmi_symbols):
        return "📐 语义PMI"
    if re.search(r'\d+(\.\d+)?\s*[+\-±]', s_strip):
        return "📐 语义PMI"
    return "📄 普通文本"

# ================= 匹配状态常量与指标口径 =================
ST_MATCH = "✅ 匹配"
ST_MATCH_BASE = "✅ 匹配 (基准)"
ST_SUSPECT = "⚠️ 疑似 (人工复核)"
ST_EXTRA = "⚠️ 多余"
ST_MISS = "❌ 缺失"
ST_PARTIAL = "⚠️ 部分覆盖"

KIND_MATCH = "match"
KIND_SUSPECT = "suspect"
KIND_EXTRA = "extra"
KIND_MISS = "miss"
KIND_PARTIAL = "partial"

_STATUS_KIND = {
    ST_MATCH: KIND_MATCH,
    ST_MATCH_BASE: KIND_MATCH,
    ST_SUSPECT: KIND_SUSPECT,
    ST_EXTRA: KIND_EXTRA,
    ST_MISS: KIND_MISS,
    ST_PARTIAL: KIND_PARTIAL,
}

def row_kind(status):
    """把匹配状态归入指标口径类型。

    指标按"行归属"统计，保证每一行都被计入、不被静默丢弃：
      应提取 E（SFA 侧行） = 匹配 + 疑似 + 缺失 + 部分覆盖
      已提取 X（开发侧行） = 匹配 + 疑似 + 多余
      命中 hit            = 匹配行中未被判为"误报"的数量
    未知状态归入疑似，宁可显式进分母也不凭空消失。
    """
    return _STATUS_KIND.get(str(status).strip(), KIND_SUSPECT)

# ================= 基准判断辅助函数 =================
def is_datum_combo(s):
    if not isinstance(s, str):
        return False
    parts = [p.strip().upper() for p in s.split('|')]
    if len(parts) < 2:
        return False
    for p in parts:
        if not re.match(r'^[A-Z]{1,3}$', p):
            return False
    return True

def is_datum_single(s):
    if not isinstance(s, str):
        return False
    return bool(re.match(r'^[A-Z]{1,3}$', s.strip().upper()))

# ================= 拆分 SFA 单元格 =================
def split_sfa_cell(cell_content):
    if not isinstance(cell_content, str) or not cell_content.strip():
        return []
    content = cell_content.replace('\r\n', '\n').replace('\r', '\n')
    raw_lines = [l.strip() for l in content.split('\n')]
    raw_lines = [l for l in raw_lines if l]
    filtered = []
    for line in raw_lines:
        if re.match(r'^[▽\s]+$', line):
            continue
        if re.match(r'^[|⎹\s|]+$', line):
            continue
        filtered.append(line)
    merged = []
    i = 0
    while i < len(filtered):
        line = filtered[i]
        if i + 1 < len(filtered):
            next_line = filtered[i + 1]
            if re.match(r'^[+\-]\d*\.?\d+$', next_line):
                merged.append(line + ' ' + next_line)
                i += 2
                continue
        merged.append(line)
        i += 1
    return merged

# ================= 解析 SFA Excel =================
def get_sfa_truths(file):
    if file is None:
        return None
    try:
        df_all = pd.read_excel(file, sheet_name=None, header=None)
        target_sheet = None
        for sheet_name, df in df_all.items():
            if 'Semantic PMI' in sheet_name and 'Summary' in sheet_name:
                target_sheet = df
                break
        if target_sheet is None:
            st.error("未找到 'Semantic PMI Summary' 工作表。")
            return []
        header_row_idx = None
        for i in range(min(20, len(target_sheet))):
            row_values = target_sheet.iloc[i].astype(str).tolist()
            if any("Semantic PMI" == str(v).strip() for v in row_values):
                header_row_idx = i
                break
        if header_row_idx is None:
            st.error("未找到 'Semantic PMI' 表头行。")
            return []
        target_df = pd.DataFrame(
            target_sheet.values[header_row_idx+1:],
            columns=target_sheet.iloc[header_row_idx].values
        )
        real_col_name = 'Semantic PMI'
        if real_col_name not in target_df.columns:
            st.error(f"当前列名有：{list(target_df.columns)}")
            return []
        vals = target_df[real_col_name].dropna().astype(str).tolist()
        candidates = []
        for v in vals:
            cleaned_full = clean_text(v)
            if ("EXPECTEDPMI" in cleaned_full or "SEEHELP" in cleaned_full or
                "SEECAX" in cleaned_full or "EXACTMATCH" in cleaned_full or
                "PARTIALMATCH" in cleaned_full or "POSSIBLEMATCH" in cleaned_full or
                "NOMATCH" in cleaned_full):
                break
            sub_items = split_sfa_cell(v)
            for sub in sub_items:
                sub_cleaned = clean_text(sub)
                if sub_cleaned and not sub_cleaned.replace('#', '').replace('.', '').isdigit():
                    candidates.append({"raw": sub.strip(), "clean": sub_cleaned, "group": "SFA"})
        return candidates
    except Exception as e:
        st.error(f"解析 Excel 出错: {e}")
        return []

# ================= 解析开发数据 =================
def get_test_data(text):
    items = []
    if not text.strip():
        return items
    if '### ' in text and 'detailData' in text:
        current_group = "未分组"
        pattern = r'^###\s+\d+\.\s+(.*)$'
        for line in text.split('\n'):
            line_strip = line.strip()
            if line_strip.startswith('## '):
                current_group = line_strip.replace('## ', '').strip()
            elif line_strip.startswith('### '):
                match = re.match(pattern, line_strip)
                if match:
                    content = match.group(1).strip()
                    if content:
                        items.append({"raw": content, "clean": clean_text(content), "group": current_group})
    else:
        try:
            data = json.loads(text)
            if isinstance(data, dict) and "extractedResults" in data:
                for group_data in data["extractedResults"]:
                    group_name = group_data.get("groupName", "未分组")
                    for item in group_data.get("items", []):
                        if isinstance(item, dict) and "text" in item:
                            raw_text = str(item["text"]).strip()
                            items.append({"raw": raw_text, "clean": clean_text(raw_text), "group": group_name})
            else:
                for item in data:
                    if isinstance(item, str):
                        items.append({"raw": item.strip(), "clean": clean_text(item.strip()), "group": "未分组"})
        except json.JSONDecodeError:
            for line in text.split('\n'):
                clean_line = line.replace('|', ' ').strip()
                if clean_line:
                    items.append({"raw": clean_line, "clean": clean_text(clean_line), "group": "未分组"})
    return [item for item in items if item["clean"]]

# ================= 智能相似度 =================
def smart_similarity(sfa_clean, dev_clean):
    if not sfa_clean or not dev_clean:
        return 0.0
    shorter = min(len(sfa_clean), len(dev_clean))
    longer = max(len(sfa_clean), len(dev_clean))
    if longer > 0 and shorter / longer >= 0.4:
        if sfa_clean in dev_clean:
            return 0.95
        if dev_clean in sfa_clean:
            return 0.95
    return SequenceMatcher(None, sfa_clean, dev_clean).ratio()

# ================= 核心匹配 =================
def pair_match_dev_first(sfa_list, test_list):
    matches = []
    used_sfa_indices = set()

    sfa_single_datum_map = {}
    sfa_datum_combo_indices = set()
    sfa_datum_combos = []
    for s_idx, sfa_item in enumerate(sfa_list):
        raw_upper = sfa_item['raw'].strip().upper()
        if is_datum_combo(raw_upper):
            letters = set(p.strip().upper() for p in raw_upper.split('|'))
            sfa_datum_combos.append((s_idx, letters, sfa_item))
            sfa_datum_combo_indices.add(s_idx)
        elif is_datum_single(raw_upper):
            if raw_upper not in sfa_single_datum_map:
                sfa_single_datum_map[raw_upper] = (s_idx, sfa_item)

    dev_datum_letters = set()
    for test_item in test_list:
        raw_upper = test_item['raw'].strip().upper()
        if is_datum_single(raw_upper):
            dev_datum_letters.add(raw_upper)

    for t_idx, dev_item in enumerate(test_list):
        dev_raw = dev_item['raw']
        dev_clean = dev_item['clean']
        dev_raw_upper = dev_raw.strip().upper()
        group_name = dev_item.get("group", "未分组")
        dev_type = classify_text(dev_raw)

        matched_sfa = None
        match_score = 0
        is_datum_ref_match = False
        remark_text = ""

        if is_datum_single(dev_raw_upper):
            if dev_raw_upper in sfa_single_datum_map:
                s_idx, sfa_item = sfa_single_datum_map[dev_raw_upper]
                if s_idx not in used_sfa_indices:
                    matched_sfa = sfa_item
                    match_score = 0.95
                    is_datum_ref_match = True
                    used_sfa_indices.add(s_idx)
                    remark_text = "SFA独立基准"
            if matched_sfa is None:
                for s_idx, letters, sfa_item in sfa_datum_combos:
                    if dev_raw_upper in letters:
                        matched_sfa = sfa_item
                        match_score = 0.95
                        is_datum_ref_match = True
                        remark_text = "SFA基准组合包含此基准"
                        break
        else:
            best_score = 0
            best_s_idx = None
            for s_idx, sfa_item in enumerate(sfa_list):
                if s_idx in used_sfa_indices:
                    continue
                if s_idx in sfa_datum_combo_indices:
                    continue
                score = smart_similarity(sfa_item['clean'], dev_clean)
                if score > best_score:
                    best_score = score
                    best_s_idx = s_idx
            if best_s_idx is not None and best_score >= 0.5:
                matched_sfa = sfa_list[best_s_idx]
                match_score = best_score
                used_sfa_indices.add(best_s_idx)

        if matched_sfa is not None:
            if is_datum_ref_match:
                status = ST_MATCH_BASE
                remark = remark_text
            elif match_score >= 0.80:
                status = ST_MATCH
                remark = "自动匹配"
            else:
                status = ST_SUSPECT
                remark = f"与SFA相似度 {match_score*100:.0f}%"

            matches.append({
                "分组": group_name,
                "SFA原始值": display_clean(matched_sfa['raw']),
                "开发原始值": display_clean(dev_raw),
                "相似度": f"{match_score*100:.1f}%",
                "匹配状态": status,
                "文本类型": dev_type,
                "人工校验": "待定",
                "备注": remark
            })
        else:
            matches.append({
                "分组": group_name,
                "SFA原始值": "",
                "开发原始值": display_clean(dev_raw),
                "相似度": "0.0%",
                "匹配状态": ST_EXTRA,
                "文本类型": dev_type,
                "人工校验": "待定",
                "备注": "无SFA对应项"
            })

    combo_partial = {}
    for s_idx, letters, sfa_item in sfa_datum_combos:
        if letters.issubset(dev_datum_letters):
            used_sfa_indices.add(s_idx)
        else:
            covered = letters & dev_datum_letters
            if covered:
                combo_partial[s_idx] = covered

    for s_idx, sfa_item in enumerate(sfa_list):
        if s_idx in used_sfa_indices:
            continue
        if s_idx in combo_partial:
            covered = combo_partial[s_idx]
            all_letters = set(p.strip().upper() for p in sfa_item['raw'].split('|'))
            missing = all_letters - covered
            matches.append({
                "分组": "⚠️ SFA部分覆盖",
                "SFA原始值": display_clean(sfa_item['raw']),
                "开发原始值": "",
                "相似度": "0.0%",
                "匹配状态": ST_PARTIAL,
                "文本类型": classify_text(sfa_item['raw']),
                "人工校验": "待定",
                "备注": f"基准组合已提取 {'/'.join(sorted(covered))}，缺 {'/'.join(sorted(missing))}"
            })
            continue
        matches.append({
            "分组": "❌ SFA未匹配",
            "SFA原始值": display_clean(sfa_item['raw']),
            "开发原始值": "",
            "相似度": "0.0%",
            "匹配状态": ST_MISS,
            "文本类型": classify_text(sfa_item['raw']),
            "人工校验": "待定",
            "备注": "开发未提取"
        })

    matches.sort(key=lambda x: (x["分组"], x["匹配状态"]))
    return matches

# ================= 指标计算 =================
def _metrics_from_rows(rows):
    """按"行归属"口径汇总指标，返回 (recall, precision, hit, miss, extra, suspect, partial)。

    每个 (kind, verdict) 行的归属：
      match   + 非误报 -> 命中：进 E、进 X
      match   + 误报   -> 命中作废：进 E（SFA 项仍存在），不进 X（开发项是误提取）
      suspect          -> 进 E；未判误报时进 X；不算命中
      miss    + 非误报 -> 进 E，计缺失
      extra   + 非误报 -> 进 X，计多余
      partial + 非误报 -> 进 E，计部分覆盖
      标"误报"的 miss/partial/extra 行从对应分母剔除（该差异不成立）
    """
    hit = miss = extra = suspect = partial = 0
    expected = 0
    extracted = 0
    for kind, verdict in rows:
        is_false_alarm = (verdict == '误报')
        if kind == KIND_MATCH:
            expected += 1
            if is_false_alarm:
                miss += 1
            else:
                hit += 1
                extracted += 1
        elif kind == KIND_SUSPECT:
            expected += 1
            suspect += 1
            if not is_false_alarm:
                extracted += 1
        elif kind == KIND_MISS:
            if not is_false_alarm:
                expected += 1
                miss += 1
        elif kind == KIND_EXTRA:
            if not is_false_alarm:
                extracted += 1
                extra += 1
        elif kind == KIND_PARTIAL:
            if not is_false_alarm:
                expected += 1
                partial += 1
    recall = hit / expected * 100 if expected > 0 else 0.0
    precision = hit / extracted * 100 if extracted > 0 else 0.0
    return recall, precision, hit, miss, extra, suspect, partial

def calculate_metrics(df):
    """机器初始指标：等价于所有行"待定"时的实时指标，两者口径一致。"""
    rows = [(row_kind(row['匹配状态']), '待定') for _, row in df.iterrows()]
    recall, precision = _metrics_from_rows(rows)[:2]
    return recall, precision

def calculate_metrics_with_verdicts(df, verdicts):
    """人工校验后的实时指标：疑似计入分母，匹配行的"误报"会作废命中。"""
    rows = []
    for idx, row in df.iterrows():
        verdict = verdicts.get(idx, row.get('人工校验', '待定'))
        rows.append((row_kind(row['匹配状态']), verdict))
    return _metrics_from_rows(rows)

# ================= 新版 UI：实体 ID 精确关联比对 =================
import pmi_core as core

_STATE = {"rows": None, "verdicts": {}, "meta": {}, "dev_raw": "",
          "sfa_warn": [], "sfa_stat": "", "sfa_checks": [], "sfa_recipes": [],
          "suspected": [], "link_stats": {}, "truth_missing": False}
for _k, _v in _STATE.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ---------------------------- 表格列视图 ----------------------------
# 核心列：判定 + 分组 + 比对内容 + 结论，默认全部显示
# 「分组」保持在核心列 —— 同一次比对可能含多个视图分组，不显示会分不清条目归属。
VIEW_COLS = [
    "匹配状态", "分组", "开发名称", "开发标注",
    "SFA语义文本",
    "提取缺陷", "人工校验", "备注",
]
# 诊断列：默认隐藏，需要溯源时在侧边栏「显示字段」里勾选恢复
# `SFA图形文本` 也在其中 —— 它依赖导出时勾选 Graphic Presentation PMI，
# 日常核查看 `SFA语义文本` 即可，比对内容不一致时再打开它。
DIAG_COLS = [
    "SFA图形文本", "关联键", "Handle", "SFA语义ID", "SFA实体类型",
    "层级", "类别", "缺陷详情", "关联路径", "多视图",
]
# 全部字段的固定顺序：表格列序一律按这个排，与用户的勾选先后无关。
ALL_COLS = VIEW_COLS + DIAG_COLS
# 列宽（未列出的按 small）
COL_WIDTH = {
    "分组": "small",
    "匹配状态": "small",
    "开发名称": "small",
    "开发标注": "medium",
    "SFA语义文本": "medium",
    "SFA图形文本": "medium",
    "备注": "large",
    "缺陷详情": "medium",
}
# 缺陷码释义（表格提示用）
DEFECT_LEGEND = ("`ENC` 编码损坏 · `CNT` 数量前缀不符 · `SYM` 符号丢失 · `NUM` 数值缺失 · "
                 "`EMP` 标题为空 · `MAP` handle-name 不一致 · "
                 "`DUP` handle 在**同一分组内**重复（跨视图复用不算缺陷）")

# 关联路径的中文释义。key 与 pmi_core._lookup_semantic_for_dev 的返回值一一对应；
# 出现字典外的 key 直接原样显示 —— 宁可显示英文，也不要静默吞掉。
PATH_LABELS = {
    "gt-1hop": "GT 1 跳（图形 → 语义表）",
    "dim-2hop": "尺寸 2 跳（图形 → dimensional_* → dcr → 语义表）",
    "datum_target-1hop": "基准目标 1 跳",
    "datum_feature-1hop": "基准 1 跳（图形 → datum_feature）",
    "datum-2hop": "基准 2 跳（ID 邻接，旧报告兜底）",
    "entity-no-semantic": "尺寸实体存在但无语义值（不进语义分母）",
    "no-truth": "不可比（SFA 报告不含语义 PMI，无真值）",
    "note": "注释 / 标签类",
    "datum_system": "基准体系",
    "reverse": "SFA 有、开发侧无",
    "none": "断链（未关联到任何实体）",
}


def build_column_config(cols):
    """只为当前显示的列生成配置，避免多余 config 干扰。"""
    cfg = {}
    for c in cols:
        if c == "人工校验":
            cfg[c] = st.column_config.SelectboxColumn(
                c, options=["待定", "误报", "确认为Bug"], width="small")
        elif c in ("Handle", "SFA语义ID"):
            cfg[c] = st.column_config.NumberColumn(c, width="small")
        else:
            cfg[c] = st.column_config.TextColumn(c, width=COL_WIDTH.get(c, "small"))
    return cfg


# ---------------------------- 侧边栏 ----------------------------
with st.sidebar:
    st.header("1. 上传 SFA 报告 (Excel)")
    sfa_file = st.file_uploader("SFA 生成的 .xlsx", type=["xlsx"])
    st.caption(
        "生成报告时建议勾选：**Semantic Representation PMI** + **Graphic Presentation PMI**；"
        "Spreadsheet 关闭「数值四舍五入」；Maximum Rows 调到最大。"
    )
    st.divider()
    st.header("2. 视图选项")
    show_note = st.checkbox("显示注释/标签类条目", value=True)
    only_issue = st.checkbox("只看非「✅ 匹配」条目", value=False)
    only_defect = st.checkbox("只看含「提取缺陷」条目", value=False,
                              help="缺陷指开发侧数据本身有问题（乱码 / 丢符号 / 丢数量前缀），"
                                   "与分析指标解耦统计。")
    st.divider()
    st.header("3. 表格字段")
    selected_cols = st.multiselect(
        "显示字段",
        ALL_COLS,
        default=VIEW_COLS,
        key="field_picker",
        placeholder="默认只显示核心列",
        help="默认显示核心列（判定 / 分组 / 比对内容 / 结论）。"
             "需要溯源时勾选其它字段即可恢复显示，例如 `SFA图形文本`、`Handle`、"
             "`SFA语义ID`、`SFA实体类型`、`关联路径`；全部勾上即完整视图，"
             "也可以反选把核心列临时收起。",
    )
    _picked = [c for c in ALL_COLS if c in set(selected_cols)]
    _hid_now = [c for c in ALL_COLS if c not in set(_picked)]
    if not _picked:
        st.warning("至少保留一列，已暂时回落为默认视图。")
    if _hid_now:
        st.caption(f"当前 {len(_picked)} 列，已隐藏 {len(_hid_now)} 列："
                   + "、".join(f"`{c}`" for c in _hid_now))
    else:
        st.caption(f"完整视图：{len(_picked)} 列，无隐藏字段。")

# ---------------------------- 输入 ----------------------------
st.header("2. 内部项目提取结果（markdown）")
md_file = st.file_uploader("上传 markdown 文件", type=["md", "txt", "json"], key="md_file")
md_paste = st.text_area("或直接粘贴内容：", height=130, key="md_paste")

# 统一取当前输入（供自检与比对共用）
raw_now = ""
if md_file is not None:
    raw_now = md_file.getvalue().decode("utf-8", errors="replace")
elif md_paste.strip():
    raw_now = md_paste

if st.button("🚀 开始比对", type="primary"):
    if sfa_file is None:
        st.warning("请先上传 SFA Excel 报告！")
    elif not raw_now.strip():
        st.warning("请提供内部项目的 markdown（上传文件或粘贴内容）！")
    else:
        truth = None
        try:
            truth = core.load_sfa(sfa_file)
        except Exception as e:
            st.error(f"SFA 报告解析失败：{e}")
        if truth is not None:
            items = core.parse_dev_markdown(raw_now)
            if not items:
                st.error("未能解析出任何标注条目，详见下方「🔎 解析自检」。")
            else:
                rows = core.match_items(truth, items)
                st.session_state.rows = rows
                st.session_state.verdicts = {r.key: "待定" for r in rows}
                st.session_state.meta = core.summarize(rows, truth, items)
                st.session_state.dev_raw = raw_now
                st.session_state.sfa_warn = list(truth.warnings)
                st.session_state.truth_missing = bool(truth.truth_missing)
                st.session_state.sfa_stat = (
                    f"语义表 {len(truth.semantic)} 项"
                    + ("（**无真值**）" if truth.truth_missing else "")
                    + f" · draughting_callout {len(truth.dc)} 条 · "
                    f"图形标注 {len(truth.ta)} 条"
                    f"（{truth.link_channel or '关联表缺失'}） · "
                    f"datum {len(truth.datum)} 项 · dcr {len(truth.dcr_by_dim)} 项"
                )
                st.session_state.sfa_checks = [ck.__dict__ for ck in truth.checks]
                st.session_state.sfa_recipes = [
                    {"key": cr.key, "sheet": cr.sheet, "header_row": cr.header_row,
                     "source": cr.source, "n_cols": cr.n_cols, "n_rows": cr.n_rows,
                     "n_loaded": cr.n_loaded, "n_rows_grouped": cr.n_rows_grouped,
                     "zero_reason": cr.zero_reason,
                     "resolved": cr.resolved, "notes": cr.notes}
                    for cr in truth.recipes
                ]
                st.session_state.link_stats = core.link_stats(rows)
                st.session_state.suspected = [
                    {"开发条目": sl.key, "标题": sl.title, "name": sl.name,
                     "疑似对应": " / ".join(
                         f"{c['ta_name']}（{c['why']}）" for c in sl.candidates)}
                    for sl in core.suspect_links(items, truth, rows)
                ]
                st.rerun()

# ---------------------------- 解析自检 ----------------------------
# 只读地跑一遍解析，把「为什么一条都没对上」显式暴露出来，避免静默全红。
if raw_now.strip():
    _items_probe, _diag = core.parse_dev_markdown_ex(raw_now)
    _probs = _diag.problems()
    _show = (not _diag.healthy) or _diag.items == 0
    with st.expander(
        ("🔎 解析自检 —— " + ("⚠️ 发现问题" if _probs else "✅ 解析正常")
         + f"　（{_diag.summary()}）"),
        expanded=_show,
    ):
        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric("解析条目", _diag.items)
        s2.metric("带 handle", f"{_diag.with_handle}/{_diag.items}" if _diag.items else "0/0")
        s3.metric("带 name", f"{_diag.with_name}/{_diag.items}" if _diag.items else "0/0")
        s4.metric("未解析到 detailData", len(_diag.detail_missing))
        s5.metric("JSON 解析失败", len(_diag.json_failed))

        if _probs:
            st.error("**handle/name 是 ID 关联的唯一依据**。以下问题会让条目全部判为「⚠️ 多余」，"
                     "同时 SFA 侧全部判为「❌ 缺失」：\n\n- "
                     + "\n- ".join(_probs))
        else:
            st.success("解析出的每条标注都带上了 handle 与 name，ID 关联链路具备前提条件。")

        if _diag.groups:
            st.caption(f"识别到 {len(_diag.groups)} 个分组：" + "、".join(_diag.groups[:12])
                       + ("…" if len(_diag.groups) > 12 else ""))
        if _diag.headings_unmatched:
            st.caption(f"以下 {len(_diag.headings_unmatched)} 行像标题但未匹配"
                       f"`{_diag.h3_pattern}`（标题行须形如 `### 1. 标题`）：")
            st.code("\n".join(_diag.headings_unmatched), language="text")
        if _diag.detail_missing or _diag.json_failed or _diag.key_missing:
            st.caption("逐条问题明细（条目前 20 条）：")
            st.code("\n".join((_diag.detail_missing + _diag.json_failed
                               + _diag.key_missing)[:20]), language="text")

        st.caption(
            "期望的 markdown 结构：`## 分组` → `- 标注数量：N 条` → `### 1. 标题` → "
            "`detailData:`（**独占一行、半角冒号**）→ JSON，"
            "JSON 字段名须为小写 `handle` / `name` / `type`。"
        )

# ---------------------------- 报告体检 ----------------------------
# 关联链断在 SFA 侧（索引表缺列 / 缺表 / 列序变）与开发侧解析失败的表象完全一样，
# 这里把「装载了什么、列是怎么定位的、引用能不能落地、哪一环断了」一次摊开。
_sfa_stat = st.session_state.get("sfa_stat") or ""
_sfa_warn = list(st.session_state.get("sfa_warn") or [])
_checks = list(st.session_state.get("sfa_checks") or [])
_recs = list(st.session_state.get("sfa_recipes") or [])
_linkstats = dict(st.session_state.get("link_stats") or {})
_suspected = list(st.session_state.get("suspected") or [])
_truth_missing = bool(st.session_state.get("truth_missing"))
_bad_checks = [c for c in _checks if not c.get("ok")]

if _sfa_stat or _checks:
    _problems = len(_bad_checks) + len(_sfa_warn) + len(_suspected) + _linkstats.get("none", 0)
    with st.expander(
        "🧭 报告体检 —— " + ("⚠️ 发现问题" if _problems else "✅ 全部正常")
        + f"　（{_sfa_stat or '未装载'}）",
        expanded=bool(_problems),
    ):
        st.caption("体检结论：" + ("全部正常" if not _problems else
                                 f"发现 {_problems} 处需要关注 —— 校验不通过 "
                                 f"{len(_bad_checks)} · 装载告警 {len(_sfa_warn)} · "
                                 f"关联断裂 {_linkstats.get('none', 0)} · "
                                 f"疑似对应 {len(_suspected)}"))
        if _truth_missing:
            # 无真值是最该被放在最前面的一条：不点破，满屏 0% 会被读成
            # 「开发侧提取全错」，而真正的原因是拖错了报告文件。
            st.error(
                "**这份 SFA 报告不含任何语义 PMI，SFA 侧没有可比真值。**\n\n"
                "开发侧条目全部记为「⛔ 无可比真值」，既不进召回率也不进精确率 —— "
                "这不是「开发侧提取全错」，而是**选错了报告文件**。\n\n"
                "- 典型原因：拿的是「图形专用」导出（NIST 命名里带 `-tg`，"
                "只有 tessellated 图形 PMI、没有 `Semantic PMI Summary`）。\n"
                "- 处理：换用同一测试件的**语义版**报告重跑。下方告警里会列出"
                "同目录下含 `Semantic PMI Summary` 的候选文件。"
            )
        for _w in _sfa_warn:
            st.warning(_w)

        st.markdown("**① 装载与列定位**")
        st.caption(
            "列位置一律按表头列名定位，不写死列号 —— SFA 的列数与列序会随版本和"
            "导出勾选变化。「列来源」显示 *默认列号（脆弱）* 时，该表没识别出表头，"
            "列序一变就会静默失配。"
        )
        if _recs:
            st.dataframe(pd.DataFrame([{
                "表": r["sheet"],
                "列来源": "表头识别" if r["source"] == "header" else "默认列号（脆弱）",
                "表头行": (r["header_row"] + 1) if r["header_row"] >= 0 else "—",
                "列数": r["n_cols"],
                # DMIA 一条 callout 占多行，行数 > 装载条数是正常的，须标出聚合
                "数据行": (f'{r["n_rows"]}（聚合为 {r["n_rows_grouped"]} 组）'
                           if r.get("n_rows_grouped") else r["n_rows"]),
                "装载": r["n_loaded"],
                "用到的列": " · ".join(f"{k}={v}" for k, v in r["resolved"].items()),
                "备注": "；".join(x for x in (
                    [r.get("zero_reason")] if r["n_loaded"] == 0 and r.get("zero_reason")
                    else []) + list(r["notes"])),
            } for r in _recs]), hide_index=True, width="stretch")

        if _checks:
            st.markdown("**② 交叉校验**")
            st.caption(
                "不依赖列语义，只看「读到没有、数量对不对、引用能否落地」。"
                "结构变化导致整表被跳过时，这里会先于结果表报警。"
            )
            st.dataframe(pd.DataFrame([{
                "校验项": c["name"], "结果": "✅" if c["ok"] else "❌",
                "值": c["value"], "说明": c["detail"],
            } for c in _checks]), hide_index=True, width="stretch")

        if _linkstats:
            st.markdown("**③ 关联路径分布**")
            st.caption(
                "`断链` 表示该条目没能关联到任何 SFA 实体。"
                "开发侧几乎全是断链 = 关联链断了，而不是两边数据真的对不上。"
            )
            st.dataframe(pd.DataFrame(
                [{"关联路径": PATH_LABELS.get(k, k), "条数": v}
                 for k, v in _linkstats.items()]),
                hide_index=True, width="stretch")

        if _suspected:
            st.markdown("**④ 疑似对应（仅供参考，不参与判定）**")
            st.caption(
                "对关联失败条目做的模糊匹配，**不计入任何指标分子分母**，"
                "仅用于判断是「名字对不上」还是「整条链断了」。"
            )
            st.dataframe(pd.DataFrame(_suspected), hide_index=True, width="stretch")

        if not _problems:
            st.caption(
                "ID 关联靠这几张表打通：`draughting_callout.ID`（＝开发侧 handle）→ "
                "关联表的语义引用 → 语义实体。关联表有**两种等价形态**，取其一即可："
                "`tessellated_annotation_occurrence`（含 tessellated 呈现，"
                "引用在 `Associated Semantic PMI` 列）或 "
                "`draughting_model_item_association`（非 tessellated 导出，"
                "引用在 `definition` 列，一条 callout 多行需聚合）。"
                "其中 GT 1 跳、尺寸经 `dimensional_characteristic_repr` 2 跳、"
                "基准经 `datum_feature` 1 跳（旧报告 ID 相邻时走 2 跳兜底）、基准目标 1 跳。"
            )

# ---------------------------- 结果 ----------------------------
if st.session_state.rows:
    rows = st.session_state.rows
    meta = st.session_state.meta
    metrics_ph = st.container()

    st.divider()
    st.subheader("📋 详细比对结果")

    _dev_rows = [r for r in rows if not r.key.startswith("SFA")]
    if _dev_rows and all(r.status == core.ST_ABSENT for r in _dev_rows):
        st.error(
            f"开发侧 {len(_dev_rows)} 条全部记为「{core.ST_ABSENT}」"
            "——**SFA 报告里没有任何语义 PMI 真值**，两边没有共同基准可比。"
            "这不是提取错误，也不是「多余」：请换用同一测试件的语义版 SFA 报告"
            "（SFA 导出时勾选语义 PMI；NIST 命名里带 `-tg` 的是图形专用变体）。"
            "展开「🧭 报告体检」可看到同目录下的候选报告。"
        )
    elif _dev_rows and all(r.status == core.ST_EXTRA for r in _dev_rows):
        st.error(
            f"开发侧 {len(_dev_rows)} 条**全部**判为「⚠️ 多余」，同时 SFA 侧大量判为「❌ 缺失」"
            "——这是 **ID 关联链没建立** 的典型特征，而不是两边数据真的对不上。"
            "两处根因按顺序排查：① 展开「🔎 解析自检」看 handle/name 有没有解析出来；"
            "② 展开「🧭 报告体检」看装载、列定位与关联路径分布，"
            "其中「疑似对应」能直接区分是名字对不上还是整条链断了。"
        )

    f1, f2, f3 = st.columns(3)
    layer_sel = f1.selectbox("层级", ["全部"] + sorted({r.layer for r in rows}))
    groups = sorted({r.group for r in rows if r.group})
    group_sel = f2.selectbox("分组", ["全部"] + groups)
    status_sel = f3.selectbox("匹配状态", ["全部"] + sorted({r.status for r in rows}))

    view = rows
    if layer_sel != "全部":
        view = [r for r in view if r.layer == layer_sel]
    if group_sel != "全部":
        view = [r for r in view if r.group == group_sel]
    if status_sel != "全部":
        view = [r for r in view if r.status == status_sel]
    if only_issue:
        view = [r for r in view if r.status != core.ST_HIT]
    if only_defect:
        view = [r for r in view if r.defects]
    if not show_note:
        view = [r for r in view if r.layer != core.LAYER_NOTE]

    if not view:
        st.info("当前筛选条件下没有条目。")
    else:
        keys = [r.key for r in view]
        # 列序固定按 ALL_COLS，与勾选先后无关；一列不剩时回落默认视图
        cols = [c for c in ALL_COLS if c in set(selected_cols)] or list(VIEW_COLS)
        df_show = pd.DataFrame([
            {c: d[c] for c in cols}
            for r in view
            for d in [{**r.as_dict(), "人工校验": st.session_state.verdicts.get(r.key, "待定")}]
        ])
        # 只有结论两列可编辑，其余只读，防止误改真值文本
        readonly = [c for c in cols if c not in ("人工校验", "备注")]
        _hid = [c for c in ALL_COLS if c not in cols]
        st.caption(
            f"共 {len(view)} 行 × {len(cols)} 列。"
            f"「提取缺陷」为空表示无缺陷；{DEFECT_LEGEND}。"
            + (f"　已隐藏 {len(_hid)} 列（侧边栏「显示字段」勾选即恢复）："
               + "、".join(f"`{c}`" for c in _hid) + "。" if _hid else "")
            + ("　列宽不够可左右拖动表头。" if len(cols) > 8 else "")
        )
        edited = st.data_editor(
            df_show,
            column_config=build_column_config(cols),
            disabled=readonly,
            hide_index=True,
            width="stretch",
            height=min(760, 42 + 35 * len(view)),
            key=f"editor_{layer_sel}_{group_sel}_{status_sel}_"
                f"{hash(tuple(cols)) & 0xFFFFFF}",
        )
        for pos, k in enumerate(keys):
            if pos < len(edited):
                st.session_state.verdicts[k] = edited.iloc[pos]["人工校验"]

    # ---------------------------- 实时指标 ----------------------------
    with metrics_ph:
        m = core.compute_metrics(rows, st.session_state.verdicts)
        st.subheader("📊 指标总览")

        st.caption("① 提取质量（开发侧提取与 SFA 真值的对齐度；非语义项与注释类不进分母）")
        _na = m.truth_missing
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("语义PMI 召回率", "N/A" if _na else f"{m.recall:.2f}%",
                  help="几何公差 + 尺寸公差。非语义项不进分母。"
                       + ("　SFA 报告不含语义 PMI，指标不成立。" if _na else ""))
        c2.metric("语义PMI 精确率", "N/A" if _na else f"{m.precision:.2f}%")
        c3.metric("基准覆盖率", "N/A" if _na else f"{m.datum_coverage:.2f}%")
        c4.metric("注释类独占提取", m.note_exclusive,
                  help="SFA 语义表与图形通道均无、内部项目能提取到的条目数（内部项目增益）")
        c5.metric("待复核", sum(1 for v in st.session_state.verdicts.values() if v == "待定"))
        if _na:
            st.warning(f"SFA 侧无真值：**{m.no_truth}** 条记为「{core.ST_ABSENT}」，"
                       "三项比率不成立（不是 0%）。换用语义版报告后重跑即可。")

        st.caption("② 开发侧提取缺陷（数据本身有问题：乱码 / 丢符号 / 丢数量前缀。"
                   "与 ① 解耦——缺陷不改判定、不进召回率与精确率分母）")
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("含缺陷条目", m.defect_rows)
        d2.metric("缺陷总数", m.defect_total)
        d3.metric("缺陷占比", f"{m.defect_rate:.2f}%",
                  help="含缺陷的开发标注数 / 参与比对的开发条目数")
        d4.metric("缺陷类型数", len(m.defect_by_code))
        if m.defect_by_code:
            st.warning("缺陷分布：" + "　".join(f"**{k}** × {v}"
                                          for k, v in m.defect_by_code.items()))
            dedup = {}
            for r in rows:
                if r.defects:
                    dedup.setdefault(r.key.split("@")[0], r)
            with st.expander(f"🔧 提取缺陷清单（{len(dedup)} 条，可直接交回开发侧排查）",
                             expanded=True):
                st.dataframe(
                    pd.DataFrame([{
                        "分组": r.group,
                        "开发名称": r.dev_name,
                        "开发标注": r.dev_title,
                        "缺陷": " ".join(r.defects),
                        "详情": r.defect_detail,
                    } for k, r in dedup.items()]),
                    column_config={
                        "分组": st.column_config.TextColumn("分组", width="small"),
                        "开发名称": st.column_config.TextColumn("开发名称", width="small"),
                        "开发标注": st.column_config.TextColumn("开发标注", width="small"),
                        "缺陷": st.column_config.TextColumn("缺陷", width="small"),
                        "详情": st.column_config.TextColumn("详情", width="large"),
                    },
                    hide_index=True, width="stretch",
                )
        else:
            st.success("✅ 未检出开发侧提取缺陷。")

        st.caption(
            f"关联方式：实体 ID 精确关联（handle/name → SFA `draughting_callout` → "
            f"关联表（tessellated_annotation_occurrence 或 "
            f"draughting_model_item_association）→ 语义表），非字符串相似度。　"
            f"开发条目 {meta.get('开发条目', 0)} 条 / SFA 语义项 {meta.get('SFA语义项', 0)} 项 / "
            f"报告工作表 {meta.get('SFA表数', 0)} 张。　"
            f"口径：标注(label/note)类完全不进 PMI 分母，仅单列统计；"
            f"人工校验标「误报」的条目从对应分母剔除。"
        )
        if meta.get("警告"):
            st.warning("；".join(meta["警告"]))

    # ---------------------------- 导出 ----------------------------
    st.divider()
    export_df = pd.DataFrame([
        {**r.as_dict(), "人工校验": st.session_state.verdicts.get(r.key, "待定")} for r in rows
    ])
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="比对结果")
        pd.DataFrame([meta]).T.rename(columns={0: "值"}).to_excel(writer, sheet_name="指标汇总")
    st.download_button(
        "📥 导出比对结果 (Excel)",
        data=buf.getvalue(),
        file_name="PMI比对结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
