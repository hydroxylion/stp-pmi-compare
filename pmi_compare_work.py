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

_STATE = {"rows": None, "verdicts": {}, "meta": {}, "dev_raw": ""}
for _k, _v in _STATE.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

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

# ---------------------------- 输入 ----------------------------
st.header("2. 内部项目提取结果（markdown）")
md_file = st.file_uploader("上传 markdown 文件", type=["md", "txt", "json"], key="md_file")
md_paste = st.text_area("或直接粘贴内容：", height=130, key="md_paste")

if st.button("🚀 开始比对", type="primary"):
    raw = ""
    if md_file is not None:
        raw = md_file.getvalue().decode("utf-8", errors="replace")
    elif md_paste.strip():
        raw = md_paste
    if sfa_file is None:
        st.warning("请先上传 SFA Excel 报告！")
    elif not raw.strip():
        st.warning("请提供内部项目的 markdown（上传文件或粘贴内容）！")
    else:
        truth = None
        try:
            truth = core.load_sfa(sfa_file)
        except Exception as e:
            st.error(f"SFA 报告解析失败：{e}")
        if truth is not None:
            items = core.parse_dev_markdown(raw)
            if not items:
                st.error("未能解析出任何标注条目。请确认格式：`## 分组` / `### N. 标题` / `detailData:` + JSON。")
            else:
                rows = core.match_items(truth, items)
                st.session_state.rows = rows
                st.session_state.verdicts = {r.key: "待定" for r in rows}
                st.session_state.meta = core.summarize(rows, truth, items)
                st.session_state.dev_raw = raw
                st.rerun()

# ---------------------------- 结果 ----------------------------
if st.session_state.rows:
    rows = st.session_state.rows
    meta = st.session_state.meta
    metrics_ph = st.container()

    st.divider()
    st.subheader("📋 详细比对结果")

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
        df_show = pd.DataFrame([
            {**r.as_dict(), "人工校验": st.session_state.verdicts.get(r.key, "待定")}
            for r in view
        ])
        edited = st.data_editor(
            df_show,
            column_config={
                "关联键": st.column_config.TextColumn("关联键", width="small"),
                "分组": st.column_config.TextColumn("分组", width="small"),
                "开发标注": st.column_config.TextColumn("开发标注", width="medium"),
                "开发名称": st.column_config.TextColumn("开发名称", width="medium"),
                "Handle": st.column_config.NumberColumn("Handle", width="small"),
                "SFA语义ID": st.column_config.NumberColumn("SFA语义ID", width="small"),
                "SFA实体类型": st.column_config.TextColumn("SFA实体类型", width="medium"),
                "SFA语义文本": st.column_config.TextColumn("SFA语义文本", width="large"),
                "SFA图形文本": st.column_config.TextColumn("SFA图形文本", width="medium"),
                "层级": st.column_config.TextColumn("层级", width="small"),
                "类别": st.column_config.TextColumn("类别", width="small"),
                "匹配状态": st.column_config.TextColumn("匹配状态", width="small"),
                "提取缺陷": st.column_config.TextColumn("提取缺陷", width="small"),
                "缺陷详情": st.column_config.TextColumn("缺陷详情", width="medium"),
                "关联路径": st.column_config.TextColumn("关联路径", width="small"),
                "人工校验": st.column_config.SelectboxColumn(
                    "人工校验", options=["待定", "误报", "确认为Bug"], width="small"),
                "备注": st.column_config.TextColumn("备注", width="large"),
            },
            hide_index=True,
            width="stretch",
            key=f"editor_{layer_sel}_{group_sel}_{status_sel}",
        )
        for pos, k in enumerate(keys):
            if pos < len(edited):
                st.session_state.verdicts[k] = edited.iloc[pos]["人工校验"]

    # ---------------------------- 实时指标 ----------------------------
    with metrics_ph:
        m = core.compute_metrics(rows, st.session_state.verdicts)
        st.subheader("📊 指标总览")

        st.caption("① 提取质量（开发侧提取与 SFA 真值的对齐度；非语义项与注释类不进分母）")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("语义PMI 召回率", f"{m.recall:.2f}%",
                  help="几何公差 + 尺寸公差。非语义项不进分母。")
        c2.metric("语义PMI 精确率", f"{m.precision:.2f}%")
        c3.metric("基准覆盖率", f"{m.datum_coverage:.2f}%")
        c4.metric("注释类独占提取", m.note_exclusive,
                  help="SFA 语义表与图形通道均无、内部项目能提取到的条目数（内部项目增益）")
        c5.metric("待复核", sum(1 for v in st.session_state.verdicts.values() if v == "待定"))

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
                        "关联键": k,
                        "分组": r.group,
                        "开发名称": r.dev_name,
                        "开发标注": r.dev_title,
                        "缺陷": " ".join(r.defects),
                        "详情": r.defect_detail,
                    } for k, r in dedup.items()]),
                    hide_index=True, width="stretch",
                )
        else:
            st.success("✅ 未检出开发侧提取缺陷。")

        st.caption(
            f"关联方式：实体 ID 精确关联（handle/name → SFA `draughting_callout` → "
            f"`tessellated_annotation_occurrence` → 语义表），非字符串相似度。　"
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
