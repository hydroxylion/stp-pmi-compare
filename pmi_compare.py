import streamlit as st
import pandas as pd
import json
import re
import unicodedata

# ================= 页面配置 =================
st.set_page_config(page_title="PMI 数据一致性比对工具", layout="wide")
st.title("PMI 数据一致性比对工具 (基于SFA真值)")
st.caption("使用方法：左侧上传 SFA 导出的 Excel，右侧粘贴开发提取的 JSON 或 Markdown，点击比对。")

# ================= 核心函数定义 =================
import unicodedata
import re

def clean_text(s):
    if not isinstance(s, str):
        return ""
        
    # 1. 强制统一全角/半角
    s = unicodedata.normalize('NFKC', s)
    
    # 2. 去除所有空白字符
    s = re.sub(r'[\s\u00a0\u3000]+', '', s) 
    
    # 3. 【终极修复：字符映射字典】把所有长得像的符号统统翻译成一种！
    # 处理SFA的轮廓度/平面度符号（⌐, ¬, ˥等）为普通减号（-）
    s = s.replace('⌐', '-').replace('¬', '-').replace('˥', '-').replace('﹁', '-')
    
    # 处理开发可能输入的各种横杠、减号（Unicode数学减号、长破折号）为普通减号
    s = s.replace('\u2212', '-').replace('\u2013', '-').replace('\u2014', '-')
    
    # 处理SFA的“或”符号（∨, ︿）为字母 V
    s = s.replace('∨', 'V').replace('︿', 'V')
    
    # 处理垂直度（⊥, ⟂, ⏊）统一为⊥
    s = s.replace('⟂', '⊥').replace('⏊', '⊥')
    
    # 处理位置度（⊕, ⌖, ⊢）统一为⊕
    s = s.replace('⌖', '⊕').replace('⊢', '⊕')
    
    # 处理直径符号
    s = s.replace('⌀', 'Ø').replace('∅', 'Ø')
    
    # 4. 转大写
    return s.upper()

def get_sfa_truths(file):
    if file is None:
        return None
    try:
        # 强制全部读取，不默认取第一行做表头
        df_all = pd.read_excel(file, sheet_name=None, header=None)
        
        # 精准定位目标Sheet
        target_sheet = None
        for sheet_name, df in df_all.items():
            if 'Semantic PMI' in sheet_name and 'Summary' in sheet_name:
                target_sheet = df
                break
                
        if target_sheet is None:
            st.error("未找到 'Semantic PMI Summary' 工作表，请检查文件名。")
            return []
            
        # 找到写着“Semantic PMI”的那一行表头
        header_row_idx = None
        for i in range(min(20, len(target_sheet))):
            row_values = target_sheet.iloc[i].astype(str).tolist()
            if any("Semantic PMI" == str(v).strip() for v in row_values):
                header_row_idx = i
                break
                
        if header_row_idx is None:
            st.error("未能在 Excel 顶部找到 'Semantic PMI' 表头行。")
            return []
            
        # 以该行为表头，重新构建DataFrame
        target_df = pd.DataFrame(target_sheet.values[header_row_idx+1:], columns=target_sheet.iloc[header_row_idx].values)
        
        # 精确提取目标列
        real_col_name = 'Semantic PMI'
        if real_col_name not in target_df.columns:
            st.error(f"当前列名有：{list(target_df.columns)}，未找到 'Semantic PMI'。")
            return []
            
        vals = target_df[real_col_name].dropna().astype(str).tolist()
        
        candidates = []
        for v in vals:
            cleaned = clean_text(v)
            
            # 【核心修复】：遇到SFA底部的说明文字（如Expected PMI），强制停止提取！
            if ("EXPECTEDPMI" in cleaned or "SEEHELP" in cleaned or "SEECAX" in cleaned 
                or "EXACTMATCH" in cleaned or "PARTIALMATCH" in cleaned or "POSSIBLEMATCH" in cleaned 
                or "NOMATCH" in cleaned):
                break
            
            # 过滤掉纯数字、空值
            if cleaned and not cleaned.replace('#', '').replace('.', '').isdigit():
                candidates.append(cleaned)
                
        # 【核心修复】：移除了去重逻辑，保留SFA中的重复标注项！
        return candidates
        
    except Exception as e:
        st.error(f"解析 Excel 出错: {e}")
        return []

def get_test_data(text):
    items = []
    if not text.strip():
        return items
    try:
        data = json.loads(text)
        
        # 如果标准结构是: {"extractedResults": [{"groupName": "...", "items": [{"id": "...", "text": "..."}]}]}
        if isinstance(data, dict) and "extractedResults" in data:
            for group in data["extractedResults"]:
                for item in group.get("items", []):
                    # 核心修复：只提取 text 字段，忽略 id 和 groupName
                    if isinstance(item, dict) and "text" in item:
                        items.append(str(item["text"]).strip())
        else:
            # 兜底：如果是纯数组（如 ['A', 'A | B | C']）
            for item in data:
                if isinstance(item, str):
                    items.append(item.strip())
                    
    except json.JSONDecodeError:
        # 如果是 Markdown 格式
        lines = text.split('\n')
        for line in lines:
            clean_line = line.replace('|', ' ').strip()
            clean_line = re.sub(r'^\d+\s+', '', clean_line)
            if clean_line:
                items.append(clean_line)
    
    # 统一清洗
    return list(dict.fromkeys([clean_text(item) for item in items if clean_text(item)]))

# ================= 主界面逻辑 =================

# 左侧侧边栏：上传SFA Excel
with st.sidebar:
    st.header("1. 上传 SFA 报告 (Excel)")
    sfa_file = st.file_uploader("请选择 SFA 生成的 Excel 文件", type=["xlsx", "xls"])

# 右侧主界面：粘贴待测数据
st.header("2. 粘贴待测数据 (JSON / Markdown)")
target_data = st.text_area("请将开发提取的 JSON 或 Markdown 内容粘贴在此处：", height=300)

# 比对执行按钮
if st.button("🚀 一键比对"):
    if not sfa_file:
        st.warning("请先上传 SFA Excel 文件！")
    elif not target_data.strip():
        st.warning("请粘贴待比对的 JSON / Markdown 数据！")
    else:
        # 获取真值与待测值
        sfa_truths = get_sfa_truths(sfa_file)
        test_vals = get_test_data(target_data)
        
        if not sfa_truths:
            st.error("SFA 提取结果为空，请检查 Excel 格式！")
        elif not test_vals:
            st.error("待测数据提取结果为空，请检查粘贴的内容格式！")
        else:
            # 将列表转换为集合以便快速比对
            sfa_set = set(sfa_truths)
            test_set = set(test_vals)
            
            # 计算指标
            tp = len(sfa_set & test_set) # 真阳性：找对的
            fn = len(sfa_set - test_set) # 假阴性：遗漏的
            fp = len(test_set - sfa_set) # 假阳性：多提取的
            
            recall = tp / (tp + fn) * 100 if (tp + fn) > 0 else 0
            precision = tp / (tp + fp) * 100 if (tp + fp) > 0 else 0
            
            # 展示指标
            st.subheader("📊 指标结果")
            col1, col2 = st.columns(2)
            col1.metric("召回率 (Recall)", f"{recall:.2f}%")
            col2.metric("精确率 (Precision)", f"{precision:.2f}%")
            
            # 展示具体差异
            st.subheader("❌ 具体比对不成功的地方 (缺失项)")
            missing_items = list(sfa_set - test_set)
            if missing_items:
                missing_df = pd.DataFrame({"SFA真值 (缺失)": missing_items})
                st.dataframe(missing_df, use_container_width=True)
            else:
                st.success("太棒了！SFA 真值全部提取成功，无缺失项。")
            
            # 展示多提取
            st.subheader("⚠️ 多余的提取项 (开发多提取了)")
            extra_items = list(test_set - sfa_set)
            if extra_items:
                extra_df = pd.DataFrame({"待测数据 (多余)": extra_items})
                st.dataframe(extra_df, use_container_width=True)
            else:
                st.success("没有多余提取项，数据精度完美。")
