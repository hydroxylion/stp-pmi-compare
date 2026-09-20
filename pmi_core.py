# -*- coding: utf-8 -*-
"""
PMI 比对核心模块
==================

真值源：NIST STEP File Analyzer (SFA) 输出的 xlsx 报告
被测源：内部项目导出的 markdown（每条含 handle / name / type / detailData）

**关联方式是实体 ID 精确关联，不是字符串相似度。**

关联链路（均已用 nist_ftc_07 NIST 测试件实测验证）::

    md.handle  ==  SFA.draughting_callout.ID
    md.name    ==  SFA.draughting_callout.name  ==  关联表的 name（见下）

    标注 ↔ 语义实体有两个等价通道，取其一即可（SfaTruth.link_channel 标明用了哪个）：
      a) tessellated_annotation_occurrence —— 一条标注一行，
         `Associated Semantic PMI` 列直接给语义引用（AP242 含 tessellated 呈现）
      b) draughting_model_item_association —— 非 tessellated 导出（如 ctc_05-e1），
         `identified_item` = draughting_callout N，`definition` 给关联实体，
         一条 callout 多行，须按 callout 聚合

    语义映射四条路径：
      GT / FCF : 关联表 → 末尾 ID                              -> 语义表 ID      (1 跳)
      DIM      : 关联表 → dimensional_size / _location ID
                    -> dcr.dimension 匹配 -> dcr.ID            -> 语义表 ID      (2 跳)
      datum    : 关联表 → datum_feature.ID                     -> 基准字母       (1 跳)
      datum_target : 关联表 → datum_target 实体 ID              -> 语义表 ID      (1 跳)

指标分层（口径已确认）：
  * 语义 PMI 层（主指标）：几何公差 + 尺寸公差；非语义项不进 precision 分母
  * 基准层（单列）：SFA `datum` 表 vs md 的 datum 标注
  * 非语义/注释层（单列）：label / note 类完全不进 PMI 分母

依赖：openpyxl
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import openpyxl

# --------------------------------------------------------------------------
# 状态常量
# --------------------------------------------------------------------------
ST_HIT      = "✅ 匹配"
ST_HIT_DIFF = "⚠️ 文本差异"
ST_SUSPECT  = "⚠️ 疑似 (人工复核)"
ST_PARTIAL  = "⚠️ 部分覆盖"
ST_MISS     = "❌ 缺失"
ST_EXTRA    = "⚠️ 多余"
ST_GAIN     = "🔷 非语义增益"
ST_NOTE     = "📝 注释文本"
# SFA 报告本身不含语义 PMI（图形专用导出，如 NIST 的 `-tg` 变体）：
# 两侧没有共同真值可比，「多余 / 缺失」无从谈起，必须与 ST_EXTRA 分开。
ST_ABSENT   = "⛔ 无可比真值"

# 条目类别
KIND_GT    = "gt"      # 几何公差（FCF）
KIND_DIM   = "dim"     # 尺寸公差
KIND_DATUM = "datum"   # 基准
KIND_NOTE  = "note"    # 注释 / 标签文本

# 指标分层
LAYER_SEM   = "语义PMI"
LAYER_DATUM = "基准"
LAYER_NOTE  = "非语义注释"
LAYER_NOTRUTH = "不可比"          # SFA 侧无真值，不进任何指标分母

KIND_LAYER = {
    KIND_GT: LAYER_SEM,
    KIND_DIM: LAYER_SEM,
    KIND_DATUM: LAYER_DATUM,
    KIND_NOTE: LAYER_NOTE,
}

# ta 表列数阈值：col11=Associated Semantic PMI（关联键），col14=Equivalent Unicode String（图形文本）。
# 导出勾选不同会导致列数在 11~14 之间浮动，缺 col14 时图形文本通道为空但不影响关联。
# 加载期交叉校验阈值：tessellated 表 col10 的引用应几乎全部能落地
# （落语义表 / 经 dcr 映射 / 经 datum-2hop）。低于此值说明该列已不是
# `Associated Semantic PMI`，或该表的列序变了。
TA_REF_MIN_HIT = 0.8

# 基准目标的实体名：**同一份 SFA 报告的两个版本写法不同**。
# ftc_10 写 `datum_target`，ctc_02 / ctc_05 / ftc_06 写 `placed_datum_target_feature`。
# 只认一种写法会让整批基准目标「静默断链」—— ta 引用能落到语义表（交叉校验全绿），
# 但后续判据不认这个实体名，结果开发侧全判「多余」、SFA 侧全判「缺失」。
# 用「含 datum 且含 target」的宽松判据，兼容现有两种写法及后续变体。
_DATUM_TARGET_ENTITIES = ("datum_target", "placed_datum_target_feature")


def is_datum_target_entity(entity: str) -> bool:
    e = (entity or "").strip().lower()
    if e in _DATUM_TARGET_ENTITIES:
        return True
    return "datum" in e and "target" in e


# SFA 语义表里实体类型 -> 条目类别
def classify_sfa_entity(entity: str) -> str:
    e = (entity or "").strip()
    if "dimensional_characteristic_representation" in e:
        return KIND_DIM
    if "datum_system" in e:
        return "datum_system"          # 组合项，不独立对应标注
    if is_datum_target_entity(e):
        return KIND_DATUM
    if "tolerance" in e:
        return KIND_GT
    return KIND_GT


def is_known_entity(entity: str) -> bool:
    """该实体名能否归入已知类别。用于加载期体检：新的实体名不该静默落到兜底分支。"""
    e = (entity or "").strip()
    if not e:
        return False
    return any(k in e for k in (
        "dimensional_characteristic_representation",
        "dimensional_size", "dimensional_location",
        "angular_location", "angular_size",
        "datum_system", "datum_feature", "tolerance",
    )) or is_datum_target_entity(e)


# --------------------------------------------------------------------------
# 归一化层
# --------------------------------------------------------------------------
# 符号统一：右值为规范形式
SYMBOL_MAP = {
    "⌀": "D", "∅": "D", "Ø": "D", "ø": "D", "Φ": "D", "φ": "D",
    "Ⓜ": "M", "Ⓛ": "L", "Ⓟ": "P", "Ⓤ": "U", "Ⓢ": "ST",
    "⫽": "∥", "∥": "∥", "//": "∥",
    "⏥": "▱", "▱": "▱",
    "⏊": "⟂", "⟂": "⟂", "⊥": "⟂",
    "⌓": "⌓", "⊿": "⌓",
    "卤": "±", "±": "±",
    "×": "X", "✕": "X",
    "⭩◎": "", "◎": "",
    # SFA 把「统计公差」符号渲染成字体私有区字形 U+F055（同一份报告里
    # 另一种写法是 `<ST>`）。不映射等于丢掉修饰符。
    "\uf055": "ST",
    # 指向性尺寸的指向符号、注释里单独出现的图形符号：非公差符号，忽略
    "↧": "", "⌴": "",
}

_CJK = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")

# CP936 把 0x80 / 0xB4 等单字节映射成了这些字符，而 Python 的 gbk codec 不认，
# 需要手工还原成原始字节，才能把「UTF-8 被按 GBK 解码」的乱码救回来。
_CP936_PATCH = {
    "€": 0x80, "‚": 0x82, "ƒ": 0x83, "„": 0x84, "…": 0x85, "†": 0x86, "‡": 0x87,
    "ˆ": 0x88, "‰": 0x89, "Š": 0x8A, "‹": 0x8B, "Œ": 0x8C, "Ž": 0x8E,
    "‘": 0x91, "’": 0x92, "“": 0x93, "”": 0x94, "•": 0x95, "–": 0x96, "—": 0x97,
    "˜": 0x98, "™": 0x99, "š": 0x9A, "›": 0x9B, "œ": 0x9C, "ž": 0x9E, "Ÿ": 0x9F,
    "´": 0xB4, "¸": 0xB8, "»": 0xBB, "¼": 0xBC, "½": 0xBD, "¾": 0xBE,
}

# 残缺乱码（字节不全，无法程序化还原）的定点映射
_MOJIBAKE_FIXED = {
    "飦": "Ⓢ",   # 统计公差符号
    "鈱": "⌀",   # 单独出现的直径符
}


def _rebuild_bytes(s: str) -> Optional[bytes]:
    """把字符还原为原始 GBK 字节；遇到无法还原的字符返回 None。"""
    buf = bytearray()
    for ch in s:
        try:
            buf += ch.encode("gbk")
            continue
        except UnicodeEncodeError:
            pass
        if ch in _CP936_PATCH:
            buf.append(_CP936_PATCH[ch])
            continue
        return None
    return bytes(buf)


def _try_rebuild(s: str) -> Optional[str]:
    """严格重建：字节还原后必须能被 UTF-8 合法解码，否则判为「不是乱码」。"""
    b = _rebuild_bytes(s)
    if b is None:
        return None
    try:
        return b.decode("utf-8")
    except UnicodeDecodeError:
        return None


def fix_mojibake(s: str) -> str:
    """修复 UTF-8 被按 GBK/CP936 解码产生的乱码（`鈱€` -> `⌀`，`鈱´` -> `⌴`）。

    严格模式：只有还原后的字节构成合法 UTF-8 才替换，避免误伤真实中文
    （`位置度` 的 GBK 字节不是合法 UTF-8，会被原样保留）。
    整串失败时按 ASCII 边界分段重试。
    """
    if not s or not _CJK.search(s):
        return s
    whole = _try_rebuild(s)
    if whole is not None:
        return whole
    out, seg = [], []

    def flush():
        if not seg:
            return
        t = "".join(seg)
        out.append(_try_rebuild(t) or t)
        seg.clear()

    for ch in s:
        if _CJK.search(ch):
            seg.append(ch)
        else:
            flush()
            out.append(ch)
    flush()
    return "".join(out)


def _fix_token(tok: str) -> str:
    """乱码修复：字节重建 + 残缺乱码兜底映射。"""
    out = fix_mojibake(tok)
    for k, v in _MOJIBAKE_FIXED.items():
        if k in out:
            out = out.replace(k, v)
    return out


_NUM = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")


def norm_number(x: str) -> str:
    """数值规范化：去前导零 / 统一小数位 / 消浮点噪声。

    保留负号（区分 `+.003` 与 `-.003`），但忽略正号前缀（`+0.003` == `0.003`）。
    """
    try:
        v = float(x)
    except (TypeError, ValueError):
        return x
    neg = str(x).strip().startswith("-")
    if abs(v - round(v)) < 1e-9:
        iv = abs(int(round(v)))
        return f"-{iv}" if neg else str(iv)
    s = f"{abs(v):.6f}".rstrip("0").rstrip(".")
    return f"-{s}" if neg else s


def normalize(text: str, *, keep_sep: bool = True) -> str:
    """结构化归一化：乱码修复 -> 符号统一 -> 数值规范化 -> 空白压缩。

    保留 `|` 分隔语义（基准序列 `A | B | C`），不再抹掉。
    """
    if text is None:
        return ""
    s = str(text)
    s = _fix_token(s)
    # 不要误删 <ST> / [C] 这类修饰符
    for k, v in SYMBOL_MAP.items():
        if k:
            s = s.replace(k, v)
    # 数值规范化
    s = _NUM.sub(lambda m: norm_number(m.group(0)), s)
    if keep_sep:
        s = re.sub(r"[ \t]+", " ", s)
        s = re.sub(r"\s*\|\s*", " | ", s)
    else:
        s = s.replace("|", " ")
        s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s*\n\s*", " ", s)
    return s.strip()


_FCF_SYM = re.compile(r"[⌓⌖⟂∥▱⏥⌭⌯−]")
# SFA 图形/语义通道把 FCF 行渲染成 `⌖ | ⌀0.8 | D | B | C`（用 | 分隔），
# 尺寸行则是 `2X ⌀5.50 ± 0.2`。同一语义行可能同时含尺寸行与 FCF 行，
# 用 | 作判据比枚举符号稳（新增符号不会漏判）。
_FCF_SEP = re.compile(r"\|")
_MOD_WORDS = {
    "MAXIMUM_MATERIAL": "M", "LEAST_MATERIAL": "L",
    "STATISTICAL_TOLERANCE": "ST", "PROJECTED": "P", "FREE_STATE": "F",
}


_MOD_CHARS = re.compile(r"[ⓂⓁⓅⓊⓈ]")
_MOD_GLYPH = {"Ⓜ": "M", "Ⓛ": "L", "Ⓟ": "P", "Ⓤ": "U", "Ⓢ": "ST"}
_STD_MODS = ("M", "L", "P", "ST", "U", "F")
# 数量前缀可能不在串首（SFA 图形通道形如 `DIM | 4X ⌀.250`），也常用 `×` 而非 `X`
_COUNT_ANY = re.compile(r"(?:^|[\s|])(\d+)\s*[Xx×✕](?![0-9A-Za-z])")


def _letters_of(text: str) -> List[str]:
    return sorted(set(re.findall(r"(?<![A-Za-z])([A-Z])(?![A-Za-z])", text)))


def sem_tokens(text: str, extra_mods: Iterable[str] = ()) -> Dict[str, Any]:
    """抽取语义指纹：数值集合 + 基准字母 + 修饰符 + 数量前缀。

    三条要点（都是为了让两侧表示差异不产生假判定）：
    - 数量前缀（`4X`）从数值与字母集合中整体剔除，单独用 count 字段比较，
      否则 `2X` 的 `X` 会被当成基准字母，报出「开发侧未呈现基准」这种假差异。
    - 修饰符优先从字形（`Ⓜ Ⓛ Ⓟ Ⓤ Ⓢ`）识别：开发侧写 `EⓂ-FⓂ-GⓂ`、SFA 写
      `E Ⓜ-F Ⓜ-G Ⓜ`，字形紧跟字母，纯字母启发式会漏掉修饰符。
    - 基准字母抽取**不做**字母剔除（`L` 既可能是基准字母也可能是 LMC 修饰符），
      修饰符与基准字母的消歧放到 compare_tokens 里对两侧对称处理。
    """
    raw = str(text or "")
    s = normalize(raw, keep_sep=False)
    m_cnt = _COUNT_ANY.search(s)
    count = int(m_cnt.group(1)) if m_cnt else None
    body = _COUNT_ANY.sub(" ", s)
    nums = sorted({norm_number(x) for x in _NUM.findall(body)})

    glyph_mods = {v for c, v in _MOD_GLYPH.items() if c in raw}
    mods = set(glyph_mods)
    if not glyph_mods:
        # 无字形时退回字母启发式（含自由状态 F、投影 P 等以字母书写的情况）
        mods |= {m for m in _STD_MODS
                 if re.search(rf"(?<![A-Za-z]){m}(?![A-Za-z])", s)}
    mods |= set(extra_mods)

    bare = _COUNT_ANY.sub(" ", normalize(_MOD_CHARS.sub(" ", raw), keep_sep=False))
    return {
        "numbers": nums,
        "letters": _letters_of(bare),
        "modifiers": sorted(mods),
        "count": count,
        "norm": s,
    }


def dev_extra_mods(it: "DevItem") -> List[str]:
    """从 detailData 补充修饰符（md 标题里常常没渲染出来）。"""
    out = set()
    raw = str(it.detail.get("modifiers", "") or "")
    for word, tag in _MOD_WORDS.items():
        if word in raw.upper():
            out.add(tag)
    if str(it.detail.get("dimensional note", "") or "").lower() == "statistical":
        out.add("ST")
    return sorted(out)


def sfa_text_for(text: str, kind: str) -> str:
    """从 SFA 语义文本中取用于比对的片段。**返回原文，不做归一化**。

    一条语义行可能把「尺寸行 + FCF 行 + 修饰行」合并在一起（如
    `3X ⌀3.50 ± 0.2\\n⌭ | 0.1`），而这两行是图纸上两个不同的标注。
    GT 只取 FCF 行，判据优先用 `|` 分隔符（SFA 固定渲染方式），
    退回符号表、再退回首行。DIM 类尺寸描述可能折行，合并全部行。

    必须返回原文：修饰符（`Ⓜ`）、直径符（`⌀`）等字形一旦归一化就丢失，
    后续 sem_tokens / symbols_of 只能靠字形判定修饰符与符号，不可提前转换。
    """
    lines = [l for l in (text or "").split("\n") if l.strip()]
    if not lines:
        return ""
    if kind == KIND_GT:
        fcf = [l for l in lines if _FCF_SEP.search(l)]
        if not fcf:
            fcf = [l for l in lines if _FCF_SYM.search(l)]
        return fcf[0].strip() if fcf else lines[0].strip()
    return " ".join(l.strip() for l in lines)


_DT_PAREN = re.compile(r"\([^)]*\)")                            # (point) / (D = 1.)
_DT_SIZE = re.compile(r"\b\d+(?:[.,]\d+)?\s*[xX]\s*\d*")         # 1.25x2
_DT_DIA = re.compile(r"(?<![A-Za-z0-9])[DR]\s*\d+(?:[.,]\d+)?")  # D85 / R5（归一化后的 ⌀85）
_DT_IDENT = re.compile(r"(?<![A-Za-z0-9])([A-Za-z]{1,2})(\d{1,3})(?![0-9])")


def datum_target_ident(text: str, expect: str = "") -> str:
    """从 SFA 的基准目标文本里抽出「基准字母 + 目标序号」标识。

    **不能取首片段** —— 标识的位置不固定，实测四种写法：

    | SFA 文本 | 首片段 | 真实标识 |
    |---|---|---|
    | `A1 (point)` | A1 | A1 |
    | `K1 (area)` | K1 | K1 |
    | `⌀85\\nK1` | ⌀85 | K1 |
    | `1.25x2\\nC1` | 1.25x2 | C1 |

    `(point)` / `(line)` / `(area)` / `(circular curve)` 是基准目标的**形式**
    （ASME Y14.5 的 point/line/area datum target），`⌀85` / `1.25x2` 是目标尺寸，
    都不是标识本身。

    给了 `expect`（开发侧 datum + target id 拼出的期望值）时优先精确找它，
    比对更稳：`expect='K1'` 在 `⌀85\\nK1` 里能命中，而按位置取首片段会拿到尺寸。
    """
    s = normalize(text or "")
    s = _DT_PAREN.sub(" ", s)
    if expect:
        if re.search(rf"(?<![A-Za-z0-9]){re.escape(expect)}(?![0-9])", s):
            return expect
    s = _DT_DIA.sub(" ", _DT_SIZE.sub(" ", s))
    m = _DT_IDENT.search(s)
    return f"{m.group(1)}{m.group(2)}" if m else ""


# `compare_tokens` 判定「开发侧把 SFA 的多条复合公差并成一条」时固定的备注前缀。
# `match_items` 靠它补 `DF_MRG` 缺陷码 —— 两侧共用同一个常量，避免措辞被改后静默失配。
MRG_REMARK = "ID 关联一致；开发侧多出数值"


def compare_tokens(dev: Dict[str, Any], sfa: Dict[str, Any]) -> Tuple[str, str]:
    """返回 (状态, 备注)。

    子集关系一律视为「ID 已确认一致，仅呈现完整度不同」；只有集合冲突才判疑似。

    基准字母比较前，先从两侧**对称地**剔除已识别的修饰符字母：`L` 既可能是
    基准字母也可能是 LMC 修饰符（如标题 `⌖ Ø0 H K L` 里的 L 同时是基准与 Ⓛ），
    不做对称剔除会产出「开发侧未呈现基准」这类假判定。剔除只损失区分度，
    不会制造冲突。
    """
    all_mods = set(dev["modifiers"]) | set(sfa["modifiers"])
    dn, sn = set(dev["numbers"]), set(sfa["numbers"])
    dl, sl = set(dev["letters"]) - all_mods, set(sfa["letters"]) - all_mods
    dm, sm = set(dev["modifiers"]), set(sfa["modifiers"])
    cnt_note = (f"；数量前缀 {dev['count'] or '-'}× vs {sfa['count'] or '-'}×"
                if dev["count"] != sfa["count"] else "")

    if dn == sn and dl == sl and dm == sm:
        if cnt_note:
            return ST_HIT_DIFF, "ID 关联一致" + cnt_note
        return ST_HIT, "ID 关联 + 语义指纹一致"

    if dn.issubset(sn) and dl.issubset(sl) and dm.issubset(sm):
        lack = []
        if sn - dn:
            lack.append("数值")
        if sl - dl:
            lack.append("基准")
        if sm - dm:
            lack.append("修饰符")
        return ST_HIT_DIFF, "ID 关联一致；开发侧未呈现：" + "/".join(lack) + cnt_note

    if sn and sn.issubset(dn) and sl.issubset(dl) and dm.issubset(sm) and len(dn) > len(sn):
        # 开发侧把 SFA 拆开的多条复合公差并成了一条（stc_09 的 handle 9010/9792：
        # SFA 拆成 `0.050 A B C` 与 `0.010 A` 两条，开发侧写成一条）。
        # 口径已定：**以 SFA 的拆分为准**，故这里保持「差异」并点名多出的数值；
        # `match_items` 会据 MRG_REMARK 补一条 `DF_MRG` 缺陷码，让它也进缺陷统计。
        extra = "/".join(str(x) for x in sorted(dn - sn))
        return ST_HIT_DIFF, (
            f"{MRG_REMARK} {extra}"
            "（把 SFA 的多条复合公差合并成一条；以 SFA 拆分口径为准）" + cnt_note)

    return ST_SUSPECT, "ID 关联但语义指纹冲突"


# --------------------------------------------------------------------------
# 开发侧提取缺陷检测层
#
# 归一化层会「修复」编码损坏，以保证内容比对不被表象差异干扰；
# 但修复动作本身必须留痕：乱码、符号丢失、数量前缀不符都是开发侧
# 提取链路的缺陷——这正是本工具要查出的目标，绝不能静默吞掉。
# 缺陷与分析指标解耦：缺陷不改判定、不进 recall/precision 分母，
# 单独成层报表，避免「数据错」和「提取漏」两类问题互相污染。
# --------------------------------------------------------------------------
DF_ENC   = "ENC 编码损坏"
DF_EMPTY = "EMP 标题为空"
DF_CNT   = "CNT 数量前缀不符"
DF_SYM   = "SYM 符号丢失"
DF_NUM   = "NUM 数值缺失"
DF_MAP   = "MAP handle/name 不一致"
DF_DUP   = "DUP handle 重复"
# 唯一的「对照真值类」缺陷码。上面七个都能只看开发侧得出结论（自检口径）；
# MRG 不行 —— 「该不该拆成多条」必须拿 SFA 的拆分当标答才判得出来。
# 由 match_items 在「SFA 拆成多条 + 开发侧数值是超集」时补记，判据见那里的注释。
DF_MRG   = "MRG 复合公差未拆分"

DEFECT_CODES = (DF_ENC, DF_EMPTY, DF_CNT, DF_SYM, DF_NUM, DF_MAP, DF_DUP, DF_MRG)

# 符号归一化别名（仅用于缺陷比对，不改变正文归一化口径）
_SYM_ALIAS = {"⊥": "⟂", "⏊": "⟂", "⊿": "⌓", "⫽": "∥", "⏥": "▱",
              "∅": "⌀", "Ø": "⌀", "Φ": "⌀", "φ": "⌀", "ø": "⌀", "⭩": ""}
# GD&T 符号全集（SFA 语义/图形通道实际用到的形位公差符号）。
# 直线度在 SFA 里渲染为 U+2212 `−`（全报告仅此一处出现，不与负号混淆）。
# `↗`/`⌰`/`○`/`⌒`/`∠` 是 ASME 的 Unicode 形位公差符号表里给出的写法，实测
# SFA 确实这么用（ctc_05 圆跳动 `↗ | 0.035 | A-B`、全跳动 `⌰ | 0.015 | B`、
# 圆度 `○ | 0.002`、ctc_02 线轮廓 `⌒ | 0.25`、ctc_03 倾斜度 `∠ | 0.04 | A`）。
# 不收录会让这些符号「既不被识别、也不被报为未知」，开发侧丢了也查不出来。
# `◎` 是同心度符号，但要与 `⭩◎` 渲染产物区分 —— 见 symbols_of() 里的剥离。
_FCF_CHARS = set("⌓⌖⟂∥▱⏥⌭⌯−↗⌰○⌒∠◎")
_DIA_CHARS = set("⌀")
# SFA 文本化标注区块时用的排版字形：分隔、填充、区块标记，**不是** GD&T 符号。
# 不能收录进 _FCF_CHARS —— 收录会让「缺符号」检查把它们当成必须出现的符号而误报
# （实测 `▽`/`⎹` 各出现 15 次，全在带基准的 FCF 里、与 `[A]` 同行，属版面元素）。
# `⭩` 是复合公差里的私有区字形，`⌮` 是复合公差的重合符号，都不参与符号比对。
# 这里登记只为让 doctor 不再把它们报成「未收录字符」，保留对真新字符的发现能力。
SFA_LAYOUT_GLYPHS = set("▽⎹◁⌮⭩")
_RAD_PREFIX = re.compile(r"(?<![A-Za-z])R\s*(?=[.\d])")
_SR_PREFIX = re.compile(r"(?<![A-Za-z])S\s*(?=[⌀.])")
# 数量标注可能不在串首（SFA 图形文本形如 `DIM | 4X ⌀.250`），也常用 `×` 而非 `X`。
# 这个「只认串首 / 竖线后」的窄版保留给需要区分位置的场景；
# 缺陷比对用的是更宽的 `_COUNT_ANY`（见 dev_defects 里的说明）。
_CNT_PREFIX = re.compile(r"(?:^|\|)\s*(\d+)\s*[Xx×✕](?![0-9A-Za-z])")


def _absnum(tok: str) -> str:
    """取数值的绝对值形式：符号在两侧呈现方式不同（`R.022-.012` vs `0.012-0.022`），
    缺陷层只关心「数值本身在不在」，正负交由匹配层判定。"""
    v = norm_number(tok)
    return v[1:] if v.startswith("-") else v


def _decimals(tok: str) -> int:
    t = str(tok).split(".")[-1] if "." in str(tok) else ""
    return len(re.sub(r"[^0-9]", "", t))


def _num_equivalent(a: str, b: str) -> bool:
    """同一数值的不同精度呈现视为等价（SFA 图形通道会同时给出 `.4375` 与 `.438`）。"""
    try:
        x, y = float(a), float(b)
    except (TypeError, ValueError):
        return a == b
    n = min(_decimals(a), _decimals(b))
    return round(x, n) == round(y, n)


def _tolerance_like(num: str) -> bool:
    """只保留「像公差的数值」：有小数部分且非零。

    整数（`/ ⌀1` 这类定义区域标注）与零值（`-0` / `-.000`，各通道写法不统一）
    都不足以定性为开发缺陷，剔除以免误报。
    """
    try:
        v = float(num)
    except (TypeError, ValueError):
        return False
    return v != 0 and v % 1 != 0


def mojibake_chars(s: str) -> List[str]:
    """列出字符串中判定为乱码产物的字符。

    判定依据（任一命中即算）：
      1. 整串能被严格重建为合法 UTF-8（典型 `鈱€` -> `⌀`）
      2. 字符命中残缺乱码定点映射表（字节不全，无法程序化还原，如 `飦`）
      3. 串内含 CJK 且夹杂 CP936 私有映射字符（`€` / `´` 等，Python gbk codec 不认）
    """
    if not s:
        return []
    if fix_mojibake(s) != s:
        hits = {c for c in s if _CJK.search(c) or c in _CP936_PATCH}
    else:
        has_cjk = bool(_CJK.search(s))
        hits = {c for c in s if c in _MOJIBAKE_FIXED or (has_cjk and c in _CP936_PATCH)}
    return [c for i, c in enumerate(s) if c in hits and c not in s[:i]]   # 去重且保留原顺序


def repaired_text(s: str) -> str:
    """还原乱码后的可读文本（仅用于展示与说明，不改变判定）。"""
    return "".join(_MOJIBAKE_FIXED.get(c, c) for c in fix_mojibake(s))


def symbols_of(text: str) -> set:
    """抽取用于比对的图形符号集合（含直径符 / 半径符 / 球面符）。

    入参必须是**原文**（it_view 不做归一化就是为了这个）：归一化会把 `⌀` 变成 `D`，
    符号就丢了。`⭩◎` 要在这里整对剥离 —— SFA 用私有区字形 `⭩` + `◎` 渲染复合公差，
    其中的 `◎` 与同心度无关，不剥离会在 ftc 系列造成「缺符号 ◎」假报；
    而独立出现的 `◎`（ctc_05 的 `◎ | ⌀0.03 | A`）是真同心度符号，要参与比对。
    """
    s = str(text or "").replace("⭩◎", "")
    s = "".join(_SYM_ALIAS.get(ch, ch) for ch in s)
    out = {ch for ch in s if ch in _FCF_CHARS or ch in _DIA_CHARS}
    if _RAD_PREFIX.search(s):
        out.add("R")
    if _SR_PREFIX.search(s):
        out.add("S")
    return out


def dev_defects(it: "DevItem", view: str = "", *, dup: bool = False,
                dup_count: int = 0, sfa_name: str = "") -> Tuple[List[str], str]:
    """检测单条开发标注的提取缺陷。

    `view` 为 SFA 侧对应条目的可呈现文本（优先图形通道，缺则语义通道）。
    `dup` 只在本条**所属分组内**同一 handle 出现多条时为真 —— 跨分组的重复是
    「同一条 STEP 标注被多个保存视图共用」，属合法结构，另由 MatchRow.multiview 提示。
    返回 (缺陷码列表, 人类可读说明)。缺陷只报不改，不影响匹配判定。
    """
    codes: List[str] = []
    notes: List[str] = []

    raw_title = str(it.title or "")
    fixed = fix_mojibake(raw_title)

    # 1) 编码损坏：修复动作产生差异（或命中残缺乱码映射）即证明原文是坏字节序列
    bad = mojibake_chars(raw_title)
    if bad:
        codes.append(DF_ENC)
        notes.append("标题含乱码 " + "".join(bad) + f"，还原为 {repaired_text(raw_title)!r}")

    # 2) 标题为空
    if not raw_title.strip():
        codes.append(DF_EMPTY)
        notes.append("标注标题为空，无可用文本")

    # 3) 与 SFA 图形通道逐维比对（数量 / 符号 / 数值）
    if view:
        raw_view = str(view)
        # 数量标注的位置两侧不一致：SFA 放串首（`2X ⌀0.238`），开发侧也可能渲染到
        # 串尾（`⌀0.237 +.005 -0.001 2X`）。只认前缀会把「已经正确提取、只是位置不同」
        # 判成「数量前缀 开发=无」——ftc_08 的 3 条假 CNT 就是这么来的。
        # 改用 `_COUNT_ANY`（允许行首 / 空白 / `|` 之后），两侧对称。
        d_cnt = _COUNT_ANY.search(fixed)
        s_cnt = _COUNT_ANY.search(raw_view)
        dv = int(d_cnt.group(1)) if d_cnt else None
        sv = int(s_cnt.group(1)) if s_cnt else None
        if dv != sv:
            codes.append(DF_CNT)
            notes.append(f"数量标注 开发={dv if dv else '无'} SFA={sv if sv else '无'}")

        lost = symbols_of(raw_view) - symbols_of(fixed)
        if lost:
            codes.append(DF_SYM)
            notes.append("缺符号 " + " ".join(sorted(lost)))

        # 数量标注单独由 DF_CNT 负责，这里先从两侧剔除；
        # 只取带小数点的数值（整数常是 `⌀1` 这类定义区域标注），零值也剔除
        # （`-0` / `-.000` 各通道写法不统一，不足以定性为开发缺陷）
        dnums = {_absnum(x) for x in _NUM.findall(_COUNT_ANY.sub("", fixed))}
        snums = {_absnum(x) for x in _NUM.findall(_COUNT_ANY.sub("", raw_view))}
        snums = {v for v in snums if _tolerance_like(v)}
        miss = {v for v in snums if not any(_num_equivalent(v, d) for d in dnums)}
        if miss:
            codes.append(DF_NUM)
            notes.append("缺数值 " + "/".join(sorted(miss)[:6]))

    # 4) 映射一致性
    if sfa_name and it.name and sfa_name != it.name:
        codes.append(DF_MAP)
        notes.append(f"handle 指向 {sfa_name!r}，标注写的是 {it.name!r}")

    # 5) 重复 handle：只认「同一分组内」重复。
    #    跨分组重复（同一 handle 出现在 MBD_A 与 MBD_B）是 SFA 里一条标注挂多个
    #    保存视图的正常结构，开发者按「视图 × 标注」导出必然出多条，不是缺陷。
    if dup and it.handle:
        codes.append(DF_DUP)
        notes.append(f"handle {it.handle} 在同一分组（{it.group or '-'}）内出现 "
                     f"{dup_count or 2} 条")

    return codes, "；".join(notes)


# --------------------------------------------------------------------------
# SFA 真值层
# --------------------------------------------------------------------------
@dataclass
class ColRecipe:
    """一张 SFA 表的列定位结果。

    `mapping` 是「规范化列名 -> 列号」，`source` 说明定位方式：
    `header` = 从表头行读出（可靠）；`fallback` = 未识别表头，退回默认列号（脆弱）。
    """
    key: str = ""                       # 逻辑表名，如 "ta" / "semantic"
    sheet: str = ""                     # 实际工作表名
    header_row: int = -1                # 表头所在行（-1 表示没找到）
    source: str = "header"              # header | fallback
    mapping: Dict[str, int] = field(default_factory=dict)
    resolved: Dict[str, int] = field(default_factory=dict)   # 需要用的字段 -> 列号
    n_cols: int = 0
    n_rows: int = 0                     # 该表 ID 列为数字的数据行数
    n_loaded: int = 0                   # 实际装载条数
    # 多行聚合成一条的表（draughting_model_item_association：一条 callout 多行，
    # 分别指向几何与语义实体）—— 聚合后的组数。装载校验要用它当分母，
    # 否则 46 行 → 20 条会被误判成「列定位错位」。
    n_rows_grouped: int = 0
    # `n_loaded == 0` 时的解释。装载校验直接引用它，避免把「这张表按设计就产出不了
    # 条目」误报成「列定位错位」。
    zero_reason: str = ""
    notes: List[str] = field(default_factory=list)

    @property
    def trustworthy(self) -> bool:
        return self.source == "header"


@dataclass
class IntegrityCheck:
    """加载期交叉校验项。不依赖列语义，只依赖「读到没有 / 数量对不对」。"""
    name: str = ""
    ok: bool = True
    detail: str = ""
    value: str = ""

    def line(self) -> str:
        return f"[{'ok' if self.ok else '!!'}] {self.name}：{self.value}" + (
            f"（{self.detail}）" if self.detail else "")


def _norm_header(v: Any) -> str:
    """表头规范化：取首行（SFA 把 `(Sec. x)` 说明放在第二行）、去括号注释、压空白、小写。"""
    s = str(v or "").replace("\r", "\n")
    s = s.split("\n")[0]
    s = re.sub(r"\(.*?\)", " ", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def _resolve_col(hmap: Dict[str, int], alias: str) -> Optional[int]:
    """在列名映射里解析别名：先精确相等，再前缀匹配（取最靠左的列）。

    前缀匹配是为了兼容 `Similar PMI / Exception` 这类带后缀的表头；
    精确优先是为了避免 `dimension` 误命中 `dimensional`。
    """
    a = alias.lower()
    if a in hmap:
        return hmap[a]
    for h, j in hmap.items():
        if h.startswith(a):
            return j
    return None


def _find_header_row(rows: List[List[str]], aliases: Sequence[str],
                     scan: int = 8) -> Tuple[int, Dict[str, int]]:
    """在前 scan 行里找表头行：取命中所需别名最多的那一行。

    SFA 的表的列名位置并不固定（语义表在 r2，tessellated/dcr 在 r3，
    且 r2 常是 Excel 生成的 `列1..列N` 占位行），所以必须扫描而不是写死行号。
    至少命中 2 个别名才认，避免把占位行误判成表头。
    """
    best_i, best_hit, best_map = -1, 0, {}
    for i, r in enumerate(rows[:scan]):
        m: Dict[str, int] = {}
        for j, cell in enumerate(r):
            h = _norm_header(cell)
            if h and h not in m:
                m[h] = j
        hit = sum(1 for a in aliases if _resolve_col(m, a) is not None)
        if hit > best_hit:
            best_i, best_hit, best_map = i, hit, m
    if best_hit >= 2:
        return best_i, best_map
    return -1, {}


def _locate(rows: List[List[str]], aliases: Sequence[str],
            wants: Sequence[str], defaults: Sequence[int]) -> ColRecipe:
    """定位一张表的列：优先按表头列名，失败才按默认列号回落。

    **表头识别成功时，找不到的列一律记为「不存在」（None），绝不套默认列号。**
    默认列号是「整表没识别出表头」时的最后手段；表头已经在手还去套默认列号，
    等于明知列名不对还按位置硬取。ftc_08 的 `-tg` 报告就踩在这里：
    ta 表只有 12 列、没有 `Associated Semantic PMI`，套默认列号 10 之后
    读到的其实是 `Saved Views`，于是从 `camera_model_d3 58866 (MBD_A)` 里
    抠出 116 个假引用 —— 装载条数、表头识别、通道判定三项全绿，
    只有结果表全红，且原因被伪装成「引用无法解释」。

    例外只有 `id`：SFA 所有表的第一列都是 ID（实测 18 份报告一致），
    而且它只决定「数据行条数」这一个口径、不参与任何内容取值，
    所以缺列表头时仍回落 0，但会在 notes 里留痕。
    """
    hi, hmap = _find_header_row(rows, aliases)
    cr = ColRecipe(header_row=hi, mapping=hmap, n_cols=max((len(r) for r in rows), default=0))
    if hi >= 0:
        for a, d in zip(wants, defaults):
            j = _resolve_col(hmap, a)
            if j is not None:
                cr.resolved[a] = j
                continue
            if a == "id":
                cr.resolved[a] = 0
                cr.notes.append("表头未见 `id`，按第 1 列取（只影响数据行计数，不参与取值）")
            else:
                cr.resolved[a] = None
                cr.notes.append(
                    f"表头未见 `{a}`，该列按**不存在**处理"
                    f"（不套默认列号 {d} —— 位置已变，按位置取会读到别的列）")
    else:
        cr.source = "fallback"
        cr.resolved = {a: d for a, d in zip(wants, defaults)}
        cr.notes.append("未识别到表头行，按默认列号取（结构变化时可能静默失配）")
    return cr


# 各表所需的列别名（与默认列号，仅在识别不到表头时使用）
SEM_ALIASES, SEM_WANTS, SEM_DEFAULT = (
    ("id", "entity", "semantic pmi", "similar pmi"),
    ("id", "entity", "semantic pmi", "similar pmi"),
    (0, 1, 2, 3),
)
DC_ALIASES, DC_WANTS, DC_DEFAULT = (("id", "name"), ("id", "name"), (0, 1))
TA_ALIASES = ("id", "name", "associated semantic pmi", "equivalent unicode string",
              "validation properties")
TA_WANTS, TA_DEFAULT = ("id", "name", "associated semantic pmi",
                        "equivalent unicode string"), (0, 1, 10, 13)
# draughting_model_item_association：**非 tessellated 导出**（如 ctc_05-e1）的图形标注
# 关联表。与 ta 表同角色，但结构不同：
#   identified_item = `draughting_callout 267`  ← 关联键（=开发侧 handle）
#   definition      = `placed_datum_target_feature 1103` / `shape_aspect 1016`
#   used_representation = `draughting_model 99`
# 一条 callout **有多行**：既指向纯几何（shape_aspect），也指向真正的语义实体，
# 必须按 callout 聚合，不能一行一条。
DMIA_ALIASES = ("id", "definition", "used_representation", "identified_item")
DMIA_WANTS, DMIA_DEFAULT = ("definition", "identified_item"), (3, 5)
# 关联通道名（写进 SfaTruth.link_channel，供体检面板与 doctor 显示）
LINK_TA   = "tessellated_annotation_occurrence"
LINK_DMIA = "draughting_model_item_association"
DATUM_ALIASES, DATUM_WANTS, DATUM_DEFAULT = (
    ("id", "identification"), ("id", "identification"), (0, 5))
# datum_feature 表：ID -> 基准字母（列名就叫 `Datum`）。这是 ta 里 `Simple Datum.n`
# 引用落地的唯一通道 —— datum 表的 ID 与它不同源。
DF_ALIASES, DF_WANTS, DF_DEFAULT = (("id", "datum"), ("id", "datum"), (0, 5))
DCR_ALIASES, DCR_WANTS, DCR_DEFAULT = (("id", "dimension"), ("id", "dimension"), (0, 1))
# 无语义文本的实体表（尺寸 / 位置 / 尺寸表示 / 形貌）：只用来登记「这个 ID 真实存在」，
# ta 引用了它们算合法，只是没有语义条目可比 —— 不该被当成「引用无法解释」。
# `shape_aspect` 是 DMIA 通道（draughting_model_item_association）必带的一列几何引用，
# 不登记会让「引用可解释」比例被几何引用拉低到 57%，从而误报断链。
ENTITY_ONLY_TABLES = ("dimensional_size", "dimensional_location",
                      "shape_dimension_representation", "shape_aspect")


@dataclass
class SfaTruth:
    path: str = ""
    units: str = ""
    semantic: Dict[int, Dict[str, Any]] = field(default_factory=dict)   # 语义表 ID -> item
    datum: Dict[int, Dict[str, Any]] = field(default_factory=dict)      # datum.ID -> item
    dc: Dict[int, str] = field(default_factory=dict)                    # draughting_callout.ID -> name
    ta: Dict[str, Dict[str, Any]] = field(default_factory=dict)         # name -> {id, sem_refs, text}
    ta_cols: int = 0                                                    # 关联表实际列数（图形文本通道判定）
    # ta 表存在、但表头里**没有** `Associated Semantic PMI` 列 —— 说明这份报告只导出了
    # 图形标注、没导出语义关联（NIST 的 `-tg` 变体就是这样：AP242 只带 tessellated
    # 图形 PMI、无语义 PMI）。此时代码拿不到任何语义引用，且**不得**按位置硬取别的列。
    graphic_only: bool = False
    # 图形标注关联通道：`tessellated_annotation_occurrence`（AP242 含 tessellated 呈现）
    # 或 `draughting_model_item_association`（非 tessellated 导出，如 ctc_05-e1）。
    # 两者取其一即可；都缺才算断链。
    link_channel: str = ""
    dcr_by_dim: Dict[int, int] = field(default_factory=dict)            # dimensional_*.ID -> dcr.ID
    # datum_feature.ID -> 基准字母。"datum.ID" 与 "datum_feature.ID" 是两套 ID，
    # 两者**不保证相邻**（FTC 系列恰好差 1，CTC 系列实测差 3），所以必须按 ID 精确映射，
    # 不能再用 `datum.ID = ref - 1` 的算术假设。
    datum_feature: Dict[int, str] = field(default_factory=dict)
    known_entities: set = field(default_factory=set)                    # 无语义文本的实体 ID（尺寸/位置），用于引用合法性判定
    sheets: List[str] = field(default_factory=list)
    sentinel_row: Optional[int] = None
    warnings: List[str] = field(default_factory=list)
    recipes: List["ColRecipe"] = field(default_factory=list)            # 每张表的列定位结果
    checks: List["IntegrityCheck"] = field(default_factory=list)        # 加载期交叉校验

    def recipe(self, key: str) -> Optional["ColRecipe"]:
        for r in self.recipes:
            if r.key == key:
                return r
        return None

    def broken_checks(self) -> List["IntegrityCheck"]:
        return [c for c in self.checks if not c.ok]

    @property
    def truth_missing(self) -> bool:
        """SFA 侧是否**完全没有**可比真值。

        判据：语义表、datum、datum_feature 三者皆空。正常报告至少会有语义表；
        三者全空说明这份导出件不含语义 PMI（如 NIST `-tg` 图形专用报告）。
        此时开发侧的每一条都「无从比对」——不是「多余」。
        """
        return not (self.semantic or self.datum or self.datum_feature)


# 引用列（ta 的 `Associated Semantic PMI`、dcr 的 `dimension`）的值形如
# `dimensional_size 120`，要从里面抠出实体 ID。这里**不能写死 ID 位数**：
# NIST FTC 系列实体 ID 是 4~7 位，而 CTC 系列（如 ctc_01）只有 2~3 位，
# 早期写死 `\d{4,7}` 会让整张 dcr 表装载为 0 条、ta 引用全部丢失 ——
# 直接表现为全量失配（0% 命中），且不报任何错。
# 排除 `(` 是为了跳过 `(1) composite_shape_aspect` 这类「第 n 个引用」序号，
# 排除前后的 `.` 是为了不把 `35.0` / `34.8` 这类小数当成 ID。
_REF_ID = re.compile(r"(?<![\d.(])(\d{1,7})(?![\d.])")


def _ref_ids(text: str) -> List[int]:
    return [int(x) for x in _REF_ID.findall(text or "")]


def _rows(ws) -> List[List[str]]:
    out = []
    for r in ws.iter_rows(values_only=True):
        out.append(["" if c is None else str(c).strip() for c in r])
    return out


def _pick_sheet(names: Sequence[str], prefix: str) -> Optional[str]:
    for n in names:
        if n.startswith(prefix):
            return n
    for n in names:
        if prefix in n:
            return n
    return None


def _pick_exact(names: Sequence[str], name: str) -> Optional[str]:
    """精确匹配表名。用于 `datum` 这类前缀会误命中 `datum_system` 的场景。"""
    low = name.strip().lower()
    for n in names:
        if n.strip().lower() == low:
            return n
    return None


# 同目录同案例的兄弟报告：`nist_ftc_08_asme1_ap242-e1-tg-sfa.xlsx` -> `nist_ftc_08`
_SIBLING_KEY = re.compile(r"^(nist_[a-z]+_\d+)", re.I)


def _source_name(src) -> str:
    """取来源的文件名。

    `src` 可以是路径，也可以是**文件对象** —— Streamlit `file_uploader` 给的是
    `UploadedFile`（不是路径！直接喂给 `os.path.basename` 会抛
    `expected str, bytes or os.PathLike object, not UploadedFile`），
    它带 `.name` 属性，值就是上传时的原始文件名。取不到就返回空串。
    """
    try:
        return os.path.basename(os.fspath(src))
    except TypeError:
        return os.path.basename(str(getattr(src, "name", "") or ""))


def _source_abspath(src) -> str:
    """取来源的绝对路径；文件对象没有路径（内容在内存里），返回空串。"""
    try:
        return os.path.abspath(os.fspath(src))
    except TypeError:
        return ""


def _has_semantic_sheet(path: str) -> bool:
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:                   # noqa: BLE001
        return False
    try:
        return any("semantic pmi summary" in s.strip().lower() for s in wb.sheetnames)
    finally:
        wb.close()


def suggest_reports_with_truth(path, limit: int = 4) -> List[str]:
    """找「含语义 PMI」的兄弟报告，供「导错文件」时给出可行替代。

    导错文件是这类 0% 结果的头号原因（图形专用变体没有 `Semantic PMI Summary`），
    而区分两份报告的唯一可靠特征就是这张表在不在。只读表名、不解析内容，代价很低；
    按文件名里的案例号（`nist_ftc_08`）收窄候选，避免把整个目录扫一遍。

    `path` 可以是路径，也可以是文件对象（UI 上传场景）：后者没有所在目录，
    同目录这一步自动跳过，只剩 `PMI_SFA_DIR` 这一路 —— 而导出件常被拖到桌面、
    标准件在 `PMI_SFA_DIR`，所以这条退化路径恰好覆盖最常见的用法。

    搜索顺序：报告同目录 → 环境变量 `PMI_SFA_DIR`（NIST 标准件目录常与工作目录分离）。
    纯提示，不参与任何判定；找不到就返回空列表。
    """
    import glob as _glob

    name = _source_name(path)
    m = _SIBLING_KEY.match(name)
    key = m.group(1).lower() if m else ""
    me = _source_abspath(path)
    dirs: List[str] = []
    if me:
        dirs.append(os.path.dirname(me))
    env_dir = os.environ.get("PMI_SFA_DIR", "")
    if env_dir:
        dirs.append(env_dir)

    out: List[str] = []
    for d in dirs:
        if not (d and os.path.isdir(d)):
            continue
        pat = f"{key}*" if key else "*"
        for full in sorted(_glob.glob(os.path.join(d, pat))):
            low = os.path.basename(full).lower()
            if not low.endswith((".xlsx", ".xlsm")) or low.startswith("~$"):
                continue
            # 排除自己：有路径时比绝对路径，文件对象时比文件名
            if (me and os.path.abspath(full) == me) or (name and low == name.lower()):
                continue
            if os.path.abspath(full) in out:
                continue
            if _has_semantic_sheet(full):
                out.append(full)
                if len(out) >= limit:
                    return out
    return out


def _cell(r: List[str], j: Optional[int]) -> str:
    """安全取列：列号越界或缺失一律返回空串，不抛异常。"""
    if j is None or j < 0 or j >= len(r):
        return ""
    return r[j]


def _open_table(wb, t: SfaTruth, key: str, prefix: str, aliases: Sequence[str],
                wants: Sequence[str], defaults: Sequence[int],
                search: Sequence[str] = (),
                optional: bool = False) -> Tuple[Optional[List[List[str]]], Optional[ColRecipe]]:
    """定位工作表并解析表头列名，返回 (数据行, 列定位结果)。

    `search` 是备用表名前缀（SFA 不同版本里同一张表可能叫不同名字）。
    `optional=True` 表示该表缺失属正常（如只有部分报告才导出的辅助索引表），不记警告。
    """
    sname = _pick_exact(t.sheets, prefix) or _pick_sheet(t.sheets, prefix)
    for alt in search:
        if not sname:
            sname = _pick_sheet(t.sheets, alt)
    if not sname:
        if not optional:
            t.warnings.append(f"未找到 `{prefix}` 表")
        return None, None
    rows = _rows(wb[sname])
    cr = _locate(rows, aliases, wants, defaults)
    cr.key, cr.sheet = key, sname
    id_col = cr.resolved.get("id", 0)
    cr.n_rows = sum(1 for r in rows if _cell(r, id_col).isdigit())
    t.recipes.append(cr)
    return rows, cr


def load_sfa(path) -> SfaTruth:
    """读取 SFA 报告，建立 ID 索引。

    `path` 可以是**文件路径**，也可以是**文件对象**（Streamlit `file_uploader`
    的 `UploadedFile`、`BytesIO`）—— openpyxl 两种都收。差别只在「找同目录兄弟
    报告」这类路径语义：文件对象没有目录，那一路自动跳过，提示改走 `PMI_SFA_DIR`。
    早期版本直接对入参做 `os.path.basename`，UI 一上传就抛
    `expected str, bytes or os.PathLike object, not UploadedFile`。

    列位置一律**按表头列名定位**，不写死列号。SFA 的列数与列序随版本和导出勾选变化
    （实测 tessellated 表 12~14 列、dcr 表 17~20 列），写死列号会造成静默失配 ——
    曾经因此整表被跳过，结果表全判「多余 / 缺失」。
    识别不到表头时才退回默认列号，并在 recipe.notes 与 warnings 里显式说明。
    """
    src_name = _source_name(path)
    t = SfaTruth(path=src_name)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    t.sheets = list(wb.sheetnames)

    def col(cr: Optional[ColRecipe], name: str) -> Optional[int]:
        return cr.resolved.get(name) if cr else None

    # --- 1. Semantic PMI Summary：语义真值
    rows, cr = _open_table(wb, t, "semantic", "Semantic PMI Summary", SEM_ALIASES,
                           SEM_WANTS, SEM_DEFAULT, search=("Semantic PMI Summ",))
    if rows is not None:
        c_id, c_ent = col(cr, "id"), col(cr, "entity")
        c_txt, c_sim = col(cr, "semantic pmi"), col(cr, "similar pmi")
        for i, r in enumerate(rows):
            if _cell(r, c_txt) == "Expected PMI":
                t.sentinel_row = i + 1
                break
        for r in rows:
            rid, txt = _cell(r, c_id), _cell(r, c_txt)
            if rid.isdigit() and txt not in ("", "Expected PMI"):
                t.semantic[int(rid)] = {
                    "id": int(rid), "entity": _cell(r, c_ent), "text": txt,
                    "similar": _cell(r, c_sim),
                    "kind": classify_sfa_entity(_cell(r, c_ent)),
                }
        if cr:
            cr.n_loaded = len(t.semantic)

    # --- 2. draughting_callout：handle <-> name
    rows, cr = _open_table(wb, t, "dc", "draughting_callout", DC_ALIASES,
                           DC_WANTS, DC_DEFAULT)
    if rows is not None:
        c_id, c_nm = col(cr, "id"), col(cr, "name")
        for r in rows:
            rid = _cell(r, c_id)
            if rid.isdigit():
                t.dc[int(rid)] = _cell(r, c_nm)
        if cr:
            cr.n_loaded = len(t.dc)

    # --- 3. 图形标注关联通道：name -> 语义引用（+ 图形文本）
    # 同一份语义数据有两种导出形态，取其一即可：
    #   a) tessellated_annotation_occurrence（含 tessellated 呈现）
    #      一条标注一行，`Associated Semantic PMI` 直接给语义引用
    #   b) draughting_model_item_association（**非 tessellated 导出**，如 ctc_05-e1）
    #      一条 callout 多行，`identified_item` = draughting_callout N，
    #      `definition` = 关联实体（同时含 shape_aspect 这类纯几何引用）
    # 只认 (a) 会让 (b) 的报告整条通道断掉：t.ta 全空 → 开发侧全判「多余」、
    # SFA 侧全判「缺失」，指标 0%，而装载 / 引用 / name 三项交叉校验**全绿**。
    # 实测 18 份 NIST 报告里 17 份有 (a)、ctc_05-e1 只有 (b)。
    #
    # 另有第三种形态（ftc_08 `-e1-tg`）：ta 表在，但**表头里没有**
    # `Associated Semantic PMI` 列 —— 该报告只导出了图形 PMI、没导出语义关联。
    # 这种表必须当「通道不可用」处理并继续试 (b)，绝不能按默认列号去读别的列
    # （默认列号 10 落在 `Saved Views` 上，会抠出一堆相机/视图 ID 冒充语义引用）。
    rows, cr = _open_table(wb, t, "ta", "tessellated_annotation_occurren", TA_ALIASES,
                           TA_WANTS, TA_DEFAULT, optional=True)
    c_ref = col(cr, "associated semantic pmi") if rows is not None else None
    ta_keyless = rows is not None and c_ref is None
    if rows is not None:
        c_id, c_nm = col(cr, "id"), col(cr, "name")
        c_uni = col(cr, "equivalent unicode string")
        t.ta_cols = cr.n_cols if cr else 0
        for r in rows:
            rid, nm = _cell(r, c_id), _cell(r, c_nm)
            if rid.isdigit() and nm:
                raw = _cell(r, c_ref)
                t.ta[nm] = {
                    "id": int(rid), "ta_name": nm,
                    "sem_refs": _ref_ids(raw), "sem_ref_raw": raw,
                    "text": _cell(r, c_uni),
                }
        if cr:
            cr.n_loaded = len(t.ta)
        if not ta_keyless:
            t.link_channel = LINK_TA
            if t.ta and not any(v.get("text") for v in t.ta.values()):
                t.warnings.append(
                    "tessellated_annotation_occurrence 未提供 `Equivalent Unicode String(s)`"
                    "（图形文本通道），「SFA图形文本」列将为空，内容比对已回落到语义表文本。"
                    "如需按图纸上实际呈现的字形比对，请在 SFA 导出时勾选该列。")
        else:
            # 表在、行也装载了，但**没有关联键列** —— 该表提供不了任何语义引用。
            # 保留 t.ta（条目名对诊断仍有价值：可以看出两侧命名体系完全不同），
            # 但绝不把别的列当关联键读。
            if cr:
                cr.notes.append("该表没有 `Associated Semantic PMI` 列，不提供任何语义关联")
            t.warnings.append(
                f"`{cr.sheet}` 表头里**没有** `Associated Semantic PMI` 列"
                f"（本表共 {t.ta_cols} 列，表头只有 "
                + "、".join(f"`{h}`" for h in list(cr.mapping)[:12]) + "）。"
                "这份导出件只含图形 PMI、不含语义关联 —— 标注 ↔ 语义实体无法建立，"
                "内容比对不成立。通常是 SFA 导出时只勾了图形 PMI，"
                "或 STEP 文件本身就是「图形专用」变体（NIST 命名里带 `-tg`）。")

    if rows is None or ta_keyless:
        rows, cr = _open_table(wb, t, "ta", "draughting_model_item_associati",
                               DMIA_ALIASES, DMIA_WANTS, DMIA_DEFAULT, optional=True)
        if rows is None:
            if not ta_keyless:
                t.warnings.append(
                    "两种图形标注关联表都没有："
                    "`tessellated_annotation_occurrence` 与 `draughting_model_item_association`，"
                    "标注 ↔ 语义实体无法建立关联（比对结果必然全为多余 / 缺失）")
        else:
            c_def, c_item = col(cr, "definition"), col(cr, "identified_item")
            t.ta_cols = cr.n_cols if cr else 0
            agg: Dict[int, Dict[str, List[str]]] = {}
            for r in rows:
                items_ref = _ref_ids(_cell(r, c_item))
                if not items_ref:
                    continue
                rec = agg.setdefault(items_ref[0], {"refs": [], "raw": []})
                raw = _cell(r, c_def)
                for x in _ref_ids(raw):
                    if x not in rec["refs"]:
                        rec["refs"].append(x)
                if raw and raw not in rec["raw"]:
                    rec["raw"].append(raw)
            added: Dict[str, Dict[str, Any]] = {}
            for hid, rec in agg.items():
                nm = t.dc.get(hid)
                if not nm:
                    continue                      # callout 未被 dc 表登记，跳过
                added[nm] = {
                    "id": hid, "ta_name": nm,
                    "sem_refs": rec["refs"], "sem_ref_raw": " ; ".join(rec["raw"]),
                    "text": "",                   # 该形态不带图形文本列
                }
            t.ta.update(added)
            if cr:
                # 只算本表贡献的条数：ta 表（若有）已装载的不能算进这张表的装载率
                cr.n_loaded = len(added)
                cr.n_rows_grouped = len(agg)      # 装载校验的分母：分组合数，不是行数
                if not added:
                    cr.zero_reason = (
                        "该报告没有 `draughting_callout` 表，关联表的 `identified_item` "
                        "指向的是 tessellated 标注本身、而不是 callout，"
                        "本表按设计产出不了关联条目（不是列定位错位）"
                        if not t.dc else
                        "关联表里的 `identified_item` 与 `draughting_callout` 表完全对不上，"
                        "两张表可能不是同一份数据")
            if added:
                t.link_channel = LINK_DMIA
                t.warnings.append(
                    "该报告没有可用的 `tessellated_annotation_occurrence` 关联列，已改用 "
                    "`draughting_model_item_association` 建立关联（等价通道，非错误）。"
                    "代价：缺 `Equivalent Unicode String(s)` 图形文本列，"
                    "内容比对回落到语义表文本，「SFA图形文本」列为空。")
            else:
                t.warnings.append(
                    "`draughting_model_item_association` 里的 `identified_item` 指向的 "
                    + ("对象不是 `draughting_callout`（该报告没有这张表，关联表指向的是 "
                       "tessellated 标注本身）" if not t.dc
                       else "callout 未被 `draughting_callout` 表登记")
                    + "，关联仍无法建立")

    # 只有「ta 缺关联键列、且没有别的通道顶上」才算图形专用导出
    t.graphic_only = ta_keyless and not t.link_channel

    # --- 4. datum：基准标识
    rows, cr = _open_table(wb, t, "datum", "datum", DATUM_ALIASES,
                           DATUM_WANTS, DATUM_DEFAULT)
    if rows is not None and cr is not None:
        if "system" in cr.sheet.lower():
            t.warnings.append(f"`datum` 表命中 `{cr.sheet}`（疑似 datum_system），已跳过该表")
            t.recipes.remove(cr)
        else:
            c_id, c_ident = col(cr, "id"), col(cr, "identification")
            for r in rows:
                rid = _cell(r, c_id)
                if rid.isdigit():
                    t.datum[int(rid)] = {"id": int(rid), "identification": _cell(r, c_ident)}
            cr.n_loaded = len(t.datum)

    # --- 4b. datum_feature：基准字母。ta 里 `Simple Datum.n` 引用的是本表 ID，
    # 而 datum 表的 ID 是另一套（两者不相邻），所以这一跳不能靠算术猜。
    rows, cr = _open_table(wb, t, "datum_feature", "datum_feature", DF_ALIASES,
                           DF_WANTS, DF_DEFAULT, optional=True)
    if rows is not None:
        c_id, c_ident = col(cr, "id"), col(cr, "datum")
        for r in rows:
            rid = _cell(r, c_id)
            if rid.isdigit():
                t.datum_feature[int(rid)] = _cell(r, c_ident)
        if cr:
            cr.n_loaded = len(t.datum_feature)

    # --- 5. dimensional_characteristic_representation：尺寸的第 2 跳
    rows, cr = _open_table(wb, t, "dcr", "dimensional_characteristic_repr", DCR_ALIASES,
                           DCR_WANTS, DCR_DEFAULT)
    if rows is not None:
        c_id, c_dim = col(cr, "id"), col(cr, "dimension")
        for r in rows:
            rid = _cell(r, c_id)
            if rid.isdigit():
                for ref in _ref_ids(_cell(r, c_dim)):
                    t.dcr_by_dim[ref] = int(rid)
        if cr:
            cr.n_loaded = len(t.dcr_by_dim)

    # --- 5b. 登记「只有实体、没有语义文本」的 ID（尺寸 / 位置 / 尺寸表示）。
    # 这些表不进 recipes（不做装载校验，它们本来就不参与比对），只把 ID 收进白名单，
    # 供「ta 引用能否落地」判定使用。
    for prefix in ENTITY_ONLY_TABLES:
        sname = _pick_exact(t.sheets, prefix) or _pick_sheet(t.sheets, prefix)
        if not sname:
            continue
        for r in _rows(wb[sname]):
            rid = _cell(r, 0)
            if rid.isdigit():
                t.known_entities.add(int(rid))

    # --- 6. 单位
    hname = _pick_sheet(t.sheets, "Header")
    if hname:
        for r in _rows(wb[hname]):
            m = re.search(r"\b(INCH|MM|MILLIMETRE|METER)\b", " ".join(r), re.I)
            if m:
                t.units = m.group(1).upper()
                break
    wb.close()

    # --- 7. 无真值兜底提示：报告里一条语义 PMI 都没有时，先把话说清楚。
    # 这是「导错文件」的头号表征 —— 用户手上往往同时有 `-e1-tg`（图形专用）与
    # `-e2`（含语义）两份导出，随手拖了一份进来就会看到满屏 0%。
    # 只做提示（不改判定），因为「哪一份才是对的」只有用户知道。
    if t.truth_missing:
        alt = suggest_reports_with_truth(path)
        tip = ("该 SFA 报告不含任何语义 PMI：既没有 `Semantic PMI Summary`，"
               "也没有 `draughting_callout` / `datum` 可作真值。"
               "开发侧的条目**无从比对**（不是提取错误），内容级比对不成立。")
        if src_name:
            tip += f"（当前文件：`{src_name}`）"
        if alt:
            tip += ("这些报告含 `Semantic PMI Summary`，应该才是你要比对的那份："
                    + "；".join(os.path.basename(x) for x in alt)
                    + "。若手上还有同一测试件的其它导出（例如不带 `-tg` 的版本），"
                      "优先换用那一份。")
        else:
            tip += ("请换用同一测试件的**语义版**导出（SFA 导出时勾选语义 PMI；"
                    "NIST 命名里带 `-tg` 的是图形专用变体，不能用）。")
        t.warnings.append(tip)

    _integrity_checks(t)
    return t


def _integrity_checks(t: SfaTruth) -> None:
    """加载后交叉校验：只看「读到没有、数量对不对、引用能否落地」，不依赖列语义。

    价值在于把「静默失配」变成显式告警：结构变了导致整表被跳过时，
    装载条数会与表内数据行数严重不符，比结果表全红更早暴露问题。
    """
    def add(name: str, ok: bool, value: str, detail: str = "") -> None:
        t.checks.append(IntegrityCheck(name=name, ok=ok, value=value, detail=detail))

    for cr in t.recipes:
        if cr.trustworthy:
            add(f"{cr.sheet} 列名识别", True, f"表头第 {cr.header_row + 1} 行")
        else:
            add(f"{cr.sheet} 列名识别", False, "退回默认列号，结构变化时会静默失配")

    for cr in t.recipes:
        # 聚合表（DMIA）的「应有条数」是分组合数，不是表内行数
        n_rows = cr.n_rows_grouped or cr.n_rows
        if n_rows == 0:
            add(f"{cr.sheet} 装载", False, "表内无可识别的数据行")
        elif cr.n_loaded == 0:
            add(f"{cr.sheet} 装载", False, f"表内 {n_rows} 行、装载 0 条",
                cr.zero_reason or "整表被跳过：列定位与代码预期不符")
        elif cr.n_loaded < n_rows * 0.5:
            add(f"{cr.sheet} 装载", False,
                f"表内 {n_rows} 行、仅装载 {cr.n_loaded} 条", "疑似列定位错位")
        else:
            extra = f"，聚合为 {cr.n_loaded} 组" if cr.n_rows_grouped else ""
            add(f"{cr.sheet} 装载", True,
                f"{cr.n_loaded}/{n_rows} 条{extra}")

    # 关联通道本身要先确认存在 —— 两种表都缺时，下面所有关联统计都是空的
    # 「假绿」，必须显式报出来。这是 ctc_05-e1 暴露的：报告只有
    # draughting_model_item_association，而代码只认 tessellated 表，
    # 于是 t.ta 全空、指标 0%，但其余校验项全部通过。
    add("图形标注关联通道", bool(t.link_channel),
        t.link_channel or "两种关联表都缺失",
        "`tessellated_annotation_occurrence` 与 `draughting_model_item_association` "
        "都没找到，标注 ↔ 语义实体无法建立关联"
        if not t.link_channel else "")

    # 关联表在、但**关联键列**不在：通道名看着有，实际一条引用也取不到。
    # ftc_08 `-e1-tg` 就是这样：表头识别通过、通道判定通过、装载 58 条，
    # 唯独没有 `Associated Semantic PMI` 列，过去被静默套上默认列号读 `Saved Views`。
    if t.graphic_only:
        add("图形标注关联键列", False,
            f"{t.link_channel or 'tessellated_annotation_occurrence'} 表缺 "
            f"`Associated Semantic PMI` 列",
            "列名已识别到位，说明该列确实不存在（不是位置变化）。"
            "这份导出件只有图形 PMI、没有语义关联，标注 ↔ 语义实体建立不起来")

    # 语义真值：没有语义表就谈不上「语义比对」。以前缺表时不留任何痕迹，
    # 开发侧每一条都被判「多余」，精确率 0% 却看不出是「没得比」还是「比错了」。
    if t.truth_missing:
        add("语义真值", False, "报告不含语义 PMI（无语义表 / datum / datum_feature）",
            "该 SFA 报告是图形专用导出，SFA 侧没有可比真值；"
            "开发侧条目标记为「⛔ 无可比真值」而非「多余」，指标不成立")

    if t.ta:
        # 引用分五类可落地：直接落语义表 / 经 dcr 映射（尺寸）/ datum_feature（基准）/
        # datum 的 ID-1 邻接（旧版报告）/ 无语义文本的实体表（尺寸·位置）。
        # 只看「全部解释不了」的比例 —— 拿「能落进语义表」当分子会把正常的
        # 尺寸类引用算成异常（实测仅 54%），阈值形同虚设。
        total = 0
        explained = 0
        for v in t.ta.values():
            for r in v["sem_refs"]:
                total += 1
                if (r in t.semantic or r in t.dcr_by_dim or r in t.datum_feature
                        or r in t.known_entities or (r - 1) in t.datum):
                    explained += 1
        if total == 0:
            # 有关联表却一条引用都没有：分子分母都为 0 时旧代码算作 100% 通过，
            # 又是一处「假绿」。正常报告里 ta 表必然带引用。
            add("ta 引用可解释", False, "0 条引用",
                "关联表装载了条目却取不到任何引用，说明关联键列缺失或为空，"
                "关联链实际上是断的")
        else:
            ratio = explained / total
            add("ta 引用可解释", ratio >= TA_REF_MIN_HIT,
                f"{explained}/{total} = {ratio:.0%}",
                "有引用既不在语义表、也不在 dcr/datum 索引里，说明该列已不是 "
                "Associated Semantic PMI，或语义表 ID 体系不同"
                if ratio < TA_REF_MIN_HIT else "")

    if t.dc and t.ta:
        dcn = {v for v in t.dc.values() if v}
        common = len(dcn & set(t.ta))
        add("dc 与 ta 的 name 对应", common >= len(dcn) * 0.5,
            f"{common}/{len(dcn)}",
            "对应率过低说明两张表不是同一套标注"
            if common < len(dcn) * 0.5 else "")

    if t.semantic:
        # 语义表实体名漂移的显式告警。这类问题的表征是「引用完全能落地、
        # 但判据不认实体名」，装载与引用检查全绿，只有结果表全红 ——
        # ctc_02 的基准目标就因为 SFA 把 entity 写成 placed_datum_target_feature
        # 而整批被判「多余 / 缺失」。
        unknown: Dict[str, int] = {}
        for v in t.semantic.values():
            e = (v.get("entity") or "").strip()
            if e and not is_known_entity(e):
                unknown[e] = unknown.get(e, 0) + 1
        add("语义表实体归类", not unknown,
            "全部已知" if not unknown else f"{len(unknown)} 种未识别",
            "；".join(f"{k!r} ×{n}" for k, n in sorted(unknown.items()))
            + " —— 未识别的实体会落到几何公差兜底分支，可能造成整批判错"
            if unknown else "")


# --------------------------------------------------------------------------
# 内部项目 markdown 解析层
# --------------------------------------------------------------------------
@dataclass
class DevItem:
    group: str = ""
    seq: int = 0
    title: str = ""
    handle: Optional[int] = None
    name: str = ""
    original_type_name: str = ""
    types: List[str] = field(default_factory=list)
    detail: Dict[str, Any] = field(default_factory=dict)
    raw: str = ""

    @property
    def kind(self) -> str:
        ts = self.types
        if any("geometric_tolerance" in x or x.endswith("_tolerance") for x in ts):
            return KIND_GT
        if any(x in ("dimensional_size", "dimensional_location",
                     "angular_location", "angular_size",
                     "dimensional_characteristic_representation") for x in ts):
            return KIND_DIM
        if any("datum_feature" in x or x == "datum" for x in ts) and "size" not in "".join(ts):
            if self.detail.get("datum"):
                return KIND_DATUM
        if "property_definition" in ts or self.original_type_name == "draughting_callout" and self.detail.get("property") == "semantic text":
            return KIND_NOTE
        if self.detail.get("datum"):
            return KIND_DATUM
        return KIND_NOTE

    @property
    def title_norm(self) -> str:
        return normalize(self.title)

    @property
    def sfa_side_key(self) -> str:
        """SFA 侧用于比对的文本：优先 detailData 中的结构化描述。"""
        return normalize(self.title)


_H2 = re.compile(r"^##\s+(.+?)\s*$")
_H3 = re.compile(r"^###\s+(\d+)\s*[.、]\s*(.*?)\s*$")
_H3_LOOSE = re.compile(r"^###\s+(.*?)\s*$")
_ANY_HEADING = re.compile(r"^#{1,6}\s+\S")
_META = re.compile(r"^-\s*(分组\s*Handle|标注数量|标注总数|视图分组)\s*[:：]\s*(.*)$")
# detailData 标签的宽松匹配：允许引号、大小写、全角冒号、JSON 与标签同行
_DETAIL_LABEL = re.compile(r'^[-\s*]*"?detailData"?\s*[:：]\s*(\{.*)?$', re.IGNORECASE)


@dataclass
class ParseDiag:
    """markdown 解析过程的自检结果。

    存在意义：解析失败时原先只会表现为「全部 ⚠️ 多余」，看不到根因。
    这里把「为什么没解析出来」显式暴露出来。
    """

    total_lines: int = 0
    groups: List[str] = field(default_factory=list)
    headings_matched: int = 0
    headings_unmatched: List[str] = field(default_factory=list)   # 「像标题但没匹配上」的行
    items: int = 0
    with_handle: int = 0
    with_name: int = 0
    detail_missing: List[str] = field(default_factory=list)       # 没找到 detailData 块的条目
    json_failed: List[str] = field(default_factory=list)          # JSON 解析失败的条目
    key_missing: List[str] = field(default_factory=list)          # JSON 在但缺 handle/name 的条目
    h3_pattern: str = r"^###\s+(\d+)\s*[.、]\s*标题$"

    @property
    def healthy(self) -> bool:
        return (self.items > 0 and self.with_handle == self.items
                and self.with_name == self.items and not self.detail_missing
                and not self.json_failed and not self.key_missing)

    def problems(self) -> List[str]:
        """返回面向用户的问题描述（空列表 = 无问题）。"""
        p: List[str] = []
        if self.items == 0:
            if self.headings_unmatched:
                p.append(f"识别到 {len(self.headings_unmatched)} 行像标题的内容，但都不符合 "
                         f"`{self.h3_pattern}`，因此没有解析出任何标注条目。")
            else:
                p.append("没有解析出任何标注条目：未找到 `### 序号. 标题` 形式的标题行。")
        if self.detail_missing:
            p.append(f"{len(self.detail_missing)} 条标注没有解析到 `detailData:` JSON 块"
                     f"（该块必须独占一行且用半角冒号），这些条目拿不到 handle，"
                     f"会全部被判为「⚠️ 多余」。")
        if self.json_failed:
            p.append(f"{len(self.json_failed)} 条的 detailData JSON 解析失败，handle 取不到。")
        if self.key_missing:
            p.append(f"{len(self.key_missing)} 条的 JSON 里缺少 `handle` 或 `name` 字段"
                     f"（字段名须为小写）。")
        if self.items and self.with_handle < self.items and not (self.detail_missing or self.json_failed):
            p.append(f"{self.items - self.with_handle} 条没有取到 handle，无法参与 ID 关联。")
        return p

    def summary(self) -> str:
        return (f"解析条目 {self.items} / 带 handle {self.with_handle} / "
                f"带 name {self.with_name} / 未解析块 {len(self.detail_missing)} / "
                f"JSON 失败 {len(self.json_failed)}")


def _lget(d: Dict[str, Any], key: str) -> Any:
    """大小写不敏感地取 JSON 字段。"""
    if key in d:
        return d[key]
    for k, v in d.items():
        if str(k).lower() == key:
            return v
    return None


def _read_detail_json(lines: Sequence[str], i: int) -> Tuple[str, int, str]:
    """从 `detailData:` 标签之后开始读 JSON 块。

    返回 (blob, 下一行下标, 错误说明)。容忍 ``` 围栏、空行、标签与 `{` 同行。
    """
    n = len(lines)
    buf: List[str] = []
    depth, started = 0, False
    while i < n:
        s = lines[i].strip()
        if not started:
            if not s or s.startswith("```"):
                i += 1
                continue
            if "{" not in s:
                return "", i, "`detailData:` 之后没有找到 JSON 起始 `{`"
            started = True
        buf.append(lines[i])
        depth += lines[i].count("{") - lines[i].count("}")
        i += 1
        if depth <= 0:
            break
    if not started:
        return "", i, "`detailData:` 块为空"
    blob = re.sub(r"^[^{]*", "", "\n".join(buf), count=1).strip()
    blob = re.sub(r"`+\s*$", "", blob).strip()
    if depth > 0:
        return blob, i, "JSON 大括号未闭合（内容可能被截断）"
    return blob, i, ""


def parse_dev_markdown_ex(text: str) -> Tuple[List[DevItem], ParseDiag]:
    """解析内部项目导出的 markdown，并返回解析过程的自检信息。"""
    diag = ParseDiag()
    items: List[DevItem] = []
    group, meta = "", {}
    lines = text.splitlines()
    diag.total_lines = len(lines)
    i, cur = 0, None

    def _flush():
        if cur is not None:
            items.append(cur)

    while i < len(lines):
        ln = lines[i]
        m2 = _H2.match(ln)
        m3 = _H3.match(ln)
        mm = _META.match(ln.strip())
        mdl = _DETAIL_LABEL.match(ln.strip())

        if m2:
            group = m2.group(1).strip()
            if group not in diag.groups:
                diag.groups.append(group)
            i += 1
            continue
        if mm:
            meta[mm.group(1)] = mm.group(2)
            i += 1
            continue
        if m3:
            _flush()
            cur = DevItem(group=group, seq=int(m3.group(1)), title=m3.group(2), raw=ln)
            diag.headings_matched += 1
            i += 1
            continue
        if mdl and cur is not None:
            cur._detail_seen = True          # type: ignore[attr-defined]
            inline = (mdl.group(1) or "").strip()
            if inline:
                # 标签与 `{` 同行：把这段接回读取器
                lines[i] = inline
                blob, i, err = _read_detail_json(lines, i)
            else:
                blob, i, err = _read_detail_json(lines, i + 1)
            cur.detail = {}
            if blob:
                try:
                    d = json.loads(blob)
                    if isinstance(d, dict):
                        cur.detail = d
                    else:
                        err = err or "detailData 不是 JSON 对象"
                except Exception as e:          # noqa: BLE001
                    err = err or f"JSON 解析失败：{e}"
            cur._detail_err = err             # type: ignore[attr-defined]
            continue
        if _ANY_HEADING.match(ln) and not m2 and not m3 and not mdl:
            if len(diag.headings_unmatched) < 10:
                diag.headings_unmatched.append(ln.strip()[:60])
        i += 1

    _flush()
    diag.items = len(items)

    # 统一按文件顺序汇总每条条目的解析结果，避免诊断信息错序
    for it in items:
        it.raw = f"{it.group}#{it.seq}"
        label = f"{it.group}#{it.seq} {it.title[:20]}"
        if not getattr(it, "_detail_seen", False):
            diag.detail_missing.append(f"{label} —— 未找到 `detailData:` 块")
        elif getattr(it, "_detail_err", ""):
            err = it._detail_err                       # type: ignore[attr-defined]
            (diag.json_failed if "JSON" in err else diag.detail_missing).append(
                f"{label} —— {err}")
        it.handle = _to_int(_lget(it.detail, "handle"))
        it.name = str(_lget(it.detail, "name") or "")
        it.original_type_name = str(_lget(it.detail, "original_type_name") or "")
        it.types = [x for x in str(_lget(it.detail, "type") or "").split("+") if x]
        if it.handle is not None:
            diag.with_handle += 1
        if it.name:
            diag.with_name += 1
        if it.detail and it.handle is None and it.name == "":
            diag.key_missing.append(label)
    return items, diag


def parse_dev_markdown(text: str) -> List[DevItem]:
    """解析内部项目导出的 markdown（只取条目，丢弃自检信息）。"""
    return parse_dev_markdown_ex(text)[0]


def _to_int(v) -> Optional[int]:
    if v is None:
        return None
    s = str(v).strip()
    m = re.search(r"\d+", s)
    return int(m.group(0)) if m else None


# --------------------------------------------------------------------------
# 关联 + 比对层
# --------------------------------------------------------------------------
@dataclass
class MatchRow:
    key: str = ""
    layer: str = ""
    kind: str = ""
    group: str = ""
    dev_title: str = ""
    dev_name: str = ""
    handle: Optional[int] = None
    sfa_sem_id: Optional[int] = None
    sfa_text: str = ""
    sfa_entity: str = ""
    sfa_graphic_text: str = ""
    status: str = ""
    remark: str = ""
    path: str = ""        # 关联路径：gt-1hop / dim-2hop / datum-2hop / none
    defects: List[str] = field(default_factory=list)   # 开发侧提取缺陷码
    defect_detail: str = ""
    # 该 handle 出现的分组列表（>1 个即「同一标注被多个保存视图共用」）。
    # 这是 STEP 的合法结构，只做提示：不进缺陷层、不影响任何指标。
    multiview: str = ""

    def as_dict(self) -> Dict[str, Any]:
        return {
            "关联键": self.key,
            "分组": self.group,
            "开发标注": self.dev_title,
            "开发名称": self.dev_name,
            "Handle": self.handle if self.handle is not None else "",
            "SFA语义ID": self.sfa_sem_id if self.sfa_sem_id is not None else "",
            "SFA实体类型": self.sfa_entity,
            "SFA语义文本": self.sfa_text,
            "SFA图形文本": self.sfa_graphic_text,
            "层级": self.layer,
            "类别": self.kind,
            "匹配状态": self.status,
            "提取缺陷": " ".join(self.defects),
            "缺陷详情": self.defect_detail,
            "关联路径": self.path,
            "多视图": self.multiview,
            "人工校验": "待定",
            "备注": self.remark,
        }


def _lookup_semantic_for_dev(it: DevItem, t: SfaTruth) -> Tuple[List[int], str]:
    """返回 (命中的 ID 列表, 关联路径)。

    路径语义：
      `gt-1hop`            ta 直接引用语义表 ID
      `dim-2hop`           ta -> dimensional_*.ID -> dcr -> 语义表（尺寸值挂在 dcr 上）
      `datum_feature-1hop` ta -> datum_feature.ID -> 基准字母
      `datum-2hop`         ta -> ID-1 -> datum 表（旧报告里两者恰好相邻时的兜底）
      `none`               没接上任何语义实体
    """
    ta = t.ta.get(it.name)
    if not ta:
        return [], "none"
    refs = list(ta["sem_refs"])

    if it.kind == KIND_GT:
        ids = [r for r in refs if r in t.semantic]
        return ids, "gt-1hop" if ids else "none"

    if it.kind == KIND_DIM:
        # ta 引用的是 dimensional_size / dimensional_location / angular_location 的 ID，
        # 尺寸值挂在 dcr 上，须经 dcr 精确映射。
        # **不做 `r+1` 兜底**：ctc_01 里 dimensional_size 127 没有 dcr 条目，
        # 一旦按 r+1 试探就会命中 128 的 dcr，把 Linear Size.6 错配成 Linear Size.9 的内容。
        ids = [i for i in (t.dcr_by_dim.get(r) for r in refs) if i and i in t.semantic]
        if ids:
            return ids, "dim-2hop"
        ids = [r for r in refs if r in t.semantic]
        return ids, ("dim-2hop" if ids else "none")

    if it.kind == KIND_DATUM:
        # 基准目标（datum_target / placed_datum_target_feature）：col11 直接给出实体 ID，1 跳
        for r in refs:
            if is_datum_target_entity(t.semantic.get(r, {}).get("entity", "")):
                return [r], "datum_target-1hop"
        # 普通基准：col11 是 datum_feature 的 ID，按 ID 精确查。
        # datum 表与 datum_feature 表是**两套 ID**，不保证相邻（FTC 差 1、CTC 差 3），
        # 所以 `ID-1` 只是最后兜底，不能当主路径。
        for r in refs:
            if r in t.datum_feature:
                return [r], "datum_feature-1hop"
        for r in refs:
            if (r - 1) in t.datum:
                return [r - 1], "datum-2hop"
        return [], "none"

    return [], "none"


def it_view(it: DevItem, t: SfaTruth) -> str:
    """SFA 侧对应条目的「可呈现文本」，用于缺陷检测的内容比对。

    优先图形通道（图纸上实际画的字，一条标注一条文本，最准）；
    缺则回落语义通道，且必须**按片段取**——语义行常把尺寸行与 FCF 行揉在一起，
    直接拿整行会导致「缺数值 / 数量前缀不符」大面积误报。
    复合公差（一条开发标注对应多个语义实体）时把各片段合并，避免只比到一段。
    """
    if it.kind == KIND_NOTE:
        return ""
    ta = t.ta.get(it.name)
    if ta and ta.get("text"):
        return str(ta["text"])
    ids, _ = _lookup_semantic_for_dev(it, t)
    segs: List[str] = []
    for sid in ids:
        s = t.semantic.get(sid)
        if not s:
            continue
        seg = sfa_text_for(s["text"], it.kind) or str(s["text"])
        if seg:
            segs.append(seg)
    # 原样返回（保留 ⌀ / Ⓜ / ⌖ 等字形）：symbols_of 按字形判定符号，
    # 归一化会把 ⌀ 变成 D 从而丢掉符号，导致「缺符号」漏报。
    return " ".join(segs)


@dataclass
class SuspectLink:
    """一条关联失败条目的「疑似对应」提示。

    **只用于诊断展示**：不进结果表、不进任何指标分子分母。
    模糊匹配一旦参与判定，指标就不再是「ID 精确关联」的口径了。
    """
    key: str = ""
    title: str = ""
    name: str = ""
    candidates: List[Dict[str, Any]] = field(default_factory=list)

    def text(self) -> str:
        return " / ".join(f"{c['ta_name']}（{c['why']}）" for c in self.candidates)


def link_stats(rows: Sequence[MatchRow]) -> Dict[str, int]:
    """关联路径分布。开发侧几乎全为 `none` = 关联链断裂，而不是数据对不上。"""
    out: Dict[str, int] = {}
    for r in rows:
        out[r.path] = out.get(r.path, 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


def _fold_name(s: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(s or "")).lower()


def _strip_seq(s: str) -> str:
    return re.sub(r"\.\d+$", "", _fold_name(s))


def suspect_links(items: Sequence[DevItem], t: SfaTruth,
                  rows: Sequence[MatchRow], limit: int = 3) -> List[SuspectLink]:
    """对关联路径为 none 的开发条目，给出可能的对应关系。

    依次尝试：name 完全一致 → 规范化一致 → 去掉末尾序号一致 → 数值指纹重叠。
    仅用于体检面板展示，便于判断「是名字对不上还是整条链断了」。
    """
    broken = {r.key.split("@")[0] for r in rows if r.path == "none"}
    if not broken or not t.ta:
        return []
    by_fold: Dict[str, List[str]] = {}
    by_base: Dict[str, List[str]] = {}
    for nm in t.ta:
        by_fold.setdefault(_fold_name(nm), []).append(nm)
        by_base.setdefault(_strip_seq(nm), []).append(nm)

    # 预先算好每个 ta 的数值指纹，避免在内层循环里重复解析
    ta_nums: Dict[str, set] = {}
    for nm, info in t.ta.items():
        text = " ".join(str(t.semantic[r]["text"]) for r in info["sem_refs"]
                        if r in t.semantic) or str(info.get("text") or "")
        ta_nums[nm] = set(sem_tokens(text, [])["numbers"]) if text else set()

    out: List[SuspectLink] = []
    for it in items:
        if it.raw not in broken:
            continue
        cands: List[Dict[str, Any]] = []
        seen: set = set()

        def push(ta_name: str, why: str, score: float) -> None:
            if ta_name in seen:
                return
            seen.add(ta_name)
            cands.append({"ta_name": ta_name, "why": why, "score": score})

        if it.name:
            if it.name in t.ta:
                push(it.name, "name 完全一致，关联本应成功", 1.0)
            for nm in by_fold.get(_fold_name(it.name), []):
                push(nm, "name 规范化后一致", 0.9)
            for nm in by_base.get(_strip_seq(it.name), []):
                push(nm, "name 去掉末尾序号后一致", 0.7)

        if len(cands) < limit:
            dev_num = set(sem_tokens(it.title, [])["numbers"])
            if dev_num:
                scored = []
                for nm, sn in ta_nums.items():
                    if nm in seen or not sn:
                        continue
                    j = len(dev_num & sn) / len(dev_num | sn)
                    if j >= 0.6:
                        scored.append((j, nm))
                scored.sort(reverse=True)
                for j, nm in scored[:limit - len(cands)]:
                    push(nm, f"数值指纹重叠 {j:.0%}", round(j, 2))

        if cands:
            cands.sort(key=lambda c: -c["score"])
            out.append(SuspectLink(key=it.raw, title=it.title, name=it.name,
                                   candidates=cands[:limit]))
    return out


def match_items(t: SfaTruth, items: Sequence[DevItem]) -> List[MatchRow]:
    """按实体 ID 精确关联，产出比对行。"""
    rows: List[MatchRow] = []
    used_sem: Dict[int, int] = {}   # 语义 ID -> 覆盖它的开发条目序号
    used_sem_handle: Dict[int, Optional[int]] = {}   # 语义 ID -> 覆盖它的开发条目 handle

    if not items:
        return rows

    # 按 (分组, handle) 建索引：只有同一分组内重复出现才是真·重复导出。
    # 跨分组重复是「同一条 STEP 标注被多个保存视图共用」—— 例如 NIST ctc_02 里
    # 基准目标 A1~B4 同时挂在 DRAUGHTING_MODEL MBD_A 与 MBD_B 下（SFA 的 ta 表
    # `Saved Views` 列会写成 `(2) camera_model_d3 28 (MBD_A) 29 (MBD_B)`），
    # 开发侧按「视图 × 标注」导出必然出两条。这是合法结构，不是缺陷。
    by_handle_group: Dict[Tuple[str, int], List[DevItem]] = {}
    groups_of_handle: Dict[int, List[str]] = {}
    for it in items:
        if not it.handle:
            continue
        by_handle_group.setdefault((it.group, it.handle), []).append(it)
        gs = groups_of_handle.setdefault(it.handle, [])
        if it.group not in gs:
            gs.append(it.group)

    for idx, it in enumerate(items, start=1):
        key = f"{it.group}#{it.seq}"
        row = MatchRow(
            key=key, layer=KIND_LAYER.get(it.kind, LAYER_NOTE), kind=it.kind,
            group=it.group, dev_title=it.title, dev_name=it.name, handle=it.handle,
        )
        mv = groups_of_handle.get(it.handle or -1) or []
        if len(mv) > 1:
            row.multiview = "、".join(mv)
        ta = t.ta.get(it.name)
        row.sfa_graphic_text = normalize(ta["text"]) if ta else ""

        # 开发侧提取缺陷检测：与判定解耦，单独成层
        # view 优先用图形通道（图纸上实际呈现的文字），缺则回落语义通道
        sfa_name = t.dc.get(it.handle) if it.handle else ""
        view = it_view(it, t)
        dup_n = len(by_handle_group.get((it.group, it.handle or -1), []))
        codes, detail = dev_defects(it, view, dup=dup_n > 1, dup_count=dup_n,
                                    sfa_name=sfa_name or "")
        row.defects, row.defect_detail = codes, detail

        if it.kind == KIND_NOTE:
            has_graphic = bool(row.sfa_graphic_text)
            row.status = ST_NOTE
            row.remark = ("SFA 语义表结构上不产出；图形通道有文本" if has_graphic
                          else "SFA 双通道均无；内部项目独占提取")
            row.path = "note"
            rows.append(row)
            continue

        # SFA 侧一条真值都没有（图形专用导出）：判「不可比」，而不是「多余」。
        # 判「多余」会让精确率显示 0%，把「选错报告」伪装成「开发侧全错提取」——
        # 这是 ftc_08 `-e1-tg` 的表现。此类条目也不进任何指标分母。
        if t.truth_missing:
            row.layer = LAYER_NOTRUTH
            row.status = ST_ABSENT
            row.path = "no-truth"
            row.remark = ((row.remark + " ；" if row.remark else "")
                          + "该 SFA 报告不含语义 PMI（无语义表 / datum / datum_feature），"
                            "SFA 侧没有可比真值，本条不可判定 —— 不代表开发侧提取错误。"
                            "请改用同一测试件的语义版 SFA 报告后重跑。")
            rows.append(row)
            continue

        sem_ids, path = _lookup_semantic_for_dev(it, t)
        row.path = path

        # handle 一致性校验
        if it.handle and it.handle in t.dc and t.dc[it.handle] != it.name:
            row.remark = f"handle/name 与 SFA 不一致（SFA={t.dc[it.handle]}）"

        if not sem_ids:
            # 「接不上」要分两种，判错方向会同时污染召回率与精确率：
            #   a) ta 的引用落在**无语义文本的实体表**（dimensional_size /
            #      dimensional_location）上 —— 标注确实存在，只是 SFA 没为它导出语义值
            #      （图纸上是个没赋值的尺寸）。双侧都无可比内容，归注释层，不进语义分母。
            #   b) 引用接不上任何已知实体 —— 开发侧多提取，才是真「多余」。
            refs = (t.ta.get(it.name) or {}).get("sem_refs") or []
            if it.kind == KIND_DIM and refs and all(r in t.known_entities for r in refs):
                row.layer = LAYER_NOTE
                row.status = ST_NOTE
                row.path = "entity-no-semantic"
                row.remark = ("SFA 有该尺寸实体但未导出语义值（图纸上未赋值的尺寸）；"
                              "双侧无可比内容")
                rows.append(row)
                continue
            if it.kind == KIND_DATUM:
                row.status = ST_EXTRA
                row.remark = (row.remark + " ；" if row.remark else "") + "SFA datum 表未匹配到对应基准"
            else:
                row.status = ST_EXTRA
                row.remark = (row.remark + " ；" if row.remark else "") + "SFA 语义表无对应实体"
            rows.append(row)
            continue

        # 多条语义实体：逐条比对（复合公差）
        # SFA 拆成多条本身不是问题（复合公差就该是多条）；开发侧若把它们并成一条，
        # 由下面的 MRG 判据记缺陷 —— 故这里只留一个前提标记。
        merged_fcf = len(sem_ids) > 1
        extra_mods = dev_extra_mods(it)
        for sid in sem_ids:
            sub = MatchRow(**{**row.__dict__})
            # 浅拷贝会让所有子行共享同一个 defects 列表：缺陷必须按子行独立记，
            # 否则给某条子行补的码会串到同条目其它子行上。
            sub.defects = list(row.defects)
            sub.key = f"{key}@{sid}"      # 复合公差一条开发标注可能映射多个语义实体
            sub.sfa_sem_id = sid
            if it.kind == KIND_DATUM:
                ent = t.semantic.get(sid, {}).get("entity", "")
                dev_ident = str(it.detail.get("datum", "") or "")
                if path == "datum_feature-1hop":
                    # 基准字母来自 datum_feature 表。顺手把 datum 表里的同名字母标记为已用，
                    # 否则反向检查会再报一条「SFA datum 表有、开发侧未提取」的重复行。
                    ident = str(t.datum_feature.get(sid, "") or "")
                    sub.sfa_entity, sub.sfa_text = "datum_feature", ident
                    for did, d in t.datum.items():
                        if d.get("identification") == ident:
                            used_sem[("datum", did)] = key
                    if ident and ident == dev_ident:
                        sub.status, sub.remark = ST_HIT, "ID 关联 + 基准标识一致"
                    else:
                        sub.status = ST_SUSPECT
                        sub.remark = f"基准标识差异：开发={dev_ident or '-'} SFA={ident or '-'}"
                elif is_datum_target_entity(ent):
                    # 语义表文本形如 `A1 (point)` / `⌀85\nK1` / `1.25x2\nC1`：
                    # 标识两侧可能挂着目标形式与目标尺寸，按位置取首片段会拿到尺寸。
                    tid = str(it.detail.get("target id")
                              or it.detail.get("target_id")
                              or it.detail.get("target") or "")
                    expect = f"{dev_ident}{tid}"
                    seg = normalize(t.semantic[sid]["text"])
                    sub.sfa_entity, sub.sfa_text = ent, seg
                    ident = datum_target_ident(seg, expect)
                    # 该目标所定义的基准字母也要标记为已用：C、D 这类基准在 SFA 里
                    # **没有 datum_feature**（图纸上由目标区域定义，dev 只能提成
                    # datum target），不标记会让反向检查多报两条
                    # 「SFA datum 表有、开发侧未提取」，把基准覆盖率压到 67%。
                    for did, d in t.datum.items():
                        if dev_ident and d.get("identification") == dev_ident:
                            used_sem[("datum", did)] = key
                    if expect and ident and ident == expect:
                        sub.status, sub.remark = ST_HIT, "ID 关联 + 基准目标标识一致"
                    else:
                        sub.status = ST_SUSPECT
                        sub.remark = f"基准目标标识差异：开发={expect or '-'} SFA={ident or '-'}"
                else:
                    ident = str(t.datum.get(sid, {}).get("identification", "") or "")
                    sub.sfa_entity = "datum"
                    sub.sfa_text = ident
                    if ident and ident == dev_ident:
                        sub.status, sub.remark = ST_HIT, "ID 关联 + 基准标识一致"
                    else:
                        sub.status = ST_SUSPECT
                        sub.remark = f"基准标识差异：开发={dev_ident or '-'} SFA={ident or '-'}"
                used_sem[("datum", sid)] = key
                used_sem[sid] = key
                rows.append(sub)
                continue
            sub.sfa_entity = t.semantic[sid]["entity"]
            seg = sfa_text_for(t.semantic[sid]["text"], it.kind)
            sub.sfa_text = seg or normalize(t.semantic[sid]["text"])
            a = sem_tokens(seg, extra_mods)
            b = sem_tokens(it.title, extra_mods)
            if sid in used_sem:
                # 复用条目**仍要跑内容比对**。「第二次命中同一个语义 ID」本身
                # 不是差异信号 —— 开发侧按「视图 × 标注」导出，同一标注在
                # MBD_A / MBD_A(Work) 天然各出一条，这是 STEP 的合法结构
                # （SFA 的 ta 表 `Saved Views` 列会同时列出两个视图）。
                # 旧代码在这里无条件判 ST_HIT_DIFF，导致 stc_09 上 18 个复用标注
                # 制造出 12 条「归一化后只差一个空格」的假差异，观感是满屏标黄。
                # 现在改为照常比对内容：一致就保持命中，冲突才降级。
                sub.status, sub.remark = compare_tokens(b, a)
                if it.handle and used_sem_handle.get(sid) == it.handle:
                    # 同一个 handle 再来一条：跨视图复用由函数末尾统一追加说明，
                    # 同分组内才是真·重复导出，这里只补后者
                    if not row.multiview:
                        sub.remark += f"；同一 handle（{it.handle}）重复导出（见缺陷 DUP）"
                else:
                    sub.remark += f"；同一语义实体已被 {used_sem[sid]} 覆盖"
            else:
                sub.status, sub.remark = compare_tokens(b, a)

            # 开发侧把 SFA 拆开的多条复合公差并成了一条：匹配层保持「差异」
            # （口径「以 SFA 拆分为准」），缺陷层同时记一条 MRG，让它进缺陷统计。
            # 双条件缺一不可 —— SFA 只有一条时「数值多出」另有原因，不能定性为未拆分。
            if merged_fcf and sub.remark.startswith(MRG_REMARK) and DF_MRG not in sub.defects:
                sub.defects.append(DF_MRG)
                tip = (f"复合公差未拆分：SFA 拆成 {len(sem_ids)} 条"
                       f"（{'、'.join(str(x) for x in sem_ids)}），开发侧合并成 1 条")
                sub.defect_detail = (f"{sub.defect_detail}；{tip}"
                                     if sub.defect_detail else tip)

            used_sem[sid] = key
            used_sem_handle[sid] = it.handle
            rows.append(sub)

    # 反向：datum 表有、开发侧无
    for did, d in t.datum.items():
        if ("datum", did) in used_sem:
            continue
        rows.append(MatchRow(
            key=f"SFA-datum#{did}", layer=LAYER_DATUM, kind=KIND_DATUM,
            sfa_sem_id=did, sfa_text=d["identification"], sfa_entity="datum",
            status=ST_MISS, path="reverse",
            remark="SFA datum 表有，开发侧未提取",
        ))

    # 反向：语义表有、开发侧无
    for sid, s in t.semantic.items():
        if sid in used_sem:
            continue
        if is_datum_target_entity(s["entity"]):
            rows.append(MatchRow(
                key=f"SFA-datum-target#{sid}", layer=LAYER_DATUM, kind=KIND_DATUM,
                sfa_sem_id=sid, sfa_text=normalize(s["text"]), sfa_entity=s["entity"],
                status=ST_MISS, path="reverse",
                remark="SFA 有该基准目标，开发侧未提取",
            ))
            continue
        if s["kind"] == "datum_system":
            rows.append(MatchRow(
                key=f"SFA#{sid}", layer=LAYER_DATUM, kind="datum_system",
                sfa_sem_id=sid, sfa_text=normalize(s["text"]), sfa_entity=s["entity"],
                status=ST_NOTE, path="datum_system",
                remark="基准体系引用（GT 的组成部分，不独立对应标注）",
            ))
            continue
        rows.append(MatchRow(
            key=f"SFA#{sid}", layer=KIND_LAYER.get(s["kind"], LAYER_SEM), kind=s["kind"],
            sfa_sem_id=sid, sfa_text=normalize(s["text"]), sfa_entity=s["entity"],
            status=ST_MISS, path="reverse",
            remark="SFA 语义表有，开发侧未提取",
        ))

    # 多视图复用标记（最后统一追加，避免被中途的 remark 赋值覆盖）。
    # 只提示、不进缺陷层、不影响指标：这是 SFA 里一条标注挂多个保存视图的合法结构。
    for r in rows:
        if not r.multiview or "多视图" in r.remark:
            continue
        tag = f"该标注在多个保存视图中复用（{r.multiview}），非重复导出"
        r.remark = f"{r.remark} ；{tag}" if r.remark else tag
    return rows


# --------------------------------------------------------------------------
# 指标层
# --------------------------------------------------------------------------
def defect_groups(rows: Sequence[MatchRow]) -> Dict[str, Dict[str, Any]]:
    """按「开发标注」聚合缺陷，供指标层与两处报表（doctor / UI）共用。

    一条开发标注可能映射到多个语义实体（复合公差），比对时被拆成多条子行
    （`key@sid`）。子行只是「映射到哪个语义实体」的切分，**不是独立条目**，
    所以缺陷要合并去重后再计数：只取第一条子行会漏掉只出现在后续子行的码 ——
    `DF_MRG` 恰恰是在每条复合公差子行上才判得出来的那一个。

    返回 {开发条目 key: {"row": 首条子行, "codes": [...], "detail": "..."}}，
    顺序与 rows 首次出现顺序一致。
    """
    out: Dict[str, Dict[str, Any]] = {}
    for r in rows:
        if not r.defects:
            continue
        g = out.setdefault(r.key.split("@")[0], {"row": r, "codes": [], "detail": ""})
        for c in r.defects:
            if c not in g["codes"]:
                g["codes"].append(c)
        d = r.defect_detail or ""
        if d and d not in g["detail"]:
            g["detail"] = f"{g['detail']}；{d}" if g["detail"] else d
    return out


@dataclass
class Metrics:
    sem_expected: int = 0
    sem_extracted: int = 0
    sem_hit: int = 0
    sem_miss: int = 0
    sem_extra: int = 0
    sem_diff: int = 0
    sem_suspect: int = 0
    datum_expected: int = 0
    datum_hit: int = 0
    note_total: int = 0
    note_graphic_only: int = 0
    note_exclusive: int = 0
    # 「不可比」层：SFA 报告不含语义 PMI 时，开发侧条目全部落到这里。
    # 不进任何分母 —— 没有真值就没有「召回 / 精确」可言。
    no_truth: int = 0
    # 开发侧提取缺陷层（与上面各层解耦，不进 recall/precision 分母）
    defect_rows: int = 0
    defect_total: int = 0
    defect_by_code: Dict[str, int] = field(default_factory=dict)

    @property
    def truth_missing(self) -> bool:
        """整份报告是否「无从比对」——判据来自 match_items 落下的「不可比」行。"""
        return self.no_truth > 0

    @property
    def defect_rate(self) -> float:
        """有缺陷的开发条目占比（缺陷条目数 / 参与比对的开发条目数）。"""
        denom = (self.sem_extracted + self.datum_expected
                 + self.note_total + self.no_truth)
        return self.defect_rows / denom * 100 if denom else 0.0

    @property
    def recall(self) -> float:
        return self.sem_hit / self.sem_expected * 100 if self.sem_expected else 0.0

    @property
    def precision(self) -> float:
        return self.sem_hit / self.sem_extracted * 100 if self.sem_extracted else 0.0

    @property
    def datum_coverage(self) -> float:
        return self.datum_hit / self.datum_expected * 100 if self.datum_expected else 0.0


def compute_metrics(rows: Sequence[MatchRow], verdicts: Optional[Dict[str, str]] = None) -> Metrics:
    """按三层口径计算指标。

    - 语义层：几何公差 + 尺寸公差；非语义项不进 precision 分母
    - 基准层：datum 单列
    - 注释层：label/note 完全不进 PMI 分母
    """
    v = verdicts or {}
    m = Metrics()
    seen_datum: set = set()
    for i, r in enumerate(rows):
        verdict = v.get(r.key, "待定")
        if r.layer == LAYER_SEM:
            if r.status == ST_MISS:
                m.sem_expected += 1
                m.sem_miss += 1
                continue
            m.sem_extracted += 1
            m.sem_expected += 1
            if r.status == ST_HIT:
                if verdict == "误报":
                    m.sem_miss += 1
                else:
                    m.sem_hit += 1
            elif r.status == ST_HIT_DIFF:
                m.sem_diff += 1
                if verdict != "误报":
                    m.sem_hit += 1
            elif r.status == ST_SUSPECT:
                m.sem_suspect += 1
                if verdict == "误报":
                    m.sem_miss += 1
            elif r.status == ST_EXTRA:
                m.sem_extra += 1
                if verdict == "误报":
                    m.sem_extracted -= 1
        elif r.layer == LAYER_DATUM:
            if r.kind == "datum_system":
                continue
            # 同一条标注跨视图重复导出只算一次：否则「基准_应提取」会按视图数翻倍
            # （ctc_02 的 8 个基准目标会算成 16），覆盖率虽然不变，数字会误导。
            dk = f"handle:{r.handle}" if r.handle else r.key
            if dk in seen_datum:
                continue
            seen_datum.add(dk)
            m.datum_expected += 1
            if r.status in (ST_HIT, ST_HIT_DIFF):
                m.datum_hit += 1
        elif r.layer == LAYER_NOTRUTH:
            # 无真值：两侧没有共同基准，既不进分子也不进分母。
            m.no_truth += 1
        else:  # LAYER_NOTE
            m.note_total += 1
            if r.status == ST_NOTE:
                if "图形通道有文本" in r.remark:
                    m.note_graphic_only += 1
                else:
                    m.note_exclusive += 1

    # 缺陷层：按「开发条目」聚合（复合公差的多个子行属同一条开发标注）
    groups = defect_groups(rows)
    m.defect_rows = len(groups)
    m.defect_total = sum(len(g["codes"]) for g in groups.values())
    for g in groups.values():
        for c in g["codes"]:
            m.defect_by_code[c] = m.defect_by_code.get(c, 0) + 1
    return m


def summarize(rows: Sequence[MatchRow], t: SfaTruth, items: Sequence[DevItem]) -> Dict[str, Any]:
    m = compute_metrics(rows)
    paths = link_stats(rows)
    return {
        "开发条目": len(items),
        "SFA语义项": len(t.semantic),
        "SFA表数": len(t.sheets),
        # 指标是否成立：SFA 侧无真值时三项比率都没有意义，导出里要能看出来
        "指标是否成立": "否（SFA 报告不含语义 PMI）" if m.truth_missing else "是",
        "不可比条目": m.no_truth,
        "关联覆盖率": (sum(1 for r in rows if r.path in ("gt-1hop", "dim-2hop", "datum-2hop"))
                   / max(1, len([x for x in items if x.kind != KIND_NOTE])) * 100),
        "关联路径分布": paths,
        "断链条目": paths.get("none", 0),
        "语义_应提取": m.sem_expected,
        "语义_已提取": m.sem_extracted,
        "语义_命中": m.sem_hit,
        "语义_缺失": m.sem_miss,
        "语义_多余": m.sem_extra,
        "语义_文本差异": m.sem_diff,
        "语义_疑似": m.sem_suspect,
        "召回率": round(m.recall, 2),
        "精确率": round(m.precision, 2),
        "基准_应提取": m.datum_expected,
        "基准_命中": m.datum_hit,
        "基准覆盖率": round(m.datum_coverage, 2),
        "注释_总数": m.note_total,
        "注释_图形通道有": m.note_graphic_only,
        "注释_内部独占": m.note_exclusive,
        "缺陷_条目数": m.defect_rows,
        "缺陷_总数": m.defect_total,
        "缺陷_占比": round(m.defect_rate, 2),
        "缺陷_分布": dict(sorted(m.defect_by_code.items(), key=lambda kv: -kv[1])),
        "警告": t.warnings,
    }


__all__ = [
    "ST_HIT", "ST_HIT_DIFF", "ST_SUSPECT", "ST_PARTIAL", "ST_MISS", "ST_EXTRA",
    "ST_GAIN", "ST_NOTE", "ST_ABSENT",
    "KIND_GT", "KIND_DIM", "KIND_DATUM", "KIND_NOTE",
    "LAYER_SEM", "LAYER_DATUM", "LAYER_NOTE", "LAYER_NOTRUTH",
    "DF_ENC", "DF_EMPTY", "DF_CNT", "DF_SYM", "DF_NUM", "DF_MAP", "DF_DUP", "DF_MRG",
    "DEFECT_CODES", "MRG_REMARK", "defect_groups",
    "SfaTruth", "DevItem", "MatchRow", "Metrics",
    "load_sfa", "parse_dev_markdown", "match_items", "compute_metrics", "summarize",
    "normalize", "sem_tokens", "fix_mojibake", "norm_number",
    "dev_defects", "symbols_of", "datum_target_ident", "suggest_reports_with_truth",
    "is_datum_target_entity", "is_known_entity", "classify_sfa_entity",
]
