# -*- coding: utf-8 -*-
"""
PMI 比对核心模块
==================

真值源：NIST STEP File Analyzer (SFA) 输出的 xlsx 报告
被测源：内部项目导出的 markdown（每条含 handle / name / type / detailData）

**关联方式是实体 ID 精确关联，不是字符串相似度。**

关联链路（均已用 nist_ftc_07 NIST 测试件实测验证）::

    md.handle  ==  SFA.draughting_callout.ID
    md.name    ==  SFA.draughting_callout.name  ==  SFA.tessellated_annotation_occurrence.name

    语义映射三条路径：
      GT / FCF : ta.col11 末尾 ID                             -> 语义表 ID      (1 跳)
      DIM      : ta.col11 的 dimensional_size / _location ID
                    -> dcr.dimension 匹配 -> dcr.ID            -> 语义表 ID      (2 跳)
      datum    : ta.col11 末尾形貌 ID - 1 -> datum.ID          -> datum.identification (2 跳)

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

# 条目类别
KIND_GT    = "gt"      # 几何公差（FCF）
KIND_DIM   = "dim"     # 尺寸公差
KIND_DATUM = "datum"   # 基准
KIND_NOTE  = "note"    # 注释 / 标签文本

# 指标分层
LAYER_SEM   = "语义PMI"
LAYER_DATUM = "基准"
LAYER_NOTE  = "非语义注释"

KIND_LAYER = {
    KIND_GT: LAYER_SEM,
    KIND_DIM: LAYER_SEM,
    KIND_DATUM: LAYER_DATUM,
    KIND_NOTE: LAYER_NOTE,
}

# SFA 语义表里实体类型 -> 条目类别
def classify_sfa_entity(entity: str) -> str:
    e = (entity or "").strip()
    if "dimensional_characteristic_representation" in e:
        return KIND_DIM
    if "datum_system" in e:
        return "datum_system"          # 组合项，不独立对应标注
    if "tolerance" in e:
        return KIND_GT
    return KIND_GT


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


_FCF_SYM = re.compile(r"[⌓⌖⟂∥▱⏊⏥⫽]")
_MOD_WORDS = {
    "MAXIMUM_MATERIAL": "M", "LEAST_MATERIAL": "L",
    "STATISTICAL_TOLERANCE": "ST", "PROJECTED": "P", "FREE_STATE": "F",
}


def sem_tokens(text: str, extra_mods: Iterable[str] = ()) -> Dict[str, Any]:
    """抽取语义指纹：数值集合 + 基准字母 + 修饰符 + 数量前缀。

    数量前缀（`4X`）从数值集合中剔除，单独用 count 字段比较。
    """
    s = normalize(text, keep_sep=False)
    cnt_m = re.match(r"^\s*(\d+)\s*[Xx]\b", s)
    count = int(cnt_m.group(1)) if cnt_m else None
    body = s[cnt_m.end():] if cnt_m else s
    nums = sorted({norm_number(x) for x in _NUM.findall(body)})
    letters = sorted(set(re.findall(r"(?<![A-Za-z])([A-Z])(?![A-Za-z])", s)))
    # 修饰符用「非字母」边界匹配：`Ø.015Ⓟ.425` 归一化后是 `D0.015P0.425`，
    # 用 \b 会漏掉（P 两侧都是数字，仍属 word 字符）。
    mods = {m for m in ("M", "L", "P", "ST", "U", "F")
            if re.search(rf"(?<![A-Za-z]){m}(?![A-Za-z])", s)}
    mods |= set(extra_mods)
    return {
        "numbers": nums,
        "letters": letters,
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
    """从 SFA 语义文本中取用于比对的片段。

    GT 类语义文本常把「尺寸行 + FCF 行 + 修饰行」合并，只取含公差符号的行。
    DIM 类尺寸描述可能折行（如 `⌀0.250 +.003` + `-.000`），需合并全部行。
    """
    lines = [l for l in (text or "").split("\n") if l.strip()]
    if not lines:
        return ""
    if kind == KIND_GT:
        fcf = [l for l in lines if _FCF_SYM.search(l)]
        if fcf:
            return " ".join(fcf[:1])
        return normalize(lines[0], keep_sep=False)
    return normalize(" ".join(lines), keep_sep=False)


def compare_tokens(dev: Dict[str, Any], sfa: Dict[str, Any]) -> Tuple[str, str]:
    """返回 (状态, 备注)。

    子集关系一律视为「ID 已确认一致，仅呈现完整度不同」；只有集合冲突才判疑似。
    """
    dn, sn = set(dev["numbers"]), set(sfa["numbers"])
    dl, sl = set(dev["letters"]), set(sfa["letters"])
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
        return ST_HIT_DIFF, "ID 关联一致；开发侧含额外数值（疑似合并多条复合公差）" + cnt_note

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

DEFECT_CODES = (DF_ENC, DF_EMPTY, DF_CNT, DF_SYM, DF_NUM, DF_MAP, DF_DUP)

# 符号归一化别名（仅用于缺陷比对，不改变正文归一化口径）
_SYM_ALIAS = {"⊥": "⟂", "⏊": "⟂", "⊿": "⌓", "⫽": "∥", "⏥": "▱",
              "∅": "⌀", "Ø": "⌀", "Φ": "⌀", "φ": "⌀", "ø": "⌀"}
_FCF_CHARS = set("⌓⌖⟂∥▱⏥")
_DIA_CHARS = set("⌀")
_RAD_PREFIX = re.compile(r"(?<![A-Za-z])R\s*(?=[.\d])")
_SR_PREFIX = re.compile(r"(?<![A-Za-z])S\s*(?=[⌀.])")
# 数量前缀可能不在串首（SFA 图形文本形如 `DIM | 4X ⌀.250`），也常用 `×` 而非 `X`
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
    """抽取用于比对的图形符号集合（含直径符 / 半径符 / 球面符）。"""
    s = "".join(_SYM_ALIAS.get(ch, ch) for ch in str(text or ""))
    out = {ch for ch in s if ch in _FCF_CHARS or ch in _DIA_CHARS}
    if _RAD_PREFIX.search(s):
        out.add("R")
    if _SR_PREFIX.search(s):
        out.add("S")
    return out


def dev_defects(it: "DevItem", view: str = "", *, dup: bool = False,
                sfa_name: str = "") -> Tuple[List[str], str]:
    """检测单条开发标注的提取缺陷。

    `view` 为 SFA 侧对应条目的可呈现文本（优先图形通道，缺则语义通道）。
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

    # 3) 与 SFA 图形通道逐维比对（数量前缀 / 符号 / 数值）
    if view:
        raw_view = str(view)
        d_cnt = _CNT_PREFIX.search(fixed)
        s_cnt = _CNT_PREFIX.search(raw_view)
        dv = int(d_cnt.group(1)) if d_cnt else None
        sv = int(s_cnt.group(1)) if s_cnt else None
        if dv != sv:
            codes.append(DF_CNT)
            notes.append(f"数量前缀 开发={dv if dv else '无'} SFA={sv if sv else '无'}")

        lost = symbols_of(raw_view) - symbols_of(fixed)
        if lost:
            codes.append(DF_SYM)
            notes.append("缺符号 " + " ".join(sorted(lost)))

        # 数量前缀单独由 DF_CNT 负责，这里先从两侧剔除；
        # 只取带小数点的数值（整数常是 `⌀1` 这类定义区域标注），零值也剔除
        # （`-0` / `-.000` 各通道写法不统一，不足以定性为开发缺陷）
        dnums = {_absnum(x) for x in _NUM.findall(_CNT_PREFIX.sub("", fixed))}
        snums = {_absnum(x) for x in _NUM.findall(_CNT_PREFIX.sub("", raw_view))}
        snums = {v for v in snums if _tolerance_like(v)}
        miss = {v for v in snums if not any(_num_equivalent(v, d) for d in dnums)}
        if miss:
            codes.append(DF_NUM)
            notes.append("缺数值 " + "/".join(sorted(miss)[:6]))

    # 4) 映射一致性
    if sfa_name and it.name and sfa_name != it.name:
        codes.append(DF_MAP)
        notes.append(f"handle 指向 {sfa_name!r}，标注写的是 {it.name!r}")

    # 5) 重复 handle
    if dup and it.handle:
        codes.append(DF_DUP)
        notes.append(f"handle {it.handle} 在同一次比对中出现多条")

    return codes, "；".join(notes)


# --------------------------------------------------------------------------
# SFA 真值层
# --------------------------------------------------------------------------
@dataclass
class SfaTruth:
    path: str = ""
    units: str = ""
    semantic: Dict[int, Dict[str, Any]] = field(default_factory=dict)   # 语义表 ID -> item
    datum: Dict[int, Dict[str, Any]] = field(default_factory=dict)      # datum.ID -> item
    dc: Dict[int, str] = field(default_factory=dict)                    # draughting_callout.ID -> name
    ta: Dict[str, Dict[str, Any]] = field(default_factory=dict)         # name -> {id, sem_refs, text}
    dcr_by_dim: Dict[int, int] = field(default_factory=dict)            # dimensional_*.ID -> dcr.ID
    sheets: List[str] = field(default_factory=list)
    sentinel_row: Optional[int] = None
    warnings: List[str] = field(default_factory=list)


_REF_ID = re.compile(r"(?<![\d.])(\d{4,7})(?![\d.])")


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


def load_sfa(path: str) -> SfaTruth:
    """读取 SFA 报告，建立 ID 索引。"""
    t = SfaTruth(path=path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    t.sheets = list(wb.sheetnames)

    # --- 1. Semantic PMI Summary：语义真值
    sname = _pick_sheet(t.sheets, "Semantic PMI Summary") or _pick_sheet(t.sheets, "Semantic PMI Summ")
    if not sname:
        t.warnings.append("未找到语义表（Semantic PMI Summary）")
    else:
        for i, r in enumerate(_rows(wb[sname])):
            if len(r) >= 3 and r[2] == "Expected PMI":
                t.sentinel_row = i + 1
                break
        for r in _rows(wb[sname]):
            if len(r) >= 3 and r[0].isdigit() and r[2] not in ("", "Expected PMI"):
                t.semantic[int(r[0])] = {
                    "id": int(r[0]), "entity": r[1], "text": r[2],
                    "similar": r[3] if len(r) > 3 else "",
                    "kind": classify_sfa_entity(r[1]),
                }

    # --- 2. draughting_callout：handle <-> name
    sname = _pick_sheet(t.sheets, "draughting_callout")
    if sname:
        for r in _rows(wb[sname]):
            if len(r) >= 2 and r[0].isdigit():
                t.dc[int(r[0])] = r[1]
    else:
        t.warnings.append("未找到 draughting_callout 表")

    # --- 3. tessellated_annotation_occurrence：name -> 语义引用 + 图形文本
    sname = _pick_sheet(t.sheets, "tessellated_annotation_occurren")
    if not sname:
        t.warnings.append("未找到 tessellated_annotation_occurrence 表（图形 PMI 通道缺失）")
    else:
        for r in _rows(wb[sname]):
            if len(r) >= 14 and r[0].isdigit():
                refs = _ref_ids(r[10])
                t.ta[r[1]] = {
                    "id": int(r[0]),
                    "ta_name": r[1],
                    "sem_refs": refs,
                    "sem_ref_raw": r[10],
                    "text": r[13],
                }

    # --- 4. datum：基准标识
    sname = _pick_sheet(t.sheets, "datum")
    if sname and sname != _pick_sheet(t.sheets, "datum_system"):
        for r in _rows(wb[sname]):
            if len(r) >= 6 and r[0].isdigit():
                t.datum[int(r[0])] = {"id": int(r[0]), "identification": r[5]}

    # --- 5. dimensional_characteristic_representation：尺寸的第 2 跳
    sname = _pick_sheet(t.sheets, "dimensional_characteristic_repr")
    if sname:
        for r in _rows(wb[sname]):
            if len(r) >= 2 and r[0].isdigit():
                refs = _ref_ids(r[1])
                for rid in refs:
                    t.dcr_by_dim[rid] = int(r[0])

    # --- 单位
    hname = _pick_sheet(t.sheets, "Header")
    if hname:
        for r in _rows(wb[hname]):
            joined = " ".join(r)
            m = re.search(r"\b(INCH|MM|MILLIMETRE|METER)\b", joined, re.I)
            if m:
                t.units = m.group(1).upper()
                break
    wb.close()
    return t


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
        if any(x in ("dimensional_size", "dimensional_location", "dimensional_characteristic_representation") for x in ts):
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
_META = re.compile(r"^-\s*(分组\s*Handle|标注数量|标注总数|视图分组)\s*[:：]\s*(.*)$")


def parse_dev_markdown(text: str) -> List[DevItem]:
    """解析内部项目导出的 markdown。"""
    items: List[DevItem] = []
    group, meta = "", {}
    i, lines = 0, text.splitlines()
    cur: Optional[DevItem] = None

    while i < len(lines):
        ln = lines[i]
        m2 = _H2.match(ln)
        m3 = _H3.match(ln)
        mm = _META.match(ln.strip())
        if m2:
            group = m2.group(1).strip()
            i += 1
            continue
        if mm:
            meta[mm.group(1)] = mm.group(2)
            i += 1
            continue
        if m3:
            if cur:
                items.append(cur)
            cur = DevItem(group=group, seq=int(m3.group(1)), title=m3.group(2), raw=ln)
            i += 1
            continue
        if cur is not None and ln.strip() == "detailData:":
            i += 1
            buf, depth, started = [], 0, False
            while i < len(lines):
                l = lines[i]
                if "{" in l:
                    started = True
                if started:
                    buf.append(l)
                    depth += l.count("{") - l.count("}")
                    if depth <= 0 and started:
                        break
                elif l.strip():
                    break
                i += 1
            blob = "\n".join(buf)
            # 去掉尾部 ``` 之类的包围
            blob = re.sub(r"^[^\{]*", "", blob, count=1).strip()
            try:
                d = json.loads(blob)
            except Exception:
                d = {}
            cur.detail = d if isinstance(d, dict) else {}
            cur.handle = _to_int(cur.detail.get("handle"))
            cur.name = str(cur.detail.get("name", "") or "")
            cur.original_type_name = str(cur.detail.get("original_type_name", "") or "")
            cur.types = [x for x in str(cur.detail.get("type", "") or "").split("+") if x]
            i += 1
            continue
        i += 1
    if cur:
        items.append(cur)

    # 视图顺序：按首次出现顺序编号
    order = []
    for it in items:
        if it.group not in order:
            order.append(it.group)
    for it in items:
        it.raw = f"{it.group}#{it.seq}"
    return items


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
            "人工校验": "待定",
            "备注": self.remark,
        }


def _lookup_semantic_for_dev(it: DevItem, t: SfaTruth) -> Tuple[List[int], str]:
    """返回 (语义表 ID 列表, 关联路径)。"""
    ta = t.ta.get(it.name)
    if not ta:
        return [], "none"
    refs = list(ta["sem_refs"])
    raw = ta["sem_ref_raw"]

    if it.kind == KIND_GT:
        ids = [r for r in refs if r in t.semantic]
        return ids, "gt-1hop" if ids else "none"

    if it.kind == KIND_DIM:
        # col11 指向 dimensional_size / dimensional_location，需经 dcr 映射
        ids = []
        for r in refs:
            dcr = t.dcr_by_dim.get(r) or t.dcr_by_dim.get(r + 1)
            if dcr and dcr in t.semantic:
                ids.append(dcr)
        if ids:
            return ids, "dim-2hop"
        ids = [r for r in refs if r in t.semantic]
        return ids, ("dim-2hop" if ids else "none")

    if it.kind == KIND_DATUM:
        # col11 里 datum_feature / shape_aspect 的 ID，datum.ID = 该 ID - 1
        for r in refs:
            if (r - 1) in t.datum:
                return [r - 1], "datum-2hop"
        return [], "none"

    return [], "none"


def it_view(it: DevItem, t: SfaTruth) -> str:
    """SFA 侧对应条目的「可呈现文本」：优先图形通道（图纸上实际画的字），
    缺则回落语义通道。注释类不做内容比对（标题只是元素名，比不出缺陷）。"""
    if it.kind == KIND_NOTE:
        return ""
    ta = t.ta.get(it.name)
    if ta and ta.get("text"):
        return str(ta["text"])
    ids, _ = _lookup_semantic_for_dev(it, t)
    for sid in ids:
        if sid in t.semantic:
            return str(t.semantic[sid]["text"])
    return ""


def match_items(t: SfaTruth, items: Sequence[DevItem]) -> List[MatchRow]:
    """按实体 ID 精确关联，产出比对行。"""
    rows: List[MatchRow] = []
    used_sem: Dict[int, int] = {}   # 语义 ID -> 覆盖它的开发条目序号

    if not items:
        return rows

    # 按 handle 建索引，便于检测重复
    by_handle: Dict[int, List[DevItem]] = {}
    for it in items:
        if it.handle:
            by_handle.setdefault(it.handle, []).append(it)

    for idx, it in enumerate(items, start=1):
        key = f"{it.group}#{it.seq}"
        row = MatchRow(
            key=key, layer=KIND_LAYER.get(it.kind, LAYER_NOTE), kind=it.kind,
            group=it.group, dev_title=it.title, dev_name=it.name, handle=it.handle,
        )
        ta = t.ta.get(it.name)
        row.sfa_graphic_text = normalize(ta["text"]) if ta else ""

        # 开发侧提取缺陷检测：与判定解耦，单独成层
        # view 优先用图形通道（图纸上实际呈现的文字），缺则回落语义通道
        sfa_name = t.dc.get(it.handle) if it.handle else ""
        view = it_view(it, t)
        codes, detail = dev_defects(it, view, dup=len(by_handle.get(it.handle or -1, [])) > 1,
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

        sem_ids, path = _lookup_semantic_for_dev(it, t)
        row.path = path

        # handle 一致性校验
        if it.handle and it.handle in t.dc and t.dc[it.handle] != it.name:
            row.remark = f"handle/name 与 SFA 不一致（SFA={t.dc[it.handle]}）"

        if not sem_ids:
            if it.kind == KIND_DATUM:
                row.status = ST_EXTRA
                row.remark = (row.remark + " ；" if row.remark else "") + "SFA datum 表未匹配到对应基准"
            else:
                row.status = ST_EXTRA
                row.remark = (row.remark + " ；" if row.remark else "") + "SFA 语义表无对应实体"
            rows.append(row)
            continue

        # 多条语义实体：逐条比对（复合公差）
        extra_mods = dev_extra_mods(it)
        for sid in sem_ids:
            sub = MatchRow(**{**row.__dict__})
            sub.key = f"{key}@{sid}"      # 复合公差一条开发标注可能映射多个语义实体
            sub.sfa_sem_id = sid
            if it.kind == KIND_DATUM:
                ident = str(t.datum.get(sid, {}).get("identification", "") or "")
                dev_ident = str(it.detail.get("datum", "") or "")
                sub.sfa_entity = "datum"
                sub.sfa_text = ident
                if ident and ident == dev_ident:
                    sub.status, sub.remark = ST_HIT, "ID 关联 + 基准标识一致"
                else:
                    sub.status = ST_SUSPECT
                    sub.remark = f"基准标识差异：开发={dev_ident or '-'} SFA={ident or '-'}"
                used_sem[("datum", sid)] = key
                rows.append(sub)
                continue
            sub.sfa_entity = t.semantic[sid]["entity"]
            seg = sfa_text_for(t.semantic[sid]["text"], it.kind)
            sub.sfa_text = seg or normalize(t.semantic[sid]["text"])
            a = sem_tokens(seg, extra_mods)
            b = sem_tokens(it.title, extra_mods)
            if sid in used_sem:
                sub.status = ST_HIT_DIFF
                sub.remark = f"同一语义实体已被 {used_sem[sid]} 覆盖"
            else:
                sub.status, sub.remark = compare_tokens(b, a)
            used_sem[sid] = key
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
    return rows


# --------------------------------------------------------------------------
# 指标层
# --------------------------------------------------------------------------
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
    # 开发侧提取缺陷层（与上面各层解耦，不进 recall/precision 分母）
    defect_rows: int = 0
    defect_total: int = 0
    defect_by_code: Dict[str, int] = field(default_factory=dict)

    @property
    def defect_rate(self) -> float:
        """有缺陷的开发条目占比（缺陷条目数 / 参与比对的开发条目数）。"""
        denom = self.sem_extracted + self.datum_expected + self.note_total
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
            m.datum_expected += 1
            if r.status in (ST_HIT, ST_HIT_DIFF):
                m.datum_hit += 1
        else:  # LAYER_NOTE
            m.note_total += 1
            if r.status == ST_NOTE:
                if "图形通道有文本" in r.remark:
                    m.note_graphic_only += 1
                else:
                    m.note_exclusive += 1

    # 缺陷层：按「开发条目」去重（复合公差的多个子行共享同一条开发标注）
    per_item: Dict[str, List[str]] = {}
    for r in rows:
        if r.defects:
            per_item.setdefault(r.key.split("@")[0], list(r.defects))
    m.defect_rows = len(per_item)
    m.defect_total = sum(len(v) for v in per_item.values())
    for codes in per_item.values():
        for c in codes:
            m.defect_by_code[c] = m.defect_by_code.get(c, 0) + 1
    return m


def summarize(rows: Sequence[MatchRow], t: SfaTruth, items: Sequence[DevItem]) -> Dict[str, Any]:
    m = compute_metrics(rows)
    return {
        "开发条目": len(items),
        "SFA语义项": len(t.semantic),
        "SFA表数": len(t.sheets),
        "关联覆盖率": (sum(1 for r in rows if r.path in ("gt-1hop", "dim-2hop", "datum-2hop"))
                   / max(1, len([x for x in items if x.kind != KIND_NOTE])) * 100),
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
    "ST_HIT", "ST_HIT_DIFF", "ST_SUSPECT", "ST_MISS", "ST_EXTRA", "ST_GAIN", "ST_NOTE",
    "KIND_GT", "KIND_DIM", "KIND_DATUM", "KIND_NOTE",
    "LAYER_SEM", "LAYER_DATUM", "LAYER_NOTE",
    "DF_ENC", "DF_EMPTY", "DF_CNT", "DF_SYM", "DF_NUM", "DF_MAP", "DF_DUP", "DEFECT_CODES",
    "SfaTruth", "DevItem", "MatchRow", "Metrics",
    "load_sfa", "parse_dev_markdown", "match_items", "compute_metrics", "summarize",
    "normalize", "sem_tokens", "fix_mojibake", "norm_number",
    "dev_defects", "symbols_of",
]
