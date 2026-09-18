# STP 比对工具 · PMI 数据一致性比对

以 **NIST SFA（STEP File Analyzer）** 报告为公差真值，与**内部项目提取的 markdown** 做逐条比对，
定位两边数据的差异来源，并把**开发侧提取缺陷**单独暴露出来。

## 为什么需要它

SFA 解析的是 STEP 里的语义 PMI，**不包含非语义信息**（图形标注、注释文本、标签）；
内部项目的 markdown 是支持的。两边口径天然不同，直接比会一片红。
本工具的作用不是"把差异抹平"，而是**把差异说清楚**：哪些是真差异、哪些是口径差异、
哪些是开发侧提取本身出错（乱码、丢符号、丢数量前缀……）。

## 关联原理

**实体 ID 精确关联，不做字符串相似度匹配。**

```
md.handle  ==  SFA.draughting_callout.ID
md.name    ==  SFA.draughting_callout.name == SFA.tessellated_annotation_occurrence.name
```

语义映射三条路径（均已在 `nist_ftc_07` NIST 测试件上实测打通）：

| 类别 | 路径 | 跳数 |
|---|---|---|
| GT / FCF | `ta.col11` 末尾 ID → 语义表 ID | 1 |
| DIM | `ta.col11` 的 dimensional_size / _location ID → `dcr.dimension` 匹配 → `dcr.ID` → 语义表 ID | 2 |
| datum | `ta.col11` 末尾形貌 ID − 1 → `datum.ID` → `datum.identification` | 2 |

## 指标口径（三层，互不污染）

| 层 | 覆盖内容 | 是否进 precision 分母 |
|---|---|---|
| 语义 PMI 层（主指标） | 几何公差 + 尺寸公差 | 是 |
| 基准层（单列） | SFA `datum` 表 vs md 基准标注 | 独立统计 |
| 非语义 / 注释层（单列） | label / note 类 | 否 |

**缺陷检测层与指标层解耦**：编码损坏等缺陷在归一化阶段会被修复以保证比对可进行，
但修复动作全程留痕，计入独立的缺陷清单，**不拉低召回率 / 精确率分母**。

## 缺陷码

| 码 | 含义 |
|---|---|
| `ENC` | 编码损坏（乱码，如 `飦` → `Ⓢ`、`鈱´` → `⌴`） |
| `CNT` | 数量前缀不符（丢 `4X` / `2X`，或 `2×` 误写为 `32X`） |
| `SYM` | 符号丢失（缺 `⌀`、`∥`、`R`） |
| `NUM` | 数值缺失 |
| `EMP` | 标题为空 |
| `MAP` | handle / name 不一致 |
| `DUP` | handle 重复 |

## 安装与运行

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
pip install -r requirements.txt

streamlit run pmi_compare_work.py
```

Windows 下也可直接双击 `启动工具.bat`。
PyCharm 中直接运行 `pmi_compare_work.py` 也可以 —— 脚本内置 bare mode 兜底，
检测到无 ScriptRunContext 时会自动转交 `streamlit run`。

## 使用流程

1. 用 SFA 导出 xlsx 报告（需包含 Semantic PMI Summary / tessellated_annotation_occurrence / draughting_callout 等表）。
2. 在页面左侧上传 SFA xlsx。
3. 上传或粘贴内部项目导出的 markdown。
4. 查看三层指标、结果明细表、以及「🔧 提取缺陷清单」。

## 文件结构

```
pmi_core.py            比对内核：真值装载 / markdown 解析 / 归一化 / ID 关联 / 指标 / 缺陷检测
pmi_compare_work.py    Streamlit 界面（当前主入口）
test_pmi_core.py       回归测试（87 项，含缺陷与指标解耦断言）
streamlit_launcher.py  启动器（规避 IDE 运行配置序列化差异）
pmi_compare.py         早期版本，留档
```

## 测试

```bash
python -m pytest test_pmi_core.py -q
```

## 环境

Python 3.13 / Streamlit 1.64 / pandas / openpyxl / xlrd

## 备注

仓库不包含 SFA 安装包、NIST 标准文档与测试用 xlsx（受版权与体积限制，见 `.gitignore`）。
