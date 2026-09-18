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

语义映射四条路径（均已在 `nist_ftc_07` / `nist_ftc_10` NIST 测试件上实测打通）：

| 类别 | 路径 | 跳数 |
|---|---|---|
| GT / FCF | `ta.col11` 末尾 ID → 语义表 ID | 1 |
| DIM | `ta.col11` 的 dimensional_size / _location ID → `dcr.dimension` 匹配 → `dcr.ID` → 语义表 ID | 2 |
| datum | `ta.col11` 末尾形貌 ID − 1 → `datum.ID` → `datum.identification` | 2 |
| datum_target | `ta.col11` 直接给出 `datum_target` 实体 ID → 语义表 ID | 1 |

> `ta` = `tessellated_annotation_occurrence`。**列位置一律按表头列名定位，不写死列号** ——
> SFA 的列数与列序会随版本和导出勾选变化（实测 `ta` 表 12~14 列、`dcr` 表 17~20 列）。
> 关联键用 `Associated Semantic PMI` 列；`Equivalent Unicode String(s)` 是图形文本通道，
> 属可选列：**缺它不影响关联**，只是「SFA图形文本」列会为空、内容比对回落到语义表文本。
> 导出时勾上该列可获得最精确的按字形比对。

## 排查「一条都没对上」

如果结果表里开发侧条目**全部**是「⚠️ 多余」、同时 SFA 侧大批「❌ 缺失」，
那不是数据真的对不上，而是 **ID 关联链没建立**。根因必在两侧之一，页面提供两个自检面板：

**①「🔎 解析自检」—— 开发侧**

| 指标 | 说明 |
|---|---|
| 解析条目 | markdown 里解析出多少条标注 |
| 带 handle / 带 name | 有多少条拿到了 ID 关联依据——**这个不是 100% 就必然全红** |
| 未解析到 detailData | 没找到 `detailData:` 块的条数 |
| JSON 解析失败 | JSON 破损 / 括号未闭合的条数 |

以及逐条问题明细、未匹配的标题行样例。`handle` 与 `name` 是 ID 关联的唯一依据。

解析器对常见写法差异做了容错：`detailData:` 允许带引号、全角冒号、JSON 与标签同行、
```json 围栏；JSON 字段名 `handle`/`name` 大小写不敏感。但**标题行必须**匹配
`^###\s+序号[.、]标题$`（如 `### 1. Simple Datum.1`）。

**②「🧭 报告体检」—— SFA 侧 + 关联链**

一个面板四段，把「工具读到了什么、列是怎么认的、哪一环断了」全部摊开：

| 段 | 内容 |
|---|---|
| ① 装载与列定位 | 每张表的列数、表头行、**用到的列是怎么定位的**；来源显示「默认列号（脆弱）」说明该表没识别出表头 |
| ② 交叉校验 | 表内数据行数 vs 实际装载条数、`ta` 引用能否落地、`dc` 与 `ta` 的 name 是否一一对应 |
| ③ 关联路径分布 | `gt-1hop / dim-2hop / datum-2hop / datum_target-1hop / none` 各多少条 |
| ④ 疑似对应 | 关联失败条目的模糊匹配候选，**仅供参考，不进任何指标** |

## 结构变化不会静默失配

数据侧的演进是无限的（SFA 换版本、换导出勾选、不同测试件、开发侧生成器改版），
所以不能只靠"把已知问题修掉"。本工具用三层把「静默失配」变成「显式告警」：

**① 列名驱动装载** — 表头列名其实就在表里（`ta` 表在 r3、其余表在 r2，带换行与
`(Sec. x)` 后缀），一律按列名定位。识别不到表头才退回默认列号，并把该表标为不可信。

**② 加载期交叉校验** — 不依赖列语义，只看「读到没有、数量对不对、引用能否落地」：

- 表内数据行数 vs 装载条数（整表被跳过时直接报出，比结果表全红更早）
- `ta` 引用中「既不在语义表、也不在 dcr / datum 索引」的比例
- `dc` 的 name 与 `ta` 的 name 对应率

**③ 一键诊断包** — 结果不对劲时跑一次，把完整上下文写成一份文件：

```bash
python doctor.py <SFA报告.xlsx> [开发markdown]
# 同时打印并写入 doctor_report.txt
```

输出包含：环境与依赖版本、全部工作表、每张关键表的原始前 5 行与列定位结果、
ID 索引规模、交叉校验逐项结果、装载告警、**未收录字符清单**、markdown 解析自检、
关联路径分布、三项指标、缺陷清单、疑似对应、全部非命中条目。

> 「未收录字符」这一段专门用来看 SFA 有没有引入新的字形。但要区分两类：
> 真公差符号（如 `⌭` 圆柱度、`⌯` 对称度、`−` 直线度）要补进符号表；
> **SFA 的排版字形**（`▽` `⎹` `⭩` `◁` `⌮` `◎`）**不能**补 —— 它们在带基准的
> 框架里与 `[A]` 同行，属版面元素，收进符号表会让「缺符号」检查误报一片。

**遗留的两类无法根治，只能保证快速发现**：新 GD&T 符号与渲染字形（白名单机制）、
开发侧 markdown 格式演进（解析自检会显式报出，不再静默全红）。

## 内容比对的分段口径

一条 `Semantic PMI Summary` 记录可能把**图纸上多个标注**的文本揉在一行：

```
3X ⌀3.50 ± 0.2        ← 尺寸标注
⌭ | 0.1               ← 圆柱度框架
```

因此取文本时必须按片段取（GT 取含 `|` 的 FCF 行，DIM 合并折行），
直接拿整行会导致「缺数值 / 数量前缀不符」大面积误报。比较用的指纹为
`数值集合 + 基准字母 + 修饰符 + 数量前缀`，其中：

- 数量前缀（`2X`）整体剔除后再取数值与字母，否则 `X` 会被当成基准字母；
- 修饰符优先按字形（`Ⓜ Ⓛ Ⓟ Ⓤ Ⓢ`）识别，因为字形紧跟字母，纯字母启发式会漏；
- 基准字母与修饰符同形时（`L` 既是基准又可能是 LMC）在判定阶段**对两侧对称剔除**，
  只损失区分度、不制造假冲突。

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

## 界面视图

结果表默认只出 **8 列**（判定 + 比对内容 + 结论），定位/溯源字段默认收起：

| 分类 | 字段 |
|---|---|
| 默认显示（8 列） | 匹配状态 · 开发名称 · 开发标注 · SFA语义文本 · SFA图形文本 · 提取缺陷 · 人工校验 · 备注 |
| 默认隐藏（9 列，侧边栏「3. 表格字段 → 附加字段」按需开启） | 关联键 · 分组 · Handle · SFA语义ID · SFA实体类型 · 层级 · 类别 · 缺陷详情 · 关联路径 |

其他约定：

- 除「人工校验」「备注」外全部只读，防止误改真值文本。
- 导出 Excel 始终是**全字段**，不受界面列显示影响。
- Streamlit 的 `column_config` 没有"默认隐藏列"能力，隐藏靠不放进 DataFrame 实现，
  因此表格 `key` 带列指纹，避免切列时行内编辑状态串位。

## 文件结构

```
pmi_core.py            比对内核：真值装载 / markdown 解析 / 归一化 / ID 关联 / 指标 / 缺陷检测
pmi_compare_work.py    Streamlit 界面（当前主入口）
doctor.py              一键诊断：结果不对劲时跑一次，产出完整上下文报告
test_pmi_core.py       内核回归（161 项：指标口径、关联链路、缺陷检测、解析自检、列名驱动）
test_ui_smoke.py       界面冒烟（29 项：AppTest 无头跑渲染分支 + 列口径 + 两个诊断面板）
test_realdata.py       真实数据回归（锁端到端数值，语料缺失自动跳过）
samples/               真实语料目录（不进版本控制，见 samples/README.md）
streamlit_launcher.py  启动器（规避 IDE 运行配置序列化差异）
pmi_compare.py         早期版本，留档
```

## 测试

```bash
python test_pmi_core.py          # 内核回归：口径、链路、缺陷、解析自检、列名驱动与交叉校验
python test_ui_smoke.py          # 界面冒烟：视图列口径、筛选、两个诊断面板
python test_ui_smoke.py <xlsx>   # 也可指定真值报告路径
python test_realdata.py          # 真实数据回归（需 samples/ 与 SFA 报告，缺失自动跳过）
python doctor.py <xlsx> [md]     # 出问题时的一键诊断（不是测试，是排查工具）
```

> `test_pmi_core.py` 的列名驱动用例是**现场用 openpyxl 构造**的报告（列序全部打乱、
> 表头带换行与 `(Sec. x)`），不依赖外部数据文件，可离线跑。

## 环境

Python 3.13 / Streamlit 1.64 / pandas / openpyxl / xlrd

## 备注

仓库不包含 SFA 安装包、NIST 标准文档与测试用 xlsx（受版权与体积限制，见 `.gitignore`）。
