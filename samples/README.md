# samples/ —— 真实数据回归语料

本目录存放**真实数据**回归语料，用于在改动比对逻辑后自动发现回归。
语料本身**不进版本控制**（体积大、含内部数据），仓库里只保留本说明。

## 放置方式

1. 开发侧导出的 markdown 放这里，命名 `dev_<报告名>.md`，例如 `dev_ftc_10.md`。
2. SFA 报告 xlsx **不必**复制进来 —— 用环境变量指向它所在目录即可（默认已是 NIST 目录）。

```bash
export PMI_SFA_DIR="F:/1【机械零件】/Step官方标准数据/NIST-PMI-STEP-Files"
python test_realdata.py
```

PowerShell：

```powershell
$env:PMI_SFA_DIR = "F:\1【机械零件】\Step官方标准数据\NIST-PMI-STEP-Files"
python test_realdata.py
```

## 行为

- 语料或 xlsx 缺失时**自动跳过**（退出码 0），不会阻塞。
- 有数据时逐条断言期望指标（召回 / 精确 / 基准 / 缺陷条数 / 缺陷分布 /
  无断链条目 / 加载期交叉校验全通过）。数值变了就说明比对逻辑被改坏了。

## 新增一条语料

在 `test_realdata.py` 的 `CASES` 里加一项：

```python
{
    "md": "samples/dev_ftc_12.md",              # 相对项目根目录
    "xlsx": "nist_ftc_12_asme1_ap242-e2-sfa.xlsx",
    "recall": 100.0, "precision": 100.0, "datum": 100.0,
    "defects": 3,
    "defect_codes": {"SYM 符号丢失": 3},
}
```
