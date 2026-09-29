# 快递账单核对skill

帮助 Codex 核对快递账单，找出应向仓库核重量、向快递核收费的异常，并提供可追溯证据。

这是供 Codex 使用的工作规则与辅助脚本，不是一键完成核算的独立软件。公开仓库只含规则、结构示例和预检脚本，不含真实账单、订单明细或具体报价金额。

## 准备三份文件

- 快递公司账单：单号、重量、收费、日期、目的地。
- 聚水潭等系统账单：单号、重量、商品编码、件数、发货日期、地区。
- 运费价格表：生效日期、地区、重量档、报价及附加费规则。

字段按实际表头识别，示例布局仅供参考。聚水潭原“运费”不参与核算，始终用系统重量和提供的价格表重算。

## 核对规则

核对单号、重量、价格和价格表规则四项。同档且不影响计价的重量小差忽略；日期、地区或报价规则不明确时保留待确认，不猜价格。

金额分别为“快递账单收费”“按快递重量计算”“按系统重量计算”。总差额是账单收费减去按系统重量计算的金额，再拆成重量影响和计价差异。每票只有一个主原因，避免重复计数；异常表示需要复核，不直接认定责任或可追回款。

## 两种输出

| 用途 | 内容 |
|---|---|
| 内部排查 | 商品异常汇总、运单明细、核对依据。按商品看异常单量、差额、原因及先找仓库还是快递。 |
| 给快递的多收费举证 | 在快递原账单副本新增总览和分类明细，只统计依据明确且总差额为正的运单，不用其他运单少收抵扣；附完整原价格表和系统账单。 |

举证明细按“计价不符、重量影响、重量＋计价”分类。单号、价格依据及系统重量、地址等可点击核查对应源行；系统源链接须核对单号一致。原系统运费可随原页保留，但不参与核价。输入文件另存保留，生成表格不自动向快递发送。

## 安装到 Codex

下载仓库后，将含 `SKILL.md` 的整个目录命名为 `express-bill-reconciliation`，复制到 Codex 的个人技能目录。本说明采用当前已使用的本地目录复制方式：

- Windows：`%USERPROFILE%\.codex\skills\express-bill-reconciliation\SKILL.md`
- macOS / Linux 默认目录形式：`~/.codex/skills/express-bill-reconciliation/SKILL.md`
- 已设置 `CODEX_HOME` 时：`<CODEX_HOME>/skills/express-bill-reconciliation/SKILL.md`

不要多套一层目录。安装后在新会话中确认技能可被识别；若所用客户端版本使用不同的技能目录，请按其设置或 [OpenAI 官方 Skills 文档](https://developers.openai.com/codex/skills) 放置。

附上三份文件，再发送：

```text
请使用 $express-bill-reconciliation 核对这三份文件，生成按商品汇总的内部异常报告。
```

需要快递举证表时：

```text
请使用 $express-bill-reconciliation 生成只统计多收费的快递举证表，
在快递原账单副本新增分类 sheet，附完整价格表和系统账单，保留源行跳转。
```

## 可选：运行预检

预检脚本使用 Python 3 和 `openpyxl`。在本目录安装依赖、查看参数：

```sh
python -m pip install -r requirements.txt
python -X utf8 scripts/inspect_inputs.py --help
```

```sh
python -X utf8 scripts/inspect_inputs.py --courier "快递公司账单.xlsx" --system "聚水潭系统账单.xlsx" --price "价格表.xlsx" --output "outputs/input_profile.json"
```

脚本只读检查文件结构、单号、重量和公式缓存，输出预检 JSON；它不计算完整运费，也不生成最终异常工作簿。完整核对及报告由 Codex 按 `SKILL.md` 执行，需要当前环境具备读写 Excel 的工具。预检 JSON 可能包含单号示例、输入路径和完整价格单元格，请保留在本地。
