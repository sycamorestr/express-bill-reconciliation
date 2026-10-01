# 财务快递账单对账 Skill

面向财务部门，核对快递账单、重量和运费，按商品汇总异常单量、金额、原因和核查对象，再生成逐单证据。**1.1.0 使用固定计算引擎：相同输入、相同配置、相同引擎版本，得到相同核对结论。**

公开仓库包含规则、引擎、配置、测试和使用说明，不含真实账单、订单明细或具体业务报价。引擎生成可复现的 JSON 和 CSV；Codex 据此制作方便仓库、快递查阅的 Excel 工作簿。

## 准备三份文件

| 文件 | 所需信息 |
|---|---|
| 快递公司账单 | 单号、重量、收费、日期、目的地 |
| 聚水潭等系统账单 | 单号、重量、商品编码、件数、发货日期、地区 |
| 运费价格表 | 生效日期、地区、重量档、报价及明确的计价条件 |

当前支持 `.xlsx`、可配置的工作表与表头别名，以及与示例同类的“日期分块、地区行 × 重量档列”价格表。默认配置适用于圆通/聚水潭示例结构，实际报价每次从本次文件读取。承运商按完整名称匹配“圆通速递/圆通快递”，允许明确网点前缀和 `_数字` 后缀；同品牌网点只在范围检查时归并，原名称全部保留。其他品牌不会套用该配置核价。它不会自动理解所有承运商的任意报价模式；结构、字段或合同口径不同时，须先明确配置及引擎支持范围。

聚水潭原“运费”不参与核算，始终使用系统重量和提供的价格表重算。规则不足的运单保留待核，不能为了得出金额而猜测。

## 结果怎么看

三个金额是“快递账单收费”“按快递重量计算”“按系统重量计算”。总差额为账单收费减去按系统重量计算的金额，再拆成重量影响与计价差异。

同档且不影响计价的重量小差忽略。每票只有一个主原因，待核依据优先；日期、地区或报价不明确时不强行归因。大于 3kg 的相邻档端点同价时，虽可得到唯一参考金额，仍保留“报价规则待确认”，不列入快递多收费举证。异常表示需要复核，不直接认定责任或可追回款。

| 用途 | Excel 内容 |
|---|---|
| 内部排查 | 商品异常汇总、运单明细、核对依据。按商品看异常单量、差额、原因及先找仓库还是快递。 |
| 给快递的多收费举证 | 在快递原账单副本新增总览和分类明细，只统计依据明确且总差额为正的运单，不用少收抵扣；附完整原价格表和系统账单。 |

举证明细按“计价不符、重量影响、重量＋计价”分类。单号、价表及系统重量、地址等可跳转到对应源行。原系统运费可随原页保留，但不参与计算。源文件不覆盖，报告不自动向快递发送。

## 安装到 Codex

仓库名称统一为 `finance-express-bill-reconciliation-skill`；`express-bill-reconciliation` 是现有调用的兼容标识，新仓库名不替代安装目录名。

建议下载固定版本 [v1.1.0](https://github.com/sycamorestr/finance-express-bill-reconciliation-skill/tree/v1.1.0)，将含 `SKILL.md` 的整个目录命名为 `express-bill-reconciliation`，复制到 Codex 的个人技能目录。本说明采用当前已使用的本地目录复制方式：

- Windows：`%USERPROFILE%\.codex\skills\express-bill-reconciliation\SKILL.md`
- macOS / Linux 默认目录形式：`~/.codex/skills/express-bill-reconciliation/SKILL.md`
- 已设置 `CODEX_HOME` 时：`<CODEX_HOME>/skills/express-bill-reconciliation/SKILL.md`

不要多套一层目录。安装后在新会话中确认技能可被识别；若所用客户端版本使用不同的技能目录，请按其设置或 [OpenAI 官方 Skills 文档](https://developers.openai.com/codex/skills) 放置。

使用 Python **3.10 或更新版本**，在技能目录安装固定依赖并运行自检：

```sh
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

附上三份文件，再发送：

```text
请使用 $express-bill-reconciliation 核对这三份文件，
使用固定引擎并保留本次配置和复现记录，生成按商品汇总的内部异常报告。
```

需要快递举证表时：

```text
请使用 $express-bill-reconciliation 生成只统计多收费的快递举证表，
在快递原账单副本新增分类 sheet，附完整价格表和系统账单，保留源行跳转。
```

## 直接运行计算引擎

从技能目录执行，把路径换为当次文件，每次使用新的空输出目录：

```sh
python -X utf8 scripts/reconcile.py --courier "快递公司账单.xlsx" --system "聚水潭系统账单.xlsx" --price "价格表.xlsx" --profile "profiles/yuantong-jushuitan-v1.json" --output-dir "outputs/run-001"
```

省略 `--profile` 时使用默认配置。可先运行 `python -X utf8 scripts/reconcile.py --help` 查看命令参数。

| 产物 | 用途 |
|---|---|
| `result.json` | 权威核对结果；金额及重量使用十进制字符串，避免浮点差异 |
| `profile-resolved.json` | 本次实际生效的字段映射、规则及配置版本 |
| `manifest.json` | 输入、配置、代码及结果的 SHA256，版本、运行环境与验证记录 |
| CSV 文件 | 便于查阅和筛选结果，不能替代权威 JSON |

引擎不直接完成所有 Excel 展示工作。Codex 按 `references/output-contract.md` 将计算结果排版为内部报告或快递举证表；不得在排版时重新核价、重新分配主原因或手改金额。输入及运行产物可能包含运单和报价，应由使用者保存在自己的工作目录。

## 换一批数据或修改规则

字段别名、工作表等受支持映射改变时，复制默认 JSON profile，修改对应配置并更新 `config_revision`（默认字符串 `"1"`，可改为其他非空文本），使用 `--profile` 指向它。`profile_version="1.1.0"` 和 `profile_id="yuantong-jushuitan-v1"` 是固定的算法兼容标识，不能随映射修订更改。实际生效配置、修订号和配置 SHA256 随每次结果保存。新的合同口径或计算模式若未被引擎支持，须先补齐实现和相应测试，再发布新引擎版本；不能只改固定策略字符串或把自然语言备注当可执行规则。

复现时使用相同输入文件、保存的配置和同一引擎版本，对照 `manifest.json` 的哈希、验证结果以及 `result.json` 的业务数据。固定版本使用同一 Git 标签；主分支持续更新，不能代替版本锁定。可用 `--compare` 自动检查完整结果：

```sh
python -X utf8 scripts/reconcile.py --courier "快递公司账单.xlsx" --system "聚水潭系统账单.xlsx" --price "价格表.xlsx" --profile "outputs/run-001/profile-resolved.json" --output-dir "outputs/run-002" --compare "outputs/run-001/result.json"
```

结果不同会报错并退出，不交付新一轮结果文件。统计、逐单金额、主原因、待核项及源依据应一致；输出目录、生成时间、Excel 容器字节及排版差异不影响语义一致性。具体边界与固定判断顺序见 [复现规则](references/reproducibility.md)。

可选预检：

```sh
python -X utf8 scripts/inspect_inputs.py --courier "快递公司账单.xlsx" --system "聚水潭系统账单.xlsx" --price "价格表.xlsx" --output "outputs/input_profile.json"
```

预检只读检查结构、单号、重量和公式缓存；完整核算必须运行 `reconcile.py`。预检 JSON 也可能包含源文件信息和价格单元格。
