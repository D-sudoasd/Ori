<p align="center">
  <img src="assets/readme/hero.png" width="100%" alt="Ori — Import research tables into editable Origin projects / 将科研表格导入可编辑的 Origin 工程. Conceptual illustration / 概念插图。">
</p>

# Ori

**Import research tables into editable Origin projects**

**将科研表格导入可编辑的 Origin 工程**

[Overview / 项目概览](#overview--项目概览) · [Start / 开始使用](#start--开始使用) · [Reference / 详细说明](#reference--详细说明)

## Overview / 项目概览

Ori imports text tables, Excel worksheets and flat JSON data into Origin projects or Excel workbooks. Inspect column types, review an import plan and retain source information with the exported tables.

Ori 将文本表格、Excel 工作表和扁平 JSON 数据导入 Origin 工程或 Excel 工作簿。先检查列类型和导入计划，再将来源信息与数据表一起导出。

- **Typed tables** — 支持数值、文本、日期时间和缺失值。
- **Editable project output** — 映射 X/Y 和误差列，生成可继续编辑的 OPJU 工程。
- **GUI, CLI and optional MCP** — 图形界面与结构化命令共用检查、计划、验证和执行流程。

## Start / 开始使用

From the checkout, install dependencies and launch the desktop tool. See [Agent guide / Agent 指南](docs/agent-guide.md) and [Example tables / 示例表格](examples_general).

在仓库目录安装并启动：

```powershell
py -3 -m pip install -r requirements.txt
py -3 spectra_to_origin.py
```

For read-only inspection / 只读检查：

```powershell
py -3 -m origin_bridge inspect -i ./examples_general
```

OPJU requires Windows and a working OriginPro installation. XLSX export does not require Origin and creates data tables, not graphs. Existing Origin sessions are protected.

OPJU 导出需要 Windows 与可用的 OriginPro；XLSX 导出不依赖 Origin，只生成数据表。程序保护已有 Origin 会话。

*Cover: AI-generated conceptual illustration. 封面为 AI 生成的概念插图。*

## Reference / 详细说明

把常见的实验、监测和工程表格导入 Origin，生成可继续编辑的 `.opju` 工程，也可以只生成 Excel。项目现已支持多列表格、Excel 多工作表和扁平 JSON 数据；原有的 XRD／光谱批量分组流程仍可使用。

程序可由人通过窗口操作，也提供 JSON 命令行和可选 MCP 服务供 agent 检查数据、生成和校验导入计划，再执行导入。类型和图形建议由本地规则产生，不需要在线大模型；导入计划可修改列映射、目标工作表名称和图形设置。

支持 TXT、DAT、XY、CSV、TSV、XLSX、XLSM、XLS、JSON、JSONL 和 NDJSON。可导入数值、文本、日期时间和缺失值。数据不会插值、拟合、平滑、基线扣除、单位换算或归一化；带时区的日期时间会先转换为 UTC，再按不带时区的值写入目标。

## 开始使用

Windows 用户从源码运行（Python 3.10+）：

```powershell
py -3 -m pip install -r requirements.txt
py -3 spectra_to_origin.py
```

不带参数启动新版通用数据窗口。也可以先把文件或文件夹作为参数传入，或直接打开旧谱线窗口：

```powershell
py -3 spectra_to_origin.py .\examples_general
py -3 spectra_to_origin.py --spectra-gui .\examples
```

`.opju` 导出需要 Windows 上已安装并可用的 OriginPro，以及当前 Python 环境中的 OriginPro 接口：

```powershell
py -3 -m pip install originpro
```

程序会检查并拒绝在已有 Origin 会话运行时创建工程，以免重置当前工作。导出 XLSX 不需要 Origin；XLSX 只生成数据表，绘图设置保存在来源信息中，图形由 OPJU 导出生成。

## 通用数据导入

新版窗口可以添加文件或文件夹，查看表名、列类型和数据预览，再选择 X、一个或多个 Y、绘图方式及输出位置。Excel 的每个非空工作表会作为单独数据表列出；窗口可设置表头、跳过行数和分隔符。命令行和导入计划可以指定 Excel 工作表。

添加文件、文件夹或拖放时，展开在后台进行，来源列表按帧刷新。默认是**每个来源一个文件**：窗口只发现路径，预览选中的来源，不把全部文件先读进内存。合并为一个工程时才逐个检查全部来源。检查结果只保留摘要（列结构、少量样例和图形建议），并用文件内容 SHA-256 加读取选项决定是否重读；文件大小或修改时间相同但内容变了也会重读。某个来源读失败时，列表会标出错误，其他来源仍可按文件导出。合并为一个工程时，任一来源失败或尚未读完会拒绝保存计划和导出。

### 每个来源一个文件

选择输出目录后，每个输入文件生成一个 OPJU 或 XLSX。同一个 Excel 工作簿的非空工作表进入这一个文件。可选为每个图再写 PDF（仅 OPJU），以及覆盖已有输出。绘图设置作用于每一个文件、每一张表：自动使用各表自己的建议；明确写出的列名或列号如果对不上，该文件失败，不会改用预览里的其他列。全数字按从 0 开始的列号。误差列需要 Y 使用原列名：有误差列时不能把 Y 写成列号，窗口会在启动前说明，不会等整批失败后再解释，也不会拿预览里的列名去替换。没有误差列时，Y 仍可以用列号。

窗口每次新建批次都会另写一对 `batch-record-<id>.json` 和 `batch-config-<id>.json`，并把这条记录路径交给核心。同一输出目录里的下一次批次使用新的名字，不会覆盖上一批的记录、配置、`.bak`、`.previous` 或日志。每个来源的 OPJU/XLSX 仍按源路径生成稳定文件名。旧的 `batch-record.json` 加 `batch-config.json` 仍可从新窗口继续。只拿到命令行留下的运行记录时，窗口会说明缺配置，并让你指出原来的配置，不会按当前界面默认值重导。命令行不传 `--record` 时，默认记录仍是该目录的 `batch-record.json`。重试仍把全部来源交给核心，不在窗口里提前丢掉已成功的文件。已成功的 OPJU 会保留。PDF 没写成时，任务可以是成功且错误为空，具体的图和原因在 `pdf_errors` 里。不要为了补 PDF 去勾选“覆盖已有输出”。批次结束时若核心带回恢复或损坏等提示，窗口会给出条数，并可在“查看完整详情”里看全文。

“停止后续”在当前这项完成后不再开始后面的任务，已写好的文件保留。它不能强行打断卡在 Origin 启动证明之前的调用。关闭窗口前应先停止并等待结束；直接销毁窗口会先请求取消，再收尾，不会立刻杀掉正在写的进程。

合并为一个工程仍使用原来的单次导出、计划 JSON 和“完成后保持打开”。PDF、继续和重试只出现在每个来源一个文件的模式里。旧的两列谱线窗口和命令行保留。

表格支持数值、文本、日期时间和缺失单元格。TXT 等文本支持逗号、分号、Tab 或空白分隔符；常见 UTF 编码和 GB 系列中文编码会自动识别。JSON 支持对象记录数组、列数组对象和 `{ "columns": [...], "data": [...] }` 矩形表；JSONL／NDJSON 每行须为字段一致的对象记录。嵌套 JSON 对象和数组不会被压平成文本。

绘图可选不绘图、折线、散点、折线加符号或柱状图；窗口支持为单个 Y 指定误差列，导入计划可分别映射多个 Y 误差列。导入时保留表中的全部列；选择 X/Y 只决定图形，不会丢弃其他列。Origin 工程和 Excel 工作簿都附带来源信息工作表，记录源路径与哈希、工作表、读入选项、列信息和图形设置。自动建议便于起步，但遇到不明确的列头或列类型时，应先检查预览并修正导入设置。

数值和日期转换服从 Origin／Excel 的单元格格式；文本数据写入为文本列。对于 JSON、TXT 等输入中的日期时间，支持 ISO 8601 格式。XLSX 默认读取公式的缓存值；若没有缓存结果会报错，可将 `formula_policy` 设为 `text` 导入公式文本。

## Agent 与命令行

新通用 CLI 每条命令均输出 UTF-8 JSON，适合脚本和 agent 自动化。Windows PowerShell 示例：

```powershell
# 检查文件、工作表、列类型、缺失值和图形建议；不会写文件或启动 Origin
py -3 -m origin_bridge inspect -i .\examples_general

# 生成计划；自动识别输出格式，另存可审阅的计划 JSON
py -3 -m origin_bridge plan -i .\examples_general\stress_strain.tsv `
  -o .\out\stress.opju --save .\out\stress-plan.json

# 校验计划后执行
py -3 -m origin_bridge validate .\out\stress-plan.json
py -3 -m origin_bridge execute .\out\stress-plan.json

# 只要表格，不需要 Origin 或图形
py -3 -m origin_bridge import -i .\examples_general\categories.json `
  -o .\out\categories.xlsx --plot none
```

计划默认不覆盖已有文件；如要替换，需要明确传入 `--overwrite`。绘图可用 `--plot auto|none|line|scatter|line_symbol|column`，并用 `--x 列名` 和 `--y 列名 ...` 选择图形列。误差列映射在导入计划中设置。可用 `--header auto|yes|no`、`--skip-rows N`、`--delimiter auto|whitespace|tab|,|;` 和 `--sheet 工作表名` 处理来源格式。Excel 省略 `--sheet` 时读取所有非空工作表。运行 `py -3 -m origin_bridge --help` 或在子命令后添加 `--help` 查看完整选项。

`batch` 子命令为每个输入文件写一个输出。不传 `--record` 时，续跑记录是输出目录的 `batch-record.json`。窗口新建批次则会传入独立的 `batch-record-<id>.json`。`--resume` 只在源内容、读取设置、绘图设置、输出路径和输出字节都未变时跳过；文件存在或修改时间相同不会单独构成跳过条件。签名、JSON 示例、取消、重试和 PDF 约定见 [docs/batch-api.md](docs/batch-api.md)。

需要 agent 通过 MCP 直接操作时，安装可选依赖并在 MCP 客户端中将项目目录设为工作目录、启动 `python -m origin_bridge.mcp_server`：

```powershell
py -3 -m pip install -r requirements-agent.txt
py -3 -m origin_bridge.mcp_server
```

服务提供 `inspect_data`、`create_import_plan`、`validate_import_plan` 和 `execute_import_plan`。MCP 计划的默认格式为 `opju`，如要生成 XLSX 请明确传入 `output_format="xlsx"`。执行工具需要显式设置 `confirm=true`；`.opju` 仍由同一台 Windows 机器上的 Origin 生成。配置细节、计划格式和完整 agent 工作流见 [docs/agent-guide.md](docs/agent-guide.md) 与 [docs/import-plan.schema.json](docs/import-plan.schema.json)。

## 支持范围

该工具面向矩形、扁平表格。它不解析 PDF、图片、任意专有二进制文件或多层嵌套对象。对列名和类型的建议是启发式结果，不代表领域判断；agent 可检查样例和类型，用户可编辑计划来修正映射。不会自动推断或改写物理单位，也不会据数据内容生成分析结论。

## 旧版 XRD／光谱批量导入

旧模式针对每个文件一条两列数值谱线。可按文件名分组、指定组数均分或将所有曲线放入一组；同组 X 网格一致时使用 XYYY，否则使用 XYXY。自动布局按每组独立判断，不以近似容差合并不同网格。

打开旧窗口：

```powershell
py -3 spectra_to_origin.py --spectra-gui .\examples
```

兼容的旧 CLI：

```powershell
# 校验输入
py -3 spectra_to_origin.py --check -i .\examples --group-by-name

# 生成 Origin 工程
py -3 spectra_to_origin.py --cli -i .\examples -o .\out\spectra.opju --layout auto --group-by-name

# 只导出 Excel 和 CSV
py -3 spectra_to_origin.py --cli --xlsx-only -i .\examples -o .\out\spectra.xlsx --n-groups 2
```

旧模式文本支持 TXT、DAT、XY、CSV、TSV，要求两列数据、至少两个点；支持空白、Tab、逗号、分号、常见 UTF 和 GB 系编码、科学计数法及 Fortran `D` 指数。可以设置 X/Y 名称和单位；不会插值、拟合、平滑或归一化。旧模式参数可通过 `py -3 spectra_to_origin.py --help` 查看。

## Windows 程序

源码构建：

```powershell
py -3 -m pip install -r requirements-build.txt
py -3 -m PyInstaller --noconfirm --clean SpectraToOrigin.spec
```

生成 `dist\SpectraToOrigin.exe`（通用 GUI，也可用 `--spectra-gui` 打开旧窗口）和 `dist\DataToOriginCLI.exe`（命令行 JSON 接口）。仓库 [examples_general](examples_general) 包含应力应变 TSV、分类性能 JSON 和监测 JSONL 演示数据；旧版两列谱线数据在 [examples](examples)。

运行测试：

```powershell
py -3 -m unittest discover -v
```

真实 Origin 集成测试需要已安装 OriginPro，并通过 `SPECTRA_TEST_ORIGIN=1` 显式启用。测试只检查导入、保存和可回读数据，不代表任何材料或实验结论。

MIT License，见 [LICENSE](LICENSE)。
