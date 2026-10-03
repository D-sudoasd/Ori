# Agent 导入指南

`origin_bridge` 提供面向脚本和 agent 的本地表格检查、计划校验及导出接口。它按明确的数据类型、列映射和绘图设置工作；计划不包含可执行代码，也不会调用在线模型推断实验结论。

## 安装与运行

常规 CLI 不需要 MCP SDK：

```powershell
py -3 -m pip install -r requirements.txt
py -3 -m origin_bridge --help
```

导出 `.opju` 需要 Windows 上已安装 OriginPro，并在运行服务或 CLI 的同一个 Python 环境安装 `originpro`：

```powershell
py -3 -m pip install originpro
```

导出 `.xlsx` 不需要 Origin。导出 `.opju` 时，如果检测到 Origin 已运行，程序会停止，避免重置现有工程。

### MCP 服务

MCP 服务使用标准 stdio transport。安装可选 SDK 并启动：

```powershell
py -3 -m pip install -r requirements-agent.txt
py -3 -m origin_bridge.mcp_server
```

在 MCP 客户端中将 `command` 设为当前 Python 解释器，将 `args` 设为 `['-m', 'origin_bridge.mcp_server']`，并将工作目录设为本项目目录。客户端和服务运行的机器必须能访问源文件；Origin 工程也由服务所在的 Windows 电脑创建。工具接受本地文件路径，建议传入绝对路径。

服务注册以下四个工具：

| 工具 | 行为 |
| --- | --- |
| `inspect_data(paths, options?)` | 读取文件并返回工作表、列类型、缺失数、样例、范围、警告和图形建议。不写文件，也不启动 Origin。 |
| `create_import_plan(paths, output_path, output_format?, options?, overwrite?, keep_open?)` | 创建版本化导入计划，不写输出，也不启动 Origin。`output_format` 默认是 `opju`；默认不覆盖已有文件。 |
| `validate_import_plan(plan, base_dir?)` | 重新读取源数据，核对 SHA-256、读入选项、列映射、图形设置和输出路径；不写输出，也不启动 Origin。 |
| `execute_import_plan(plan, confirm, base_dir?)` | 再次校验并执行。必须显式传入 `confirm=true`；只会在计划 `overwrite=true` 时替换已有文件。 |

服务在本机解析源文件并写入输出；源路径、列信息和样例会通过 MCP 返回给客户端。客户端如何保留或处理这些返回数据，取决于所用 MCP 客户端的策略。源路径须为服务进程可访问的本地路径。

## 建议的 agent 流程

1. 调用 `inspect_data` 检查所有输入文件。确认自动识别的列名、类型、样例、工作表以及缺失值处理符合来源数据。
2. 调用 `create_import_plan` 生成计划。若默认图形或命名不合适，直接修改返回的计划对象；也可通过 CLI 的 `plan --save` 写成 JSON 后编辑。
3. 调用 `validate_import_plan` 校验最终计划。若源文件在检查后发生变化，必须重新检查并生成计划；不要忽略哈希不匹配。
4. 用户确认输出路径和图形后，将同一份计划传入 `execute_import_plan`，并设置 `confirm=true`。记录返回的输出路径和 SHA-256。

`inspect_data` 对自动表头和绘图的建议是规则推断。遇到全文本列、混合类型或没有明确的 X 列时，先检查返回的样例和警告，再确定表头与映射。所有源列都会保留；X/Y 选择只控制绘图。

## CLI 示例

每条 CLI 命令输出 UTF-8 JSON。Windows PowerShell：

```powershell
# 递归检查目录下所有支持的文件
py -3 -m origin_bridge inspect -i .\examples_general

# 创建可审阅的 OPJU 计划
py -3 -m origin_bridge plan `
  -i .\examples_general\stress_strain.tsv `
  -o .\out\stress.opju `
  --save .\out\stress-plan.json

# 校验后执行
py -3 -m origin_bridge validate .\out\stress-plan.json
py -3 -m origin_bridge execute .\out\stress-plan.json

# 立即生成无图形的 XLSX 表格
py -3 -m origin_bridge import `
  -i .\examples_general\monitoring.jsonl `
  -o .\out\monitoring.xlsx `
  --plot none
```

按文件批量导出使用 `batch`。每个输入文件生成一个 OPJU 或 XLSX；同一 Excel 工作簿的非空工作表进入同一个输出。默认不覆盖已有文件。命令行未指定记录时用 `batch-record.json`。通用数据窗口的“每个来源一个文件”调用同一套请求，但每次新建批次写入独立的 `batch-record-<id>.json` 和 `batch-config-<id>.json`，避免同目录的下一批覆盖上一批的恢复文件；旧文件名仍可加载。核心回执本身不含完整全局请求。完整请求、结果、进度、取消、重试和续跑约定见 [batch-api.md](batch-api.md)。启动 Origin 前若 COM 卡住，还没有归属证明，不能结束未知的 Origin 进程。

```powershell
py -3 -m origin_bridge batch -i .\examples_general -o .\out\batch --format xlsx --plot none
py -3 -m origin_bridge batch -i .\examples_general -o .\out\batch --resume
```

计划命令只创建 JSON 计划，`import` 命令直接执行。CLI 的 `plan --save` 不覆盖已有计划文件；导入目标默认不覆盖，只有指定 `--overwrite` 才允许替换。计划文件中的相对源路径由 `validate` 和 `execute` 按计划文件所在目录解析；自动生成的计划通常记录源文件绝对路径。

CLI 可通过 `--plot auto|none|line|scatter|line_symbol|column` 指定图形，通过 `--x 列名` 和 `--y 列名 ...` 选择列。Y 误差列需在保存的计划 JSON 中设置 `plot.y_error`，或直接编辑 MCP 返回的计划对象，然后调用 `validate_import_plan`。

来源读取参数包括：

| 选项 | 说明 |
| --- | --- |
| `--header auto\|yes\|no` | 自动判断、将首行作为列名或将首行保留为数据。JSON 字段名属于数据结构，不能关闭表头。 |
| `--skip-rows N` | 在读取表格前跳过的物理行或工作表行数。 |
| `--delimiter auto\|whitespace\|tab\|,\|;` | 文本表格的分隔符。默认自动检测。 |
| `--sheet NAME` | 只读指定 Excel 工作表；省略时读取所有非空表。 |
| `--encoding NAME` | 文本编码；默认自动识别。 |
| `--formula-policy cached\|text` | Excel 公式默认读缓存结果；无缓存时可用 `text` 保留公式文本。 |
| `--missing-value TEXT` | 增加缺失值标记；默认只把空单元格或空字段识别为缺失。可重复指定。 |

列名中的 `name (unit)` 或 `name [unit]` 会拆分成列名和单位。类型推断只用于导入：有限数字会成为数值，ISO 日期时间会成为日期时间，其他内容作为文本；前导零标识符会保留为文本。数据行不完整、非有限数值或表格列数不齐时会报错，不会删除行或插补缺失值。

## 导入计划格式

计划的规范见 [import-plan.schema.json](import-plan.schema.json)。根对象包含 `schema_version`、`output` 和 `tables`。验证器会拒绝未知字段、无效类型、未找到的列、列角色冲突、输入输出路径相同、已有目标未允许覆盖，以及与计划记录哈希不同的源文件。

每个 table 项由源文件、目标工作表名、绘图和可选列标签组成：

```json
{
  "schema_version": 1,
  "output": {
    "path": "C:/data/stress.opju",
    "format": "opju",
    "overwrite": false,
    "keep_open": false
  },
  "tables": [
    {
      "source": {
        "path": "C:/data/stress_strain.tsv",
        "sha256": "由 inspect 或 create_import_plan 返回的 64 位哈希",
        "options": {
          "sheet": null,
          "header": true,
          "skip_rows": 0,
          "delimiter": "\t",
          "encoding": "utf-8-sig",
          "missing_values": [""],
          "formula_policy": "cached"
        }
      },
      "name": "stress_strain",
      "plot": {
        "kind": "line",
        "x": "Strain",
        "y": ["Stress"],
        "y_error": {"Stress": "Error"},
        "title": "Stress–strain",
        "x_label": "Strain (%)",
        "y_label": "Stress (MPa)"
      },
      "column_labels": {}
    }
  ]
}
```

示例中的哈希和读入选项必须原样取自检查或计划结果；不要手工编造。具体规则：

- `source.path` 指向一份原始文件；`sha256` 用来确保执行时文件与检查时相同；`options` 记录完整且已解析的读入设置。
- `name` 是目标工作表名，在一个工程内必须唯一。
- `plot.kind` 可选 `none`、`line`、`scatter`、`line_symbol`、`column`。`none` 必须同时设置 `x: null`、`y: []`、`y_error: {}`。
- `plot.x` 和 `plot.y` 可用源文件中的原始唯一列名或从 0 开始的列索引。`x: null` 表示以行号为 X。折线和散点图需要数值或日期时间 X；分类 X 使用柱状图。
- `plot.y_error` 的键是被选中 Y 列的**原始唯一列名**，值是对应数值误差列的原始列名或从 0 开始的索引。误差列必须为非负数值，且不能与 X/Y 列重合。
- `column_labels` 使用原始唯一列名作为键，可设置 `{ "name": "新列名", "unit": "单位" }`。只更改 Origin／Excel 中的标签，不更改值或数据类型。X/Y 选择应基于原始列名或原始索引，而不是标签改名后的内容。
- `output.format` 为 `opju` 或 `xlsx`，应与目标后缀对应。已有输出默认拒绝；`overwrite` 只能在用户明确选择替换时设为 `true`。`keep_open` 只适用于 `opju`。

同一 Excel 文件有多个来源表时，每个 table 的 `source.options.sheet` 都要明确指定对应工作表名。生成计划时已自动记录这些信息。完整字段约束以 JSON Schema 和 `validate_import_plan` 的实际验证为准。

`.opju` 和 `.xlsx` 输出都包含 `Provenance` 工作表，记录源路径、工作表、SHA-256、行数、列名、原列名、类型、单位、缺失数、读取选项和绘图设置。执行结果还会返回输出文件的路径、格式、字节数及 SHA-256，可用于记录和后续核对。

## 输入范围和数据保真

支持常见扁平矩形表：TXT、DAT、XY、CSV、TSV、XLSX、XLSM、XLS、JSON、JSONL 和 NDJSON。Excel 默认读取全部非空工作表。JSON 顶层可为字段一致的记录数组、等长列数组对象，或 `{ "columns": [...], "data": [[...]] }`；JSONL／NDJSON 要求每行是字段一致的标量记录。嵌套对象和数组、字段不一致、不同长度的列数组都明确报错。

格式读取和 Origin 写出不等于科学数据分析。日期、数值和文本按识别到的数据类型写入目标；带时区的日期时间会先转成 UTC，再以不带时区的值写入。数值精度最终受 Origin／Excel 数值单元格格式限制。缺失值保留为空单元格。导入器不会推测单位、补值、重采样、求平均或转换数据。

## GUI 来源检查（当前阶段）

通用数据窗口在后台展开文件夹，并按**单个路径**调用检查。摘要缓存只保存列结构、最多几行样例和图形建议，不保存完整原始表。缓存是否有效看文件内容 SHA-256 和本次读取选项，不看文件大小或修改时间。

`GeneralDataApp.batch_source_snapshot()` 是留给后续批量阶段的只读入口，内容是路径、成功摘要和来源错误。它不启动导出，也没有取消或 manifest。单工程在存在失败或未完成来源时会拒绝计划；批量阶段不能假设每个来源都检查成功。前缀分组、批量任务配置和批量导出都不在这一阶段。
