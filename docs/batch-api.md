# 批量导出 API

批量核心在 `origin_bridge.batch`。GUI 下一阶段只调用这里的 JSON 安全请求和 `origin_bridge.worker.batch_worker`，不要在 Tk 里启动 Origin，也不要改 `origin_bridge/gui.py`。

`session_factory`、进程内 `progress` 回调和 `cancel_event` 不是 JSON 字段。跨进程时请求本体必须能 `json.dumps`，取消事件单独作为 spawn 参数传入。

## 签名

```python
from origin_bridge.batch import (
    BATCH_SCHEMA_VERSION,          # 1
    RECORD_KIND,                   # "ori-batch-record"
    build_batch_request,
    execute_batch,
    load_batch_record,
    save_batch_record,
    read_batch_log,
    retry_batch_tasks,
    log_path_for,
    stable_output_path,
    task_id_for,
)
from origin_bridge.worker import execute_batch_request, batch_worker

build_batch_request(
    inputs: list[str | Path],
    output_dir: str | Path,
    *,
    layout: str = "per_file",          # "per_file" | "unified"
    format: str = "opju",              # "opju" | "xlsx"
    read_options: Mapping | None = None,
    plot: Mapping | None = None,
    task_overrides: list[Mapping] | None = None,
    overwrite: bool = False,
    pdf: bool = False,
    keep_open: bool = False,
    recycle_every: int = 0,            # 0 表示不按文件重启 Origin
    prefetch: int = 1,
    max_resident_tasks: int = 2,       # 至少为 1；含当前任务
    origin_timeout_s: float = 300,
    origin_retries: int = 1,           # 会话丢失后的额外尝试次数
    output_name: str | None = None,    # 只用于 unified
    record_path: str | Path | None = None,
    resume: bool = False,
    retry_task_ids: list[str] | None = None,
) -> dict

execute_batch(
    request: Mapping,
    *,
    progress: Callable[[dict], None] | None = None,
    cancel_event: Any = None,          # 任何带 is_set() 的对象，含 multiprocessing.Event
    session_factory: Callable[[], Any] | None = None,  # 测试或嵌入替身；不是 JSON 字段
    resume: bool | None = None,        # 传入时覆盖请求里的 resume
) -> dict

execute_batch_request(request, cancel_event=None, progress=None) -> dict
batch_worker(send_conn, request, cancel_event=None) -> None
```

`build_batch_request` 只检查并规范化请求，不读数据，也不启动 Origin。`pdf=True` 要求 `format="opju"`。`keep_open=True` 同样只适用于 OPJU。

默认读取选项：

```json
{
  "header": "auto",
  "skip_rows": 0,
  "delimiter": "auto",
  "encoding": "auto",
  "missing_values": [""],
  "formula_policy": "cached"
}
```

省略 `sheet` 表示 Excel 的全部非空工作表进入同一个任务。一个 xlsx/xlsm/xls 文件只产生一个 OPJU 或 XLSX。

默认绘图请求：

```json
{
  "kind": "auto",
  "x": "auto",
  "y": "auto",
  "y_error": "auto",
  "title": "auto",
  "x_label": "auto",
  "y_label": "auto"
}
```

`auto` 使用现有建议。显式值按列名或从 0 开始的列号解释，列不存在、类型不匹配或重复时只让该任务失败，不会改选其他列。`kind: "none"` 不绘图。`y_error` 为对象时，键必须是已选 Y 列的原始列名，值是误差列的列名或列号。任务级 `task_overrides[].plot` 和 `read_options` 覆盖同名字段；未写的字段仍用全局值。

`task_overrides` 的每一项至少包含 `source`。路径会解析成绝对路径后再与发现到的文件比较。

## 请求示例

```json
{
  "schema_version": 1,
  "inputs": ["D:/data/光谱.csv", "D:/data/book.xlsx"],
  "output_dir": "D:/out/batch",
  "layout": "per_file",
  "format": "opju",
  "read_options": {
    "header": "auto",
    "skip_rows": 0,
    "delimiter": "auto",
    "encoding": "auto",
    "missing_values": [""],
    "formula_policy": "cached"
  },
  "plot": {
    "kind": "line",
    "x": "x",
    "y": ["y"],
    "y_error": "auto",
    "title": "auto",
    "x_label": "auto",
    "y_label": "auto"
  },
  "task_overrides": [
    {
      "source": "D:/data/book.xlsx",
      "plot": {"kind": "none"}
    }
  ],
  "overwrite": false,
  "pdf": false,
  "keep_open": false,
  "recycle_every": 0,
  "prefetch": 1,
  "max_resident_tasks": 2,
  "origin_timeout_s": 300,
  "origin_retries": 1,
  "output_name": "batch.opju",
  "record_path": "D:/out/batch/batch-record.json",
  "resume": false,
  "retry_task_ids": []
}
```

## 输出命名与发现

`per_file` 的输出名是 `<清洗后的 stem，最长 160>__<路径 SHA-256 前 12 位>.opju|xlsx`。中文保留，同名文件因完整路径不同而不会撞名。任务 ID 是同一路径哈希的前 16 位十六进制。

PDF 为 `<opju stem>__g01_<表名最多 40 字符>.pdf`，序号按图形顺序从 01 起。没有图形的任务不会生成 PDF。

递归发现输入时会排除：

- 输出目录里面的文件（输出目录本身就是扫描根目录时除外）
- 记录、`.bak`、`.log.jsonl` 和 `.previous`
- 统一布局的目标文件
- 符合 `__` + 12 位十六进制 + 可选 `__gNN_` + `.opju/.xlsx/.pdf` 的生成文件名

因此把输出目录放进输入目录，或把上一轮产物留在输入树里，都不会被再次当成输入。

## 执行与内存

`layout="per_file"` 在一个由本程序创建的 Origin 会话里串行执行 COM。默认只保留当前任务再加 `prefetch` 个已解析任务，并且不超过 `max_resident_tasks`。默认是当前任务加 1 个预读，最多 2 份完整表。这不是跨任务的转换缓存。数值和日期仍在写入前做完整性检查；下溢、非数值文本和非法日期会让该任务失败。同一次转换产生的值和 warning 会交给填充、回读和 provenance，不另建缓存框架。

`recycle_every` 默认 0。设为正整数时，每成功写完这么多个 OPJU 就关闭并重新启动本程序拥有的 Origin。

`layout="unified"` 是旧的单工程导出。它会一直持有全部表，不受 `max_resident_tasks` 限制。`output_name` 只在这种布局下决定文件名。

已有 Origin 进程会拒绝附着，不会被修改。所有导出路径，包括旧的 `spectra_to_origin` 入口，共用同一把可重入的跨进程锁。同一线程可以嵌套获取；其他线程或进程得到「Another generic Origin export is already active」。超时或失败只允许结束本任务记录到的 Origin PID，不会全局 `taskkill` Origin。

## 结果示例

`ok` 在存在失败、阻塞、取消或 PDF 失败时为 false。成功任务的 OPJU 不会因为 PDF 失败被删除。

```json
{
  "ok": false,
  "schema_version": 1,
  "cancelled": false,
  "output_dir": "D:/out/batch",
  "layout": "per_file",
  "format": "opju",
  "record_path": "D:/out/batch/batch-record.json",
  "log_path": "D:/out/batch/batch-record.log.jsonl",
  "origin": {"starts": 1, "stops": 1, "owned_pids": [12345], "recycled": 0},
  "resident": {"max_full_tasks": 2, "limit": 2},
  "counts": {
    "succeeded": 1,
    "skipped": 0,
    "failed": 0,
    "cancelled": 0,
    "blocked": 0,
    "pdf_failed": 1
  },
  "tasks": [
    {
      "task_id": "0123456789abcdef",
      "source": "D:/data/光谱.csv",
      "sources": ["D:/data/光谱.csv"],
      "status": "succeeded",
      "output": {
        "path": "D:/out/batch/光谱__0123456789ab.opju",
        "format": "opju",
        "size_bytes": 4096,
        "sha256": "…"
      },
      "pdfs": [],
      "pdf_status": "failed",
      "error": null,
      "warnings": [],
      "resolved_plot": [
        {
          "table": "光谱",
          "kind": "line",
          "x": 0,
          "y": [1],
          "y_error": {},
          "title": "光谱",
          "x_label": "x",
          "y_label": "y"
        }
      ]
    }
  ],
  "warnings": []
}
```

任务 `status`：`succeeded`、`failed`、`skipped`、`cancelled`、`blocked`。`blocked` 表示目标已存在且未允许覆盖，已有字节保持不变。

`pdf_status`：`not_applicable`（未请求 PDF，或该任务没有图形）、`ok`、`failed`。PDF 要有 `%PDF-` 头、`%%EOF`、`startxref` 或 `/XRef`，以及至少一个 `/Type /Page`。只检查扩展名或非空文件不够。

`error` 为 `{"type": "...", "message": "..."}` 或 null。进度回调自身抛出的异常记入 `warnings`，不会取消批次。

## 进度示例

进程内回调收到的对象：

```json
{"type": "batch", "status": "started", "count": 2}
```

```json
{
  "type": "task",
  "task_id": "0123456789abcdef",
  "index": 0,
  "count": 2,
  "status": "succeeded",
  "source": "D:/data/光谱.csv",
  "output": "D:/out/batch/光谱__0123456789ab.opju",
  "error": null,
  "pdf_status": "not_applicable"
}
```

这些事件同时追加到 `log_path`。损坏的最后一行在 `read_batch_log` 里变成 `{"type": "corrupt_log_line", "raw": "..."}`，前面的完整行仍然有效。

## 取消

取消是协作式的，只在任务之间检查 `cancel_event.is_set()`。已经成功安装的文件保留。尚未开始的任务记为 `cancelled`。当前正在写的那一个任务会跑完本次尝试。

Spawn 入口：

```python
import multiprocessing as mp
from origin_bridge.worker import batch_worker

ctx = mp.get_context("spawn")
recv, send = ctx.Pipe(duplex=False)
cancel = ctx.Event()
proc = ctx.Process(target=batch_worker, args=(send, request, cancel))
proc.start()
send.close()  # 父进程关掉自己这一端的发送连接
cancel.set()  # 当前任务结束后不再开始后续任务
while True:
    message = recv.recv()
    if message["type"] == "progress":
        event = message["event"]
    elif message["type"] == "result":
        result = message["result"]
        break
proc.join()
```

子进程消息只有两种：`{"type": "progress", "event": <上面的进度对象>}` 和 `{"type": "result", "result": <结果对象>}`。子进程异常时 `result` 为 `{"ok": false, "error": "..."}`。`batch_worker` 必须保持模块级函数，供 Windows spawn 导入。

## 记录、续跑和重试

默认记录是 `output_dir/batch-record.json`，日志是把后缀换成 `.log.jsonl`。写入使用临时文件、`fsync` 和 `os.replace`。上一个完整文件保留为 `.bak`。主文件损坏时 `load_batch_record` 使用 `.bak`，并在结果警告里说明已恢复。主文件和备份都损坏时得到空记录和 `_corrupt` 警告，不会把已有输出当成成功。

非续跑会把已有记录改名为 `batch-record.json.previous`，然后从空记录开始。`.previous` 不会在下次启动时被自动恢复。

续跑只有在下面全部成立时把任务标为 `skipped`：

- 记录状态是 `succeeded`
- 源文件 SHA-256 与记录一致
- 读取选项、绘图请求、格式、是否请求 PDF、指纹和输出路径一致
- 输出文件存在、非空，且 SHA-256 与记录一致
- 若当时 PDF 状态是 `ok`，每个 PDF 仍存在、哈希一致且结构仍有效

文件存在、文件大小相同或 mtime 相同都不会单独导致跳过。源内容变化、输出被替换、输出变空、PDF 之前失败、格式或绘图请求变化，都会重新执行。`overwrite` 仍为 false 时，重新执行遇到已有文件会 `blocked`，且不改那个文件的字节。跳过不会覆盖记录里的成功条目。

成功安装之后、记录落盘之前中断时，磁盘上可以已经有输出，但记录里没有成功条目。再次续跑且不允许覆盖时，该任务是 `blocked`，已有字节保持不变。

`retry_task_ids` 会在续跑前从记录中去掉这些 ID，即使它们曾经成功。输出文件不会被删除；不允许覆盖时，再次执行会阻塞而不是替换。

```python
retry_batch_tasks(
    record_path,
    task_ids=None,                         # None 表示按 statuses 选择
    statuses=("failed", "cancelled", "blocked"),
    force_task_ids=None,                   # 即使状态是 succeeded 也去掉
)
```

CLI `--retry-failed` 等价于续跑，并只去掉状态为 `failed`、`cancelled` 和 `blocked` 的任务。失败任务在普通 `--resume` 时本来就会重跑；这个开关把选择写进记录。

Origin 会话丢失（进程退出、无响应、COM/`com_error`）会在 `origin_retries` 允许的额外次数内换一个新的本程序会话再试该任务。`DataImportError`、`FileExistsError` 和 `ValueError` 不是会话丢失，不重启 Origin。重试耗尽后该任务失败，后面的任务仍会尝试新会话。

记录本身写失败时，该任务在本次结果里改为 `failed`，内存中的对应记录条目去掉，警告说明输出可能已安装但记录未保存。已安装文件不会回滚。

## CLI

```powershell
py -3 -m origin_bridge batch -i .\data -o .\out --format xlsx --plot none
py -3 -m origin_bridge batch -i .\data -o .\out --pdf --plot line --x x --y y
py -3 -m origin_bridge batch -i .\data -o .\out --resume
py -3 -m origin_bridge batch -i .\data -o .\out --retry-failed
py -3 -m origin_bridge batch -i .\data -o .\out --layout unified --output-name combined.opju
```

读取参数与 `inspect` / `import` 相同：`--header`、`--skip-rows`、`--sheet`、`--delimiter`、`--encoding`、`--missing-value`、`--formula-policy`。全数字的 `--x` / `--y` 按列号解释，其他文本按列名解释。

标准输出仍是 UTF-8 JSON。`batch` 在结果 `ok` 为 false 时退出码 4，同时仍打印完整结果。参数错误也是 `{"ok": false, "error": {"type", "message"}}` 和退出码 4。`inspect`、`plan`、`import`、`validate`、`execute` 的输出形状不变。

## 真机验收

普通单元测试不启动 Origin。显式真机检查是：

```powershell
$env:ORI_BATCH_ORIGIN = "1"
py -3 -m unittest test_batch_origin.py
```

Origin 已在运行时测试会跳过或失败，不会关掉用户的 Origin。临时输入和输出在系统临时目录，不在仓库里。
