# 批量导出 API

批量核心在 `origin_bridge.batch`。通用数据窗口调用 JSON 安全的 `build_batch_request` 和 `origin_bridge.worker.batch_worker`，取消事件单独作为 spawn 参数。不要在 Tk 线程里启动 Origin。

核心审查阶段曾经要求不要改 `origin_bridge/gui.py`。那只约束当时的核心补丁。批量窗口可以修改 GUI。窗口新建批次时把可恢复的完整请求写到 `batch-config-<id>.json`，记录是同目录的 `batch-record-<id>.json`（名称、`record_path`、输入路径和请求本体），不改核心成功回执。旧版 `batch-record.json` 与 `batch-config.json` 仍能配对加载。只有运行记录、没有配套配置时，不能凭窗口当前默认值重新导出。`y_error` 的键必须是 Y 的原始列名；窗口在有误差列而 Y 为列号时会在保存配置前拒绝。

“停止后续”只在任务之间生效：当前这项会跑完，后面的任务取消。它不能打断还没证明归属的 Origin 启动；那种 COM 调用可能一直不返回，此时没有可以安全结束的进程。

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
- 目录扫描中符合 `__` + 12 位十六进制 + 可选 `__gNN_` + `.opju/.xlsx/.pdf` 的生成文件名

目录扫描排除生成文件名时，结果 `warnings` 里有 `excluded <路径>: generated batch output name`。调用方显式传入的文件，即使名字符合这个形态，也不会被这条规则丢掉。输入发现或“排除后没有文件”在打开记录之前失败：已有 `batch-record.json`、`.bak`、`.previous` 和 `.log.jsonl` 的字节保持不变，不会先归档再写空记录。输出目录可以先被创建。

## 执行与内存

`layout="per_file"` 在一个由本程序创建的 Origin 会话里串行执行 COM。默认只保留当前任务再加 `prefetch` 个已解析任务，并且不超过 `max_resident_tasks`。默认是当前任务加 1 个预读，最多 2 份完整表。这不是跨任务的转换缓存。数值和日期仍在写入前做完整性检查；下溢、非数值文本和非法日期会让该任务失败。同一次转换产生的值和 warning 会交给填充、回读和 provenance，不另建缓存框架。

`recycle_every` 默认 0。设为正整数时，每成功写完这么多个 OPJU 就关闭并重新启动本程序拥有的 Origin。

`layout="unified"` 是旧的整表单工程路径。它会一直持有全部表，不受 `max_resident_tasks`、预读或续跑跳过的约束。`output_name` 只在这种布局下决定文件名。`unified` 不能带 `pdf=True`，也不能带 `resume=True` 或 `retry_task_ids`：构建和执行都会拒绝，而不是返回 `ok: true` 却没有 PDF，也不是无提示地把整个工程重导一遍。PDF 和续跑只用 `per_file`。

已有 Origin 进程会拒绝附着，不会被修改，也不会被结束。本程序只在 COM 启动之后用该实例打开一个私有哨兵文件，再用 Windows Restart Manager 查询持有者。路径经 `set_lt_str("file.filename$", path)` 写入，再执行 `file.mode=2; file.open();`。不要把带反斜杠的路径直接放进 `file.open(路径)`：LabTalk 会把它当成转义，打开结果不能当作持有证明。只有恰好一个持有者，且其可执行映像是 `Origin64.exe` 或 `Origin.exe` 时才认领。认领同时保存进程句柄和创建时间。Restart Manager 的应用显示名不是映像路径，不能当作证明。零个持有者、多个持有者、映像不是 Origin，或句柄创建时间对不上，都算没有证明：调用该 COM 对象的 `exit()` 后失败，错误写明没有结束候选进程。不会因为“启动后新出现的 Origin”或“当时只有一个 Origin”去结束进程。

超时、关闭等待和 `terminate_owned` 使用同一套身份：句柄仍对应原创建时间，并且该 PID 当前的创建时间也相同。PID 被复用时视为不健康，不结束那个新进程。句柄在关闭时释放，重复关闭不再释放一次。

看门狗只保护已经证明归属的阶段，并且只在 `origin_timeout_s > 0` 时启动。`set_show(False)` 以及哨兵证明里的 `set_lt_str("file.filename$")` 和 `lt_exec` 都在这之前，是直接 COM 调用，没有截止时间。许可证对话框或首次向导如果卡在这里，`execute_batch` 停在 `session.start()`，调用不返回。这不是返回失败。此时还没有证明过的进程句柄，不能安全结束任何 Origin，也不会为了让调用返回去结束一个未知进程。

启动前有三次“当前没有 Origin”检查：加载 originpro 之前、加载之后，以及活会话枚举到的 PID。这三次不是原子操作，不能排除检查之后、COM 对象真正建起来之前出现的进程。当前安装的 `originpro/config.py` 里，`APP.__getattr__` 在第一次取属性时构造 `OriginExt.Application()`，并执行 `LT_execute('sec -poc')`。`APP.Attach` 的说明是附着已有实例，用的是 `OriginExt.ApplicationSI()`。这条批量路径不调用 `Attach`。文件里只说明默认构造和附着入口不同，不能据此说三次检查已经消除竞争。

`close_origin_app` 在 `exit()` 之后仍最多等 8 秒；等待期间只要还能看到 `Origin64.exe` 或 `Origin.exe` 就继续等。这是旧导出入口的关闭等待，不证明该进程属于本次会话，也不结束未证明的进程。会话自己的收尾只核对已证明的句柄，身份仍成立时才结束那个句柄。这 8 秒的全局可见性等待是已知限制。

`origin.owned_pids` 是本次批次证明过的进程号，按号排序，包含已经关闭的会话。`origin.owned_processes` 是对应的 `pid`、`creation_filetime` 和 `image`。它不是“期间出现过的全部 Origin”。之后若要清理，必须同时匹配创建时间。只提供 PID 的测试替身会记成仅含 `pid` 的条目。

所有导出路径，包括旧的 `spectra_to_origin` 入口，共用同一把锁。同一线程可以嵌套进入。其他线程用非阻塞获取，其他进程用非阻塞文件锁；两边都立即得到「Another generic Origin export is already active」，不会等到当前导出结束。异常退出时锁会释放。不会全局 `taskkill` Origin。

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
  "origin": {
    "starts": 1,
    "stops": 1,
    "owned_pids": [12345],
    "owned_processes": [{"pid": 12345, "creation_filetime": 133000000000000000, "image": "C:/Program Files/OriginLab/Origin2025/Origin64.exe"}],
    "recycled": 0
  },
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
      "pdf_errors": [
        {"table": "光谱", "graph": "光谱", "short_name": "Graph1", "path": "D:/out/batch/光谱__0123456789ab__g01_光谱.pdf", "message": "pdf export failed"}
      ],
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

`pdf_status`：`not_applicable`（未请求 PDF，或该任务没有图形）、`ok`、`failed`。PDF 恢复被拒绝时，这次结果的 `pdf_status` 是 `failed`，成功记录本身不改。检查器只认未压缩的 `/Type /Page` 字节，外加 `%PDF-` 头、`%%EOF`、`startxref` 或 `/XRef`。内容流上的 `/Filter` 不影响这一判断。没有未压缩页、且文件含 `/ObjStm` 时，错误写明不支持 ObjStm。没有未压缩页的其他文件也会被拒绝。这个检查器不解码对象流，也不假装支持压缩页字典。

PDF 失败时任务仍是 `succeeded`，`error` 为 null，`ok` 为 false。失败原因在结果、进度和记录的 `pdf_errors` 里，每项含 `table`、`graph`、`short_name`、`path`、`message`。已安装的 OPJU 和已经校验通过的 PDF 保留。

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
  "pdf_status": "failed",
  "pdf_errors": [
    {"table": "光谱", "graph": "光谱", "short_name": "Graph1", "path": "D:/out/光谱__0123456789ab__g01_光谱.pdf", "message": "pdf export failed"}
  ]
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

未传入 `record_path` 时，默认记录是 `output_dir/batch-record.json`，日志是把后缀换成 `.log.jsonl`。窗口新建批次会传入 `output_dir/batch-record-<id>.json`，配置文件按同一标识命名为 `batch-config-<id>.json`，因此同目录的下一批不会替换上一批的记录、配置、`.bak`、`.previous` 或日志。写入使用临时文件、`fsync` 和 `os.replace`。上一个完整文件保留为 `.bak`。主文件损坏时 `load_batch_record` 使用 `.bak`，并在结果警告里说明已恢复。主文件和备份都损坏时得到空记录和 `_corrupt` 警告，不会把已有输出当成成功。这些警告出现在结果的 `warnings` 里；窗口汇总条数并提供只读详情，而不是为每条警告各弹一次。

非续跑只有在输入发现成功之后，才把已有记录改名为 `batch-record.json.previous` 并从空记录开始。`.previous` 不会在下次启动时被自动恢复。发现失败不会归档、不会删除 `.bak`、也不会写空记录。

续跑先要求记录状态是 `succeeded`，并且源字节、读取选项、绘图请求、格式、是否请求 PDF、指纹、输出路径和 OPJU 哈希全部一致。文件存在、文件大小相同或 mtime 相同都不会单独导致跳过。

在此之上：

- 未请求 PDF，或 `pdf_status` 是 `not_applicable`：任务 `skipped`。
- 每个记录中的图都有 PDF，哈希与回执一致且仍能通过未压缩页检查：任务 `skipped`。
- 其中一些 PDF 缺失：先把回执核对过的 OPJU 流式复制到该任务独占的临时目录，再核对副本哈希、源配置和原 OPJU 哈希。COM 只打开这个副本，不打开生产 OPJU，也不保存。自动保存或异常写只会改到这个副本。导出返回后、安装每个 PDF 之前再次核对生产 OPJU。生产文件若已变化，不安装这次的新 PDF，不覆盖成功记录，保留读到的新字节；这次结果是 `blocked`，`pdf_status` 是 `failed`。不会为了让哈希回到旧值而把旧字节写回去。
- 表长名、图的长短名必须与记录一致，且各只有一个匹配。provenance 必须包含当时的源路径。对不上就不导出另一张图。
- PDF 已存在但哈希与回执不同，或回执里没有这个文件：结果是 `blocked`，成功记录原样保留。`overwrite=True` 时可以在恢复中替换该 PDF，仍然不重写 OPJU。
- 源、读取或绘图配置、格式、指纹或 OPJU 字节变化：走完整导出，不进入 PDF 补导。`overwrite` 仍为 false 时，已有 OPJU 使该任务 `blocked`，不改那个文件的字节。

成功安装之后、记录落盘之前中断时，磁盘上可以已经有输出，但记录里没有成功条目。再次续跑且不允许覆盖时，该任务是 `blocked`，已有字节保持不变。

`retry_task_ids` 必须和 `resume=True` 一起使用。`resume=False` 时构建和执行都会拒绝，不会悄悄忽略。续跑前只从记录中去掉状态为 `failed`、`cancelled` 或 `blocked` 的指定任务。状态为 `succeeded` 的任务即使出现在 `retry_task_ids` 里也会保留，这样 PDF 失败或缺失仍走上面的补导，而不是删掉回执后撞上已有 OPJU。`force_task_ids` 才会去掉成功条目；去掉之后不允许覆盖时，完整重跑会阻塞。输出文件不会被删除。

```python
retry_batch_tasks(
    record_path,
    task_ids=None,                         # None 表示按 statuses 选择
    statuses=("failed", "cancelled", "blocked"),
    force_task_ids=None,                   # 即使状态是 succeeded 也去掉
)
```

CLI `--retry-failed` 等于 `resume=True`，并且只把 `failed`、`cancelled`、`blocked` 放进 `retry_task_ids`。已成功但 PDF 失败或缺失的任务留在记录里，同一次续跑会补这些 PDF。`keep_open=True` 时，最后一次健康会话保持打开；`recycle_every` 到点仍会关闭并重启，包括最后一个任务。回收计数在 `origin.recycled`。

Origin 会话丢失（进程退出、无响应、COM/`com_error`）会在 `origin_retries` 允许的额外次数内换一个新的本程序会话再试该任务。`DataImportError`、`FileExistsError` 和 `ValueError` 不是会话丢失，不重启 Origin。重试耗尽后该任务失败，后面的任务仍会尝试新会话。

记录本身写失败时，该任务在本次结果里改为 `failed`，内存中的对应记录条目去掉，警告说明输出可能已安装但记录未保存。已安装文件不会回滚。

## CLI

```powershell
py -3 -m origin_bridge batch -i .\data -o .\out --format xlsx --plot none
py -3 -m origin_bridge batch -i .\data -o .\out --pdf --plot line --x x --y y
py -3 -m origin_bridge batch -i .\data -o .\out --resume
py -3 -m origin_bridge batch -i .\data -o .\out --retry-failed
py -3 -m origin_bridge batch -i .\data -o .\out --layout unified --format xlsx --output-name combined.xlsx
```

`unified` 不要加 `--pdf`、`--resume` 或 `--retry-failed`，这三种都会在执行前拒绝。读取参数与 `inspect` / `import` 相同：`--header`、`--skip-rows`、`--sheet`、`--delimiter`、`--encoding`、`--missing-value`、`--formula-policy`。全数字的 `--x` / `--y` 按列号解释，其他文本按列名解释。

标准输出仍是 UTF-8 JSON。`batch` 在结果 `ok` 为 false 时退出码 4，同时仍打印完整结果。参数错误也是 `{"ok": false, "error": {"type", "message"}}` 和退出码 4。`inspect`、`plan`、`import`、`validate`、`execute` 的输出形状不变。

## 真机验收

普通单元测试不启动 Origin。显式真机检查是：

```powershell
$env:ORI_BATCH_ORIGIN = "1"
py -3 -m unittest test_batch_origin.py
```

Origin 已在运行时测试会跳过或失败，不会关掉用户的 Origin。临时输入和输出在系统临时目录，不在仓库里。
