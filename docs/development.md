# 开发与验证

从当前 diff 和相关调用方开始。已知来源或函数时使用精确路径和局部读取；
源格式不明才检查数据，接口问题不需要预读全部 README、schema 和批次文档。
任务内已有的读取、摘要和检查结果，在相关内容与环境未变时继续使用。

## 检查选择

在仓库目录，使用任务所用 Python：

```powershell
python scripts/check_changes.py                      # staged / unstaged / 未忽略的新文件
python scripts/check_changes.py --base origin/main   # 本分支与基准的差异
python scripts/check_changes.py --run                # 执行选中的 unittest
python scripts/check_changes.py --full               # 明确选择完整路径
```

JSON 输出列出改动、测试范围、MCP 检查和构建条件。脚本不安装依赖、不保存
“已通过”状态，也不把相同检查自动重复运行。`--run` 只执行 unittest；
`build=true` 表示还需适用的打包检查。MCP 测试须使用装有
`requirements-agent.txt` 的环境，跳过不代表 MCP 已验证。

| 改动范围 | 默认检查 |
| --- | --- |
| 文档、图像 | diff、引用与命令核对；不运行代码测试或构建 |
| 单个测试 | 对应 unittest 模块 |
| CLI、MCP、GUI、批次等已知模块 | 模块及其直接调用边界的测试 |
| readers、planning、models、dates、source_cache | 完整 unittest、MCP 和打包检查 |
| 启动入口、worker、Origin session、依赖、打包配置 | 完整检查及打包 |
| 新代码、未知配置、不可用的 Git 基准 | 保守进入完整路径，再据事实缩小后续检查 |

CI 使用同一选择器；文档改动不启动测试矩阵或打包，MCP 专项使用 SDK
环境，完整路径保留 Python 3.10 / 3.13 矩阵。手动 `workflow_dispatch`
执行完整检查。只取当前比较所需的 Git 提交，不为路由拉取完整历史。

失败后先复测失败项及其边界；修复跨共享契约才再跑完整路径。通过结果
只有在新修改、环境变化、检查失败或具体风险未解决时才需要重复/扩大。
提交、merge 和 push 不会使未变源码的测试结果失效。

## 原生与打包检查

改动 Origin 写入、会话生命周期或所有权判断时，运行受影响的真实 Origin
测试；已有会话必须保留。需要已安装且可用的 Origin，入口为
`SPECTRA_TEST_ORIGIN=1` / `ORI_BATCH_ORIGIN=1`。Tk 交互改动才设置
`SPECTRA_TEST_GUI=1`；测试跳过时记录相应未验证范围。

新增运行模块、改动包装入口/worker、依赖或 spec 时，构建一次并验证
新包的相关入口。优先给构建指定独立 `--distpath` 和 `--workpath`；无需
每次 `--clean`。例如已有构建环境可运行：

```powershell
python -m PyInstaller --noconfirm --distpath <独立输出目录> --workpath <独立构建目录> SpectraToOrigin.spec
python scripts/verify_packaged_worker.py --executable <新包目录>/SpectraToOrigin.exe
```

CLI 的 `import` 验证直接导入，`plan` / `execute` 验证保存计划重放；
只有独立校验接口改动或未覆盖风险才额外调用 `validate`。包内的 worker
验证 Windows spawn 和打包模块，不能用源代码 worker 的成功代替。

## 来源复用

CLI 的计划/导入/重放、MCP、合并批次和 GUI worker 使用 `SourceCache`。
单次 CLI `inspect` 直接读取，不创建后续不会复用的缓存。每次独立操作
核对 SHA-256，同一操作内复用来源身份。缓存只保留不可变列值和受复制
保护的选项/摘要；不缓存计划批准、转换校验或输出存在状态。新内容、新
读取选项、文件消失/读失败、容量淘汰都会触发适用的重新读取。

调用 Python API 需要连续检查/计划/校验时，可给 `inspect_inputs`、
`create_plan`、`prepare_plan` 传同一个 `cache=SourceCache()`；每次调用
仍独立核对内容。在一条已授权命令内部，用 `with cache.operation():`
共享快照身份，写出器仍独立核对来源。只使用当前任务相关来源，不能用
工作区根目录充当默认输入。

效果与复现命令见 [效率验证记录](agent-efficiency.md)。
