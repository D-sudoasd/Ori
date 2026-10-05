# Ori

已知任务位于本仓库时，直接从用户给出的文件、diff 或对应入口开始。
本说明取代上级通用接手流程中的固定文档预读；无需先读工作区地图、
完整 README 或建立全仓文件地图。已有信息在内容未变时继续使用。

## 按任务查阅

| 当前任务 | 源码入口 | 需要接口细节时再读 |
| --- | --- | --- |
| 来源格式、类型或导入计划 | `origin_bridge/readers.py`、`planning.py` | `docs/import-plan.schema.json` |
| CLI / MCP 导入 | `origin_bridge/cli.py`、`mcp_server.py` | `docs/agent-guide.md` |
| 按文件导出、重试或续跑 | `origin_bridge/batch.py`、`session.py` | `docs/batch-api.md` |
| 来源预览或窗口操作 | `origin_bridge/gui.py`、`source_summary.py` | 对应 GUI 测试 |
| 写出与打包 | `origin_bridge/exporter.py`、`SpectraToOrigin.spec` | `docs/development.md` |
| 旧两列光谱流程 | `spectra_to_origin.py` | 对应 `test_import.py` / `test_export.py` |

先局部检索、读取相关函数及调用方。只有证据不足、检查失败、范围变化
或跨越数据/执行边界时扩大范围；不要每轮扫描全仓或重读同一份说明。
独立读取可合并调用；仅在实际减少等待时间时并行，不固定启用多个 agent。

## 必须保留的边界

- 保留来源值、类型、单位与缺失值，不补值或生成科学结论。
- 计划 schema v1 的字段、列映射、输出覆盖规则仍严格校验。
- 导入工具的每次独立调用核对来源内容 SHA-256；不能只用大小/mtime 判断未变。
  `SourceCache` 在一次 CLI 命令或一个 MCP 服务内有界复用解析和摘要，
  不保存到磁盘、不缓存失败或执行授权。写出前仍独立核对来源。
- 导出须已有用户授权；MCP 执行须 `confirm=true`。已有 Origin 会话和
  非本任务拥有的进程必须保留。XLSX 导出仅有表格，不声称创建图形。

## 验证与交付

用 `python scripts/check_changes.py` 查看本次改动所需检查；分支比较可传
`--base origin/main`，`--run` 执行，`--full` 进入完整路径。选中检查通过且
相关内容/环境未变时复用结果；新修改只补相关检查，失败先复测失败边界。
共享数据契约、未知代码/配置或明确未解决风险才扩大到完整测试。
原生 Origin、GUI、打包与跨 Python 版本的触发条件见 `docs/development.md`。

交付前 review 最终 diff，确认范围和适用验证。用户要求 Git 收尾时完成
commit / merge / push，并核对远端；不因提交或切换分支重复未变内容的测试。
