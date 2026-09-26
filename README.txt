Data to Origin
==============

把常见的表格数据导入 Origin Pro，生成 .opju 工程；也可只导出 Excel。
源码新版默认打开通用数据窗口，可处理多列表格和 Excel 多工作表。

支持 TXT、DAT、XY、CSV、TSV、XLSX、XLSM、XLS、JSON、JSONL、NDJSON。
数据可包含数值、文本、日期时间和缺失值。数据不会自动插值、拟合、平滑、
单位换算或归一化。详细支持范围见 README.md。

开始使用（Python 3.10+，Windows）：

  py -3 -m pip install -r requirements.txt
  py -3 spectra_to_origin.py

可把文件或文件夹路径传给程序，也可打开旧的 XRD／光谱批量导入窗口：

  py -3 spectra_to_origin.py .\examples_general
  py -3 spectra_to_origin.py --spectra-gui .\examples

Agent 命令行检查、创建并校验导入计划：

  py -3 -m origin_bridge inspect -i .\examples_general
  py -3 -m origin_bridge plan -i .\examples_general\stress_strain.tsv -o .\out\stress.opju --save .\out\stress-plan.json
  py -3 -m origin_bridge validate .\out\stress-plan.json
  py -3 -m origin_bridge execute .\out\stress-plan.json

也可安装 requirements-agent.txt 并启动 `python -m origin_bridge.mcp_server`，
通过 MCP 客户端调用 inspect_data、create_import_plan、validate_import_plan、
execute_import_plan。完整 agent 指南见 docs/agent-guide.md。

生成 .opju 需要 Windows 上已安装并可用的 Origin Pro；已有 Origin 会话运行时
程序会停止导出，以保护当前工程。导出 .xlsx 不需要 Origin。

构建的 SpectraToOrigin.exe 打开通用 GUI；DataToOriginCLI.exe 输出命令行 JSON。
MIT License。
