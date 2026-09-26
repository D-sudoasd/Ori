# SpectraToOrigin v1.1.0（源码更新）

- 支持 TXT、DAT、XY、CSV、TSV，两列分隔文本、列名表头、UTF-16 BOM 和 Fortran D 指数。
- 修复不同目录同名文件误去重；自然排序，准确报告无效路径和数据行。
- 按每组的 X 网格独立判断布局；所有分组均导出 CSV。
- 增加曲线预览、轴名称与单位设置、独立 Excel/CSV 导出、后台导出进度。
- 保护已经运行的 Origin 工程；修复旧目标文件导致的保存成功误判及跨卷替换问题。
- 使用可移植测试数据，增加 Windows Python 3.10/3.13 测试和程序构建工作流。

此版本的程序需要从当前源码构建，或下载对应工作流的构建产物；此前发布的 exe 不包含这些改动。

# SpectraToOrigin v1.0.0

Windows 双击即可使用，不必先安装 Python。

## 下载

资源里的 **SpectraToOrigin.exe**（约 18 MB）。放到任意文件夹，双击打开。

1. 把谱线 txt 或整个文件夹拖进窗口
2. 确认自动识别的 XYYY / XYXY
3. 点「生成 Origin 工程 (.opju)」

## 说明

- **不需要 Python**
- 写出 `.opju` 需要本机已安装 **Origin Pro**（经 Origin 2025 验证）
- 没有 Origin 时可加载谱线；v1.0.0 的窗口导出依赖 Origin，表格单独导出需使用下方命令行
- 命令行同样可用：`SpectraToOrigin.exe --cli --xlsx-only -i 谱线文件夹 -o out.xlsx`

请从本 Release 下载 exe，不要用源码树里的 `dist/`（未纳入 git）。
