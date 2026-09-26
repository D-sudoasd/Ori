SpectraToOrigin 1.1.0
====================

把两列 XRD／光谱文件批量导入 Origin：每组一张工作表、一张曲线图。
也可以不启动 Origin，直接导出全部分组的 Excel 和 CSV。

1. 双击 run.bat，或运行构建好的 SpectraToOrigin.exe。
2. 拖入 txt/dat/xy/csv/tsv 文件或文件夹，选中谱线查看预览。
3. 选择自动布局、分组方式，填写坐标轴名称和单位。
4. 点击生成 Origin 工程，或只导出 Excel/CSV。

生成 .opju 需要安装并激活 Origin Pro；先保存并关闭已打开的 Origin。
表格导出不需要 Origin。多组 CSV 位于“输出名_csv”文件夹，每组一个文件。
本工具只整理和绘图，不做插值、拟合、平滑或强度归一化。

源代码运行：
py -3 -m pip install -r requirements.txt
py -3 spectra_to_origin.py

检查演示数据：
py -3 spectra_to_origin.py --check -i examples --group-by-name

只导出表格：
py -3 spectra_to_origin.py --cli --xlsx-only -i examples -o demo.xlsx --n-groups 2

详细格式、参数、测试与构建说明请阅读 README.md。
