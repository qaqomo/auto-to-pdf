📄 Word & Excel 批量转 PDF 工具

一个轻量级、易用的 Windows 图形界面工具，用于批量将 Word (.doc/.docx) 和 Excel (.xls/.xlsx/.xlsm/.xlsb) 文件转换为 PDF 格式。支持中文和空格路径，适合办公自动化需求。

✅ 支持批量转换

✅ 支持中文和空格路径

✅ 图形界面简单明了

✅ 可直接运行 .exe 文件，无需安装 Python

🚀 程序用途

此程序可将指定文件夹中的所有 Word 和 Excel 文件批量转换为 PDF。转换过程自动进行，支持选择输出目录及指定 Excel 的工作表索引。

适用于：

办公自动化

教师批量出卷、保存

财务、行政等批量文档转换

📦 安装依赖（如运行 .py 版）

若你使用源码运行，请确保已安装以下依赖：

pip install pywin32

此外，系统必须安装 Microsoft Office Word 与 Excel。

🛠 使用说明
✅ 使用 .exe 可执行文件（推荐）
下载发布版本（请替换为你 GitHub Release 页的链接）；

解压并双击运行 auto-to-pdf.exe；

按照界面提示：

选择源文件夹（包含 Word/Excel 文件）；

可选设置输出目录；

可选设置 Excel 的工作表索引（默认第一个）；

点击“开始转换”；

完成后界面会显示绿色对勾 ✅。

无需安装 Python 或 Office SDK，仅需本地 Office 即可运行。

🐍 使用 .py 源码文件
安装 Python 3.8 及以上版本；

安装依赖：
pip install pywin32

终端运行：
python auto-to-pdf.py

📁 示例界面



💡 注意事项
Excel 默认转换第一个工作表，如需指定，请修改工作表索引；

输出文件名将保留原文件名，仅更换为 .pdf 后缀；

转换过程中建议不要打开 Word/Excel。
