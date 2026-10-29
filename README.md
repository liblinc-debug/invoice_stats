# 发票整理工具 v1.1

这是一个带操作界面的发票整理工具。它从 PDF 发票提取开票日期、发票号码、项目名称和价税合计（小写），按日期正序写入 Excel 模板。

## Windows 安装和启动

1. 安装 Python 3.9 或更高版本，安装时勾选 `Add Python to PATH`。
2. 双击 `Windows安装依赖.bat`，等待出现“依赖安装完成”。
3. 以后双击 `Windows启动统计.bat` 即可打开工具。

如果安装依赖失败，请先检查网络和 Python 是否可从命令行启动。重新双击安装脚本会复用原有 `.venv` 环境。

## macOS 安装和启动

1. 双击 `安装依赖.command`，或在终端运行 `./install.sh`。
2. 双击 `启动统计.command`，或在终端运行 `./startup.sh`。

## 使用流程

1. 点击“选择模板”，选择 `.xlsx` 台账模板。工作区会按子表显示模板内容。
2. 点击“选择发票”，选择包含 PDF 的目录。
3. 点击“统计”。进度窗口会显示当前文件，可点击“取消”中止；统计期间其他按钮会暂时禁用。
4. 在工作区切换子表、检查结果。单元格可直接修改。
5. 点击“导出”，保存当前工作区的全部子表。默认文件名为 `模板名_YYYYMMDD_HHMM.xlsx`。

### 发票目录规则

- 如果所选目录下有包含 PDF 的直接子目录，子目录名必须与 Excel 子表名对应。例如 `材料费/` 对应“材料费”子表，`差旅费/` 对应“差旅费”子表。
- 如果没有包含 PDF 的子目录，所选目录内的 PDF 全部写入模板第一张子表。
- 根目录 PDF 和子目录 PDF 同时存在时，工具会要求先整理目录，防止发票写错子表。

## 统计规则

- 自动识别表头中的日期、发票号、名称和金额列，列顺序可以不同。
- 每张发票生成一行，按开票日期和发票号正序排列。
- 同一子表内按发票号去重。重复文件内容不一致时会报错。
- 表头中没有要求提取的字段保持空白。
- 任一 PDF 解析失败时不会修改工作区，状态栏和弹窗会显示原因。
- 导出时保留模板的全部子表和已有格式。

## 命令行模式

`invoice_stats.py` 作为 v1.0 命令行入口保留，便于批量运行或排查：

```bat
.venv\Scripts\python.exe 发票统计工具\invoice_stats.py --check-only
```

```bash
.venv/bin/python 发票统计工具/invoice_stats.py --check-only
```

## 测试

Windows：

```bat
set QT_QPA_PLATFORM=offscreen
.venv\Scripts\python.exe -m unittest discover -s 发票统计工具 -p "test_*.py"
```

macOS / Linux：

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m unittest discover -s 发票统计工具 -p 'test_*.py'
```

## 代码和脚本

- `发票统计工具/gui.py`：PyQt5 界面和通用目录/表头映射。
- `发票统计工具/invoice_stats.py`：PDF 解析后端及 v1.0 命令行入口。
- `发票统计工具/test_*.py`：解析、目录映射和工作表写入测试。
- `发票统计工具/requirements.txt`：固定版本的 Python 依赖。
- `发票统计工具/参考脚本/`：首次整理中使用的检查脚本，供排查保留。

## 已知限制

- 支持带可提取文字层的中国电子发票 PDF。纯图片扫描件需先做 OCR。
- 项目名称使用规则提取。遇到新发票版式时应人工复核。
- 工作区是轻量表格预览/编辑器，不完整复制 Excel 的图表、公式计算和条件格式编辑功能。未编辑的模板对象会由 openpyxl 原样保留。
