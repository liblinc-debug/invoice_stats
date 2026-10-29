# IDEA 项目发票统计工具

本工具从当前目录下的 `材料费/`、`差旅费/` PDF 发票中提取开票日期、发票号码、项目名称和价税合计（小写），写入 `IDEA项目材料台账.xlsx` 对应子表。

## Windows 首次使用

1. 安装 Python 3.9 或更高版本。安装时勾选 `Add Python to PATH`。
2. 双击 `Windows安装依赖.bat`。
3. 保持以下目录和文件名不变：
   - `IDEA项目材料台账.xlsx`
   - `材料费/`
   - `差旅费/`
4. 双击 `Windows启动统计.bat`。
5. 结果保存在 `输出/`，文件名包含运行时间。原台账不会被覆盖。

## macOS 首次使用

1. 双击 `安装依赖.command`（也可在终端运行 `install.sh`）。
2. 保持以下目录和文件名不变：
   - `IDEA项目材料台账.xlsx`
   - `材料费/`
   - `差旅费/`
3. 双击 `启动统计.command`（也可在终端运行 `startup.sh`）。
4. 结果保存在 `输出/`，文件名包含运行时间。原台账不会被覆盖。

## 统计规则

- 每张发票生成一行，按开票日期正序排列。
- 同一目录内按发票号去重；重复文件内容不一致时停止生成并报错。
- 材料费写入购买日期、发票号、物料品名、金额。
- 差旅费写入日期、发票号、发票名称、发票金额。
- 其余台账字段保持空白。
- 任意 PDF 解析失败时不生成半成品台账。

## 命令行用法

Windows：

```bat
Windows启动统计.bat
Windows启动统计.bat --check-only
Windows启动统计.bat --workbook "其他台账.xlsx" --output "输出\指定文件名.xlsx"
```

macOS：

```bash
./启动统计.command
./启动统计.command --check-only
./启动统计.command --workbook "其他台账.xlsx" --output "输出/指定文件名.xlsx"
```

查看完整参数：

```bat
.venv\Scripts\python.exe 发票统计工具\invoice_stats.py --help
```

```bash
.venv/bin/python 发票统计工具/invoice_stats.py --help
```

## 测试

```bat
.venv\Scripts\python.exe -m unittest 发票统计工具\test_invoice_stats.py
Windows启动统计.bat --check-only
```

```bash
.venv/bin/python -m unittest 发票统计工具/test_invoice_stats.py
./启动统计.command --check-only
```

## 目录说明

- `发票统计工具/invoice_stats.py`：正式统计入口。
- `发票统计工具/test_invoice_stats.py`：核心解析测试。
- `发票统计工具/requirements.txt`：Python 依赖。
- `发票统计工具/参考脚本/`：首次统计过程中使用的文本提取、PDF 接触表、工作簿检查和 Codex 工作簿写入脚本，完整保留供排查参考。

## 限制

- 目前支持含可提取文字层的中国电子发票 PDF；纯图片扫描件会报错，需要先做 OCR。
- 项目名称采用规则提取。新增发票版式差异较大时，请先运行 `--check-only` 并人工复核结果。
- `参考脚本` 中的 `.mjs` 文件依赖 Codex 的 `@oai/artifact-tool` 运行时；日常使用不需要运行这些文件。
