# excel-template-filler

## 项目介绍

Excel 多模板自动填充系统。用户放入一个主 Excel 文件、多个模板 Excel 文件和一个 `mapping_config.json`，系统会按配置从主表提取数据、清洗数据、执行简单计算，并安全写入模板副本。

系统设计原则：

- 不修改原始主表和原始模板。
- 写入前复制模板，所有写入只发生在 `output/filled_files/` 的副本中。
- 写入时只修改 `cell.value`。
- 不修改字体、字号、颜色、边框、合并单元格、行高、列宽、公式、图片、页眉页脚、打印区域和页面设置。
- 字段找不到、公式覆盖、明细行不足等问题都会写入报告。

## 安装方法

建议使用 Python 3.10 或更高版本。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 目录说明

- `config/`：配置文件和自然语言规则文件。
- `input/main_excel/`：主 Excel 文件目录，只放一个主表。
- `input/templates/`：模板 Excel 文件目录，可放多个模板。
- `output/filled_files/`：填充后的模板副本。
- `output/logs/`：结构分析、填充计划、提取数据、计算日志、流程日志。
- `output/reports/`：异常报告和校验报告。
- `src/`：核心 Python 模块。
- `tests/`：pytest 测试用例。
- `docs/`：项目设计、用户指南、开发者指南。

## 如何放置主表

把主 Excel 放到：

```text
input/main_excel/
```

要求：

- 目录中应只有一个主表文件。
- 支持 `.xlsx`、`.xlsm`。
- `.xls` 会自动尝试转换为 `.xlsx`，转换日志见 `output/logs/conversion_log.xlsx`。

## 如何放置模板

把模板 Excel 放到：

```text
input/templates/
```

要求：

- 可以放多个模板。
- 支持 `.xlsx`、`.xlsm`。
- `.xls` 会先备份到 `output/backups/`，再转换到 `output/converted/` 后处理。

## 如何配置 mapping_config.json

参考：

```text
config/mapping_config.example.json
```

复制为：

```text
config/mapping_config.json
```

核心字段：

- `main_file`：主表文件名。
- `template_name`：模板文件名。
- `output_name`：填充后输出文件名。
- `source_sheet`：主表来源 Sheet。
- `target_sheet`：模板目标 Sheet。
- `mappings`：单值字段映射。
- `detail_area`：明细区域映射。
- `calculations`：简单计算规则。

## 如何运行分析

```powershell
python main.py analyze
```

输出：

- `output/logs/main_excel_structure.json`
- `output/logs/template_structure.json`

## 如何生成填充计划

```powershell
python main.py plan --config config/mapping_config.json
```

输出：

- `output/logs/fill_plan.json`
- `output/logs/fill_plan.xlsx`
- `output/reports/config_error_report.xlsx`

## 如何提取、清洗和计算数据

```powershell
python main.py extract --config config/mapping_config.json
```

输出：

- `output/logs/extracted_data.json`
- `output/logs/calculation_log.xlsx`
- `output/reports/data_error_report.xlsx`

## 如何执行填充

```powershell
python main.py fill --config config/mapping_config.json
```

输出：

- `output/filled_files/`
- `output/reports/error_report.xlsx`
- `output/summary.txt`

## 一键完整流程

```powershell
python main.py run --config config/mapping_config.json
```

完整流程会执行：生成计划、提取数据、清洗计算、安全填充、输出校验。

输出：

- `output/filled_files/`
- `output/reports/validation_report.xlsx`
- `output/reports/error_report.xlsx`
- `output/logs/process_log.xlsx`
- `output/summary.txt`

## 如何查看报告

- `output/summary.txt`：本次处理摘要。
- `output/reports/error_report.xlsx`：异常报告。
- `output/reports/validation_report.xlsx`：输出校验报告。
- `output/logs/process_log.xlsx`：流程日志。
- `output/logs/fill_plan.xlsx`：人工可读填充计划。

## 自然语言规则解析

示例文件：

```text
config/sample_rules_prompt.txt
```

运行：

```powershell
python main.py parse-rules --rules config/sample_rules_prompt.txt
```

输出：

- `config/mapping_config.draft.json`
- `config/mapping_config.need_confirm.json`

自然语言解析只生成配置草稿，不会写入 Excel。人工确认后，才能作为正式 `mapping_config.json` 使用。

## Web 页面

启动：

```powershell
python -m uvicorn web_app:app --host 127.0.0.1 --port 8000
```

打开：

```text
http://127.0.0.1:8000/
```

页面支持上传主表、多个模板，并直接输入自然语言规则。上传后系统会自动生成 `config/mapping_config.json` 和 `config/mapping_config.need_confirm.json`，再通过按钮执行分析、生成填充计划、执行填充，并下载 `output` 压缩包。

## 常见错误说明

- `input/main_excel must contain exactly one supported main Excel file`：主表目录没有文件，或放了多个主表。
- `target_template does not exist`：配置中的模板文件名和 `input/templates/` 中的文件名不一致。
- `source_sheet does not exist`：配置中的来源 Sheet 不在主表中。
- `source_field does not exist`：配置中的字段名与主表表头不一致。
- `target cell contains formula and overwrite_formula is false`：目标单元格有公式，默认不能覆盖。
- `detail rows exceed target range`：明细数据行数超过模板配置区域，超出部分不会写入。
- `.xls conversion failed`：当前环境没有可用 Excel 或 LibreOffice，或文件无法转换。

## 测试

```powershell
pytest
```
