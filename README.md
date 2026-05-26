# excel-template-filler

第一阶段：本地命令行版 Excel 多模板自动填充系统。

## 安装

```powershell
python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
pip install -e ".[dev]"
```

## 运行

```powershell
python -m excel_template_filler --config configs/example_config.json
```

如果需要生成示例 Excel 和配置：

```powershell
python examples/create_example_files.py
python -m excel_template_filler --config configs/example_config.json
```

## 测试

```powershell
pytest
```

## 配置说明

当前支持两类写入规则：

- `cell`：从主 Excel 某个 Sheet 的键值表或指定单元格取值，写入模板副本的指定单元格。
- `range`：从主 Excel 某个 Sheet 的表格区域取多行多列，写入模板副本的指定区域。

系统默认不会覆盖目标公式；如确需覆盖，需要在规则中设置 `allow_overwrite_formula: true`。

系统默认不会自动插入行；目标区域行数不足时会记录异常并跳过该规则。
