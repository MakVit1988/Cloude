# kompas-mcp

MCP-сервер для КОМПАС-3D (v22+): Claude Desktop / Claude Code управляют КОМПАС через COM (API7 + API5).

> **Статус: черновик, не проверен на живом КОМПАС.** Каркас MCP, режимы доступа и проверка путей покрыты тестами
> на заглушке. Вызовы COM написаны по памяти об API КОМПАС и **не запускались**. Места с неуверенными сигнатурами
> (переменные, МЦХ, свойства, 2D-размеры, `SaveAs` в PDF/STEP) перебирают варианты вызова, а при неудаче пишут,
> что именно не сработало. Для отладки есть инструмент `inspect_com`.

## Установка (Windows, на машине с КОМПАС-3D)

```bash
git clone https://github.com/MakVit1988/Cloude && cd Cloude
git checkout claude/mcp-kompas-3d-server-it9wq8
python -m venv .venv
.venv\Scripts\pip install -e .
```

Claude Desktop, `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "kompas-3d": {
      "command": "C:\\path\\to\\Cloude\\.venv\\Scripts\\kompas-mcp.exe",
      "env": { "KOMPAS_MCP_ROOTS": "D:\\Projects" }
    }
  }
}
```

Claude Code: `claude mcp add kompas-3d -- C:\path\to\Cloude\.venv\Scripts\kompas-mcp.exe`

## Безопасность (переменные окружения)

| Переменная | Эффект |
|---|---|
| `KOMPAS_MCP_READONLY=1` | не регистрируются инструменты, меняющие модели и файлы |
| `KOMPAS_MCP_ROOTS=D:\A;D:\B` | работа с файлами только внутри этих каталогов |
| `KOMPAS_MCP_ALLOW_MACROS=1` | включает `run_library_command` и `execute_command` (по умолчанию выключены) |

## Инструменты

Чтение: `kompas_version`, `list_documents`, `get_properties`, `get_variables`, `get_mass_properties`,
`get_assembly_structure`, `get_bom`, `export_document` (pdf, dxf, dwg, step, stl, igs, sat, x_t), `inspect_com`,
`open_document`, `activate_document`, `close_document`.

Изменение: `create_document`, `save_document`, `save_document_as`, `set_variable`,
`sketch_extrude` (эскиз из rect/circle/line/polygon + выдавливание или вырезание),
`draw_line`, `draw_circle`, `draw_linear_dimension`.

Макросы (только с `KOMPAS_MCP_ALLOW_MACROS=1`): `run_library_command`, `execute_command`.

Не реализовано: чтение самого документа-спецификации (`.spw`), состав сборки берётся из дерева 3D-модели.

## Разработка и тесты

```bash
pip install -e ".[dev]"
pytest
```
