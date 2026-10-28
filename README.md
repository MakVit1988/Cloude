# kompas-mcp

MCP-сервер для КОМПАС-3D (v22+): Claude Desktop / Claude Code управляют КОМПАС через COM (API7 + API5).

> **Статус: проверено на живом КОМПАС-3D v22 (Windows, Python 3.14).**
>
> Работает: `kompas_version`, `list_documents`, `open_document`, `create_document`, `get_mass_properties`,
> `get_variables`, `set_variable`, `get_assembly_structure`, `get_bom`, `export_document` (STEP, STL, PDF, DXF, DWG),
> `sketch_extrude`, `draw_line`, `draw_circle`, `draw_linear_dimension` (PDF сверен глазами), работа по stdio,
> подключение к Claude Code (v2.1.311), режим `KOMPAS_MCP_READONLY` (12 инструментов вместо 20),
> скрытие и показ инструментов макросов по `KOMPAS_MCP_ALLOW_MACROS`.
>
> Не проверено: сам запуск макросов (`run_library_command`, `execute_command`), `inspect_com`, `save_document`,
> `save_document_as`, `close_document`, сборки с сопряжениями, Claude Desktop.
>
> Единицы МЦХ приходят в единицах КОМПАС (объём в см³, масса в г), не в мм.

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
`sketch_extrude` (эскиз из rect/circle/line/polygon + выдавливание: base — первая операция, boss — приклеить, cut — вырезать),
`draw_line`, `draw_circle`, `draw_linear_dimension`.

Макросы (только с `KOMPAS_MCP_ALLOW_MACROS=1`): `run_library_command`, `execute_command`.

Не реализовано: чтение самого документа-спецификации (`.spw`), состав сборки берётся из дерева 3D-модели.

## Разработка и тесты

```bash
pip install -e ".[dev]"
pytest
```
