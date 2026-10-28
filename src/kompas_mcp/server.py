"""MCP-сервер КОМПАС-3D: регистрация инструментов."""

from __future__ import annotations

from typing import Any, Literal

from mcp.server.fastmcp import FastMCP

from .backend import Kompas
from .config import Settings

DocType = Literal["part", "assembly", "drawing", "fragment", "specification", "text"]
ExportFormat = Literal["pdf", "dxf", "dwg", "step", "stp", "stl", "igs", "iges", "sat", "x_t"]


def build_server(backend: Any | None = None, settings: Settings | None = None) -> FastMCP:
    settings = settings or Settings.from_env()
    k = backend if backend is not None else Kompas(settings)
    mcp = FastMCP("kompas-3d")

    # ---------------------------------------------------------- чтение

    @mcp.tool()
    def kompas_version() -> dict:
        """Версия запущенного КОМПАС-3D (запускает КОМПАС, если он не запущен)."""
        return k.version()

    @mcp.tool()
    def list_documents() -> dict:
        """Открытые документы и путь активного."""
        return k.list_documents()

    @mcp.tool()
    def get_properties(path: str | None = None) -> dict:
        """Свойства документа (обозначение, наименование, материал и т.п.). Без path — активный документ."""
        return k.get_properties(path)

    @mcp.tool()
    def get_variables(path: str | None = None) -> dict:
        """Переменные 3D-модели: имя, значение, выражение, примечание."""
        return k.get_variables(path)

    @mcp.tool()
    def get_mass_properties(path: str | None = None) -> dict:
        """Массо-центровочные характеристики детали/сборки: масса, объём, центр масс, моменты инерции."""
        return k.get_mass_properties(path)

    @mcp.tool()
    def get_assembly_structure(path: str | None = None, max_depth: int = 10) -> dict:
        """Дерево состава сборки (наименование, обозначение, файл, признак стандартного изделия)."""
        return k.get_assembly_structure(path, max_depth)

    @mcp.tool()
    def get_bom(path: str | None = None) -> dict:
        """Плоская ведомость сборки с количеством одинаковых компонентов."""
        return k.get_bom(path)

    @mcp.tool()
    def inspect_com(target: Literal["app", "api5", "document", "part"] = "app", path: str | None = None) -> dict:
        """Диагностика: список членов COM-объекта КОМПАС (для отладки API на вашей версии)."""
        return k.inspect_com(target, path)

    @mcp.tool()
    def open_document(path: str, read_only: bool = False) -> dict:
        """Открыть документ КОМПАС (.m3d, .a3d, .cdw, .frw, .spw)."""
        return k.open_document(path, read_only)

    @mcp.tool()
    def activate_document(path: str) -> dict:
        """Сделать открытый документ активным."""
        return k.activate_document(path)

    @mcp.tool()
    def close_document(path: str | None = None, save: bool = False) -> dict:
        """Закрыть документ; save=true — сохранить изменения, иначе отбросить."""
        return k.close_document(path, save)

    @mcp.tool()
    def export_document(format: ExportFormat, target: str, path: str | None = None) -> dict:
        """Экспорт в pdf/dxf/dwg (чертежи) или step/stl/igs/sat/x_t (модели). Исходный документ не меняется."""
        return k.export_document(format, target, path)

    if settings.readonly:
        return mcp

    # --------------------------------------------------------- изменение

    @mcp.tool()
    def create_document(type: DocType, save_as: str | None = None) -> dict:
        """Создать документ. save_as — сразу сохранить по этому пути."""
        return k.create_document(type, save_as)

    @mcp.tool()
    def save_document(path: str | None = None) -> dict:
        """Сохранить документ (без path — активный)."""
        return k.save_document(path)

    @mcp.tool()
    def save_document_as(target: str, path: str | None = None) -> dict:
        """Сохранить документ под новым именем."""
        return k.save_document_as(target, path)

    @mcp.tool()
    def set_variable(name: str, expression: str, path: str | None = None) -> dict:
        """Изменить переменную модели (выражение, например '25' или 'L*2') и перестроить модель."""
        return k.set_variable(name, expression, path)

    @mcp.tool()
    def sketch_extrude(
        shapes: list[dict],
        depth: float,
        plane: Literal["XOY", "XOZ", "YOZ"] = "XOY",
        operation: Literal["base", "boss", "cut"] = "base",
        direction: Literal["normal", "reverse", "both"] = "normal",
        path: str | None = None,
    ) -> dict:
        """Создать эскиз и выдавить в активной 3D-модели. Размеры в мм.

        operation: base — первая операция в пустой детали, boss — приклеить к телу, cut — вырезать.

        shapes: список фигур —
          {"type":"rect","x":0,"y":0,"w":50,"h":30},
          {"type":"circle","cx":0,"cy":0,"r":10},
          {"type":"line","x1":0,"y1":0,"x2":10,"y2":0},
          {"type":"polygon","points":[[0,0],[10,0],[5,8]]}.
        Контур должен быть замкнутым.
        """
        return k.sketch_extrude(shapes, depth, plane, operation, direction, path)

    @mcp.tool()
    def draw_line(x1: float, y1: float, x2: float, y2: float, path: str | None = None) -> dict:
        """Отрезок на активном виде чертежа/фрагмента."""
        return k.draw_line(x1, y1, x2, y2, path)

    @mcp.tool()
    def draw_circle(cx: float, cy: float, radius: float, path: str | None = None) -> dict:
        """Окружность на активном виде чертежа/фрагмента."""
        return k.draw_circle(cx, cy, radius, path)

    @mcp.tool()
    def draw_linear_dimension(
        x1: float, y1: float, x2: float, y2: float, x3: float, y3: float, path: str | None = None
    ) -> dict:
        """Линейный размер между точками 1 и 2; точка 3 задаёт положение размерной линии."""
        return k.draw_linear_dimension(x1, y1, x2, y2, x3, y3, path)

    if settings.allow_macros:

        @mcp.tool()
        def run_library_command(library_path: str, command: int) -> dict:
            """Подключить библиотеку КОМПАС и выполнить её команду по номеру (макрос)."""
            return k.run_library_command(library_path, command)

        @mcp.tool()
        def execute_command(command_id: int) -> dict:
            """Выполнить встроенную команду КОМПАС по номеру."""
            return k.execute_command(command_id)

    return mcp


def main() -> None:
    build_server().run(transport="stdio")


if __name__ == "__main__":
    main()
