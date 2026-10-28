"""Слой работы с КОМПАС-3D через COM (API7 + API5).

Работает только на Windows с установленным КОМПАС-3D v22+ и пакетом pywin32.
Вызовы, в точных сигнатурах которых нет уверенности, обёрнуты в _first_ok():
он перебирает варианты и при неудаче сообщает, что именно не сработало.
Для разбора на конкретной версии есть inspect_com().
"""

from __future__ import annotations

import os
import threading
from typing import Any, Callable, Iterable

from .config import Settings

# DocumentTypeEnum
DOC_TYPES = {
    "drawing": 1,
    "fragment": 2,
    "specification": 3,
    "part": 4,
    "assembly": 5,
    "text": 6,
}
DOC_TYPE_NAMES = {v: k for k, v in DOC_TYPES.items()}

# Форматы, которые пробуем сохранять через IKompasDocument.SaveAs (по расширению).
EXPORT_FORMATS = {"pdf", "dxf", "dwg", "step", "stp", "stl", "igs", "iges", "sat", "x_t"}
FORMATS_3D = {"step", "stp", "stl", "igs", "iges", "sat", "x_t"}
FORMATS_2D = {"pdf", "dxf", "dwg"}

PLANES = {"XOY": "o3d_planeXOY", "XOZ": "o3d_planeXOZ", "YOZ": "o3d_planeYOZ"}

# Запасные значения констант, если библиотеки констант не удалось подключить.
FALLBACK_CONSTANTS = {
    "o3d_planeXOY": 1,
    "o3d_planeXOZ": 2,
    "o3d_planeYOZ": 3,
    "o3d_sketch": 5,
    "o3d_baseExtrusion": 24,
    "o3d_bossExtrusion": 25,
    "o3d_cutExtrusion": 26,
    "pTop_Part": -1,
    "dtNormal": 0,
    "dtReverse": 1,
    "dtBoth": 2,
    "etBlind": 0,
}

# D3FormatConvType (SaveAsToAdditionFormat)
FORMAT_3D_CODES = {"sat": 1, "x_t": 2, "step": 3, "stp": 3, "igs": 4, "iges": 4, "stl": 6}

API7_TLB = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"
API5_TLB = "{0422828C-F174-495E-AC5D-D31014DBBE87}"


class KompasError(RuntimeError):
    """Ошибка работы с КОМПАС (понятное сообщение для клиента MCP)."""


def _first_ok(label: str, attempts: Iterable[tuple[str, Callable[[], Any]]]) -> Any:
    errors = []
    for name, fn in attempts:
        try:
            return fn()
        except Exception as e:  # noqa: BLE001 — COM бросает разные типы
            errors.append(f"{name}: {e}")
    raise KompasError(
        f"{label}: ни один вариант вызова не сработал ({'; '.join(errors)}). "
        "Используйте inspect_com, чтобы посмотреть доступные члены в вашей версии КОМПАС."
    )


def _scalar(result: Any) -> Any:
    """COM-методы с out-параметрами возвращают кортеж; берём первое небулево значение."""
    if isinstance(result, tuple):
        for item in result:
            if not isinstance(item, bool):
                return item
        return result[0] if result else None
    return result


def _iter(coll: Any) -> list[Any]:
    if coll is None:
        return []
    try:
        return list(coll)
    except Exception:  # noqa: BLE001
        pass
    count = int(coll.Count)
    for getter in ("Item", "Part", "Document"):
        fn = getattr(coll, getter, None)
        if fn is None:
            continue
        try:
            return [fn(i) for i in range(count)]
        except Exception:  # noqa: BLE001
            continue
    raise KompasError("Не удалось перечислить элементы коллекции")


class Kompas:
    def __init__(self, settings: Settings):
        self.settings = settings
        self._lock = threading.RLock()
        self._m7 = None
        self._m5 = None
        self._app = None
        self._api5 = None
        self._consts: Any = None

    # ------------------------------------------------------------------ COM

    def _connect(self) -> None:
        if self._app is not None:
            return
        if os.name != "nt":
            raise KompasError("КОМПАС-3D доступен только в Windows: запустите сервер на машине с КОМПАС.")
        try:
            import pythoncom
            from win32com.client import Dispatch, gencache
        except ImportError as e:  # pragma: no cover
            raise KompasError("Не установлен pywin32: pip install pywin32") from e
        pythoncom.CoInitialize()
        self._m7 = self._ensure_module(API7_TLB, "KompasAPI7")
        raw = Dispatch("Kompas.Application.7")
        api7 = self._m7.IKompasAPIObject(
            raw._oleobj_.QueryInterface(self._m7.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch)
        )
        self._app = api7.Application
        self._app.Visible = True
        self._consts = _load_constants(gencache)

    def _ensure_module(self, guid: str, desc_hint: str):
        from win32com.client import gencache

        try:
            return gencache.EnsureModule(guid, 0, 1, 0)
        except Exception:  # noqa: BLE001
            from win32com.client import selecttlb

            for info in selecttlb.EnumTlbs():
                if desc_hint.lower() in (info.desc or "").lower():
                    return gencache.EnsureModule(info.clsid, info.lcid, int(info.major), int(info.minor))
            raise KompasError(f"Не найдена type library {desc_hint} — установлен ли КОМПАС-3D?")

    @property
    def app(self):
        self._connect()
        return self._app

    @property
    def api5(self):
        self._connect()
        if self._api5 is None:
            import pythoncom
            from win32com.client import Dispatch

            self._m5 = self._ensure_module(API5_TLB, "Kompas6API5")
            raw = Dispatch("Kompas.Application.5")
            self._api5 = self._m5.KompasObject(
                raw._oleobj_.QueryInterface(self._m5.KompasObject.CLSID, pythoncom.IID_IDispatch)
            )
        return self._api5

    def const(self, name: str) -> int:
        if self._consts is not None:
            try:
                return int(getattr(self._consts, name))
            except Exception:  # noqa: BLE001
                pass
        if name in FALLBACK_CONSTANTS:
            return FALLBACK_CONSTANTS[name]
        raise KompasError(f"Неизвестная константа КОМПАС: {name}")

    def cast(self, obj: Any, iface: str) -> Any:
        """Привести COM-объект к интерфейсу API7 (QueryInterface)."""
        import pythoncom

        self._connect()
        cls = getattr(self._m7, iface)
        return cls(obj._oleobj_.QueryInterface(cls.CLSID, pythoncom.IID_IDispatch))

    # ------------------------------------------------------------ документы

    def _doc_info(self, doc: Any) -> dict:
        t = None
        try:
            t = int(doc.DocumentType)
        except Exception:  # noqa: BLE001
            pass
        return {
            "name": _safe(lambda: doc.Name),
            "path": _safe(lambda: doc.PathName),
            "type": DOC_TYPE_NAMES.get(t, t),
        }

    def _find(self, path: str | None) -> Any:
        """Документ по пути (из открытых) или активный."""
        if path is None:
            doc = self.app.ActiveDocument
            if doc is None:
                raise KompasError("Нет активного документа")
            return doc
        want = os.path.normcase(self.settings.check_path(path))
        for doc in _iter(self.app.Documents):
            if os.path.normcase(str(_safe(lambda d=doc: d.PathName) or "")) == want:
                return doc
        raise KompasError(f"Документ не открыт: {path}")

    def list_documents(self) -> dict:
        with self._lock:
            active = _safe(lambda: self.app.ActiveDocument.PathName)
            docs = [self._doc_info(d) for d in _iter(self.app.Documents)]
            return {"active": active, "documents": docs}

    def open_document(self, path: str, read_only: bool = False) -> dict:
        with self._lock:
            path = self.settings.check_path(path)
            if not os.path.isfile(path):
                raise KompasError(f"Файл не найден: {path}")
            doc = self.app.Documents.Open(path, True, read_only)
            if doc is None:
                raise KompasError(f"КОМПАС не смог открыть {path}")
            return self._doc_info(doc)

    def create_document(self, doc_type: str, save_as: str | None = None) -> dict:
        with self._lock:
            if doc_type not in DOC_TYPES:
                raise KompasError(f"Тип документа: один из {sorted(DOC_TYPES)}")
            if save_as:
                save_as = self.settings.check_path(save_as)
            doc = self.app.Documents.Add(DOC_TYPES[doc_type], True)
            if doc is None:
                raise KompasError("КОМПАС не создал документ")
            if save_as:
                self._save_as(doc, save_as)
            return self._doc_info(doc)

    def save_document(self, path: str | None = None) -> dict:
        with self._lock:
            doc = self._find(path)
            if not _safe(lambda: doc.PathName):
                raise KompasError("У документа нет файла: используйте save_document_as")
            doc.Save()  # void
            return self._doc_info(doc)

    def _save_as(self, doc: Any, target: str) -> None:
        os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
        doc.SaveAs(target)  # void: успех проверяем по файлу
        if not os.path.isfile(target):
            raise KompasError(f"SaveAs не создал файл: {target}")

    def save_document_as(self, target: str, path: str | None = None) -> dict:
        with self._lock:
            doc = self._find(path)
            self._save_as(doc, self.settings.check_path(target))
            return self._doc_info(doc)

    def close_document(self, path: str | None = None, save: bool = False) -> dict:
        with self._lock:
            doc = self._find(path)
            info = self._doc_info(doc)
            doc.Close(1 if save else 0)  # kdSaveChanges / kdDoNotSaveChanges
            return {"closed": info, "saved": save}

    def activate_document(self, path: str) -> dict:
        with self._lock:
            doc = self._find(path)
            doc.Active = True
            return self._doc_info(doc)

    # ------------------------------------------- свойства и переменные модели

    def get_properties(self, path: str | None = None) -> dict:
        with self._lock:
            doc = self._find(path)
            result: dict[str, Any] = {"document": self._doc_info(doc), "properties": {}}
            try:
                mng = self.cast(self.app, "IPropertyMng")
                keeper = self.cast(doc, "IPropertyKeeper")
                raw = mng.GetProperties(doc)
                props = list(raw) if isinstance(raw, (tuple, list)) else [raw]
                for prop in props:
                    name = _safe(lambda p=prop: p.Name)
                    value = _safe(lambda p=prop: _prop_value(keeper, p))
                    if name:
                        result["properties"][name] = value
            except Exception as e:  # noqa: BLE001
                result["properties_error"] = str(e)
            return result

    def _top_part(self, doc: Any) -> Any:
        try:
            return self.cast(doc, "IKompasDocument3D").TopPart
        except Exception as e:  # noqa: BLE001
            raise KompasError("Документ не является 3D-моделью (деталью или сборкой)") from e

    def _as(self, obj: Any, iface: str) -> Any:
        """QueryInterface к iface; если не вышло, вернуть объект как есть."""
        try:
            return self.cast(obj, iface)
        except Exception:  # noqa: BLE001
            return obj

    def _variables(self, part: Any) -> list[Any]:
        # Проверено на v22: у IPart7 нет Variables, переменные читаются через IFeature7.
        feat = self._as(part, "IFeature7")
        count = _safe(lambda: int(feat.VariablesCount))
        if count is None:
            raise KompasError("Не удалось получить число переменных (IFeature7.VariablesCount)")
        return [feat.Variable(False, False, i) for i in range(count)]

    def get_variables(self, path: str | None = None) -> dict:
        with self._lock:
            doc = self._find(path)
            part = self._top_part(doc)
            out = []
            for v in self._variables(part):
                out.append(
                    {
                        "name": _safe(lambda v=v: v.Name),
                        "value": _safe(lambda v=v: v.Value),
                        "expression": _safe(lambda v=v: v.Expression),
                        "note": _safe(lambda v=v: v.Note),
                        "external": _safe(lambda v=v: v.External),
                    }
                )
            return {"document": self._doc_info(doc), "variables": out}

    def set_variable(self, name: str, expression: str, path: str | None = None) -> dict:
        with self._lock:
            doc = self._find(path)
            part = self._top_part(doc)
            for v in self._variables(part):
                if _safe(lambda v=v: v.Name) == name:
                    v.Expression = expression
                    _first_ok(
                        "Перестроение модели",
                        [
                            ("RebuildDocument", lambda: self.cast(doc, "IKompasDocument3D").RebuildDocument()),
                            ("Part.RebuildModel", lambda: part.RebuildModel()),
                        ],
                    )
                    # после перестроения старый объект хранит прежнее значение — читаем заново
                    fresh = next((x for x in self._variables(part) if _safe(lambda x=x: x.Name) == name), v)
                    return {
                        "name": name,
                        "expression": _safe(lambda: fresh.Expression),
                        "value": _safe(lambda: fresh.Value),
                    }
            raise KompasError(f"Переменная не найдена: {name}")

    def get_mass_properties(self, path: str | None = None) -> dict:
        """МЦХ через IMassInertiaParam7 (доп. интерфейс IPart7, берётся QueryInterface)."""
        with self._lock:
            doc = self._find(path)
            part = self._top_part(doc)
            res: dict[str, Any] = {}
            try:
                mi = self.cast(part, "IMassInertiaParam7")
                mi.Calculate()
                fields = [
                    "Mass", "Volume", "Area", "Density", "Xc", "Yc", "Zc",
                    "Jx", "Jy", "Jz", "Jxy", "Jxz", "Jyz", "Jx0", "Jy0", "Jz0",
                    "Lx", "Ly", "Lz", "Lxy", "Lxz", "Lyz", "LengthUnits", "MassUnits",
                ]
                res = {f: _safe(lambda f=f: getattr(mi, f)) for f in fields}
                res = {k: v for k, v in res.items() if v is not None}
            except Exception as e:  # noqa: BLE001
                mass = _safe(lambda: part.Mass)
                if mass is None:
                    raise KompasError(f"МЦХ недоступны: {e}") from e
                res = {"Mass": mass, "Density": _safe(lambda: part.Density), "note": f"IMassInertiaParam7 недоступен: {e}"}
            if not res:
                raise KompasError("Результат расчёта МЦХ пуст — проверьте материал и плотность детали")
            res["Material"] = _safe(lambda: part.Material)
            return {"document": self._doc_info(doc), "mass_properties": res}

    # ------------------------------------------------------ состав сборки

    def _part_node(self, part: Any, depth: int, max_depth: int) -> dict:
        part = self._as(part, "IPart7")
        node = {
            "name": _safe(lambda: part.Name),
            "marking": _safe(lambda: part.Marking),
            "file": _safe(lambda: part.FileName),
            "standard": _safe(lambda: part.Standard),
        }
        if depth < max_depth:
            children = _safe(lambda: _iter(part.Parts))
            if children:
                node["children"] = [self._part_node(c, depth + 1, max_depth) for c in children]
        return node

    def get_assembly_structure(self, path: str | None = None, max_depth: int = 10) -> dict:
        with self._lock:
            doc = self._find(path)
            top = self._top_part(doc)
            return {"document": self._doc_info(doc), "structure": self._part_node(top, 0, max_depth)}

    def get_bom(self, path: str | None = None) -> dict:
        """Плоская ведомость: одинаковые (обозначение, наименование, файл) складываются в количество."""
        tree = self.get_assembly_structure(path)["structure"]
        counts: dict[tuple, dict] = {}

        def walk(node: dict, root: bool = False) -> None:
            if not root:
                key = (node.get("marking"), node.get("name"), node.get("file"))
                row = counts.setdefault(key, {**{k: node.get(k) for k in ("marking", "name", "file", "standard")}, "qty": 0})
                row["qty"] += 1
            for c in node.get("children", []):
                walk(c)

        walk(tree, root=True)
        rows = sorted(counts.values(), key=lambda r: (str(r["marking"] or ""), str(r["name"] or "")))
        return {"assembly": tree.get("name"), "items": rows}

    # ----------------------------------------------------------- экспорт

    def export_document(self, fmt: str, target: str, path: str | None = None) -> dict:
        """3D: ksDocument3D.SaveAsToAdditionFormat (API5). 2D (pdf/dxf/dwg): IKompasDocument.SaveAs по расширению."""
        with self._lock:
            fmt = fmt.lower().lstrip(".")
            if fmt not in EXPORT_FORMATS:
                raise KompasError(f"Формат не поддерживается: {fmt}. Доступны: {sorted(EXPORT_FORMATS)}")
            doc = self._find(path)
            info = self._doc_info(doc)
            kind = info["type"]
            if fmt in FORMATS_3D and kind not in ("part", "assembly"):
                raise KompasError(f"Формат {fmt} — для деталей и сборок, а документ: {kind}")
            if fmt in FORMATS_2D and kind not in ("drawing", "fragment", "specification"):
                raise KompasError(f"Формат {fmt} — для чертежей/фрагментов/спецификаций, а документ: {kind}")
            target = self.settings.check_path(target)
            if os.path.splitext(target)[1].lower().lstrip(".") != fmt:
                target = f"{target}.{fmt}"
            os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
            if fmt in FORMATS_3D:
                if fmt not in FORMAT_3D_CODES:
                    raise KompasError(f"Формат {fmt} пока не поддержан для 3D-экспорта")
                doc.Active = True
                doc3d = self.api5.ActiveDocument3D()
                par = doc3d.AdditionFormatParam()
                par.Init()
                par.format = FORMAT_3D_CODES[fmt]
                if fmt == "stl":
                    par.formatBinary = True
                if not doc3d.SaveAsToAdditionFormat(target, par):
                    raise KompasError(f"SaveAsToAdditionFormat вернул FALSE для {target}")
            else:
                doc.SaveAs(target)
            if not os.path.isfile(target):
                raise KompasError(f"Файл не создан: {target}. Формат {fmt} не поддержан в вашей версии?")
            return {"exported": target, "size": os.path.getsize(target), "source": info}

    # --------------------------------------------------- построение 3D-модели

    def sketch_extrude(
        self,
        shapes: list[dict],
        depth: float,
        plane: str = "XOY",
        operation: str = "base",
        direction: str = "normal",
        path: str | None = None,
    ) -> dict:
        """Эскиз на плоскости из примитивов и операция выдавливания/вырезания (API5).

        base — первая операция в пустой детали, boss — приклеить к существующему телу, cut — вырезать.

        Размеры — в мм. shapes: rect {x,y,w,h}, circle {cx,cy,r}, line {x1,y1,x2,y2},
        polygon {points:[[x,y],...]} (замкнутый).
        """
        with self._lock:
            if plane not in PLANES:
                raise KompasError(f"plane: одна из {sorted(PLANES)}")
            if operation not in ("base", "boss", "cut"):
                raise KompasError("operation: base (первая операция детали), boss (добавить) или cut (вырезать)")
            if direction not in ("normal", "reverse", "both"):
                raise KompasError("direction: normal, reverse или both")
            if depth <= 0:
                raise KompasError("depth должна быть > 0")
            segments = _expand_shapes(shapes)
            if not segments:
                raise KompasError("shapes пуст")
            if path:
                self.activate_document(path)
            doc3d = self.api5.ActiveDocument3D()
            if doc3d is None:
                raise KompasError("Нет активной 3D-модели")
            part = doc3d.GetPart(self.const("pTop_Part"))
            plane_ent = part.GetDefaultEntity(self.const(PLANES[plane]))
            sketch = part.NewEntity(self.const("o3d_sketch"))
            sdef = sketch.GetDefinition()
            sdef.SetPlane(plane_ent)
            if not sketch.Create():
                raise KompasError("Не удалось создать эскиз")
            s2d = sdef.BeginEdit()
            try:
                for seg in segments:
                    if seg[0] == "line":
                        s2d.ksLineSeg(seg[1], seg[2], seg[3], seg[4], 1)
                    else:
                        s2d.ksCircle(seg[1], seg[2], seg[3], 1)
            finally:
                sdef.EndEdit()
            kind = {"base": "o3d_baseExtrusion", "boss": "o3d_bossExtrusion", "cut": "o3d_cutExtrusion"}[operation]
            ext = part.NewEntity(self.const(kind))
            edef = ext.GetDefinition()
            edef.directionType = self.const({"normal": "dtNormal", "reverse": "dtReverse", "both": "dtBoth"}[direction])
            edef.SetSideParam(True, self.const("etBlind"), depth, 0, False)
            if direction == "both":
                edef.SetSideParam(False, self.const("etBlind"), depth, 0, False)
            edef.SetSketch(sketch)
            if not ext.Create():
                raise KompasError("Операция выдавливания не создана (проверьте замкнутость контура и глубину)")
            _safe(lambda: doc3d.RebuildDocument())
            return {"ok": True, "operation": operation, "plane": plane, "depth": depth, "segments": len(segments)}

    # --------------------------------------------------------- 2D-чертёж

    def _drawing_container(self, doc: Any) -> Any:
        doc2d = self.cast(doc, "IKompasDocument2D")
        view = doc2d.ViewsAndLayersManager.Views.ActiveView
        return self.cast(view, "IDrawingContainer")

    def draw_line(self, x1: float, y1: float, x2: float, y2: float, path: str | None = None) -> dict:
        with self._lock:
            dc = self._drawing_container(self._find(path))
            seg = dc.LineSegments.Add()
            seg.X1, seg.Y1, seg.X2, seg.Y2 = x1, y1, x2, y2
            seg.Update()
            return {"ok": True}

    def draw_circle(self, cx: float, cy: float, radius: float, path: str | None = None) -> dict:
        with self._lock:
            dc = self._drawing_container(self._find(path))
            c = dc.Circles.Add()
            c.Xc, c.Yc, c.Radius = cx, cy, radius
            c.Update()
            return {"ok": True}

    def draw_linear_dimension(
        self, x1: float, y1: float, x2: float, y2: float, x3: float, y3: float, path: str | None = None
    ) -> dict:
        """Линейный размер между (x1,y1) и (x2,y2); (x3,y3) — положение размерной линии."""
        with self._lock:
            doc2d = self.cast(self._find(path), "IKompasDocument2D")
            view = doc2d.ViewsAndLayersManager.Views.ActiveView
            dim = self.cast(view, "ISymbols2DContainer").LineDimensions.Add()
            dim.X1, dim.Y1, dim.X2, dim.Y2, dim.X3, dim.Y3 = x1, y1, x2, y2, x3, y3
            dim.Update()
            return {"ok": True}

    # ---------------------------------------------------------- макросы

    def run_library_command(self, library_path: str, command: int) -> dict:
        """Подключить библиотеку КОМПАС (.rtw/.dll/.lcm…) и выполнить команду по номеру."""
        with self._lock:
            if not self.settings.allow_macros:
                raise PermissionError("Запуск макросов выключен: задайте KOMPAS_MCP_ALLOW_MACROS=1")
            library_path = self.settings.check_path(library_path)
            if not os.path.isfile(library_path):
                raise KompasError(f"Файл библиотеки не найден: {library_path}")
            k5 = self.api5
            lib_id = k5.ksAttachKompasLibrary(library_path)
            if not lib_id:
                raise KompasError("Не удалось подключить библиотеку")
            try:
                result = k5.ksExecuteKompasLibraryCommand(lib_id, command)
            finally:
                _safe(lambda: k5.ksDetachKompasLibrary(lib_id))
            return {"library": library_path, "command": command, "result": result}

    def execute_command(self, command_id: int) -> dict:
        """Выполнить встроенную команду КОМПАС по её номеру."""
        with self._lock:
            if not self.settings.allow_macros:
                raise PermissionError("Запуск команд выключен: задайте KOMPAS_MCP_ALLOW_MACROS=1")
            self.app.ExecuteKompasCommand(command_id, False)
            return {"command": command_id}

    # ------------------------------------------------------ диагностика

    def inspect_com(self, target: str = "app", path: str | None = None) -> dict:
        """Список членов COM-объекта: app, api5, document, part, view."""
        with self._lock:
            if target == "app":
                obj = self.app
            elif target == "api5":
                obj = self.api5
            elif target == "document":
                obj = self._find(path)
            elif target == "part":
                obj = self._top_part(self._find(path))
            else:
                raise KompasError("target: app, api5, document или part")
            names = sorted(n for n in dir(obj) if not n.startswith("_"))
            return {"target": target, "type": type(obj).__name__, "members": names}

    def version(self) -> dict:
        with self._lock:
            self._connect()
            # IApplication.GetSystemVersion(0) -> (ok, major, minor, release, build)
            raw = _safe(lambda: self.app.GetSystemVersion(0))
            version = None
            if isinstance(raw, tuple) and len(raw) >= 5:
                version = ".".join(str(n) for n in raw[1:5])
            return {"version": version, "visible": _safe(lambda: self.app.Visible)}


def _prop_value(keeper: Any, prop: Any) -> Any:
    """IPropertyKeeper.GetPropertyValue(Property, out Value, BaseUnit, out FromSource) -> (ok, value, fromSource)."""
    for args in (("", True, False), (None, True, False), (True,)):
        try:
            r = keeper.GetPropertyValue(prop, *args)
        except Exception:  # noqa: BLE001
            continue
        if isinstance(r, tuple):
            return r[1] if len(r) > 1 and isinstance(r[0], bool) else _scalar(r)
        return r
    return None


def _safe(fn: Callable[[], Any]) -> Any:
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return None


def _expand_shapes(shapes: list[dict]) -> list[tuple]:
    out: list[tuple] = []
    for s in shapes:
        t = s.get("type")
        try:
            if t == "rect":
                x, y, w, h = float(s["x"]), float(s["y"]), float(s["w"]), float(s["h"])
                pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
                out += [("line", *pts[i], *pts[(i + 1) % 4]) for i in range(4)]
            elif t == "circle":
                out.append(("circle", float(s["cx"]), float(s["cy"]), float(s["r"])))
            elif t == "line":
                out.append(("line", float(s["x1"]), float(s["y1"]), float(s["x2"]), float(s["y2"])))
            elif t == "polygon":
                pts = [(float(p[0]), float(p[1])) for p in s["points"]]
                if len(pts) < 3:
                    raise KompasError("polygon: нужно минимум 3 точки")
                out += [("line", *pts[i], *pts[(i + 1) % len(pts)]) for i in range(len(pts))]
            else:
                raise KompasError(f"Неизвестный тип фигуры: {t!r} (rect, circle, line, polygon)")
        except KeyError as e:
            raise KompasError(f"В фигуре {t!r} не хватает поля {e}") from e
    return out


def _load_constants(gencache) -> Any:  # pragma: no cover — только Windows
    """Подключить type library констант КОМПАС, чтобы брать значения по именам."""
    try:
        import win32com.client as wc
        from win32com.client import selecttlb

        for info in selecttlb.EnumTlbs():
            desc = (info.desc or "").lower()
            if "kompas6constants" in desc:
                gencache.EnsureModule(info.clsid, info.lcid, int(info.major), int(info.minor))
        return wc.constants
    except Exception:  # noqa: BLE001
        return None
