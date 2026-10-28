"""Регрессионные тесты backend на фейковых COM-объектах (поведение подтверждено на КОМПАС v22)."""

from types import SimpleNamespace

import pytest

from kompas_mcp.backend import Kompas, KompasError
from kompas_mcp.config import Settings


class FakeVar:
    def __init__(self, name, expression, value):
        self.Name, self.Expression, self.Value, self.Note, self.External = name, expression, value, "", False


class Snapshot:
    """Объект переменной: Value — снимок на момент выдачи, запись Expression уходит в модель."""

    def __init__(self, model_var):
        self._m = model_var
        self.Name, self.Value = model_var.Name, model_var.Value
        self.Note, self.External = "", False

    @property
    def Expression(self):
        return self._m.Expression

    @Expression.setter
    def Expression(self, v):
        self._m.Expression = v


class FakeFeature:
    """IFeature7: переменные только через VariablesCount и Variable(External, InSource, i)."""

    def __init__(self, variables, doc):
        self._vars, self._doc = variables, doc

    @property
    def VariablesCount(self):
        return len(self._vars)

    def Variable(self, external, in_source, index):
        return Snapshot(self._vars[index])


class FakePart:
    """IPart7 без Variables/GetVariables, как в реальном КОМПАС."""

    Marking = "ABC.001"
    FileName = "box.m3d"
    Standard = False
    Name = "box"
    Parts = ()


class FakeModelObject:
    """Дочерний компонент до приведения к IPart7: нет Marking/FileName."""

    Name = "box"

    def __init__(self, part):
        self.part = part


def make_kompas(top, variables=()):
    k = Kompas(Settings())
    doc = SimpleNamespace(PathName="C:/m.m3d", Name="m.m3d", DocumentType=4)
    rebuilt = []
    store = list(variables)

    def cast(obj, iface):
        if iface == "IKompasDocument3D":
            return SimpleNamespace(TopPart=top, RebuildDocument=lambda: rebuilt.append(1) or _apply(store))
        if iface == "IFeature7":
            return FakeFeature(store, doc)
        if iface == "IPart7":
            return obj.part if isinstance(obj, FakeModelObject) else obj
        raise AssertionError(iface)

    def _apply(vs):
        for v in vs:  # перестроение пересчитывает значения
            v.Value = float(eval(v.Expression, {}, {x.Name: x.Value for x in vs}))

    k.cast = cast
    k._find = lambda path=None: doc
    k._store, k._rebuilt = store, rebuilt
    return k


def test_get_variables_uses_ifeature7():
    k = make_kompas(FakePart(), [FakeVar("L", "40", 40.0), FakeVar("W", "L/2+1", 21.0)])
    out = k.get_variables()["variables"]
    assert [(v["name"], v["value"]) for v in out] == [("L", 40.0), ("W", 21.0)]


def test_set_variable_returns_fresh_value():
    k = make_kompas(FakePart(), [FakeVar("L", "25", 25.0)])
    res = k.set_variable("L", "40")
    assert res["value"] == 40.0 and res["expression"] == "40" and k._rebuilt


def test_set_variable_unknown_name():
    k = make_kompas(FakePart(), [FakeVar("L", "25", 25.0)])
    with pytest.raises(KompasError, match="не найдена"):
        k.set_variable("Q", "1")


def test_assembly_children_cast_to_ipart7():
    child = FakePart()
    top = SimpleNamespace(Name="asm", Marking="", FileName="asm.a3d", Standard=False,
                          Parts=[FakeModelObject(child), FakeModelObject(child)])
    k = make_kompas(top)
    bom = k.get_bom()["items"]
    assert bom == [{"marking": "ABC.001", "name": "box", "file": "box.m3d", "standard": False, "qty": 2}]
