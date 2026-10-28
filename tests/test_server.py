"""Тесты MCP-слоя на заглушке: настоящий КОМПАС нужен только на Windows."""

import pytest
from mcp.shared.memory import create_connected_server_and_client_session

from kompas_mcp.backend import KompasError, _expand_shapes
from kompas_mcp.config import Settings
from kompas_mcp.server import build_server


class FakeKompas:
    def __init__(self):
        self.calls = []

    def __getattr__(self, name):
        def call(*args, **kwargs):
            self.calls.append((name, args))
            return {"called": name, "args": list(args)}

        return call


async def tool_names(settings):
    server = build_server(FakeKompas(), settings)
    async with create_connected_server_and_client_session(server._mcp_server) as c:
        return {t.name for t in (await c.list_tools()).tools}


async def test_full_toolset_without_macros():
    names = await tool_names(Settings())
    assert {"open_document", "sketch_extrude", "export_document", "set_variable"} <= names
    assert "run_library_command" not in names


async def test_macros_only_when_allowed():
    names = await tool_names(Settings(allow_macros=True))
    assert {"run_library_command", "execute_command"} <= names


async def test_readonly_hides_mutating_tools():
    names = await tool_names(Settings(readonly=True, allow_macros=True))
    assert "get_mass_properties" in names and "export_document" in names
    assert not names & {"create_document", "set_variable", "sketch_extrude", "save_document", "run_library_command"}


async def test_tool_call_reaches_backend():
    fake = FakeKompas()
    server = build_server(fake, Settings())
    async with create_connected_server_and_client_session(server._mcp_server) as c:
        r = await c.call_tool("export_document", {"format": "step", "target": "C:/out/a.step"})
        assert not r.isError
        r = await c.call_tool("export_document", {"format": "exe", "target": "x"})
        assert r.isError
    assert fake.calls[0][0] == "export_document"


def test_path_policy(tmp_path):
    s = Settings(roots=(tmp_path.resolve(),))
    s.check_path(str(tmp_path / "a.m3d"))
    with pytest.raises(PermissionError):
        s.check_path(str(tmp_path / ".." / "evil.m3d"))
    Settings().check_path("anything/at/all.m3d")


def test_expand_shapes():
    segs = _expand_shapes([{"type": "rect", "x": 0, "y": 0, "w": 2, "h": 1}, {"type": "circle", "cx": 0, "cy": 0, "r": 1}])
    assert len(segs) == 5 and segs[0] == ("line", 0, 0, 2, 0) and segs[-1][0] == "circle"
    with pytest.raises(KompasError):
        _expand_shapes([{"type": "rect", "x": 0}])
    with pytest.raises(KompasError):
        _expand_shapes([{"type": "polygon", "points": [[0, 0], [1, 1]]}])
    with pytest.raises(KompasError):
        _expand_shapes([{"type": "blob"}])
