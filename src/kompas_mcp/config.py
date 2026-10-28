"""Настройки из переменных окружения и проверка путей."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    #: Только чтение: инструменты, меняющие модели и файлы, не регистрируются.
    readonly: bool = False
    #: Разрешить запуск макросов/библиотек и команд КОМПАС.
    allow_macros: bool = False
    #: Если непусто, работать можно только с файлами внутри этих каталогов.
    roots: tuple[Path, ...] = field(default_factory=tuple)

    @classmethod
    def from_env(cls) -> "Settings":
        raw = os.environ.get("KOMPAS_MCP_ROOTS", "")
        roots = tuple(Path(p).resolve() for p in raw.split(os.pathsep) if p.strip())
        return cls(
            readonly=_flag("KOMPAS_MCP_READONLY"),
            allow_macros=_flag("KOMPAS_MCP_ALLOW_MACROS"),
            roots=roots,
        )

    def check_path(self, path: str) -> str:
        """Вернуть абсолютный путь или бросить ошибку, если он вне разрешённых каталогов."""
        p = Path(path).expanduser().resolve()
        if self.roots and not any(p == r or r in p.parents for r in self.roots):
            allowed = ", ".join(str(r) for r in self.roots)
            raise PermissionError(f"Путь {p} вне разрешённых каталогов (KOMPAS_MCP_ROOTS): {allowed}")
        return str(p)
