"""Preferencias persistentes de la interfaz (últimas carpetas y plantilla)."""

import json
import logging
import os
import tempfile
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

APP_DIRECTORY_NAME = "jw-program-formatter"


def default_settings_path() -> Path:
    """Devuelve la ruta XDG de configuración del usuario."""
    base = os.environ.get("XDG_CONFIG_HOME")
    config_home = Path(base) if base else Path.home() / ".config"
    return config_home / APP_DIRECTORY_NAME / "settings.json"


class AppSettings:
    """Almacén clave-valor en JSON; nunca interrumpe la aplicación si falla."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = path if path is not None else default_settings_path()
        self._values: dict[str, Any] = self._load()

    def _load(self) -> dict[str, Any]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return {}
        except (OSError, ValueError) as exc:
            logger.warning("No se pudieron leer las preferencias: %s", exc)
            return {}
        return data if isinstance(data, dict) else {}

    def get_path(self, key: str) -> Path | None:
        """Devuelve una ruta guardada solo si sigue existiendo."""
        value = self._values.get(key)
        if not isinstance(value, str) or not value:
            return None
        path = Path(value)
        try:
            return path if path.exists() else None
        except OSError:
            return None

    def set_path(self, key: str, path: Path) -> None:
        """Guarda una ruta y persiste las preferencias."""
        self._values[key] = str(path)
        self._save()

    def _save(self) -> None:
        temporary_name: str | None = None
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(
                "w",
                encoding="utf-8",
                dir=self.path.parent,
                delete=False,
            ) as temporary:
                temporary_name = temporary.name
                json.dump(self._values, temporary, ensure_ascii=False, indent=2)
            os.replace(temporary_name, self.path)
        except OSError as exc:
            logger.warning("No se pudieron guardar las preferencias: %s", exc)
            if temporary_name is not None:
                try:
                    os.unlink(temporary_name)
                except OSError:
                    pass
