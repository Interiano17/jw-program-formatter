"""
Utilidades comunes para el Meeting Generator.

Funciones auxiliares para limpieza de texto, manejo de archivos,
configuración de logging y otras operaciones compartidas.
"""

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

from .config import (
    LOG_BACKUP_COUNT,
    LOG_DATE_FORMAT,
    LOG_FILENAME,
    LOG_FORMAT,
    LOG_LEVEL,
    LOG_MAX_BYTES,
    RE_NAME_SEPARATOR,
    RE_NAME_SUFFIX,
)


def setup_logging(log_path: Optional[Path] = None) -> bool:
    """
    Configura el logging del paquete con rotación de archivo y salida a consola.

    Args:
        log_path: Ruta opcional para el archivo de log.
                  Si no se especifica, se usa el directorio actual.

    Returns:
        True si se abrió el archivo; False si se continúa con logging en consola.
    """
    formatter = logging.Formatter(LOG_FORMAT, LOG_DATE_FORMAT)
    package_logger = logging.getLogger("meeting_generator")
    handler_names = ("meeting_generator.file", "meeting_generator.console")
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(LOG_LEVEL)
    console_handler.setFormatter(formatter)

    try:
        if log_path is None:
            log_path = Path.cwd() / LOG_FILENAME
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT,
            encoding="utf-8",
        )
    except OSError as exc:
        if any(handler.name == handler_names[1] for handler in package_logger.handlers):
            console_handler.close()
        else:
            console_handler.set_name(handler_names[1])
            package_logger.setLevel(LOG_LEVEL)
            package_logger.propagate = False
            package_logger.addHandler(console_handler)
        package_logger.warning(
            "No se pudo abrir el archivo de log solicitado (%s). "
            "La aplicación continuará con logging en consola.",
            type(exc).__name__,
        )
        return False

    handlers: tuple[logging.Handler, ...] = (file_handler, console_handler)

    # Sustituye solo los handlers propios; el logger global pertenece al anfitrión.
    for handler in package_logger.handlers[:]:
        if handler.name in handler_names:
            package_logger.removeHandler(handler)
            handler.close()

    package_logger.setLevel(LOG_LEVEL)
    package_logger.propagate = False
    for name, handler in zip(handler_names, handlers):
        handler.set_name(name)
        handler.setLevel(LOG_LEVEL)
        handler.setFormatter(formatter)
        package_logger.addHandler(handler)

    package_logger.info("Logging configurado con rotación de archivo.")
    return True


def clean_text(text: str) -> str:
    """
    Limpia una cadena de texto: elimina espacios múltiples, tabulaciones,
    saltos de línea y espacios al inicio y final.

    Args:
        text: Texto a limpiar.

    Returns:
        Texto limpio y normalizado.
    """
    if not text:
        return ""

    cleaned = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")

    cleaned = " ".join(cleaned.split())

    return cleaned.strip()


def normalize_person_name(text: str) -> str:
    """Normaliza espacios y mayúsculas sin corregir la escritura del nombre."""
    cleaned = clean_text(text)
    suffix_match = RE_NAME_SUFFIX.search(cleaned)
    name = cleaned[: suffix_match.start()].strip() if suffix_match else cleaned
    suffix = suffix_match.group().strip() if suffix_match else ""
    letters = [character for character in name if character.isalpha()]
    if letters and all(character.isupper() for character in letters):
        name = name.title()
    return f"{name} {suffix}".strip()


def normalize_sentence_case(text: str) -> str:
    """Deja en mayúscula solo la primera letra de un texto."""
    return clean_text(text).capitalize()


def split_names(text: str) -> tuple[str, str]:
    """
    Divide un texto que contiene nombres separados por '//'.

    Ejemplos:
        'Carlos Pérez // Juan López' -> ('Carlos Pérez', 'Juan López')
        'Marvin Zúniga // Samuel Gómez' -> ('Marvin Zúniga', 'Samuel Gómez')
        'Lorena Cabrera' -> ('Lorena Cabrera', '')

    Args:
        text: Texto a dividir.

    Returns:
        Tupla (nombre_principal, nombre_ayudante).
        El ayudante será cadena vacía si no hay separador.

    Raises:
        ValueError: Si hay separadores inválidos, grupos vacíos o más de dos nombres.
    """
    text = clean_text(text)
    if not text:
        return ("", "")

    parts = RE_NAME_SEPARATOR.split(text)
    if len(parts) > 2:
        raise ValueError("El grupo debe contener como máximo dos participantes.")
    if any("/" in part for part in parts):
        raise ValueError("El único separador de participantes permitido es '//'.")
    names = [normalize_person_name(part) for part in parts]
    if any(not remove_name_suffix(name).strip() for name in names):
        raise ValueError("Cada grupo de participantes debe contener un nombre.")

    return (names[0], names[1] if len(names) == 2 else "")


def parse_single_name(text: str) -> str:
    """Valida y normaliza un campo que debe identificar a una sola persona."""
    primary, secondary = split_names(text)
    if not primary or secondary:
        raise ValueError("El campo debe contener exactamente un participante.")
    return primary


def remove_name_suffix(text: str) -> str:
    """
    Elimina sufijos entre paréntesis de un nombre, como (p) o (h).

    Args:
        text: Nombre que puede contener un sufijo.

    Returns:
        Nombre sin el sufijo.
    """
    if not text:
        return ""
    return RE_NAME_SUFFIX.sub("", text).strip()


def extract_number(text: str) -> Optional[int]:
    """
    Extrae el primer número encontrado en un texto.

    Args:
        text: Texto del que extraer el número.

    Returns:
        El número entero encontrado, o None si no hay número.
    """
    if not text:
        return None

    digits = "".join(character for character in text if character.isdigit())
    if digits:
        return int(digits)
    return None


def get_default_output_path(source_path: Path, template_path: Path) -> Path:
    """
    Genera la ruta sugerida inicialmente en el diálogo Guardar como.

    El usuario todavía puede elegir otro directorio o nombre antes de generar.

    Args:
        source_path: Ruta del documento fuente.
        template_path: Ruta de la plantilla.

    Returns:
        Ruta completa para el archivo de salida.
    """
    from .config import OUTPUT_FILENAME

    output_dir = template_path.parent
    return output_dir / OUTPUT_FILENAME
