"""
Utilidades comunes para el Meeting Generator.

Funciones auxiliares para limpieza de texto, manejo de archivos,
configuración de logging y otras operaciones compartidas.
"""

import logging
import sys
from pathlib import Path
from typing import Optional

from .config import (
    LOG_DATE_FORMAT,
    LOG_FILENAME,
    LOG_FORMAT,
    LOG_LEVEL,
    RE_NAME_SEPARATOR,
    RE_NAME_SUFFIX,
)


def setup_logging(log_path: Optional[Path] = None) -> None:
    """
    Configura el sistema de logging para escribir a archivo y consola.

    Args:
        log_path: Ruta opcional para el archivo de log.
                  Si no se especifica, se usa el directorio actual.
    """
    if log_path is None:
        log_path = Path.cwd() / LOG_FILENAME

    # Crear handlers
    file_handler = logging.FileHandler(str(log_path), encoding="utf-8")
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT, LOG_DATE_FORMAT))

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(LOG_LEVEL)
    console_handler.setFormatter(logging.Formatter(LOG_FORMAT, LOG_DATE_FORMAT))

    # Configurar el logger raíz
    root_logger = logging.getLogger()
    root_logger.setLevel(LOG_LEVEL)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    root_logger.info("Sistema de logging configurado. Archivo: %s", log_path)


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

    # Reemplazar saltos de línea y tabulaciones por espacios
    cleaned = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")

    # Eliminar espacios múltiples
    cleaned = " ".join(cleaned.split())

    return cleaned.strip()


def normalize_person_name(text: str) -> str:
    """Convierte un nombre completamente en mayúsculas a formato título."""
    cleaned = clean_text(text)
    letters = [character for character in cleaned if character.isalpha()]
    if letters and all(character.isupper() for character in letters):
        return cleaned.title()
    return cleaned


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
    """
    if not text:
        return ("", "")

    parts = RE_NAME_SEPARATOR.split(text, maxsplit=1)

    primary = normalize_person_name(parts[0]) if len(parts) > 0 else ""
    secondary = normalize_person_name(parts[1]) if len(parts) > 1 else ""

    return (primary, secondary)


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


def is_empty_or_whitespace(text: Optional[str]) -> bool:
    """
    Verifica si un texto es None, vacío o solo contiene espacios.

    Args:
        text: Texto a verificar.

    Returns:
        True si el texto está vacío o es solo espacios.
    """
    if text is None:
        return True
    return len(text.strip()) == 0


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

    digits = "".join(c for c in text if c.isdigit())
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


def ensure_directory_exists(directory: Path) -> None:
    """
    Asegura que un directorio exista, creándolo si es necesario.

    Args:
        directory: Ruta del directorio a verificar.
    """
    directory.mkdir(parents=True, exist_ok=True)
