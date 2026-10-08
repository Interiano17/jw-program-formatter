"""
Módulo de configuración centralizada para el Meeting Generator.

Contiene los patrones y ajustes compartidos que usan el parser y el writer,
incluido el catálogo efectivo de marcadores pendientes de la plantilla.
"""

import logging
import os
import re

# ---------------------------------------------------------------------------
# Configuración de archivos y logging
# ---------------------------------------------------------------------------

OUTPUT_FILENAME: str = "S-140_COMPLETADO.docx"

# Nombre mostrado en el encabezado de cada página del formulario. Configure el
# nombre real fuera del repositorio mediante MEETING_GENERATOR_CONGREGATION_NAME.
CONGREGATION_NAME: str = os.environ.get(
    "MEETING_GENERATOR_CONGREGATION_NAME",
    "Congregación Loarque",
)

LOG_FILENAME: str = "meeting_generator.log"

LOG_LEVEL: int = logging.INFO
LOG_MAX_BYTES: int = 5 * 1024 * 1024
LOG_BACKUP_COUNT: int = 3

LOG_FORMAT: str = "%(asctime)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"

# ---------------------------------------------------------------------------
# Patrones Regex compilados
# ---------------------------------------------------------------------------

# Detección del inicio de una semana
# Ejemplos:
#   "SEMANA DEL 6  AL 12  JULIO. JEREMÍAS 13-15"
#   "SEMANA DEL13  AL 19 JULIO. JEREMÍAS, 16, 17"
#   "SEMANA DEL 27 DE JULIO  AL  2  DE  AGOSTO."
#   "SEMANA DEL 3 AL 9 AGOSTO.JEREMÍAS 22, 23"
RE_WEEK_HEADER: re.Pattern = re.compile(r"^\s*SEMANA\s+DEL\s*\d+.*$", re.IGNORECASE)

# Extracción del contenido de fecha de la cabecera de semana.
# El parser limita este contenido por la posición de la lectura semanal.
RE_DATE: re.Pattern = re.compile(
    r"^\s*SEMANA\s+DEL\s*(.*)$",
    re.IGNORECASE,
)

# Los 66 libros en la forma utilizada por el programa en español.
BIBLE_BOOK_NAMES: tuple[str, ...] = (
    "GÉNESIS",
    "ÉXODO",
    "LEVÍTICO",
    "NÚMEROS",
    "DEUTERONOMIO",
    "JOSUÉ",
    "JUECES",
    "RUT",
    "1 SAMUEL",
    "2 SAMUEL",
    "1 REYES",
    "2 REYES",
    "1 CRÓNICAS",
    "2 CRÓNICAS",
    "ESDRAS",
    "NEHEMÍAS",
    "ESTER",
    "JOB",
    "SALMOS",
    "PROVERBIOS",
    "ECLESIASTÉS",
    "CANTAR DE LOS CANTARES",
    "ISAÍAS",
    "JEREMÍAS",
    "LAMENTACIONES",
    "EZEQUIEL",
    "DANIEL",
    "OSEAS",
    "JOEL",
    "AMÓS",
    "ABDÍAS",
    "JONÁS",
    "MIQUEAS",
    "NAHÚM",
    "HABACUC",
    "SOFONÍAS",
    "AGEO",
    "ZACARÍAS",
    "MALAQUÍAS",
    "MATEO",
    "MARCOS",
    "LUCAS",
    "JUAN",
    "HECHOS",
    "ROMANOS",
    "1 CORINTIOS",
    "2 CORINTIOS",
    "GÁLATAS",
    "EFESIOS",
    "FILIPENSES",
    "COLOSENSES",
    "1 TESALONICENSES",
    "2 TESALONICENSES",
    "1 TIMOTEO",
    "2 TIMOTEO",
    "TITO",
    "FILEMÓN",
    "HEBREOS",
    "SANTIAGO",
    "1 PEDRO",
    "2 PEDRO",
    "1 JUAN",
    "2 JUAN",
    "3 JUAN",
    "JUDAS",
    "APOCALIPSIS",
)

_BIBLE_BOOK_ALIASES: tuple[str, ...] = (
    "EL CANTAR DE LOS CANTARES",
    "HECHOS DE LOS APÓSTOLES",
    "CANTAR",
    "CANTARES",
    "HAGEO",
)

_BIBLE_BOOK_PATTERN = "|".join(
    re.escape(book).replace(r"\ ", r"\s+")
    for book in sorted(
        BIBLE_BOOK_NAMES + _BIBLE_BOOK_ALIASES,
        key=len,
        reverse=True,
    )
)

# Extracción de la lectura semanal (libro + capítulos).
# Ejemplos: "JEREMÍAS 13-15", "1 CORINTIOS 1-2", "OSEAS 1-3".
RE_WEEKLY_READING: re.Pattern = re.compile(
    rf"\b(?:{_BIBLE_BOOK_PATTERN})\s*[,;:]?\s*"
    r"\d+(?:\s*[-–—,;:.]\s*\d+)*\b(?!\s*[-–—,;:])",
    re.IGNORECASE,
)

# Detección del presidente y canción inicial
# Ejemplo: "PRESIDENTE: ANA EJEMPLO. CANCIÓN 123:"
# También: "PRESIDENTE. CARLOS EJEMPLO. CANCIÓN 44"
RE_PRESIDENT_SONG: re.Pattern = re.compile(
    r"PRESIDENTE\s*[:\.,]\s*[\.,]?\s*(.+?)\s*[\.,]?\s*CANCI[OÓ]N\s*(\d+)", re.IGNORECASE
)

# Detección de canción intermedia o final
# Ejemplo: "CANCIÓN 49", "CANCIÓN  22"
RE_SONG: re.Pattern = re.compile(r"^\s*CANCI[OÓ]N\s*(\d+)\s*$", re.IGNORECASE)

# Separador de nombres: estudiante // ayudante o conductor // lector
RE_NAME_SEPARATOR: re.Pattern = re.compile(r"\s*//\s*")

# Detección de nombres con sufijo entre paréntesis como (p) o (h)
RE_NAME_SUFFIX: re.Pattern = re.compile(r"\s*\([^)]*\)\s*$")

# Palabras de conclusión
RE_CONCLUSION_WORDS: re.Pattern = re.compile(
    r"^\s*PALABRAS\s+DE\s+CONCLUSI[OÓ]N[\.\s]*.*$", re.IGNORECASE
)

# Oración final
RE_FINAL_PRAYER: re.Pattern = re.compile(r"^\s*ORACI[OÓ]N\s+FINAL\s*$", re.IGNORECASE)

# Detección de "Estudio bíblico de la congregación"
RE_BIBLE_STUDY: re.Pattern = re.compile(
    r"^\s*Estudio\s+b[ií]blico\s+de\s+la\s*congregaci[oó]n[\.\s]*.*$", re.IGNORECASE
)

# ---------------------------------------------------------------------------
# Marcadores en la plantilla S-140 que deben ser reemplazados
# ---------------------------------------------------------------------------

# Marcadores variables que se limpian en las filas sobrantes.
PLACEHOLDERS: tuple[str, ...] = (
    "[FECHA]",
    "[Nombre]",
    "[Título]",
    "[Nombre/Nombre]",
    "Canción [Número]",
    "[Número]",
    "[X mins.]",
    "[XX mins.]",
    "(X mins.)",
    "(XX mins.)",
    "[NOMBRE DE LA CONGREGACIÓN]",
)

# Catálogo completo para detectar contenido pendiente antes de guardar.
# Los rótulos fijos, como "Presidente:" u "Oración:", deben permanecer.
TEMPLATE_MARKERS: tuple[str, ...] = (*PLACEHOLDERS, "LECTURA SEMANAL DE LA BIBLIA")

# ---------------------------------------------------------------------------
# Constantes de tiempo por defecto para puntos
# ---------------------------------------------------------------------------

# Duraciones usadas para el horario y la inferencia de discursos (en minutos).
DEFAULT_DURATIONS: dict[str, int] = {
    "introduccion": 1,
    "conclusion": 3,
    "ministerio_discurso": 5,
}

# Duración total prevista para las asignaciones de Nuestra Vida Cristiana
# anteriores al estudio bíblico. Permite completar una única duración omitida
# en el documento fuente sin inventar tiempos cuando hay más ambigüedad.
CHRISTIAN_LIFE_PRE_STUDY_MINUTES: int = 15

# Reglas para calcular las horas de inicio mostradas en el formulario.
MEETING_START_MINUTES: int = 19 * 60
SONG_DURATION_MINUTES: int = 4
OPENING_PRAYER_DURATION_MINUTES: int = 1
MINISTRY_TRANSITION_MINUTES: int = 1

# ponytail: duración excepcional ligada a fecha y lectura; usar metadatos
# de la fuente si se añaden más excepciones de horario.
INTERMEDIATE_SONG_DURATION_OVERRIDES: dict[tuple[str, str], int] = {
    ("7 al 13 de septiembre", "Jeremías 32-33"): 6,
}
