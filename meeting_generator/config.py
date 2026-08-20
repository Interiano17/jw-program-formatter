"""
Módulo de configuración centralizada para el Meeting Generator.

Contiene todas las constantes, expresiones regulares compiladas,
palabras clave y configuraciones utilizadas por el parser y el writer.
"""

import re
import logging
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuración de archivos y logging
# ---------------------------------------------------------------------------

# Nombre del archivo de salida generado
OUTPUT_FILENAME: str = "S-140_COMPLETADO.docx"

# Nombre mostrado en el encabezado de cada página del formulario.
CONGREGATION_NAME: str = "Congregación Loarque"

# Nombre del archivo de log
LOG_FILENAME: str = "meeting_generator.log"

# Nivel de logging por defecto
LOG_LEVEL: int = logging.INFO

# Formato del log
LOG_FORMAT: str = "%(asctime)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"

# ---------------------------------------------------------------------------
# Extensiones de archivo permitidas
# ---------------------------------------------------------------------------

ALLOWED_EXTENSIONS: tuple[str, ...] = (".docx",)

# ---------------------------------------------------------------------------
# Patrones Regex compilados
# ---------------------------------------------------------------------------

# Detección del inicio de una semana
# Ejemplos:
#   "SEMANA DEL 6  AL 12  JULIO. JEREMÍAS 13-15"
#   "SEMANA DEL13  AL 19 JULIO. JEREMÍAS, 16, 17"
#   "SEMANA DEL 27 DE JULIO  AL  2  DE  AGOSTO."
#   "SEMANA DEL 3 AL 9 AGOSTO.JEREMÍAS 22, 23"
RE_WEEK_HEADER: re.Pattern = re.compile(
    r"^\s*SEMANA\s+DEL\s*\d+.*$",
    re.IGNORECASE
)

# Extracción de la fecha de la cabecera de semana
# Captura todo el texto después de "SEMANA DEL "
RE_DATE: re.Pattern = re.compile(
    r"SEMANA\s+DEL\s*(.+?)(?:\.\s*JEREMÍAS|\s*JEREMÍAS|\.\s*$|\s*$)",
    re.IGNORECASE
)

# Extracción de la lectura semanal (libro + capítulos)
# Ejemplos: "JEREMÍAS 13-15", "JEREMÍAS, 16, 17", "JEREMÍAS 18,19"
RE_WEEKLY_READING: re.Pattern = re.compile(
    r"(JEREMÍAS|ISAÍAS|EZEQUIEL|DANIEL|GÉNESIS|ÉXODO|LEVÍTICO|NÚMEROS|DEUTERONOMIO|"
    r"MATEO|MARCOS|LUCAS|JUAN|HECHOS|ROMANOS|CORINTIOS|GÁLATAS|EFESIOS|FILIPENSES|"
    r"COLOSENSES|TESALONICENSES|TIMOTEO|TITO|FILEMÓN|HEBREOS|SANTIAGO|PEDRO|"
    r"APOCALIPSIS|PROVERBIOS|SALMOS|JOB|CANTAR|RUT|LAMENTACIONES|ESTER|ECLESIASTÉS|"
    r"JOEL|AMÓS|ABDÍAS|JONÁS|MIQUEAS|NAHÚM|HABACUC|SOFONÍAS|HAGEO|ZACARÍAS|MALAQUÍAS|"
    r"JOSUÉ|JUECES|SAMUEL|REYES|CRÓNICAS|ESDRAS|NEHEMÍAS|ÉXODO|LEVÍTICO)\s*[\d,\-:\s;]+",
    re.IGNORECASE
)

# Detección del presidente y canción inicial
# Ejemplo: "PRESIDENTE: GERMAN CERRATO. CANCIÓN 123:"
# También: "PRESIDENTE. MIGUEL HERNÁNDEZ. CANCIÓN 44"
RE_PRESIDENT_SONG: re.Pattern = re.compile(
    r"PRESIDENTE\s*[:\.,]\s*[\.,]?\s*(.+?)\s*[\.,]?\s*CANCI[OÓ]N\s*(\d+)",
    re.IGNORECASE
)

# Detección de secciones principales del programa
RE_SECTION_TREASURES: re.Pattern = re.compile(
    r"^\s*TESOROS\s+DE\s+(LA|LAS)\s+BIBLIA[SE]?[S]?\s*$",
    re.IGNORECASE
)

RE_SECTION_MINISTRY: re.Pattern = re.compile(
    r"^\s*SEAMOS\s+MEJORES\s+MAESTROS[\.\s]*$",
    re.IGNORECASE
)

RE_SECTION_CHRISTIAN_LIFE: re.Pattern = re.compile(
    r"^\s*NUESTRA\s+VIDA\s+CRISTIANA[\.\s]*$",
    re.IGNORECASE
)

# Detección de números de punto (1 al 10 como marcadores de sección)
RE_POINT_NUMBER: re.Pattern = re.compile(
    r"^\s*(\d{1,2})\s*$"
)

# Detección de título de punto con duración entre paréntesis
# Ejemplo: "Jehová merece que le obedezcamos . (10 minutos)"
RE_POINT_TITLE_WITH_DURATION: re.Pattern = re.compile(
    r"^(.+?)\s*[\(\.]\s*(\d+)\s*(?:minutos?|mins?\.?)\s*\)?\s*$",
    re.IGNORECASE
)

# Detección de duración en cualquier texto entre paréntesis
RE_DURATION: re.Pattern = re.compile(
    r"\(?\s*(\d+)\s*(?:minutos?|mins?\.?)\s*\)?",
    re.IGNORECASE
)

# Detección de canción intermedia o final
# Ejemplo: "CANCIÓN 49", "CANCIÓN  22"
RE_SONG: re.Pattern = re.compile(
    r"^\s*CANCI[OÓ]N\s*(\d+)\s*$",
    re.IGNORECASE
)

# Separador de nombres: estudiante // ayudante o conductor // lector
RE_NAME_SEPARATOR: re.Pattern = re.compile(
    r"\s*//\s*"
)

# Detección de nombres con sufijo entre paréntesis como (p) o (h)
RE_NAME_SUFFIX: re.Pattern = re.compile(
    r"\s*\([^)]*\)\s*$"
)

# Palabras de conclusión
RE_CONCLUSION_WORDS: re.Pattern = re.compile(
    r"^\s*PALABRAS\s+DE\s+CONCLUSI[OÓ]N[\.\s]*.*$",
    re.IGNORECASE
)

# Oración final
RE_FINAL_PRAYER: re.Pattern = re.compile(
    r"^\s*ORACI[OÓ]N\s+FINAL\s*$",
    re.IGNORECASE
)

# Palabras de introducción
RE_INTRODUCTION: re.Pattern = re.compile(
    r"^\s*ORACI[OÓ]N[,\.]?\s*PALABRAS\s+DE\s+INTRODUCCI[OÓ]N\s*$",
    re.IGNORECASE
)

# Detección de "Estudio bíblico de la congregación"
RE_BIBLE_STUDY: re.Pattern = re.compile(
    r"^\s*Estudio\s+b[ií]blico\s+de\s+la\s*congregaci[oó]n[\.\s]*.*$",
    re.IGNORECASE
)

# Detección de "haga discípulos"
RE_MAKE_DISCIPLES: re.Pattern = re.compile(
    r"^\s*Haga\s+disc[ií]pulos[\.\s]*.*$",
    re.IGNORECASE
)

# Detección de "explique sus creencias"
RE_EXPLAIN_BELIEFS: re.Pattern = re.compile(
    r"^\s*Explique\s+sus\s+creencias[\.\s]*.*$",
    re.IGNORECASE
)

# Detección de "discurso"
RE_DISCOURSE: re.Pattern = re.compile(
    r"^\s*DISCURSO[\.\s]*.*$",
    re.IGNORECASE
)

# Detección de "empiece conversaciones"
RE_START_CONVERSATIONS: re.Pattern = re.compile(
    r"^\s*Empiece\s+conversaciones[\.\s]*.*$",
    re.IGNORECASE
)

# Detección de "haga revisitas"
RE_RETURN_VISITS: re.Pattern = re.compile(
    r"^\s*Haga\s+revisitas[\.\s]*.*$",
    re.IGNORECASE
)

# ---------------------------------------------------------------------------
# Marcadores en la plantilla S-140 que deben ser reemplazados
# ---------------------------------------------------------------------------

# Marcadores para búsqueda en la plantilla destino
TEMPLATE_MARKERS: dict[str, str] = {
    "fecha": "[FECHA]",
    "lectura_semanal": "LECTURA SEMANAL DE LA BIBLIA",
    "presidente": "Presidente:",
    "nombre_generico": "[Nombre]",
    "cancion": "Canción [Número]",
    "oracion": "Oración:",
    "titulo": "[Título]",
    "estudiante": "Estudiante:",
    "estudiante_ayudante": "Estudiante/Ayudante:",
    "conductor_lector": "Conductor/Lector:",
    "nombre_nombre": "[Nombre/Nombre]",
    "palabras_conclusion": "Palabras de conclusión",
    "nombre_congregacion": "[NOMBRE DE LA CONGREGACIÓN]",
}

# Orden de las secciones en la plantilla (para reemplazo secuencial)
TEMPLATE_SECTION_ORDER: list[str] = [
    "fecha",
    "presidente",
    "oracion",
    "cancion_inicio",
    "punto1_titulo",
    "punto1_nombre",
    "punto2_nombre",
    "punto3_estudiante",
    "punto4_titulo",
    "punto4_estudiante_ayudante",
    "punto5_titulo",
    "punto5_estudiante_ayudante",
    "punto6_titulo",
    "punto6_estudiante_ayudante",
    "punto7_titulo",
    "punto7_estudiante_ayudante",
    "cancion_intermedia",
    "punto8_titulo",
    "punto8_nombre",
    "punto9_titulo",
    "punto9_nombre",
    "punto10_conductor_lector",
    "cancion_final",
    "oracion_final",
    "conclusion_nombre",
]

# ---------------------------------------------------------------------------
# Constantes de tiempo por defecto para puntos
# ---------------------------------------------------------------------------

# Tiempos por defecto para cada tipo de punto (en minutos)
DEFAULT_DURATIONS: dict[str, int] = {
    "tesoros_punto1": 10,
    "tesoros_punto2": 10,
    "tesoros_punto3": 4,
    "estudio_biblico": 30,
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

# ---------------------------------------------------------------------------
# Límites y validaciones
# ---------------------------------------------------------------------------

# Máximo número de semanas que puede tener la plantilla
MAX_TEMPLATE_WEEKS: int = 10

# Mínimo de campos requeridos para considerar una semana válida
MIN_REQUIRED_FIELDS: int = 5
