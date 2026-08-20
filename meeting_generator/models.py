"""
Modelos de datos para el Meeting Generator.

Define las dataclasses que representan la estructura de una semana
de reunión. Arquitectura completamente dinámica: no hay campos fijos
ni límites predefinidos para ninguna sección.
"""

from dataclasses import dataclass, field


@dataclass
class Participant:
    """
    Representa un participante en una asignación.

    El rol puede ser: "Estudiante", "Ayudante", "Conductor", "Lector",
    "Sala auxiliar", "Auditorio principal", o vacío si no se determina.
    """

    name: str = ""
    role: str = ""

    def __str__(self) -> str:
        return self.name


@dataclass
class Assignment:
    """
    Representa una asignación individual del programa.

    Campos:
        number: Número exacto del documento fuente (1, 2, 3... sin límite).
        title: Título o descripción de la asignación.
        duration_minutes: Duración en minutos (entero).
        duration_text: Texto original de duración ("10 mins.", "4 minutos", etc.).
        participants: Lista de participantes con sus roles.
        is_bible_study: True si es el estudio bíblico de la congregación.
        is_conclusion: True si es "Palabras de conclusión".
        is_prayer: True si es la oración final.
        raw_lines: Líneas originales de la celda (para debug).
    """

    number: int = 0
    title: str = ""
    duration_minutes: int = 0
    duration_text: str = ""
    participants: list[Participant] = field(default_factory=list)
    is_bible_study: bool = False
    is_conclusion: bool = False
    is_prayer: bool = False
    raw_lines: list[str] = field(default_factory=list)

    def has_participants(self) -> bool:
        """Indica si la asignación tiene al menos un participante."""
        return len(self.participants) > 0 and any(
            p.name.strip() for p in self.participants
        )

    def participant_count(self) -> int:
        """Número de participantes con nombre no vacío."""
        return sum(1 for p in self.participants if p.name.strip())

    def first_participant_name(self) -> str:
        """Nombre del primer participante, o cadena vacía."""
        for p in self.participants:
            if p.name.strip():
                return p.name
        return ""

    def second_participant_name(self) -> str:
        """Nombre del segundo participante, o cadena vacía."""
        count = 0
        for p in self.participants:
            if p.name.strip():
                count += 1
                if count == 2:
                    return p.name
        return ""

    def formatted_participants(self) -> str:
        """
        Formatea los participantes para mostrar en la plantilla.

        Si hay 2 participantes: "Nombre1 / Nombre2"
        Si hay 1 participante: "Nombre1"
        Si no hay: cadena vacía.
        """
        names = [p.name for p in self.participants if p.name.strip()]
        if len(names) == 0:
            return ""
        elif len(names) == 1:
            return names[0]
        else:
            return " / ".join(names)


@dataclass
class MeetingWeek:
    """
    Representa una semana completa del programa de reunión.

    Las tres secciones son listas dinámicas sin límite de elementos:
        bible_treasures: Puntos que aparecen bajo "TESOROS DE LA BIBLIA".
        ministry: Puntos bajo "SEAMOS MEJORES MAESTROS".
        christian_life: Puntos bajo "NUESTRA VIDA CRISTIANA".
    """

    week_index: int = 0
    raw_header: str = ""
    date: str = ""
    weekly_reading: str = ""
    president: str = ""
    opening_song: str = ""
    opening_prayer: str = ""

    bible_treasures: list[Assignment] = field(default_factory=list)
    ministry: list[Assignment] = field(default_factory=list)
    christian_life: list[Assignment] = field(default_factory=list)

    intermediate_song: str = ""
    closing_song: str = ""
    closing_prayer: str = ""

    def is_valid(self) -> bool:
        """
        Valida que la semana tenga los campos mínimos requeridos.

        Returns:
            True si la semana tiene al menos fecha, presidente
            y algún contenido en tesoros.
        """
        has_date = bool(self.date.strip())
        has_president = bool(self.president.strip())
        has_treasures = len(self.bible_treasures) > 0 and any(
            a.has_participants() for a in self.bible_treasures
        )
        return has_date and has_president and has_treasures

    def all_assignments(self) -> list[Assignment]:
        """
        Retorna todas las asignaciones de la semana en orden:
        bible_treasures + ministry + christian_life.
        """
        return self.bible_treasures + self.ministry + self.christian_life
