"""
Modelos de datos para el Meeting Generator.

Define las dataclasses que representan la estructura de una semana
de reunión y el contrato mínimo necesario para generar el formulario.
"""

import re
from dataclasses import dataclass, field


def _is_positive_integer(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


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
            isinstance(p.name, str) and p.name.strip() for p in self.participants
        )

    def participant_count(self) -> int:
        """Número de participantes con nombre no vacío."""
        return sum(
            1
            for p in self.participants
            if isinstance(p.name, str) and p.name.strip()
        )

    def first_participant_name(self) -> str:
        """Nombre del primer participante, o cadena vacía."""
        for p in self.participants:
            if isinstance(p.name, str) and p.name.strip():
                return p.name
        return ""

    def second_participant_name(self) -> str:
        """Nombre del segundo participante, o cadena vacía."""
        count = 0
        for p in self.participants:
            if isinstance(p.name, str) and p.name.strip():
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
        names = [
            p.name
            for p in self.participants
            if isinstance(p.name, str) and p.name.strip()
        ]
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

    Las tres secciones conservan las asignaciones extraídas del documento:
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

    def validation_errors(self) -> list[str]:
        """Devuelve todos los incumplimientos del contrato de una semana."""
        errors: list[str] = []

        required_fields = (
            ("fecha", self.date),
            ("lectura semanal", self.weekly_reading),
            ("presidente", self.president),
            ("oración inicial", self.opening_prayer),
            ("oración final", self.closing_prayer),
        )
        for field_name, value in required_fields:
            if not isinstance(value, str) or not value.strip():
                errors.append(f"falta {field_name}")

        songs = (
            ("canción inicial", self.opening_song),
            ("canción intermedia", self.intermediate_song),
            ("canción final", self.closing_song),
        )
        for field_name, value in songs:
            normalized_value = value.strip() if isinstance(value, str) else ""
            if not normalized_value:
                errors.append(f"falta {field_name}")
            elif not normalized_value.isdecimal() or int(normalized_value) <= 0:
                errors.append(f"{field_name} debe ser un número positivo")

        if len(self.bible_treasures) != 3:
            errors.append(
                "Tesoros de la Biblia debe contener exactamente los puntos 1, 2 y 3"
            )
        elif [assignment.number for assignment in self.bible_treasures] != [1, 2, 3]:
            errors.append("Tesoros de la Biblia debe estar numerado 1, 2 y 3")

        if not self.ministry:
            errors.append("Seamos Mejores Maestros no contiene asignaciones")

        christian_life_normal = [
            assignment
            for assignment in self.christian_life
            if not assignment.is_bible_study and not assignment.is_conclusion
        ]
        bible_studies = [
            assignment
            for assignment in self.christian_life
            if assignment.is_bible_study
        ]
        conclusions = [
            assignment
            for assignment in self.christian_life
            if assignment.is_conclusion
        ]

        if not christian_life_normal:
            errors.append("Nuestra Vida Cristiana no contiene asignaciones")
        if len(bible_studies) != 1:
            errors.append("debe haber exactamente un estudio bíblico")
        if len(conclusions) != 1:
            errors.append("debe haber exactamente una conclusión")
        elif not self.christian_life[-1].is_conclusion:
            errors.append("la conclusión debe ser la última asignación")

        numbered_assignments = [
            assignment
            for assignment in self.all_assignments()
            if not assignment.is_conclusion
        ]
        numbers = [assignment.number for assignment in numbered_assignments]
        if any(not _is_positive_integer(number) for number in numbers):
            errors.append("todas las asignaciones deben tener un número positivo")
        else:
            if len(set(numbers)) != len(numbers):
                errors.append("la numeración de las asignaciones contiene duplicados")
            if numbers != sorted(numbers):
                errors.append("la numeración de las asignaciones no es ascendente")

        section_assignments = (
            ("Tesoros de la Biblia", self.bible_treasures),
            ("Seamos Mejores Maestros", self.ministry),
            ("Nuestra Vida Cristiana", self.christian_life),
        )
        for section_name, assignments in section_assignments:
            for position, assignment in enumerate(assignments, start=1):
                label = self._assignment_label(section_name, position, assignment)
                errors.extend(self._assignment_validation_errors(label, assignment))

                participant_count = assignment.participant_count()
                if any(
                    not isinstance(participant.name, str)
                    or not participant.name.strip()
                    for participant in assignment.participants
                ):
                    errors.append(f"{label} contiene un participante sin nombre")

                if section_name != "Nuestra Vida Cristiana" and (
                    assignment.is_bible_study or assignment.is_conclusion
                ):
                    errors.append(f"{label} tiene un tipo incompatible con su sección")
                if assignment.is_bible_study and assignment.is_conclusion:
                    errors.append(f"{label} no puede ser estudio y conclusión a la vez")

                if section_name == "Tesoros de la Biblia" and participant_count != 1:
                    errors.append(f"{label} debe tener exactamente un participante")
                elif section_name == "Seamos Mejores Maestros" and not (
                    1 <= participant_count <= 2
                ):
                    errors.append(f"{label} debe tener uno o dos participantes")
                elif assignment.is_bible_study and participant_count != 2:
                    errors.append(f"{label} debe tener conductor y lector")
                elif (
                    section_name == "Nuestra Vida Cristiana"
                    and not assignment.is_bible_study
                    and participant_count != 1
                ):
                    errors.append(f"{label} debe tener exactamente un participante")

        return errors

    def is_valid(self) -> bool:
        """
        Valida el contrato completo requerido para generar el formulario.

        Returns:
            True si no se detecta ningún error de validación.
        """
        return not self.validation_errors()

    def all_assignments(self) -> list[Assignment]:
        """
        Retorna todas las asignaciones de la semana en orden:
        bible_treasures + ministry + christian_life.
        """
        return self.bible_treasures + self.ministry + self.christian_life

    def _assignment_label(
        self,
        section_name: str,
        position: int,
        assignment: Assignment,
    ) -> str:
        if assignment.is_bible_study:
            return "estudio bíblico"
        if assignment.is_conclusion:
            return "conclusión"
        number = (
            assignment.number if _is_positive_integer(assignment.number) else position
        )
        return f"{section_name}, punto {number}"

    def _assignment_validation_errors(
        self,
        label: str,
        assignment: Assignment,
    ) -> list[str]:
        errors: list[str] = []
        if not isinstance(assignment.title, str) or not assignment.title.strip():
            errors.append(f"{label}: falta el título")
        if not _is_positive_integer(assignment.duration_minutes):
            errors.append(f"{label}: falta una duración positiva")
        if (
            not isinstance(assignment.duration_text, str)
            or not assignment.duration_text.strip()
        ):
            errors.append(f"{label}: falta el texto de duración")
        else:
            duration_match = re.fullmatch(
                r"\s*(\d+)\s*(?:mins?\.?|minutos?)\s*",
                assignment.duration_text,
                re.IGNORECASE,
            )
            if duration_match is None:
                errors.append(
                    f"{label}: el texto de duración tiene un formato inválido"
                )
            elif int(duration_match.group(1)) != assignment.duration_minutes:
                errors.append(
                    f"{label}: la duración escrita no coincide con los minutos"
                )
        return errors
