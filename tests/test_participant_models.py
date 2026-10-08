"""Regresiones de identidades individuales entregadas mediante la API."""

import copy
import tempfile
import unittest
from pathlib import Path

from meeting_generator.models import NameCorrection, Participant
from meeting_generator.parser import parse_document
from meeting_generator.template_writer import fill_template
from tests.source_fixtures import build_primary_source
from tests.template_fixture import build_template

INVALID_NAMES = (
    "Ana Uno // Juan Dos",
    "Ana Uno // Juan Dos // Eva Tres",
    "Ana Uno /// Juan Dos",
    "Ana Uno / Juan Dos",
    "Ana Uno // ",
    " // Juan Dos",
    "//",
    "(h)",
)


class NameCorrectionParticipantTests(unittest.TestCase):
    def test_corrections_reject_ambiguous_original_and_replacement(self) -> None:
        for field in ("original_name", "corrected_name"):
            for name in INVALID_NAMES:
                with self.subTest(field=field, name=name):
                    values = {
                        "original_name": "Ana Ejemplo",
                        "corrected_name": "Ana Prueba",
                        "reason": "Verificado en una referencia",
                    }
                    values[field] = name

                    with self.assertRaisesRegex(ValueError, "corrección"):
                        NameCorrection(**values)

    def test_corrections_preserve_valid_spelling_suffix_and_reason(self) -> None:
        correction = NameCorrection(
            " ANA  EJEMPLO(h) ",
            "Ana Prueba (h)",
            "Referencias A // B",
        )

        self.assertEqual(correction.original_name, " ANA  EJEMPLO(h) ")
        self.assertEqual(correction.corrected_name, "Ana Prueba (h)")
        self.assertEqual(correction.reason, "Referencias A // B")


class MeetingWeekParticipantTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.directory = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.directory.cleanup)
        cls.root = Path(cls.directory.name)
        cls.weeks = parse_document(build_primary_source(cls.root / "source.docx"))
        cls.template = build_template(cls.root / "template.docx")

    def test_header_identities_require_one_person(self) -> None:
        fields = (
            ("president", "presidente"),
            ("opening_prayer", "oración inicial"),
            ("closing_prayer", "oración final"),
        )
        for field, label in fields:
            for name in INVALID_NAMES:
                with self.subTest(field=field, name=name):
                    week = copy.deepcopy(self.weeks[0])
                    setattr(week, field, name)

                    errors = week.validation_errors()

                    self.assertTrue(any(label in error for error in errors), errors)
                    self.assertFalse(week.is_valid())

    def test_every_assignment_identity_requires_one_person(self) -> None:
        for position in range(len(self.weeks[0].all_assignments())):
            for name in INVALID_NAMES:
                with self.subTest(position=position, name=name):
                    week = copy.deepcopy(self.weeks[0])
                    week.all_assignments()[position].participants[0].name = name

                    errors = week.validation_errors()

                    self.assertTrue(
                        any("participante 1" in error for error in errors), errors
                    )
                    self.assertFalse(week.is_valid())

    def test_invalid_second_participant_is_not_ignored(self) -> None:
        for section in ("ministry", "christian_life"):
            for name in ("", " ", *INVALID_NAMES):
                with self.subTest(section=section, name=name):
                    week = copy.deepcopy(self.weeks[0])
                    assignments = getattr(week, section)
                    assignment = next(
                        item for item in assignments if len(item.participants) == 2
                    )
                    assignment.participants[1] = Participant(name)

                    errors = week.validation_errors()

                    self.assertTrue(
                        any("participante" in error for error in errors), errors
                    )
                    self.assertFalse(week.is_valid())

    def test_validation_does_not_rewrite_names_or_parse_nonperson_fields(self) -> None:
        week = copy.deepcopy(self.weeks[0])
        week.date = "Semana // de prueba"
        week.weekly_reading = "Lectura // de prueba"
        week.president = " ANA  EJEMPLO(h) "
        week.bible_treasures[0].participants[0].name = " CARLOS EJEMPLO(p) "

        self.assertEqual(week.validation_errors(), [])
        self.assertEqual(week.president, " ANA  EJEMPLO(h) ")
        self.assertEqual(
            week.bible_treasures[0].participants[0].name, " CARLOS EJEMPLO(p) "
        )

    def test_direct_fill_rejects_invalid_names_and_preserves_previous_output(
        self,
    ) -> None:
        output = self.root / "output.docx"
        previous_output = b"salida anterior"
        for target in ("header", "assignment"):
            for name in INVALID_NAMES:
                with self.subTest(target=target, name=name):
                    weeks = copy.deepcopy(self.weeks)
                    if target == "header":
                        weeks[0].president = name
                    else:
                        weeks[0].ministry[0].participants[1].name = name
                    output.write_bytes(previous_output)

                    with self.assertLogs(
                        "meeting_generator.template_writer", level="ERROR"
                    ):
                        result = fill_template(self.template, weeks, output)

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, len(weeks))
                    self.assertIsNone(result.output_path)
                    self.assertTrue(
                        any("semana 1" in error.lower() for error in result.errors),
                        result.errors,
                    )
                    self.assertEqual(output.read_bytes(), previous_output)
