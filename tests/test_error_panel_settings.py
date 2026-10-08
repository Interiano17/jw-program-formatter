import json
import tempfile
import unittest
from pathlib import Path

from meeting_generator.error_panel import GENERAL_GROUP, group_errors
from meeting_generator.settings import AppSettings


class GroupErrorsTests(unittest.TestCase):
    def test_groups_by_week_in_numeric_order_with_general_first(self) -> None:
        groups = group_errors(
            [
                "semana 10: falta presidente",
                "la cantidad de encabezados (9) no coincide con la de tablas (8)",
                "semana 2: falta canción inicial",
                "semana 2: Nuestra Vida Cristiana, punto 7: falta duración",
                "semana 10: falta lectura semanal",
            ]
        )

        self.assertEqual(
            groups,
            [
                (
                    GENERAL_GROUP,
                    ["la cantidad de encabezados (9) no coincide con la de tablas (8)"],
                ),
                (
                    "Semana 2",
                    [
                        "Falta canción inicial",
                        "Nuestra Vida Cristiana, punto 7: falta duración",
                    ],
                ),
                ("Semana 10", ["Falta presidente", "Falta lectura semanal"]),
            ],
        )

    def test_keeps_row_context_and_ignores_blank_errors(self) -> None:
        groups = group_errors(["semana 3, fila 6: fila ambigua", "  "])

        self.assertEqual(groups, [("Semana 3", ["Fila 6: fila ambigua"])])


class AppSettingsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.settings_path = self.root / "config" / "settings.json"

    def test_round_trips_existing_paths_only(self) -> None:
        template = self.root / "template.docx"
        template.write_bytes(b"x")
        settings = AppSettings(self.settings_path)
        settings.set_path("template_file", template)
        settings.set_path("source_dir", self.root / "missing")

        reloaded = AppSettings(self.settings_path)

        self.assertEqual(reloaded.get_path("template_file"), template)
        self.assertIsNone(reloaded.get_path("source_dir"))
        self.assertIsNone(reloaded.get_path("unknown"))

    def test_corrupt_file_is_ignored(self) -> None:
        self.settings_path.parent.mkdir(parents=True)
        self.settings_path.write_text("{no es json", encoding="utf-8")

        settings = AppSettings(self.settings_path)

        self.assertIsNone(settings.get_path("template_file"))
        settings.set_path("source_dir", self.root)
        self.assertEqual(
            json.loads(self.settings_path.read_text(encoding="utf-8")),
            {"source_dir": str(self.root)},
        )

    def test_unwritable_location_does_not_raise(self) -> None:
        blocker = self.root / "blocker"
        blocker.write_text("x", encoding="utf-8")

        settings = AppSettings(blocker / "settings.json")
        settings.set_path("source_dir", self.root)


if __name__ == "__main__":
    unittest.main()
