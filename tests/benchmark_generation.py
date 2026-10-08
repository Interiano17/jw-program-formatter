"""Medición manual de nueve semanas con los mismos fixtures en cada ejecución."""

import hashlib
import json
import platform
import statistics
import tempfile
from importlib.metadata import version
from pathlib import Path
from time import perf_counter
from zipfile import ZipFile

from meeting_generator.generator import generate_document
from tests.source_fixtures import build_primary_source
from tests.template_fixture import build_template


def main() -> None:
    with tempfile.TemporaryDirectory() as directory:
        folder = Path(directory)
        source = build_primary_source(folder / "source.docx")
        template = build_template(folder / "template.docx")
        output = folder / "output.docx"
        durations: list[float] = []
        # La primera generación sirve de calentamiento; los fixtures quedan fuera.
        for iteration in range(6):
            start = perf_counter()
            result = generate_document(source, template, output)
            elapsed = perf_counter() - start
            if not result.succeeded or result.written_weeks != 9:
                raise RuntimeError(f"La medición no generó nueve semanas: {result}")
            if iteration:
                durations.append(elapsed)

        digest = hashlib.sha256()
        with ZipFile(output) as document:
            for name in sorted(document.namelist()):
                digest.update(name.encode())
                digest.update(document.read(name))

        print(
            json.dumps(
                {
                    "python": platform.python_version(),
                    "python_docx": version("python-docx"),
                    "weeks": 9,
                    "runs_seconds": durations,
                    "median_seconds": statistics.median(durations),
                    "docx_parts_sha256": digest.hexdigest(),
                },
                indent=2,
            )
        )


if __name__ == "__main__":
    main()
