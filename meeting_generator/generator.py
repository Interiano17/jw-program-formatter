"""Orquestación fail-closed del parser y el escritor."""

import logging
from pathlib import Path

from .parser import ProgramParser
from .results import GenerationResult
from .template_writer import fill_template
from .validation import GenerationValidationError

logger = logging.getLogger(__name__)


def generate_document(
    source_path: Path,
    template_path: Path,
    output_path: Path,
) -> GenerationResult:
    """Genera el documento o devuelve un fallo estructurado sin resultado parcial."""
    parser = ProgramParser(source_path)
    try:
        weeks = parser.parse()
    except GenerationValidationError as exc:
        logger.error("Documento fuente rechazado: %s", exc)
        return GenerationResult.failure(parser.detected_week_count, list(exc.issues))
    except Exception as exc:
        logger.error("No se pudo analizar el documento fuente", exc_info=True)
        return GenerationResult.failure(
            parser.detected_week_count,
            [f"No se pudo analizar el documento fuente: {exc}"],
        )

    try:
        resolved_output = output_path.resolve()
        resolved_source = source_path.resolve()
        resolved_template = template_path.resolve()
    except (OSError, RuntimeError) as exc:
        return GenerationResult.failure(
            len(weeks),
            [f"No se pudo validar la ruta de salida: {exc}"],
        )

    if resolved_output == resolved_source:
        return GenerationResult.failure(
            len(weeks),
            ["La ruta de salida coincide con el documento fuente"],
        )
    if resolved_output == resolved_template:
        return GenerationResult.failure(
            len(weeks),
            ["La ruta de salida coincide con la plantilla"],
        )

    try:
        result = fill_template(template_path, weeks, output_path)
    except Exception as exc:
        logger.error("El escritor no devolvió un resultado", exc_info=True)
        return GenerationResult.failure(
            len(weeks),
            [f"El escritor falló inesperadamente: {exc}"],
        )
    logger.info(
        "Resultado: escritas=%d omitidas=%d errores=%d salida=%s",
        result.written_weeks,
        result.omitted_weeks,
        len(result.errors),
        result.output_path,
    )
    return result
