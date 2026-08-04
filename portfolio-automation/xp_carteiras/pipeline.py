"""Orquestra o pipeline completo ou apenas uma de suas etapas."""

from __future__ import annotations

from .output_pipeline import generate_output_files
from .pipeline_data import prepare_pipeline_context
from .powerpoint_pipeline import generate_powerpoints
from .settings import Settings, load_settings


def _prepare(settings: Settings | None):
    resolved = settings or load_settings()
    return prepare_pipeline_context(resolved)


def main_output(settings: Settings | None = None) -> None:
    """Atualiza somente os arquivos Excel da pasta de output."""
    generate_output_files(_prepare(settings))


def main_powerpoints(settings: Settings | None = None) -> None:
    """Gera somente as Lâminas Comerciais e Prestações de Contas."""
    generate_powerpoints(_prepare(settings))


def main(settings: Settings | None = None) -> None:
    """Executa dados de output e PowerPoint reutilizando a mesma preparação."""
    context = _prepare(settings)
    generate_output_files(context)
    generate_powerpoints(context)
