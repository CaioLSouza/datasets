"""Itens que normalmente precisam de revisão a cada fechamento mensal.

Este é o único arquivo que um usuário operacional deve editar mês a mês.
"""

from __future__ import annotations

from .settings import Settings


# Atualize os nomes abaixo quando os templates do mês mudarem.
COMMERCIAL_TEMPLATE_FILES = {
    "Carteira - TOP Ações XP": "Lâmina Comercial - Top Ações - Junho 2026.pptx",
    "Carteira - TOP DIVIDENDOS XP": "Lâmina Comercial - Top Dividendos - Junho 2026.pptx",
    "Carteira - TOP SMALL CAPS XP": "Lâmina Comercial - Top Small Caps - Junho 2026.pptx",
}

COMMERCIAL_OUTPUT_FILES = {
    "Carteira - TOP Ações XP": "Lâmina Comercial - Top Ações.pptx",
    "Carteira - TOP DIVIDENDOS XP": "Lâmina Comercial - Top Dividendos.pptx",
    "Carteira - TOP SMALL CAPS XP": "Lâmina Comercial - Top Small Caps.pptx",
}

def commercial_ppt_config(settings: Settings) -> dict[str, dict[str, str]]:
    """Monta os caminhos de entrada e saída das lâminas comerciais."""
    return {
        portfolio: {
            "template": str(settings.templates_dir / template),
            "saida": str(settings.commercial_deck_dir / COMMERCIAL_OUTPUT_FILES[portfolio]),
        }
        for portfolio, template in COMMERCIAL_TEMPLATE_FILES.items()
    }
