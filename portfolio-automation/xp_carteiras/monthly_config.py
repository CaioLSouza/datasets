"""Itens que normalmente precisam de revisão a cada fechamento mensal.

Este é o único arquivo que um usuário operacional deve editar mês a mês.
"""

from __future__ import annotations

from pathlib import Path

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

# A Prestação de Contas é habilitada inicialmente para a Top Ações, que é a
# carteira do template fornecido. Para incluir outra carteira, adicione aqui o
# template correspondente depois de validar sua estrutura visual.
ACCOUNTABILITY_TEMPLATE_FILES = {
    "Carteira - TOP Ações XP": "Prestação de Contas - Top Ações - Julho 2026.pptx",
}

ACCOUNTABILITY_OUTPUT_LABELS = {
    "Carteira - TOP Ações XP": "Top Ações",
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


def accountability_ppt_config(settings: Settings) -> dict[str, dict[str, str]]:
    """Monta a configuração da Prestação de Contas.

    Dá preferência ao template mensal da pasta corporativa e usa a cópia
    empacotada no projeto quando ele não estiver disponível.
    """
    packaged_dir = Path(__file__).resolve().parent.parent / "templates"
    config = {}
    for portfolio, filename in ACCOUNTABILITY_TEMPLATE_FILES.items():
        corporate = settings.templates_dir / filename
        packaged = packaged_dir / filename
        template = corporate if corporate.exists() else packaged
        config[portfolio] = {
            "template": str(template),
            "output_label": ACCOUNTABILITY_OUTPUT_LABELS[portfolio],
        }
    return config
