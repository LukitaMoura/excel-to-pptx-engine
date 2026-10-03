# -*- coding: utf-8 -*-
"""Identidade visual da apresentacao.

Centraliza nome da empresa, cores, fonte e imagens usados tanto na planilha
modelo quanto no PowerPoint. Para aplicar outra marca, basta editar este
arquivo e colocar as imagens em hr_deck/assets/ (todas opcionais: sem elas
o gerador desenha fundos solidos e escreve o nome da empresa em texto).
"""
import os
from pptx.dml.color import RGBColor

COMPANY = "Empresa Exemplo"
COMPANY_SHORT = "EXEMPLO"

# ---------------------------------------------------------------------------
# Imagens opcionais (pasta hr_deck/assets)
# ---------------------------------------------------------------------------
_ASSETS = os.path.join(os.path.dirname(__file__), "assets")
COVER_BG     = os.path.join(_ASSETS, "cover_bg.jpg")       # fundo da capa
LOGO_ICON    = os.path.join(_ASSETS, "logo_icon.png")      # icone do cabecalho
LOGO_WORD    = os.path.join(_ASSETS, "logo_wordmark.png")  # logotipo da capa

# ---------------------------------------------------------------------------
# Paleta: cor de destaque + cor primaria (estrutura/titulos) e seus tints
# ---------------------------------------------------------------------------
ACCENT       = "E4572E"   # destaques, CTA, acentos dos cards
ACCENT_DK    = "B8401D"
ACCENT_80    = "E97958"
ACCENT_60    = "EF9A82"
ACCENT_40    = "F4BCAB"
ACCENT_20    = "FADDD5"
ACCENT_10    = "FCEEEA"

PRIMARY      = "1D3557"   # estrutura, titulos, cabecalhos
PRIMARY_80   = "4A5D79"
PRIMARY_60   = "77869A"
PRIMARY_40   = "A5AEBC"
PRIMARY_20   = "D2D7DD"
PRIMARY_10   = "E8EBEE"

GRAY       = "C1C5C8"
GRAY_DARK  = "7A8084"   # textos secundarios
GRAY_LT    = "F0F2F5"   # surface / backgrounds suaves
WHITE      = "FFFFFF"
BORDER     = "D8DCDE"
TEXT       = "1A2B33"   # texto principal

# Status (semaforo de conformidade: dentro x fora da cota)
GREEN      = "1E8E5A"
GREEN_10   = "E6F4EC"
AMBER      = "B7791F"

FONT = "Calibri"
FONT_FALLBACK = "Segoe UI"

# Sequencia de cores para series de graficos
SERIES_COLORS = [ACCENT, PRIMARY, ACCENT_60, PRIMARY_60, GRAY_DARK, ACCENT_20, PRIMARY_20]


def rgb(hex_str):
    """'E4572E' -> RGBColor."""
    return RGBColor.from_string(hex_str)
