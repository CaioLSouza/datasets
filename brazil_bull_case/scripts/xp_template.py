"""Look of the XP model template (template modelo XP.xlsx, the house WEG model), shared by the build scripts.

Roboto 10; navy 1F2F44 sections; yellow F9C113 blocks; grey band sub-blocks; inputs in blue 0070C0; assumptions to
change in pink FFE7E7 with a dotted border; formulas (links included) in black; logo in the print header and the
classification footer; zoom 85%; a Cover with navigation buttons and a yellow 'Support >' divider.
"""
from __future__ import annotations

from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as cn

HERE = Path(__file__).parent
MEDIA = HERE / "template_media"
LOGO_BLACK, LOGO_WHITE = MEDIA / "xp_research_black_header.png", MEDIA / "xp_research_white.png"

NAVY, LBLUE, YEL, GREY, GREEN = "#1F2F44", "#8EB3DF", "#F9C113", "#74797C", "#008000"
INK, TITLE_C = "#18191A", "#1D1E1F"
INBLUE, PINK, GBAND, KEY = "#0070C0", "#FFE7E7", "#F2F2F2", "#E9EAEC"
PTS, PTS1, MULT, PCT, PCT2 = "#,##0", "#,##0.0", '0.0"x"', "0.0%", "0.00%"
UPS = "+0.0%;-0.0%;0.0%"
PP = '+0.0" pp";-0.0" pp";0.0" pp"'
PPC = '0.00" pp"'
BP = '+0" bp";-0" bp";0" bp"'
PX, DATE, DATED = "#,##0.00", "[$-409]mmm-yy", "[$-409]d-mmm-yy"
FOOTER = '&R&"Calibri"&10&K008000[ CLASSIFICAÇÃO: PÚBLICA ]'
HEAT = {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0, "mid_color": "#FFEB84",
        "max_color": "#63BE7B"}


class Book:
    def __init__(self, path, order, tab_colors=None, prefix="Ibovespa - "):
        self.wb = xlsxwriter.Workbook(str(path))
        self.wb.set_size(1700, 1050)
        self._fc = {}
        self.prefix = prefix
        self.W = {}
        for s in order:
            ws = self.wb.add_worksheet(s)
            if s in (tab_colors or {}):
                ws.set_tab_color(tab_colors[s])
            ws.hide_gridlines(2)
            ws.set_default_row(14)
            ws.set_zoom(85)
            if s not in ("Cover", "Support >"):
                ws.set_header("&L&G", {"image_left": str(LOGO_BLACK)})
                ws.set_footer(FOOTER)
                ws.set_margins(0.4, 0.4, 1.25, 0.5)
                ws.set_landscape()
                ws.set_paper(9)
                ws.fit_to_pages(1, 0)
            self.W[s] = ws
        self.IN = dict(font_color=INBLUE)
        self.LV = dict(font_color=INBLUE, bold=True, bg_color=PINK, border=4)
        self.HDR = self.F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", text_wrap=True)
        self.HDRL = self.F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="left", text_wrap=True)
        self.SUB = self.F(bold=True, bg_color=GBAND, bottom=7)
        self.TXT = self.F()
        self.TXTW = self.F(text_wrap=True, valign="top")
        self.NOTE = self.F(italic=True, font_color=GREY)
        self.BOLD = self.F(bold=True)
        self.NA = self.F(font_color=GREY, align="right")
        self.GROUPH = self.F(bold=True, bg_color=GBAND, align="center", bottom=7)

    def F(self, **k):
        key = tuple(sorted(k.items()))
        if key not in self._fc:
            p = {"font_name": "Roboto", "font_size": 10, "valign": "vcenter", "font_color": INK}
            p.update(k)
            self._fc[key] = self.wb.add_format(p)
        return self._fc[key]

    def name(self, nm, sheet, ref):
        q = f"'{sheet}'" if any(ch in sheet for ch in " &->") else sheet
        self.wb.define_name(nm, f"={q}!{ref}")

    def title(self, ws, text, sub=None, legend=True):
        ws.set_row(0, 30)
        ws.write("B1", self.prefix + text, self.F(bold=True, font_size=18, font_color=TITLE_C))
        if sub:
            ws.write("B2", sub, self.NOTE)
        if legend:
            ws.write("B3", "Legend:", self.NOTE)
            ws.write("C3", "input", self.F(**self.IN, italic=True))
            ws.write("D3", "assumption", self.F(**self.LV, italic=True))
            ws.write("E3", "formula", self.F(italic=True))

    def section(self, ws, row, text, c0="B", c1="H", fill=NAVY, color="#FFFFFF"):
        ws.merge_range(f"{c0}{row}:{c1}{row}", text, self.F(bold=True, font_color=color, bg_color=fill))

    def header_row(self, ws, row, cells, first_left=True):
        """cells: list of (col letter, text)."""
        for i, (c, h) in enumerate(cells):
            ws.write(f"{c}{row}", h, self.HDRL if (i == 0 and first_left) else self.HDR)

    def button(self, ws, rng, text, target, fill, color):
        first = rng.split(":")[0]
        f = self.F(bold=True, font_size=11, font_color=color, bg_color=fill, align="center", text_wrap=True, border=5,
                   border_color="#FFFFFF")
        ws.merge_range(rng, "", f)
        q = f"'{target}'"
        ws.write_url(first, f"internal:{q}!A1", f, text, tip=f"Go to {target}")

    def cover(self, title_bold, title_light, subtitle, main, aux_rows, readme=None):
        """Cover in the template's layout: navy band, white logo, navy main buttons, yellow auxiliary buttons."""
        ws = self.W["Cover"]
        ws.set_column("A:A", 2.2)
        ws.set_column("B:B", 1.4)
        ws.set_column("C:N", 11.5)
        ws.set_column("O:O", 1.4)
        nav = self.F(bg_color=NAVY)
        for r in range(2, 6):
            for c in "BCDEFGHIJKLMNO":
                ws.write_blank(f"{c}{r}", None, nav)
        ws.set_row(1, 8)
        ws.set_row(2, 34)
        ws.set_row(3, 18)
        ws.set_row(4, 8)
        ws.write_rich_string("C3", self.F(bold=True, font_size=18, font_color="#FFFFFF"), title_bold,
                             self.F(font_name="Roboto Light", font_size=18, font_color="#FFFFFF"), title_light, nav)
        ws.write("C4", subtitle, self.F(font_name="Roboto Light", font_color="#FFFFFF", bg_color=NAVY))
        ws.insert_image("L2", str(LOGO_WHITE), {"x_scale": 0.62, "y_scale": 0.62, "x_offset": 40, "y_offset": 6,
                                                "object_position": 3})
        lab = self.F(bold=True, italic=True, font_size=9, font_color=NAVY)
        ws.write("C7", "Main tabs >>", lab)
        ws.write("C11", "Auxiliary tabs >>", lab)
        for i, (t, tgt) in enumerate(main):
            self.button(ws, f"{cn(2 + 2 * i)}8:{cn(3 + 2 * i)}9", t, tgt, NAVY, "#FFFFFF")
        for k, row in enumerate(aux_rows):
            r = 12 + 3 * k
            for i, (t, tgt) in enumerate(row):
                self.button(ws, f"{cn(2 + 2 * i)}{r}:{cn(3 + 2 * i)}{r + 1}", t, tgt, YEL, INK)
            ws.set_row(r - 1, 19)
            ws.set_row(r, 19)
        ws.set_row(7, 19)
        ws.set_row(8, 19)
        if readme:
            ws.write_url("K15", f"internal:'{readme}'!A1", self.F(italic=True, underline=1, font_color=INBLUE),
                         "How this model works >>")
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.activate()
        return ws

    def close(self):
        self.wb.close()
