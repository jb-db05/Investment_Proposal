#!/usr/bin/env python3
"""
Rebuild the "proposed portfolio" line-items pages as real PowerPoint tables.

The reference deck (work16.pptx) built these two slides out of tall
textboxes with one paragraph per row, manually positioned to line up like a
table. It breaks the moment a category has more lines than fit the fixed
paragraph count it was built for, and it can't repaginate. This script
replaces that with actual `a:tbl` tables (python-pptx `add_table`), grouped
by category, with real subtotal/total rows and slide-count that adapts to
however many holdings `parsed.json` actually has.

Hard rule (matches the task spec): a category (the `group` field in
parsed.json's `line_items` — e.g. "Govies 1-10 (local)", "US" equities,
"Cash") never splits across two slides. If it doesn't fit on the current
slide, the whole group moves to the next one.

Usage (also called from build_proposal.py):
    python build_line_items_tables.py deck.pptx parsed.json -o deck_out.pptx
"""
from __future__ import annotations

import argparse
import copy
import json
from dataclasses import dataclass

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# --- visual constants, matched to the template's own title-block geometry ---
TABLE_LEFT = Emu(665163)
TABLE_WIDTH = Emu(9398001)
TABLE_TOP = Emu(2260000)
TABLE_BOTTOM_MARGIN = Emu(450000)  # keep clear of the footer/page-number

COL_WIDTHS = [Emu(1900000), Emu(1500000), Emu(4573001), Emu(1425000)]  # category, ISIN, instrument, weight
ROW_HEIGHT = Emu(228600)  # ~0.25in, comfortable for 10-10.5pt text

HEADER_FILL = RGBColor(0x1E, 0x27, 0x61)
HEADER_FONT = RGBColor(0xFF, 0xFF, 0xFF)
SUBTOTAL_FILL = RGBColor(0xEA, 0xEC, 0xF5)
TOTAL_FILL = RGBColor(0xF4, 0x7B, 0x20)
BODY_FONT = RGBColor(0x22, 0x22, 0x22)
FONT_NAME = "Calibri"


@dataclass
class Group:
    top_category: str
    group: str
    rows: list[dict]

    @property
    def subtotal_pct(self) -> float:
        return sum(r["weight_pct"] for r in self.rows)


def group_line_items(line_items: list[dict]) -> list[Group]:
    groups: list[Group] = []
    for li in line_items:
        if groups and groups[-1].top_category == li["top_category"] and groups[-1].group == li["group"]:
            groups[-1].rows.append(li)
        else:
            groups.append(Group(li["top_category"], li["group"], [li]))
    return groups


def paginate(groups: list[Group], rows_per_slide: int) -> list[list[Group]]:
    """Greedy bin-packing where a Group is the atomic, never-split unit.
    Emits a fresh category-header row whenever top_category changes,
    including right after a page break, so a continued category is still
    clearly labeled."""
    pages: list[list[Group]] = []
    current: list[Group] = []
    current_rows = 0
    last_top_category = None

    def group_cost(g: Group, top_cat_changed: bool) -> int:
        return (1 if top_cat_changed else 0) + len(g.rows) + 1  # header? + data rows + subtotal row

    for g in groups:
        top_cat_changed = g.top_category != last_top_category
        cost = group_cost(g, top_cat_changed)
        if current and current_rows + cost > rows_per_slide:
            pages.append(current)
            current = []
            current_rows = 0
            last_top_category = None
            top_cat_changed = True
            cost = group_cost(g, top_cat_changed)
        current.append(g)
        current_rows += cost
        last_top_category = g.top_category
    if current:
        pages.append(current)
    return pages


def rows_per_slide_estimate() -> int:
    from pptx.util import Emu as _E
    usable = 7559675 - TABLE_TOP - TABLE_BOTTOM_MARGIN  # slide height is fixed 4:3->16:9 EMU from prs; recomputed properly in build()
    return max(8, usable // ROW_HEIGHT)


def _set_cell(cell, text, *, bold=False, fill=None, font_color=BODY_FONT, size=10, align=PP_ALIGN.LEFT):
    cell.margin_left = Emu(45720)
    cell.margin_right = Emu(45720)
    cell.margin_top = Emu(9144)
    cell.margin_bottom = Emu(9144)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run() if not p.runs else p.runs[0]
    run.text = str(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = FONT_NAME
    run.font.color.rgb = font_color


def _merge_col1(table, r0, r1):
    if r1 > r0:
        table.cell(r0, 0).merge(table.cell(r1, 0))


def render_page(slide, page_groups: list[Group], n_rows: int, is_last_page: bool, grand_total_pct: float | None):
    n_cols = 4
    graphic_frame = slide.shapes.add_table(n_rows, n_cols, TABLE_LEFT, TABLE_TOP,
                                            TABLE_WIDTH, ROW_HEIGHT * n_rows)
    table = graphic_frame.table
    table.first_row = False
    table.horz_banding = False
    for i, w in enumerate(COL_WIDTHS):
        table.columns[i].width = w
    for r in range(n_rows):
        table.rows[r].height = ROW_HEIGHT

    # column headers
    headers = ["Asset Allocation", "ISIN", "Instrument", "Weight"]
    for c, h in enumerate(headers):
        _set_cell(table.cell(0, c), h, bold=True, fill=HEADER_FILL, font_color=HEADER_FONT, size=10)
    r = 1
    last_top_category = None
    for g in page_groups:
        if g.top_category != last_top_category:
            for c in range(n_cols):
                _set_cell(table.cell(r, c), g.top_category if c == 0 else "", bold=True,
                          fill=HEADER_FILL, font_color=HEADER_FONT, size=10.5)
            r += 1
            last_top_category = g.top_category
        group_first_row = r
        for li in g.rows:
            _set_cell(table.cell(r, 1), li["isin"] or "", size=9.5)
            _set_cell(table.cell(r, 2), li["instrument"] or "", size=9.5)
            _set_cell(table.cell(r, 3), f"{li['weight_pct']:.1f}%", size=9.5, align=PP_ALIGN.RIGHT)
            r += 1
        _set_cell(table.cell(group_first_row, 0), g.group, bold=True, size=9.5)
        _merge_col1(table, group_first_row, r - 1)
        # subtotal row for this group
        _set_cell(table.cell(r, 0), "", fill=SUBTOTAL_FILL)
        _set_cell(table.cell(r, 1), "", fill=SUBTOTAL_FILL)
        _set_cell(table.cell(r, 2), f"Subtotal — {g.group}", bold=True, fill=SUBTOTAL_FILL, size=9)
        _set_cell(table.cell(r, 3), f"{g.subtotal_pct:.1f}%", bold=True, fill=SUBTOTAL_FILL, size=9, align=PP_ALIGN.RIGHT)
        r += 1

    if is_last_page and grand_total_pct is not None:
        for c in range(n_cols):
            text = "Total" if c == 0 else (f"{grand_total_pct:.1f}%" if c == 3 else "")
            _set_cell(table.cell(r, c), text, bold=True, fill=TOTAL_FILL, font_color=RGBColor(0xFF, 0xFF, 0xFF), size=10,
                      align=(PP_ALIGN.RIGHT if c == 3 else PP_ALIGN.LEFT))
        r += 1
    assert r == n_rows, f"row count mismatch: built {r}, allocated {n_rows}"


def find_slide(prs, predicate):
    for i, slide in enumerate(prs.slides):
        if predicate(slide):
            return i, slide
    return None, None


def is_line_items_slide(slide) -> bool:
    names = {sh.name for sh in slide.shapes}
    return names == {"Title 1", "Text Placeholder 2"}


def duplicate_slide_after(prs, index):
    """Insert a new slide using the same layout as prs.slides[index],
    positioned immediately after it. Returns the new slide."""
    src = prs.slides[index]
    new_slide = prs.slides.add_slide(src.slide_layout)  # appended at the end
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    new_el = slides[-1]
    xml_slides.remove(new_el)
    xml_slides.insert(index + 1, new_el)
    return new_slide


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[index].get(qn("r:id"))
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[index])


def set_placeholder_text(slide, name, text):
    for sh in slide.shapes:
        if sh.name == name and sh.has_text_frame:
            tf = sh.text_frame
            p = tf.paragraphs[0]
            for run in list(p.runs)[1:]:
                run._r.getparent().remove(run._r)
            if p.runs:
                p.runs[0].text = text
            else:
                p.add_run().text = text
            for extra_p in list(tf.paragraphs[1:]):
                extra_p._p.getparent().remove(extra_p._p)
            return


def build(prs: Presentation, parsed: dict) -> Presentation:
    slide_h = prs.slide_height
    usable = slide_h - TABLE_TOP - TABLE_BOTTOM_MARGIN
    rows_per_slide = max(8, int(usable // ROW_HEIGHT))

    line_items = parsed["line_items"]
    groups = group_line_items(line_items)
    pages = paginate(groups, rows_per_slide)
    n_pages = len(pages)
    grand_total = sum(li["weight_pct"] for li in line_items)

    anchor_idx, anchor_slide = find_slide(prs, is_line_items_slide)
    if anchor_idx is None:
        raise RuntimeError("No blank line-items slide found in the template "
                            "(expected a slide with only 'Title 1' and 'Text Placeholder 2' shapes)")
    # collect every consecutive blank line-items slide starting at anchor (the
    # template ships two: the old 'cash and fixed income' / 'equity...' slots)
    slot_indices = [anchor_idx]
    probe = anchor_idx + 1
    while probe < len(prs.slides) and is_line_items_slide(prs.slides[probe]):
        slot_indices.append(probe)
        probe += 1

    while len(slot_indices) < n_pages:
        new_slide = duplicate_slide_after(prs, slot_indices[-1])
        new_idx = slot_indices[-1] + 1
        slot_indices.append(new_idx)
    while len(slot_indices) > n_pages:
        delete_slide(prs, slot_indices.pop())

    ccy = parsed["base_currency"]
    total_m = parsed["total_value_eur"] / 1_000_000
    date = parsed["valuation_date"]

    for page_num, (idx, page_groups) in enumerate(zip(slot_indices, pages), start=1):
        slide = prs.slides[idx]
        title = "Proposed portfolio: full holdings list" if n_pages > 1 else "Proposed portfolio"
        subtitle = (f"All positions of the proposed {ccy} {total_m:.1f}m allocation, "
                    f"part {page_num} of {n_pages}. Weights as at {date}.")
        set_placeholder_text(slide, "Title 1", title)
        set_placeholder_text(slide, "Text Placeholder 2", subtitle)

        n_rows = 1  # column header
        last_tc = None
        for g in page_groups:
            if g.top_category != last_tc:
                n_rows += 1
                last_tc = g.top_category
            n_rows += len(g.rows) + 1
        is_last = page_num == n_pages
        if is_last:
            n_rows += 1
        render_page(slide, page_groups, n_rows, is_last, grand_total if is_last else None)

    return prs


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx_path")
    ap.add_argument("parsed_json")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()

    prs = Presentation(args.pptx_path)
    parsed = json.loads(open(args.parsed_json).read())
    build(prs, parsed)
    prs.save(args.output)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
