"""
build_presentation.py
----------------------
Takes the coded/pooled team data and slots it into the pool-team text boxes
of a PPTX template, using python-pptx (a standard library — runs anywhere,
no LibreOffice or sandbox-only tooling required).

How it finds the right text boxes: it looks for any text frame whose first
paragraph matches "Pool X Teams ... so far...". It then treats the *first*
paragraph as the header (left untouched), any *trailing* paragraphs with no
text runs as blank/placeholder lines (left untouched), and everything in
between as the team list (replaced).
"""

import copy
import re

from pptx import Presentation

POOL_HEADER_RE = re.compile(r"Pool\s+([A-Z])\s+Teams.*so far", re.IGNORECASE)


def _clone_paragraph_as_template(paragraph):
    """Deep-copy a paragraph's underlying XML so we can reuse its exact
    run formatting (font, size, color, bold) for new team entries."""
    return copy.deepcopy(paragraph._p)


def _set_paragraph_text(p_element, text, nsmap):
    """Given a cloned <a:p> XML element, replace its run text with `text`,
    keeping the first run's formatting and dropping any extra runs (some
    team names span multiple runs in the original e.g. for a misspelled-word
    correction; we don't need that for freshly generated names)."""
    a_ns = nsmap["a"]
    runs = p_element.findall(f"{{{a_ns}}}r")
    if not runs:
        return
    # Keep the first run, set its text, remove any additional runs.
    first_run = runs[0]
    t_el = first_run.find(f"{{{a_ns}}}t")
    if t_el is None:
        t_el = first_run.makeelement(f"{{{a_ns}}}t", {})
        first_run.append(t_el)
    t_el.text = text
    for extra in runs[1:]:
        p_element.remove(extra)


def build_presentation(template_path, output_path, pools_by_category):
    """
    pools_by_category: {
        "Grade 10": {"A": [...], "B": [...], ...},
        "Senior":   {"A": [...], "B": [...], ...},
    }
    Only categories present in pools_by_category get touched; any slide/
    shape that doesn't match a "Pool X Teams" header is left completely
    alone, so unrelated slides (topics, intro, closing) are untouched.
    """
    prs = Presentation(template_path)
    nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    # Flatten all pool letter->teams across categories isn't safe (different
    # categories reuse "Pool A" labels on different slides) so we process
    # slide by slide, matching each slide's pool letters against whichever
    # category dict has matching pool keys -- in practice each category's
    # pools live on their own slide, so we just try each category's pool
    # dict against each slide's shapes and use whichever one has a header
    # matching its lettering. Simpler: caller passes one category's pools
    # at a time by calling this function once per slide-category pairing,
    # OR -- as implemented here -- we match by slide order: the Nth slide
    # containing pool-team shapes gets the Nth category in pools_by_category.

    categories = list(pools_by_category.keys())
    cat_index = 0

    for slide in prs.slides:
        # Determine whether this slide has any pool-team shapes at all.
        pool_shapes = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if not shape.text_frame.paragraphs:
                continue
            header_text = shape.text_frame.paragraphs[0].text
            m = POOL_HEADER_RE.search(header_text)
            if m:
                pool_shapes.append((shape, m.group(1)))

        if not pool_shapes:
            continue

        if cat_index >= len(categories):
            break
        category = categories[cat_index]
        pools = pools_by_category[category]
        cat_index += 1

        for shape, pool_letter in pool_shapes:
            if pool_letter not in pools:
                continue
            tf = shape.text_frame
            paragraphs = tf.paragraphs

            header_p = paragraphs[0]
            # Identify trailing blank paragraphs (no runs) from the end.
            trailing = []
            for p in reversed(paragraphs):
                if len(p.runs) == 0 and p is not header_p:
                    trailing.append(p)
                else:
                    break
            trailing_count = len(trailing)

            team_paragraphs = paragraphs[1: len(paragraphs) - trailing_count] if trailing_count else paragraphs[1:]
            if not team_paragraphs:
                continue
            template_p_xml = _clone_paragraph_as_template(team_paragraphs[0])

            body = tf._txBody
            a_ns = nsmap["a"]

            # Remove all old team paragraphs from the XML body.
            for p in team_paragraphs:
                body.remove(p._p)

            # Build new paragraphs from the template XML and insert them
            # right after the header paragraph, before any trailing blanks.
            insert_after = header_p._p
            for team_name in pools[pool_letter]:
                new_p = copy.deepcopy(template_p_xml)
                _set_paragraph_text(new_p, team_name, nsmap)
                insert_after.addnext(new_p)
                insert_after = new_p

    prs.save(output_path)


if __name__ == "__main__":
    import json
    import sys

    template_path, output_path, pools_json_path = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(pools_json_path) as f:
        pools_by_category = json.load(f)
    build_presentation(template_path, output_path, pools_by_category)
    print(f"Wrote {output_path}")
