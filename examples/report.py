"""A structured Word report: sections, headers and footers, a table of
contents, links, notes, a comment, and template placeholders.

    python examples/report.py report.docx
"""

import sys

from oox import WordprocessingDocument
from oox.ns import w


def main(path: str) -> None:
    doc = WordprocessingDocument.new()
    doc.properties.title = "{{client}} annual report"

    doc.add_heading("{{client}} annual report", 0)
    doc.add_paragraph("Prepared on {{date}}.")
    first_chapter = doc.add_heading("Summary", 1)
    p = doc.add_paragraph("Revenue grew in every region. Details are in ")
    doc.add_hyperlink(p, "the appendix", anchor="appendix")
    p.add(w.R(w.T("; the method follows ", space="preserve")))
    doc.add_hyperlink(p, "ECMA-376", target="https://ecma-international.org/publications-and-standards/standards/ecma-376/")
    p.add(w.R(w.T(".")))
    doc.add_footnote(p, "All figures are unaudited.")
    doc.add_comment(p, "Confirm with finance before sending.", author="Ada Lovelace")
    doc.add_heading("Regions", 2)
    regions = doc.new_list("bullet")
    for line in ("North: +12%", "South: +8%", "West: +15%"):
        regions.add(line)
    doc.add_heading("Outlook", 1)
    doc.add_paragraph("We expect growth to continue next year.")

    # Page setup: different first page, page numbers in the footer.
    cover = doc.sections[0]
    cover.different_first_page = True
    cover.header_text("{{client}}\t\tConfidential")
    cover.footer_text("Page {PAGE} of {NUMPAGES}", align="center")

    # A landscape section for wide material; it keeps the footer above.
    appendix = doc.add_section(landscape=True)
    doc.add_bookmark("appendix", doc.add_heading("Appendix: figures by quarter", 1))
    doc.add_paragraph("Wide tables go here.")
    doc.add_endnote(doc.add_paragraph("Sources are listed at the end."), "Internal sales database.")
    appendix.header_text("{{client}} - appendix", align="right")

    doc.add_toc(levels=(1, 2), before=first_chapter)

    # Fill in the template placeholders everywhere (body, headers, notes...).
    counts = doc.replace_text({"{{client}}": "Example Ltd", "{{date}}": "14 November 2026"})
    doc.properties.title = "Example Ltd annual report"

    problems = doc.validate(target="Office2007")
    assert not problems, problems
    doc.save(path)
    print(f"wrote {path}; replaced {counts}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "report.docx")
