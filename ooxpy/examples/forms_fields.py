"""A styled report with captions and cross-references, and a form that is
filled in from data.

    python examples/forms_fields.py outdir
"""

import datetime as dt
import sys
from pathlib import Path

from oox import WordprocessingDocument, open_document


def report(path: Path) -> None:
    doc = WordprocessingDocument.new()
    doc.properties.title = "Quarterly figures"
    doc.extended_properties.company = "Example Ltd"
    doc.custom_properties["Quarter"] = "Q3"

    doc.add_style("Note", italic=True, color="595959", indent_left=720, space_before=6)

    title = doc.add_paragraph("")
    doc.add_field(title, "TITLE")                    # filled from the properties
    doc.add_run(title, " – ")
    doc.add_field(title, "DOCPROPERTY Quarter")
    doc.format_paragraph(title, align="center", space_after=18)

    intro = doc.add_heading("Summary", 1)
    p = doc.add_paragraph("Revenue rose in every region ")
    doc.add_run(p, "except the west", bold=True, color="C00000")
    doc.add_run(p, ". ")
    table_caption = doc.add_caption("Revenue by region", label="Table")
    doc.add_table([["Region", "Q2", "Q3"], ["North", 120, 135], ["West", 90, 84]],
                  header_rows=1)
    q = doc.add_paragraph("The figures are in ")
    doc.add_cross_reference(q, table_caption)            # "Table 1"
    doc.add_run(q, ", ")
    doc.add_cross_reference(q, table_caption, what="position")
    doc.add_run(q, "; the method is described under ")
    doc.add_cross_reference(q, intro)
    doc.add_run(q, ".")
    doc.add_paragraph("Figures are unaudited.", style="Note")
    doc.save(path)


def form(path: Path) -> None:
    doc = WordprocessingDocument.new()
    doc.add_heading("Leave request", 1)
    doc.add_text_control(doc.add_paragraph("Name: "), tag="name", placeholder="Your name")
    doc.add_dropdown(doc.add_paragraph("Type: "), tag="type",
                     items=["Annual leave", "Sick leave", "Unpaid"])
    doc.add_date_control(doc.add_paragraph("From: "), tag="from", fmt="d MMMM yyyy")
    doc.add_checkbox(doc.add_paragraph("Manager informed: "), tag="informed")
    doc.save(path)


def fill(template: Path, out: Path, data: dict) -> None:
    doc = open_document(template)
    doc.fill_controls(data)
    doc.save(out)
    print({c["tag"]: c["value"] for c in doc.content_controls()})


def main(outdir: str) -> None:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    report(out / "report.docx")
    form(out / "leave-form.docx")
    fill(out / "leave-form.docx", out / "leave-ada.docx",
         {"name": "Ada Lovelace", "type": "Annual leave", "from": dt.date(2026, 12, 21),
          "informed": True})


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "out")
