"""A short project deck: section header, two-column plan, results table with
speaker notes, a process drawn with shapes, a picture slide and a chart.

    python examples/slides.py plan.pptx
"""

import sys

from oox import PresentationDocument
from oox.units import Inches


def main(path: str, picture: str | None = None) -> None:
    deck = PresentationDocument.new()

    cover = deck.add_slide("Title Slide")
    deck.set_placeholder_text(cover, "title", "Project Falcon")
    deck.set_placeholder_text(cover, "subtitle", "Status, November 2026")

    section = deck.add_slide("Section Header")
    deck.set_placeholder_text(section, "title", "Where we are")
    deck.set_placeholder_text(section, "body", "Plan, results and next steps")

    plan = deck.add_slide("Two Content")
    deck.set_placeholder_text(plan, "title", "Plan")
    deck.set_placeholder_text(plan, 1, ["Research", (1, "12 interviews"), "Prototype"])
    deck.set_placeholder_text(plan, 2, ["Build", "Pilot in North", "Launch"])
    deck.set_notes(plan, "The pilot moved two weeks because of the holidays.")

    results = deck.add_slide("Title Only")
    deck.set_placeholder_text(results, "title", "Results")
    deck.add_table(results, [["Region", "Q1", "Q2", "Q3"], ["North", 120, 135, 150],
                             ["South", 95, 110, 99], ["West", 80, 84, 91],
                             ["Total", None, None, 340]],
                   Inches(1), Inches(1.8), Inches(8), merges=["A5:C5"])
    deck.set_notes(results, "Totals are rounded to the nearest thousand.")

    process = deck.add_slide("Title Only")
    deck.set_placeholder_text(process, "title", "Next steps")
    for i, (label, fill) in enumerate([("Pilot", "4472C4"), ("Review", "5B9BD5"),
                                       ("Launch", "70AD47")]):
        deck.add_shape(process, "chevron" if i else "homePlate", Inches(1 + i * 3.6),
                       Inches(3), Inches(3.4), Inches(1.2), text=label, fill=fill, size=20)

    if picture:
        photo = deck.add_slide("Picture with Caption")
        deck.set_placeholder_text(photo, "title", "The team")
        deck.set_placeholder_text(photo, 2, "Offsite, May 2026")
        deck.set_placeholder_picture(photo, "pic", picture)

    trend = deck.add_slide("Title Only")
    deck.set_placeholder_text(trend, "title", "Trend")
    deck.add_chart(trend, "line", ["Q1", "Q2", "Q3"],
                   {"North": [120, 135, 150], "South": [95, 110, 99], "West": [80, 84, 91]},
                   Inches(1), Inches(1.8), Inches(10), Inches(5), y_title="EUR k")

    assert not deck.validate(target="Office2007")
    deck.save(path)
    print(len(deck.slides), "slides;", [deck.slide_layout(s) for s in deck.slides])


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "plan.pptx",
         sys.argv[2] if len(sys.argv) > 2 else None)
