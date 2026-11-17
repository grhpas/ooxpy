# PowerPoint

`PresentationDocument` creates and edits .pptx files. Slides are elements; most methods
take the slide as their first argument. Every block below runs as written, in order.
The [reference](../reference/powerpoint.md) lists every option.

## A new deck and its layouts

```python
from oox import PresentationDocument
from oox.units import Inches

deck = PresentationDocument.new()                   # 16:9 (13.33 x 7.5 in)
print(deck.layout_names[:4])
# ['Title Slide', 'Title and Content', 'Section Header', 'Two Content']

old_style = PresentationDocument.new(width=Inches(10), height=Inches(7.5))   # 4:3
```

`new()` builds the slide master, the Office theme and Office's eleven standard layouts
with their placeholder geometry, scaled to the slide size. `layouts="basic"` keeps only
Title Slide, Title and Content and Blank.

## Slides and placeholders

```python
cover = deck.add_slide("Title Slide")
deck.set_placeholder_text(cover, "title", "Project Falcon")
deck.set_placeholder_text(cover, "subtitle", "Status, November 2026")

plan = deck.add_slide("Two Content")
deck.set_placeholder_text(plan, "title", "Plan")
deck.set_placeholder_text(plan, 1, ["Research", (1, "12 interviews"), "Prototype"])
deck.set_placeholder_text(plan, 2, "Build\nPilot in North\nLaunch")

print([(ph.type, ph.idx) for ph in deck.placeholders(plan)])
```

`add_slide(layout)` copies the layout's placeholders onto the slide, empty and
inheriting position and style, like PowerPoint's New Slide. Placeholders are chosen by
type (`"title"`, `"body"`, `"subtitle"`, `"pic"` …), by index, or by shape name. Text is
a string (`\n` starts a paragraph) or a list of paragraphs, where `(level, text)`
indents a bullet.

## Text boxes, shapes and pictures

```python
steps = deck.add_slide("Title Only")
deck.set_placeholder_text(steps, "title", "Next steps")
for i, (label, fill) in enumerate([("Pilot", "4472C4"), ("Review", "5B9BD5"),
                                   ("Launch", "70AD47")]):
    deck.add_shape(steps, "chevron" if i else "homePlate", x=Inches(1 + i * 3.6), y=Inches(3),
                   width=Inches(3.4), height=Inches(1.2), text=label, fill=fill, size=20)
deck.add_text_box(steps, "Dates are provisional.", x=Inches(1), y=Inches(5), width=Inches(10),
                  height=Inches(0.8), size=14, color="595959")
deck.add_picture(steps, "logo.png", x=Inches(11.5), y=Inches(0.3), width=Inches(1.2))

print([(s.kind, s.text) for s in deck.shapes(steps)][:3])
```

Positions and sizes are keywords: `x`, `y`, `width` and `height`, in points or a length
from `oox.units`. Leave them out and the shape, table, chart or picture goes in the
slide's content area (`deck.content_area(slide)`, where "Title and Content" puts its
text): `deck.add_table(slide, rows)` is enough for a table under the title.

Shapes use PowerPoint's preset geometries (`rect`, `roundRect`, `ellipse`, `rightArrow`,
`chevron` …) with fill, line and centred text. `shapes(slide)` describes everything on
a slide as `Shape` records: placeholders, text boxes, shapes, pictures, tables, charts
and groups, with their position and size in points.

A picture placeholder (on the "Picture with Caption" layout, say) is filled the way
PowerPoint fills it, cropped to the placeholder:

```python
photo = deck.add_slide("Picture with Caption")
deck.set_placeholder_text(photo, "title", "The team")
deck.set_placeholder_picture(photo, "pic", "chart.png")
```

## Tables

```python
results = deck.add_slide("Title Only")
deck.set_placeholder_text(results, "title", "Results")
deck.add_table(results, [["Region", "Q1", "Q2", "Q3"],
                         ["North", 120, 135, 150],
                         ["South", 95, 110, 99],
                         ["Total", None, None, 384]],
               x=Inches(1), y=Inches(1.8), width=Inches(8), merges=["A4:C4"],
               style="Medium Style 2 - Accent 1")
```

Tables use PowerPoint's built-in styles (by name or GUID) with the header, banded and
first-column switches; merges are A1-style ranges.

**Editing** an existing table, in this deck or one made by PowerPoint:

```python
# the first table on the slide (or by name)
table = deck.table(results)
table.set_text(1, 3, 155)                        # keeps the cell's font and alignment
table.add_row(["West", 80, 84, 91], index=3)      # formatting copied from a neighbour
table.add_column(["Q4", 160, 101, 95, None], width=Inches(1.5))
table.delete_row(2)
print(table.values()[0])                          # ['Region', 'Q1', 'Q2', 'Q3', 'Q4']
```

Rows and columns can be added, deleted, merged and unmerged; merges shrink or move
the way PowerPoint does it, and the table frame follows the grid.

## Speaker notes

```python
deck.set_notes(results, "Totals are rounded to the nearest thousand.")
print(deck.notes(results))
```

The notes master is created on first use (or an existing one is reused).

## Links

```python
deck.set_link(steps, deck.shapes(steps)[-2].name, "https://example.org/plan",
              tooltip="The full plan")
deck.set_link(cover, deck.shapes(cover)[0].name, plan)    # jump to a slide
deck.set_link(results, deck.shapes(results)[0].name, "next")
print([link.kind for link in deck.links(steps)])         # ['url']
```

A link goes on a whole shape, or on `text=` inside it. Targets are URLs (`http`,
`https`, `mailto`, `file`), slides, or `"next"`, `"previous"`, `"first"`, `"last"`,
`"end"`; `None` removes a link.

## Sections

```python
deck.set_sections([("Introduction", cover), ("Plan", plan), ("Results", results)])
print(deck.sections())
```

Sections are a PowerPoint 2010 feature: PowerPoint 2007 and LibreOffice ignore them
and show the slides in order.

## Charts

```python
trend = deck.add_slide("Title Only")
deck.set_placeholder_text(trend, "title", "Trend")
deck.add_chart(trend, "line", ["Q1", "Q2", "Q3"],
               {"North": [120, 135, 150], "South": [95, 110, 99]}, x=Inches(1), y=Inches(1.8),
               width=Inches(10), height=Inches(5), y_title="EUR k")
```

See [Charts](charts.md).

## Saving

```python
assert not deck.validate(target="Office2007")
deck.save("falcon.pptx")
print(len(deck.slides), [deck.slide_layout(s) for s in deck.slides][:3])
```
