"""Build a small report with the raw factory API, odfpy style.

    python examples/hello.py out.docx
"""

import sys

from oox import WordprocessingDocument
from oox.ns import w


def para(*runs, align=None, space_after=None, bold=False, size=None):
    ppr = w.PPr()
    if align:
        ppr.add(w.Jc(val=align))
    if space_after is not None:
        ppr.add(w.Spacing(after=space_after))
    return w.P(ppr, *runs) if len(ppr) else w.P(*runs)


def run(text, bold=False, italic=False, size=None, color=None):
    rpr = w.RPr()
    if color:
        rpr.add(w.Color(val=color))
    if size:  # half-points
        rpr.add(w.Sz(val=size))
    if italic:
        rpr.add(w.I())
    if bold:
        rpr.add(w.B())
    return w.R(rpr, w.T(text)) if len(rpr) else w.R(w.T(text))


def cell(text, width, shade=None, bold=False):
    tcpr = w.TcPr(w.TcW(w=width, type="dxa"))
    if shade:
        tcpr.add(w.Shd(val="clear", color="auto", fill=shade))
    return w.Tc(tcpr, w.P(run(text, bold=bold)))


def table(rows, widths):
    border = dict(val="single", sz=4, space=0, color="999999")
    tbl = w.Tbl(
        w.TblPr(
            w.TblW(w=sum(widths), type="dxa"),
            w.TblBorders(w.Top(**border), w.Left(**border), w.Bottom(**border),
                         w.Right(**border), w.InsideH(**border), w.InsideV(**border)),
        ),
        w.TblGrid(*(w.GridCol(w=x) for x in widths)),
    )
    for i, row in enumerate(rows):
        tbl.add(w.Tr(*(cell(t, widths[j], "D9E2F3" if i == 0 else None, bold=i == 0)
                      for j, t in enumerate(row))))
    return tbl


def main(path: str) -> None:
    doc = WordprocessingDocument.new()
    doc.body.add(
        para(run("oox prototype", bold=True, size=48, color="1F3864"), align="center",
             space_after=240),
        para(run("Generated from the Open XML SDK schema metadata. "),
             run("Every element below was validated", italic=True),
             run(" and inserted in schema order.")),
        table([["Layer", "Source", "Status"],
               ["Grammar tables", "Generated", "Prototype"],
               ["Element factories", "Generated", "Prototype"],
               ["OPC packaging", "Hand-written", "Prototype"]],
              widths=[3000, 3000, 3000]),
        para(run("  Leading spaces are preserved automatically.")),
    )
    # New documents come with an A4 section; replace it with our own page
    # setup.  Order inside sectPr is enforced for us (pgMar before pgSz here).
    doc.body.add(w.SectPr(w.PgMar(top=1134, right=1134, bottom=1134, left=1134, header=708,
                                  footer=708, gutter=0),
                          w.PgSz(w=11906, h=16838)), replace=True)
    problems = doc.document.validate()
    assert not problems, problems
    doc.save(path)
    print(f"wrote {path}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "hello.docx")
