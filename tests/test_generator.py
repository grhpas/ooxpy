"""Generator unit tests (synthetic input) + optional full-regeneration check."""

import filecmp
import os
from pathlib import Path

import pytest

from tools.gen.sdk_frontend import _occurs, _ParticleWalker


def el(name, **occ):
    p = {"Name": name}
    if occ:
        p["Occurs"] = [occ]
    return p


def test_occurs_defaults():
    assert _occurs({}) == (1, 1)  # absent -> exactly once
    assert _occurs({"Occurs": [{}]}) == (0, None)  # {} -> 0..unbounded
    assert _occurs({"Occurs": [{"Max": 1}]}) == (0, 1)  # optional
    assert _occurs({"Occurs": [{"Min": 1}]}) == (1, None)
    # versioned entries are skipped in favour of the baseline
    assert _occurs({"Occurs": [{"Max": 2, "Version": "Office2010"}, {"Max": 1}]}) == (0, 1)


def test_sequence_ranks_and_singletons():
    w = _ParticleWalker()
    w.walk({"Kind": "Sequence", "Items": [el("T/a", Max=1), el("T/b"), el("T/c", Min=0)]})
    a, b, c = (w.children[k] for k in ("T/a", "T/b", "T/c"))
    assert a.rank < b.rank < c.rank
    assert a.max_one and b.max_one and not c.max_one


def test_choice_members_share_rank():
    w = _ParticleWalker()
    w.walk({"Kind": "Sequence", "Items": [
        el("T/first"),
        {"Kind": "Choice", "Items": [el("T/x"), el("T/y")]},
        el("T/last"),
    ]})
    ch = w.children
    assert ch["T/x"].rank == ch["T/y"].rank
    assert ch["T/first"].rank < ch["T/x"].rank < ch["T/last"].rank


def test_repeating_compositor_freezes_order():
    """(a, b)* must keep insertion order, not regroup a b a -> a a b."""
    w = _ParticleWalker()
    w.walk({"Kind": "Sequence", "Items": [
        el("T/head"),
        {"Kind": "Sequence", "Occurs": [{}], "Items": [el("T/a"), el("T/b")]},
    ]})
    assert w.children["T/a"].rank == w.children["T/b"].rank
    assert not w.children["T/a"].max_one


def test_any_is_recorded():
    w = _ParticleWalker()
    w.walk({"Kind": "Sequence", "Items": [{"Kind": "Any", "Occurs": [{}]}]})
    assert w.any_child


SDK = os.environ.get("OOX_SDK_DATA")


@pytest.mark.skipif(not SDK, reason="set OOX_SDK_DATA=<Open-XML-SDK>/data to run")
def test_committed_code_matches_regeneration(tmp_path):
    from tools.gen import emit, overrides, sdk_frontend

    schema = sdk_frontend.load(SDK)
    overrides.apply(schema)
    emit.emit(schema, tmp_path)
    src = Path(__file__).parents[1] / "src" / "oox"
    for sub in ("_generated", "ns"):
        cmp = filecmp.dircmp(src / sub, tmp_path / sub, ignore=["__pycache__"])
        assert not (cmp.diff_files or cmp.left_only or cmp.right_only), sub


def test_stub_keyword_rule_matches_runtime():
    """tools/gen/emit.kwarg_names (for the .pyi stubs) must name keywords
    exactly as oox.grammar._kwarg_map does at run time."""
    import importlib
    import pkgutil

    import oox._generated as gen
    from oox import grammar
    from tools.gen.emit import kwarg_names

    checked = 0
    for m in pkgutil.iter_modules(gen.__path__):
        if not m.name.startswith("g_"):
            continue
        for key, entry in importlib.import_module(f"oox._generated.{m.name}").TYPES.items():
            qname = entry[0]
            if not qname or not entry[2]:
                continue
            specs = {a[0]: grammar.AttrSpec(*a) for a in entry[2]}
            runtime = {k: v.qname for k, v in grammar._kwarg_map(specs, qname.partition(":")[0]).items()}
            assert kwarg_names(list(specs), qname.partition(":")[0]) == runtime, key
            checked += 1
    assert checked > 2500          # every type with attributes


def test_stubs_list_attribute_keywords():
    from pathlib import Path

    stub = (Path(__file__).parents[1] / "src" / "oox" / "ns" / "w.pyi").read_text(encoding="utf-8")
    jc = stub[stub.index("def Jc("):].split(") -> OoxElement", 1)[0]
    assert "val: Literal['left', 'start', 'center'" in jc and "**attrs: object" in jc
    assert (Path(__file__).parents[1] / "src" / "oox" / "py.typed").exists()
