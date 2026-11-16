"""Full content-model checking: does an element's sequence of children match
its schema particle (sequences, choices, ``all`` groups, min/max occurrences)?

The basic validator checks each child on its own (allowed? singleton? in
order?).  This module checks the whole sequence, which also catches a missing
required child inside a choice, two alternatives of an exclusive choice, too
many repetitions, and a group that is started but not finished.

Each particle tree (generated into ``_generated/p_<prefix>.py``) is compiled
once into a Thompson NFA; matching simulates it over the children's qnames.
Children the particle doesn't mention at all are left to the basic validator
(it reports unknown ones; ignorable extension elements are fine), as are
``mc:AlternateContent`` blocks.
"""

from __future__ import annotations

import importlib
from functools import cache

ANY = object()
_UNROLL = 8          # repeat bounds above this are treated as unbounded
_STATES_LIMIT = 4000  # very large models: give up rather than compile slowly


class _NFA:
    __slots__ = ("accept", "alphabet", "eps", "has_any", "start", "trans")

    def __init__(self) -> None:
        self.trans: list[list[tuple[object, int]]] = []
        self.eps: list[list[int]] = []
        self.alphabet: set[str] = set()
        self.has_any = False

    def state(self) -> int:
        if len(self.trans) > _STATES_LIMIT:
            raise OverflowError
        self.trans.append([])
        self.eps.append([])
        return len(self.trans) - 1

    def build(self, node: tuple) -> tuple[int, int]:
        """Fragment ``(start, end)`` for ``node`` with its occurrence bounds."""
        if node[0] == "E":
            mn, mx = node[2], node[3]
        elif node[0] == "A":
            mn, mx = node[1], node[2]
        else:
            mn, mx = node[1], node[2]
        if node[0] == "L":           # xsd:all: any order, approximated as (a|b|...)*
            mx = None
        if mx is not None and mx > _UNROLL:
            mx = None
        mn = min(mn, _UNROLL)
        start = cur = self.state()
        for _ in range(mn):
            s, e = self._once(node)
            self.eps[cur].append(s)
            cur = e
        if mx is None:
            s, e = self._once(node)
            self.eps[cur].append(s)
            self.eps[e].append(cur)     # loop back: zero or more further copies
            end = self.state()
            self.eps[cur].append(end)
            return start, end
        end = self.state()
        self.eps[cur].append(end)
        for _ in range(mx - mn):
            s, e = self._once(node)
            self.eps[cur].append(s)
            self.eps[e].append(end)
            cur = e
        return start, end

    def _once(self, node: tuple) -> tuple[int, int]:
        kind = node[0]
        s, e = self.state(), self.state()
        if kind == "E":
            self.trans[s].append((node[1], e))
            self.alphabet.add(node[1])
        elif kind == "A":
            self.trans[s].append((ANY, e))
            self.has_any = True
        elif kind == "S":
            cur = s
            for item in node[3]:
                fs, fe = self.build(item)
                self.eps[cur].append(fs)
                cur = fe
            self.eps[cur].append(e)
        else:  # "C" choice, "L" all (one item per pass)
            if not node[3]:
                self.eps[s].append(e)
            for item in node[3]:
                fs, fe = self.build(item)
                self.eps[s].append(fs)
                self.eps[fe].append(e)
        return s, e

    def closure(self, states) -> frozenset[int]:
        seen = set(states)
        todo = list(states)
        while todo:
            for n in self.eps[todo.pop()]:
                if n not in seen:
                    seen.add(n)
                    todo.append(n)
        return frozenset(seen)

    def step(self, states: frozenset[int], q: str) -> frozenset[int]:
        nxt = [t for s in states for label, t in self.trans[s] if label == q or label is ANY]
        return self.closure(nxt) if nxt else frozenset()

    def expected(self, states: frozenset[int]) -> list[str]:
        out = sorted({label for s in states for label, _ in self.trans[s]
                      if isinstance(label, str)})
        if any(label is ANY for s in states for label, _ in self.trans[s]):
            out.append("(any element)")
        return out


@cache
def _particles(mod: str) -> dict:
    try:
        return importlib.import_module(f"oox._generated.p_{mod}").PARTICLES
    except ModuleNotFoundError:
        return {}


@cache
def compiled(key: str) -> _NFA | None:
    """The NFA for a type key (``w:CT_Tbl/w:tbl``), or ``None`` if unknown."""
    tree = _particles(key.partition("/")[0].partition(":")[0]).get(key)
    if tree is None:
        return None
    nfa = _NFA()
    try:
        s, e = nfa.build(tree)
    except (OverflowError, RecursionError):
        return None
    nfa.start, nfa.accept = s, e
    return nfa


def check(key: str, tag: str, child_qnames: list[str]) -> str | None:
    """``None`` if the children match the particle of ``key``, else a message."""
    nfa = compiled(key)
    if nfa is None:
        return None
    seq = [q for q in child_qnames
           if q in nfa.alphabet or (nfa.has_any and not q.startswith("mc:"))]
    states = nfa.closure([nfa.start])
    for i, q in enumerate(seq):
        nxt = nfa.step(states, q)
        if not nxt:
            exp = nfa.expected(states)
            where = "here" if i else "first"
            return (f"<{tag}>: <{q}> is not allowed {where} (child {i + 1} of {len(seq)}); "
                    f"expected {_list(exp) if exp else 'no more children'}")
        states = nxt
    if nfa.accept not in states:
        exp = nfa.expected(states)
        return f"<{tag}>: incomplete, missing {_list(exp)}"
    return None


def _list(names: list[str]) -> str:
    if len(names) > 8:
        return "one of " + ", ".join(f"<{n}>" for n in names[:8]) + f" (+{len(names) - 8})"
    if len(names) > 1:
        return "one of " + ", ".join(f"<{n}>" for n in names)
    return f"<{names[0]}>" if names else "nothing"
