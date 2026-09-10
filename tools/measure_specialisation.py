#!/usr/bin/env python3
"""Measure what cross-operation specialisation prevents, on any workbook.

Version 6.6.

The README quotes a figure for the reference workbook. This is the script that
produces it, so the figure is reproducible rather than asserted.

Both numbers are counted from the workbook tree rather than from the generated
document, because counting the document would confuse two effects: how many
attributes an operation shows, and how a `$ref` graph happens to be shared.
The two figures here differ in exactly one respect, whether a group that
several operations share is allowed to differ between them.

    with specialisation   a leaf counts for an operation when the System of
                          Record behind that operation supplies it
    without it            a leaf counts for an operation when any operation
                          sharing the group supplies it, which is the single
                          shared shape a namespace without specialisation
                          would give

A container counts as one occurrence itself, and only when something beneath it
counts, which is the same rule the generator prunes by.

Usage:
    python tools/measure_specialisation.py samples/creditcard_v2.0.0.xlsx
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import polymorphize_classify as cls           # noqa: E402
import polymorphize_mapping as mapfmt         # noqa: E402
import polymorphize_workbook as wbk           # noqa: E402

__version__ = "6.6"


def key(name):
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def roots(result):
    out = []
    for section in (result.request, result.response):
        if section is not None and section.root is not None:
            out.append(section.root)
    return out


def collect_union(node, endpoints, union):
    """Record, per group name, every child any operation supplies."""
    for child in node.children or []:
        if child.children:
            collect_union(child, endpoints, union)
        elif any(child.sor.get(e) for e in endpoints):
            union.setdefault(key(node.name), set()).add(key(child.name))


def count(node, endpoints, union, shared):
    total = 0
    for child in node.children or []:
        if child.children:
            beneath = count(child, endpoints, union, shared)
            if beneath:
                total += 1 + beneath
        else:
            here = any(child.sor.get(e) for e in endpoints)
            if here:
                total += 1
            elif shared and key(child.name) in union.get(key(node.name), set()):
                total += 1
    return total


def read(path):
    verdict = cls.classify(path)
    if verdict.fmt == cls.MAPPING:
        return mapfmt.read_workbook(path), verdict.label
    if verdict.fmt == cls.UNKNOWN:
        raise SystemExit("%s: %s" % (path, verdict.line()))
    return wbk.read_workbook(path), verdict.label


def measure(path):
    results, label = read(path)
    ok = [r for r in results if r.status == "ok" and roots(r)]

    union = {}
    for result in ok:
        for root in roots(result):
            collect_union(root, list(result.endpoints or []), union)

    specialised = shared = 0
    for result in ok:
        endpoints = list(result.endpoints or [])
        for root in roots(result):
            specialised += count(root, endpoints, union, shared=False)
            shared += count(root, endpoints, union, shared=True)
    return label, len(ok), specialised, shared


def main(argv):
    if len(argv) != 2:
        print(__doc__)
        return 2
    label, sheets, specialised, shared = measure(argv[1])
    print("Workbook:            %s" % os.path.basename(argv[1]))
    print("Input format:        %s" % label)
    print("Operations counted:  %d" % sheets)
    print()
    print("With specialisation: %d attribute occurrences" % specialised)
    print("Without it:          %d" % shared)
    if specialised:
        print("Prevented:           %d occurrences, %.0f%% more"
              % (shared - specialised,
                 (shared - specialised) * 100.0 / specialised))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
