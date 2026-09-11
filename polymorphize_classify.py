"""
polymorphize_classify — which input format is this workbook?

Version 6.8.

Two input formats are in use and they are told apart by structure rather than
by style, so the test is cheap and unambiguous.

Level format
    Header on row 1 carrying ``Level 1`` to ``Level N``, banner in rows 2 to 6
    **below** it, one column per SOR endpoint headed ``SOR <endpoint>``, and
    the section labels in the column to the left of ``Level 1``.

Field mapping format
    Banner in rows 1 to 9, header on row 10 carrying ``Parameter Type`` and a
    dotted path in ``Reusable API Field Name``, and a single
    ``SOR API Field Name`` column.

Why this module exists
======================

Before it, a field mapping workbook handed to the tool produced this::

    skip    Get All Operations    support sheet
    0 operation sheet(s): 0 would generate, 0 would fail. 0 error(s).
    exit code: 0

Every sheet was classified as a support sheet, because the Level reader
searches only the first six rows for a header and this format's header is on
row 10. The run reported nothing wrong and exited zero, so a build step would
go green having produced nothing. That is the failure mode this closes.

It classifies, it does not validate
===================================

The result says what the file is and how much of it, never whether it is good.
A correctly classified workbook can still be full of defects, and a pass
indicator here would be read as a promise this module cannot keep. Validation
stays in :mod:`polymorphize_validate`.

Speed
=====

Opening a workbook the ordinary way to look at its header row is a trap. On the
8.7 MB CASA mapping workbook, ``load_workbook(path)`` takes about 76 seconds,
where the bounded read-only peek below takes about 0.3 seconds, a factor of
266. At a file-entry field that difference is the window appearing to hang, so
the peek is always ``read_only=True`` and always bounded on both rows and
sheets. ``tests/test_workbook_contract.py`` holds it under :data:`PEEK_BUDGET`.
"""

from __future__ import annotations

import os
import re
import time
from dataclasses import dataclass, field

__version__ = "6.8"

#: Rows of each sampled sheet the peek reads. The Level banner ends by row 6
#: and the mapping header sits on row 10, so 14 covers both with room for a
#: sheet that carries a title row above its banner.
PEEK_ROWS = 14

#: Columns of each sampled row the peek reads. The widest sheet seen carries
#: 25, and a Level sheet can reach the low twenties with six Level columns and
#: five SOR columns.
PEEK_COLS = 30

#: Sheets sampled before a verdict is reached. The first sheet of a workbook is
#: often a cover or a vocabulary sheet, so one is not enough.
PEEK_SHEETS = 6

#: Seconds the peek is allowed on any workbook. Held by a test.
PEEK_BUDGET = 3.0

LEVEL = "level"
MAPPING = "mapping"
UNKNOWN = "unknown"

FORMAT_LABEL = {
    LEVEL: "Level format",
    MAPPING: "Field mapping document",
    UNKNOWN: "Not recognised",
}

LEVEL_RE = re.compile(r"^level\s*(\d+)$")


def key(value):
    """A cell reduced to letters and digits, for comparing headings."""
    return re.sub(r"[^a-z0-9]", "", str(value).lower()) if value is not None else ""


# --------------------------------------------------------------------------- #
# The peek
# --------------------------------------------------------------------------- #


def peek(path, rows=PEEK_ROWS, cols=PEEK_COLS, sheets=PEEK_SHEETS):
    """``[(sheet_title, [row, ...]), ...]`` for the top of the first sheets.

    Bounded on every axis and read-only throughout. Never open a workbook any
    other way to answer a question about its shape: see the module docstring.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        out = []
        for ws in wb.worksheets[:sheets]:
            grid = []
            for i, row in enumerate(ws.iter_rows(max_row=rows, max_col=cols,
                                                 values_only=True)):
                grid.append(list(row))
                if i + 1 >= rows:
                    break
            out.append((ws.title, grid))
        return out
    finally:
        try:
            wb.close()
        except Exception:                                      # noqa: BLE001
            pass


# --------------------------------------------------------------------------- #
# The verdict
# --------------------------------------------------------------------------- #


@dataclass
class SheetVerdict:
    title: str
    fmt: str = UNKNOWN
    header_row: int = 0
    detail: str = ""


@dataclass
class Classification:
    """What a file is, and how much of it. Never whether it is good."""

    path: str
    fmt: str = UNKNOWN
    sheets_sampled: int = 0
    sheets_total: int = 0
    operation_sheets: int = 0
    support_sheets: int = 0
    per_sheet: list = field(default_factory=list)
    mixed: bool = False
    readable: bool = True
    error: str = ""
    elapsed: float = 0.0

    @property
    def needs_specification(self):
        """A field mapping document describes an interface that already exists.

        The engine that consumes it trims a published specification rather than
        composing a new one, so a second input is required. The window reveals
        the field for it on this flag and the command line asks for it.
        """
        return self.fmt == MAPPING

    @property
    def label(self):
        return FORMAT_LABEL.get(self.fmt, FORMAT_LABEL[UNKNOWN])

    def line(self):
        """One line for a file field or a command line. States, never judges."""
        if not self.readable:
            return "Cannot read this file yet: %s" % self.error
        if self.fmt == UNKNOWN:
            return ("Not recognised: no Level columns and no Parameter Type "
                    "column found in the first %d rows of %d sheet(s)"
                    % (PEEK_ROWS, self.sheets_sampled))
        if self.sheets_total > self.sheets_sampled:
            # Be explicit that the counts describe the sample, not the file, or
            # "6 operation sheets" reads as a total when 13 sheets exist.
            bits = ["%d sheets" % self.sheets_total,
                    "%d sampled, %d of those carry an operation"
                    % (self.sheets_sampled, self.operation_sheets)]
        else:
            bits = ["%d operation sheet%s" % (self.operation_sheets,
                                              "" if self.operation_sheets == 1 else "s")]
            if self.support_sheets:
                bits.append("%d support sheet%s" % (self.support_sheets,
                                                    "" if self.support_sheets == 1 else "s"))
        out = "%s · %s" % (self.label, " · ".join(bits))
        if self.needs_specification:
            out += " · needs a specification to trim"
        if self.mixed:
            out += " · mixed formats in one workbook"
        return out


def _sheet_verdict(title, grid):
    """Classify one sheet from its top rows alone."""
    for i, row in enumerate(grid, start=1):
        keys = [key(c) for c in row]
        if any(LEVEL_RE.match(k) for k in keys):
            n = sum(1 for k in keys if LEVEL_RE.match(k))
            return SheetVerdict(title, LEVEL, i, "%d Level column(s)" % n)
        if "parametertype" in keys:
            named = [k for k in keys if k in ("reusableapifieldname", "apifieldname",
                                              "fieldname", "attributename")]
            return SheetVerdict(title, MAPPING, i,
                                "Parameter Type%s" % (" and a field name column"
                                                      if named else ""))
    return SheetVerdict(title, UNKNOWN, 0, "no recognised header row")


def classify(path):
    """What format is this workbook, and how much of it. Never raises."""
    out = Classification(path=path)
    started = time.perf_counter()
    try:
        sampled = peek(path)
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            out.sheets_total = len(wb.sheetnames)
        finally:
            try:
                wb.close()
            except Exception:                                  # noqa: BLE001
                pass
    except Exception as exc:                                   # noqa: BLE001
        # Fail soft. These files are routinely open in Excel while being
        # dropped on a field, and a classifier must never block the run.
        out.readable = False
        out.error = "%s: %s" % (type(exc).__name__, exc)
        out.elapsed = time.perf_counter() - started
        return out

    out.per_sheet = [_sheet_verdict(t, g) for t, g in sampled]
    out.sheets_sampled = len(out.per_sheet)
    seen = {v.fmt for v in out.per_sheet if v.fmt != UNKNOWN}
    out.operation_sheets = sum(1 for v in out.per_sheet if v.fmt != UNKNOWN)
    out.support_sheets = sum(1 for v in out.per_sheet if v.fmt == UNKNOWN)

    if not seen:
        out.fmt = UNKNOWN
    elif len(seen) == 1:
        out.fmt = next(iter(seen))
    else:
        # Both formats in one workbook. The Level format wins, because it is
        # the format the generate path was built for, and the mixture is
        # reported rather than resolved silently.
        out.fmt = LEVEL
        out.mixed = True
    out.elapsed = time.perf_counter() - started
    return out


def describe(path):
    """The one-line verdict, for a caller that wants nothing else."""
    return classify(path).line()


__all__ = ["LEVEL", "MAPPING", "UNKNOWN", "FORMAT_LABEL", "PEEK_ROWS",
           "PEEK_COLS", "PEEK_SHEETS", "PEEK_BUDGET", "Classification",
           "SheetVerdict", "peek", "classify", "describe"]
