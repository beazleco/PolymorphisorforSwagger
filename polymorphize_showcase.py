"""
polymorphize_showcase — the HTML report that shows what was eliminated.

Version 6.8.

This is the validation instrument. It exists so a business analyst can open one
file and answer, for every class in the generated specification and every
endpoint that uses it, two questions:

* which elements are published, and
* for every element that is **not** published, why not

Nothing here is inferred from the generated specification. Every row is taken
from the workbook's own tree and annotated with the decision the tool made
about it, so the report is an account of the tool's reasoning rather than a
description of its output. That is what makes it usable as evidence.

The five dispositions
=====================

===========================  ==============================================
Published                    an SOR field backs it, so it reaches the interface
No SOR field                 the workbook left the SOR cell empty
Not in the SOR               the workbook named a field that does not exist
                             on the SOR endpoint, so it was excluded
Variant only                 published by some variants of the operation and
                             not others
Container                    a group, published only because something
                             beneath it is
===========================  ==============================================

A sixth line appears only if something has gone wrong: an element that is
mapped and verified but did not reach the interface. That should always be
zero, and the report says so, because a silent loss is the one failure mode
this tool must not have.
"""

from __future__ import annotations

import datetime as _dt
import html as _h
import os
from collections import OrderedDict, defaultdict

import polymorphize_generate as gen
import polymorphize_sor as sor
import polymorphize_sor as sormod
import polymorphize_workbook as wbk
from polymorphize_workbook import VARIANT_PROPERTY, cell_ref, text

__version__ = "6.8"

PUBLISHED = "published"
NO_FIELD = "no-field"
NOT_IN_SOR = "not-in-sor"
VARIANT_ONLY = "variant-only"
CONTAINER = "container"
LOST = "lost"

DISPOSITION_LABEL = {
    PUBLISHED: "Published",
    NO_FIELD: "No SOR field",
    NOT_IN_SOR: "Not in the SOR",
    VARIANT_ONLY: "Variant only",
    CONTAINER: "Container",
    LOST: "LOST",
}

DISPOSITION_NOTE = {
    PUBLISHED: "An SOR field backs this element on every variant of the "
               "operation, so it is published.",
    NO_FIELD: "The SOR column is empty, so the workbook says the System of "
              "Record does not supply this element. It is removed.",
    NOT_IN_SOR: "The workbook named an SOR field, but no element of that name "
                "exists in the relevant message of the SOR endpoint the "
                "sheet's banner names. It is removed.",
    VARIANT_ONLY: "Backed by the SOR on some variants of this operation and "
                  "not others, so it appears only in those variants.",
    CONTAINER: "A group rather than an element. It is published only where "
               "something beneath it is.",
    LOST: "Mapped and verified, yet absent from the generated interface. This "
          "is a fault in the tool, not in the workbook.",
}


# --------------------------------------------------------------------------- #
# Working out what happened to each element
# --------------------------------------------------------------------------- #


class Row:
    """One declared element, with the decision made about it."""

    __slots__ = ("node", "depth", "path", "disposition", "per_variant",
                 "declared", "resolutions", "cell")

    def __init__(self, node, depth, path):
        self.node = node
        self.depth = depth
        self.path = path
        self.disposition = NO_FIELD
        self.per_variant = OrderedDict()   # endpoint -> (state, detail)
        self.declared = OrderedDict()      # endpoint -> the workbook's word
        self.resolutions = OrderedDict()   # endpoint -> Resolution or None
        self.cell = ""


class ClassView:
    """One class of the specification, as it appears in one message."""

    def __init__(self, name, kind, node):
        self.name = name
        self.kind = kind                  # "message" | "group" | "array item"
        self.node = node
        self.rows = []

    @property
    def counts(self):
        c = defaultdict(int)
        for r in self.rows:
            c[r.disposition] += 1
        return c


class MessageView:
    def __init__(self, label, root):
        self.label = label
        self.root = root
        self.classes = []


class EndpointView:
    def __init__(self, result, spec):
        self.result = result
        self.spec = spec
        self.messages = []
        self.endpoints = []
        self.variants = []
        self.verification = None

    @property
    def sheet(self):
        return self.result.sheet

    @property
    def api(self):
        """The integration API this sheet publishes: the thing consumers call."""
        return "%s %s" % (self.spec.method.upper(), self.spec.path)

    @property
    def operation_id(self):
        op = self.spec.document["paths"][self.spec.path][self.spec.method]
        return op.get("operationId", "")

    @property
    def eliminated(self):
        """Every declared element the application kept out of the interface."""
        out = []
        for m in self.messages:
            for cls in m.classes:
                for row in cls.rows:
                    if row.disposition in (NO_FIELD, NOT_IN_SOR):
                        out.append((m.label, cls, row))
        return out

    @property
    def narrowed(self):
        """Elements published by some variants of this operation and not all."""
        out = []
        for m in self.messages:
            for cls in m.classes:
                for row in cls.rows:
                    if row.disposition == VARIANT_ONLY:
                        out.append((m.label, cls, row))
        return out

    @property
    def counts(self):
        c = defaultdict(int)
        for m in self.messages:
            for cls in m.classes:
                for k, v in cls.counts.items():
                    c[k] += v
        return c


def _classify(node, endpoints, verification, published_paths, path):
    """The disposition of one element."""
    row = Row(node, 0, path)
    layout_cells = {}
    for endpoint in endpoints:
        row.declared[endpoint] = node.sor_declared.get(endpoint, "")
        res = verification.of(node, endpoint) if verification else None
        row.resolutions[endpoint] = res
        effective = bool(node.sor.get(endpoint))
        if effective:
            state = "yes"
            detail = ", ".join(res.at[:3]) if res and res.at else ""
        elif row.declared[endpoint] and res and res.status == sormod.MISSING:
            state = "not-in-sor"
            detail = res.declared
        elif row.declared[endpoint]:
            state = "not-in-sor"
            detail = row.declared[endpoint]
        else:
            state = "no"
            detail = ""
        row.per_variant[endpoint] = (state, detail)

    states = {s for s, _d in row.per_variant.values()}
    if node.children:
        row.disposition = CONTAINER
    elif "yes" in states and len(endpoints) > 1 and \
            ("no" in states or "not-in-sor" in states):
        row.disposition = VARIANT_ONLY
    elif "yes" in states:
        row.disposition = PUBLISHED
    elif "not-in-sor" in states:
        row.disposition = NOT_IN_SOR
    else:
        row.disposition = NO_FIELD

    # The safety check: mapped, verified, and yet not in the interface.
    if row.disposition in (PUBLISHED, VARIANT_ONLY) and published_paths is not None \
            and path not in published_paths and not node.children:
        row.disposition = LOST
    return row


def _published_leaf_paths(spec):
    """Every leaf path the generated document actually publishes."""
    if spec.document is None:
        return None
    schemas = spec.document.get("components", {}).get("schemas", {})
    out = set()

    def walk(node, path, seen):
        if not isinstance(node, dict):
            return
        ref = node.get("$ref")
        if isinstance(ref, str):
            name = ref.rsplit("/", 1)[-1]
            if name in seen:
                return
            walk(schemas.get(name, {}), path, seen | {name})
            return
        for kw in ("allOf", "oneOf", "anyOf"):
            if isinstance(node.get(kw), list):
                for sub in node[kw]:
                    walk(sub, path, seen)
        if node.get("type") == "array":
            walk(node.get("items") or {}, path, seen)
            return
        props = node.get("properties")
        if isinstance(props, dict):
            for name, sub in props.items():
                out.add(path + (name,))
                walk(sub, path + (name,), seen)

    for name, schema in schemas.items():
        walk(schema, (), {name})
    return {p[-1:] and p for p in out}


def disposition_of(node, endpoints, verification, published_paths, path):
    """The decision made about one element, as a :class:`Row`.

    Public because the field mapping exporter needs the same verdict the
    showcase draws, and the two must never disagree. Computing it twice by two
    routes is how a report and a spreadsheet end up telling an analyst
    different things about the same cell.
    """
    return _classify(node, endpoints, verification, published_paths, path)


def published_leaf_paths(spec):
    """Every leaf path a generated specification actually publishes."""
    return _published_leaf_paths(spec)


def _collect_classes(root, endpoints, verification, published, message):
    """The classes of one message, in the order the workbook declares them."""
    if root is None:
        return

    def add_class(name, kind, node, path_prefix):
        view = ClassView(name, kind, node)
        message.classes.append(view)
        for child in node.children:
            path = path_prefix + (child.name,)
            row = _classify(child, endpoints, verification, published, path)
            row.depth = 0
            row.cell = cell_ref(child.row, 0)
            view.rows.append(row)
        for child in node.children:
            if not child.children:
                continue
            path = path_prefix + (child.name,)
            if child.json_type == "array":
                kids = child.children
                if len(kids) == 1 and kids[0].children and \
                        kids[0].json_type == "object":
                    add_class(kids[0].name, "array item", kids[0], path)
                else:
                    add_class(child.name + "Item", "array item", child, path)
            else:
                add_class(child.name, "group", child, path)

    add_class(root.name, "message", root, ())


def build_views(out):
    """One :class:`EndpointView` per generated endpoint."""
    views = []
    for spec in out.specs:
        if spec.status != "ok":
            continue
        result = next((r for r in _results_of(out) if r.sheet == spec.sheet), None)
        if result is None:
            continue
        view = EndpointView(result, spec)
        view.endpoints = ([c.endpoint for c in result.layout.sor_cols]
                          if result.layout is not None else [])
        view.variants = spec.variants or []
        view.verification = out.verifications.get(spec.sheet)
        published = _published_leaf_paths(spec)
        for label, section in (("Request", result.request),
                               ("Response", result.response)):
            if section is None or section.root is None:
                continue
            m = MessageView(label, section.root)
            _collect_classes(section.root, view.endpoints, view.verification,
                             published, m)
            view.messages.append(m)
        views.append(view)
    return views


_RESULT_CACHE = {}


def _results_of(out):
    """The sheet results behind a run.

    The run carries the results it actually read, so this is normally a
    lookup. Until 6.6 it re-read the workbook with the Level reader whatever
    the input format was, which meant that for a field mapping document every
    sheet came back empty and the showcase reported no elements at all. The
    re-read survives only for a caller that builds a ``RunResult`` by hand,
    and it now dispatches on the format.
    """
    carried = getattr(out, "results", None)
    if carried:
        return carried
    cached = getattr(out, "_showcase_results", None)
    if cached is None:
        if getattr(out, "input_format", "level") == "mapping":
            import polymorphize_mapping as mapfmt
            cached = mapfmt.read_workbook(out.workbook)
        else:
            cached = wbk.read_workbook(out.workbook)
        # Re-apply verification so the trees match what generation saw.
        if out.sor_index is not None and not out.sor_index.empty:
            sormod.verify_workbook(cached, out.sor_index)
        out._showcase_results = cached
    return cached


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #

CSS = """
:root{
  --ink:#16202c; --muted:#5c6a7a; --faint:#8a97a6; --line:#dbe3ec;
  --bg:#f2f5f9; --panel:#ffffff; --accent:#1f3864; --accent-soft:#e8eef8;
  --ok-bg:#e6f4ea; --ok-fg:#1c6b38; --ok-line:#a8d8b9;
  --no-bg:#eef1f5; --no-fg:#5c6a7a; --no-line:#cdd6e0;
  --bad-bg:#fdeaea; --bad-fg:#a11b1b; --bad-line:#f0b8b8;
  --var-bg:#fff6e0; --var-fg:#8a6100; --var-line:#ecd39a;
  --con-bg:#eef4fb; --con-fg:#2c4f80; --con-line:#c4d6ea;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
 font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
code,.mono{font-family:ui-monospace,SFMono-Regular,Consolas,"Liberation Mono",monospace;
 font-size:12.5px}
a{color:var(--accent)}
.wrap{display:grid;grid-template-columns:262px minmax(0,1fr);gap:0;
 min-height:100vh;align-items:start}
nav{position:sticky;top:0;max-height:100vh;overflow:auto;background:var(--panel);
 border-right:1px solid var(--line);padding:22px 0}
nav h2{font-size:11px;letter-spacing:.09em;text-transform:uppercase;
 color:var(--faint);margin:20px 20px 8px;font-weight:600}
nav a{display:block;padding:6px 20px;text-decoration:none;color:var(--ink);
 border-left:3px solid transparent;font-size:13px}
nav a:hover{background:var(--accent-soft)}
nav a.fail{color:var(--bad-fg)}
nav .badge{float:right;color:var(--faint);font-size:11.5px}
main{padding:26px 34px 90px;max-width:1500px;min-width:0}
header.top{margin-bottom:22px}
h1{font-size:25px;margin:0 0 4px;color:var(--accent);letter-spacing:-.2px}
.sub{color:var(--muted);margin:0}
.callout{margin:18px 0;padding:14px 16px;border-radius:8px;border:1px solid;
 background:var(--ok-bg);border-color:var(--ok-line);color:var(--ok-fg)}
.callout.warn{background:var(--var-bg);border-color:var(--var-line);color:var(--var-fg)}
.callout.bad{background:var(--bad-bg);border-color:var(--bad-line);color:var(--bad-fg)}
.callout strong{display:block;margin-bottom:3px;font-size:14.5px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));
 gap:12px;margin:18px 0 26px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:8px;
 padding:14px 16px}
.card .n{font-size:27px;font-weight:600;letter-spacing:-.5px;line-height:1.1}
.card .k{color:var(--muted);font-size:12px;margin-top:3px}
.card.ok .n{color:var(--ok-fg)} .card.bad .n{color:var(--bad-fg)}
.card.var .n{color:var(--var-fg)} .card.no .n{color:var(--no-fg)}
section{background:var(--panel);border:1px solid var(--line);border-radius:9px;
 padding:20px 22px;margin:0 0 20px}
section > h2{margin:0 0 3px;font-size:18px;color:var(--accent)}
section > .meta{color:var(--muted);font-size:12.5px;margin:0 0 14px}
h3{font-size:14px;margin:22px 0 8px;color:var(--accent)}
h4{font-size:13px;margin:16px 0 6px;color:var(--ink)}
h4 .kind{color:var(--faint);font-weight:400;font-size:12px;margin-left:6px}
table{border-collapse:collapse;width:100%;font-size:13px;margin:0 0 6px}
th,td{text-align:left;padding:5px 8px;border-bottom:1px solid var(--line);
 vertical-align:top}
table.grid{table-layout:fixed}
.c-el{width:210px}
.c-ty{width:96px}
.c-us{width:80px}
.c-di{width:112px}
.c-va{width:38px;text-align:center}
.c-sor{width:250px;overflow-wrap:anywhere}
.c-why{min-width:170px}
table.grid th.c-va{font-size:11px;padding:5px 2px;text-align:center}
h2 .verb,h3 .verb{display:inline-block;font-size:11.5px;font-weight:700;
 letter-spacing:.06em;background:var(--accent);color:#fff;border-radius:4px;
 padding:2px 7px;vertical-align:2px;margin-right:7px}
h2 code.api,h3 code.api{font-size:16px;color:var(--ink);font-weight:600}
h3 code.api{font-size:13.5px}
h3.sep{margin-top:26px;padding-top:16px;border-top:1px solid var(--line)}
nav .navverb{display:inline-block;min-width:38px;font-size:10.5px;
 font-weight:700;color:var(--faint);letter-spacing:.04em}
table.elim{table-layout:fixed;font-size:12.5px}
table.elim td{overflow-wrap:anywhere}
table.elim td.name{white-space:normal}
.e-ms{width:78px}.e-cl{width:180px}.e-el{width:180px}.e-pa{width:230px}
.e-re{width:118px}.e-so{width:180px}.e-wh{width:auto}
table.fam{table-layout:fixed;font-size:12.5px}
table.fam td{overflow-wrap:anywhere}
.f-cl{width:210px}.f-ba{width:200px}.f-de{width:300px}
.f-us{width:190px}.f-ad{width:auto}
table.legend-table{width:auto;margin:4px 0 10px;font-size:12.5px}
table.legend-table th,table.legend-table td{padding:3px 14px 3px 0;border:0}
table.legend-table thead th{background:none;position:static}
th{font-size:11px;letter-spacing:.05em;text-transform:uppercase;
 color:var(--faint);font-weight:600;background:#fafcfe;position:sticky;top:0}
tr.d-no td:first-child,tr.d-not-in-sor td:first-child{color:var(--muted)}
tr.d-not-in-sor{background:#fef7f7}
tr.d-lost{background:var(--bad-bg);font-weight:600}
td.name{white-space:nowrap}
td.name .ind{color:var(--faint)}
.pill{display:inline-block;padding:1px 8px;border-radius:9px;font-size:11.5px;
 font-weight:600;border:1px solid;white-space:nowrap}
.p-published{background:var(--ok-bg);color:var(--ok-fg);border-color:var(--ok-line)}
.p-no-field{background:var(--no-bg);color:var(--no-fg);border-color:var(--no-line)}
.p-not-in-sor{background:var(--bad-bg);color:var(--bad-fg);border-color:var(--bad-line)}
.p-variant-only{background:var(--var-bg);color:var(--var-fg);border-color:var(--var-line)}
.p-container{background:var(--con-bg);color:var(--con-fg);border-color:var(--con-line)}
.p-lost{background:var(--bad-fg);color:#fff;border-color:var(--bad-fg)}
.tick{color:var(--ok-fg);font-weight:700}
.cross{color:var(--faint)}
.xsor{color:var(--bad-fg);font-weight:700}
.why{color:var(--muted);font-size:12px}
details{margin:8px 0}
summary{cursor:pointer;color:var(--accent);font-size:13px;font-weight:600;
 padding:4px 0}
.legend{display:flex;flex-wrap:wrap;gap:8px 18px;margin:8px 0 4px}
.legend div{font-size:12.5px;color:var(--muted)}
.scroll{overflow-x:auto}
.kv{display:grid;grid-template-columns:auto 1fr;gap:3px 14px;font-size:13px;
 margin:0 0 14px}
.kv dt{color:var(--muted)} .kv dd{margin:0}
.small{font-size:12px;color:var(--muted)}
footer{color:var(--faint);font-size:12px;padding:6px 34px 40px}
@media print{
  nav{display:none} .wrap{grid-template-columns:1fr} main{padding:0}
  section{break-inside:avoid;box-shadow:none}
  th{position:static}
}
"""


def e(v):
    return _h.escape("" if v is None else str(v))


def _pill(disposition):
    return '<span class="pill p-%s">%s</span>' % (
        disposition, e(DISPOSITION_LABEL[disposition]))


def _variant_header(view):
    """Column headings for the per-variant matrix.

    The heading is the variant code alone, because five columns of prose makes
    every row three lines tall. The codes are expanded in a legend directly
    above the tables, where there is room to name the SOR endpoint too.
    """
    if len(view.endpoints) <= 1:
        return []
    if view.variants:
        return [(v.code, v.endpoint, v.label) for v in view.variants]
    return [("%d" % (i + 1), ep, ep or "(unnamed)")
            for i, ep in enumerate(view.endpoints)]


def _variant_legend(view, columns):
    if not columns:
        return ""
    items = "".join(
        '<tr><td class="mono">%s</td><td>%s</td><td class="mono">%s</td></tr>'
        % (e(code), e(label), e(endpoint or "(from the banner)"))
        for code, endpoint, label in columns)
    return ('<p class="small">This operation is served by %d SOR endpoints. '
            'The numbered columns in the tables below are these variants.</p>'
            '<div class="scroll"><table class="legend-table"><thead><tr>'
            '<th>Variant</th><th>Name</th><th>SOR endpoint</th></tr></thead>'
            '<tbody>%s</tbody></table></div>' % (len(columns), items))


def _row_html(row, view, columns):
    cells = ['<td class="name c-el"><code>%s</code></td>' % e(row.node.name)]
    kind = row.node.json_type or "string"
    if row.node.max_length:
        kind += " (%d)" % row.node.max_length
    cells.append('<td class="mono c-ty">%s</td>' % e(kind))
    usage = {wbk.USAGE_REQUIRED: "Mandatory",
             wbk.USAGE_CONDITIONAL: "Conditional"}.get(row.node.usage, "")
    cells.append('<td class="c-us">%s</td>' % e(usage))
    cells.append('<td class="c-di">%s</td>' % _pill(row.disposition))

    if columns:
        for _code, endpoint, _label in columns:
            state, detail = row.per_variant.get(endpoint, ("no", ""))
            if state == "yes":
                mark = '<span class="tick" title="%s">&#10003;</span>' % e(detail)
            elif state == "not-in-sor":
                mark = '<span class="xsor" title="%s">&#10007;</span>' % e(
                    "named %s, absent from the SOR" % detail)
            else:
                mark = '<span class="cross">&middot;</span>'
            cells.append('<td class="c-va">%s</td>' % mark)

    declared = [v for v in row.declared.values() if v]
    cells.append('<td class="mono c-sor">%s</td>'
                 % (e(" | ".join(sorted(set(declared)))) or
                    '<span class="why">empty</span>'))

    why = _why(row, view)
    cells.append('<td class="why c-why">%s</td>' % why)
    return '<tr class="d-%s">%s</tr>' % (row.disposition, "".join(cells))


def _why(row, view):
    """The one-line explanation for this element."""
    if row.disposition == CONTAINER:
        return ""
    if row.disposition == NO_FIELD:
        return e("the SOR column is empty")
    if row.disposition == NOT_IN_SOR:
        bad = [(ep, d) for ep, (s, d) in row.per_variant.items()
               if s == "not-in-sor"]
        if bad:
            ep, declared = bad[0]
            return e("%s does not exist in %s" % (declared, ep or "the SOR endpoint"))
        return e("named an SOR field that does not exist")
    if row.disposition == VARIANT_ONLY:
        yes = [ep for ep, (s, _d) in row.per_variant.items() if s == "yes"]
        return e("only %s" % ", ".join(x or "(unnamed)" for x in yes))
    if row.disposition == LOST:
        return e("mapped and verified, yet not published. Report this.")
    res = [r for r in row.resolutions.values() if r and r.at]
    if res:
        first = res[0]
        if first.status == sormod.AMBIGUOUS:
            return e("matched %s in %d places: %s"
                     % (first.matched, len(first.at), ", ".join(first.at[:2])))
        return e("matched %s" % (first.at[0] if first.at else first.matched))
    return ""


def _class_table(cls, view, columns):
    head = [("Element", "c-el"), ("Type", "c-ty"), ("Usage", "c-us"),
            ("Disposition", "c-di")]
    head += [(c[0], "c-va") for c in columns]
    head += [("SOR field in the workbook", "c-sor"), ("Why", "c-why")]
    rows = "".join(_row_html(r, view, columns) for r in cls.rows)
    c = cls.counts
    caption = ", ".join(
        "%d %s" % (c[k], DISPOSITION_LABEL[k].lower())
        for k in (PUBLISHED, VARIANT_ONLY, NOT_IN_SOR, NO_FIELD, CONTAINER, LOST)
        if c.get(k))
    return (
        '<h4><code>%s</code><span class="kind">%s &middot; %s</span></h4>'
        '<div class="scroll"><table class="grid"><thead><tr>%s</tr></thead>'
        '<tbody>%s</tbody></table></div>'
        % (e(cls.name), e(cls.kind), e(caption),
           "".join('<th class="%s">%s</th>' % (klass, e(h))
                   for h, klass in head), rows))


def _endpoint_section(view, anchor):
    r, spec = view.result, view.spec
    banner = r.banner
    columns = _variant_header(view)
    c = view.counts
    total = sum(c.values())
    kept = c.get(PUBLISHED, 0) + c.get(VARIANT_ONLY, 0)

    kv = [
        ("Operation", "<code>%s</code>" % e(view.operation_id)),
        ("Pattern", e(spec.pattern) + " &middot; " + e({
            "P1": "one SOR endpoint, no polymorphism needed",
            "P2": "several SOR endpoints chosen at run time by %s"
                  % VARIANT_PROPERTY,
            "P3": "several SOR endpoints, chosen by which identifier is sent",
        }.get(spec.pattern, ""))),
    ]
    if banner.get("service_domain"):
        kv.append(("BIAN service domain",
                   e(banner["service_domain"]) +
                   (" &middot; BQ " + e(banner["behaviour_qualifier"])
                    if banner.get("behaviour_qualifier") else "")))
    if banner.get("use_case"):
        kv.append(("Use case", e(text(banner["use_case"])[:400])))
    sor_paths = [d.endpoint for d in (view.verification.decisions
                                     if view.verification else [])
                 if getattr(d, "endpoint", "")]
    if not sor_paths:
        sor_paths = [x for x in sor.split_endpoints(
            (banner or {}).get("sor_endpoint", ""))]
    kv.append(("Backed by, downstream",
               '<span class="why">%s</span>'
               % e(", ".join(sor_paths) or "not stated")))
    v = view.verification
    if v is not None and v.checked:
        kv.append(("SOR verification",
                   "%d field names resolved, %d ambiguous, "
                   "<strong>%d absent from the SOR</strong>"
                   % (v.counts.get(sormod.RESOLVED, 0),
                      v.counts.get(sormod.AMBIGUOUS, 0),
                      v.counts.get(sormod.MISSING, 0))))
    else:
        kv.append(("SOR verification",
                   '<span class="why">not verified: no SOR specification '
                   'covered this endpoint</span>'))

    body = ['<section id="%s"><h2><span class="verb">%s</span> '
            '<code class="api">%s</code></h2>'
            % (anchor, e(spec.method.upper()), e(spec.path)),
            '<p class="meta">%s &middot; %d elements declared in the workbook, '
            '<strong>%d published in this API</strong>, %d eliminated</p>'
            % (e(view.sheet), total - c.get(CONTAINER, 0), kept,
               total - c.get(CONTAINER, 0) - kept),
            '<dl class="kv">%s</dl>'
            % "".join("<dt>%s</dt><dd>%s</dd>" % (e(k), val) for k, val in kv)]

    if columns:
        body.append(_variant_legend(view, columns))
        body.append('<p class="small">In the numbered columns: '
                    '<span class="tick">&#10003;</span> that variant publishes '
                    'the element &middot; '
                    '<span class="xsor">&#10007;</span> the workbook named an '
                    'SOR field that does not exist on that endpoint &middot; '
                    '<span class="cross">&middot;</span> the SOR cell is '
                    'empty. Hover a mark for the detail.</p>')

    for m in view.messages:
        body.append("<h3>%s body</h3>" % e(m.label))
        for cls in m.classes:
            body.append(_class_table(cls, view, columns))

    findings = [f for f in spec.findings if f.severity == "warning"]
    if findings:
        body.append("<details><summary>%d finding%s on this sheet</summary>"
                    % (len(findings), "" if len(findings) == 1 else "s"))
        body.append('<div class="scroll"><table><thead><tr><th>Code</th>'
                    '<th>Cell</th><th>Found</th><th>Fix</th></tr></thead><tbody>')
        for f in findings:
            body.append("<tr><td class=\"mono\">%s</td><td class=\"mono\">%s</td>"
                        "<td>%s</td><td>%s</td></tr>"
                        % (e(f.code), e(f.cell) or "&#8212;", e(f.what),
                           e(f.fix)))
        body.append("</tbody></table></div></details>")
    body.append("</section>")
    return "".join(body)


REASON_ORDER = (NOT_IN_SOR, NO_FIELD)


def _eliminated_section(views):
    """Every element this application kept out of the published interface.

    The point of the tool is the elements that are *not* there, so they get a
    section of their own rather than only appearing greyed out among the
    published ones. Grouped by integration API, because that is the unit a
    consumer sees.
    """
    total = sum(len(v.eliminated) for v in views)
    narrowed = sum(len(v.narrowed) for v in views)
    if not total and not narrowed:
        return ('<section id="eliminated"><h2>Elements eliminated from the '
                'interface</h2><p class="meta">Nothing was eliminated: every '
                'element the workbook declared is backed by an SOR field and '
                'reaches the interface.</p></section>')

    by_reason = defaultdict(int)
    for v in views:
        for _msg, _cls, row in v.eliminated:
            by_reason[row.disposition] += 1

    cards = "".join(
        '<div class="card %s"><div class="n">%d</div><div class="k">%s</div>'
        '</div>' % (kind, n, e(label)) for kind, n, label in (
            ("bad", total, "elements eliminated from the interface"),
            ("no", by_reason.get(NO_FIELD, 0),
             "because the workbook names no SOR field"),
            ("bad", by_reason.get(NOT_IN_SOR, 0),
             "because the named field is not in the SOR"),
            ("var", narrowed, "narrowed to some variants only"),
        ))

    blocks = []
    for v in views:
        rows = v.eliminated
        if not rows:
            continue
        body = []
        for msg, cls, row in sorted(
                rows, key=lambda t: (REASON_ORDER.index(t[2].disposition), t[0])):
            declared = " | ".join(sorted({x for x in row.declared.values() if x}))
            body.append(
                '<tr class="d-%s"><td>%s</td><td><code>%s</code></td>'
                '<td class="name"><code>%s</code></td><td class="mono">%s</td>'
                '<td>%s</td><td class="mono">%s</td><td class="why">%s</td></tr>'
                % (row.disposition, e(msg), e(cls.name), e(row.node.name),
                   e(".".join(row.path)), _pill(row.disposition),
                   e(declared) or '<span class="why">empty</span>',
                   _why(row, v)))
        blocks.append(
            '<h3><span class="verb">%s</span> <code class="api">%s</code>'
            '<span class="kind"> &middot; %d eliminated</span></h3>'
            '<div class="scroll"><table class="elim"><thead><tr>'
            '<th class="e-ms">Message</th><th class="e-cl">Class</th>'
            '<th class="e-el">Element</th><th class="e-pa">Path in the message</th>'
            '<th class="e-re">Reason</th><th class="e-so">SOR field named</th>'
            '<th class="e-wh">Detail</th></tr></thead><tbody>%s</tbody>'
            '</table></div>'
            % (e(v.spec.method.upper()), e(v.spec.path), len(rows),
               "".join(body)))

    narrow_blocks = []
    for v in views:
        rows = v.narrowed
        if not rows:
            continue
        body = []
        for msg, cls, row in rows:
            yes = [ep for ep, (st, _d) in row.per_variant.items() if st == "yes"]
            body.append(
                '<tr><td>%s</td><td><code>%s</code></td>'
                '<td class="name"><code>%s</code></td><td>%s</td>'
                '<td class="why">%s</td></tr>'
                % (e(msg), e(cls.name), e(row.node.name),
                   e("%d of %d variants" % (len(yes), len(row.per_variant))),
                   e(", ".join(x or "(unnamed)" for x in yes))))
        narrow_blocks.append(
            '<h3><span class="verb">%s</span> <code class="api">%s</code>'
            '<span class="kind"> &middot; %d narrowed</span></h3>'
            '<div class="scroll"><table><thead><tr><th>Message</th>'
            '<th>Class</th><th>Element</th><th>Published by</th>'
            '<th>Which variants</th></tr></thead><tbody>%s</tbody></table>'
            '</div>' % (e(v.spec.method.upper()), e(v.spec.path), len(rows),
                        "".join(body)))

    out_html = ['<section id="eliminated">'
                '<h2>Elements eliminated from the interface</h2>'
                '<p class="meta">These are the elements the workbook declared '
                'that this application kept out of the published integration '
                'APIs. They are the reason the tool exists: every one of them '
                'is a field a consumer would otherwise have seen, and could '
                'have populated, with nothing behind it in the System of '
                'Record.</p>',
                '<div class="cards">%s</div>' % cards]
    if blocks:
        out_html.append("".join(blocks))
    if narrow_blocks:
        out_html.append(
            '<h3 class="sep">Narrowed rather than removed</h3>'
            '<p class="small">These elements are backed by the SOR on some '
            'variants of their operation and not others. Rather than showing '
            'them to every caller, each variant publishes only its own, '
            'through <code>allOf</code> inheritance from a shared base.</p>')
        out_html.append("".join(narrow_blocks))
    out_html.append("</section>")
    return "".join(out_html)


def _families_section(out):
    m = out.merged
    if m is None or m.document is None:
        return ""
    split = [(name, base, derived) for name, base, derived in m.profiles
             if len(derived) > 1]
    schemas = m.document["components"]["schemas"]
    rows = []
    for name, base, derived in sorted(split):
        for d in sorted(derived):
            entry = schemas.get(d) or {}
            used = entry.get("x-used-by") or []
            delta = entry.get("allOf", [{}, {}])[1] if "allOf" in entry else entry
            props = sorted((delta.get("properties") or {}))
            rows.append(
                "<tr><td><code>%s</code></td><td><code>%s</code></td>"
                "<td><code>%s</code></td><td>%s</td>"
                "<td class=\"mono\">%s</td></tr>"
                % (e(name), e(base) if base else '<span class="why">none</span>',
                   e(d), e(", ".join(used)) if used else
                   '<span class="why">one operation</span>',
                   e(", ".join(props)) if props else
                   '<span class="why">nothing, it is the base alone</span>'))
    if not rows:
        return ""
    base_counts = {}
    for name, base, _d in split:
        if base and base in schemas:
            base_counts[base] = len(schemas[base].get("properties") or {})
    return (
        '<section id="families"><h2>Classes shared by more than one endpoint</h2>'
        '<p class="meta">%d classes are used by several operations and do not '
        'agree on what they publish. Each keeps its own name for the shared '
        'core, and every operation that uses it gets a derived class '
        'inheriting through <code>allOf</code>. That is what stops one '
        'endpoint being shown another endpoint\'s elements.</p>'
        '<div class="scroll"><table class="fam"><thead><tr>'
        '<th class="f-cl">Class</th><th class="f-ba">Shared base</th>'
        '<th class="f-de">Derived class</th><th class="f-us">Used by</th>'
        '<th class="f-ad">Elements it adds</th></tr></thead>'
        '<tbody>%s</tbody></table></div></section>'
        % (len(split), "".join(rows)))


def _totals(views):
    c = defaultdict(int)
    for v in views:
        for k, n in v.counts.items():
            c[k] += n
    return c


def render(out, *, title=None):
    """The showcase, as one self-contained HTML document."""
    views = build_views(out)
    c = _totals(views)
    declared = sum(c.values()) - c.get(CONTAINER, 0)
    kept = c.get(PUBLISHED, 0) + c.get(VARIANT_ONLY, 0)
    removed = declared - kept
    verified = out.sor_verified

    nav = ['<nav><h2>Overview</h2>',
           '<a href="#summary">Summary</a>',
           '<a href="#method">How this was decided</a>']
    if out.merged is not None and out.merged.document is not None:
        nav.append('<a href="#families">Shared classes</a>')
    nav.append('<a href="#eliminated">Elements eliminated</a>')
    nav.append("<h2>Integration APIs</h2>")
    sections = []
    for i, v in enumerate(views, 1):
        anchor = "op%d" % i
        vc = v.counts
        nav.append('<a href="#%s"><span class="navverb">%s</span> %s'
                   '<span class="badge">%d/%d</span></a>'
                   % (anchor, e(v.spec.method.upper()), e(v.spec.path),
                      vc.get(PUBLISHED, 0) + vc.get(VARIANT_ONLY, 0),
                      sum(vc.values()) - vc.get(CONTAINER, 0)))
        sections.append(_endpoint_section(v, anchor))
    failed = [s for s in out.specs if s.status == "failed"]
    if failed:
        nav.append("<h2>Not generated</h2>")
        for s in failed:
            nav.append('<a class="fail" href="#failed">%s</a>' % e(s.sheet))
    nav.append("</nav>")

    if verified:
        s = out.sor_index.summary()
        gate = (
            '<div class="callout"><strong>Every mapping was checked against '
            'the System of Record specification.</strong>'
            '%d SOR file%s supplied, covering %d operations across %d paths. '
            'An element is published only where the workbook names an SOR '
            'field <em>and</em> that field exists in the relevant message of '
            'the SOR endpoint the sheet points at. Matching is on the last '
            'segment of the field name.</div>'
            % (s["files"], "" if s["files"] == 1 else "s", s["operations"],
               s["paths"]))
    else:
        gate = (
            '<div class="callout warn"><strong>No System of Record '
            'specification was supplied, so no mapping was verified.</strong>'
            'Every SOR field name in the workbook has been taken on trust. An '
            'element is published wherever the workbook names a field, whether '
            'or not that field exists in the SOR. Supply the SOR '
            'specification to turn this report into evidence.</div>')

    lost = c.get(LOST, 0)
    if lost:
        gate += ('<div class="callout bad"><strong>%d element%s mapped, '
                 'verified and yet absent from the generated interface.</strong>'
                 'That is a fault in the tool. The rows are marked LOST '
                 'below.</div>' % (lost, " is" if lost == 1 else "s are"))

    cards = [
        ("", declared, "elements declared in the workbook"),
        ("ok", kept, "published in the interface"),
        ("bad", removed, "removed, %d%% of what was declared"
         % (round(100 * removed / declared) if declared else 0)),
        ("no", c.get(NO_FIELD, 0), "removed: no SOR field in the workbook"),
        ("bad", c.get(NOT_IN_SOR, 0), "removed: named field absent from the SOR"),
        ("var", c.get(VARIANT_ONLY, 0), "published by some variants only"),
    ]

    merged_bits = ""
    if out.merged is not None and out.merged.document is not None:
        ms = out.merged.stats
        merged_bits = (
            "<p>The generated specification holds <strong>%d operations</strong> "
            "and <strong>%d classes</strong>. %d groups were published as named "
            "classes so they can be reused, and <strong>%d of those are used by "
            "more than one endpoint with different content</strong>, so each is "
            "split into a shared base and one derived class per endpoint. "
            "Without that split every endpoint would be shown the union of what "
            "all of them use.</p>"
            % (ms["operations"], ms["schemas"], ms["hoisted_groups"],
               ms["split_groups"]))

    legend = "".join(
        '<div>%s %s</div>' % (_pill(k), e(DISPOSITION_NOTE[k]))
        for k in (PUBLISHED, VARIANT_ONLY, NOT_IN_SOR, NO_FIELD, CONTAINER))

    failed_section = ""
    if failed:
        rows = []
        for s in failed:
            for f in s.findings:
                if f.severity != "error":
                    continue
                rows.append("<tr><td>%s</td><td class=\"mono\">%s</td>"
                            "<td class=\"mono\">%s</td><td>%s</td><td>%s</td></tr>"
                            % (e(s.sheet), e(f.code), e(f.cell) or "&#8212;",
                               e(f.what), e(f.fix)))
        failed_section = (
            '<section id="failed"><h2>Endpoints not generated</h2>'
            '<p class="meta">No interface was produced for these sheets, so '
            'nothing about them appears above. Each needs a correction in the '
            'workbook.</p><div class="scroll"><table><thead><tr><th>Sheet</th>'
            '<th>Code</th><th>Cell</th><th>Found</th><th>Fix</th></tr></thead>'
            '<tbody>%s</tbody></table></div></section>' % "".join(rows))

    doc = """<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%(title)s</title><style>%(css)s</style></head><body>
<div class="wrap">
%(nav)s
<main>
<header class="top">
<h1>%(title)s</h1>
<p class="sub">What the integration interface publishes, what it does not, and
why. Generated from <code>%(workbook)s</code> by SOR Polymorphizer %(version)s
on %(when)s.</p>
</header>
%(gate)s
<section id="summary"><h2>Summary</h2>
<p class="meta">The purpose of this tool is to keep elements the System of
Record cannot supply out of the integration interface. These are the numbers
that show whether it did.</p>
<div class="cards">%(cards)s</div>
%(merged)s
<h3>What each disposition means</h3>
<div class="legend">%(legend)s</div>
</section>
<section id="method"><h2>How this was decided</h2>
<p>For every row of every operation sheet the tool asked, in order:</p>
<ol>
<li><strong>Does the workbook name an SOR field for this element?</strong> An
empty SOR cell is the analyst saying the System of Record does not supply it,
and the element is removed.</li>
<li><strong>Does that field exist in the System of Record?</strong> The
endpoint comes from the sheet's <code>SOR API Endpoint</code> banner cell,
which is the definitive statement of it. A Request Body row is looked up in
that endpoint's request, meaning its body for POST and PUT or its parameters
for GET, and a Response Body row in its response. Matching is on the last
segment of the field name, so <code>accountInfo.currency</code> resolves to any
element named <code>currency</code>. A field that is not there means the
element is removed, and the removal is listed against its cell.</li>
<li><strong>Do all variants of the operation supply it?</strong> Where an
operation is served by several SOR endpoints, an element supplied by some and
not others appears only in those, through <code>allOf</code> inheritance from a
shared base.</li>
<li><strong>Do other endpoints share this class?</strong> A class used by more
than one endpoint with different content is split, so no endpoint is shown
another endpoint's elements.</li>
</ol>
<p class="small">Only the <code>SOR API Endpoint</code> banner cell decides
which SOR endpoint applies. An endpoint written on an SOR column header is a
label: where the two differ the banner is used and the difference is reported,
so the workbook can be tidied without the interface changing underneath
anyone.</p>
</section>
%(eliminated)s
%(families)s
%(sections)s
%(failed)s
</main></div>
<footer>SOR Polymorphizer %(version)s &middot; %(when)s &middot; %(mode)s</footer>
</body></html>
"""
    return doc % {
        "title": e(title or "Integration interface, element by element"),
        "css": CSS,
        "nav": "".join(nav),
        "workbook": e(os.path.basename(out.workbook)),
        "version": e(__version__),
        "when": e(_dt.datetime.now().strftime("%d %B %Y at %H:%M")),
        "gate": gate,
        "cards": "".join(
            '<div class="card %s"><div class="n">%s</div><div class="k">%s</div>'
            '</div>' % (kind, n, e(label)) for kind, n, label in cards),
        "merged": merged_bits,
        "legend": legend,
        "eliminated": _eliminated_section(views),
        "families": _families_section(out),
        "sections": "".join(sections),
        "failed": failed_section,
        "mode": e("verified against the SOR specification" if verified
                  else "NOT verified: no SOR specification supplied"),
    }


def default_filename(out, fallback="openapi"):
    """The showcase file name, taken from the first service domain.

    The report is read by service domain, so it is named by service domain.
    The first one encountered in sheet order wins, because a workbook that
    spans more than one is still one interface and needs one name.
    """
    for spec in out.specs:
        if spec.status != "ok":
            continue
        result = next((r for r in _results_of(out) if r.sheet == spec.sheet),
                      None)
        domain = text((result.banner or {}).get("service_domain", "")) \
            if result is not None else ""
        if domain:
            return (gen.slug(domain) or fallback) + ".html"
    return "%s_showcase.html" % fallback


def write(out, path, *, title=None):
    import polymorphize_core as core
    core._write(path, render(out, title=title))
    return path


__all__ = ["ClassView", "EndpointView", "MessageView", "Row", "build_views",
           "render", "write", "disposition_of", "published_leaf_paths",
           "DISPOSITION_LABEL", "DISPOSITION_NOTE",
           "PUBLISHED", "NO_FIELD", "NOT_IN_SOR",
           "VARIANT_ONLY", "CONTAINER", "LOST"]
