"""
polymorphize_errors — the Apigee error set, and the standard headers.

Version 6.8.

Every operation the tool publishes carries the same twelve error responses,
referenced from ``components/responses`` rather than inlined. The set is not a
judgement call: it is taken from the API COE's own specification,
``apicoeissueddeviceadministrationswaggerv1.0.8.yaml``, where all twelve carry
this sentence in their description, so the specification itself states which
layer raises them.

    Disclaimer: The error message is Apigee error only and not mapped to the
    System API error. Consumers should not rely on the exact error text as it
    may change depends on the format of the System API its calling.

In that specification all twelve are structurally identical: the same three
response headers, the same ``ErrorResponse`` schema, the same disclaimer, and
every one of the 264 error entries across its 22 operations is a ``$ref`` with
nothing inlined. That is the shape reproduced here.

Two departures from the sample, both deliberate
===============================================

The three ``x-BDO-Client-Request-*`` headers appear on the sample's ``200``
responses as well as its errors, so they are a house convention rather than an
error convention. They are attached to **every** response this tool emits.

The sample's ``ErrorResponse`` example does not conform to its own schema: the
schema declares ``errors`` as an array and the example provides ``error`` as a
single object, it omits ``status`` and ``realm``, and its ``title`` reads "The
request was successful, but there is no content in the response", which is a
204 message in an error example. It is not copied. The example below conforms,
and the defect has been raised as a note for the API COE rather than silently
propagated into every specification the tool generates.
"""

from __future__ import annotations

from collections import OrderedDict

__version__ = "6.8"

#: The disclaimer, quoted from the source specification, wording unchanged.
DISCLAIMER = (
    "Disclaimer: The error message is Apigee error only and not mapped to the "
    "System API error. Consumers should not rely on the exact error text as it "
    "may change depends on the format of the System API its calling."
)

#: ``(component name, status code, reason phrase)``, in status order.
APIGEE_ERRORS = (
    ("BadRequest",          "400", "Bad Request"),
    ("Unauthorized",        "401", "Unauthorized"),
    ("Forbidden",           "403", "Forbidden"),
    ("NotFound",            "404", "Not Found"),
    ("MethodNotAllowed",    "405", "Method Not Allowed"),
    ("Conflict",            "409", "Conflict"),
    ("UnprocessableEntity", "422", "Unprocessable Entity"),
    ("TooManyRequests",     "429", "Too Many Requests"),
    ("InternalServerError", "500", "Internal Server Error"),
    ("BadGateway",          "502", "Bad Gateway"),
    ("ServiceUnavailable",  "503", "Service Unavailable"),
    ("GatewayTimeout",      "504", "Gateway Timeout"),
)

ERROR_SCHEMA_NAME = "ErrorResponse"

#: The three standard headers, on every response rather than only the errors.
STANDARD_HEADERS = OrderedDict((
    ("x-BDO-Client-Request-Id",
     "The unique id (preferable uuidv4) provided by the client for a specific "
     "HTTP request."),
    ("x-BDO-Client-Request-Trace-Id",
     "The unique identifier for an entire trace in OpenTelemetry (OTEL). A "
     "trace represents a single request or transaction as it flows through "
     "multiple services in a distributed system."),
    ("x-BDO-Client-Request-Span-Id",
     "Uniquely identifies a single span within a trace. A span represents one "
     "unit of work, for example an HTTP request or a database query."),
))


def headers_block():
    """``components/headers``, the three standard response headers."""
    return OrderedDict(
        (name, OrderedDict([("schema", {"type": "string"}),
                            ("description", desc)]))
        for name, desc in STANDARD_HEADERS.items())


def header_refs():
    """The ``headers`` mapping to attach to any response."""
    return OrderedDict(
        (name, {"$ref": "#/components/headers/%s" % name})
        for name in STANDARD_HEADERS)


def error_schema():
    """``ErrorResponse``, with an example that conforms to it."""
    return OrderedDict([
        ("description",
         "An array of detail error codes, and messages, and URLs to "
         "documentation to help remediation."),
        ("type", "object"),
        ("properties", OrderedDict([
            ("status", {"type": "string",
                        "description": "The HTTP status code of the response"}),
            ("title", {"type": "string",
                       "description": "A short, human-readable summary of the "
                                      "error."}),
            ("timestamp", {"type": "string", "format": "date-time",
                           "description": "A standardized date and time format "
                                          "defined by the ISO 8601 standard"}),
            ("errors", OrderedDict([
                ("type", "array"),
                ("description", "Array of errors within an object to provide "
                                "additional information."),
                ("items", OrderedDict([
                    ("type", "object"),
                    ("properties", OrderedDict([
                        ("realm", {"type": "string",
                                   "description": "Indicates the architectural "
                                                  "layer or system component "
                                                  "where the error originated."}),
                        ("code", {"type": "string",
                                  "description": "Represents a unique "
                                                 "identifier for the specific "
                                                 "error condition."}),
                        ("errordesc", {"type": "string",
                                       "description": "Provides a short, "
                                                      "human-readable "
                                                      "description of the error "
                                                      "associated with the code."}),
                        ("detail", {"type": "string",
                                    "description": "A more comprehensive "
                                                   "explanation of the error, "
                                                   "including what went wrong "
                                                   "and how to fix it."}),
                        ("instance", {"type": "string",
                                      "description": "A URI that identifies the "
                                                     "specific occurrence of the "
                                                     "problem."}),
                    ])),
                ])),
            ])),
        ])),
        # Conforms to the schema above: errors is an array, status and realm
        # are present, and the title describes a failure. See the module
        # docstring for why the source example is not reproduced.
        ("example", OrderedDict([
            ("status", "400"),
            ("title", "Bad Request"),
            ("timestamp", "2026-07-14T08:16:49.876Z"),
            ("errors", [OrderedDict([
                ("realm", "Apigee"),
                ("code", "001"),
                ("errordesc", "QUERY_PARAMETER_FAILURE"),
                ("detail", "The input query parameters do not match the open "
                           "api spec"),
                ("instance", "/v1/<resource>/<action>"),
            ])]),
        ])),
    ])


def responses_block():
    """``components/responses``, the twelve Apigee errors."""
    out = OrderedDict()
    for name, _code, reason in APIGEE_ERRORS:
        out[name] = OrderedDict([
            ("description", "%s\n\n%s" % (reason, DISCLAIMER)),
            ("headers", header_refs()),
            ("content", {"application/json": {
                "schema": {"$ref": "#/components/schemas/%s" % ERROR_SCHEMA_NAME}}}),
        ])
    return out


def operation_responses():
    """``{status: {$ref}}`` for the twelve, to merge into an operation."""
    return OrderedDict(
        (code, {"$ref": "#/components/responses/%s" % name})
        for name, code, _reason in APIGEE_ERRORS)


def attach(document):
    """Add the error set to a finished document, in place, and return it.

    Idempotent, so a document that already carries the set is unchanged. The
    twelve are added to every operation that does not already declare that
    status, the three headers are added to every response, and the shared
    components are created once.
    """
    comp = document.setdefault("components", OrderedDict())

    schemas = comp.setdefault("schemas", OrderedDict())
    if ERROR_SCHEMA_NAME not in schemas:
        schemas[ERROR_SCHEMA_NAME] = error_schema()

    headers = comp.setdefault("headers", OrderedDict())
    for name, block in headers_block().items():
        headers.setdefault(name, block)

    responses = comp.setdefault("responses", OrderedDict())
    for name, block in responses_block().items():
        responses.setdefault(name, block)

    refs = operation_responses()
    for _path, entry in (document.get("paths") or {}).items():
        if not isinstance(entry, dict):
            continue
        for method, op in entry.items():
            if not isinstance(op, dict) or "responses" not in op:
                continue
            for code, ref in refs.items():
                op["responses"].setdefault(code, dict(ref))
            # The standard headers belong on the success responses too: the
            # source specification carries them on its 200s.
            for code, response in op["responses"].items():
                if "$ref" in response:
                    continue                     # the shared errors carry them
                response.setdefault("headers", header_refs())
    return document


def error_codes():
    """The twelve status codes, for a report or a test."""
    return [code for _n, code, _r in APIGEE_ERRORS]


__all__ = ["DISCLAIMER", "APIGEE_ERRORS", "ERROR_SCHEMA_NAME",
           "STANDARD_HEADERS", "headers_block", "header_refs", "error_schema",
           "responses_block", "operation_responses", "attach", "error_codes"]
