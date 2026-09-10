# Archive: output from version 5.3

These three files were produced by version 5.3 from a workbook that is not
part of the sample set. They are kept for reference only and they do not
describe how 6.5 behaves. Specifically they predate three changes:

* the `__` class and context naming convention, so their schema names carry
  neither `x-class` nor `x-context`
* the plain name being given to the most used shape rather than to the shared
  core
* the Apigee error set, which 6.5 attaches to every operation

Nothing in the tool reads them, and no test refers to them. Regenerate from a
current workbook rather than treating these as a baseline.
