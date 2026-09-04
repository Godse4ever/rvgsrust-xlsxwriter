# API parity gaps vs rust_xlsxwriter 0.98.2

Only open items below. Anything closed has been removed from this file
rather than marked done, so this stays a to-do list, not a changelog —
see [CHANGELOG.md](CHANGELOG.md) for what shipped.

Worksheet, Workbook, and Charts are all fully closed now (PRs #30-#47).
serde serialisation is the one deliberate exception, left open below
rather than closed, because it's a real feature (new Cargo feature
flag + a Python-dict-to-serde_json::Value bridge) rather than a small
gap, and hasn't been attempted yet. Gridline formatting on `Chart` is
a separate, permanent exception -- not a to-do, since upstream has no
API to bind to (see Known limitations below).

All `file:line` references below were read from source, re-checked
against
[`v0.98.2`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.98.2/src).

---

## serde serialisation — deferred, not attempted

`Worksheet.serialize()`/`serialize_headers()` and friends are feature-gated
upstream behind Cargo's `serde` feature. `serde`/`serde_json` are already
present transitively in `Cargo.lock` (pulled in by other dependencies), so
enabling the feature and adding a direct `serde_json` dependency is lower-risk
than it looks at first. The real work is the bridge: upstream's methods are
generic over `T: Serialize`, which only works for actual Rust structs deriving
`Serialize` -- there's no such thing from Python. A binding would need to
accept a Python dict (or list of dicts) and convert it to `serde_json::Value`
(which does implement `Serialize`) before calling through, and would need to
be checked against upstream's serializer to see whether it tolerates a plain
JSON value the way it tolerates a derived struct (field-renaming/skip
attributes on a real struct have no JSON equivalent). Not a small gap --
left as a dedicated follow-up rather than rushed.

---

## Known limitations (not parity gaps)

**Chart gridline formatting:** not implementable from this binding --
upstream has no `major_gridlines()`/`minor_gridlines()` accessor
returning a formattable object on `Chart`, only the on/off toggle
already exposed (`set_x_axis_major_gridlines()` etc.). There is no
`ChartFormat`-compatible path to style gridline color/weight/dash type.

**`group_rows()` in `constant_memory=True` mode:** individual `<row>`
elements never get an `outlineLevel` attribute -- only the sheet-wide
`outlineLevelRow` maximum is written. `group_rows()` still succeeds and
doesn't raise, but the per-row visual grouping in Excel won't appear.
Appears to be an upstream `constant_memory` streaming limitation, not
fixable from this binding without writing worksheet XML directly.
`group_columns()` is unaffected (columns are a separate `<cols>`
section, outside the row-streaming mechanism).

**Suspected upstream bug:** `Worksheet::set_print_first_page_number`
(`worksheet.rs:18697`) writes the page number into the
`useFirstPageNumber` attribute and never emits a `firstPageNumber`
attribute:

```rust
if self.first_page_number > 0 {
    attributes.push(("useFirstPageNumber", self.first_page_number.to_string()));
}
```

Per ECMA-376, `pageSetup@useFirstPageNumber` is a boolean and
`pageSetup@firstPageNumber` is the uint carrying the value. Excel treats a
nonzero boolean as true, so the feature is enabled but the first page
number probably defaults to 1 rather than the requested value.

Our binding passes the call straight through, so it inherits the
behaviour. `tests/test_page_setup.py::test_print_first_page_number` pins
what is currently written rather than what ought to be, and will fail if
upstream changes it. Not worked around locally: doing so would mean
writing worksheet XML ourselves. Worth reporting upstream.
