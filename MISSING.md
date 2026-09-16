# API parity gaps vs rust_xlsxwriter 0.99.0

Only open items below. Anything closed has been removed from this file
rather than marked done, so this stays a to-do list, not a changelog —
see [CHANGELOG.md](CHANGELOG.md) for what shipped.

Worksheet, Workbook, and Charts are all fully closed now (PRs #30-#47).
serde serialisation was evaluated and deliberately declined -- see
below for why -- rather than closed as done. Gridline formatting on
`Chart` is a separate, permanent exception -- not a to-do, since
upstream has no API to bind to (see Known limitations below). With
both of those settled, there is currently nothing actionable left in
this file.

All `file:line` references below were read from source, re-checked
against
[`v0.98.2`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.98.2/src)
(the pin in place when the parity work was done; the crate has since
moved to 0.99.0 -- see CHANGELOG.md -- with no API changes affecting
this binding).

---

## serde serialisation — declined

`Worksheet.serialize()`/`serialize_headers()` and friends are feature-gated
upstream behind Cargo's `serde` feature. `serde`/`serde_json` are already
present transitively in `Cargo.lock` (pulled in by other dependencies), so
enabling the feature and adding a direct `serde_json` dependency would have
been lower-risk than it looks at first. The real work would have been the
bridge: upstream's methods are generic over `T: Serialize`, which only works
for actual Rust structs deriving `Serialize` -- there's no such thing from
Python. A binding would have needed to accept a Python dict (or list of
dicts) and convert it to `serde_json::Value` (which does implement
`Serialize`) before calling through, with no guarantee upstream's serializer
tolerates a plain JSON value the way it tolerates a derived struct
(field-renaming/skip attributes on a real struct have no JSON equivalent).
Evaluated and explicitly decided not worth the effort -- not on the
roadmap.

---

## Known limitations (not parity gaps)

**`set_row_height()` quantizes to the nearest 0.75pt (1/288") step:**
confirmed via source trace, not fixable from this binding. Upstream's
own `Worksheet::set_row_height()` converts the point value to pixels
and rounds to the nearest integer (`(height * 4.0 / 3.0).round() as
u32`), then stores *only* that integer pixel count -- there is no
fractional-point storage anywhere in `rust_xlsxwriter`'s row metadata,
even internally. Converting back to points for the XML `ht` attribute
(`pixels as f64 * 0.75`) reproduces the original value exactly only
when the input is a multiple of 3; otherwise it's off by exactly
0.25pt in a fixed direction per residue class (`height % 3 == 1` reads
back 0.25pt low, `height % 3 == 2` reads back 0.25pt high). E.g. input
`20` round-trips as `20.25`, input `8` as `8.25`, input `7` as `6.75`.
`set_row_height_pixels()` doesn't avoid this either -- it's the same
underlying `u32` storage, just skipping the point-to-pixel conversion
step. Reported and confirmed via a real user's byte-level comparison
against classic `xlsxwriter` (which stores/writes the exact point
value with no unit conversion). Not filed upstream with
`jmcnamara/rust_xlsxwriter` yet.

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
