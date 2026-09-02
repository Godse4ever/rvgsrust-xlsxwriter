# API parity gaps vs rust_xlsxwriter 0.98.2

Only open items below. Anything closed has been removed from this file
rather than marked done, so this stays a to-do list, not a changelog —
see [CHANGELOG.md](CHANGELOG.md) for what shipped.

Worksheet and Workbook are both fully closed now (PRs #30-#38) --
Charts is the only section with open gaps. serde serialisation is the
one deliberate exception, left open below rather than closed, because
it's a real feature (new Cargo feature flag + a Python-dict-to-
serde_json::Value bridge) rather than a small gap, and hasn't been
attempted yet.

All `file:line` references below were read from source, re-checked
against
[`v0.98.2`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.98.2/src).

Priorities are a judgement call about request frequency, weighted toward
this project's actual use — survey and market-research reporting, where
page setup, print layout and per-column formatting come up far more often
than VBA or chartsheets. They are not upstream's opinion.

---

## Charts — ~40 gaps

Parts 1–3 covered the core (series, formatting, markers/trendlines/data
labels). Secondary axes (`set_x2_axis_*`/`set_y2_axis_*`, PR #28) and
error bars (`ChartErrorBars`, PR #29) have shipped. Up-down bars, drop
lines, and high-low lines (all Line-chart-only decorations) shipped in
PR #39. Chart/plot area formatting shipped in PR #40. `Chart.combine()`
for combined charts shipped in PR #41. Axis label placement, tick
marks, date min/max/units, crossing, and display units all shipped in
PR #42 -- note two OOXML quirks documented there: `set_x_axis_crossing()`
actually controls the *value* axis's `<c:crosses>` (upstream stores
crossing on the opposite axis from where it renders), and the same for
`set_x_axis_position_between_ticks()` and `<c:crossBetween>`. Legend
entry deletion, object movement, decorative, and scale width/height
shipped in PR #43. `ChartDataTable` (data tables) shipped in PR #44.
What remains:

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Per-point formatting | `chart.rs:7750` | `ChartPoint` pyclass, via `ChartSeries.set_points([...])` | Medium |
| Gradient fills | `chart.rs:13980` | `ChartFormat.set_gradient_fill(...)` plus `ChartGradientStop` | Medium |
| Pattern fills | `chart.rs:13915` | `ChartFormat.set_pattern_fill(...)` | Low |
| Gridline formatting | — | Not implementable -- upstream has no `major_gridlines()`/`minor_gridlines()` accessor returning a formattable object, only the on/off toggle already exposed | Medium |
| Manual layouts | `chart.rs:9147` | `ChartLayout` pyclass for `title`, `legend`, `plot_area` | Low |

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
