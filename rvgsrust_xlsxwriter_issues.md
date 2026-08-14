# Proposed issues — rvgsrust-xlsxwriter 0.2.1

All of these came out of a real migration evaluation, not a synthetic test. Shared
environment block for any issue you file:

```
rvgsrust-xlsxwriter 0.2.1 (wheel from PyPI)
Python 3.11.14, Windows (conda env), polars 1.41.2
Workload: 1,261 rows × 76,480 columns SPSS .sav, exported as 2 workbooks
          × 7 sheets each (5 grid sheets of ≤16,000 cols + 2 metadata sheets)
Comparison baseline: rustpy-xlsxwriter 0.5.2 (same rust_xlsxwriter core)
```

Priority order below. Only **#1** is a defect; the rest are gaps or polish.

---

## 1. `Workbook.close()` requires the path again even when it was given to the constructor

**Labels:** bug, api

### What happens

```python
wb = Workbook("out.xlsx")
ws = wb.add_worksheet("Sheet1", False)
ws.write_dataframe(0, 0, df, header_format=fmt, write_header=True)
wb.close()          # TypeError
wb.close("out.xlsx")  # works
```

`Workbook` exposes a `path` attribute (visible in `dir(wb)`), and the constructor
already accepts the path, so `close()` has everything it needs. Requiring the
caller to pass it a second time means the path has to be threaded through
application code purely to satisfy `close()`, and it creates a silent
opportunity to write to the wrong file if the two ever disagree.

### Expected

`close()` with no argument closes to the constructor path. `close(path)` stays
supported as an override, so this is backwards compatible.

### Also worth confirming

If `with Workbook(path) as wb:` exits via `close()` with no args, it likely hits
the same `TypeError`. I have not tested the context manager, so please verify
that behavior before assuming it's unaffected.

---

## 2. `write_dataframe()` has no `column_formats` parameter, so per-column formats can't ride the fast path

**Labels:** enhancement, api

### Current signature

```
Worksheet.write_dataframe(start_row, start_col, data, header_format, write_header)
```

There is no per-column format argument. The README lists per-column formats in
the v0.3 roadmap, so this is expected — filing it to record the use case and
to note what the workarounds actually cost.

### Use case

Applying cell borders to a data grid. The two available routes both have
drawbacks:

- `set_column_range_format(0, n-1, fmt)` — cheap (measured ~+0.3s, about 5% on
  the workload above) but **column-scoped**. In OOXML a cell's own format takes
  precedence over the column format, so any cell written with its own format
  (dates, number formats) can end up with no border. Whether that bites depends
  on what `write_dataframe` emits per dtype.
- `set_range_format(first_row, first_col, last_row, last_col, fmt)` — correct,
  but on a 25,000 × 76,000 grid that's ~1.9 billion cells. Not viable.

A `column_formats` argument that merges the format into the cells as they are
written from Arrow would avoid both problems. For reference,
`rustpy-xlsxwriter` 0.5.2 exposes exactly this on its
`.sheet(name, data, column_width, column_widths, column_formats, header_format)`
builder, so the shape is proven at the binding level.

### Suggested

`write_dataframe(..., column_formats: dict[str, Format] | None = None)`, keyed by
column name. If the implementation must fall back to a per-cell loop, state
that in the docstring so callers can measure before adopting it.

---

## 3. No `set_column_range_width()` to match `set_column_range_format()`

**Labels:** enhancement, performance

### Problem

`Worksheet` has `set_column_range_format(first_col, last_col, format)` but the
width API is per-column only: `set_column_width(col, width)`. Setting a uniform
width across a wide sheet therefore costs one Python→Rust call per column.

### Measured

On the workload above (14 sheets, ~153,000 columns total, uniform width 15):

| | build phase | save phase | total |
|---|---:|---:|---:|
| rvgsrust-xlsxwriter 0.2.1 | 2.7s | 2.6s | 5.7s |
| rustpy-xlsxwriter 0.5.2 (`column_width=15` kwarg) | 0.0s | 6.5s | 6.5s |

The 2.7s of eager build time is dominated by the width loop; the comparison
library takes a single `column_width=` argument and does it inside Rust. Net
throughput is still slightly better here, but this is avoidable overhead that
grows linearly with column count.

### Suggested

`set_column_range_width(first_col, last_col, width)`, mirroring the existing
range-format method. A `default_column_width` on the worksheet or workbook
would also solve the uniform case.

---

## 4. `Format` border setters are named inconsistently

**Labels:** api, breaking-change-candidate

Style setters and colour setters use opposite word orders:

| edge | style setter | colour setter |
|---|---|---|
| bottom | `set_bottom_border` | `set_border_bottom_color` |
| top | `set_top_border` | `set_border_top_color` |
| left | `set_left_border` | `set_border_left_color` |
| right | `set_right_border` | `set_border_right_color` |
| diagonal | `set_border_diagonal` | `set_border_diagonal_color` |

Diagonal already uses the `set_border_*` prefix, so the four main edges are the
odd ones out — and autocomplete on `set_border_` silently hides half the API.
`rustpy-xlsxwriter` 0.5.2 uses `set_border_bottom` / `set_border_bottom_color`
consistently.

**Suggested:** add `set_border_top/bottom/left/right` as the canonical names,
keep the current spellings as deprecated aliases until the next major version.

---

## 5. No docstrings or type stubs; `set_border()`'s accepted values are undocumented

**Labels:** documentation

Introspecting the wheel shows parameter names but no docstrings:

```
Format.set_border sig=['border'] doc=[]
Worksheet.write_dataframe sig=['start_row','start_col','data','header_format','write_header'] doc=[]
```

`border` has no documented type or accepted value set. I found `"thin"` works
only by trying candidates (`"thin"`, `"Thin"`, `"THIN"`, `1`, `7`) until one
stopped raising. Callers integrating this into production code shouldn't have to
brute-force an enum.

**Suggested:** ship a `.pyi` stub with literal types for the string enums
(`border`, `align`, `pattern`, `underline`, …), or at minimum add short
docstrings listing the accepted values and what happens on an invalid one.

---

## 6. `annotations` leaks into the public namespace

**Labels:** good first issue

```python
>>> [n for n in dir(rvgsrust_xlsxwriter) if not n.startswith("_")]
['Chart', ..., 'Workbook', 'Worksheet', 'annotations']
```

`annotations` is the object from `from __future__ import annotations` leaking
through `__init__.py`. Harmless, but it appears in `dir()`, in autocomplete,
and in `from ... import *`.

**Suggested:** define `__all__` in `__init__.py`.

---

## 7. Pin or commit dependencies for the audit release

**Labels:** build, supply-chain

Worth confirming a `Cargo.lock` is committed and that transitive dependency
versions are reproducible for a tagged release. Files produced by this library
go to external clients in my case, and "the wheel I tested" and "the wheel I
ship" need to be the same build. If this is already handled on `main` and just
wasn't in the 0.2.1 sdist, ignore this.

**Suggested:** commit `Cargo.lock`, and note the MSRV and the resolved
`rust_xlsxwriter` version in the release notes so downstream users can pin
confidently.

---

## 8. README's "6–8× faster" needs the benchmark shape stated alongside it

**Labels:** documentation

The headline figure is measured against pure-Python `xlsxwriter` at
100,000 rows × 8 columns. That's a reasonable benchmark, but it reads as a
general claim, and the two comparisons people actually make are different:

- **vs. another Rust-backed writer** (`rustpy-xlsxwriter`, same
  `rust_xlsxwriter` core, also zero-copy Arrow): I measured **~12% faster**
  overall (5.7s vs 6.5s), not 6–8×.
- **wide rather than tall**: 76,480 columns × 1,261 rows behaves very
  differently from 8 columns × 100,000 rows.

Where 0.2.1 *did* win decisively in my testing was GIL behaviour — a background
Python thread polling every 5 ms saw a worst-case stall of **~20 ms** during
`close()`, versus **~1,300 ms** for the comparison library. In practice, that
means a GUI stays responsive during the write instead of freezing — a much
stronger selling point for desktop-app users than the raw throughput number,
and it isn't called out prominently.

**Suggested:** state the benchmark shape and the comparison target inline with
the figure, add a wide-workbook row, and promote the "save releases the GIL"
property to the feature list.

---

## Not filed — measured but inconclusive or untested

Listing these so nothing looks like it was verified when it wasn't:

- **`constant_memory=True`**: not yet exercised on this workload. The README's
  non-decreasing-row-order requirement (raises `ValueError` where upstream would
  emit a corrupt file) is a good behaviour to have, but I haven't tested it, so
  I'm not filing anything about it.
- **Border coverage per column**: my first check sampled only 50 cells of one
  row and reported "confirmed", which was not a sound basis for the claim in
  issue #2. A full per-cell scan is what should decide whether
  `set_column_range_format` actually covers date/number-formatted columns. Until
  that run is done, treat the precedence concern in #2 as a hypothesis about
  OOXML semantics rather than an observed failure in this library.
- **Output size**: 16.8 MB vs 18.3 MB for the same data. Smaller is fine and
  needs no issue, but I haven't checked whether the difference is shared-string
  handling, style de-duplication, or compression level.
