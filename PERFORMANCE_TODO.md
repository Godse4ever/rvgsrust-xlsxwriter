# Performance TODO

Log only. **Nothing here is to be implemented in the same session as
feature work** — these change hot paths and need their own
before/after benchmark runs to be worth anything.

Baseline numbers live in [PERFORMANCE.md](PERFORMANCE.md). All timings
below were measured on an Intel i5-6267U (MacBook Pro); anything marked
*unmeasured* is a hypothesis, not a result.

---

## Measured bottlenecks

### 1. `RefCell<RustWorkbook>` borrow cost — ~5–10%

Every `Workbook` and `Worksheet` method pays a `borrow_mut()` refcount
check. With ~800k calls for a 100k × 8 write, that adds up.

`UnsafeCell` would eliminate it, since the GIL already serialises access
to the workbook — no two Python threads can be inside a `pymethod` for
the same object at once.

- **Effort:** medium. Mechanical, but every `with_sheet` / `borrow_mut`
  site has to be audited to confirm no re-entrancy.
- **Risk:** high. Getting this wrong is UB rather than a wrong answer.
  Worth writing down explicitly *why* the GIL makes it sound, in a
  comment at the `UnsafeCell` definition, before touching it.
- **Caveat:** if `Py::allow_threads` is ever introduced to release the
  GIL during a long write, this reasoning collapses. Note it there too.

### 2. `Py<Workbook>` strong reference in `Worksheet`

Each `Worksheet` holds a strong ref to its `Workbook`, so the pair forms
a reference cycle that only Python's cyclic GC can collect. In
long-running processes that delays reclamation of the whole workbook,
including its string table.

Converting to `PyWeakref` breaks the cycle. Needs a clear error when the
workbook has already been dropped — currently that case cannot arise.

- **Effort:** low–medium. **Risk:** low, but it is an API-visible change
  in failure behaviour, so it wants a test.

### 3. `write_rows` is slower than `write_records` at 100k

2.02s vs 1.94s — the opposite of the intent, since `write_rows` was
meant to be the faster pre-classified path. The
`Vec<(Vec<CellValue>, bool)>` allocation dominates: one inner `Vec` per
row, 100k allocations.

Options, cheapest first:

- `SmallVec<[CellValue; 16]>` for the inner row — most sheets are under
  16 columns, so this removes the per-row heap allocation entirely.
- A reusable row buffer, cleared per row rather than reallocated.
- A pool allocator, if the above two are not enough.

- **Effort:** low for `SmallVec` (new dependency). **Risk:** low.
- Start here. It is the clearest win-to-effort ratio on this list.

### 4. Polars beats PyArrow at 100k in `write_dataframe`

Same code path, so the difference is cache locality in the source
buffers rather than anything in this crate. Suggests a **column-major
write ordering** would help both: walk one column to completion before
moving to the next, instead of row-major.

- **Blocker:** `constant_memory` mode requires monotonically
  non-decreasing row order, so column-major is incompatible with it.
  Would need either a mode check or two code paths.
- **Effort:** high. **Risk:** medium — the row-order guard exists for a
  reason and has tests.

---

## Not yet tried

Unmeasured. Listed roughly best-guess-first.

- **`target-cpu=native` for source installs.** Claimed 5–8% on modern
  hardware. Cannot go in released wheels — they must stay portable — so
  this is an opt-in profile for people building from source.
- **Direct `RecordBatch` iteration in `write_dataframe`**, skipping the
  `ArrowColumn` enum dispatch on the hot path. The enum now has 20
  variants after the extended-types work, so the match is wider than
  when it was written. Worth measuring whether the branch predictor
  still handles it.
- **Deduplicated string writes.** Repeated categorical values currently
  allocate a fresh `String` per cell in `arrow_cell_value`. A
  `HashMap<&str, ExcelString>` keyed off the Arrow buffer would make
  repeats near-free. This is the single biggest theoretical win for
  survey-style data, where a column may have five distinct values across
  100k rows.
- **Multi-threaded row classification in `write_records`** via rayon.
  Note the GIL: classification touches Python objects, so the parallel
  section would have to be strictly after extraction into `CellValue`.
  That may leave too little work to be worth the coordination.
- **`constant_memory=True` as the default.** Breaking change — defer to
  v0.3 and pair with a migration note.

---

## Observations from the feature work

Found while implementing Arrow types, conditional formats, sparklines and
charts. None are on a hot path, but they are worth knowing before
anyone benchmarks and gets confused.

- **`write_dataframe` builds two `Format`s unconditionally.** The
  `yyyy-mm-dd` and `yyyy-mm-dd hh:mm:ss` formats for temporal columns
  are created once per call even when the frame has no date or timestamp
  column. Two wasted allocations per call — trivial in absolute terms,
  but it is pure waste and a lazy construction is a two-line change.
  This is the cheapest item on the whole page.
- **Per-column format resolution is already hoisted.** `col_fmts` is
  computed once per `RecordBatch`, not per cell, so the temporal support
  added no per-cell cost for non-temporal data. Do not "optimise" this
  by inlining the match back into the inner loop.
- **Builder setters clone on every call.** The conditional format and
  sparkline types are consuming builders upstream (`mut self -> Self`),
  so each `pymethod` does `inner = inner.clone().set_x(...)`.
  `ChartFormat` additionally clones its `ChartLine` / `ChartSolidFill`
  state per setter. All of this is configuration-time, called a handful
  of times per object, and none of it scales with row count — so it is
  **not** a target. Recorded here only so it is not mistaken for one.
- **Nanosecond timestamps lose sub-microsecond precision.** An i64
  nanosecond count for a modern date exceeds f64's 2^53 exact-integer
  range. This is a correctness note rather than a performance one, and
  it is not fixable: Excel serial dates cannot represent that precision
  either, so nothing is lost that could have been stored.

---

## Suggested order

1. `SmallVec` for `write_rows` (#3) — clearest win, lowest risk.
2. Lazy temporal `Format` construction — two lines, trivially safe.
3. String deduplication — biggest theoretical win, self-contained.
4. `PyWeakref` for the workbook ref (#2) — correctness as much as speed.
5. `target-cpu=native` opt-in profile.
6. `UnsafeCell` (#1) — only with the soundness argument written down.
7. Column-major ordering (#4) — only if the `constant_memory`
   interaction can be resolved cleanly.
