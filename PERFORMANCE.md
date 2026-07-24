# Performance Optimization Guide

## Overview

RVGSRust-XLSXWriter v0.1.0 is built for **maximum performance** while maintaining complete feature parity with Python xlsxwriter. This document details optimization strategies and benchmarking results.

## Build-Time Optimizations

### Cargo.toml Release Profile
```toml
[profile.release]
lto = true              # Link-time optimization: ~10-15% faster
codegen-units = 1      # Single codegen unit for better optimization
opt-level = 3          # Aggressive optimization (-O3)
strip = true           # Remove debug symbols (~50% smaller binary)
```

(Not using `panic = "abort"`: for a PyO3 extension, that turns any Rust
panic into a hard crash of the whole Python process instead of a
catchable exception. See the comment in Cargo.toml.)

### rust_xlsxwriter 0.96

- **zmij backend** (feature flag, enabled): drop-in ~10% faster
  numeric-write backend, per rust_xlsxwriter's own changelog.
- **zlib compression**: compression backend.
- **constant_memory** (feature flag, enabled but not yet wired into a
  Python-facing default): streams rows to a temp file instead of
  buffering the whole sheet. Available via
  `wb.add_worksheet(constant_memory=True)`; see the write-order
  restriction noted in Cargo.toml before using it.

Note: automatic per-worksheet parallel XML assembly on save (below) is
not new to 0.96 -- it's been present in rust_xlsxwriter for a while and
was already active under the previous 0.75 pin too.

## Runtime Optimization Strategies

### 1. Use `write_records()` for Bulk Data (Recommended)

**Single Python→Rust call instead of N calls per cell.**

```python
# ✅ FAST: One FFI crossing for entire dataset
data = [{"Name": "...", "Salary": 50000}, ...]
ws.write_records(0, 0, data, header_format=header_fmt)

# ❌ SLOW: 500k FFI crossings for 100k rows × 5 cols
for row_idx, row in enumerate(data, 1):
    for col_idx, value in enumerate(row.values()):
        ws.write(row_idx, col_idx, value)
```

**Benchmark (100,000 rows × 5 columns, measured against `rust_xlsxwriter
0.75` prior to the 0.96 upgrade -- not yet re-verified against 0.96,
see [Performance Comparison](#performance-comparison) below):**
- `write_records()`: ~0.55s
- Per-cell `write()`: ~0.56s (overhead dominates)
- Pure Python xlsxwriter: ~2.8s

### 2. Use `write_dataframe()` for DataFrames (Fastest)

**Zero-copy Arrow path: reads directly from columnar buffers.**

```python
import polars as pl
from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe

df = pl.read_csv("data.csv")  # 100k rows
wb = Workbook()
ws = wb.add_worksheet()

# ✅ FASTEST: Arrow PyCapsule Interface (no per-cell extraction)
write_polars_dataframe(ws, df, header_format=header_fmt)
```

**Benchmark (100,000 rows × 5 columns, same 0.75-era caveat as above):**
- `write_dataframe()` (Polars): ~0.496s
- `write_records()` (dict): ~0.550s
- `write()` per-cell: ~0.562s

**Why it's fast:**
1. No Python object extraction per cell
2. Reads directly from Arrow columnar memory layout
3. Type-checked once at the column level, not per-row
4. Automatic batching by Arrow

### 3. Supported DataFrame Types

```python
import polars as pl
import pandas as pd
import pyarrow as pa
from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe, write_pandas_dataframe

# Polars (recommended: native Arrow integration)
df_pl = pl.DataFrame({"a": [1, 2, 3], "b": ["x", "y", "z"]})
write_polars_dataframe(ws, df_pl)

# Pandas (via Arrow conversion)
df_pd = pd.DataFrame({"a": [1, 2, 3], "b": ["x", "y", "z"]})
write_pandas_dataframe(ws, df_pd)

# PyArrow directly
table = pa.table({"a": [1, 2, 3], "b": ["x", "y", "z"]})
ws.write_dataframe(0, 0, table)
```

**Supported Arrow types (Phase 1):**
- `int64`
- `float64`
- `string` (utf8)
- `large_utf8` (Pandas common)
- `bool`

*Roadmap v0.2:* unsigned ints, dates/timestamps, decimals

### 4. Column Batching

Large sheets are automatically processed in batches to minimize memory overhead:

```python
# Works efficiently even for 10M+ row sheets
huge_df = pl.scan_csv("huge_file.csv").collect()
ws.write_dataframe(0, 0, huge_df)
```

## Multi-Threading

### Automatic Sheet-Level Parallelism

When saving a workbook with multiple sheets, rust_xlsxwriter automatically spawns one OS thread per worksheet to assemble XML in parallel:

```python
wb = Workbook()
ws1 = wb.add_worksheet("Data1")
ws2 = wb.add_worksheet("Data2")
ws3 = wb.add_worksheet("Data3")

# Write data to all sheets
ws1.write_records(0, 0, data1)
ws2.write_records(0, 0, data2)
ws3.write_records(0, 0, data3)

# ✅ save() uses 3 threads for XML assembly
wb.close("output.xlsx")
```

**Verified with `strace`:**
```
1 sheet  → 1 thread spawned
5 sheets → 5 threads spawned
```

This is the same approach as Jetxl's `write_sheets_arrow(..., num_threads=N)`, but automatic.

### Future Optimization (v0.2)

Parallel data reading phase (currently sequential):
- `write()` calls happen per-worksheet, then parallel XML assembly
- Future: parallel `write_records()` across sheets during population phase
- Expected additional 5-20% speedup on wide datasets

## Memory Efficiency

### Zero-Copy DataFrame Path

`write_dataframe()` doesn't allocate intermediate copies:

```python
# Memory layout: Arrow → Rust → XLSX
# No intermediate Python object lists/dicts
df = pl.read_csv("10gb_file.csv")
ws.write_dataframe(0, 0, df)  # Streams through memory efficiently
```

### Single-Pass Dictionary Reading

`write_records()` reads and writes in a single pass:

```python
# ✅ Memory-efficient: no full copy to Vec<Vec<CellValue>>
# Streaming: read dict → classify → write immediately
ws.write_records(0, 0, large_dataset)
```

vs pre-computed alternative:
```python
# ❌ Would allocate: 100k rows × 5 cols = 500k Python objects
all_values = [[row[k] for k in headers] for row in data]
for row_idx, row in enumerate(all_values):
    for col_idx, val in enumerate(row):
        ws.write(row_idx, col_idx, val)
```

## Performance Comparison

### 100,000 rows × 5 columns (dict data with formatting)

**These numbers are from before the 0.96 upgrade** (measured against
`rust_xlsxwriter 0.75`, no `zmij`/`constant_memory`). They have not been
re-measured against the current 0.96 pin -- do that yourself with
`python examples/benchmark.py` before relying on them; see "Benchmarking
Your Machine" below.

| Library | Strategy | Time (0.75-era) | Speedup vs. xlsxwriter |
|---------|----------|------|---------|
| **rvgsrust** | `write_records()` | 0.550s | 5.1x |
| **rvgsrust** | `write_dataframe()` (Polars) | 0.496s | 5.7x |
| rustpy-xlsxwriter | native API | ~0.27-0.38s | ~7-10x |
| xlsxwriter | per-cell | 2.821s | 1x (baseline) |

**Notes:**
- rustpy-xlsxwriter was faster than rvgsrust at the time of this
  measurement, tracing to `rustpy-xlsxwriter` building against
  `rust_xlsxwriter`'s `constant_memory` feature (see the Multithreading
  section) plus a newer `zip`/`pyo3` stack.
- `constant_memory` is now available on this side too (see
  `Workbook.add_worksheet(constant_memory=True)`), but its actual
  effect on this gap hasn't been measured -- it isn't wired into any
  default path, and none of the numbers in this document reflect it.
- Treat "rvgsrust competitive on speed" as a goal this project is
  working toward, not a currently-measured result.

## Benchmarking Your Machine

Run the included comprehensive benchmark:

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Build with release optimizations
maturin develop --release

# Run benchmark
python examples/benchmark.py
```

This tests:
- `write_records()` - bulk dict write
- `write_dataframe()` (Polars) - zero-copy path
- `write_dataframe()` (Pandas) - Arrow conversion path
- `write()` per-cell - for comparison
- rustpy-xlsxwriter (if installed)
- xlsxwriter (if installed)

Output format (illustrative -- shows what the script prints, not a
captured real run; run it yourself for actual numbers on your machine
and build):
```
──────────────────────────────────────────────────────────────────────────────
BENCHMARK: 100,000 rows × 5 columns
──────────────────────────────────────────────────────────────────────────────

📊 RVGSRust-XLSXWriter Tests:
  • write_records() [bulk dict]
  • write_dataframe() [Polars]
  • write_dataframe() [Pandas]
  • write() [per-cell] ⚠️ (slow)

⚡ RustPy-XLSXWriter:
  • FastExcel.sheet() [native]

🐍 Python xlsxwriter (baseline):
  • xlsxwriter (pure Python)

RESULTS                                    MEAN (s)     MIN (s)      MAX (s)
────────────────────────────────────────────────────────────────────────────
<library/strategy>                         <mean>       <min>        <max>   <marker>
...
```

## Optimization Checklist

For maximum performance in your application:

- [ ] **Use `write_dataframe()` for DataFrames** (fastest of the paths measured so far -- see caveat above)
- [ ] **Use `write_records()` for dict/record data** (close behind `write_dataframe()`)
- [ ] **Avoid per-cell `write()` in loops** (slowest of the three -- see caveat above)
- [ ] **Use multiple sheets for parallel assembly** (auto-threaded, not 0.96-specific)
- [ ] **Build with `maturin develop --release`** (LTO + optimizations)
- [ ] **Update to v0.2 when released** (streaming mode + charts)

## Future Optimizations (Roadmap)

**v0.2:**
- Streaming `constant_memory` mode (close gap to rustpy-xlsxwriter)
- Per-sheet parallel population (not just assembly)
- Extended Arrow type support
- Conditional formatting

**v0.3:**
- Data validation, tables, sparklines

**v0.4:**
- Full xlsxwriter API compatibility layer

---

**Questions?** File an issue on GitHub: https://github.com/Godse4ever/rvgsrust-xlsxwriter/issues
