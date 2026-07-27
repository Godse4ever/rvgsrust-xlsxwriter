# Performance

rvgsrust-xlsxwriter is substantially faster than pure-Python XLSX libraries
for bulk data writes. The key advantage is a single Python→Rust FFI crossing
for an entire dataset, plus a zero-copy Arrow C-stream interface for DataFrame
sources.

## Benchmark Results

**System:** Intel Core i5-6267U @ 2.90GHz, macOS, Python 3.12.11  
**Methodology:** 7 timed runs, 2 warmup runs discarded, `gc.collect()` before each run.  
*Note: high stdev reflects thermal throttling on a 2-core laptop. Apple Silicon results will be tighter.*

### 1,000 rows × 8 columns (8,000 cells)

| Library / Method                      | Mean   | vs fastest |
|---------------------------------------|--------|-----------|
| **rvgs** `write_dataframe()` pyarrow  | 0.013s | 🏆 fastest |
| **rvgs** `write_records()`            | 0.016s | 1.25×      |
| **rvgs** `write_dataframe()` polars   | 0.018s | 1.36×      |
| **rvgs** `write_dataframe()` pandas   | 0.018s | 1.36×      |
| **rvgs** `write_rows()`               | 0.020s | 1.53×      |
| xlsxwriter `write_row()`              | 0.074s | 5.6×       |
| openpyxl normal                       | 0.146s | 11×        |
| pandas `to_excel()` openpyxl          | 0.174s | 13×        |

### 10,000 rows × 8 columns (80,000 cells)

| Library / Method                      | Mean   | vs fastest |
|---------------------------------------|--------|-----------|
| **rvgs** `write_dataframe()` pyarrow  | 0.106s | 🏆 fastest |
| **rvgs** `write_dataframe()` polars   | 0.123s | 1.16×      |
| **rvgs** `write_dataframe()` pandas   | 0.128s | 1.21×      |
| **rvgs** `write_records()`            | 0.183s | 1.72×      |
| **rvgs** `write_rows()`               | 0.186s | 1.76×      |
| xlsxwriter `write_row()`              | 0.722s | **6.8× slower** |
| openpyxl normal                       | 1.566s | 14.8× slower |
| pandas `to_excel()` openpyxl          | 1.862s | 17.5× slower |

### 100,000 rows × 8 columns (800,000 cells)

| Library / Method                      | Mean    | vs fastest |
|---------------------------------------|---------|-----------|
| **rvgs** `write_dataframe()` pandas   | 1.433s  | 🏆 fastest |
| **rvgs** `write_dataframe()` polars   | 1.528s  | 1.07×      |
| **rvgs** `write_dataframe()` pyarrow  | 1.653s  | 1.15×      |
| **rvgs** `write_records()`            | 1.939s  | 1.35×      |
| **rvgs** `write_rows()`               | 2.025s  | 1.41×      |
| xlsxwriter `write_row()`              | 11.867s | **8.3× slower** |
| xlsxwriter `constant_memory`          | 14.825s | 10.3× slower |
| openpyxl write-only                   | 18.668s | 13× slower |
| pandas `to_excel()` xlsxwriter        | 20.049s | 14× slower |
| pandas `to_excel()` openpyxl          | 25.030s | 17.5× slower |

## Key Takeaways

- **DataFrame path is fastest** for all scales. If your data is already in
  pandas, polars, or pyarrow, `ws.write_dataframe()` is the right call — it
  uses the Arrow C-stream zero-copy interface.

- **6–17× faster than pure-Python alternatives** at 100k rows depending on
  which library and method you compare against.

- **`write_records()`** (list of dicts) and **`write_rows()`** (list of lists)
  are comparable and both significantly faster than xlsxwriter's best path.

- **`constant_memory=True`** reduces peak RAM for very large sheets by
  streaming rows to disk instead of buffering, but is not faster than normal
  mode (and slightly slower in practice due to I/O overhead).

## Reproducing

```bash
maturin develop --release
pip install xlsxwriter openpyxl pandas pyarrow polars
python benchmarks/run_benchmarks.py --runs 7 --warmup 2
```
