"""
Benchmark: rvgsrust-xlsxwriter vs rustpy-xlsxwriter vs xlsxwriter
=================================================================
Comprehensive performance testing across different write strategies.
Tests bulk write, DataFrame, and per-cell methods.
"""
import time
import statistics
import os
import sys

try:
    from rvgsrust_xlsxwriter import Workbook as RVGSWorkbook
    HAS_RVGS = True
except ImportError:
    HAS_RVGS = False
    print("rvgsrust-xlsxwriter not installed. Install with: maturin develop --release")

try:
    from rustpy_xlsxwriter import FastExcel
    HAS_RUSTPY = True
except ImportError:
    HAS_RUSTPY = False
    print("rustpy-xlsxwriter not installed. Install with: pip install rustpy-xlsxwriter")

try:
    import xlsxwriter
    HAS_XLSXWRITER = True
except ImportError:
    HAS_XLSXWRITER = False
    print("xlsxwriter not installed. Install with: pip install xlsxwriter")

try:
    import polars as pl
    HAS_POLARS = True
except ImportError:
    HAS_POLARS = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


def generate_data(rows):
    """Generate test data as list of dicts."""
    return [
        {
            "Name": f"Employee_{i:05d}",
            "Department": ["Sales", "Engineering", "Marketing", "HR"][i % 4],
            "Salary": 50000 + (i % 100) * 1000,
            "Bonus": (i % 50) * 100,
            "Active": i % 3 == 0,
        }
        for i in range(rows)
    ]


def generate_polars_df(rows):
    """Generate test data as Polars DataFrame."""
    if not HAS_POLARS:
        return None
    data = generate_data(rows)
    return pl.DataFrame(data)


def generate_pandas_df(rows):
    """Generate test data as Pandas DataFrame."""
    if not HAS_PANDAS:
        return None
    data = generate_data(rows)
    return pd.DataFrame(data)


# ============================================
# RVGSRUST-XLSXWRITER BENCHMARKS
# ============================================

def benchmark_rvgs_write_records(data, filename="bench_rvgs_records.xlsx"):
    """Bulk write via write_records(): optimized for dict data.
    Single Python->Rust call for entire dataset.
    """
    wb = RVGSWorkbook()
    ws = wb.add_worksheet()

    fmt = wb.add_format()
    fmt.set_bold()

    ws.write_records(0, 0, data, header_format=fmt)
    wb.close(filename)
    return filename


def benchmark_rvgs_write_dataframe_polars(df, filename="bench_rvgs_polars.xlsx"):
    """Zero-copy DataFrame write via Arrow PyCapsule Interface (Polars).
    Fastest path: directly reads from Arrow columnar buffers.
    """
    if df is None:
        return None
    
    from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe
    
    wb = RVGSWorkbook()
    ws = wb.add_worksheet()

    fmt = wb.add_format()
    fmt.set_bold()

    write_polars_dataframe(ws, df, row=0, col=0, header_format=fmt)
    wb.close(filename)
    return filename


def benchmark_rvgs_write_dataframe_pandas(df, filename="bench_rvgs_pandas.xlsx"):
    """Zero-copy DataFrame write via Arrow PyCapsule Interface (Pandas).
    Fast path: uses Arrow conversion internally.
    """
    if df is None:
        return None
    
    from rvgsrust_xlsxwriter.dataframe import write_pandas_dataframe
    
    wb = RVGSWorkbook()
    ws = wb.add_worksheet()

    fmt = wb.add_format()
    fmt.set_bold()

    write_pandas_dataframe(ws, df, row=0, col=0, header_format=fmt)
    wb.close(filename)
    return filename


def benchmark_rvgs_cell_by_cell(data, filename="bench_rvgs_cell.xlsx"):
    """Per-cell write: slowest path, for comparison.
    Shows cost of individual Python->Rust FFI calls.
    For 100k rows x 5 cols = 500k FFI crossings.
    """
    wb = RVGSWorkbook()
    ws = wb.add_worksheet()

    fmt = wb.add_format()
    fmt.set_bold()

    headers = list(data[0].keys())
    for col, header in enumerate(headers):
        ws.write(0, col, header, fmt)

    for row_idx, row in enumerate(data, 1):
        for col_idx, key in enumerate(headers):
            ws.write(row_idx, col_idx, row[key])

    wb.close(filename)
    return filename


# ============================================
# RUSTPY-XLSXWRITER BENCHMARK
# ============================================

def benchmark_rustpy(data, filename="bench_rustpy.xlsx"):
    """RustPy-XLSXWriter's native API."""
    fe = FastExcel(filename)
    fe.sheet("Sheet1", data)
    fe.save()
    return filename


# ============================================
# XLSXWRITER (PURE PYTHON) BENCHMARK
# ============================================

def benchmark_xlsxwriter(data, filename="bench_xlsxwriter.xlsx"):
    """Pure Python xlsxwriter: baseline."""
    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet()

    fmt = wb.add_format({"bold": True})

    headers = list(data[0].keys())
    for col, header in enumerate(headers):
        ws.write(0, col, header, fmt)

    for row_idx, row in enumerate(data, 1):
        for col_idx, key in enumerate(headers):
            ws.write(row_idx, col_idx, row[key])

    wb.close()
    return filename


# ============================================
# BENCHMARK RUNNER
# ============================================

def run_benchmark(name, func, data, runs=3):
    """Run benchmark and collect timing statistics."""
    try:
        # Warmup run (discarded)
        warmup_file = func(data)
        if warmup_file and os.path.exists(warmup_file):
            os.remove(warmup_file)

        times = []
        for _ in range(runs):
            start = time.perf_counter()
            filename = func(data)
            elapsed = time.perf_counter() - start
            times.append(elapsed)
            if filename and os.path.exists(filename):
                os.remove(filename)

        mean = statistics.mean(times)
        return name, mean, min(times), max(times)
    except Exception as e:
        print(f"  ⚠️  {name}: {e}")
        return None


def main():
    print("=" * 80)
    print("RVGSRust-XLSXWriter v0.1.0 - Comprehensive Benchmark")
    print("=" * 80)
    print(f"Using rust_xlsxwriter 0.96 with zmij backend + LTO optimization")
    print()

    for rows in [1000, 10000, 100000]:
        print(f"\n{'─' * 80}")
        print(f"BENCHMARK: {rows:,} rows × 5 columns")
        print(f"{'─' * 80}")
        
        data = generate_data(rows)
        polars_df = generate_polars_df(rows) if HAS_POLARS else None
        pandas_df = generate_pandas_df(rows) if HAS_PANDAS else None

        results = []

        # RVGSRUST benchmarks
        if HAS_RVGS:
            print("\n📊 RVGSRust-XLSXWriter Tests:")
            
            result = run_benchmark("  • write_records() [bulk dict]", benchmark_rvgs_write_records, data)
            if result:
                results.append(result)
            
            if polars_df is not None:
                result = run_benchmark("  • write_dataframe() [Polars]", 
                                      lambda _: benchmark_rvgs_write_dataframe_polars(polars_df), 
                                      data)
                if result:
                    results.append(result)
            
            if pandas_df is not None:
                result = run_benchmark("  • write_dataframe() [Pandas]", 
                                      lambda _: benchmark_rvgs_write_dataframe_pandas(pandas_df), 
                                      data)
                if result:
                    results.append(result)
            
            result = run_benchmark("  • write() [per-cell] ⚠️ (slow)", benchmark_rvgs_cell_by_cell, data)
            if result:
                results.append(result)

        # RustPy benchmark
        if HAS_RUSTPY:
            print("\n⚡ RustPy-XLSXWriter:")
            result = run_benchmark("  • FastExcel.sheet() [native]", benchmark_rustpy, data)
            if result:
                results.append(result)

        # Pure Python baseline
        if HAS_XLSXWRITER:
            print("\n🐍 Python xlsxwriter (baseline):")
            result = run_benchmark("  • xlsxwriter (pure Python)", benchmark_xlsxwriter, data)
            if result:
                results.append(result)

        # Sort by speed
        results.sort(key=lambda x: x[1])

        print(f"\n{'RESULTS':<50} {'MEAN (s)':<12} {'MIN (s)':<12} {'MAX (s)':<12}")
        print(f"{'─' * 80}")

        if results:
            baseline = results[0][1]
            for name, mean, min_t, max_t in results:
                speedup = baseline / mean
                if speedup == 1.0:
                    marker = " 🏆 FASTEST"
                else:
                    marker = f" ({speedup:.2f}x)"
                print(f"{name:<50} {mean:<12.4f} {min_t:<12.4f} {max_t:<12.4f}{marker}")

    print(f"\n{'=' * 80}")
    print("📈 Summary:")
    print("  • write_records() - Recommended for dict/record data")
    print("  • write_dataframe() - Recommended for Polars/Pandas (zero-copy)")
    print("  • write() per-cell - Avoid for large datasets")
    print(f"{'=' * 80}\n")


if __name__ == "__main__":
    main()
