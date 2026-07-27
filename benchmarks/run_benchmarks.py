#!/usr/bin/env python3
"""
rvgsrust-xlsxwriter  –  MacBook Benchmark Suite
================================================
Compares rvgsrust-xlsxwriter against every mainstream Python XLSX writer.

Libraries tested (all optional; missing ones are skipped gracefully):
  - rvgsrust-xlsxwriter   our library (maturin develop --release)
  - xlsxwriter            pure-Python, the most-used library
  - openpyxl              pure-Python, most-complete API
  - pandas.to_excel()     via openpyxl and xlsxwriter engines

Usage:
  # 1. Build and install our library in release mode
  maturin develop --release

  # 2. Install comparison libraries
  pip install xlsxwriter openpyxl pandas pyarrow polars

  # 3. Run
  python benchmarks/run_benchmarks.py

  # Quicker smoke-test (1k + 10k rows only):
  python benchmarks/run_benchmarks.py --small

  # Fewer timed runs:
  python benchmarks/run_benchmarks.py --runs 3
"""

from __future__ import annotations

import argparse
import gc
import os
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from typing import Callable

# ─────────────────────────────────────────────────────────
# OPTIONAL LIBRARY DETECTION
# ─────────────────────────────────────────────────────────

def _try_import(name: str):
    try:
        return __import__(name)
    except ImportError:
        return None


rvgs  = _try_import("rvgsrust_xlsxwriter")
xlsxw = _try_import("xlsxwriter")
opx   = _try_import("openpyxl")
pd    = _try_import("pandas")
pl    = _try_import("polars")
pa    = _try_import("pyarrow")


def _pkg_version(pkg_name: str) -> str:
    try:
        import importlib.metadata
        return importlib.metadata.version(pkg_name)
    except Exception:
        return "not installed"


# ─────────────────────────────────────────────────────────
# SYSTEM INFO
# ─────────────────────────────────────────────────────────

def print_system_info() -> None:
    uname = platform.uname()
    print("=" * 72)
    print("SYSTEM")
    print("=" * 72)
    print(f"  OS          : {uname.system} {uname.release} ({uname.machine})")

    cpu_label = uname.processor or uname.machine
    if uname.system == "Darwin":
        try:
            cpu_label = subprocess.check_output(
                ["sysctl", "-n", "machdep.cpu.brand_string"],
                stderr=subprocess.DEVNULL,
            ).decode().strip()
        except Exception:
            pass

    print(f"  CPU         : {cpu_label}")
    print(f"  Python      : {sys.version.split()[0]}")
    print()
    print(f"  {'Library':<28} Version")
    print(f"  {'─'*28} ─────────────────")
    for display, pkg in [
        ("rvgsrust-xlsxwriter",    "rvgsrust_xlsxwriter"),
        ("xlsxwriter",             "xlsxwriter"),
        ("openpyxl",               "openpyxl"),
        ("pandas",                 "pandas"),
        ("polars",                 "polars"),
        ("pyarrow",                "pyarrow"),
    ]:
        ver = _pkg_version(pkg)
        print(f"  {display:<28} {ver}")
    print()


# ─────────────────────────────────────────────────────────
# DATA GENERATION
# ─────────────────────────────────────────────────────────

HEADERS = ["ID", "Name", "Department", "Salary", "Bonus", "Active", "Score", "Region"]
DEPTS   = ["Sales", "Engineering", "Marketing", "HR", "Finance", "Legal"]
REGIONS = ["North", "South", "East", "West", "Central"]


def make_records(n: int) -> list[dict]:
    return [
        {
            "ID":         i,
            "Name":       f"Employee_{i:06d}",
            "Department": DEPTS[i % len(DEPTS)],
            "Salary":     50_000 + (i % 100) * 500,
            "Bonus":      (i % 50) * 200,
            "Active":     i % 3 != 0,
            "Score":      round(60.0 + (i % 400) / 10.0, 1),
            "Region":     REGIONS[i % len(REGIONS)],
        }
        for i in range(n)
    ]


def make_pandas_df(n: int):
    return pd.DataFrame(make_records(n)) if pd else None


def make_polars_df(n: int):
    return pl.DataFrame(make_records(n)) if pl else None


def make_pyarrow_table(n: int):
    if pa is None:
        return None
    rows = make_records(n)
    return pa.table({k: [r[k] for r in rows] for k in HEADERS})


# ─────────────────────────────────────────────────────────
# RESULT TYPE
# ─────────────────────────────────────────────────────────

@dataclass
class Result:
    label: str
    times: list[float] = field(default_factory=list)
    file_bytes: int = 0
    error: str = ""

    @property
    def mean(self) -> float:
        return statistics.mean(self.times) if self.times else float("inf")

    @property
    def stdev(self) -> float:
        return statistics.stdev(self.times) if len(self.times) > 1 else 0.0

    @property
    def best(self) -> float:
        return min(self.times) if self.times else float("inf")


# ─────────────────────────────────────────────────────────
# TIMING HARNESS
# ─────────────────────────────────────────────────────────

def bench(
    label: str,
    fn: Callable[[str], None],
    runs: int = 5,
    warmup: int = 1,
) -> Result:
    """
    Time fn(path) where fn writes an xlsx to path.
    One warmup run is discarded. gc.collect() before each timed run.
    File is deleted between runs to avoid caching effects.
    """
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name

    times: list[float] = []
    file_bytes = 0

    try:
        for i in range(warmup + runs):
            if os.path.exists(path):
                os.remove(path)
            gc.collect()
            t0 = time.perf_counter()
            fn(path)
            t1 = time.perf_counter()
            if i >= warmup:
                times.append(t1 - t0)
                if os.path.exists(path):
                    file_bytes = os.path.getsize(path)
    except Exception as e:
        return Result(label=label, error=str(e)[:140])
    finally:
        if os.path.exists(path):
            os.remove(path)

    return Result(label=label, times=times, file_bytes=file_bytes)


# ─────────────────────────────────────────────────────────
# WRITE FUNCTIONS — each takes a path: str argument
# ─────────────────────────────────────────────────────────

# ── rvgsrust-xlsxwriter ──────────────────────────────────

def fn_rvgs_write_records(records):
    def _f(path):
        wb = rvgs.Workbook()
        ws = wb.add_worksheet()
        ws.write_records(0, 0, records)
        wb.close(path)
    return _f


def fn_rvgs_write_records_cm(records):
    """Same as above, using context-manager constructor."""
    def _f(path):
        with rvgs.Workbook(path) as wb:
            ws = wb.add_worksheet()
            ws.write_records(0, 0, records)
    return _f


def fn_rvgs_write_records_constmem(records):
    def _f(path):
        wb = rvgs.Workbook()
        ws = wb.add_worksheet(constant_memory=True)
        ws.write_records(0, 0, records)
        wb.close(path)
    return _f


def fn_rvgs_dataframe_pandas(df):
    def _f(path):
        wb = rvgs.Workbook()
        ws = wb.add_worksheet()
        ws.write_dataframe(0, 0, df)
        wb.close(path)
    return _f


def fn_rvgs_dataframe_polars(df):
    def _f(path):
        from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe
        wb = rvgs.Workbook()
        ws = wb.add_worksheet()
        write_polars_dataframe(ws, df)
        wb.close(path)
    return _f


def fn_rvgs_dataframe_pyarrow(tbl):
    def _f(path):
        wb = rvgs.Workbook()
        ws = wb.add_worksheet()
        ws.write_dataframe(0, 0, tbl)
        wb.close(path)
    return _f


def fn_rvgs_per_cell(records):
    def _f(path):
        wb = rvgs.Workbook()
        ws = wb.add_worksheet()
        for col, h in enumerate(HEADERS):
            ws.write(0, col, h)
        for r, row in enumerate(records, 1):
            for c, key in enumerate(HEADERS):
                ws.write(r, c, row[key])
        wb.close(path)
    return _f


# ── xlsxwriter ───────────────────────────────────────────

def fn_xlsxw_write_row(records):
    def _f(path):
        wb = xlsxw.Workbook(path)
        ws = wb.add_worksheet()
        ws.write_row(0, 0, HEADERS)
        for r, row in enumerate(records, 1):
            ws.write_row(r, 0, [row[k] for k in HEADERS])
        wb.close()
    return _f


def fn_xlsxw_per_cell(records):
    def _f(path):
        wb = xlsxw.Workbook(path)
        ws = wb.add_worksheet()
        ws.write_row(0, 0, HEADERS)
        for r, row in enumerate(records, 1):
            for c, key in enumerate(HEADERS):
                ws.write(r, c, row[key])
        wb.close()
    return _f


def fn_xlsxw_constmem(records):
    def _f(path):
        wb = xlsxw.Workbook(path, {"constant_memory": True})
        ws = wb.add_worksheet()
        ws.write_row(0, 0, HEADERS)
        for r, row in enumerate(records, 1):
            ws.write_row(r, 0, [row[k] for k in HEADERS])
        wb.close()
    return _f


# ── openpyxl ─────────────────────────────────────────────

def fn_opx_write_only(records):
    def _f(path):
        wb = opx.Workbook(write_only=True)
        ws = wb.create_sheet()
        ws.append(HEADERS)
        for row in records:
            ws.append([row[k] for k in HEADERS])
        wb.save(path)
    return _f


def fn_opx_normal(records):
    def _f(path):
        wb = opx.Workbook(write_only=False)
        ws = wb.active
        ws.append(HEADERS)
        for row in records:
            ws.append([row[k] for k in HEADERS])
        wb.save(path)
    return _f


# ── pandas.to_excel ──────────────────────────────────────

def fn_pandas_openpyxl(df):
    def _f(path):
        df.to_excel(path, index=False, engine="openpyxl")
    return _f


def fn_pandas_xlsxwriter(df):
    def _f(path):
        with pd.ExcelWriter(path, engine="xlsxwriter") as w:
            df.to_excel(w, index=False)
    return _f


# ─────────────────────────────────────────────────────────
# RESULTS TABLE
# ─────────────────────────────────────────────────────────

def print_results(results: list[Result]) -> None:
    valid = [r for r in results if not r.error and r.times]
    failed = [r for r in results if r.error]

    if not valid and not failed:
        return

    fastest_mean = min(r.mean for r in valid) if valid else 1.0

    W_label = 42
    print()
    header = (
        f"  {'METHOD':<{W_label}} {'MEAN':>8}  {'BEST':>8}  {'±':>7}  "
        f"{'vs FASTEST':>13}  {'FILE':>8}"
    )
    print(header)
    print("  " + "─" * (len(header) - 2))

    for r in sorted(valid, key=lambda x: x.mean):
        ratio = r.mean / fastest_mean
        ratio_str = "🏆 fastest" if ratio < 1.0005 else f"{ratio:.2f}x slower"
        size_kb = f"{r.file_bytes / 1024:.0f} KB" if r.file_bytes else "—"
        sd = f"±{r.stdev * 1000:.0f}ms" if r.stdev > 0 else "—"
        print(
            f"  {r.label:<{W_label}} {r.mean:>7.3f}s  {r.best:>7.3f}s  "
            f"{sd:>7}  {ratio_str:>13}  {size_kb:>8}"
        )

    for r in failed:
        print(f"  ✗  {r.label:<{W_label}} SKIPPED: {r.error}")


# ─────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--small",  action="store_true",
                        help="Only 1k and 10k rows")
    parser.add_argument("--runs",   type=int, default=5,
                        help="Timed runs per benchmark (default 5)")
    parser.add_argument("--warmup", type=int, default=1,
                        help="Discarded warmup runs (default 1)")
    args = parser.parse_args()

    row_sizes = [1_000, 10_000] if args.small else [1_000, 10_000, 100_000]
    R, W = args.runs, args.warmup

    print_system_info()
    print("=" * 72)
    print("CONFIGURATION")
    print("=" * 72)
    print(f"  Timed runs : {R}  |  Warmup runs : {W}")
    print(f"  Row sizes  : {', '.join(f'{n:,}' for n in row_sizes)}")
    print(f"  Columns    : {len(HEADERS)}  ({', '.join(HEADERS)})")
    print()

    for n in row_sizes:
        cells = n * len(HEADERS)
        print("=" * 72)
        print(f"  {n:,} rows × {len(HEADERS)} cols  =  {cells:,} cells")
        print("=" * 72)

        records     = make_records(n)
        pandas_df   = make_pandas_df(n)
        polars_df   = make_polars_df(n)
        pyarrow_tbl = make_pyarrow_table(n)

        results: list[Result] = []

        # ── rvgsrust-xlsxwriter ──────────────────────────
        if rvgs:
            print("\n  ── rvgsrust-xlsxwriter (Rust) ──────────────────")
            results += [
                bench("rvgs  write_records()",            fn_rvgs_write_records(records),        R, W),
                bench("rvgs  write_records()  ctx-mgr",   fn_rvgs_write_records_cm(records),     R, W),
                bench("rvgs  write_records()  const-mem", fn_rvgs_write_records_constmem(records),R,W),
            ]
            if pandas_df is not None:
                results.append(bench("rvgs  write_dataframe()  pandas",  fn_rvgs_dataframe_pandas(pandas_df),  R, W))
            if polars_df is not None:
                results.append(bench("rvgs  write_dataframe()  polars",  fn_rvgs_dataframe_polars(polars_df),  R, W))
            if pyarrow_tbl is not None:
                results.append(bench("rvgs  write_dataframe()  pyarrow", fn_rvgs_dataframe_pyarrow(pyarrow_tbl),R,W))
            if n <= 10_000:
                results.append(bench("rvgs  write()  per-cell",          fn_rvgs_per_cell(records),           R, W))

        # ── xlsxwriter ───────────────────────────────────
        if xlsxw:
            print("\n  ── xlsxwriter (pure Python) ─────────────────────")
            results += [
                bench("xlsxwriter  write_row()",          fn_xlsxw_write_row(records),   R, W),
                bench("xlsxwriter  per-cell",             fn_xlsxw_per_cell(records),    R, W),
                bench("xlsxwriter  constant_memory",      fn_xlsxw_constmem(records),    R, W),
            ]

        # ── openpyxl ─────────────────────────────────────
        if opx:
            print("\n  ── openpyxl (pure Python) ───────────────────────")
            results += [
                bench("openpyxl  write-only mode",        fn_opx_write_only(records),    R, W),
                bench("openpyxl  normal mode",            fn_opx_normal(records),        R, W),
            ]

        # ── pandas.to_excel ──────────────────────────────
        if pd and pandas_df is not None:
            print("\n  ── pandas.DataFrame.to_excel() ──────────────────")
            if opx:
                results.append(bench("pandas  to_excel()  openpyxl",    fn_pandas_openpyxl(pandas_df),    R, W))
            if xlsxw:
                results.append(bench("pandas  to_excel()  xlsxwriter",  fn_pandas_xlsxwriter(pandas_df),  R, W))

        print_results(results)
        print()

    print("=" * 72)
    print("NOTES")
    print("=" * 72)
    print("  Times      : wall-clock seconds (mean of timed runs).")
    print("  Warmup     : first run discarded to allow JIT/cache warmup.")
    print("  gc.collect : called before each timed run to reduce GC jitter.")
    print("  File       : written to temp dir, deleted between each run.")
    print()
    print("  rvgs write_records()   : single Python→Rust FFI call for whole dataset.")
    print("  rvgs write_dataframe() : zero-copy via Arrow C-stream interface.")
    print("  rvgs const-mem         : streams rows to disk, lower peak RAM.")
    print("  rvgs per-cell          : one FFI call per cell — intentionally slow,")
    print("                           shown to demonstrate the overhead cost.")
    print()


if __name__ == "__main__":
    main()
