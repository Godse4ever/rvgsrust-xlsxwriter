"""
Benchmark: rvgsrust-xlsxwriter vs rustpy-xlsxwriter vs xlsxwriter
=================================================================
Run this to compare speeds on your machine.
"""
import time
import statistics
import os

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


def generate_data(rows):
    """Generate test data."""
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


def benchmark_rvgs(data, filename="bench_rvgs.xlsx"):
    wb = RVGSWorkbook()
    ws = wb.add_worksheet()

    fmt = RVGSWorkbook().add_format()
    fmt.set_bold()

    # Headers
    headers = list(data[0].keys())
    for col, header in enumerate(headers):
        ws.write(0, col, header, fmt)

    # Data
    for row_idx, row in enumerate(data, 1):
        for col_idx, key in enumerate(headers):
            ws.write(row_idx, col_idx, row[key])

    wb.close(filename)
    return filename


def benchmark_rustpy(data, filename="bench_rustpy.xlsx"):
    fe = FastExcel(filename)
    fe.sheet("Sheet1", data)
    fe.save()
    return filename


def benchmark_xlsxwriter(data, filename="bench_xlsxwriter.xlsx"):
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


def run_benchmark(name, func, data, runs=3):
    times = []
    for _ in range(runs):
        start = time.perf_counter()
        filename = func(data)
        elapsed = time.perf_counter() - start
        times.append(elapsed)
        if os.path.exists(filename):
            os.remove(filename)

    mean = statistics.mean(times)
    return name, mean, min(times), max(times)


def main():
    print("=" * 60)
    print("RVGSRust-XLSXWriter Benchmark")
    print("=" * 60)

    for rows in [1000, 10000, 100000]:
        print(f"
--- {rows:,} rows ---")
        data = generate_data(rows)

        results = []

        if HAS_RVGS:
            results.append(run_benchmark("rvgsrust", benchmark_rvgs, data))
        if HAS_RUSTPY:
            results.append(run_benchmark("rustpy", benchmark_rustpy, data))
        if HAS_XLSXWRITER:
            results.append(run_benchmark("xlsxwriter", benchmark_xlsxwriter, data))

        # Sort by speed (fastest first)
        results.sort(key=lambda x: x[1])

        print(f"{'Library':<15} {'Mean (s)':<12} {'Min (s)':<12} {'Max (s)':<12}")
        print("-" * 51)

        baseline = results[0][1]
        for name, mean, min_t, max_t in results:
            speedup = baseline / mean
            marker = " <-- FASTEST" if mean == baseline else f" ({speedup:.2f}x)"
            print(f"{name:<15} {mean:<12.4f} {min_t:<12.4f} {max_t:<12.4f}{marker}")


if __name__ == "__main__":
    main()
