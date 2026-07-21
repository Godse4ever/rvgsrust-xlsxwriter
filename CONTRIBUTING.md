# Contributing to RVGSRust-XLSXWriter

Thank you for your interest in contributing! This project is built with Rust (PyO3) and Python.

## Development Setup

### Prerequisites

- Rust 1.70+ (`rustup update`)
- Python 3.8+
- `maturin` (`pip install maturin`)

### Build

```bash
# Clone the repo
git clone https://github.com/Godse4ever/rvgsrust-xlsxwriter.git
cd rvgsrust-xlsxwriter

# Build and install in development mode
maturin develop --release

# Run tests
pytest tests/ -v
```

### Project Structure

```
rust_xlsxwriter/
├── src/
│   └── lib.rs          # Rust source with PyO3 bindings
├── python/
│   └── rvgsrust_xlsxwriter/
│       ├── __init__.py   # Python package
│       └── dataframe.py  # DataFrame helpers
├── examples/             # Usage examples
├── tests/                # Test suite
├── Cargo.toml           # Rust dependencies
└── pyproject.toml       # Python package config
```

### Adding a New Feature

1. Add Rust binding in `src/lib.rs`
2. Add Python wrapper if needed in `python/rvgsrust_xlsxwriter/`
3. Add test in `tests/`
4. Add example in `examples/`
5. Update README.md

### Code Style

- Rust: `cargo fmt` and `cargo clippy`
- Python: `black` and `ruff`

## Reporting Issues

Please include:
- Python version
- Rust version (`rustc --version`)
- Operating system
- Minimal reproducible example

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
