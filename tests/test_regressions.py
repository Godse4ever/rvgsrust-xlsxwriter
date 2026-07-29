"""Regression tests for bugs found in the 0.2.0 audit.

Each test here pins the *observable* consequence of a specific fixed bug,
so a future refactor that reintroduces the root cause fails loudly. They
deliberately assert on the saved workbook (via openpyxl) rather than on
internal state, because every bug in this file was silent at the API
level -- the calls all returned successfully and produced a valid .xlsx
that simply contained the wrong thing.
"""
import os

import openpyxl
import pytest

from rvgsrust_xlsxwriter import Workbook

TEST_FILE = "test_regressions.xlsx"


def teardown_function(function):
    if os.path.exists(TEST_FILE):
        os.remove(TEST_FILE)


# ---------------------------------------------------------------------
# add_worksheet() index desync after a rejected sheet name
# ---------------------------------------------------------------------
# Previously add_worksheet() appended the worksheet, then called
# set_name(), then advanced its index counter. A rejected name returned
# early from the middle of that sequence, leaving an unnamed worksheet in
# the workbook while the counter had not moved. The next add_worksheet()
# therefore returned an index pointing at the orphan, and every write
# through that handle landed on the wrong sheet while the intended sheet
# saved empty. Nothing raised.


@pytest.mark.parametrize(
    "bad_name",
    [
        "A" * 32,      # exceeds Excel's 31-character limit
        "bad/name",    # contains an invalid character
        "bad:name",
        "bad[name]",
        "",            # blank
    ],
)
def test_rejected_sheet_name_leaves_workbook_unchanged(bad_name):
    """A rejected name must not append a worksheet, and must not shift
    the index of any worksheet added afterwards."""
    wb = Workbook()
    wb.add_worksheet("First")

    with pytest.raises(ValueError):
        wb.add_worksheet(bad_name)

    second = wb.add_worksheet("Second")
    second.write(0, 0, "landed-correctly")
    wb.close(TEST_FILE)

    book = openpyxl.load_workbook(TEST_FILE)
    # No orphan sheet from the failed call.
    assert book.sheetnames == ["First", "Second"]
    # The write went to the sheet the caller actually asked for.
    assert book["Second"]["A1"].value == "landed-correctly"
    assert book["First"]["A1"].value is None


def test_many_sheets_keep_stable_indices_across_failures():
    """Interleaving failures with successes must not corrupt any handle."""
    wb = Workbook()
    handles = {}
    for i in range(5):
        handles[f"S{i}"] = wb.add_worksheet(f"S{i}")
        with pytest.raises(ValueError):
            wb.add_worksheet("X" * 32)

    for name, ws in handles.items():
        ws.write(0, 0, name)
    wb.close(TEST_FILE)

    book = openpyxl.load_workbook(TEST_FILE)
    assert book.sheetnames == [f"S{i}" for i in range(5)]
    # Each sheet must contain its own name, not a neighbour's.
    for name in handles:
        assert book[name]["A1"].value == name


# ---------------------------------------------------------------------
# add_table() bypassed the constant_memory row-order guard
# ---------------------------------------------------------------------
# add_table() writes header/total cells immediately rather than at save
# time, but was the only cell-writing Worksheet method with no
# check_row_order call. On a constant_memory worksheet, a table anchored
# above the current high-water mark was accepted and silently produced a
# corrupt file instead of raising.


def test_add_table_rejects_backward_write_in_constant_memory():
    from rvgsrust_xlsxwriter import Table

    wb = Workbook()
    ws = wb.add_worksheet("CM", constant_memory=True)

    # Advance the high-water mark well past where the table would start.
    ws.write(50, 0, "later-row")

    table = Table()
    with pytest.raises(ValueError, match="constant_memory"):
        ws.add_table(0, 0, 10, 2, table)


def test_add_table_still_allowed_in_order_in_constant_memory():
    """The guard must not reject a legitimate forward-ordered table."""
    from rvgsrust_xlsxwriter import Table

    wb = Workbook()
    ws = wb.add_worksheet("CM", constant_memory=True)
    ws.write(0, 0, "header-area")

    table = Table()
    ws.add_table(5, 0, 10, 2, table)  # forward of row 0 -- must not raise
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_add_table_unaffected_on_normal_worksheet():
    """Non-constant_memory sheets keep unrestricted ordering."""
    from rvgsrust_xlsxwriter import Table

    wb = Workbook()
    ws = wb.add_worksheet("Normal")
    ws.write(50, 0, "later-row")

    table = Table()
    ws.add_table(0, 0, 10, 2, table)  # backward, but allowed here
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


# ---------------------------------------------------------------------
# Re-entrant workbook access panicked instead of raising
# ---------------------------------------------------------------------
# write_records() holds the workbook's RefCell borrow across the whole
# loop while calling classify() per cell, and classify() falls back to
# value.str() for unrecognised types. A __str__ that touches the same
# Workbook re-entered borrow_mut() and hit a Rust panic surfacing as
# pyo3's PanicException with only "already mutably borrowed". It is now a
# RuntimeError explaining the cause.


def test_reentrant_write_from_dunder_str_raises_runtime_error():
    wb = Workbook()
    ws = wb.add_worksheet("S")

    class ReentrantValue:
        """Has no __float__/__index__ and is not a str, so classify()
        falls through to value.str() -- running this __str__ while the
        workbook borrow is held."""

        def __str__(self):
            ws.write(99, 0, "re-entrant")
            return "value"

    with pytest.raises(RuntimeError, match="already being modified"):
        ws.write_records(0, 0, [{"col": ReentrantValue()}])


def test_reentrant_value_is_fine_on_the_per_cell_path():
    """write() classifies before taking the borrow, so the same value is
    harmless there. Pins that the fix did not over-restrict."""
    wb = Workbook()
    ws = wb.add_worksheet("S")

    class Chatty:
        def __str__(self):
            return "converted"

    ws.write(0, 0, Chatty())
    wb.close(TEST_FILE)

    book = openpyxl.load_workbook(TEST_FILE)
    assert book["S"]["A1"].value == "converted"
