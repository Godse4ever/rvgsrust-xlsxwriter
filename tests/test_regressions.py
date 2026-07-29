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
