from collections.abc import Iterator

import pytest
from xlwings import Range as RangeImpl
from xlwings import Sheet

from xlviews.range.range import Range
from xlviews.range.range_collection import RangeCollection
from xlviews.testing import is_excel_installed

pytestmark = pytest.mark.skipif(not is_excel_installed(), reason="Excel not installed")


@pytest.fixture(scope="module", params=["A1", "A1:A3", "F4:I4", "C1:E3"])
def addr(request: pytest.FixtureRequest):
    return request.param


@pytest.fixture(scope="module")
def rng_impl(addr, sheet_module: Sheet):
    return sheet_module.range(addr)


@pytest.mark.parametrize("include_sheetname", [False, True])
@pytest.mark.parametrize("external", [False, True])
def test_range_str(
    addr,
    rng_impl: RangeImpl,
    include_sheetname,
    external,
    sheet_module: Sheet,
):
    addr = rng_impl.get_address(include_sheetname=include_sheetname, external=external)
    rng = Range(addr)
    assert range_str(rng) == addr


# @pytest.mark.parametrize("rng", ["A1", "A1:A3", "C1,E3", "B2:C4,E3:G4"])
# @pytest.mark.parametrize("apply", [lambda x: x, lambda x: RangeCollection([x])])
# def test_iter_addresses_str(rng, apply, sheet_module: Sheet):
#     from xlviews.range.address import iter_addresses

#     rngs = apply(sheet_module.range(rng))
#     x = list(iter_addresses(rngs, row_absolute=False, column_absolute=False))
#     assert x == [rng]


# @pytest.mark.parametrize(
#     "rngs",
#     [["$A$1"], ["$A$1", "$B$2"], ["$A$1:$B$3", "$C$4:$E$6"]],
# )
# @pytest.mark.parametrize("apply", [lambda x: x, lambda x: [RangeCollection(x)]])
# def test_iter_addresses_list(rngs, apply, sheet_module: Sheet):
#     from xlviews.range.address import iter_addresses

#     rngs_ = apply([sheet_module.range(r) for r in rngs])
#     assert list(iter_addresses(rngs_)) == rngs


# @pytest.mark.parametrize(
#     ("rngs", "addrs"),
#     [
#         (["A1"], ["=$A$1"]),
#         (["A1:A3"], ["=$A$1", "=$A$2", "=$A$3"]),
#         (["E4:F5"], ["=$E$4", "=$F$4", "=$E$5", "=$F$5"]),
#         (["E4:E5", "AAA7:AAA8"], ["=$E$4", "=$E$5", "=$AAA$7", "=$AAA$8"]),
#         (["E4:F4", "AB7:AD7"], ["=$E$4", "=$F$4", "=$AB$7", "=$AC$7", "=$AD$7"]),
#     ],
# )
# @pytest.mark.parametrize("apply", [lambda x: x, lambda x: [RangeCollection(x)]])
# def test_iter_addresses_cellwise_formula(rngs, addrs, apply, sheet_module: Sheet):
#     from xlviews.range.address import iter_addresses

#     rngs_ = apply([sheet_module.range(r) for r in rngs])
#     assert list(iter_addresses(rngs_, cellwise=True, formula=True)) == addrs


# @pytest.mark.parametrize(("rng", "row"), [("A5", 5), ("B6:C9", 6)])
# def test_range_row(sheet_module: Sheet, rng, row):
#     assert sheet_module.range(rng).row == row


# @pytest.mark.parametrize(("rng", "row"), [("A5", 5), ("B6:C9", 9)])
# def test_range_row_end(sheet_module: Sheet, rng, row):
#     assert sheet_module.range(rng)[-1].row == row


# @pytest.mark.parametrize(("rng", "column"), [("A5", 1), ("B6:C9", 2)])
# def test_range_column(sheet_module: Sheet, rng, column):
#     assert sheet_module.range(rng).column == column


# @pytest.mark.parametrize(("rng", "column"), [("A5", 1), ("B6:C9", 3)])
# def test_range_column_end(sheet_module: Sheet, rng, column):
#     assert sheet_module.range(rng)[-1].column == column


# def iter_addresses_from_range_impl(
#     rng: Range,
#     *,
#     row_absolute: bool = True,
#     column_absolute: bool = True,
#     include_sheetname: bool = False,
#     external: bool = False,
# ) -> Iterator[str]:
#     for r in rng:
#         yield r.get_address(
#             row_absolute=row_absolute,
#             column_absolute=column_absolute,
#             include_sheetname=include_sheetname,
#             external=external,
#         )


# @pytest.mark.parametrize("rng", ["A1", "A1:A3", "AA10:AZ10", "C10:H30"])
# @pytest.mark.parametrize("column_absolute", [False, True])
# @pytest.mark.parametrize("row_absolute", [False, True])
# @pytest.mark.parametrize("include_sheetname", [False, True])
# @pytest.mark.parametrize("external", [False, True])
# def test_iter_addresses_from_range(
#     rng,
#     column_absolute,
#     row_absolute,
#     include_sheetname,
#     external,
#     sheet_module: Sheet,
# ):
#     from xlviews.range.address import _iter_addresses_from_range

#     rng = sheet_module.range(rng)

#     xs = []
#     for f in [_iter_addresses_from_range, iter_addresses_from_range_impl]:
#         xs.append(
#             list(
#                 f(
#                     rng,
#                     row_absolute=row_absolute,
#                     column_absolute=column_absolute,
#                     include_sheetname=include_sheetname,
#                     external=external,
#                 ),
#             ),
#         )

#     assert xs[0] == xs[1]
