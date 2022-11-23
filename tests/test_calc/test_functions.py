from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.calc import Calc


# Bare bones APSO test
#
# import uno
# smgr = XSCRIPTCONTEXT.ctx.getServiceManager()
# fa_obj = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", XSCRIPTCONTEXT.ctx)
# fa = fa_obj.queryInterface(uno.getTypeByName("com.sun.star.sheet.XFunctionAccess"))

# # test MAX
# args = (10 ,23, 33)
# result = fa.callFunction("MAX", args)
# assert result == 33.0

# # test ABS
# args = ((-1, 2, 3), (4, -5, 6), (7, 8, -9))
# result = fa.callFunction("ABS", args)
# # uno.com.sun.star.lang.IllegalArgumentException

# result = fa.callFunction("ABS", (args,))
# # uno.com.sun.star.lang.IllegalArgumentException


def test_abs_with_sheet(loader) -> None:
    # https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=31229
    # it seems loading args from array is causing a problem but from a sheet works better.
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    # test abs
    arr = ((-1, 2, 3), (4, -5, 6), (7, 8, -9))
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C3")
    result = Calc.call_fun("ABS", rng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result[2][2] == 9.0


def test_percentile_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = ((1.0, 2.0, 3.0),)
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
    result = Calc.call_fun("PERCENTILE", rng, 0.35)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result == 1.7


def test_transpose_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C3")
    result = Calc.call_fun("TRANSPOSE", rng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result[0][0] == 1.0
    assert result[0][1] == 4.0
    assert result[0][2] == ""
    assert result[1][0] == 2.0
    assert result[1][1] == 5.0
    assert result[1][2] == ""
    assert result[2][0] == 3.0
    assert result[2][1] == 6.0
    assert result[2][2] == ""


def test_ztest_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = ((1.0, 2.0, 3.0),)
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
    result = Calc.call_fun("ZTEST", rng, 2.0)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result == 0.5


def test_round(loader) -> None:
    result = Calc.call_fun("ROUND", 1.999)
    assert result == 2.0


def test_round_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = ((1.999,),)
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:A1")
    result = Calc.call_fun("ROUND", rng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result[0][0] == 2.0


def test_radians(loader) -> None:
    result = Calc.call_fun("Radians", 30)
    assert result == pytest.approx(0.5235987755982988, rel=1e-4)


def test_radians_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = ((30,),)
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:A1")
    result = Calc.call_fun("RADIANS", rng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result[0][0] == pytest.approx(0.5235987755982988, rel=1e-4)


def test_average(loader) -> None:
    result = Calc.call_fun("AVERAGE", 1, 2, 3, 4, 5)
    assert result == 3.0


def test_average_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = ((1, 2, 3, 4, 5),)
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:E1")
    result = Calc.call_fun("AVERAGE", rng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result == 3.0


def test_max(loader) -> None:
    result = Calc.call_fun("MAX", 10, 23, 33)
    assert result == 33.0


def test_max_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = ((10, 23, 33),)
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
    result = Calc.call_fun("MAX", rng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result == 33.0


def test_slop_with_sheet(loader) -> None:
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    if visible:
        GUI.set_visible(visible, doc)
    arr = [[1.0, 2.0, 3.0], [3.0, 6.0, 9.0]]
    Calc.set_array(values=arr, sheet=sheet, name="A1")
    Lo.delay(delay)
    xrng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
    yrng = Calc.get_cell_range(sheet=sheet, range_name="A2:C2")
    result = Calc.call_fun("SLOPE", yrng, xrng)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert result == 3.0


def test_sum_imaginary(loader) -> None:
    result = Calc.call_fun("IMSUM", "13+4j", "5+3j")
    assert result == "18+7j"


def test_dec_hex(loader) -> None:
    result = Calc.call_fun("DEC2HEX", 100, 4)
    assert result == "0064"


def test_rot13(loader) -> None:
    result = Calc.call_fun("ROT13", "hello")
    assert result == "uryyb"


def test_roman_numbers(loader) -> None:
    # http://cs.stackexchange.com/questions/7777/is-the-language-of-roman-numerals-ambiguous
    result = Calc.call_fun("ROMAN", 999)
    assert result == "CMXCIX"
    result = Calc.call_fun("ROMAN", 999, 4)
    assert result == "IM"


def test_address(loader) -> None:
    result = Calc.call_fun("ADDRESS", 2, 5, 4)
    assert result == "E2"


def test_get_recent_functions(loader) -> None:
    # result = Calc.get_recent_functions()
    # assert len(result) > 0
    show_recent_functions()


def show_recent_functions() -> None:
    recents = Calc.get_recent_functions()
    print(f"Recently used functions ({len(recents)}):")
    for i in recents:
        props = Calc.find_function(idx=i)
        print(f"  {Props.get_value(name='Name',props=props)}")
    print()


def _test_funcs(loader) -> None:
    # https://help.libreoffice.org/latest/ro/text/sbasic/shared/03/sf_session.html?DbPAR=BASIC

    # basic
    # session.ExecuteCalcFunction("AVERAGE", 1, 5, 3, 7) ' 4
    # session.ExecuteCalcFunction("ABS", Array(Array(-1, 2, 3), Array(4, -5, 6), Array(7, 8, -9)))(2)(2) ' 9
    # session.ExecuteCalcFunction("LN", -3)
    # ' Generates an error.

    # python
    # session.ExecuteCalcFunction("AVERAGE", 1, 5, 3, 7) # 4
    # session.ExecuteCalcFunction("ABS", ((-1, 2, 3), (4, -5, 6), (7, 8, -9)))[2][2] # 9
    # session.ExecuteCalcFunction("LN", -3)

    # ABS
    # session.ExecuteCalcFunction("ABS", ((-1, 2, 3), (4, -5, 6), (7, 8, -9)))[2][2] # 9
    from com.sun.star.sheet import XFunctionAccess

    fa = Lo.create_instance_mcf(XFunctionAccess, "com.sun.star.sheet.FunctionAccess")
    args = (30,)
    result = fa.callFunction("RADIANS", args)
    assert result == pytest.approx(0.5235987755982988, rel=1e-4)

    args = ((1.0, 2.0, 3.0), 0.35)
    result = fa.callFunction("PERCENTILE", args)
    assert result == 33.0

    args = ((-1, 2, 3), (4, -5, 6), (7, 8, -9))
    result = fa.callFunction("ABS", args)
    assert result[2][2] == 9

    # transpose a matrix
    arr = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
    args = [arr]
    result = Calc.call_fun("TRANSPOSE", arr)

    assert result is not None
