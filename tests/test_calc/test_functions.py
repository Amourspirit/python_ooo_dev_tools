from __future__ import annotations
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.uno_util import UnoEnum
from ooodev.office.calc import Calc
import uno

def test_fill_series(loader) -> None:
    # https://help.libreoffice.org/latest/ro/text/sbasic/shared/03/sf_session.html?DbPAR=BASIC
    from com.sun.star.sheet import XFunctionAccess
    result = Calc.call_fun(func_name="ROUND", arg=1.999)
    assert result == 2.0
    
    # sine of 30 degrees
    result = Calc.call_fun("RADIANS", 30)
    assert result == pytest.approx(0.5235987755982988, rel=1e-4)
    
    # average function
    # test tuple
    args = (1, 2, 3, 4, 5)
    result = Calc.call_fun("AVERAGE", args)
    assert result == 3.0
    # test list
    args = [1, 2, 3, 4, 5]
    result = Calc.call_fun("AVERAGE", args)
    assert result == 3.0
    
    # basic
    # session.ExecuteCalcFunction("AVERAGE", 1, 5, 3, 7) ' 4
    # session.ExecuteCalcFunction("ABS", Array(Array(-1, 2, 3), Array(4, -5, 6), Array(7, 8, -9)))(2)(2) ' 9
    # session.ExecuteCalcFunction("LN", -3)
    # ' Generates an error.
    
    # python
    # session.ExecuteCalcFunction("AVERAGE", 1, 5, 3, 7) # 4
    # session.ExecuteCalcFunction("ABS", ((-1, 2, 3), (4, -5, 6), (7, 8, -9)))[2][2] # 9
    # session.ExecuteCalcFunction("LN", -3)
    
    #ABS
    # session.ExecuteCalcFunction("ABS", ((-1, 2, 3), (4, -5, 6), (7, 8, -9)))[2][2] # 9
    fa = Lo.create_instance_mcf(XFunctionAccess, "com.sun.star.sheet.FunctionAccess")
    args = (30,)
    result = fa.callFunction("RADIANS", args)
    assert result == pytest.approx(0.5235987755982988, rel=1e-4)
    args = ((-1, 2, 3), (4, -5, 6), (7, 8, -9))
    result = fa.callFunction("ABS", args)
    assert result[2][2] == 9
    
    # transpose a matrix
    arr = [[1.0, 2.0, 3.0],[4.0, 5.0, 6.0]]
    args = [arr]
    result = Calc.call_fun("TRANSPOSE", arr)
    
    assert result is not None
    
    # zTest function
    data = ((1.0, 2.0, 3.0),)
    args = (data, 2.0)
    result = Calc.call_fun("ZTEST", args)
    assert result is not None
    
    # slope function
    # xdata = [[1.0, 2.0, 3.0]] # must be a matrix
    # ydata = [[3.0, 6.0, 9.0]] # must be a matrix
    # args = [ydata, xdata]
    # result = Calc.call_fun("SLOPE", args)
    # assert result == 3.0
    
    