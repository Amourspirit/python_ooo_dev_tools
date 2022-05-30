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
    round
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
    
    # transpose a matrix
    arr = ((1, 2, 3),(4, 5, 6))
    args = (arr,)
    result = Calc.call_fun("TRANSPOSE", list(arr))
    
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
    
    