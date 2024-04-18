from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.exceptions.ex import GoalDivergenceError


def test_goal_seek(loader) -> None:
    from com.sun.star.sheet import XGoalSeek

    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 1000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    gs = Lo.qi(XGoalSeek, doc)
    assert gs is not None

    # -------------------------------------------------

    Calc.set_val(value=9, sheet=sheet, cell_name="C1")  # x-variable and starting value
    Calc.set_val(value="=SQRT(C1)", sheet=sheet, cell_name="C2")  # formula
    x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="C1", formula_cell_name="C2", result=4.0)
    assert x == 16.0

    with pytest.raises(GoalDivergenceError) as ge:
        # fails (i.e. divergence is big)
        x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="C1", formula_cell_name="C2", result=-4.0)

    # arg 0 is divergence value
    assert isinstance(ge.value.args[0], float)

    # -------------------------------------------------

    Calc.set_val(sheet=sheet, cell_name="D1", value=0.8)  # x-variable and starting value
    Calc.set_val(sheet=sheet, cell_name="D2", value="=(D1^2 - 1)/(D1 - 1)")  # formula
    # The formula is y = (x^2 -1)/(x-1)
    # After factoring, this is just y = x+1
    x = Calc.goal_seek(gs, sheet, "D1", "D2", 2)
    assert x == pytest.approx(1.0, rel=1e-5)

    # -------------------------------------------------
    Calc.set_val(100000, sheet, "B1")  # x-variable and starting value
    Calc.set_val(1, sheet, "B2")  # n, no. of years
    Calc.set_val(0.075, sheet, "B3")  # i, interest rate (7.5%)
    Calc.set_val("=B1*B2*B3", sheet, "B4")  # formula
    # The formula is Annual interest = x*n*r
    # where capital (x), number of years (n), and interest rate (r).
    # Find the capital, if the other values are given.
    x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="B1", formula_cell_name="B4", result=15000)
    assert x == 200_000.0

    # -------------------------------------------------
    Calc.set_val(value=0, sheet=sheet, cell_name="E1")  # x-variable and starting value
    Calc.set_val(value="=(E1^3 - 2*E1 + 2", sheet=sheet, cell_name="E2")  # formula
    x = Calc.goal_seek(gs=gs, sheet=sheet, cell_name="E1", formula_cell_name="E2", result=0)

    assert x == pytest.approx(-1.7692923428381226, rel=1e-5)
    # so not using Newton's method which oscillates between 0 and 1

    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
