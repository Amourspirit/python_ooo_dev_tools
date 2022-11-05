from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from typing import cast
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.calc import Calc, SolverConstraintOperator


def test_linear_solver(loader, capsys: pytest.CaptureFixture) -> None:
    from com.sun.star.sheet import XSolver

    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 1000
    try:
        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)
        sheet = Calc.get_sheet(doc=doc, index=0)
        capsys.readouterr()  # clear buffer
        Calc.list_solvers()
        caputure = capsys.readouterr()
        # Services offered by the solver
        #   com.sun.star.comp.Calc.CoinMPSolver
        #   com.sun.star.comp.Calc.LpsolveSolver
        #   com.sun.star.comp.Calc.SwarmSolver
        lst_result = [line.strip() for line in cast(str, caputure.out).splitlines()]
        # originally solver was set to com.sun.star.comp.Calc.CoinMPSolver
        # for unknown reason this stopped working on linux.
        # Ubuntu 22.04 LibreOffice 7.3 no-longer list com.sun.star.comp.Calc.CoinMPSolver
        # as a reported service.
        # strangly Windows 10, LibreOffice 7.3 does still list com.sun.star.comp.Calc.CoinMPSolver
        # as a service.
        # for these reason switched to com.sun.star.comp.Calc.LpsolveSolver
        # Just need to get XSolver from a service for this example
        solver = "com.sun.star.comp.Calc.LpsolveSolver"
        assert solver in lst_result

        # specify the variable cells
        xpos = Calc.get_cell_address(sheet=sheet, cell_name="B1")  # X
        ypos = Calc.get_cell_address(sheet=sheet, cell_name="B2")  # Y

        vars = (xpos, ypos)

        # specify profit equation
        Calc.set_val(value="=143*B1 + 60*B2", sheet=sheet, cell_name="B3")
        profit_eq = Calc.get_cell_address(sheet, "B3")

        # set up equation formulae without inequalities
        Calc.set_val(value="=120*B1 + 210*B2", sheet=sheet, cell_name="B4")
        Calc.set_val(value="=110*B1 + 30*B2", sheet=sheet, cell_name="B5")
        Calc.set_val(value="=B1 + B2", sheet=sheet, cell_name="B6")

        # create the constraints
        # constraints are equations and their inequalities
        sc1 = Calc.make_constraint(num=15000, op="<=", sheet=sheet, cell_name="B4")
        #   20x + 210y <= 15000
        #   B4 is the address of the cell that is constrained
        sc2 = Calc.make_constraint(num=4000, op=SolverConstraintOperator.LESS_EQUAL, sheet=sheet, cell_name="B5")
        #   110x + 30y <= 4000
        sc3 = Calc.make_constraint(num=75, op="<=", sheet=sheet, cell_name="B6")
        #   x + y <= 75

        # could also include x >= 0 and y >= 0
        constraints = (sc1, sc2, sc3)

        # initialize the linear solver (CoinMP or basic linear)
        solver = Lo.create_instance_mcf(XSolver, solver, raise_err=True)
        solver.Document = doc
        solver.Objective = profit_eq
        solver.Variables = vars
        solver.Constraints = constraints
        solver.Maximize = True

        # restrict the search to the top-right quadrant of the graph
        Props.set_property(prop_set=solver, name="NonNegative", value=True)

        # execute the solver
        solver.solve()
        capsys.readouterr()  # clear buffer
        Calc.solver_report(solver=solver)
        # Profit max == $6315.625, with x == 21.875 and y == 53.125
        caputure = capsys.readouterr()
        # Services offered by the solver
        #   com.sun.star.comp.Calc.CoinMPSolver
        #   com.sun.star.comp.Calc.LpsolveSolver
        #   com.sun.star.comp.Calc.SwarmSolver
        lst_rpt_result = [line.strip() for line in cast(str, caputure.out).splitlines()]
        # ['Solver result:', 'B3 == 6315.6250', 'Solver variables:', 'B1 == 21.8750', 'B2 == 53.1250', '']
        assert len(lst_rpt_result) == 6
        assert lst_rpt_result[1] == "B3 == 6315.6250"
        assert lst_rpt_result[3] == "B1 == 21.8750"
        assert lst_rpt_result[4] == "B2 == 53.1250"

        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
