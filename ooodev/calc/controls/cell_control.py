from __future__ import annotations
import uno
from typing import TYPE_CHECKING, Tuple
from ooodev.calc.controls.sheet_control_base import SheetControlBase
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial


if TYPE_CHECKING:
    from ooodev.calc.calc_cell import CalcCell
    from ooodev.loader.inst.lo_inst import LoInst


class CellControl(SheetControlBase):
    """A partial class for a cell control."""

    def __init__(self, calc_obj: CalcCell, lo_inst: LoInst | None = None) -> None:
        super().__init__(calc_obj, lo_inst)

    def _init_calc_sheet_prop(self) -> None:
        CalcSheetPropPartial.__init__(self, self.calc_obj.calc_sheet)

    def _get_pos_size(self) -> Tuple[int, int, int, int]:
        ps = self.calc_obj.component.Position
        size = self.calc_obj.component.Size
        return (ps.X, ps.Y, size.Width, size.Height)

    @property
    def calc_obj(self) -> CalcCell:
        return super().calc_obj
