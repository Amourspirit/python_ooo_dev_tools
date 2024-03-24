from __future__ import annotations
import uno
from typing import cast, Any, Iterable, TYPE_CHECKING, Tuple
from ooodev.calc.controls.sheet_control_base import SheetControlBase
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import Shape  # service
    from ooodev.calc.calc_cell_range import CalcCellRange
    from ooodev.loader.inst.lo_inst import LoInst


class CellRangeControl(SheetControlBase):
    """A partial class for a cell control."""

    def __init__(self, calc_obj: CalcCellRange, lo_inst: LoInst | None = None) -> None:
        super().__init__(calc_obj, lo_inst)

    def _init_calc_sheet_prop(self) -> None:
        CalcSheetPropPartial.__init__(self, self.calc_obj.calc_sheet)

    def _get_pos_size(self) -> Tuple[int, int, int, int]:
        ps = self.calc_obj.component.Position
        size = self.calc_obj.component.Size
        return (ps.X, ps.Y, size.Width, size.Height)

    def on_setting_shape_props(self, event_args: CancelEventArgs) -> None:
        """
        Event handler for setting shape properties.

        Triggers the ``setting_shape_props`` event.
        """
        co = self.calc_obj.range_obj.cell_start
        cell = self.calc_obj.calc_sheet[co]
        event_args.event_data["Anchor"] = cell.component
        super().on_setting_shape_props(event_args)

    @property
    def calc_obj(self) -> CalcCellRange:
        return super().calc_obj
