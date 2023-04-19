from __future__ import annotations
from typing import cast
import uno
from ooo.dyn.chart.missing_value_treatment import MissingValueTreatment
from com.sun.star.chart2 import XChartDocument

from enum import IntEnum
from .plot_simple import PlotSimple


class MissingValueKind(IntEnum):
    """
    This specifies how empty or invalid cells in the provided data should be handled when plotted.
    """

    LEAVE_GAP = MissingValueTreatment.LEAVE_GAP
    """Leave a gap in the line or bar chart."""
    USE_ZERO = MissingValueTreatment.USE_ZERO
    """Use zero as the value for the empty or invalid cell."""
    CONTINUE = MissingValueTreatment.CONTINUE
    """Use the value of the previous cell for the empty or invalid cell."""


class Plot(PlotSimple):
    """Chart Data Series Plot Options"""

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        missing_values: MissingValueKind | None = None,
        hidden_cell_values: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            missing_values (MissingValueTreatmentEnum | None, optional): Specifies plot missing values option. Defaults to ``None``.
            hidden_cell_values (bool | None, optional): Specifies if values from hidden cells are to be included. Defaults to ``None``.

        Returns:
            None:
        """
        super().__init__(chart_doc=chart_doc, hidden_cell_values=hidden_cell_values)
        if missing_values is not None:
            self.prop_missing_values = missing_values

    @property
    def prop_missing_values(self) -> MissingValueKind | None:
        """Gets or sets the missing value treatment."""
        pv = cast(int, self._get("MissingValueTreatment"))
        if pv is None:
            return None
        return MissingValueKind(pv)

    @prop_missing_values.setter
    def prop_missing_values(self, value: MissingValueKind | None) -> None:
        if value is None:
            self._remove("MissingValueTreatment")
            return
        self._set("MissingValueTreatment", value.value)
