from __future__ import annotations
from typing import Any, Tuple, TypeVar, overload
import uno
from com.sun.star.chart2 import XChartDocument

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo

_TPlotSimple = TypeVar("_TPlotSimple", bound="PlotSimple")


class PlotSimple(StyleBase):
    """
    Data Series Plot Simple

    ..seealso::

        - :ref:`help_chart2_format_direct_series_series_options`
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        hidden_cell_values: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            hidden_cell_values (bool, optional): Specifies if values from hidden cells are to be included. Defaults to ``None``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_series_options`
        """
        self._chart_doc = chart_doc
        super().__init__()
        if hidden_cell_values is not None:
            self.prop_hidden_cell_values = hidden_cell_values

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.DataSeries",)
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("apply")
            return

        try:
            diagram = self._chart_doc.getDiagram()  # type: ignore
        except Exception as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply() - Unable to get chart diagram")
            mLo.Lo.print(f"  Error: {e}")
            return

        super().apply(obj=diagram, validate=False, **kwargs)

    # endregion apply()

    # region Copy()
    @overload
    def copy(self: _TPlotSimple) -> _TPlotSimple: ...

    @overload
    def copy(self: _TPlotSimple, **kwargs) -> _TPlotSimple: ...

    def copy(self: _TPlotSimple, **kwargs) -> _TPlotSimple:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion Copy()
    # endregion Overrides

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_hidden_cell_values(self) -> bool | None:
        """Gets or sets whether hidden cells are included."""
        return self._get("IncludeHiddenCells")

    @prop_hidden_cell_values.setter
    def prop_hidden_cell_values(self, value: bool | None) -> None:
        if value is None:
            self._remove("IncludeHiddenCells")
            return
        self._set("IncludeHiddenCells", value)

    # endregion Properties
