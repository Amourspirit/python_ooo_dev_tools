from __future__ import annotations
from typing import Any, Tuple, cast
import uno
from ooo.dyn.chart2.legend_position import LegendPosition

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.abstract.abstract_writing_mode import AbstractWritingMode
from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo


class _WritingMode(AbstractWritingMode):
    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.Legend",
                "com.sun.star.style.ParagraphPropertiesComplex",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    # endregion overrides

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_mode(self) -> DirectionModeKind:
        """Gets/Sets writing mode."""
        pv = cast(int, self._get(self._get_property_name()))
        return DirectionModeKind(pv)

    @prop_mode.setter
    def prop_mode(self, value: DirectionModeKind):
        self._set(self._get_property_name(), value.value)

    # endregion properties


class Position(StyleMulti):
    """
    Chart2 legend position.

    .. seealso::

        - :ref:`help_chart2_format_direct_legend_position`
    """

    def __init__(
        self,
        *,
        pos: LegendPosition | None = None,
        mode: DirectionModeKind | None = None,
        no_overlap: bool | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            pos (LegendPosition | None, optional): Specifies the position of the legend.
            mode (DirectionModeKind, optional): Specifies the writing direction.
            no_overlap (bool | None, optional): Show the legend without overlapping the chart.

        See Also:
            - :ref:`help_chart2_format_direct_legend_position`
        """
        super().__init__()
        if pos is not None:
            self.prop_pos = pos
        if no_overlap is not None:
            self.prop_no_overlap = no_overlap
        if mode is not None:
            self._set_style("writing_mode", self._get_writing_mode(mode))

    # region internal methods
    def _get_writing_mode(self, mode: DirectionModeKind) -> _WritingMode:
        def get_cattribs() -> dict:
            return {
                "_format_kind_prop": self.prop_format_kind,
                "_property_name": "WritingMode",
                "_supported_services_values": self._supported_services(),
            }

        return _WritingMode(mode=mode, _cattribs=get_cattribs())

    # endregion internal methods

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Legend",)
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion overrides

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_mode(self) -> DirectionModeKind | None:
        """Gets or set the writing direction"""
        pv = cast(_WritingMode, self._get_style("writing_mode"))
        if pv is None:
            return None
        return pv.prop_mode

    @prop_mode.setter
    def prop_mode(self, value: DirectionModeKind | None) -> None:
        if value is None:
            self._remove_style("writing_mode")
            return
        self._set_style("writing_mode", self._get_writing_mode(value))

    @property
    def prop_pos(self) -> LegendPosition | None:
        """Gets or set the position of the legend"""
        return self._get("AnchorPosition")

    @prop_pos.setter
    def prop_pos(self, value: LegendPosition | None) -> None:
        if value is None:
            self._remove("AnchorPosition")
            return
        self._set("AnchorPosition", value)

    @property
    def prop_no_overlap(self) -> bool | None:
        """Gets or set the position of the legend"""
        pv = self._get("Overlay")
        if pv is None:
            return None
        return not pv

    @prop_no_overlap.setter
    def prop_no_overlap(self, value: bool | None) -> None:
        if value is None:
            self._remove("Overlay")
            return
        self._set("Overlay", not value)

    # endregion properties
