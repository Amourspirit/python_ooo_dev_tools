from __future__ import annotations
from typing import cast, Tuple
import uno
from com.sun.star.chart2 import XAxis
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo


class PositionAxis(StyleBase):
    """
    Chart Axis Line Position.

    Note:
        Does not apply to all axis types.

    .. seealso::

        - :ref:`help_chart2_format_direct_axis_positioning`

    .. versionadded:: 0.9.4
    """

    def __init__(self, on_mark: bool = True) -> None:
        """
        Constructor

        Args:
            on_mark(bool, optional): Specifies that the axis is position.
                If ``True``, specifies that the axis is positioned on the first/last tickmarks. This makes the data points visual representation begin/end at the value axis.
                If ``False``, specifies that the axis is positioned between the tickmarks. This makes the data points visual representation begin/end at a distance from the value axis.
                Default is ``True``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_axis_positioning`
        """
        super().__init__()
        self.prop_on_mark = on_mark

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values

    # region apply()
    def apply(self, obj: object, **kwargs) -> None:
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("apply")
            return
        try:
            axis = cast(XAxis, obj)
            sd = axis.getScaleData()
            cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
            cargs.event_data = self
            self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
            if cargs.cancel:
                return
            sd.ShiftedCategoryPosition = self.prop_on_mark
            axis.setScaleData(sd)
            eargs = EventArgs.from_args(cargs)
            self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)
        except Exception as ex:
            mLo.Lo.print(f"{self.__class__.__name__}.apply() error:")
            mLo.Lo.print(f"  {ex}")

    # endregion apply()

    # endregion overrides

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
    def prop_on_mark(self) -> bool:
        return self._get("ShiftedCategoryPosition")

    @prop_on_mark.setter
    def prop_on_mark(self, value: bool) -> None:
        self._set("ShiftedCategoryPosition", value)

    # endregion Properties
