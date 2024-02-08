from __future__ import annotations
from typing import Any, TYPE_CHECKING, TypeVar, Generic

from ooodev.adapter.chart2.title_comp import TitleComp
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.format.inner.partial.font_effects_partial import FontEffectsPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.units import Angle

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT
    from ooodev.events.args.cancel_event_args import CancelEventArgs

_T = TypeVar("_T", bound="ComponentT")


class ChartTitle(
    Generic[_T],
    LoInstPropsPartial,
    EventsPartial,
    TitleComp,
    PropPartial,
    FontEffectsPartial,
    QiPartial,
    ServicePartial,
    StylePartial,
):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(self, owner: _T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        EventsPartial.__init__(self)
        TitleComp.__init__(self, component=component)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        FontEffectsPartial.__init__(self, factory_name="ooodev.chart2.title", component=component, lo_inst=lo_inst)
        self._owner = owner
        self._init_events()

    # region Events
    def _init_events(self) -> None:
        self._fn_on_apply_style_text = self._on_apply_style_text
        self._fn_on_global_cancel = self._on_global_cancel
        self.subscribe_event("before_style_font_effect", self._fn_on_apply_style_text)
        self.subscribe_event(GblNamedEvent.EVENT_CANCELED, self._fn_on_global_cancel)

    def _on_apply_style_text(self, source: Any, args: CancelEventArgs) -> None:
        fo_strs = self.component.getText()
        if fo_strs:
            args.event_data["component"] = fo_strs[0]
        else:
            args.cancel = True

    def _on_global_cancel(self, source: Any, args: CancelEventArgs) -> None:
        initial_event = args.get("initial_event", "")

        if initial_event == "before_style_font_effect":
            args.handled = True

    # endregion Events

    # region StylePartial Overrides

    def apply_styles(self, *styles: StyleT, **kwargs) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        mChart2.Chart2._style_title(self.component, styles)

    # endregion

    @property
    def owner(self) -> _T:
        """Chart Document"""
        return self._owner

    @property
    def rotation(self) -> Angle:
        """Gets or sets the rotation angle of the title."""
        return Angle(self.get_property("TextRotation", 0))

    @rotation.setter
    def rotation(self, value: Angle | int) -> None:
        rotation = Angle(int(value))
        self.set_property(TextRotation=rotation.value)
