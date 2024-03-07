from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.text.graphic_crop import GraphicCrop

from ooodev.adapter.struct_base import StructBase

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class GraphicCropStructComp(StructBase[GraphicCrop]):
    """
    GraphicCrop Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_text_GraphicCrop_changing``.
    The event raised after the property is changed is called ``com_sun_star_text_GraphicCrop_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: GraphicCrop, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (GraphicCrop): Graphic Crop.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_text_GraphicCrop_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_text_GraphicCrop_changed"

    def _copy(self, src: GraphicCrop | None = None) -> GraphicCrop:
        if src is None:
            src = self.component
        return GraphicCrop(
            Top=src.Top,
            Bottom=src.Bottom,
            Left=src.Left,
            Right=src.Right,
        )

    # endregion Overrides
    # region Properties

    @property
    def top(self) -> int:
        """
        Gets/Sets the top value to cut (if negative) or to extend (if positive)
        """
        return self.component.Top

    @top.setter
    def top(self, value: int) -> None:
        old_value = self.component.Top
        if old_value != value:
            event_args = self._trigger_cancel_event("Top", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def bottom(self) -> int:
        """
        Gets/Sets the bottom value to cut (if negative) or to extend (if positive)
        """
        return self.component.Bottom

    @bottom.setter
    def bottom(self, value: int) -> None:
        old_value = self.component.Bottom
        if old_value != value:
            event_args = self._trigger_cancel_event("Bottom", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def left(self) -> int:
        """
        Gets/Sets the left value to cut (if negative) or to extend (if positive)
        """
        return self.component.Left

    @left.setter
    def left(self, value: int) -> None:
        old_value = self.component.Left
        if old_value != value:
            event_args = self._trigger_cancel_event("Left", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def right(self) -> int:
        """
        Gets/Sets the right value to cut (if negative) or to extend (if positive)
        """
        return self.component.Right

    @right.setter
    def right(self, value: int) -> None:
        old_value = self.component.Right
        if old_value != value:
            event_args = self._trigger_cancel_event("Right", old_value, value)
            self._trigger_done_event(event_args)

    # endregion Properties
