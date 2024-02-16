from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import calc_align_orientation_factory

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.inner.direct.calc.alignment.text_orientation import EdgeKind
    from ooodev.format.proto.calc.alignment.text_orientation_t import TextOrientationT
    from ooodev.units import Angle
else:
    TextAlignT = Any
    LoInst = Any
    EdgeKind = Any
    Angle = Any


class TextOrientationPartial:
    """
    Partial class for Alignment Text Orientation.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_align_orientation"
        self.__styler.before_event_name = "before_style_align_orientation"

    def style_align_orientation(
        self, vert_stack: bool | None = None, rotation: int | Angle | None = None, edge: EdgeKind | None = None
    ) -> TextOrientationT | None:
        """
        Style Text Orientation.

        Args:
            vert_stack (bool, optional): Specifies if vertical stack is to be used.
            rotation (int, Angle, optional): Specifies if the rotation.
            edge (EdgeKind, optional): Specifies the Reference Edge.

        Raises:
            CancelEventError: If the event ``before_style_align_orientation`` is cancelled and not handled.

        Returns:
            TextOrientationT | None: Text Alignment instance or ``None`` if cancelled.

        Hint:
            - ``EdgeKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.text_orientation``
        """
        factory = calc_align_orientation_factory
        kwargs = {"vert_stack": vert_stack, "rotation": rotation, "edge": edge}
        return self.__styler.style(factory=factory, **kwargs)

    def style_align_orientation_get(self) -> TextOrientationT | None:
        """
        Gets the Alignment Text Orientation Style.

        Raises:
            CancelEventError: If the event ``before_style_align_orientation_get`` is cancelled and not handled.

        Returns:
            TextOrientationT | None: Text Alignment style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=calc_align_orientation_factory)
