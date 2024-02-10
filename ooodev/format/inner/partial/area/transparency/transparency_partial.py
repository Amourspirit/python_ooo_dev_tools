from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import area_transparency_transparency_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.format.proto.area.transparency.transparency_t import TransparencyT
else:
    LoInst = Any
    TransparencyT = Any
    Intensity = Any


class TransparencyPartial:
    """
    Partial class for FillColor.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def style_area_transparency_transparency(self, value: Intensity | int = 0) -> TransparencyT | None:
        """
        Style Area Color.

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Raises:
            CancelEventError: If the event ``before_style_area_transparency_gradient`` is cancelled and not handled.

        Returns:
            TransparencyT | None: FillColor instance or ``None`` if cancelled.
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_transparency_transparency.__qualname__)
            event_data: Dict[str, Any] = {
                "value": value,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_transparency_gradient", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_transparency_gradient")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                    else:
                        return None
                else:
                    return None
            value = cargs.event_data.get("value", value)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = area_transparency_transparency_factory(factory_name)
        fe = styler(value=value)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_transparency_gradient", EventArgs.from_args(cargs))  # type: ignore
        return fe
