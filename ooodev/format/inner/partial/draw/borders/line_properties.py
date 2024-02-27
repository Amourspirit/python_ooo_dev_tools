from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.preset.preset_border_line import BorderLineKind
from ooodev.format.inner.style_factory import draw_border_line_factory
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.utils import color as mColor
from ooodev.utils.context.lo_context import LoContext
from ooodev.format.inner.partial.factory_name_base import FactoryNameBase

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.borders.line_properties_t import LinePropertiesT
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.data_type.intensity import Intensity
else:
    LoInst = Any
    LinePropertiesT = Any
    UnitT = Any
    Intensity = Any


class LineProperties(FactoryNameBase):
    """
    Class for Line Properties.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        super().__init__(factory_name, component, lo_inst)
        self.before_event_name = "before_style_border_line"
        self.after_event_name = "after_style_border_line"

    def style(
        self,
        color: mColor.Color = mColor.Color(0),
        width: float | UnitT = 0,
        transparency: int | Intensity = 0,
        style: BorderLineKind = BorderLineKind.CONTINUOUS,
    ) -> LinePropertiesT | None:
        """
        Style Font.

        Args:
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float | UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int | Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.
            style (BorderLineKind, optional): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.

        Raises:
            CancelEventError: If the event ``before_style_border_line`` is cancelled and not handled.

        Returns:
            LinePropertiesT | None: Font Only instance or ``None`` if cancelled.

        Hint:
            - ``BorderLineKind`` can be imported from ``ooodev.format.inner.preset.preset_border_line``
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        comp = self._component
        factory_name = self._factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style.__qualname__)
            event_data: Dict[str, Any] = {
                "style": style,
                "color": color,
                "width": width,
                "transparency": transparency,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event(self.before_event_name, cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", self.before_event_name)
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                else:
                    return None
            style = cargs.event_data.get("style", style)
            color = cargs.event_data.get("color", color)
            width = cargs.event_data.get("width", width)
            transparency = cargs.event_data.get("transparency", transparency)
            comp = cargs.event_data.get("this_component", comp)
            factory_name = cargs.event_data.get("factory_name", factory_name)

        styler = draw_border_line_factory(factory_name)
        fe = styler(
            style=style,
            color=color,
            width=width,
            transparency=transparency,
        )
        # fe.factory_name = factory_name

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self._lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            event_args.event_data["styler_object"] = fe
            self.trigger_event(self.after_event_name, EventArgs.from_args(cargs))  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore
        return fe

    def style_get(self) -> LinePropertiesT | None:
        """
        Gets the Style.

        Raises:
            CancelEventError: If the event is cancelled and not handled.

        Returns:
            LinePropertiesT | None: Line style or ``None`` if cancelled.
        """
        comp = self._component
        factory_name = self._factory_name
        cargs = None
        event_name = f"{self.before_event_name}_get"
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_get.__qualname__)
            event_data: Dict[str, Any] = {
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(event_name, cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", event_name)
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = draw_border_line_factory(factory_name)
        try:
            style = styler.from_obj(comp)
        except mEx.DisabledMethodError:
            return None
        style.set_update_obj(comp)
        return style
