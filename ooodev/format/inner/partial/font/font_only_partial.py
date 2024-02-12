from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.format.inner.style_factory import font_only_factory
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.font.font_lang_t import FontLangT
    from ooodev.format.proto.font.font_only_t import FontOnlyT
    from ooodev.units import UnitT
else:
    FontLangT = Any
    FontOnlyT = Any
    UnitT = Any


class FontOnlyPartial:
    """
    Partial class for FontEffects.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def style_font(
        self,
        name: str | None = None,
        size: float | UnitT | None = None,
        font_style: str | None = None,
        lang: FontLangT | None = None,
    ) -> FontOnlyT | None:
        """
        Style Font.

        Args:
            name (str | None, optional): Font Name.
            size (float | UnitT | None, optional): Font Size in ``PT`` units or ``UnitT``.
            font_style (str | None, optional): Font Style such as ``Bold Italics``
            lang (FontLangT | None, optional): Font Language.

        Raises:
            CancelEventError: If the event ``before_style_font_only`` is cancelled and not handled.

        Returns:
            FontOnlyT | None: Font Only instance or ``None`` if cancelled.

        See Also:
            :py:class:`~ooodev.format.inner.direct.write.char.font.font_only.FontLang`
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_font.__qualname__)
            event_data: Dict[str, Any] = {
                "name": name,
                "size": size,
                "font_style": font_style,
                "lang": lang,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_font_only", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_font_only")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                    else:
                        return None
                else:
                    return None
            name = cargs.event_data.get("name", name)
            size = cargs.event_data.get("size", size)
            font_style = cargs.event_data.get("font_style", font_style)
            lang = cargs.event_data.get("lang", lang)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = font_only_factory(factory_name)
        fe = styler(
            name=name,
            size=size,
            font_style=font_style,
            lang=lang,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_font_only", EventArgs.from_args(cargs))  # type: ignore
        return fe
