from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.format.inner.style_factory import font_effects_factory
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent

if TYPE_CHECKING:
    from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
    from ooo.dyn.style.case_map import CaseMapEnum
    from ooo.dyn.awt.font_relief import FontReliefEnum
    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
    from ooodev.utils.color import Color
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.font.font_effects_t import FontEffectsT


class FontEffectsPartial:
    """
    Partial class for FontEffects.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def style_font_effect(
        self,
        *,
        color: Color | None = None,
        transparency: Intensity | int | None = None,
        overline: FontLine | None = None,
        underline: FontLine | None = None,
        strike: FontStrikeoutEnum | None = None,
        word_mode: bool | None = None,
        case: CaseMapEnum | None = None,
        relief: FontReliefEnum | None = None,
        outline: bool | None = None,
        hidden: bool | None = None,
        shadowed: bool | None = None,
    ) -> FontEffectsT | None:
        """
        Style Font options.

        Args:
            color (:py:data:`~.utils.color.Color`, optional): The value of the text color.
                If value is ``-1`` the automatic color is applied.
            transparency (Intensity, int, optional): The transparency value from ``0`` to ``100`` for the font color.
            overline (FontLine, optional): Character overline values.
            underline (FontLine, optional): Character underline values.
            strike (FontStrikeoutEnum, optional): Determines the type of the strike out of the character.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied
                to white spaces.
            case (CaseMapEnum, optional): Specifies the case of the font.
            relief (FontReliefEnum, optional): Specifies the relief of the font.
            outline (bool, optional): Specifies if the font is outlined.
            hidden (bool, optional): Specifies if the font is hidden.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.

        Raises:
            CancelEventError: If the event ``before_style_font_effect`` is cancelled and not handled.

        Returns:
            None:
        """
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_font_effect.__qualname__)
            event_data: Dict[str, Any] = {
                "color": color,
                "transparency": transparency,
                "overline": overline,
                "underline": underline,
                "strike": strike,
                "word_mode": word_mode,
                "case": case,
                "relief": relief,
                "outline": outline,
                "hidden": hidden,
                "shadowed": shadowed,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_font_effect", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_font_effect")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                    else:
                        return None
                else:
                    return None
            color = cargs.event_data.get("color", color)
            transparency = cargs.event_data.get("transparency", transparency)
            overline = cargs.event_data.get("overline", overline)
            underline = cargs.event_data.get("underline", underline)
            strike = cargs.event_data.get("strike", strike)
            word_mode = cargs.event_data.get("word_mode", word_mode)
            case = cargs.event_data.get("case", case)
            relief = cargs.event_data.get("relief", relief)
            outline = cargs.event_data.get("outline", outline)
            hidden = cargs.event_data.get("hidden", hidden)
            shadowed = cargs.event_data.get("shadowed", shadowed)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = font_effects_factory(factory_name)
        fe = styler(
            color=color,
            transparency=transparency,
            overline=overline,
            underline=underline,
            strike=strike,
            word_mode=word_mode,
            case=case,
            relief=relief,
            outline=outline,
            hidden=hidden,
            shadowed=shadowed,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_font_effect", EventArgs.from_args(cargs))  # type: ignore
        return fe
