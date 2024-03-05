from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.mock import mock_g
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import font_effects_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
    from ooo.dyn.style.case_map import CaseMapEnum
    from ooo.dyn.awt.font_relief import FontReliefEnum
    from ooo.dyn.awt.font_underline import FontUnderlineEnum
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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_font_effect",
            after_event="after_style_font_effect",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

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
            FontEffectsT | None: Font Effects instance or ``None`` if cancelled.

        Hint:
            - ``CaseMapEnum`` can be imported from ``ooo.dyn.style.case_map``
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.awt.font_relief``
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``
            - ``FontLine`` can be imported from ``ooodev.format.inner.direct.write.char.font.font_effects``
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
            - ``FontUnderlineEnum`` can be imported from ``ooo.dyn.awt.font_underline``
        """
        factory = font_effects_factory
        kwargs: Dict[str, Any] = {
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
        }
        return self.__styler.style(factory=factory, **kwargs)

    def style_font_effect_line(
        self,
        line: FontUnderlineEnum | None = None,
        color: Color | None = None,
        overline: bool = False,
    ) -> FontEffectsT | None:
        """
        Style Font Underline or Overline.

        This method is a subset of ``style_font_effect()`` method for convenience.

        Args:
            color (:py:data:`~.utils.color.Color`, optional): The value of the text color.
                If value is ``-1`` the automatic color is applied.
            line (FontUnderlineEnum, optional): Font Line kind.
            overline (bool, optional): If ``True`` the line is overline, otherwise it is underline.

        Raises:
            CancelEventError: If the event ``before_style_font_effect`` is cancelled and not handled.

        Returns:
            FontEffectsT | None: Font Effects instance or ``None`` if cancelled.

        Hint:
            - ``FontUnderlineEnum`` can be imported from ``ooo.dyn.awt.font_underline``
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine

        factory = font_effects_factory
        fl = FontLine(line=line, color=color)
        kwargs: Dict[str, Any] = {"color": color}
        if overline:
            kwargs["overline"] = fl
        else:
            kwargs["underline"] = fl

        return self.__styler.style(factory=factory, **kwargs)

    def style_font_effect_get(self) -> FontEffectsT | None:
        """
        Gets the font effect Style.

        Raises:
            CancelEventError: If the event ``before_style_font_effect_get`` is cancelled and not handled.

        Returns:
            FontEffectsT | None: Font Effect style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=font_effects_factory)


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
