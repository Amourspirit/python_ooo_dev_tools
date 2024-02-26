from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial as StandardFontEffectsPartial


if TYPE_CHECKING:
    from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
    from ooo.dyn.style.case_map import CaseMapEnum
    from ooo.dyn.awt.font_relief import FontReliefEnum
    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
    from ooodev.utils.color import Color
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.format.proto.font.font_effects_t import FontEffectsT


class FontEffectsPartial(StandardFontEffectsPartial):

    # region StandardFontEffectsPartial Overrides
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
        # calc does not have a hidden property for font effects.
        kwargs = {
            "color": color,
            "transparency": transparency,
            "overline": overline,
            "underline": underline,
            "strike": strike,
            "word_mode": word_mode,
            "case": case,
            "relief": relief,
            "outline": outline,
            "shadowed": shadowed,
        }
        return super().style_font_effect(**kwargs)

    # endregion StandardFontEffectsPartial Overrides
