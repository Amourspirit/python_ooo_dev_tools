from __future__ import annotations
import uno
from ooo.dyn.awt.font_relief import FontReliefEnum
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.style.case_map import CaseMapEnum

from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects as CharFontEffects
from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity

from ooodev.meta.deleted_attrib import DeletedAttrib
from ooodev.meta.disabled_method import DisabledMethod


class FontEffects(CharFontEffects):
    """
    Calc Character Font Effects

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties can be chained together.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_font_effects`

    .. versionadded:: 0.9.4
    """

    prop_hidden = DeletedAttrib()  # type: ignore
    fmt_hidden = DisabledMethod()  # type: ignore

    def __init__(
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
    ) -> None:
        """
        Font options used in styles.

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
            hidden (bool, optional): This property is Ignored in Calc. Included for compatibility.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_direct_cell_font_effects`
        """
        hidden = None  # Included for compatibility with style_factory.font_effects_factory()
        super().__init__(
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
