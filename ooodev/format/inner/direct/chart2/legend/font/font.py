# region Import
from __future__ import annotations
import uno
from ooo.dyn.awt.char_set import CharSetEnum
from ooo.dyn.awt.font_family import FontFamilyEnum
from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.awt.font_weight import FontWeightEnum
from ooo.dyn.table.shadow_format import ShadowFormat

from ooodev.utils.color import Color
from ooodev.units.angle import Angle as Angle
from ooodev.units.unit_obj import UnitT
from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
from ooodev.format.inner.direct.write.char.font.font import Font as CharFont
from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind

# endregion Import


class Font(CharFont):
    """
    Character Font for a chart legend.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties such as ``bold``, ``italic``, ``underline`` can be chained together.

    ..seealso::

        - :ref:`help_chart2_format_direct_legend_font`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        *,
        b: bool | None = None,
        i: bool | None = None,
        u: bool | None = None,
        bg_color: Color | None = None,
        bg_transparent: bool | None = None,
        charset: CharSetEnum | None = None,
        color: Color | None = None,
        family: FontFamilyEnum | None = None,
        name: str | None = None,
        overline: FontLine | None = None,
        rotation: int | Angle | None = None,
        shadow_fmt: ShadowFormat | None = None,
        shadowed: bool | None = None,
        size: float | UnitT | None = None,
        slant: FontSlant | None = None,
        spacing: CharSpacingKind | float | UnitT | None = None,
        strike: FontStrikeoutEnum | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        underline: FontLine | None = None,
        weight: FontWeightEnum | None = None,
        word_mode: bool | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            b (bool, optional): Shortcut to set ``weight`` to bold.
            i (bool, optional): Shortcut to set ``slant`` to italic.
            u (bool, optional): Shortcut ot set ``underline`` to underline.
            bg_color (:py:data:`~.utils.color.Color`, optional): The value of the text background color.
            bg_transparent (bool, optional): Determines if the text background color is set to transparent.
            charset (CharSetEnum, optional): The text encoding of the font.
            color (:py:data:`~.utils.color.Color`, optional): The value of the text color. Setting to ``-1`` will cause automatic color.
            family (FontFamilyEnum, optional): Font Family.
            name (str, optional): This property specifies the name of the font style.
                It may contain more than one name separated by comma.
            overline (FontLine, optional): Character overline values.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees.
                Depending on the implementation only certain values may be allowed.
            shadow_fmt: (ShadowFormat, optional): Determines the type, color, and width of the shadow.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            size (float, UnitT, optional): This value contains the size of the characters in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            slant (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
            spacing (CharSpacingKind, float, UnitT, optional): Specifies character spacing in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            strike (FontStrikeoutEnum, optional): Determines the type of the strike out of the character.
            subscript (bool, optional): Subscript option.
            superscript (bool, optional): Superscript option.
            underline (FontLine, optional): Character underline values.
            weight (FontWeightEnum, optional): The value of the font weight.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to
                white spaces.

        Returns:
            None:

        See Also:

            - :ref:`help_chart2_format_direct_legend_font`
        """
        super().__init__(
            b=b,
            i=i,
            u=u,
            bg_color=bg_color,
            bg_transparent=bg_transparent,
            charset=charset,
            color=color,
            family=family,
            name=name,
            overline=overline,
            rotation=rotation,
            shadow_fmt=shadow_fmt,
            shadowed=shadowed,
            size=size,
            slant=slant,
            spacing=spacing,
            strike=strike,
            subscript=subscript,
            superscript=superscript,
            underline=underline,
            weight=weight,
            word_mode=word_mode,
        )
