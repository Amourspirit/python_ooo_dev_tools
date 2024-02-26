from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol

    from typing_extensions import Self
    from ooo.dyn.awt.char_set import CharSetEnum
    from ooo.dyn.awt.font_family import FontFamilyEnum
    from ooo.dyn.awt.font_slant import FontSlant
    from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
    from ooo.dyn.awt.font_underline import FontUnderlineEnum
    from ooo.dyn.awt.font_weight import FontWeightEnum
    from ooo.dyn.table.shadow_format import ShadowFormat

    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
    from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind
    from ooodev.units.angle import Angle
    from ooodev.units.unit_pt import UnitPT
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.color import Color
else:
    Protocol = object
    Self = Any
    CharSetEnum = Any
    FontFamilyEnum = Any
    FontSlant = Any
    FontStrikeoutEnum = Any
    FontUnderlineEnum = Any
    FontWeightEnum = Any
    ShadowFormat = Any
    FontLine = Any
    CharSpacingKind = Any
    Angle = Any
    UnitPT = Any
    UnitT = Any
    Color = Any


class FontT(StyleT, Protocol):
    """Font Protocol"""

    def __init__(
        self,
        *,
        b: bool | None = ...,
        i: bool | None = ...,
        u: bool | None = ...,
        bg_color: Color | None = ...,
        bg_transparent: bool | None = ...,
        charset: CharSetEnum | None = ...,
        color: Color | None = ...,
        family: FontFamilyEnum | None = ...,
        name: str | None = ...,
        overline: FontLine | None = ...,
        rotation: int | Angle | None = ...,
        shadow_fmt: ShadowFormat | None = ...,
        shadowed: bool | None = ...,
        size: float | UnitT | None = ...,
        slant: FontSlant | None = ...,
        spacing: CharSpacingKind | float | UnitT | None = ...,
        strike: FontStrikeoutEnum | None = ...,
        subscript: bool | None = ...,
        superscript: bool | None = ...,
        underline: FontLine | None = ...,
        weight: FontWeightEnum | float | None = ...,
        word_mode: bool | None = ...,
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
            weight (FontWeightEnum, float, optional): The value of the font weight.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to
                white spaces.

        Returns:
            None:
        """
        ...

    # region Format Methods

    def fmt_bg_color(self, value: Color | None = None) -> Self:
        """
        Get copy of instance with text background color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text background color.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_bg_transparent(self, value: bool | None = None) -> Self:
        """
        Get copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_charset(self, value: CharSetEnum | None = None) -> Self:
        """
        Gets a copy of instance with charset set or removed.

        Args:
            value (CharSetEnum, optional): The text encoding of the font.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_color(self, value: Color | None = None) -> Self:
        """
        Get copy of instance with text color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_family(self, value: FontFamilyEnum | None = None) -> Self:
        """
        Gets a copy of instance with charset set or removed.

        Args:
            value (FontFamilyEnum, optional): The Font Family.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_size(self, value: float | UnitT | None = None) -> Self:
        """
        Get copy of instance with text size set or removed.

        Args:
            value (float, UnitT, optional): The size of the characters in ``pt`` (point) units :ref:`proto_unit_obj`.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_name(self, value: str | None = None) -> Self:
        """
        Get copy of instance with name set or removed.

        Args:
            value (str, optional): The name of the font style. It may contain more than one name separated by comma.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_overline(self, value: FontUnderlineEnum | None = None) -> Self:
        """
        Get copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_overline_color(self, value: Color | None = None) -> Self:
        """
        Get copy of instance with text overline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_rotation(self, value: int | Angle | None = None) -> Self:
        """
        Get copy of instance with rotation set or removed.

        Args:
            value (int, Angle, optional): The rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_slant(self, value: FontSlant | None = None) -> Self:
        """
        Get copy of instance with slant set or removed.

        Args:
            value (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed

        Note:
            This method changes or removes any italic settings.
        """
        ...

    def fmt_spacing(self, value: float | UnitT | None = None) -> Self:
        """
        Get copy of instance with spacing set or removed.

        Args:
            value (float, UnitT, optional): The character spacing in ``pt`` (point) units :ref:`proto_unit_obj`.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_shadow_fmt(self, value: ShadowFormat | None = None) -> Self:
        """
        Get copy of instance with shadow format set or removed.

        Args:
            value (ShadowFormat, optional): The type, color, and width of the shadow.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_strike(self, value: FontStrikeoutEnum | None = None) -> Self:
        """
        Get copy of instance with strike set or removed.

        Args:
            value (FontStrikeoutEnum, optional): The type of the strike out of the character.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_subscript(self, value: bool | None = None) -> Self:
        """
        Get copy of instance with text subscript set or removed.

        Args:
            value (bool, optional): The subscript.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_superscript(self, value: bool | None = None) -> Self:
        """
        Get copy of instance with text superscript set or removed.

        Args:
            value (bool, optional): The superscript.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_underline(self, value: FontUnderlineEnum | None = None) -> Self:
        """
        Get copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_underline_color(self, value: Color | None = None) -> Self:
        """
        Gets copy of instance with text underline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    def fmt_weight(self, value: FontWeightEnum | None = None) -> Self:
        """
        Get copy of instance with weight set or removed.

        Args:
            value (FontWeightEnum, optional): The value of the font weight.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed

        Note:
            This method changes or removes any bold settings.
        """
        ...

    def fmt_word_mode(self, value: bool | None = None) -> Self:
        """
        Get copy of instance with word mode set or removed.

        The underline and strike-through properties are not applied to white spaces when set to ``True``.

        Args:
            value (bool, optional): The word mode.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ...

    # endregion Format Methods

    # region Style Properties

    @property
    def bold(self) -> Self:
        """Gets copy of instance with bold set"""
        ...

    @property
    def italic(self) -> Self:
        """Gets copy of instance with italic set"""
        ...

    @property
    def underline(self) -> Self:
        """Gets copy of instance with underline set"""
        ...

    @property
    def bg_transparent(self) -> Self:
        """Gets copy of instance with background transparent set"""
        ...

    @property
    def overline(self) -> Self:
        """Gets copy of instance with overline set"""
        ...

    @property
    def shadowed(self) -> Self:
        """Gets copy of instance with shadow set"""
        ...

    @property
    def strike(self) -> Self:
        """Gets copy of instance with strike set"""
        ...

    @property
    def subscript(self) -> Self:
        """Gets copy of instance with sub script set"""
        ...

    @property
    def superscript(self) -> Self:
        """Gets copy of instance with super script set"""
        ...

    @property
    def word_mode(self) -> Self:
        """Gets copy of instance with word mode set"""
        ...

    @property
    def spacing_very_tight(self) -> Self:
        """Gets copy of instance with spacing set to very tight value"""
        ...

    @property
    def spacing_tight(self) -> Self:
        """Gets copy of instance with spacing set to tight value"""
        ...

    @property
    def spacing_normal(self) -> Self:
        """Gets copy of instance with spacing set to normal value"""
        ...

    @property
    def spacing_loose(self) -> Self:
        """Gets copy of instance with spacing set to loose value"""
        ...

    @property
    def spacing_very_loose(self) -> Self:
        """Gets copy of instance with spacing set to very loose value"""
        ...

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_is_bold(self) -> bool:
        """Specifies bold"""
        ...

    @prop_is_bold.setter
    def prop_is_bold(self, value: bool | None) -> None: ...

    @property
    def prop_bg_color(self) -> Color | None:
        """This property contains the text background color."""
        ...

    @prop_bg_color.setter
    def prop_bg_color(self, value: Color | None) -> None: ...

    @property
    def prop_bg_color_transparent(self) -> bool | None:
        """This property contains the text background color."""
        ...

    @prop_bg_color_transparent.setter
    def prop_bg_color_transparent(self, value: bool | None) -> None: ...

    @property
    def prop_is_italic(self) -> bool | None:
        """Specifies italic"""
        ...

    @prop_is_italic.setter
    def prop_is_italic(self, value: bool | None) -> None: ...

    @property
    def prop_is_underline(self) -> bool | None:
        """Specifies underline"""
        ...

    @prop_is_underline.setter
    def prop_is_underline(self, value: bool | None) -> None: ...

    @property
    def prop_underline(self) -> FontLine:
        """This property contains the value for the character underline."""
        ...

    @prop_underline.setter
    def prop_underline(self, value: FontLine | None) -> None: ...

    @property
    def prop_charset(self) -> CharSetEnum | None:
        """This property contains the text encoding of the font."""
        ...

    @prop_charset.setter
    def prop_charset(self, value: CharSetEnum | None) -> None: ...

    @property
    def prop_color(self) -> Color | None:
        """This property contains the value of the text color."""
        ...

    @prop_color.setter
    def prop_color(self, value: Color | None) -> None: ...

    @property
    def prop_family(self) -> FontFamilyEnum | None:
        """This property contains font family."""
        ...

    @prop_family.setter
    def prop_family(self, value: FontFamilyEnum | None) -> None: ...

    @property
    def prop_size(self) -> UnitPT | None:
        """This value contains the size of the characters in ``pt`` (point) units."""
        ...

    @prop_size.setter
    def prop_size(self, value: float | UnitT | None) -> None: ...

    @property
    def prop_name(self) -> str | None:
        """This property specifies the name of the font style. It may contain more than one name separated by comma."""
        ...

    @prop_name.setter
    def prop_name(self, value: str | None) -> None: ...

    @property
    def prop_strike(self) -> FontStrikeoutEnum | None:
        """This property determines the type of the strike out of the character."""
        ...

    @prop_strike.setter
    def prop_strike(self, value: FontStrikeoutEnum | None) -> None: ...

    @property
    def prop_weight(self) -> FontWeightEnum | None:
        """This property contains the value of the font weight."""
        ...

    @prop_weight.setter
    def prop_weight(self, value: FontWeightEnum | None) -> None: ...

    @property
    def prop_slant(self) -> FontSlant | None:
        """This property contains the value of the posture of the document such as  ``FontSlant.ITALIC``"""
        ...

    @prop_slant.setter
    def prop_slant(self, value: FontSlant | None) -> None: ...

    @property
    def prop_spacing(self) -> UnitPT | None:
        """This value contains character spacing in ``pt`` (point) units"""
        ...

    @prop_spacing.setter
    def prop_spacing(self, value: float | CharSpacingKind | UnitT | None) -> None: ...

    @property
    def prop_shadowed(self) -> bool | None:
        """This property specifies if the characters are formatted and displayed with a shadow effect."""
        ...

    @prop_shadowed.setter
    def prop_shadowed(self, value: bool | None) -> None: ...

    @property
    def prop_shadow_fmt(self) -> ShadowFormat | None:
        """This property specifies the type, color, and width of the shadow."""
        ...

    @prop_shadow_fmt.setter
    def prop_shadow_fmt(self, value: ShadowFormat | None) -> None: ...

    @property
    def prop_superscript(self) -> bool | None:
        """Specifies if the font is super script"""
        ...

    @prop_superscript.setter
    def prop_superscript(self, value: bool | None) -> None: ...

    @property
    def prop_subscript(self) -> bool | None:
        """Specifies if the font is sub script"""
        ...

    @prop_subscript.setter
    def prop_subscript(self, value: bool | None) -> None: ...

    @property
    def prop_overline(self) -> FontLine:
        """This property contains the value for the character overline."""
        ...

    @prop_overline.setter
    def prop_overline(self, value: FontLine | None) -> None: ...

    @property
    def prop_rotation(self) -> Angle | None:
        """
        This optional property determines the rotation of a character in degrees.

        Depending on the implementation only certain values may be allowed.
        """
        ...

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle | None) -> None: ...

    @property
    def prop_word_mode(self) -> bool | None:
        """If this property is ``True``, the underline and strike-through properties are not applied to white spaces."""
        ...

    @prop_word_mode.setter
    def prop_word_mode(self, value: bool | None) -> None: ...

    # endregion Prop Properties
