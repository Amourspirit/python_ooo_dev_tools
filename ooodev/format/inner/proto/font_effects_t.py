from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Type
import uno
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.style.case_map import CaseMapEnum
from ooo.dyn.awt.font_relief import FontReliefEnum

from .sytle_t import StyleT
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity as Intensity

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
    from ooo.dyn.awt.font_underline import FontUnderlineEnum
else:
    Protocol = object


class FontEffectsT(StyleT, Protocol):
    """Font Effect Protocol"""

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
    ) -> None: ...

    @overload
    @classmethod
    def from_obj(cls: Type[FontEffectsT], obj: Any) -> FontEffectsT: ...

    @overload
    @classmethod
    def from_obj(cls: Type[FontEffectsT], obj: Any, **kwargs) -> FontEffectsT: ...

    # region Format Methods
    def fmt_color(self: FontEffectsT, value: Color | None = None) -> FontEffectsT:
        """
        Gets copy of instance with text color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_transparency(self: FontEffectsT, value: bool | None = None) -> FontEffectsT:
        """
        Gets copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_overline(self: FontEffectsT, value: FontUnderlineEnum | None = None) -> FontEffectsT:
        """
        Gets copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_overline_color(self: FontEffectsT, value: Color | None = None) -> FontEffectsT:
        """
        Gets copy of instance with text overline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_strike(self: FontEffectsT, value: FontStrikeoutEnum | None = None) -> FontEffectsT:
        """
        Gets copy of instance with strike set or removed.

        Args:
            value (FontStrikeoutEnum, optional): The type of the strike out of the character.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_underline(self: FontEffectsT, value: FontUnderlineEnum | None = None) -> FontEffectsT:
        """
        Gets copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_underline_color(self: FontEffectsT, value: Color | None = None) -> FontEffectsT:
        """
        Gets copy of instance with text underline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_word_mode(self: FontEffectsT, value: bool | None = None) -> FontEffectsT:
        """
        Gets copy of instance with word mode set or removed.

        The underline and strike-through properties are not applied to white spaces when set to ``True``.

        Args:
            value (bool, optional): The word mode.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_case(self: FontEffectsT, value: CaseMapEnum | None = None) -> FontEffectsT:
        """
        Gets copy of instance with case set or removed.

        Args:
            value (CaseMapEnum, optional): The case value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_relief(self: FontEffectsT, value: FontReliefEnum | None = None) -> FontEffectsT:
        """
        Gets copy of instance with relief set or removed.

        Args:
            value (FontReliefEnum, optional): The relief value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_outline(self: FontEffectsT, value: bool | None = None) -> FontEffectsT:
        """
        Gets copy of instance with outline set or removed.

        Args:
            value (bool, optional): The outline value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_hidden(self: FontEffectsT, value: bool | None = None) -> FontEffectsT:
        """
        Gets copy of instance with hidden set or removed.

        Args:
            value (bool, optional): The hidden value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    def fmt_shadowed(self: FontEffectsT, value: bool | None = None) -> FontEffectsT:
        """
        Gets copy of instance with shadowed set or removed.

        Args:
            value (bool, optional): The shadowed value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ...

    # endregion Format Methods

    # region Style Properties
    @property
    def color_auto(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with color set to automatic"""
        ...

    @property
    def overline(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with overline set"""
        ...

    @property
    def overline_color_auto(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with overline color set to automatic"""
        ...

    @property
    def underline(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with underline set"""
        ...

    @property
    def under_color_auto(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with underline color set to automatic"""
        ...

    @property
    def strike(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with strike set"""
        ...

    @property
    def word_mode(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with word mode set"""
        ...

    @property
    def outline(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with outline set"""
        ...

    @property
    def hidden(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with hidden set"""
        ...

    @property
    def shadowed(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with shadow set"""
        ...

    @property
    def case_upper(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with case set to upper"""
        ...

    @property
    def case_lower(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with case set to lower"""
        ...

    @property
    def case_title(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with case set to title"""
        ...

    @property
    def case_small_caps(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with case set to small caps"""
        ...

    @property
    def case_none(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with no case set"""
        ...

    @property
    def relief_none(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with no relief set"""
        ...

    @property
    def relief_embossed(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with relief set to embossed"""
        ...

    @property
    def relief_engraved(self: FontEffectsT) -> FontEffectsT:
        """Gets copy of instance with relief set to engraved"""
        ...

    # endregion Style Properties

    # region Prop Properties

    @property
    def prop_color(self) -> Color | None:
        """Gets/Sets the value of the text color."""
        ...

    @prop_color.setter
    def prop_color(self, value: Color | None) -> None: ...

    @property
    def prop_transparency(self) -> Intensity | None:
        """Gets/Sets The transparency value"""
        ...

    @prop_transparency.setter
    def prop_transparency(self, value: Intensity | int | None) -> None: ...

    @property
    def prop_overline(self) -> FontLine:
        """This property contains the value for the character overline."""
        ...

    @prop_overline.setter
    def prop_overline(self, value: FontLine | None) -> None: ...

    @property
    def prop_underline(self) -> FontLine:
        """This property contains the value for the character underline."""
        ...

    @prop_underline.setter
    def prop_underline(self, value: FontLine | None) -> None: ...

    @property
    def prop_strike(self) -> FontStrikeoutEnum | None:
        """Gets/Sets the type of the strike out of the character."""
        ...

    @prop_strike.setter
    def prop_strike(self, value: FontStrikeoutEnum | None) -> None: ...

    @property
    def prop_word_mode(self) -> bool | None:
        """Gets/Sets Character word mode. If this property is ``True``, the underline and strike-through properties are not applied to white spaces."""
        ...

    @prop_word_mode.setter
    def prop_word_mode(self, value: bool | None) -> None: ...

    @property
    def prop_case(self) -> CaseMapEnum | None:
        """Gets/Sets Font Case Value"""
        ...

    @prop_case.setter
    def prop_case(self, value: CaseMapEnum | None) -> None: ...

    @property
    def prop_relief(self) -> FontReliefEnum | None:
        """Gets/Sets Font Relief Value"""
        ...

    @prop_relief.setter
    def prop_relief(self, value: FontReliefEnum | None) -> None: ...

    @property
    def prop_outline(self) -> bool | None:
        """Gets/Sets if the font is outlined"""
        ...

    @prop_outline.setter
    def prop_outline(self, value: bool | None) -> None: ...

    @property
    def prop_hidden(self) -> bool | None:
        """Gets/Sets if the font is hidden"""
        ...

    @prop_hidden.setter
    def prop_hidden(self, value: bool | None) -> None: ...

    @property
    def prop_shadowed(self) -> bool | None:
        """Gets/Sets if the characters are formatted and displayed with a shadow effect."""
        ...

    @prop_shadowed.setter
    def prop_shadowed(self, value: bool | None) -> None: ...

    @property
    def default(self: FontEffectsT) -> FontEffectsT:  # type: ignore[misc]
        """Gets Font Position default."""
        ...

    # endregion Prop Properties
