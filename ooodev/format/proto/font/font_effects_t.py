from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Type
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
    from ooo.dyn.awt.font_underline import FontUnderlineEnum
    from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
    from ooo.dyn.style.case_map import CaseMapEnum
    from ooo.dyn.awt.font_relief import FontReliefEnum
    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
    from ooodev.utils.color import Color
    from ooodev.utils.data_type.intensity import Intensity
else:
    Protocol = object
    Self = Any
    FontUnderlineEnum = Any
    FontStrikeoutEnum = Any
    CaseMapEnum = Any
    FontReliefEnum = Any
    FontLine = Any
    Color = Any
    Intensity = Any


class FontEffectsT(StyleT, Protocol):
    """Font Effect Protocol"""

    def __init__(
        self,
        *,
        color: Color | None = ...,
        transparency: Intensity | int | None = ...,
        overline: FontLine | None = ...,
        underline: FontLine | None = ...,
        strike: FontStrikeoutEnum | None = ...,
        word_mode: bool | None = ...,
        case: CaseMapEnum | None = ...,
        relief: FontReliefEnum | None = ...,
        outline: bool | None = ...,
        hidden: bool | None = ...,
        shadowed: bool | None = ...,
    ) -> None: ...

    @overload
    @classmethod
    def from_obj(cls: Type[FontEffectsT], obj: Any) -> FontEffectsT: ...

    @overload
    @classmethod
    def from_obj(cls: Type[FontEffectsT], obj: Any, **kwargs) -> FontEffectsT: ...

    # region Format Methods
    def fmt_color(self, value: Color | None = None) -> FontEffectsT:
        """
        Gets copy of instance with text color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_transparency(self, value: bool | None = None) -> Self:
        """
        Gets copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_overline(self, value: FontUnderlineEnum | None = None) -> Self:
        """
        Gets copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed.

        Hint:
            - ``FontUnderlineEnum`` can be imported from ``ooo.dyn.awt.font_underline``
        """
        ...

    def fmt_overline_color(self, value: Color | None = None) -> Self:
        """
        Gets copy of instance with text overline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_strike(self, value: FontStrikeoutEnum | None = None) -> Self:
        """
        Gets copy of instance with strike set or removed.

        Args:
            value (FontStrikeoutEnum, optional): The type of the strike out of the character.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed.

        Hint:
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``
        """
        ...

    def fmt_underline(self, value: FontUnderlineEnum | None = None) -> Self:
        """
        Gets copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed.

        Hint:
            - ``FontUnderlineEnum`` can be imported from ``ooo.dyn.awt.font_underline``
        """
        ...

    def fmt_underline_color(self, value: Color | None = None) -> Self:
        """
        Gets copy of instance with text underline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_word_mode(self, value: bool | None = None) -> Self:
        """
        Gets copy of instance with word mode set or removed.

        The underline and strike-through properties are not applied to white spaces when set to ``True``.

        Args:
            value (bool, optional): The word mode.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_case(self, value: CaseMapEnum | None = None) -> Self:
        """
        Gets copy of instance with case set or removed.

        Args:
            value (CaseMapEnum, optional): The case value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed.

        Hint:
            - ``CaseMapEnum`` can be imported from ``ooo.dyn.style.case_map``
        """
        ...

    def fmt_relief(self, value: FontReliefEnum | None = None) -> Self:
        """
        Gets copy of instance with relief set or removed.

        Args:
            value (FontReliefEnum, optional): The relief value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed.

        Hint:
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.awt.font_relief``
        """
        ...

    def fmt_outline(self, value: bool | None = None) -> Self:
        """
        Gets copy of instance with outline set or removed.

        Args:
            value (bool, optional): The outline value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_hidden(self, value: bool | None = None) -> Self:
        """
        Gets copy of instance with hidden set or removed.

        Args:
            value (bool, optional): The hidden value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    def fmt_shadowed(self, value: bool | None = None) -> Self:
        """
        Gets copy of instance with shadowed set or removed.

        Args:
            value (bool, optional): The shadowed value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffectsT: Font with style added or removed
        """
        ...

    # endregion Format Methods

    # region Style Properties
    @property
    def color_auto(self) -> FontEffectsT:
        """Gets copy of instance with color set to automatic"""
        ...

    @property
    def overline(self) -> FontEffectsT:
        """Gets copy of instance with overline set"""
        ...

    @property
    def overline_color_auto(self) -> FontEffectsT:
        """Gets copy of instance with overline color set to automatic"""
        ...

    @property
    def underline(self) -> FontEffectsT:
        """Gets copy of instance with underline set"""
        ...

    @property
    def under_color_auto(self) -> FontEffectsT:
        """Gets copy of instance with underline color set to automatic"""
        ...

    @property
    def strike(self) -> FontEffectsT:
        """Gets copy of instance with strike set"""
        ...

    @property
    def word_mode(self) -> FontEffectsT:
        """Gets copy of instance with word mode set"""
        ...

    @property
    def outline(self) -> FontEffectsT:
        """Gets copy of instance with outline set"""
        ...

    @property
    def hidden(self) -> FontEffectsT:
        """Gets copy of instance with hidden set"""
        ...

    @property
    def shadowed(self) -> FontEffectsT:
        """Gets copy of instance with shadow set"""
        ...

    @property
    def case_upper(self) -> FontEffectsT:
        """Gets copy of instance with case set to upper"""
        ...

    @property
    def case_lower(self) -> FontEffectsT:
        """Gets copy of instance with case set to lower"""
        ...

    @property
    def case_title(self) -> FontEffectsT:
        """Gets copy of instance with case set to title"""
        ...

    @property
    def case_small_caps(self) -> FontEffectsT:
        """Gets copy of instance with case set to small caps"""
        ...

    @property
    def case_none(self) -> FontEffectsT:
        """Gets copy of instance with no case set"""
        ...

    @property
    def relief_none(self) -> FontEffectsT:
        """Gets copy of instance with no relief set"""
        ...

    @property
    def relief_embossed(self) -> FontEffectsT:
        """Gets copy of instance with relief set to embossed"""
        ...

    @property
    def relief_engraved(self) -> FontEffectsT:
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
        """
        Gets/Sets The transparency value

        Hint:
            - The transparency value from ``0`` to ``100``.
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        ...

    @prop_transparency.setter
    def prop_transparency(self, value: Intensity | int | None) -> None: ...

    @property
    def prop_overline(self) -> FontLine:
        """
        This property contains the value for the character overline.

        Hint:
            - ``FontLine`` can be imported from ``ooodev.format.inner.direct.write.char.font.font_effects``
        """
        ...

    @prop_overline.setter
    def prop_overline(self, value: FontLine | None) -> None: ...

    @property
    def prop_underline(self) -> FontLine:
        """
        This property contains the value for the character underline.

        Hint:
            - ``FontLine`` can be imported from ``ooodev.format.inner.direct.write.char.font.font_effects``
        """
        ...

    @prop_underline.setter
    def prop_underline(self, value: FontLine | None) -> None: ...

    @property
    def prop_strike(self) -> FontStrikeoutEnum | None:
        """
        Gets/Sets the type of the strike out of the character.

        Hint:
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``
        """
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
        """
        Gets/Sets Font Case Value.

        Hint:
            - ``CaseMapEnum`` can be imported from ``ooo.dyn.style.case_map``
        """
        ...

    @prop_case.setter
    def prop_case(self, value: CaseMapEnum | None) -> None: ...

    @property
    def prop_relief(self) -> FontReliefEnum | None:
        """
        Gets/Sets Font Relief Value.

        Hint:
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.awt.font_relief``
        """
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
    def default(self) -> FontEffectsT:  # type: ignore[misc]
        """Gets Font Position default."""
        ...

    # endregion Prop Properties
