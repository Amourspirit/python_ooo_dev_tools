"""
Module for managing character fonts.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar

import uno
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum
from ooo.dyn.style.case_map import CaseMapEnum
from ooo.dyn.awt.font_relief import FontReliefEnum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils import props as mProps
from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor
from ooodev.utils.data_type.intensity import Intensity

# endregion Imports

_TFontEffects = TypeVar(name="_TFontEffects", bound="FontEffects")


class FontLine:
    """Font Line such as overline and underline."""

    def __init__(self, line: FontUnderlineEnum | None = None, color: Color | None = None) -> None:
        """
        Constructor

        Args:
            line (FontUnderlineEnum, optional): Font Line kind.
            color (:py:data:`~.utils.color.Color`, optional): Line color. If value is ``-1`` the automatic color is applied.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_char_font_effects`
        """
        self._line = line
        self._color = color

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, FontLine):
            return self.color == oth.color and self.line == oth.line
        return NotImplemented

    def has_values(self) -> bool:
        """Gets if instance has any values."""
        return self._line is not None or self._color is not None

    @property
    def line(self) -> FontUnderlineEnum | None:
        """
        Gets/Sets line.
        """
        return self._line

    @line.setter
    def line(self, value: FontUnderlineEnum | None):
        self._line = value

    @property
    def color(self) -> Color | None:
        """
        Gets/Sets color.
        """
        return self._color

    @color.setter
    def color(self, value: Color | None):
        self._color = value


class FontEffects(StyleBase):
    """
    Character Font Effects

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties can be chained together.

    .. versionadded:: 0.9.0
    """

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
            hidden (bool, optional): Specifies if the font is hidden.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_direct_char_font_effects`
        """
        # could not find any documentation in the API or elsewhere online for Overline
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html

        super().__init__()
        if transparency is not None:
            self.prop_transparency = transparency
        if color is not None:
            self.prop_color = color
        if overline is not None:
            self.prop_overline = overline
        if underline is not None:
            self.prop_underline = underline
        if strike is not None:
            self.prop_strike = strike
        if word_mode is not None:
            self.prop_word_mode = word_mode
        if case is not None:
            self.prop_case = case
        if relief is not None:
            self.prop_relief = relief
        if outline is not None:
            self.prop_outline = outline
        if hidden is not None:
            self.prop_hidden = hidden
        if shadowed is not None:
            self.prop_shadowed = shadowed

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.drawing.ControlShape",
                "com.sun.star.chart2.Legend",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    @overload
    def apply(self, obj: Any, **kwargs: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs: Any) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TFontEffects], obj: Any) -> _TFontEffects: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TFontEffects], obj: Any, **kwargs) -> _TFontEffects: ...

    @classmethod
    def from_obj(cls: Type[_TFontEffects], obj: Any, **kwargs) -> _TFontEffects:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            FontEffects: ``FontEffects`` instance that represents ``obj`` font effects.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, fe: FontEffects):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                fe._set(key, val)

        set_prop("CharColor", inst)
        set_prop("CharOverline", inst)
        set_prop("CharOverlineColor", inst)
        set_prop("CharOverlineHasColor", inst)
        set_prop("CharUnderline", inst)
        set_prop("CharUnderlineColor", inst)
        set_prop("CharUnderlineHasColor", inst)
        set_prop("CharWordMode", inst)
        set_prop("CharShadowed", inst)
        set_prop("CharContoured", inst)
        set_prop("CharHidden", inst)
        set_prop("CharTransparence", inst)
        set_prop("CharStrikeout", inst)
        set_prop("CharCaseMap", inst)
        set_prop("CharRelief", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion methods

    # region Format Methods
    def fmt_color(self: _TFontEffects, value: Color | None = None) -> _TFontEffects:
        """
        Gets copy of instance with text color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_color = value
        return ft

    def fmt_transparency(self: _TFontEffects, value: bool | None = None) -> _TFontEffects:
        """
        Gets copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_transparency = value
        return ft

    def fmt_overline(self: _TFontEffects, value: FontUnderlineEnum | None = None) -> _TFontEffects:
        """
        Gets copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_overline
        fl.line = value
        ft.prop_overline = fl
        return ft

    def fmt_overline_color(self: _TFontEffects, value: Color | None = None) -> _TFontEffects:
        """
        Gets copy of instance with text overline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_overline
        fl.color = value
        ft.prop_overline = fl
        return ft

    def fmt_strike(self: _TFontEffects, value: FontStrikeoutEnum | None = None) -> _TFontEffects:
        """
        Gets copy of instance with strike set or removed.

        Args:
            value (FontStrikeoutEnum, optional): The type of the strike out of the character.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_strike = value
        return ft

    def fmt_underline(self: _TFontEffects, value: FontUnderlineEnum | None = None) -> _TFontEffects:
        """
        Gets copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_underline
        fl.line = value
        ft.prop_underline = fl
        return ft

    def fmt_underline_color(self: _TFontEffects, value: Color | None = None) -> _TFontEffects:
        """
        Gets copy of instance with text underline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_underline
        fl.color = value
        ft.prop_underline = fl
        return ft

    def fmt_word_mode(self: _TFontEffects, value: bool | None = None) -> _TFontEffects:
        """
        Gets copy of instance with word mode set or removed.

        The underline and strike-through properties are not applied to white spaces when set to ``True``.

        Args:
            value (bool, optional): The word mode.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_word_mode = value
        return ft

    def fmt_case(self: _TFontEffects, value: CaseMapEnum | None = None) -> _TFontEffects:
        """
        Gets copy of instance with case set or removed.

        Args:
            value (CaseMapEnum, optional): The case value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_case = value
        return ft

    def fmt_relief(self: _TFontEffects, value: FontReliefEnum | None = None) -> _TFontEffects:
        """
        Gets copy of instance with relief set or removed.

        Args:
            value (FontReliefEnum, optional): The relief value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_relief = value
        return ft

    def fmt_outline(self: _TFontEffects, value: bool | None = None) -> _TFontEffects:
        """
        Gets copy of instance with outline set or removed.

        Args:
            value (bool, optional): The outline value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_outline = value
        return ft

    def fmt_hidden(self: _TFontEffects, value: bool | None = None) -> _TFontEffects:
        """
        Gets copy of instance with hidden set or removed.

        Args:
            value (bool, optional): The hidden value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_hidden = value
        return ft

    def fmt_shadowed(self: _TFontEffects, value: bool | None = None) -> _TFontEffects:
        """
        Gets copy of instance with shadowed set or removed.

        Args:
            value (bool, optional): The shadowed value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_shadowed = value
        return ft

    # endregion Format Methods

    # region Style Properties
    @property
    def color_auto(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with color set to automatic"""
        ft = self.copy()
        ft.prop_color = StandardColor.AUTO_COLOR
        return ft

    @property
    def overline(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with overline set"""
        ft = self.copy()
        ft.prop_overline = FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.AUTO_COLOR)
        return ft

    @property
    def overline_color_auto(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with overline color set to automatic"""
        ft = self.copy()
        fl = ft.prop_overline
        fl.color = StandardColor.AUTO_COLOR
        ft.prop_overline = fl
        return ft

    @property
    def underline(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with underline set"""
        ft = self.copy()
        ft.prop_underline = FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.AUTO_COLOR)
        return ft

    @property
    def under_color_auto(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with underline color set to automatic"""
        ft = self.copy()
        fl = ft.prop_underline
        fl.color = StandardColor.AUTO_COLOR
        ft.prop_underline = fl
        return ft

    @property
    def strike(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with strike set"""
        ft = self.copy()
        ft.prop_strike = FontStrikeoutEnum.SINGLE
        return ft

    @property
    def word_mode(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with word mode set"""
        ft = self.copy()
        ft.prop_word_mode = True
        return ft

    @property
    def outline(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with outline set"""
        ft = self.copy()
        ft.prop_outline = True
        return ft

    @property
    def hidden(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with hidden set"""
        ft = self.copy()
        ft.prop_hidden = True
        return ft

    @property
    def shadowed(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with shadow set"""
        ft = self.copy()
        ft.prop_shadowed = True
        return ft

    @property
    def case_upper(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with case set to upper"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.UPPERCASE
        return ft

    @property
    def case_lower(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with case set to lower"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.LOWERCASE
        return ft

    @property
    def case_title(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with case set to title"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.TITLE
        return ft

    @property
    def case_small_caps(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with case set to small caps"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.SMALLCAPS
        return ft

    @property
    def case_none(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with no case set"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.NONE
        return ft

    @property
    def relief_none(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with no relief set"""
        ft = self.copy()
        ft.prop_relief = FontReliefEnum.NONE
        return ft

    @property
    def relief_embossed(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with relief set to embossed"""
        ft = self.copy()
        ft.prop_relief = FontReliefEnum.EMBOSSED
        return ft

    @property
    def relief_engraved(self: _TFontEffects) -> _TFontEffects:
        """Gets copy of instance with relief set to engraved"""
        ft = self.copy()
        ft.prop_relief = FontReliefEnum.ENGRAVED
        return ft

    # endregion Style Properties

    # region Prop Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_color(self) -> Color | None:
        """Gets/Sets the value of the text color."""
        return self._get("CharColor")

    @prop_color.setter
    def prop_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharColor")
            return
        if value < 0:
            self._set("CharColor", -1)
        else:
            self._set("CharColor", value)

    @property
    def prop_transparency(self) -> Intensity | None:
        """Gets/Sets The transparency value"""
        pv = cast(int, self._get("CharTransparence"))
        return Intensity(pv) if pv is not None else None

    @prop_transparency.setter
    def prop_transparency(self, value: Intensity | int | None) -> None:
        if value is None:
            self._remove("CharTransparence")
            return
        self._set("CharTransparence", Intensity(int(value)).value)

    @property
    def prop_overline(self) -> FontLine:
        """This property contains the value for the character overline."""
        pv = cast(int, self._get("CharOverline"))
        line = None if pv is None else FontUnderlineEnum(pv)
        return FontLine(line=line, color=cast(Color, self._get("CharOverlineColor")))

    @prop_overline.setter
    def prop_overline(self, value: FontLine | None) -> None:
        if value is None:
            self._remove("CharOverline")
            self._remove("CharOverlineColor")
            self._remove("CharOverlineHasColor")
            return
        if value.line is None:
            self._remove("CharOverline")
        else:
            self._set("CharOverline", value.line.value)

        if value.color is None:
            self._remove("CharOverlineColor")
            self._remove("CharOverlineHasColor")
        elif value.color < 0:
            # automatic color
            self._set("CharOverlineHasColor", False)
            self._set("CharOverlineColor", -1)
        else:
            self._set("CharOverlineHasColor", True)
            self._set("CharOverlineColor", value.color)

    @property
    def prop_underline(self) -> FontLine:
        """This property contains the value for the character underline."""
        pv = cast(int, self._get("CharUnderline"))
        line = None if pv is None else FontUnderlineEnum(pv)
        return FontLine(line=line, color=cast(Color, self._get("CharUnderlineColor")))

    @prop_underline.setter
    def prop_underline(self, value: FontLine | None) -> None:
        if value is None:
            self._remove("CharUnderline")
            self._remove("CharUnderlineColor")
            self._remove("CharUnderlineHasColor")
            return
        if value.line is None:
            self._remove("CharUnderline")
        else:
            self._set("CharUnderline", value.line.value)

        if value.color is None:
            self._remove("CharUnderlineColor")
            self._remove("CharUnderlineHasColor")
        elif value.color < 0:
            # automatic color
            self._set("CharUnderlineHasColor", False)
            self._set("CharUnderlineColor", -1)
        else:
            self._set("CharUnderlineHasColor", True)
            self._set("CharUnderlineColor", value.color)

    @property
    def prop_strike(self) -> FontStrikeoutEnum | None:
        """Gets/Sets the type of the strike out of the character."""
        pv = cast(int, self._get("CharStrikeout"))
        return FontStrikeoutEnum(pv) if pv is not None else None

    @prop_strike.setter
    def prop_strike(self, value: FontStrikeoutEnum | None) -> None:
        if value is None:
            self._remove("CharStrikeout")
            return
        self._set("CharStrikeout", value.value)

    @property
    def prop_word_mode(self) -> bool | None:
        """Gets/Sets Character word mode. If this property is ``True``, the underline and strike-through properties are not applied to white spaces."""
        return self._get("CharWordMode")

    @prop_word_mode.setter
    def prop_word_mode(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharWordMode")
            return
        self._set("CharWordMode", value)

    @property
    def prop_case(self) -> CaseMapEnum | None:
        """Gets/Sets Font Case Value"""
        pv = cast(int, self._get("CharCaseMap"))
        return None if pv is None else CaseMapEnum(pv)

    @prop_case.setter
    def prop_case(self, value: CaseMapEnum | None) -> None:
        if value is None:
            self._remove("CharCaseMap")
            return
        self._set("CharCaseMap", value.value)

    @property
    def prop_relief(self) -> FontReliefEnum | None:
        """Gets/Sets Font Relief Value"""
        pv = cast(int, self._get("CharRelief"))
        return None if pv is None else FontReliefEnum(pv)

    @prop_relief.setter
    def prop_relief(self, value: FontReliefEnum | None) -> None:
        if value is None:
            self._remove("CharRelief")
            return
        self._set("CharRelief", value.value)

    @property
    def prop_outline(self) -> bool | None:
        """Gets/Sets if the font is outlined"""
        return self._get("CharContoured")

    @prop_outline.setter
    def prop_outline(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharContoured")
            return
        self._set("CharContoured", value)

    @property
    def prop_hidden(self) -> bool | None:
        """Gets/Sets if the font is hidden"""
        return self._get("CharHidden")

    @prop_hidden.setter
    def prop_hidden(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharHidden")
            return
        self._set("CharHidden", value)

    @property
    def prop_shadowed(self) -> bool | None:
        """Gets/Sets if the characters are formatted and displayed with a shadow effect."""
        return self._get("CharShadowed")

    @prop_shadowed.setter
    def prop_shadowed(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharShadowed")
            return
        self._set("CharShadowed", value)

    @property
    def default(self: _TFontEffects) -> _TFontEffects:  # type: ignore[misc]
        """Gets Font Position default."""
        try:
            return self._default_inst
        except AttributeError:
            # pylint: disable=protected-access
            # pylint: disable=unexpected-keyword-arg
            fe = self.__class__(_cattribs=self._get_internal_cattribs())
            fe._set("CharColor", -1)
            fe._set("CharOverline", 0)
            fe._set("CharOverlineColor", -1)
            fe._set("CharOverlineHasColor", False)
            fe._set("CharUnderline", 0)
            fe._set("CharUnderlineColor", -1)
            fe._set("CharUnderlineHasColor", False)
            fe._set("CharWordMode", False)
            fe._set("CharShadowed", False)
            fe._set("CharContoured", False)
            fe._set("CharHidden", False)
            fe._set("CharTransparence", 100)
            fe._set("CharStrikeout", 0)
            fe._set("CharCaseMap", 0)
            fe._set("CharRelief", 0)
            fe._is_default_inst = True
            self._default_inst = fe
        return self._default_inst

    # endregion Prop Properties
