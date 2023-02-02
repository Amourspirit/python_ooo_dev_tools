"""
Module for managing character fonts.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, overload

from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.color import Color
from .....utils.data_type.intensity import Intensity as Intensity
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.style.case_map import CaseMapEnum as CaseMapEnum
from ooo.dyn.awt.font_relief import FontReliefEnum as FontReliefEnum


class FontEffects(StyleBase):
    """
    Character Font Effects

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties can be chained together.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    def __init__(
        self,
        *,
        color: Color | None = None,
        transparency: Intensity | int | None = None,
        overline: FontUnderlineEnum | None = None,
        overline_color: Color | None = None,
        strike: FontStrikeoutEnum | None = None,
        underine: FontUnderlineEnum | None = None,
        underline_color: Color | None = None,
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
            color (Color, optional): The value of the text color.
                If value is ``-1`` the automatic color is applied.
            transparency (Intensity, int, optional): The transparency value from ``0`` to ``100`` for the font color.
            overline (FontUnderlineEnum, optional): The value for the character overline.
            overline_color (Color, optional): Specifies if the property ``CharOverlinelineColor`` is used for an overline.
                If value is ``-1`` the automatic color is applied.
            strike (FontStrikeoutEnum, optional): Detrmines the type of the strike out of the character.
            underine (FontUnderlineEnum, optional): The value for the character underline.
            underline_color (Color, optional): Specifies if the property ``CharUnderlineColor`` is used for an underline.
                If value is ``-1`` the automatic color is applied.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to white spaces.
            case (CaseMapEnum, optional): Specifies the case of the font.
            releif (FontReliefEnum, optional): Specifies the relief of the font.
            outline (bool, optional): Specifies if the font is outlined.
            hidden (bool, optional): Specifies if the font is hidden.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
        """
        # could not find any documention in the API or elsewhere online for Overline
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
        init_vals = {
            "CharWordMode": word_mode,
            "CharShadowed": shadowed,
            "CharContoured": outline,
            "CharHidden": hidden,
        }
        if not color is None:
            if color < 0:
                init_vals["CharColor"] = -1
            else:
                init_vals["CharColor"] = color

        if not transparency is None:
            transparency = Intensity(int(transparency))
            init_vals["CharTransparence"] = transparency.value

        if not overline_color is None:
            if overline_color < 0:
                # -1 is set color to automatic
                init_vals["CharOverlineColor"] = -1
                init_vals["CharOverlineHasColor"] = False
            else:
                init_vals["CharOverlineColor"] = overline_color
                init_vals["CharOverlineHasColor"] = True

        if not underline_color is None:
            if underline_color < 0:
                # -1 is set color to automatic
                init_vals["CharUnderlineColor"] = -1
                init_vals["CharUnderlineHasColor"] = False
            else:
                init_vals["CharUnderlineColor"] = underline_color
                init_vals["CharUnderlineHasColor"] = True

        if not strike is None:
            init_vals["CharStrikeout"] = strike.value

        if not overline is None:
            init_vals["CharOverline"] = overline.value

        if not underine is None:
            init_vals["CharUnderline"] = underine.value

        if not case is None:
            init_vals["CharCaseMap"] = case.value

        if not relief is None:
            init_vals["CharRelief"] = relief.value

        super().__init__(**init_vals)

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Font.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    @staticmethod
    def from_obj(obj: object) -> FontEffects:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            FontEffects: ``FontEffects`` instance that represents ``obj`` font effects.
        """
        if not FontEffects.default._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported to get FontEffects from.")

        def set_prop(key: str, fe: FontEffects):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                fe._set(key, val)

        inst = FontEffects()
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
        return inst

    # endregion methods

    # region Format Methods
    def fmt_color(self, value: Color | None = None) -> FontEffects:
        """
        Gets copy of instance with text color set or removed.

        Args:
            value (Color, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_color = value
        return ft

    def fmt_transparency(self, value: bool | None = None) -> FontEffects:
        """
        Gets copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_bg_transparent = value
        return ft

    def fmt_overline(self, value: FontUnderlineEnum | None = None) -> FontEffects:
        """
        Gets copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_overline = value
        return ft

    def fmt_overline_color(self, value: Color | None = None) -> FontEffects:
        """
        Gets copy of instance with text overline color set or removed.

        Args:
            value (Color, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_overline_color = value
        return ft

    def fmt_strike(self, value: FontStrikeoutEnum | None = None) -> FontEffects:
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

    def fmt_underline(self, value: FontUnderlineEnum | None = None) -> FontEffects:
        """
        Gets copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_underline = value
        return ft

    def fmt_underline_color(self, value: Color | None = None) -> FontEffects:
        """
        Gets copy of instance with text underline color set or removed.

        Args:
            value (Color, optional): The color is used for an underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_underline_color = value
        return ft

    def fmt_word_mode(self, value: bool | None = None) -> FontEffects:
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

    def fmt_case(self, value: CaseMapEnum | None = None) -> FontEffects:
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

    def fmt_relief(self, value: FontReliefEnum | None = None) -> FontEffects:
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

    def fmt_outline(self, value: bool | None = None) -> FontEffects:
        """
        Gets copy of instance with outline set or removed.

        Args:
            value (bool, optional): The outline value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontEffects: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_relief = value
        return ft

    def fmt_hidden(self, value: bool | None = None) -> FontEffects:
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

    def fmt_shadowed(self, value: bool | None = None) -> FontEffects:
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
    def color_auto(self) -> FontEffects:
        """Gets copy of instance with color set to automatic"""
        ft = self.copy()
        ft.prop_color = -1
        return ft

    @property
    def overline(self) -> FontEffects:
        """Gets copy of instance with overline set"""
        ft = self.copy()
        ft.prop_overline = FontUnderlineEnum.SINGLE
        return ft

    @property
    def overline_color_auto(self) -> FontEffects:
        """Gets copy of instance with overline color set to automatic"""
        ft = self.copy()
        ft.prop_overline_color = -1
        return ft

    @property
    def strike(self) -> FontEffects:
        """Gets copy of instance with strike set"""
        ft = self.copy()
        ft.prop_strike = FontStrikeoutEnum.SINGLE
        return ft

    @property
    def underline(self) -> FontEffects:
        """Gets copy of instance with underline set"""
        ft = self.copy()
        ft.prop_underline = FontUnderlineEnum.SINGLE
        return ft

    @property
    def under_color_auto(self) -> FontEffects:
        """Gets copy of instance with underline color set to automatic"""
        ft = self.copy()
        ft.prop_underline_color = -1
        return ft

    @property
    def word_mode(self) -> FontEffects:
        """Gets copy of instance with word mode set"""
        ft = self.copy()
        ft.prop_word_mode = True
        return ft

    @property
    def outline(self) -> FontEffects:
        """Gets copy of instance with outline set"""
        ft = self.copy()
        ft.prop_outline = True
        return ft

    @property
    def hidden(self) -> FontEffects:
        """Gets copy of instance with hidden set"""
        ft = self.copy()
        ft.prop_hidden = True
        return ft

    @property
    def shadowed(self) -> FontEffects:
        """Gets copy of instance with shadow set"""
        ft = self.copy()
        ft.prop_shadowed = True
        return ft

    @property
    def case_upper(self) -> FontEffects:
        """Gets copy of instance with case set to upper"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.UPPERCASE
        return ft

    @property
    def case_lower(self) -> FontEffects:
        """Gets copy of instance with case set to lower"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.LOWERCASE
        return ft

    @property
    def case_title(self) -> FontEffects:
        """Gets copy of instance with case set to title"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.TITLE
        return ft

    @property
    def case_small_caps(self) -> FontEffects:
        """Gets copy of instance with case set to small caps"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.SMALLCAPS
        return ft

    @property
    def case_none(self) -> FontEffects:
        """Gets copy of instance with no case set"""
        ft = self.copy()
        ft.prop_case = CaseMapEnum.NONE
        return ft

    @property
    def relief_none(self) -> FontEffects:
        """Gets copy of instance with no relief set"""
        ft = self.copy()
        ft.prop_relief = FontReliefEnum.NONE
        return ft

    @property
    def relief_embossed(self) -> FontEffects:
        """Gets copy of instance with relief set to embossed"""
        ft = self.copy()
        ft.prop_relief = FontReliefEnum.EMBOSSED
        return ft

    @property
    def relief_engraved(self) -> FontEffects:
        """Gets copy of instance with relief set to engraved"""
        ft = self.copy()
        ft.prop_relief = FontReliefEnum.ENGRAVED
        return ft

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    @property
    def prop_color(self) -> Color | None:
        """Gets/Sets the value of the text color."""
        return self._get("CharColor")

    @prop_color.setter
    def prop_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharColor")
            return
        self._set("CharColor", value)

    @property
    def prop_transparency(self) -> Intensity | None:
        """Gets/Sets The transparency value"""
        pv = cast(int, self._get("CharTransparence"))
        if not pv is None:
            return Intensity(pv)
        return None

    @prop_transparency.setter
    def prop_transparency(self, value: Intensity | int | None) -> None:
        if value is None:
            self._remove("CharTransparence")
            return
        self._set("CharTransparence", Intensity(int(value)).value)

    @property
    def prop_overline(self) -> FontUnderlineEnum | None:
        """This property contains the value for the character overline."""
        pv = cast(int, self._get("CharOverline"))
        if not pv is None:
            return FontUnderlineEnum(pv)
        return None

    @prop_overline.setter
    def prop_overline(self, value: FontUnderlineEnum | None) -> None:
        if value is None:
            self._remove("CharOverline")
            return
        self._set("CharOverline", value.value)

    @property
    def prop_overline_color(self) -> Color | None:
        """This property specifies if the property ``CharOverlineColor`` is used for an overline."""
        return self._get("CharOverlineColor")

    @prop_overline_color.setter
    def prop_overline_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharOverlineColor")
            self._remove("CharOverlineHasColor")
            return
        if value < 0:
            # automatic color
            self._set("CharOverlineHasColor", False)
            self._set("CharOverlineColor", -1)
        else:
            self._set("CharOverlineHasColor", True)
            self._set("CharOverlineColor", value)

    @property
    def prop_strike(self) -> FontStrikeoutEnum | None:
        """Gets/Sets the type of the strike out of the character."""
        pv = cast(int, self._get("CharStrikeout"))
        if not pv is None:
            return FontStrikeoutEnum(pv)
        return None

    @prop_strike.setter
    def prop_strike(self, value: FontStrikeoutEnum | None) -> None:
        if value is None:
            self._remove("CharStrikeout")
            return
        self._set("CharStrikeout", value.value)

    @property
    def prop_underline(self) -> FontUnderlineEnum | None:
        """Gets/Sets underline"""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            return FontUnderlineEnum(pv)
        return None

    @prop_underline.setter
    def prop_underline(self, value: FontUnderlineEnum | None) -> None:
        if value is None:
            self._remove("CharUnderline")
            return
        self._set("CharUnderline", value.value)

    @property
    def prop_underline_color(self) -> Color | None:
        """Gets/Sets if the property ``CharUnderlineColor`` is used for an underline."""
        return self._get("CharUnderlineColor")

    @prop_underline_color.setter
    def prop_underline_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharUnderlineColor")
            self._remove("CharUnderlineHasColor")
            return
        if value < 0:
            # automatic color
            self._set("CharUnderlineHasColor", False)
            self._set("CharUnderlineColor", -1)
        else:
            self._set("CharUnderlineHasColor", True)
            self._set("CharUnderlineColor", value)

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
        if pv is None:
            return None
        return CaseMapEnum(pv)

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
        if pv is None:
            return None
        return FontReliefEnum(pv)

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

    @static_prop
    def default() -> FontEffects:  # type: ignore[misc]
        """Gets Font Position default. Static Property."""
        if FontEffects._DEFAULT is None:
            fe = FontEffects()
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
            FontEffects._DEFAULT = fe
        return FontEffects._DEFAULT

    # endregion Prop Properties
