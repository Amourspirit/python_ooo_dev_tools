from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import contextlib
import uno

from ooo.dyn.awt.font_underline import FontUnderlineEnum
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.text.ruby_position import RubyPositionEnum
from ooodev.adapter.table.border_line2_struct_comp import BorderLine2StructComp
from ooodev.adapter.table.shadow_format_struct_comp import ShadowFormatStructComp
from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.events.events import Events
from ooodev.units.unit_pt import UnitPT
from ooodev.utils import info as mInfo
from ooodev.units.angle10 import Angle10
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XNameContainer
    from com.sun.star.lang import Locale
    from com.sun.star.style import CharacterProperties
    from com.sun.star.table import BorderLine2  # struct
    from com.sun.star.table import ShadowFormat  # struct
    from ooo.dyn.awt.font_slant import FontSlant
    from ooodev.events.args.key_val_args import KeyValArgs
    from ooodev.utils.color import Color  # type def
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.angle_t import AngleT
    from ooodev.units.unit_obj import UnitT


class CharacterPropertiesPartial:
    """
    Partial class for CharacterProperties.

    See Also:
        `API CharacterProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html>`_
    """

    def __init__(self, component: CharacterProperties) -> None:
        """
        Constructor

        Args:
            component (CharacterProperties): UNO Component that implements ``com.sun.star.style.CharacterProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``CharacterProperties``.
        """
        self.__component = component
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.__event_provider.subscribe_event(
            "com_sun_star_table_BorderLine2_changed", self.__fn_on_comp_struct_changed
        )
        self.__event_provider.subscribe_event(
            "com_sun_star_table_ShadowFormat_changed", self.__fn_on_comp_struct_changed
        )

    # region CharacterProperties
    @property
    def char_interop_grab_bag(self) -> Tuple[PropertyValue, ...] | None:
        """
        Gets/Sets grab bag of character properties, used as a string-any map for interim interop purposes.

        This property is intentionally not handled by the ODF filter. Any member that should be handled there should be first moved out from this grab bag to a separate property.

        **Optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharInteropGrabBag
        return None

    @char_interop_grab_bag.setter
    def char_interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharInteropGrabBag = value

    @property
    def char_style_names(self) -> Tuple[str, ...] | None:
        """
        Gets/Sets - This optional property specifies the names of the all styles applied to the font.

        It is not guaranteed that the order in the sequence reflects the order of the evaluation of the character style attributes.

        **Optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharStyleNames
        return None

    @char_style_names.setter
    def char_style_names(self, value: Tuple[str, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharStyleNames = value

    @property
    def char_auto_kerning(self) -> bool | None:
        """
        Gets/Sets - This optional property determines whether the kerning tables from the current font are used.

        Automatic kerning applies a spacing in between certain pairs of characters to make the text look better.

        **Optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharAutoKerning
        return None

    @char_auto_kerning.setter
    def char_auto_kerning(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharAutoKerning = value

    @property
    def char_back_color(self) -> Color | None:
        """
        Get/Sets - This optional property contains the text background color.

        **Optional**

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not supported.
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharBackColor  # type: ignore
        return None

    @char_back_color.setter
    def char_back_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharBackColor = value  # type: ignore

    @property
    def char_back_transparent(self) -> bool | None:
        """
        Gets/Sets if the text background color is set to transparent.

        **Optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharBackTransparent
        return None

    @char_back_transparent.setter
    def char_back_transparent(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharBackTransparent = value

    @property
    def char_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets the distance from the border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **Optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.CharBorderDistance)
        return None

    @char_border_distance.setter
    def char_border_distance(self, value: int | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def char_bottom_border(self) -> BorderLine2StructComp | None:
        """
        This property contains the bottom border of the object.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        **optional**

        Returns:
            BorderLine2StructComp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "CharBottomBorder"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.CharBottomBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @char_bottom_border.setter
    def char_bottom_border(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        key = "CharBottomBorder"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.CharBottomBorder = value.copy()
        else:
            self.__component.CharBottomBorder = cast("BorderLine2", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def char_bottom_border_distance(self) -> UnitMM100 | None:
        """
        This property contains the distance from the bottom border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.CharBottomBorderDistance)
        return None

    @char_bottom_border_distance.setter
    def char_bottom_border_distance(self, value: int | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharBottomBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def char_case_map(self) -> int | None:
        """
        Gets/Sets - This optional property contains the value of the case-mapping of the text for formatting and displaying.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharCaseMap
        return None

    @char_case_map.setter
    def char_case_map(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharCaseMap = value

    @property
    def char_color(self) -> Color:
        """
        This property contains the value of the text color.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.__component.CharColor  # type: ignore

    @char_color.setter
    def char_color(self, value: Color) -> None:
        self.__component.CharColor = value  # type: ignore

    @property
    def char_color_theme(self) -> int | None:
        """
        Gets/Sets - If available, keeps the color theme index, so that the character can be re-colored easily based on a theme.

        **since**

            LibreOffice ``7.3``

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharColorTheme
        return None

    @char_color_theme.setter
    def char_color_theme(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharColorTheme = value

    @property
    def char_color_tint_or_shade(self) -> int | None:
        """
        Gets/Sets the tint or shade of the character color.

        **since**

            LibreOffice ``7.3``

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharColorTintOrShade
        return None

    @char_color_tint_or_shade.setter
    def char_color_tint_or_shade(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharColorTintOrShade = value

    @property
    def char_combine_is_on(self) -> bool | None:
        """
        Gets/Sets - This optional property determines whether text is formatted in two lines.

        It is linked to the properties ``char_combine_prefix`` and ``char_combine_suffix``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharCombineIsOn
        return None

    @char_combine_is_on.setter
    def char_combine_is_on(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharCombineIsOn = value

    @property
    def char_combine_prefix(self) -> str | None:
        """
        Gets/Sets - This optional property contains the prefix (usually parenthesis) before text that is formatted in two lines.

        It is linked to the properties ``char_combine_is_on`` and ``char_combine_suffix``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharCombinePrefix
        return None

    @char_combine_prefix.setter
    def char_combine_prefix(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharCombinePrefix = value

    @property
    def char_combine_suffix(self) -> str | None:
        """
        Gets/Sets - This optional property contains the suffix (usually parenthesis) after text that is formatted in two lines.

        It is linked to the properties CharCombineIsOn and CharCombinePrefix.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharCombineSuffix
        return None

    @char_combine_suffix.setter
    def char_combine_suffix(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharCombineSuffix = value

    @property
    def char_contoured(self) -> bool | None:
        """
        Gets/Sets - This optional property specifies if the characters are formatted and displayed with a contour effect.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharContoured
        return None

    @char_contoured.setter
    def char_contoured(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharContoured = value

    @property
    def char_crossed_out(self) -> bool | None:
        """
        Gets/Sets - This property is ``True`` if the characters are crossed out.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharCrossedOut
        return None

    @char_crossed_out.setter
    def char_crossed_out(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharCrossedOut = value

    @property
    def char_emphasis(self) -> int | None:
        """
        Gets/Sets - This optional property contains the font emphasis value.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharEmphasis
        return None

    @char_emphasis.setter
    def char_emphasis(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharEmphasis = value

    @property
    def char_escapement(self) -> int | None:
        """
        Gets/Sets the percentage by which to raise/lower superscript/subscript characters.

        Negative values denote subscripts and positive values superscripts.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharEscapement
        return None

    @char_escapement.setter
    def char_escapement(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharEscapement = value

    @property
    def char_escapement_height(self) -> int | None:
        """
        Gets/Sets - This is the relative height used for subscript or superscript characters in units of percent.

        The value 100 denotes the original height of the characters.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharEscapementHeight
        return None

    @char_escapement_height.setter
    def char_escapement_height(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharEscapementHeight = value

    @property
    def char_flash(self) -> bool | None:
        """
        Gets/Sets - If this optional property is ``True``, then the characters are flashing.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharFlash
        return None

    @char_flash.setter
    def char_flash(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharFlash = value

    @property
    def char_font_char_set(self) -> int:
        """
        Gets/Sets - This property contains the text encoding of the font.
        """
        return self.__component.CharFontCharSet

    @char_font_char_set.setter
    def char_font_char_set(self, value: int) -> None:
        self.__component.CharFontCharSet = value

    @property
    def char_font_family(self) -> int:
        """
        Get/Sets the font family.
        """
        return self.__component.CharFontFamily

    @char_font_family.setter
    def char_font_family(self, value: int) -> None:
        self.__component.CharFontFamily = value

    @property
    def char_font_name(self) -> str:
        """
        Gets/Sets the name of the font style.

        It may contain more than one name separated by comma.
        """
        return self.__component.CharFontName

    @char_font_name.setter
    def char_font_name(self, value: str) -> None:
        self.__component.CharFontName = value

    @property
    def char_font_pitch(self) -> int:
        """
        Gets/Sets the font pitch.
        """
        return self.__component.CharFontPitch

    @char_font_pitch.setter
    def char_font_pitch(self, value: int) -> None:
        self.__component.CharFontPitch = value

    @property
    def char_font_style_name(self) -> str:
        """
        Gets/Sets the name of the font style.

        This property may be empty.
        """
        return self.__component.CharFontStyleName

    @char_font_style_name.setter
    def char_font_style_name(self, value: str) -> None:
        self.__component.CharFontStyleName = value

    @property
    def char_font_type(self) -> int | None:
        """
        Gets/Sets - This optional property specifies the fundamental technology of the font.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharFontType
        return None

    @char_font_type.setter
    def char_font_type(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharFontType = value

    @property
    def char_height(self) -> UnitPT:
        """
        Gets/Sets - This value contains the height of the characters in point.

        When setting the value can be a float (in points) or a ``UnitT`` instance.
        """
        return UnitPT(self.__component.CharHeight)

    @char_height.setter
    def char_height(self, value: float | UnitT) -> None:
        self.__component.CharHeight = UnitPT.from_unit_val(value).value

    @property
    def char_hidden(self) -> bool | None:
        """
        Gets/Sets - If this optional property is ``True``, then the characters are invisible.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharHidden
        return None

    @char_hidden.setter
    def char_hidden(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharHidden = value

    @property
    def char_highlight(self) -> Color | None:
        """
        Gets/Sets the color of the highlight.

        **optional**

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharHighlight  # type: ignore
        return None

    @char_highlight.setter
    def char_highlight(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharHighlight = value  # type: ignore

    @property
    def char_keep_together(self) -> bool | None:
        """
        Gets/Sets - This optional property marks a range of characters to prevent it from being broken into two lines.

        A line break is applied before the range of characters if the layout makes a break necessary within the range.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharKeepTogether
        return None

    @char_keep_together.setter
    def char_keep_together(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharKeepTogether = value

    @property
    def char_kerning(self) -> int | None:
        """
        Gets/Sets - This optional property contains the value of the kerning of the characters.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharKerning
        return None

    @char_kerning.setter
    def char_kerning(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharKerning = value

    @property
    def char_left_border(self) -> BorderLine2StructComp | None:
        """
        Gets/Sets - This property contains the left border of the object.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        **optional**

        Returns:
            BorderLine2StructComp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "CharLeftBorder"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.CharLeftBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @char_left_border.setter
    def char_left_border(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        key = "CharLeftBorder"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.CharLeftBorder = value.copy()
        else:
            self.__component.CharLeftBorder = cast("BorderLine2", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def char_left_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets - This property contains the distance from the left border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.CharLeftBorderDistance)
        return None

    @char_left_border_distance.setter
    def char_left_border_distance(self, value: int | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharLeftBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def char_locale(self) -> Locale:
        """
        Gets/Sets - This property contains the value of the locale.
        """
        return self.__component.CharLocale

    @char_locale.setter
    def char_locale(self, value: Locale) -> None:
        self.__component.CharLocale = value

    @property
    def char_no_hyphenation(self) -> bool | None:
        """
        Gets/Sets - This optional property determines if the word can be hyphenated at the character.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharNoHyphenation
        return None

    @char_no_hyphenation.setter
    def char_no_hyphenation(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharNoHyphenation = value

    @property
    def char_no_line_break(self) -> bool | None:
        """
        Gets/Sets - This optional property marks a range of characters to ignore a line break in this area.

        A line break is applied behind the range of characters if the layout makes a break necessary within the range. That means that the text may go through the border.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharNoLineBreak
        return None

    @char_no_line_break.setter
    def char_no_line_break(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharNoLineBreak = value

    @property
    def char_posture(self) -> FontSlant:
        """
        Gets/Sets - This property contains the value of the posture of the document.

        Returns:
            FontSlant: Returns FontSlant

        Hint:
            - ``FontSlant`` can be imported from ``ooo.dyn.awt.font_slant``
        """
        # from ooo.dyn.awt.font_slant import FontSlant
        return self.__component.CharPosture  # type: ignore

    @char_posture.setter
    def char_posture(self, value: FontSlant) -> None:
        self.__component.CharPosture = value  # type: ignore

    @property
    def char_relief(self) -> int | None:
        """
        Gets/Sets - This optional property contains the relief style of the characters.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharRelief
        return None

    @char_relief.setter
    def char_relief(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharRelief = value

    @property
    def char_right_border(self) -> BorderLine2StructComp | None:
        """
        Gets/Sets - This property contains the right border of the object.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        **optional**

        Returns:
            BorderLine2StructComp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "CharRightBorder"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.CharRightBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @char_right_border.setter
    def char_right_border(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        key = "CharRightBorder"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.CharRightBorder = value.copy()
        else:
            self.__component.CharRightBorder = cast("BorderLine2", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def char_right_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets - This property contains the distance from the right border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.CharRightBorderDistance)
        return None

    @char_right_border_distance.setter
    def char_right_border_distance(self, value: int | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharRightBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def char_rotation(self) -> Angle10 | None:
        """
        Gets/Sets - This optional property determines the rotation of a character in tenths of a degree.

        Depending on the implementation only certain values may be allowed.

        **optional**

        Returns:
            Angle10: Returns Angle10, ``1/10th`` degrees, or None if not supported.

        Hint:
            - ``Angle10`` can be imported from ``ooodev.units``
        """
        with contextlib.suppress(AttributeError):
            angle = self.__component.CharRotation
            return Angle10(angle)
        return None

    @char_rotation.setter
    def char_rotation(self, value: int | AngleT) -> None:
        if not hasattr(self.__component, "CharRotation"):
            return
        val = Angle10.from_unit_val(value)
        self.__component.CharRotation = val.value

    @property
    def char_rotation_is_fit_to_line(self) -> bool | None:
        """
        Gets/Sets - This optional property determines whether the text formatting tries to fit rotated text into the surrounded line height.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharRotationIsFitToLine
        return None

    @char_rotation_is_fit_to_line.setter
    def char_rotation_is_fit_to_line(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharRotationIsFitToLine = value

    @property
    def char_scale_width(self) -> int | None:
        """
        Gets/Sets - This optional property determines the percentage value for scaling the width of characters.

        The value refers to the original width which is denoted by ``100``, and it has to be greater than ``0``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharScaleWidth
        return None

    @char_scale_width.setter
    def char_scale_width(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharScaleWidth = value

    @property
    def char_shading_value(self) -> int | None:
        """
        Gets/Sets - This optional property contains the text shading value.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharShadingValue
        return None

    @char_shading_value.setter
    def char_shading_value(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharShadingValue = value

    @property
    def char_shadow_format(self) -> ShadowFormatStructComp | None:
        """
        Gets/Sets the type, color, and width of the shadow.

        When setting the value can be an instance of ``ShadowFormatStructComp`` or ``ShadowFormat``.

        **optional**

        Returns:
            ShadowFormatStructComp: Shadow Format or None if not supported.

        Hint:
            - ``ShadowFormat`` can be imported from ``ooo.dyn.table.shadow_format``
        """
        key = "CharShadowFormat"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = ShadowFormatStructComp(self.__component.CharShadowFormat, key, self.__event_provider)
            self.__props[key] = prop
        return cast(ShadowFormatStructComp, prop)

    @char_shadow_format.setter
    def char_shadow_format(self, value: ShadowFormat | ShadowFormatStructComp) -> None:
        key = "ShadowFormat"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, ShadowFormatStructComp):
            self.__component.CharShadowFormat = value.copy()
        else:
            self.__component.CharShadowFormat = cast("ShadowFormat", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def char_shadowed(self) -> bool | None:
        """
        Gets/Sets - This optional property specifies if the characters are formatted and displayed with a shadow effect.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharShadowed
        return None

    @char_shadowed.setter
    def char_shadowed(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharShadowed = value

    @property
    def char_strikeout(self) -> FontStrikeoutEnum | None:
        """
        Gets/Sets - This property determines the type of the strike out of the character.

        **optional**

        Returns:
            FontStrikeoutEnum | None: Returns FontStrikeoutEnum or None if not supported.

        Hint:
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``
        """
        with contextlib.suppress(AttributeError):
            return FontStrikeoutEnum(self.__component.CharStrikeout)
        return None

    @char_strikeout.setter
    def char_strikeout(self, value: int | FontStrikeoutEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharStrikeout = FontStrikeoutEnum(value).value

    @property
    def char_style_name(self) -> str | None:
        """
        Gets/Sets - This optional property specifies the name of the style of the font.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharStyleName
        return None

    @char_style_name.setter
    def char_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharStyleName = value

    @property
    def char_top_border(self) -> BorderLine2StructComp | None:
        """
        Gets/Sets - This property contains the top border of the object.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        **optional**

        Returns:
            BorderLine2StructComp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "CharTopBorder"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.__component.CharTopBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @char_top_border.setter
    def char_top_border(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        key = "CharTopBorder"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            self.__component.CharTopBorder = value.copy()
        else:
            self.__component.CharTopBorder = cast("BorderLine2", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def char_top_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets - This property contains the distance from the top border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.CharTopBorderDistance)
        return None

    @char_top_border_distance.setter
    def char_top_border_distance(self, value: int | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharTopBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def char_transparence(self) -> int | None:
        """
        Gets/Sets - This is the transparency of the character text.

        The value ``100`` means entirely transparent, while ``0`` means not transparent at all.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharTransparence
        return None

    @char_transparence.setter
    def char_transparence(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharTransparence = value

    @property
    def char_underline(self) -> FontUnderlineEnum:
        """
        This property contains the value for the character underline.
        """
        return FontUnderlineEnum(self.__component.CharUnderline)

    @char_underline.setter
    def char_underline(self, value: int | FontUnderlineEnum) -> None:
        self.__component.CharUnderline = FontUnderlineEnum(value).value

    @property
    def char_underline_color(self) -> Color:
        """
        Gets/Sets the color of the underline for the characters.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.__component.CharUnderlineColor  # type: ignore

    @char_underline_color.setter
    def char_underline_color(self, value: Color) -> None:
        self.__component.CharUnderlineColor = value  # type: ignore

    @property
    def char_underline_has_color(self) -> bool:
        """
        Gets/Sets if the property ``char_underline_color`` is used for an underline.
        """
        return self.__component.CharUnderlineHasColor

    @char_underline_has_color.setter
    def char_underline_has_color(self, value: bool) -> None:
        self.__component.CharUnderlineHasColor = value

    @property
    def char_weight(self) -> float:
        """
        Gets/Sets the value of the font weight.

        Example:
            .. code-block:: python

                from com.sun.star.awt import FontWeight

                my_char_properties.char_weight = FontWeight.BOLD
        """
        return self.__component.CharWeight

    @char_weight.setter
    def char_weight(self, value: float) -> None:
        self.__component.CharWeight = value

    @property
    def char_word_mode(self) -> bool | None:
        """
        Gets/Sets - If this property is ``True``, the underline and strike-through properties are not applied to white spaces.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CharWordMode
        return None

    @char_word_mode.setter
    def char_word_mode(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CharWordMode = value

    @property
    def hyper_link_name(self) -> str | None:
        """
        Gets/Sets - This optional property contains the name of the hyperlink.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.HyperLinkName
        return None

    @hyper_link_name.setter
    def hyper_link_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.HyperLinkName = value

    @property
    def hyper_link_target(self) -> str | None:
        """
        Gets/Sets - This optional property contains the name of the target for a hyperlink.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.HyperLinkTarget
        return None

    @hyper_link_target.setter
    def hyper_link_target(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.HyperLinkTarget = value

    @property
    def hyper_link_url(self) -> str | None:
        """
        Gets/Sets - This optional property contains the URL of a hyperlink.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.HyperLinkURL
        return None

    @hyper_link_url.setter
    def hyper_link_url(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.HyperLinkURL = value

    @property
    def ruby_adjust(self) -> int | None:
        """
        Gets/Sets -  This optional property determines the adjustment of the ruby .

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RubyAdjust
        return None

    @ruby_adjust.setter
    def ruby_adjust(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RubyAdjust = value

    @property
    def ruby_char_style_name(self) -> str | None:
        """
        Gets/Sets - This optional property contains the name of the character style that is applied to RubyText.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RubyCharStyleName
        return None

    @ruby_char_style_name.setter
    def ruby_char_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RubyCharStyleName = value

    @property
    def ruby_is_above(self) -> bool | None:
        """
        Gets/Sets - This optional property determines whether the ruby text is printed above/left or below/right of the text.

        This property is replaced by RubyPosition.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RubyIsAbove
        return None

    @ruby_is_above.setter
    def ruby_is_above(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RubyIsAbove = value

    @property
    def ruby_position(self) -> RubyPositionEnum | None:
        """
        Gets/Sets - This optional property determines the position of the ruby .

        **optional**

        Returns:
            RubyPositionEnum | None: Returns RubyPositionEnum or None if not supported.

        Hint:
            - ``RubyPositionEnum`` can be imported from ``ooo.dyn.text.ruby_position``
        """
        with contextlib.suppress(AttributeError):
            return RubyPositionEnum(self.__component.RubyPosition)
        return None

    @ruby_position.setter
    def ruby_position(self, value: int | RubyPositionEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RubyPosition = RubyPositionEnum(value).value

    @property
    def ruby_text(self) -> str | None:
        """
        Gets/Sets - This optional property contains the text that is set as ruby.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RubyText
        return None

    @ruby_text.setter
    def ruby_text(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RubyText = value

    @property
    def text_user_defined_attributes(self) -> NameContainerComp | None:
        """
        Gets/Sets - This property stores XML attributes.

        They will be saved to and restored from automatic styles inside XML files.

        When setting the value, it can be a ``NameContainerComp`` or a ``XNameContainer`` instance.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            comp = self.__component.TextUserDefinedAttributes
            return None if comp is None else NameContainerComp(comp)
        return None

    @text_user_defined_attributes.setter
    def text_user_defined_attributes(self, value: XNameContainer | NameContainerComp) -> None:
        if not hasattr(self.__component, "TextUserDefinedAttributes"):
            return
        if mInfo.Info.is_instance(value, NameContainerComp):
            self.__component.TextUserDefinedAttributes = value.component
        else:
            self.__component.TextUserDefinedAttributes = value  # type: ignore

    @property
    def unvisited_char_style_name(self) -> str | None:
        """
        Gets/Sets - This optional property contains the character style name for unvisited hyperlinks.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.UnvisitedCharStyleName
        return None

    @unvisited_char_style_name.setter
    def unvisited_char_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.UnvisitedCharStyleName = value

    @property
    def visited_char_style_name(self) -> str | None:
        """
        Gets/Sets - This optional property contains the character style name for visited hyperlinks.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.VisitedCharStyleName
        return None

    @visited_char_style_name.setter
    def visited_char_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.VisitedCharStyleName = value

    # endregion CharacterProperties
