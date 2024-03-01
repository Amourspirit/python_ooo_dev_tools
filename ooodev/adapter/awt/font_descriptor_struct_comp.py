from __future__ import annotations
from typing import TYPE_CHECKING
from ooo.dyn.awt.char_set import CharSetEnum
from ooo.dyn.awt.font_descriptor import FontDescriptor
from ooo.dyn.awt.font_family import FontFamilyEnum
from ooo.dyn.awt.font_pitch import FontPitchEnum
from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.awt.font_type import FontTypeEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum
from ooodev.adapter.struct_base import StructBase
from ooodev.units.angle10 import Angle10

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.
# Example:
# fd = self.__copy()
# fd.Name = event_args.value
# self.component = fd


class FontDescriptorStructComp(StructBase[FontDescriptor]):
    """
    Font Descriptor Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_awt_FontDescriptor_changing``.
    The event raised after the property is changed is called ``com_sun_star_awt_FontDescriptor_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: FontDescriptor, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (FontDescriptor): Font Descriptor.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsPartial | None): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_awt_FontDescriptor_changed"

    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_awt_FontDescriptor_changing"

    def _copy(self, src: FontDescriptor | None = None) -> FontDescriptor:
        if src is None:
            src = self.component
        return FontDescriptor(
            Name=src.Name,
            Height=src.Height,
            StyleName=src.StyleName,
            Family=src.Family,
            CharSet=src.CharSet,
            Pitch=src.Pitch,
            CharacterWidth=src.CharacterWidth,
            Width=src.Width,
            Weight=src.Weight,
            Slant=src.Slant,
            Strikeout=src.Strikeout,
            Orientation=src.Orientation,
            Underline=src.Underline,
            Kerning=src.Kerning,
            WordLineMode=src.WordLineMode,
            Type=src.Type,
        )

    # endregion Overrides

    # region Properties

    @property
    def name(self) -> str:
        """
        Gets/Sets the exact name of the font.
        """
        return self.component.Name

    @name.setter
    def name(self, value: str) -> None:
        old_value = self.component.Name
        if old_value != value:
            event_args = self._trigger_cancel_event("Name", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def height(self) -> int:
        """
        Gets/Sets the height of the font in the measure of the destination.
        """
        return self.component.Height

    @height.setter
    def height(self, value: int) -> None:
        old_value = self.component.Height
        if old_value != value:
            event_args = self._trigger_cancel_event("Height", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def width(self) -> int:
        """
        Gets/Sets the width of the font in the measure of the destination.
        """
        return self.component.Width

    @width.setter
    def width(self, value: int) -> None:
        old_value = self.component.Width
        if old_value != value:
            event_args = self._trigger_cancel_event("Width", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def style_name(self) -> str:
        """
        Gets/Sets the style name of the font.
        """
        return self.component.StyleName

    @style_name.setter
    def style_name(self, value: str) -> None:
        old_value = self.component.StyleName
        if old_value != value:
            event_args = self._trigger_cancel_event("StyleName", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def family(self) -> FontFamilyEnum:
        """
        Gets/Sets the general style of the font.

        When setting the value, the value can be an integer or a ``FontFamilyEnum`` object.

        Hint:
            - ``FontFamilyEnum`` can be imported from ``oo.dyn.awt.font_family``.
        """
        return FontFamilyEnum(self.component.Family)

    @family.setter
    def family(self, value: int | FontFamilyEnum) -> None:
        old_value = FontFamilyEnum(self.component.Family)
        new_value = FontFamilyEnum(value)
        if old_value.value != new_value.value:
            event_args = self._trigger_cancel_event("Family", old_value.value, new_value.value)
            _ = self._trigger_done_event(event_args)

    @property
    def char_set(self) -> CharSetEnum:
        """
        Gets/Sets the character set which is supported by the font.

        When setting the value, the value can be an integer or a ``CharSetEnum`` object.

        Hint:
            - ``CharSetEnum`` can be imported from ``ooo.dyn.awt.char_set``.
        """
        return CharSetEnum(self.component.CharSet)

    @char_set.setter
    def char_set(self, value: int | CharSetEnum) -> None:
        old_value = CharSetEnum(self.component.CharSet)
        new_value = CharSetEnum(value)
        if old_value.value != new_value.value:
            event_args = self._trigger_cancel_event("CharSet", old_value.value, new_value.value)
            _ = self._trigger_done_event(event_args)

    @property
    def pitch(self) -> FontPitchEnum:
        """
        Gets/Sets the pitch of the font.

        When setting the value, the value can be an integer or a ``FontPitchEnum`` object.

        Hint:
            - ``FontPitchEnum`` can be imported from ``ooo.dyn.awt.font_pitch``.
        """
        return FontPitchEnum(self.component.Pitch)

    @pitch.setter
    def pitch(self, value: int | FontPitchEnum) -> None:
        old_value = FontPitchEnum(self.component.Pitch)
        new_value = FontPitchEnum(value)
        if old_value.value != new_value.value:
            event_args = self._trigger_cancel_event("Pitch", old_value.value, new_value.value)
            _ = self._trigger_done_event(event_args)

    @property
    def character_width(self) -> float:
        """
        Gets/Sets the character width.

        Depending on the specified width, a font that supports this width may be selected.

        The value is expressed as a percentage.
        """
        return self.component.CharacterWidth

    @character_width.setter
    def character_width(self, value: float) -> None:
        old_value = self.component.CharacterWidth
        if old_value != value:
            event_args = self._trigger_cancel_event("CharacterWidth", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def weight(self) -> float:
        """
        Gets/Sets the thickness of the line.

        Depending on the specified weight, a font that supports this thickness may be selected.

        The value is expressed as a percentage. ``150.0`` usually means bold.
        """
        return self.component.Weight

    @weight.setter
    def weight(self, value: float) -> None:
        old_value = self.component.Weight
        if old_value != value:
            event_args = self._trigger_cancel_event("Weight", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def slant(self) -> FontSlant:
        """
        Gets/Sets the slant of the font.

        Hint:
            - ``FontSlant`` is a ``ooo.dyn.awt.font_slant``.
        """
        return self.component.Slant

    @slant.setter
    def slant(self, value: FontSlant) -> None:
        old_value = self.component.Slant
        if old_value.value != value.value:
            event_args = self._trigger_cancel_event("Slant", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def underline(self) -> FontUnderlineEnum:
        """
        Gets/Sets the kind of underlining.

        When setting the value, the value can be an integer or a ``FontUnderlineEnum`` object.

        Hint:
            ``FontUnderlineEnum`` can be imported from ``ooo.dyn.awt.font_underline``.
        """
        return FontUnderlineEnum(self.component.Underline)

    @underline.setter
    def underline(self, value: int | FontUnderlineEnum) -> None:
        old_value = FontUnderlineEnum(self.component.Underline)
        new_value = FontUnderlineEnum(value)
        if old_value.value != new_value.value:
            event_args = self._trigger_cancel_event("Underline", old_value.value, new_value.value)
            _ = self._trigger_done_event(event_args)

    @property
    def strikeout(self) -> FontStrikeoutEnum:
        """
        Gets/Sets the kind of strikeout.

        When setting the value, the value can be an integer or a ``FontStrikeoutEnum`` object.

        Hint:
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``.
        """
        return FontStrikeoutEnum(self.component.Strikeout)

    @strikeout.setter
    def strikeout(self, value: int | FontStrikeoutEnum) -> None:
        old_value = FontStrikeoutEnum(self.component.Strikeout)
        new_value = FontStrikeoutEnum(value)
        if old_value.value != new_value.value:
            event_args = self._trigger_cancel_event("Strikeout", old_value.value, new_value.value)
            _ = self._trigger_done_event(event_args)

    @property
    def orientation(self) -> Angle10:
        """
        Gets/Sets the rotation of the font.

        The unit of measure is ``1/10 degrees``; ``0`` is the baseline.

        When setting the value, the value can be an float in ``1/10 degrees`` or an ``Angle10`` object.

        Returns:
            Angle10: Rotation of the font.
        """
        return Angle10(round(self.component.Orientation * 10))

    @orientation.setter
    def orientation(self, value: float | Angle10) -> None:
        old_value = self.component.Orientation
        if isinstance(value, Angle10):
            new_value = 0 if value.value == 0 else value.value / 10
        else:
            new_value = 0 if value == 0 else value / 10
        if old_value != value:
            event_args = self._trigger_cancel_event("Orientation", old_value, new_value)
            _ = self._trigger_done_event(event_args)

    @property
    def kerning(self) -> bool:
        """
        For requesting, it specifies if there is a kerning table available.

        For selecting, it specifies if the kerning table is to be used.
        """
        return self.component.Kerning

    @kerning.setter
    def kerning(self, value: bool) -> None:
        old_value = self.component.Kerning
        if old_value != value:
            event_args = self._trigger_cancel_event("Kerning", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def word_line_mode(self) -> bool:
        """
        Gets/Sets if only words get underlined.

        ``True`` means that only non-space characters get underlined, ``False`` means that the spacing also gets underlined.

        This property is only valid if the property ``com.sun.star.awt.FontDescriptor.Underline`` is not ``FontUnderline.NONE``.
        """
        return self.component.WordLineMode

    @word_line_mode.setter
    def word_line_mode(self, value: bool) -> None:
        old_value = self.component.WordLineMode
        if old_value != value:
            event_args = self._trigger_cancel_event("WordLineMode", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def type(self) -> FontTypeEnum:
        """
        Gets/Sets the technology of the font representation. Flags Enum.

        One or more values out of the constant group ``com.sun.star.awt.FontType`` can be combined by an arithmetical or-operation.

        When setting the value, the value can be an integer or a ``FontTypeEnum`` object.

        Hint:
            - ``FontTypeEnum`` can be imported from ``ooo.dyn.awt.font_type``.
        """
        return FontTypeEnum(self.component.Type)

    @type.setter
    def type(self, value: int | FontTypeEnum) -> None:
        old_value = FontTypeEnum(self.component.Type)
        new_value = FontTypeEnum(value)
        if old_value.value != new_value.value:
            event_args = self._trigger_cancel_event("Type", old_value.value, new_value.value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
