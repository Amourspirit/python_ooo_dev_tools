from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooo.dyn.awt.char_set import CharSetEnum
from ooo.dyn.awt.font_descriptor import FontDescriptor
from ooo.dyn.awt.font_family import FontFamilyEnum
from ooo.dyn.awt.font_pitch import FontPitchEnum
from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.awt.font_type import FontTypeEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.adapter.component_base import ComponentBase

if TYPE_CHECKING:
    from ooodev.events.partial.events_partial import EventsPartial

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.
# Example:
# fd = self.__copy()
# fd.Name = event_args.value
# self.component = fd


class FontDescriptorStructComp(ComponentBase):
    """
    Font Descriptor Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``font_descriptor_struct_changing``.
    The event raised after the property is changed is called ``font_descriptor_struct_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: FontDescriptor, event_provider: EventsPartial | None) -> None:
        """
        Constructor

        Args:
            component (FontDescriptor): Font Descriptor.
            event_provider (EventsPartial | None): Event Provider.
        """
        ComponentBase.__init__(self, component)
        self._event_provider = event_provider

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # PropertySetPartial will validate
        return ()

    # endregion Overrides

    def _on_property_changing(self, event_args: KeyValCancelArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event("font_descriptor_struct_changing", event_args)

    def _on_property_changed(self, event_args: KeyValArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event("font_descriptor_struct_changed", event_args)

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

    def copy(self) -> FontDescriptor:
        """
        Makes a copy of the Font Descriptor.

        Returns:
            FontDescriptor: Copied Font Descriptor.
        """
        return self._copy()

    # region Properties

    @property
    def component(self) -> FontDescriptor:
        """FontDescriptor Component"""
        # pylint: disable=no-member
        return cast("FontDescriptor", self._ComponentBase__get_component())  # type: ignore

    @component.setter
    def component(self, value: FontDescriptor) -> None:
        # pylint: disable=no-member
        self._ComponentBase__set_component(self._copy(src=value))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="name",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                fd = self._copy()
                fd.Name = event_args.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="height",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                fd = self._copy()
                fd.Height = event_args.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="width",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                fd = self._copy()
                fd.Width = event_args.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="style_name",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                fd = self._copy()
                fd.StyleName = event_args.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def family(self) -> FontFamilyEnum:
        """
        specifies the general style of the font.

        Hint:
            - ``FontFamilyEnum`` can be imported from ``oo.dyn.awt.font_family``.
        """
        return FontFamilyEnum(self.component.Family)

    @family.setter
    def family(self, value: int | FontFamilyEnum) -> None:
        old_value = FontFamilyEnum(self.component.Family)
        new_value = FontFamilyEnum(value)
        if old_value.value != new_value.value:
            event_args = KeyValCancelArgs(
                source=self,
                key="family",
                value=new_value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(FontFamilyEnum, event_args.value)
                fd = self._copy()
                fd.Family = val.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def char_set(self) -> CharSetEnum:
        """
        Gets/Sets the character set which is supported by the font.

        Hint:
            - ``CharSetEnum`` can be imported from ``ooo.dyn.awt.char_set``.
        """
        return CharSetEnum(self.component.CharSet)

    @char_set.setter
    def char_set(self, value: int | CharSetEnum) -> None:
        old_value = CharSetEnum(self.component.CharSet)
        new_value = CharSetEnum(value)
        if old_value.value != new_value.value:
            event_args = KeyValCancelArgs(
                source=self,
                key="char_set",
                value=new_value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(CharSetEnum, event_args.value)
                fd = self._copy()
                fd.CharSet = val.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def pitch(self) -> FontPitchEnum:
        """
        Gets/Sets the pitch of the font.

        Hint:
            - ``FontPitchEnum`` can be imported from ``ooo.dyn.awt.font_pitch``.
        """
        return FontPitchEnum(self.component.Pitch)

    @pitch.setter
    def pitch(self, value: int | FontPitchEnum) -> None:
        old_value = FontPitchEnum(self.component.Pitch)
        new_value = FontPitchEnum(value)
        if old_value.value != new_value.value:
            event_args = KeyValCancelArgs(
                source=self,
                key="pitch",
                value=new_value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(FontPitchEnum, event_args.value)
                fd = self._copy()
                fd.Pitch = val.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def character_width(self) -> float:
        """
        specifies the character width.

        Depending on the specified width, a font that supports this width may be selected.

        The value is expressed as a percentage.
        """
        return self.component.CharacterWidth

    @character_width.setter
    def character_width(self, value: float) -> None:
        old_value = self.component.CharacterWidth
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="character_width",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                fd = self._copy()
                fd.CharacterWidth = event_args.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def weight(self) -> float:
        """
        Gets/Sets the thickness of the line.

        Depending on the specified weight, a font that supports this thickness may be selected.

        The value is expressed as a percentage.
        """
        return self.component.Weight

    @weight.setter
    def weight(self, value: float) -> None:
        old_value = self.component.Weight
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="weight",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                fd = self._copy()
                fd.Weight = event_args.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="slant",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast("FontSlant", event_args.value)
                fd = self._copy()
                fd.Slant = val
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def underline(self) -> FontUnderlineEnum:
        """
        Gets/Sets the kind of underlining.

        Use one value out of the constant group com.sun.star.awt.FontUnderline.

        Hint:
            ``FontUnderlineEnum`` can be imported from ``ooo.dyn.awt.font_underline``.
        """
        return FontUnderlineEnum(self.component.Underline)

    @underline.setter
    def underline(self, value: int | FontUnderlineEnum) -> None:
        old_value = FontUnderlineEnum(self.component.Underline)
        new_value = FontUnderlineEnum(value)
        if old_value.value != new_value.value:
            event_args = KeyValCancelArgs(
                source=self,
                key="underline",
                value=new_value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(FontUnderlineEnum, event_args.value)
                fd = self._copy()
                fd.Underline = val.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def strikeout(self) -> FontStrikeoutEnum:
        """
        Gets/Sets the kind of strikeout.

        Hint:
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``.
        """
        return FontStrikeoutEnum(self.component.Strikeout)

    @strikeout.setter
    def strikeout(self, value: int | FontStrikeoutEnum) -> None:
        old_value = FontStrikeoutEnum(self.component.Strikeout)
        new_value = FontStrikeoutEnum(value)
        if old_value.value != new_value.value:
            event_args = KeyValCancelArgs(
                source=self,
                key="strikeout",
                value=new_value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(FontStrikeoutEnum, event_args.value)
                fd = self._copy()
                fd.Strikeout = val.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def orientation(self) -> float:
        """
        Gets/Sets the rotation of the font.

        The unit of measure is degrees; 0 is the baseline.
        """
        return self.component.Orientation

    @orientation.setter
    def orientation(self, value: float) -> None:
        old_value = self.component.Orientation
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="orientation",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(float, event_args.value)
                fd = self._copy()
                fd.Orientation = val
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="kerning",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(bool, event_args.value)
                fd = self._copy()
                fd.Kerning = val
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="word_line_mode",
                value=value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(bool, event_args.value)
                fd = self._copy()
                fd.WordLineMode = val
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def type(self) -> FontTypeEnum:
        """
        specifies the technology of the font representation.

        One or more values out of the constant group ``com.sun.star.awt.FontType`` can be combined by an arithmetical or-operation.

        Hint:
            - ``FontTypeEnum`` can be imported from ``ooo.dyn.awt.font_type``.
        """
        return FontTypeEnum(self.component.Type)

    @type.setter
    def type(self, value: int | FontTypeEnum) -> None:
        old_value = FontTypeEnum(self.component.Type)
        new_value = FontTypeEnum(value)
        if old_value.value != new_value.value:
            event_args = KeyValCancelArgs(
                source=self,
                key="type",
                value=new_value,
            )
            event_args.event_data = {"old_value": old_value}
            self._on_property_changing(event_args)
            if not event_args.cancel:
                val = cast(FontTypeEnum, event_args.value)
                fd = self._copy()
                fd.Type = val.value
                self.component = fd
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    # endregion Properties
