from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING, Tuple
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.awt.mouse_wheel_behavior import MouseWheelBehaviorEnum
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.color import Color
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlListBoxModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlListBoxModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlListBoxModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlListBoxModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlListBoxModel
        event_provider = self if isinstance(self, EventsPartial) else None
        UnoControlModelPartial.__init__(self)
        self.__font_descriptor = FontDescriptorComp(self.model.FontDescriptor, event_provider)

        if event_provider is not None:

            def on_font_descriptor_changed(src: Any, event_args: KeyValArgs) -> None:
                self.model.FontDescriptor = self.__font_descriptor.component

            self.__fn_on_font_descriptor_changed = on_font_descriptor_changed
            # pylint: disable=no-member
            event_provider.subscribe_event("font_descriptor_struct_changed", self.__fn_on_font_descriptor_changed)

    def set_font_descriptor(self, font_descriptor: FontDescriptor) -> None:
        """
        Sets the font descriptor of the control.

        Args:
            font_descriptor (FontDescriptor): UNO Struct - Font descriptor to set.
        """
        # FontDescriptorComp do not have any state, so we can directly assign the component.
        self.__font_descriptor.component = font_descriptor
        self.model.FontDescriptor = self.__font_descriptor.component

    # region Properties

    @property
    def font_descriptor(self) -> FontDescriptorComp:
        """
        Gets the Font Descriptor

        Hint:
            ``set_font_descriptor()`` can be used to set the font descriptor.
        """
        return self.__font_descriptor

    @property
    def selected_items(self) -> Tuple[int, ...]:
        """Gets specifies the sequence of selected items, identified by the position."""
        # SelectedItems is a miss reported type in current typings.
        # It reports as ``uno.ByteSequence`` but it is actually a ``Tuple[int, ...]``.
        return cast(Tuple[int, ...], self.model.SelectedItems)

    @selected_items.setter
    def selected_items(self, value: Tuple[int, ...]) -> None:
        self.model.SelectedItems = value  # type: ignore

    @property
    def string_item_list(self) -> Tuple[str, ...]:
        """
        specifies the list of items.
        """
        return self.model.StringItemList

    @string_item_list.setter
    def string_item_list(self, value: Tuple[str, ...]) -> None:
        self.model.StringItemList = value

    @property
    def typed_item_list(self) -> Tuple[Any, ...]:
        """
        Gets/Sets the list of raw typed (not stringized) items.

        This list corresponds with the StringItemList and if given has to be of the same length,
        the elements' positions matching those of their string representation in ``string_item_list``.
        """
        return self.model.TypedItemList

    @typed_item_list.setter
    def typed_item_list(self, value: Tuple[Any, ...]) -> None:
        self.model.TypedItemList = value

    @property
    def align(self) -> AlignKind | None:
        """
        Get/Sets the horizontal alignment of the text in the control.

        **optional**

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        with contextlib.suppress(AttributeError):
            return AlignKind(self.model.Align)
        return None

    @align.setter
    def align(self, value: AlignKind | int) -> None:
        kind = AlignKind(int(value))
        with contextlib.suppress(AttributeError):
            self.model.Align = kind.value

    @property
    def background_color(self) -> Color:
        """
        Gets/Set the background color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.model.BackgroundColor)

    @background_color.setter
    def background_color(self, value: Color) -> None:
        self.model.BackgroundColor = value  # type: ignore

    @property
    def border(self) -> BorderKind:
        """
        Gets/Sets the border style.

        Note:
            Value can be set with ``BorderKind`` or ``int``.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``
        """
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: int | BorderKind) -> None:
        self.model.Border = int(value)

    @property
    def border_color(self) -> Color | None:
        """
        Gets/Set the color of the border, if present.

        Not every border style (see Border) may support coloring.
        For instance, usually a border with 3D effect will ignore the ``border_color`` setting.

        **optional**

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        with contextlib.suppress(AttributeError):
            return Color(self.model.BorderColor)
        return None

    @border_color.setter
    def border_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.model.BorderColor = value  # type: ignore

    @property
    def drop_down(self) -> bool:
        """
        Gets/Sets if the control has a drop down button.
        """
        return self.model.Dropdown

    @drop_down.setter
    def drop_down(self, value: bool) -> None:
        self.model.Dropdown = value

    @property
    def enabled(self) -> bool:
        """
        Gets/Sets whether the control is enabled or disabled.
        """
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def font_emphasis_mark(self) -> FontEmphasisEnum:
        """
        Gets/Sets the ``FontEmphasis`` value of the text in the control.

        Note:
            Value can be set with ``FontEmphasisEnum`` or ``int``.

        Hint:
            - ``FontEmphasisEnum`` can be imported from ``ooo.dyn.text.font_emphasis``.
        """
        return FontEmphasisEnum(self.model.FontEmphasisMark)

    @font_emphasis_mark.setter
    def font_emphasis_mark(self, value: int | FontEmphasisEnum) -> None:
        self.model.FontEmphasisMark = int(value)

    @property
    def font_relief(self) -> FontReliefEnum:
        """
        Gets/Sets ``FontRelief`` value of the text in the control.

        Note:
            Value can be set with ``FontReliefEnum`` or ``int``.

        Hint:
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.text.font_relief``.
        """
        return FontReliefEnum(self.model.FontRelief)

    @font_relief.setter
    def font_relief(self, value: int | FontReliefEnum) -> None:
        self.model.FontRelief = int(value)

    @property
    def help_text(self) -> str:
        """
        Get/Sets the help text of the control.
        """
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """
        Gets/Sets the help URL of the control.
        """
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def item_separator_pos(self) -> int | None:
        """
        specifies where an item separator - a horizontal line - is drawn.

        If this is not NULL, then a horizontal line will be drawn between the item at the given position, and the following item.

        **optional**

        **Maybe None**
        """
        return self.model.ItemSeparatorPos

    @item_separator_pos.setter
    def item_separator_pos(self, value: int) -> None:
        self.model.ItemSeparatorPos = value

    @property
    def line_count(self) -> int:
        """
        Gets/Sets the maximum line count displayed in the drop down box.
        """
        return self.model.LineCount

    @line_count.setter
    def line_count(self, value: int) -> None:
        self.model.LineCount = value

    @property
    def mouse_wheel_behavior(self) -> MouseWheelBehaviorEnum:
        """
        Gets/Sets how the mouse wheel can be used to scroll through the control's content.

        Usually, the mouse wheel scroll through the control's entry list.
        Using this property,you can control under which circumstances this is possible.

        Note:
            Value can be set with ``MouseWheelBehaviorEnum`` or ``int``.

        Hint:
            - ``MouseWheelBehaviorEnum`` can be imported from ``ooo.dyn.awt.mouse_wheel_behavior``
        """
        return MouseWheelBehaviorEnum(self.model.MouseWheelBehavior)

    @mouse_wheel_behavior.setter
    def mouse_wheel_behavior(self, value: int | MouseWheelBehaviorEnum) -> None:
        self.model.MouseWheelBehavior = int(value)

    @property
    def multi_selection(self) -> bool:
        """
        Gets/Sets if more than one entry can be selected.
        """
        return self.model.MultiSelection

    @multi_selection.setter
    def multi_selection(self, value: bool) -> None:
        self.model.MultiSelection = value

    @property
    def printable(self) -> bool:
        """
        Gets/Sets that the control will be printed with the document.
        """
        return self.model.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.model.Printable = value

    @property
    def read_only(self) -> bool:
        """
        Gets/Sets that the content of the control cannot be modified by the user.
        """
        return self.model.ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        self.model.ReadOnly = value

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets that the control can be reached with the TAB key.
        """
        return self.model.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.model.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.model.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color of the control.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return Color(self.model.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.model.TextLineColor = value  # type: ignore

    @property
    def writing_mode(self) -> int | None:
        """
        Denotes the writing mode used in the control, as specified in the ``com.sun.star.text.WritingMode2`` constants group.

        Only LR_TB (``0``) and RL_TB (``1``) are supported at the moment.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.WritingMode
        return None

    @writing_mode.setter
    def writing_mode(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.model.WritingMode = value

    # endregion Properties
