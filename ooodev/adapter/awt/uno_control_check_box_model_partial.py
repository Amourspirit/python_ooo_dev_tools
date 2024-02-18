from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.image_position import ImagePositionEnum
from ooo.dyn.awt.visual_effect import VisualEffectEnum
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.tri_state_kind import TriStateKind
from ooodev.utils.color import Color
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCheckBoxModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from com.sun.star.graphic import XGraphic
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlCheckBoxModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlCheckBoxModel."""

    def __init__(self, component: UnoControlCheckBoxModel):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlCheckBoxModel`` service.
        """
        # pylint: disable=unused-argument
        self.__component = component
        event_provider = self if isinstance(self, EventsPartial) else None
        UnoControlModelPartial.__init__(self, component=component)
        self.__font_descriptor = FontDescriptorComp(self.__component.FontDescriptor, event_provider)

        if event_provider is not None:

            def on_font_descriptor_changed(src: Any, event_args: KeyValArgs) -> None:
                self.__component.FontDescriptor = self.__font_descriptor.component

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
    def align(self) -> AlignKind:
        """
        Get/Sets the horizontal alignment of the text in the control.

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        return AlignKind(self.__component.Align)

    @align.setter
    def align(self, value: AlignKind | int) -> None:
        kind = AlignKind(int(value))
        self.__component.Align = kind.value

    @property
    def background_color(self) -> Color:
        """
        Gets/Set the background color of the control.
        """
        return Color(self.__component.BackgroundColor)

    @background_color.setter
    def background_color(self, value: Color) -> None:
        self.__component.BackgroundColor = value  # type: ignore

    @property
    def enabled(self) -> bool:
        """
        Gets/Sets whether the control is enabled or disabled.
        """
        return self.__component.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.__component.Enabled = value

    @property
    def font_emphasis_mark(self) -> FontEmphasisEnum:
        """
        Gets/Sets the ``FontEmphasis`` value of the text in the control.

        Note:
            Value can be set with ``FontEmphasisEnum`` or ``int``.

        Hint:
            - ``FontEmphasisEnum`` can be imported from ``ooo.dyn.text.font_emphasis``.
        """
        return FontEmphasisEnum(self.__component.FontEmphasisMark)

    @font_emphasis_mark.setter
    def font_emphasis_mark(self, value: int | FontEmphasisEnum) -> None:
        self.__component.FontEmphasisMark = int(value)

    @property
    def font_relief(self) -> FontReliefEnum:
        """
        Gets/Sets ``FontRelief`` value of the text in the control.

        Note:
            Value can be set with ``FontReliefEnum`` or ``int``.

        Hint:
            - ``FontReliefEnum`` can be imported from ``ooo.dyn.text.font_relief``.
        """
        return FontReliefEnum(self.__component.FontRelief)

    @font_relief.setter
    def font_relief(self, value: int | FontReliefEnum) -> None:
        self.__component.FontRelief = int(value)

    @property
    def graphic(self) -> XGraphic:
        """
        specifies a graphic to be displayed at the button

        If this property is present, it interacts with the ``image_url`` in the following way:

        - If ``image_url`` is set, ``graphic`` will be reset to an object as loaded from the given image URL, or None if ``image_url`` does not point to a valid image file.
        - If ``graphic`` is set, ``image_url`` will be reset to an empty string.
        """
        return self.__component.Graphic

    @graphic.setter
    def graphic(self, value: XGraphic) -> None:
        self.__component.Graphic = value

    @property
    def help_text(self) -> str:
        """
        Get/Sets the help text of the control.
        """
        return self.__component.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.__component.HelpText = value

    @property
    def help_url(self) -> str:
        """
        Gets/Sets the help URL of the control.
        """
        return self.__component.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.__component.HelpURL = value

    @property
    def image_position(self) -> ImagePositionEnum:
        """
        Gets/Sets the position of the image, if any, relative to the text, if any

        Valid values of this property are specified with image_position.

        If this property is present, it supersedes the ImageAlign property - setting one of both properties sets the other one to the best possible match.

        Note:
            Value can be set with ``ImagePositionEnum`` or ``int``.

        Hint:
            - ``ImagePositionEnum`` can be imported from ``ooo.dyn.awt.image_position``
        """
        return ImagePositionEnum(self.__component.ImagePosition)

    @image_position.setter
    def image_position(self, value: int | ImagePositionEnum) -> None:
        self.__component.ImagePosition = int(value)

    @property
    def image_url(self) -> str:
        """
        Gets/Sets a URL to an image to use for the button.
        """
        return self.__component.ImageURL

    @image_url.setter
    def image_url(self, value: str) -> None:
        self.__component.ImageURL = value

    @property
    def label(self) -> str:
        """
        Gets/Sets the label of the control.
        """
        return self.__component.Label

    @label.setter
    def label(self, value: str) -> None:
        self.__component.Label = value

    @property
    def multi_line(self) -> bool:
        """
        Gets/Sets that the text may be displayed on more than one line.
        """
        return self.__component.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.__component.MultiLine = value

    @property
    def printable(self) -> bool:
        """
        Gets/Sets that the control will be printed with the document.
        """
        return self.__component.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.__component.Printable = value

    @property
    def state(self) -> TriStateKind:
        """
        Gets/Sets the state of the control.

        If Toggle property is set to ``True``, the pressed state is enabled and its pressed state can be obtained with this property.

        Note:
            Value can be set with ``TriStateKind`` or ``int``.

        Hint:
            - ``TriStateKind`` can be imported from ``ooo.dyn.kind.tri_state_kind``
        """
        return TriStateKind(self.__component.State)

    @state.setter
    def state(self, value: int | TriStateKind) -> None:
        self.__component.State = int(value)

    @property
    def tabstop(self) -> bool:
        """
        Gets/Sets that the control can be reached with the TAB key.
        """
        return self.__component.Tabstop

    @tabstop.setter
    def tabstop(self, value: bool) -> None:
        self.__component.Tabstop = value

    @property
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.
        """
        return Color(self.__component.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.__component.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color of the control.
        """
        return Color(self.__component.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.__component.TextLineColor = value  # type: ignore

    @property
    def tri_state(self) -> bool:
        """
        Gets/Sets that the control may have the state ``don't know``.
        """
        return self.__component.TriState

    @tri_state.setter
    def tri_state(self, value: bool) -> None:
        self.__component.TriState = value

    @property
    def vertical_align(self) -> VerticalAlignment:
        """
        specifies the vertical alignment of the text in the control.

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        return self.__component.VerticalAlign  # type: ignore

    @vertical_align.setter
    def vertical_align(self, value: VerticalAlignment) -> None:
        self.__component.VerticalAlign = value  # type: ignore

    @property
    def visual_effect(self) -> VisualEffectEnum:
        """
        specifies a visual effect to apply to the check box control

        Possible values for this property are VisualEffect.FLAT and VisualEffect.LOOK3D.

        Note:
            Value can be set with ``VisualEffectEnum`` or ``int``.

        Hint:
            - ``VisualEffectEnum`` can be imported from ``ooo.dyn.awt.visual_effect``
        """
        return VisualEffectEnum(self.__component.VisualEffect)

    @visual_effect.setter
    def visual_effect(self, value: int | VisualEffectEnum) -> None:
        self.__component.VisualEffect = int(value)

    @property
    def writing_mode(self) -> int:
        """
        Denotes the writing mode used in the control, as specified in the ``com.sun.star.text.WritingMode2`` constants group.

        Only LR_TB (``0``) and RL_TB (``1``) are supported at the moment.
        """
        return self.__component.WritingMode

    @writing_mode.setter
    def writing_mode(self, value: int) -> None:
        self.__component.WritingMode = value

    # endregion Properties
