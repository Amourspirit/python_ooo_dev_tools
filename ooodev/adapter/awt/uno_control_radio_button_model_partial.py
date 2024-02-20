from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING
import uno
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.image_position import ImagePositionEnum
from ooo.dyn.awt.visual_effect import VisualEffectEnum
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.state_kind import StateKind
from ooodev.utils.color import Color
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlRadioButtonModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from com.sun.star.graphic import XGraphic
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlRadioButtonModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlRadioButtonModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlRadioButtonModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlRadioButtonModel
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
    def background_color(self) -> Color | None:
        """
        Gets/Set the background color of the control.

        **optional**

        Returns:
            ~ooodev.utils.color.Color | None: Color or None if not set.
        """
        with contextlib.suppress(AttributeError):
            return Color(self.model.BackgroundColor)
        return None

    @background_color.setter
    def background_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.model.BackgroundColor = value  # type: ignore

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
    def graphic(self) -> XGraphic | None:
        """
        specifies a graphic to be displayed at the button

        If this property is present, it interacts with the ``image_url`` in the following way:

        - If ``image_url`` is set, ``graphic`` will be reset to an object as loaded from the given image URL, or None if ``image_url`` does not point to a valid image file.
        - If ``graphic`` is set, ``image_url`` will be reset to an empty string.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.Graphic
        return None

    @graphic.setter
    def graphic(self, value: XGraphic) -> None:
        with contextlib.suppress(AttributeError):
            self.model.Graphic = value

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
    def image_position(self) -> ImagePositionEnum | None:
        """
        Gets/Sets the position of the image, if any, relative to the text, if any

        Valid values of this property are specified with image_position.

        If this property is present, it supersedes the ImageAlign property - setting one of both properties sets the other one to the best possible match.

        **optional**

        Note:
            Value can be set with ``ImagePositionEnum`` or ``int``.

        Hint:
            - ``ImagePositionEnum`` can be imported from ``ooo.dyn.awt.image_position``
        """
        with contextlib.suppress(AttributeError):
            return ImagePositionEnum(self.model.ImagePosition)
        return None

    @image_position.setter
    def image_position(self, value: int | ImagePositionEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.model.ImagePosition = int(value)

    @property
    def image_url(self) -> str | None:
        """
        Gets/Sets a URL to an image to use for the button.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.ImageURL
        return None

    @image_url.setter
    def image_url(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.model.ImageURL = value

    @property
    def label(self) -> str:
        """
        Gets/Sets the label of the control.
        """
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    @property
    def multi_line(self) -> bool | None:
        """
        Gets/Sets that the text may be displayed on more than one line.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.model.MultiLine
        return None

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.model.MultiLine = value

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
    def state(self) -> StateKind:
        """
        Gets/Sets the state of the control.

        If Toggle property is set to ``True``, the pressed state is enabled and its pressed state can be obtained with this property.

        Note:
            Value can be set with ``StateKind`` or ``int``.

        Hint:
            - ``StateKind`` can be imported from ``ooodev.utils.kind.state_kind``
        """
        return StateKind(self.model.State)

    @state.setter
    def state(self, value: int | StateKind) -> None:
        self.model.State = int(value)

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
    def vertical_align(self) -> VerticalAlignment | None:
        """
        Gets/Sets the vertical alignment of the text in the control.

        **optional**

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        with contextlib.suppress(AttributeError):
            return self.model.VerticalAlign  # type: ignore
        return None

    @vertical_align.setter
    def vertical_align(self, value: VerticalAlignment) -> None:
        with contextlib.suppress(AttributeError):
            self.model.VerticalAlign = value  # type: ignore

    @property
    def visual_effect(self) -> VisualEffectEnum | None:
        """
        specifies a visual effect to apply to the check box control

        Possible values for this property are VisualEffect.FLAT and VisualEffect.LOOK3D.

        **optional**

        Note:
            Value can be set with ``VisualEffectEnum`` or ``int``.

        Hint:
            - ``VisualEffectEnum`` can be imported from ``ooo.dyn.awt.visual_effect``
        """
        with contextlib.suppress(AttributeError):
            return VisualEffectEnum(self.model.VisualEffect)
        return None

    @visual_effect.setter
    def visual_effect(self, value: int | VisualEffectEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.model.VisualEffect = int(value)

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
