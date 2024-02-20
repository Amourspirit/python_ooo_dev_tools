from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooo.dyn.text.font_emphasis import FontEmphasisEnum
from ooo.dyn.text.font_relief import FontReliefEnum
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.utils.color import Color
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from .uno_control_model_partial import UnoControlModelPartial
from .font_descriptor_comp import FontDescriptorComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedLineModel  # Service
    from com.sun.star.awt import FontDescriptor  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class UnoControlFixedLineModelPartial(UnoControlModelPartial):
    """Partial class for UnoControlFixedLineModel."""

    def __init__(self):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFixedLineModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: UnoControlFixedLineModel
        event_provider = self if isinstance(self, EventsPartial) else None
        UnoControlModelPartial.__init__(self, component=self.model)
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
        self.model.FontDescriptor = font_descriptor

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
    def label(self) -> str:
        """
        Gets/Sets the label of the control.
        """
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    @property
    def orientation(self) -> OrientationKind:
        """
        Gets/Sets the orientation of the control.

        Note:
            Value can be set with ``OrientationKind`` or ``int``.

        Hint:
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        return OrientationKind(self.model.Orientation)

    @orientation.setter
    def orientation(self, value: int | OrientationKind) -> None:
        self.model.Orientation = int(value)

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
    def text_color(self) -> Color:
        """
        Gets/Sets the text color of the control.
        """
        return Color(self.model.TextColor)

    @text_color.setter
    def text_color(self, value: Color) -> None:
        self.model.TextColor = value  # type: ignore

    @property
    def text_line_color(self) -> Color:
        """
        Gets/Sets the text line color (RGB) of the control.
        """
        return Color(self.model.TextLineColor)

    @text_line_color.setter
    def text_line_color(self, value: Color) -> None:
        self.model.TextLineColor = value  # type: ignore

    # endregion Properties
