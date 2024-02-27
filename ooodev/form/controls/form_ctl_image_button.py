from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from pathlib import Path
import uno

from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum as ImageScaleModeEnum
from ooodev.adapter.form.approve_action_events import ApproveActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.file_io import FileIO
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.form_component_kind import FormComponentKind

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import ImageButton as ControlModel  # service
    from com.sun.star.form.control import ImageButton as ControlView  # service
    from ooodev.utils.type_var import PathOrStr
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlImageButton(FormCtlBase, ApproveActionEvents):
    """``com.sun.star.form.component.ImageButton`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.ImageButton`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            None:

        Note:
            If the :ref:`LoContext <ooodev.utils.context.lo_context.LoContext>` manager is use before this class is instantiated,
            then the Lo instance will be set using the current Lo instance. That the context manager has set.
            Generally speaking this means that there is no need to set ``lo_inst`` when instantiating this class.

        See Also:
            :ref:`ooodev.form.Forms`.
        """
        FormCtlBase.__init__(self, ctl=ctl, lo_inst=lo_inst)
        generic_args = self._get_generic_args()
        ApproveActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_approve_action_add_remove)

    # region Lazy Listeners

    def _on_approve_action_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addApproveActionListener(self.events_listener_approve_action)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_view(self) -> ControlView:
            """Gets the view of this control"""
            return cast("ControlView", super().get_view())

        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.IMAGE_BUTTON

    # endregion Overrides

    # region Properties
    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def help_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """Gets/Sets the help url"""
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def image_scale_mode(self) -> ImageScaleModeEnum:
        """Gets/Sets the Image ScaleMode"""
        return ImageScaleModeEnum(self.model.ScaleMode)

    @image_scale_mode.setter
    def image_scale_mode(self, value: ImageScaleModeEnum) -> None:
        """Sets the Image ScaleMode"""
        self.model.ScaleMode = value.value

    @property
    def image_url(self) -> str:
        """
        Gets/Sets the URL for the image

        Returns:
            str: The URL for the image or empty string if no image is set

        Note:
            It recommended to use the ``picture`` property instead of ``image_url`` property.
            The ``image_url`` property is does not provide any validation.
        """
        return self.model.ImageURL

    @image_url.setter
    def image_url(self, value: str) -> None:
        """Sets the URL for the image"""
        self.model.ImageURL = value

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def picture(self) -> str:
        """
        Gets/Sets the picture for the control

        When setting the value it can be a string or a Path object.
        If a string is passed it can be a URL or a path to a file.
        Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
        Relative paths are supported.

        Returns:
            str: The picture URL in the format of ``file:///path/to/image.png`` or empty string if no picture is set.

        Note:
            It recommended to use the ``picture`` property instead of ``image_url`` property.
            The ``image_url`` property is does not provide any validation.

        Warning:
            If an error occurs when setting the value the ``picture`` then the ``model.ImageURL`` property will be set to an empty string.
        """
        return self.model.ImageURL

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
        try:
            pth_str = str(value)
            if pth_str == "":
                self.model.ImageURL = ""
                return
            if isinstance(value, str):
                if value.startswith("file://"):
                    self.model.ImageURL = value
                    return
                value = Path(value)
            if not FileIO.is_valid_path_or_str(value):
                raise ValueError(f"Invalid path or str: {value}")
            self.model.ImageURL = FileIO.fnm_to_url(value)
        except:
            self.model.ImageURL = ""

    @property
    def printable(self) -> bool:
        """Gets/Sets the printable property"""
        return self.model.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.model.Printable = value

    @property
    def scale_image(self) -> bool:
        """Gets/Sets whether the image should be scaled"""
        return self.model.ScaleImage

    @scale_image.setter
    def scale_image(self, value: bool) -> None:
        """Sets whether the image should be scaled"""
        self.model.ScaleImage = value

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
