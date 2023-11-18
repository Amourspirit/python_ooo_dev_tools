# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
from pathlib import Path
import contextlib
import uno  # pylint: disable=unused-import

# pylint: disable=useless-import-alias
from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum as ImageScaleModeEnum

from ooodev.utils.file_io import FileIO
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.type_var import PathOrStr
from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlImageControl  # service
    from com.sun.star.awt import UnoControlImageControlModel  # service
# endregion imports


class CtlImage(DialogControlBase):
    """Class for Image Control"""

    # region init
    def __init__(self, ctl: UnoControlImageControl) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlImageControl): Image Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)

    # endregion init

    # region Overrides
    def get_view_ctl(self) -> UnoControlImageControl:
        return cast("UnoControlImageControl", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlImageControl``"""
        return "com.sun.star.awt.UnoControlImageControl"

    def get_model(self) -> UnoControlImageControlModel:
        """Gets the Model for the control"""
        return cast("UnoControlImageControlModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.IMAGE``"""
        return DialogControlKind.IMAGE

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.IMAGE``"""
        return DialogControlNamedKind.IMAGE

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlImageControl:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlImageControlModel:
        return self.get_model()

    @property
    def image_url(self) -> str:
        """Gets/Sets the URL for the image"""
        return self.model.ImageURL

    @image_url.setter
    def image_url(self, value: str) -> None:
        """Sets the URL for the image"""
        self.model.ImageURL = value

    @property
    def image_scale_mode(self) -> ImageScaleModeEnum:
        """Gets/Sets the Image ScaleMode"""
        return ImageScaleModeEnum(self.model.ScaleMode)

    @image_scale_mode.setter
    def image_scale_mode(self, value: ImageScaleModeEnum) -> None:
        """Sets the Image ScaleMode"""
        self.model.ScaleMode = value.value

    @property
    def scale_image(self) -> bool:
        """Gets/Sets whether the image should be scaled"""
        return self.model.ScaleImage

    @scale_image.setter
    def scale_image(self, value: bool) -> None:
        """Sets whether the image should be scaled"""
        self.model.ScaleImage = value

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
        """
        with contextlib.suppress(Exception):
            return self.model.ImageURL
        return ""

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
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

    # endregion Properties
