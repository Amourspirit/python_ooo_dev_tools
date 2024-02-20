# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
from pathlib import Path
import contextlib
import uno  # pylint: disable=unused-import

# pylint: disable=useless-import-alias
from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum as ImageScaleModeEnum

from ooodev.adapter.awt.uno_control_image_control_model_partial import UnoControlImageControlModelPartial
from ooodev.utils.file_io import FileIO
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlImageControl  # service
    from com.sun.star.awt import UnoControlImageControlModel  # service
    from ooodev.utils.type_var import PathOrStr
# endregion imports


class CtlImage(DialogControlBase, UnoControlImageControlModelPartial):
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
        UnoControlImageControlModelPartial.__init__(self)

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
    def image_scale_mode(self) -> ImageScaleModeEnum:
        """
        Gets/Sets the Image ScaleMode.

        Same as ``scale_mode`` property.

        Note:
            Value can be set with ``ImageScaleModeEnum`` or ``int``.

        Hint:
            - ``ImageScaleModeEnum`` can be imported from ``ooo.dyn.awt.image_scale_mode``
        """
        return ImageScaleModeEnum.NONE if self.scale_mode is None else self.scale_mode

    @image_scale_mode.setter
    def image_scale_mode(self, value: int | ImageScaleModeEnum) -> None:
        """Sets the Image ScaleMode"""
        self.scale_mode = value

    # region UnoControlImageControlModelPartial Overrides
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

    # endregion UnoControlImageControlModelPartial Overrides

    @property
    def model(self) -> UnoControlImageControlModel:
        # pylint: disable=no-member
        return cast("UnoControlImageControlModel", super().model)

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
        with contextlib.suppress(Exception):
            return self.model.ImageURL
        return ""

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
        # pylint: disable=broad-exception-caught
        try:
            pth_str = str(value)
            if not pth_str:
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
        except Exception:
            self.model.ImageURL = ""

    @property
    def view(self) -> UnoControlImageControl:
        # pylint: disable=no-member
        return cast("UnoControlImageControl", super().view)

    # endregion Properties
