# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum as ImageScaleModeEnum

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlImageControl  # service
    from com.sun.star.awt import UnoControlImageControlModel  # service
# endregion imports


class CtlImage(CtlBase):
    """Class for Image Control"""

    # region init
    def __init__(self, ctl: UnoControlImageControl) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlImageControl): Image Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)

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

    # endregion Properties
