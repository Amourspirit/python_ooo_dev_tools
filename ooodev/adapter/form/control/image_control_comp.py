from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_image_control_comp import UnoControlImageControlComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import ImageControl


class ImageControlComp(UnoControlImageControlComp, BoundControlPartial):
    """Class for ImageControl Control"""

    def __init__(self, component: ImageControl):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.ImageControl`` service.
        """
        UnoControlImageControlComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.ImageControl",)

    @property
    def component(self) -> ImageControl:
        """ImageControl Component"""
        # pylint: disable=no-member
        return cast("ImageControl", self._ComponentBase__get_component())  # type: ignore
