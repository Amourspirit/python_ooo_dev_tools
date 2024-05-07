from __future__ import annotations
from typing import cast
import uno
from com.sun.star.io import XOutputStream

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.input_stream_partial import InputStreamPartial


class OutputStreamComp(ComponentProp, InputStreamPartial):
    """
    Class for managing XOutputStream Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XOutputStream) -> None:
        """
        Constructor

        Args:
            component (XOutputStream): UNO Component that implements ``com.sun.star.io.XOutputStream`` interface.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        InputStreamPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> XOutputStream:
        """XOutputStream Component"""
        # pylint: disable=no-member
        return cast("XOutputStream", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
