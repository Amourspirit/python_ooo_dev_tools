from __future__ import annotations
from typing import cast
import uno
from com.sun.star.io import XInputStream

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.input_stream_partial import InputStreamPartial


class InputStreamComp(ComponentProp, InputStreamPartial):
    """
    Class for managing XInputStream Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XInputStream) -> None:
        """
        Constructor

        Args:
            component (XInputStream): UNO Component that implements ``com.sun.star.io.XInputStream`` interface.
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
    def component(self) -> XInputStream:
        """XInputStream Component"""
        # pylint: disable=no-member
        return cast("XInputStream", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
