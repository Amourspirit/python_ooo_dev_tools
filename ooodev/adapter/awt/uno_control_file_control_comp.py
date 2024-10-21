from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFileControl


class UnoControlFileControlComp(UnoControlEditComp):

    def __init__(self, component: UnoControlFileControl):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFileControl`` service.
        """
        UnoControlEditComp.__init__(self, component=component)

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlFileControl",)

    @property
    @override
    def component(self) -> UnoControlFileControl:
        """UnoControlFileControl Component"""
        # pylint: disable=no-member
        return cast("UnoControlFileControl", self._ComponentBase__get_component())  # type: ignore
