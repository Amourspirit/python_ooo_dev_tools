from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.graphic.graphic_descriptor_comp import GraphicDescriptorComp
from ooodev.adapter.graphic.graphic_partial import GraphicPartial

if TYPE_CHECKING:
    from com.sun.star.graphic import Graphic  # service


class GraphicComp(GraphicDescriptorComp, GraphicPartial):
    """
    Class for managing Graphic Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Graphic) -> None:
        """
        Constructor

        Args:
            component (Graphic): UNO Component that implements ``com.sun.star.graphic.Graphic`` service.
        """
        GraphicDescriptorComp.__init__(self, component)  # type: ignore
        GraphicPartial.__init__(self, component=component, interface=None)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.graphic.Graphic",)

    # endregion Overrides
    # region Properties
    @property
    @override
    def component(self) -> Graphic:
        """Graphic Component"""
        # pylint: disable=no-member
        return cast("Graphic", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
