from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.drawing import ShadowProperties  # service


class ShadowPropertiesComp(ComponentBase):
    """
    Class for managing table ShadowProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ShadowProperties) -> None:
        """
        Constructor

        Args:
            component (ShadowProperties): UNO ShadowProperties Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.ShadowProperties",)

    # endregion Overrides
    # region Properties
    @property
    @override
    def component(self) -> ShadowProperties:
        """ShadowProperties Component"""
        # pylint: disable=no-member
        return cast("ShadowProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
