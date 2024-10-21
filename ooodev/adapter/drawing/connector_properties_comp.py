from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.connector_properties_partial import ConnectorPropertiesPartial


if TYPE_CHECKING:
    from com.sun.star.drawing import ConnectorProperties  # service


class ConnectorPropertiesComp(ComponentBase, ConnectorPropertiesPartial):
    """
    Class for managing table ConnectorProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ConnectorProperties) -> None:
        """
        Constructor

        Args:
            component (ConnectorProperties): UNO ConnectorProperties Component.
        """
        ComponentBase.__init__(self, component)
        ConnectorPropertiesPartial.__init__(self, component=component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.ConnectorProperties",)

    # endregion Overrides
    # region Properties
    @property
    @override
    def component(self) -> ConnectorProperties:
        """ConnectorProperties Component"""
        # pylint: disable=no-member
        return cast("ConnectorProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
