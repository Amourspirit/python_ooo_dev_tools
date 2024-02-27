from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from com.sun.star.container import XNameAccess


class NameAccessComp(ComponentBase, NameAccessPartial):
    """
    Class for managing XNameAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameAccess) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess``.
        """

        ComponentBase.__init__(self, component)
        NameAccessPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XNameAccess:
        """XNameAccess Component"""
        # pylint: disable=no-member
        return cast("XNameAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
