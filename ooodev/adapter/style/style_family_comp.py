from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.name_access_partial import NameAccessPartial


if TYPE_CHECKING:
    from com.sun.star.style import StyleFamily  # service
    from com.sun.star.container import XNameAccess


class StyleFamilyComp(ComponentBase, NameAccessPartial):
    """
    Class for managing StyleFamily Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameAccess) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that support ``com.sun.star.style.StyleFamily`` service.
        """
        ComponentBase.__init__(self, component)
        # no need to validate NameAccessPartial will be validated by ComponentBase
        NameAccessPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.StyleFamily",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> StyleFamily:
        """StyleFamily Component"""
        # pylint: disable=no-member
        return cast("StyleFamily", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
