from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.name_access_partial import NameAccessPartial


if TYPE_CHECKING:
    from com.sun.star.style import StyleFamilies  # service
    from com.sun.star.container import XNameAccess


class StyleFamiliesComp(ComponentBase, NameAccessPartial):
    """
    Class for managing StyleFamilies Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameAccess) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that support ``com.sun.star.style.StyleFamilies`` service.
        """
        ComponentBase.__init__(self, component)
        # no need to validate NameAccessPartial will be validated by ComponentBase
        NameAccessPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.StyleFamilies",)

    # endregion Overrides
    # region Properties
    @property
    @override
    def component(self) -> StyleFamilies:
        """StyleFamilies Component"""
        # pylint: disable=no-member
        return cast("StyleFamilies", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
