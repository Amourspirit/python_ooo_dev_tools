from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.name_replace_partial import NameReplacePartial

if TYPE_CHECKING:
    from com.sun.star.container import XNameReplace


class NameReplaceComp(ComponentBase, NameReplacePartial):
    """
    Class for managing XNameContainer Component.
    """

    def __init__(self, component: XNameReplace) -> None:
        """
        Constructor

        Args:
            component (XNameReplace): UNO Component that implements ``com.sun.star.container.XNameReplace``.
        """

        ComponentBase.__init__(self, component)
        NameReplacePartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XNameReplace:
        """XNameReplace Component"""
        # pylint: disable=no-member
        return cast("XNameReplace", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
