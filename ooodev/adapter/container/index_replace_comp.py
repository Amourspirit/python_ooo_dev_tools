from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.index_replace_partial import IndexReplacePartial

if TYPE_CHECKING:
    from com.sun.star.container import XIndexReplace


class IndexReplaceComp(ComponentProp, IndexReplacePartial):
    """
    Class for managing XIndexReplace Component.
    """

    def __init__(self, component: XIndexReplace) -> None:
        """
        Constructor

        Args:
            component (XIndexReplace): UNO Component that implements ``com.sun.star.container.XIndexReplace``.
        """

        ComponentProp.__init__(self, component)
        IndexReplacePartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XIndexReplace:
        """XIndexReplace Component"""
        # overrides base property
        # pylint: disable=no-member
        return cast("XIndexReplace", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
