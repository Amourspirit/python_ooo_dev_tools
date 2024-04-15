from __future__ import annotations
from typing import cast, TYPE_CHECKING, Generic, TypeVar

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess

T = TypeVar("T")


class IndexAccessComp(
    ComponentProp,
    IndexAccessPartial[T],
    Generic[T],
):
    """
    Class for managing XIndexAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexAccess) -> None:
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Component that implements ``com.sun.star.container.XIndexAccess``.
        """

        ComponentProp.__init__(self, component)
        IndexAccessPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XIndexAccess:
        """XIndexAccess Component"""
        # overrides base property
        # pylint: disable=no-member
        return cast("XIndexAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
