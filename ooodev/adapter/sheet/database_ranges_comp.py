from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.sheet.database_ranges_partial import DatabaseRangesPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial
from ooodev.adapter.lang.type_provider_partial import TypeProviderPartial
from ooodev.adapter.sheet.database_range_comp import DatabaseRangeComp


if TYPE_CHECKING:
    from com.sun.star.sheet import DatabaseRanges  # service


class DatabaseRangesComp(
    ComponentBase,
    DatabaseRangesPartial,
    EnumerationAccessPartial[DatabaseRangeComp],
    IndexAccessPartial[DatabaseRangeComp],
    ServiceInfoPartial,
    TypeProviderPartial,
):
    """
    Class for managing Database Ranges Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DatabaseRanges) -> None:
        """
        Constructor

        Args:
            component (DatabaseRanges): UNO Sheet Cell Component
        """
        ComponentBase.__init__(self, component)
        DatabaseRangesPartial.__init__(self, component=component, interface=None)
        EnumerationAccessPartial.__init__(self, component=component, interface=None)
        IndexAccessPartial.__init__(self, component=component, interface=None)
        ServiceInfoPartial.__init__(self, component=component, interface=None)  # type: ignore
        TypeProviderPartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.DatabaseRanges",)

    def get_by_name(self, name: str) -> DatabaseRangeComp:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            Any: The element with the specified name.
        """
        result = self.component.getByName(name)
        if result is None:
            return None  # type: ignore
        return DatabaseRangeComp(result)  # type: ignore

    def get_by_index(self, idx: int) -> DatabaseRangeComp:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element.

        Returns:
            Any: The element at the specified index.
        """
        result = self.component.getByIndex(idx)
        if result is None:
            return None  # type: ignore
        return DatabaseRangeComp(result)  # type: ignore

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> DatabaseRanges:
        """Sheet Cell Component"""
        # pylint: disable=no-member
        return cast("DatabaseRanges", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
