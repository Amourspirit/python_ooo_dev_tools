from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.name_access_partial import NameAccessPartial
from ooodev.adapter.container.element_access_partial import ElementAccessPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from com.sun.star.table import TableColumns  # service
    from com.sun.star.table import XTableColumns


class TableColumnsComp(ComponentBase, NameAccessPartial, EnumerationAccessPartial, IndexAccessPartial):
    """
    Class for managing TableColumns Component.

    Provides methods to access columns via index and to insert and remove columns.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableColumns) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO Component that implements ``com.sun.star.table.TableColumns`` service.
        """
        ComponentBase.__init__(self, component)
        ElementAccessPartial.__init__(self, component=self.component)
        NameAccessPartial.__init__(self, component=self.component)
        EnumerationAccessPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.TableColumns",)

    # endregion Overrides
    # region XTableColumns
    def insert_by_index(self, index: int, count: int) -> None:
        """
        Inserts a new column at the specified index.

        Args:
            index (int): The index at which the column will be inserted.
            count (int): The number of columns to insert.
        """
        self.component.insertByIndex(index, count)

    def remove_by_index(self, index: int, count: int) -> None:
        """
        Removes columns from the specified index.

        Args:
            index (int): The index at which the column will be removed.
            count (int): The number of columns to remove.
        """
        self.component.removeByIndex(index, count)

    # endregion XTableColumns
    # region Properties
    @property
    def component(self) -> TableColumns:
        """Cell Component"""
        # pylint: disable=no-member
        return cast("TableColumns", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
