from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.table.table_columns_partial import TableColumnsPartial

if TYPE_CHECKING:
    from com.sun.star.table import TableColumns  # service
    from com.sun.star.table import XTableColumns


class TableColumnsComp(ComponentBase, TableColumnsPartial, EnumerationAccessPartial):
    """
    Class for managing TableColumns Component.

    Provides methods to access columns via index and to insert and remove columns.
    """

    # this class is very similar to ooodev.adapter.text.table_columns_comp.TableColumnsComp
    # don't get them confused.
    # pylint: disable=unused-argument

    def __init__(self, component: XTableColumns) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO Component that implements ``com.sun.star.table.TableColumns`` service.
        """
        ComponentBase.__init__(self, component)
        TableColumnsPartial.__init__(self, component=component, interface=None)
        # Name access is optional
        # NameAccessPartial.__init__(self, component=self.component)
        EnumerationAccessPartial.__init__(self, component=self.component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.TableColumns",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> TableColumns:
        """Table columns Component"""
        # pylint: disable=no-member
        return cast("TableColumns", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
