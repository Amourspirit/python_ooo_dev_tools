from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.adapter.table.table_rows_partial import TableRowsPartial

if TYPE_CHECKING:
    from com.sun.star.text import TableRows  # service
    from com.sun.star.table import XTableRows


class TableRowsComp(IndexAccessComp, TableRowsPartial):
    """
    Class for managing TableRows Component.

    Provides methods to access rows via index and to insert and remove rows.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableRows) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO Component that implements ``com.sun.star.text.TableRows`` service.
        """
        IndexAccessComp.__init__(self, component)
        TableRowsPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TableRows",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TableRows:
        """TableRows Component"""
        # pylint: disable=no-member
        return cast("TableRows", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
