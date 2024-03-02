from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextTables  # service


class TextTablesComp(IndexAccessComp, NameAccessPartial):
    """
    Class for managing Text Tables Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.text.Text`` service.
        """

        IndexAccessComp.__init__(self, component)
        NameAccessPartial.__init__(self, component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextTables",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> TextTables:
        """TextTables Component"""
        # pylint: disable=no-member
        return cast("TextTables", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
