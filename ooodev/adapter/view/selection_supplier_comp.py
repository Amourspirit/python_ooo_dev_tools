from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.adapter.view.selection_supplier_partial import SelectionSupplierPartial
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.view import XSelectionSupplier


class SelectionSupplierComp(ComponentBase, SelectionSupplierPartial):
    """
    Class for managing Window Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSelectionSupplier) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports `com.sun.star.view.XSelectionSupplier`` interface.

        Returns:
            None:
        """

        ComponentBase.__init__(self, component)  # type: ignore
        SelectionSupplierPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> XSelectionSupplier:
        """XSelectionSupplier Component"""
        # pylint: disable=no-member
        return cast("XSelectionSupplier", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
