from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.table import CellCursor  # service


class CellComp(ComponentBase):
    """
    Class for managing table Cell Cursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellCursor) -> None:
        """
        Constructor

        Args:
            component (CellCursor): UNO table CellCursor Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.CellCursor",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> CellCursor:
        """CellCursor Component"""
        # pylint: disable=no-member
        return cast("CellCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
