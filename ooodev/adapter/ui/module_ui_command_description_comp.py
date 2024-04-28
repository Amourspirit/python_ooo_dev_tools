from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import uno
from com.sun.star.beans import PropertyValue

from ooodev.adapter.container.name_access_comp import NameAccessComp

if TYPE_CHECKING:
    from com.sun.star.ui import ModuleUICommandDescription  # service


class ModuleUICommandDescriptionComp(NameAccessComp[Tuple[PropertyValue, ...]]):
    """
    Class for managing Action Trigger Container Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.ui.ModuleUICommandDescription`` service.
        """

        NameAccessComp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> ModuleUICommandDescription:
        """ModuleUICommandDescription Component"""
        # pylint: disable=no-member
        return cast("ModuleUICommandDescription", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
