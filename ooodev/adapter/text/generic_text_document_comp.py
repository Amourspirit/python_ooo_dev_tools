from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.frame.model_partial import ModelPartial


if TYPE_CHECKING:
    from com.sun.star.text import GenericTextDocument  # service


class GenericTextDocumentComp(ComponentBase, ModelPartial):
    """
    Class for managing GenericTextDocument Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GenericTextDocument) -> None:
        """
        Constructor

        Args:
            component (GenericTextDocument): UNO GenericTextDocument Component
        """

        ComponentBase.__init__(self, component)
        ModelPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.GenericTextDocument",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> GenericTextDocument:
        """Sheet Cell Cursor Component"""
        # pylint: disable=no-member
        return cast("GenericTextDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
