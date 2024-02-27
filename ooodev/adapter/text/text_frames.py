from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.container.name_access_comp import NameAccessComp

if TYPE_CHECKING:
    from com.sun.star.text import TextFrames


class TextFramesComp(NameAccessComp, IndexAccessPartial):
    """
    Class for managing TextFrames Component.
    """

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.text.TextFrames`` service.
        """

        NameAccessComp.__init__(self, component)
        IndexAccessPartial.__init__(self, component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextFrames",)

    # endregion Overrides

    # region XEnumerationAccess

    # endregion XEnumerationAccess

    # region Properties
    @property
    def component(self) -> TextFrames:
        """TextFrames Component"""
        # pylint: disable=no-member
        return cast("TextFrames", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
