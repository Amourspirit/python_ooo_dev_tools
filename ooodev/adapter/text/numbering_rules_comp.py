from __future__ import annotations
from typing import cast, TYPE_CHECKING, Generic, TypeVar
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.index_replace_partial import IndexReplacePartial


if TYPE_CHECKING:
    from com.sun.star.text import NumberingRules  # service
    from com.sun.star.container import XIndexReplace

T = TypeVar("T")


class NumberingRulesComp(ComponentBase, IndexReplacePartial[T], Generic[T]):
    """
    Class for managing NumberingRules Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexReplace) -> None:
        """
        Constructor

        Args:
            component (XIndexReplace): UNO Component that supports ``com.sun.star.text.NumberingRules`` service.
        """

        ComponentBase.__init__(self, component)
        IndexReplacePartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.NumberingRules",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> NumberingRules:
        """Sheet Cell Cursor Component"""
        # pylint: disable=no-member
        return cast("NumberingRules", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
