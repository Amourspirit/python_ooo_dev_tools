from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.style import CharacterPropertiesComplex  # service


class CharacterPropertiesComplexComp(ComponentBase):
    """
    Class for managing table CharacterPropertiesComplex Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CharacterPropertiesComplex) -> None:
        """
        Constructor

        Args:
            component (CharacterPropertiesComplex): UNO CharacterPropertiesComplex Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.CharacterPropertiesComplex",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> CharacterPropertiesComplex:
        """CharacterPropertiesComplex Component"""
        # pylint: disable=no-member
        return cast("CharacterPropertiesComplex", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
