from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial


if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


class CharacterPropertiesComp(ComponentBase, CharacterPropertiesPartial):
    """
    Class for managing table CharacterProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CharacterProperties) -> None:
        """
        Constructor

        Args:
            component (CharacterProperties): UNO CharacterProperties Component.
        """
        ComponentBase.__init__(self, component)
        CharacterPropertiesPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.TextProperties", "com.sun.star.style.CharacterProperties")

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> CharacterProperties:
        """CharacterProperties Component"""
        # pylint: disable=no-member
        return cast("CharacterProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
