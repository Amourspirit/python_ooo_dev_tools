from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from .text_properties_comp import TextPropertiesComp
from ooodev.adapter.style.character_properties_comp import CharacterPropertiesComp
from ooodev.adapter.style.paragraph_properties_comp import ParagraphPropertiesComp
from ooodev.adapter.style.character_properties_asian_comp import CharacterPropertiesAsianComp
from ooodev.adapter.style.character_properties_complex_comp import CharacterPropertiesComplexComp
from ooodev.exceptions import ex as mEx


if TYPE_CHECKING:
    from com.sun.star.drawing import Text  # service


class TextComp(ComponentBase):
    """
    Class for managing table Text Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Text) -> None:
        """
        Constructor

        Args:
            component (Text): UNO Text Component.
        """
        ComponentBase.__init__(self, component)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.Text",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Text:
        """Text Component"""
        return cast("Text", self._get_component())

    @property
    def character_properties(self) -> CharacterPropertiesComp:
        """CharacterProperties Component"""
        try:
            return self.__character_properties
        except AttributeError:
            self.__character_properties = CharacterPropertiesComp(self.component)
            return self.__character_properties

    @property
    def paragraph_properties(self) -> ParagraphPropertiesComp:
        """ParagraphProperties Component"""
        try:
            return self.__paragraph_properties
        except AttributeError:
            self.__paragraph_properties = ParagraphPropertiesComp(self.component)
            return self.__paragraph_properties

    @property
    def text_properties(self) -> TextPropertiesComp:
        """TextProperties Component"""
        try:
            return self.__text_properties
        except AttributeError:
            self.__text_properties = TextPropertiesComp(self.component)
            return self.__text_properties

    @property
    def character_properties_asian(self) -> CharacterPropertiesAsianComp | None:
        """Optional, CharacterPropertiesAsian Component"""
        try:
            return self.__character_properties_asian
        except AttributeError:
            try:
                self.__character_properties_asian = CharacterPropertiesAsianComp(self.component)
            except mEx.NotSupportedServiceError:
                self.__character_properties_asian = None
            return self.__character_properties_asian

    @property
    def character_properties_complex(self) -> CharacterPropertiesComplexComp | None:
        """Optional, CharacterPropertiesComplex Component"""
        try:
            return self.__character_properties_complex
        except AttributeError:
            try:
                self.__character_properties_complex = CharacterPropertiesComplexComp(self.component)
            except mEx.NotSupportedServiceError:
                self.__character_properties_complex = None
            return self.__character_properties_complex
        # endregion Properties
