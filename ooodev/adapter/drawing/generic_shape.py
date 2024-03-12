from __future__ import annotations
from typing import TypeVar, Generic
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.shape_partial import ShapePartial
from ooodev.adapter.drawing.shape_descriptor_partial import ShapeDescriptorPartial
from ooodev.adapter.text.text_partial import TextPartial
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial


T = TypeVar("T")


class GenericShapeComp(
    Generic[T],
    ComponentBase,
    ShapePartial,
    ShapeDescriptorPartial,
    TextPartial,
    CharacterPropertiesPartial,
    ParagraphPropertiesPartial,
):
    """
    Class for managing Shape Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: T) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.ClosedBezierShape`` service.
        """
        ComponentBase.__init__(self, component)
        ShapePartial.__init__(self, component=component, interface=None)  # type: ignore
        ShapeDescriptorPartial.__init__(self, component=component, interface=None)  # type: ignore
        TextPartial.__init__(self, component, interface=None)  # type: ignore
        CharacterPropertiesPartial.__init__(self, component=component)  # type: ignore
        ParagraphPropertiesPartial.__init__(self, component=component)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> T:
        """ClosedBezierShape Component"""
        # pylint: disable=no-member
        return self._ComponentBase__get_component()  # type: ignore

    # endregion Properties
