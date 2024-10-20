from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.reflection.constant_type_description_partial import ConstantTypeDescriptionPartial

if TYPE_CHECKING:
    from com.sun.star.reflection import XConstantTypeDescription


class ConstantTypeDescriptionComp(ComponentProp, ConstantTypeDescriptionPartial):
    """
    Class for managing XConstantTypeDescription.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XConstantTypeDescription) -> None:
        """
        Constructor

        Args:
            component (XConstantTypeDescription): UNO Component that implements ``com.sun.star.reflection.XConstantTypeDescription`` interface.
        """
        ComponentProp.__init__(self, component)
        ConstantTypeDescriptionPartial.__init__(self, component=component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by TypeDescriptionEnumerationPartial
        return ()

    # endregion Overrides

    # region Properties
    @property
    @override
    def component(self) -> XConstantTypeDescription:
        """XConstantTypeDescription Component"""
        # pylint: disable=no-member
        return cast("XConstantTypeDescription", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
