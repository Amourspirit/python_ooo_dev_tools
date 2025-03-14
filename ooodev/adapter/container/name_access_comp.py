from __future__ import annotations
from typing import cast, TYPE_CHECKING, Generic, TypeVar

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from com.sun.star.container import XNameAccess

T = TypeVar("T")


class NameAccessComp(ComponentProp, NameAccessPartial[T], Generic[T]):
    """
    Class for managing XNameAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameAccess) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess``.
        """

        ComponentProp.__init__(self, component)
        NameAccessPartial.__init__(self, component=self.component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    @override
    def component(self) -> XNameAccess:
        """XNameAccess Component"""
        # overrides base property
        # pylint: disable=no-member
        return cast("XNameAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
