from __future__ import annotations
from typing import cast, TYPE_CHECKING, Generic, TypeVar

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.index_container_partial import IndexContainerPartial
from ooodev.utils import gen_util as mGenUtil

if TYPE_CHECKING:
    from com.sun.star.container import XIndexContainer

T = TypeVar("T")


class IndexContainerComp(ComponentProp, IndexContainerPartial[T], Generic[T]):
    """
    Class for managing XIndexContainer Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexContainer) -> None:
        """
        Constructor

        Args:
            component (XIndexContainer): UNO Component that implements ``com.sun.star.container.XIndexContainer``.
        """

        ComponentProp.__init__(self, component)
        IndexContainerPartial.__init__(self, component=self.component)

    def __getitem__(self, idx: int) -> T:
        """
        Gets the element at the specified index.

        Args:
            index (int): The Zero-based index of the element to get. Negative values are allowed and are interpreted as counting from the end of the list.

        Returns:
            T: The element at the specified index.

        Raises:
            IndexOutOfBoundsException: ``com.sun.star.lang.IndexOutOfBoundsException``
        """
        if idx < 0:
            count = self.component.getCount()
            index = mGenUtil.Util.get_index(idx, count, False)
        else:
            index = idx
        return self.component.getByIndex(index)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XIndexContainer:
        """XIndexContainer Component"""
        # overrides base property
        # pylint: disable=no-member
        return cast("XIndexContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
