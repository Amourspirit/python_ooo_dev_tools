from __future__ import annotations
from typing import Any, TypeVar, Generic
import uno
from com.sun.star.container import XEnumerationAccess

from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from . import write_text_portion as mWriteTextPortion

T = TypeVar("T", bound="ComponentT")


class WriteTextPortions(Generic[T], EnumerationAccessPartial, QiPartial):
    """
    Represents writer Text Portions.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: XEnumerationAccess) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.TextPortion`` service.
        """
        self.__owner = owner
        EnumerationAccessPartial.__init__(self, component=component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore

    # region Overrides
    def _is_next_element_valid(self, element: Any) -> bool:
        """
        Gets if the next element is valid.
        This method is called when iterating over the elements of this class.

        Args:
            element (Any): Element

        Returns:
            bool: True if element supports service com.sun.star.text.TextPortion.
        """
        return mInfo.Info.support_service(element, "com.sun.star.text.TextPortion")

    def __next__(self) -> mWriteTextPortion.WriteTextPortion[T]:
        result = super().__next__()
        return mWriteTextPortion.WriteTextPortion(self.owner, result)

    # endregion Overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
