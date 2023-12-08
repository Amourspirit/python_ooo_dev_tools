from __future__ import annotations
from typing import Any, TypeVar, Generic
import uno

from ooodev.adapter.text.text_comp import TextComp
from ooodev.proto.component_proto import ComponentT

from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils import info as mInfo
from . import write_paragraph as mWriteParagraph

T = TypeVar("T", bound="ComponentT")


class WriteParagraphs(Generic[T], TextComp, QiPartial):
    """
    Represents writer paragraphs.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: Any) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
        """
        self.__owner = owner
        TextComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore

    # region Overrides
    def _is_next_element_valid(self, element: Any) -> bool:
        """
        Gets if the next element is valid.
        This method is called when iterating over the elements of this class.

        Args:
            element (Any): Element

        Returns:
            bool: True if element supports service com.sun.star.text.Paragraph.
        """
        return mInfo.Info.support_service(element, "com.sun.star.text.Paragraph")

    def __next__(self) -> mWriteParagraph.WriteParagraph[WriteParagraphs]:
        result = super().__next__()
        return mWriteParagraph.WriteParagraph(self, result)

    # endregion Overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
