from __future__ import annotations
from typing import Any, TypeVar, Generic, TYPE_CHECKING
import uno
from com.sun.star.container import XEnumerationAccess

from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write import write_text_portion as mWriteTextPortion

if TYPE_CHECKING:
    from ooodev.proto.component_proto import ComponentT

T = TypeVar("T", bound="ComponentT")


class WriteTextPortions(
    LoInstPropsPartial, WriteDocPropPartial, EnumerationAccessPartial, QiPartial, TheDictionaryPartial, Generic[T]
):
    """
    Represents writer Text Portions.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: XEnumerationAccess, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.TextPortion`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        self._owner = owner
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        EnumerationAccessPartial.__init__(self, component=component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        TheDictionaryPartial.__init__(self)

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
        """
        Gets the next element.

        Returns:
            WriteTextPortion[T]: Next element.
        """
        result = super().__next__()
        return mWriteTextPortion.WriteTextPortion(self.owner, result)

    # endregion Overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
