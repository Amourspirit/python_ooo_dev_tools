from __future__ import annotations
from typing import Any, TypeVar, Generic
import uno

from ooodev.adapter.text.text_comp import TextComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from . import write_text_table as mWriteTextTable

T = TypeVar("T", bound="ComponentT")


class WriteTextTables(Generic[T], LoInstPropsPartial, TextComp, QiPartial):
    """
    Represents writer text tables.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TextComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore

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

    def __next__(self) -> mWriteTextTable.WriteTextTable[T]:
        result = super().__next__()
        return mWriteTextTable.WriteTextTable(owner=self.owner, component=result, lo_inst=self.lo_inst)

    # endregion Overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
