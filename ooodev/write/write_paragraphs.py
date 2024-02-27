from __future__ import annotations
from typing import Any, TypeVar, Generic
import uno

from ooodev.adapter.text.text_comp import TextComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write import write_paragraph as mWriteParagraph

T = TypeVar("T", bound="ComponentT")


class WriteParagraphs(Generic[T], LoInstPropsPartial, WriteDocPropPartial, TextComp, QiPartial):
    """
    Represents writer paragraphs.

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
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
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

    def __next__(self) -> mWriteParagraph.WriteParagraph[T]:
        """
        Gets the next Paragraph.

        Returns:
            WriteParagraph[T]: The next paragraph.
        """
        result = super().__next__()
        return mWriteParagraph.WriteParagraph(owner=self.owner, component=result, lo_inst=self.lo_inst)

    # endregion Overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
