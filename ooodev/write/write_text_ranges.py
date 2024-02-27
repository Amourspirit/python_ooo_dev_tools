from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.write.write_text_range import WriteTextRange

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess


class WriteTextRanges(LoInstPropsPartial, IndexAccessComp, WriteDocPropPartial, QiPartial):
    """
    Class for managing Writer Forms.

    This class is Enumerable and returns ``WriteTextCursor`` instance on iteration.

    .. versionadded:: 0.30.0
    """

    def __init__(self, owner: WriteDocPropPartial, component: XIndexAccess) -> None:
        """
        Constructor

        Args:
            owner (WriteDrawPage): Owner Component
            forms (XForms): Forms instance.
            lo_inst (LoInst, optional): Lo instance. Used when creating multiple documents. Defaults to ``None``.
        """
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        IndexAccessComp.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)

    def __next__(self) -> WriteTextRange:
        """
        Gets the next Text Range.

        Returns:
            WriteTextRange: Text Range instance.
        """
        return WriteTextRange(owner=self, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, key: int) -> WriteTextRange:
        """
        Gets the element at the specified index.

        Args:
            key (int): The Zero-based index of the element. Key can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            WriteTextRange: The element at the specified index.
        """
        return self.get_by_index(key)

    def __len__(self) -> int:
        """
        Gets the number of Text ranges in this instance.

        Returns:
            int: Number of Text ranges.
        """
        return self.component.getCount()

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> WriteTextRange:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            WriteTextRange: The element at the specified index.
        """
        idx = self._get_index(idx, True)
        result = super().get_by_index(idx)
        return WriteTextRange(owner=self, component=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides
