from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.write.write_text_cursor import WriteTextCursor

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess


class WriteTextCursors(
    LoInstPropsPartial,
    IndexAccessComp,
    WriteDocPropPartial,
    QiPartial,
    TheDictionaryPartial,
):
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
        TheDictionaryPartial.__init__(self)

    def __next__(self) -> WriteTextCursor:
        return WriteTextCursor(owner=self, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, index: int) -> WriteTextCursor:
        return self.get_by_index(index)

    def __len__(self) -> int:
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

    def get_by_index(self, idx: int) -> WriteTextCursor:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            WriteTextCursor: The element at the specified index.
        """
        idx = self._get_index(idx, True)
        result = super().get_by_index(idx)
        return WriteTextCursor(owner=self, component=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides
