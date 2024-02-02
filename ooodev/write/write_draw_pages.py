"""DrawPages class for Draw documents."""

from __future__ import annotations
from typing import TYPE_CHECKING
import contextlib
import uno
from com.sun.star.drawing import XDrawPage

from ooodev.adapter.drawing.draw_pages_comp import DrawPagesComp
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from .write_draw_page import WriteDrawPage


if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPages
    from .write_doc import WriteDoc


class WriteDrawPages(LoInstPropsPartial, DrawPagesComp, QiPartial):
    """
    Class for managing Writer Draw Pages.
    """

    def __init__(self, owner: WriteDoc, slides: XDrawPages, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Owner Document
            slides (XDrawPages): Document Pages.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DrawPagesComp.__init__(self, slides)  # type: ignore
        # The API does not show that DrawPages implements XNameAccess, but it does.
        QiPartial.__init__(self, component=slides, lo_inst=self.lo_inst)
        self._current_index = 0

    def __getitem__(self, idx: int) -> WriteDrawPage[WriteDoc]:
        return self.get_by_index(idx=idx)

    def __len__(self) -> int:
        return self.component.getCount()

    def __iter__(self):
        self._current_index = 0
        return self

    def __next__(self) -> WriteDrawPage[WriteDoc]:
        if self._current_index >= len(self):
            self._current_index = 0
            raise StopIteration
        self._current_index += 1
        return self[self._current_index - 1]

    def __delitem__(self, _item: int | WriteDrawPage[WriteDoc] | XDrawPage) -> None:
        # Delete slide by index, name, or object
        if mInfo.Info.is_instance(_item, int):
            self.delete_page(_item)
        elif mInfo.Info.is_instance(_item, WriteDrawPage):
            super().remove(_item.component)
        elif mInfo.Info.is_instance(_item, XDrawPage):
            super().remove(_item)
        else:
            raise TypeError(f"Unsupported type: {type(_item)}")

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

    def insert_page(self, idx: int) -> WriteDrawPage[WriteDoc]:
        """
        Inserts a draw page at the given position in the document

        Args:
            idx (int): Index, can be a negative value to insert from the end of the document.
                For example, -1 will insert at the end of the document.

        Returns:
            WriteDrawPage: New slide that was inserted.
        """
        idx = self._get_index(idx=idx, allow_greater=True)
        return WriteDrawPage(owner=self.owner, component=self.component.insertNewByIndex(idx), lo_inst=self.lo_inst)

    def delete_page(self, idx: int) -> bool:
        """
        Deletes a draw page

        Args:
            idx (int): Index. Can be a negative value to delete from the end of the document.
                For example, -1 will delete the last slide.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        idx = self._get_index(idx=idx, allow_greater=False)
        with contextlib.suppress(Exception):
            # get the slide as UNO object and remove it
            result = super().get_by_index(idx)
            if result is None:
                return False
            self.remove(result)
            return True
        return False

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> WriteDrawPage[WriteDoc]:
        """
        Gets the element with the specified index.

        Args:
            idx (int): The index of the element. Idx can be a negative value to get from the end of the document.
                For example, -1 will get the last slide.

        Raises:
            IndexError: If unable to find slide with index.

        Returns:
            WriteDrawPage: The drawpage with the specified index.
        """
        idx = self._get_index(idx=idx, allow_greater=False)
        if idx >= len(self):
            raise IndexError(f"Index out of range: '{idx}'")

        result = super().get_by_index(idx)
        return WriteDrawPage(owner=self.owner, component=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides

    # region XDrawPages overrides
    def insert_new_by_index(self, idx: int) -> WriteDrawPage[WriteDoc]:
        """
        Creates and inserts a new GenericDrawPage or MasterPage into this container.

        Args:
            idx (int): The index at which the new page will be inserted.
                ``idx`` can be a negative value to insert from the end of the document.
                For example, ``-1`` will insert at the end of the document.

        Returns:
            WriteDrawPage: The new page.
        """
        idx = self._get_index(idx=idx, allow_greater=True)
        result = super().insert_new_by_index(idx)
        return WriteDrawPage(owner=self.owner, component=result, lo_inst=self.lo_inst)

    # endregion XDrawPages overrides

    # region Properties
    @property
    def owner(self) -> WriteDoc:
        """
        Returns:
            _T: Draw or Impress document.
        """
        return self._owner

    # endregion Properties
