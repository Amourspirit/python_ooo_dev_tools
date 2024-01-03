"""DrawPages class for Draw documents."""
from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import contextlib
import uno
from com.sun.star.drawing import XDrawPage

from ooodev.adapter.drawing.draw_pages_comp import DrawPagesComp
from .spreadsheet_draw_page import SpreadsheetDrawPage
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.partial.qi_partial import QiPartial

from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPages

_T = TypeVar("_T", bound="ComponentT")


class SpreadsheetDrawPages(Generic[_T], DrawPagesComp, QiPartial):
    """
    Class for managing Spreadsheet Draw Pages.
    """

    def __init__(self, owner: _T, slides: XDrawPages) -> None:
        """
        Constructor

        Args:
            owner (DrawDoc): Owner Document
            sheet (XDrawPages): Document Pages.
        """
        self.__owner = owner
        DrawPagesComp.__init__(self, slides)  # type: ignore
        # The API does not show that DrawPages implements XNameAccess, but it does.
        QiPartial.__init__(self, component=slides, lo_inst=mLo.Lo.current_lo)
        self._current_index = 0

    def __getitem__(self, _itm: int) -> SpreadsheetDrawPage[_T]:
        if _itm < 0:
            _itm = len(self) + _itm
            if _itm < 0:
                raise IndexError("list index out of range")
        return self.get_by_index(idx=_itm)

    def __len__(self) -> int:
        return self.component.getCount()

    def __iter__(self):
        return self

    def __next__(self) -> SpreadsheetDrawPage[_T]:
        if self._current_index >= len(self):
            self._current_index = 0
            raise StopIteration
        self._current_index += 1
        return self[self._current_index - 1]

    def __delitem__(self, _item: int | SpreadsheetDrawPage[_T] | XDrawPage) -> None:
        # Delete slide by index, name, or object
        if mInfo.Info.is_instance(_item, int):
            self.delete_page(_item)
        elif mInfo.Info.is_instance(_item, SpreadsheetDrawPage):
            super().remove(_item.component)
        elif mInfo.Info.is_instance(_item, XDrawPage):
            super().remove(_item)
        else:
            raise TypeError(f"Unsupported type: {type(_item)}")

    def insert_page(self, idx: int) -> SpreadsheetDrawPage[_T]:
        """
        Inserts a draw page at the given position in the document

        Args:
            idx (int): Index, can be a negative value to insert from the end of the document.
                For example, -1 will insert at the end of the document.

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            SpreadsheetDrawPage: New slide that was inserted.
        """
        if idx < 0:
            idx = len(self) + idx
            if idx < 0:
                raise IndexError("list index out of range")
        return SpreadsheetDrawPage(self.owner, self.component.insertNewByIndex(idx))

    def delete_page(self, idx: int) -> bool:
        """
        Deletes a draw page

        Args:
            idx (int): Index. Can be a negative value to delete from the end of the document.
                For example, -1 will delete the last slide.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        if idx < 0:
            idx = len(self) + idx
            if idx < 0:
                raise IndexError("list index out of range")
        with contextlib.suppress(Exception):
            # get the slide as UNO object and remove it
            result = super().get_by_index(idx)
            if result is None:
                return False
            self.remove(result)
            return True
        return False

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> SpreadsheetDrawPage[_T]:
        """
        Gets the element with the specified index.

        Args:
            idx (int): The index of the element. Idx can be a negative value to get from the end of the document.
                For example, -1 will get the last slide.

        Raises:
            IndexError: If unable to find slide with index.

        Returns:
            SpreadsheetDrawPage: The drawpage with the specified index.
        """
        if idx < 0:
            idx = len(self) + idx
            if idx < 0:
                raise IndexError("Index out of range")
        if idx >= len(self):
            raise IndexError(f"Index out of range: '{idx}'")

        result = super().get_by_index(idx)
        return SpreadsheetDrawPage(owner=self.owner, component=result)

    # endregion XIndexAccess overrides

    # region XDrawPages overrides
    def insert_new_by_index(self, idx: int) -> SpreadsheetDrawPage[_T]:
        """
        Creates and inserts a new GenericDrawPage or MasterPage into this container.

        Args:
            idx (int): The index at which the new page will be inserted.
                ``idx`` can be a negative value to insert from the end of the document.
                For example, ``-1`` will insert at the end of the document.

        Returns:
            SpreadsheetDrawPage: The new page.
        """
        if idx >= len(self):
            # if index is greater than the number of slides, then insert at the end
            idx = -1
        if idx < 0:
            idx = len(self) + idx
            if idx < 0:
                raise IndexError("Index out of range")
        result = super().insert_new_by_index(idx)
        return SpreadsheetDrawPage(owner=self.owner, component=result)

    # endregion XDrawPages overrides

    # region Properties
    @property
    def owner(self) -> _T:
        """
        Returns:
            _T: Draw or Impress document.
        """
        return self.__owner

    # endregion Properties
