"""DrawPages class for Draw documents."""
from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.adapter.drawing.draw_pages_comp import DrawPagesComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils import lo as mLo
from ooodev.draw import draw_page as mDrawPage

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPages
    from ooodev.draw import DrawDoc


class DrawPages(DrawPagesComp, QiPartial):
    """
    Class for managing Draw Pages.
    """

    def __init__(self, owner: DrawDoc, slides: XDrawPages) -> None:
        """
        Constructor

        Args:
            owner (DrawDoc): Owner Document
            sheet (XDrawPages): Document Pages.
        """
        self.__owner = owner
        DrawPagesComp.__init__(self, slides)  # type: ignore
        QiPartial.__init__(self, component=slides, lo_inst=mLo.Lo.current_lo)
        self._current_index = 0

    def __getitem__(self, idx: int) -> mDrawPage.DrawPage[DrawDoc]:
        if idx < 0:
            idx = len(self) + idx
            if idx < 0:
                raise IndexError("list index out of range")
        return self.owner.get_slide(slides=self.component, idx=idx)

    def __len__(self) -> int:
        return self.component.getCount()

    def __iter__(self):
        return self

    def __next__(self) -> mDrawPage.DrawPage[DrawDoc]:
        if self._current_index >= len(self) - 1:
            self._current_index = 0
            raise StopIteration
        self._current_index += 1
        return self[self._current_index - 1]

    def insert_slide(self, idx: int) -> mDrawPage.DrawPage[DrawDoc]:
        """
        Inserts a slide at the given position in the document

        Args:
            idx (int): Index, can be a negative value to insert from the end of the document.
                For example, -1 will insert at the end of the document.

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPage: New slide that was inserted.
        """
        if idx < 0:
            idx = len(self) + idx
            if idx < 0:
                raise IndexError("list index out of range")
        return mDrawPage.DrawPage(self.owner, self.component.insertNewByIndex(idx))

    def delete_slide(self, idx: int) -> bool:
        """
        Deletes a slide

        Args:
            idx (int): Index. Can be a negative value to delete from the end of the document.
                For example, -1 will delete the last slide.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        return self.owner.delete_slide(idx=idx)

    # region Properties
    @property
    def owner(self) -> DrawDoc:
        """
        Returns:
            DrawDoc: Draw document.
        """
        return self.__owner

    # endregion Properties
