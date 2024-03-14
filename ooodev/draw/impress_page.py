from __future__ import annotations
from typing import TYPE_CHECKING
from typing import TypeVar, Generic
import uno
from com.sun.star.drawing import XDrawPage

from ooodev.draw.draw_page import DrawPage
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.office import draw as mDraw

if TYPE_CHECKING:
    from ooodev.proto.component_proto import ComponentT

_T = TypeVar("_T", bound="ComponentT")

# ShapeFactoryPartial of DrawPage -> GenericDrawPage implements OfficeDocumentPropPartial


class ImpressPage(DrawPage[_T], Generic[_T]):
    """Represents an Impress page."""

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        DrawPage.__init__(self, owner=owner, component=component, lo_inst=lo_inst)

    def get_master_page(self) -> ImpressPage[_T]:
        """
        Gets master page

        Raises:
            DrawError: If error occurs.

        Returns:
            ImpressPage: Master Page.
        """
        page = mDraw.Draw.get_master_page(self.component)  # type: ignore
        return ImpressPage(owner=self._owner, component=page, lo_inst=self.lo_inst)

    def get_notes_page(self) -> ImpressPage[_T]:
        """
        Gets the notes page of a slide.

        Each draw page has a notes page.

        Raises:
            DrawPageMissingError: If notes page is ``None``.
            DrawPageError: If any other error occurs.

        Returns:
            ImpressPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page_by_index`
        """
        page = mDraw.Draw.get_notes_page(self.component)  # type: ignore
        return ImpressPage(owner=self._owner, component=page, lo_inst=self.lo_inst)

    def remove_master_page(self) -> None:
        """
        Removes this page as a master page.

        Raises:
            DrawError: If unable to remove master page.

        Returns:
            None:
        """
        if self._owner is None:
            raise mEx.DrawPageError("Owner is None")
        if not self.lo_inst.is_uno_interfaces(self._owner, XDrawPage):
            raise mEx.DrawPageError("Owner component is not XDrawPage")
        mDraw.Draw.remove_master_page(doc=self._owner, slide=self.__component)  # type: ignore

    def set_master_footer(self, text: str) -> None:
        """
        Sets master footer text.

        Args:
            text (str): Footer text.

        Raises:
            ShapeMissingError: If unable to find footer shape.
            DrawPageError: If any other error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_master_footer(master=self.__component, text=text)

    def title_only_slide(self, header: str) -> None:
        """
        Creates a slide with only a title.

        Args:
            slide (XDrawPage): Slide.
            header (str): Header text.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.title_only_slide(self.component, header)  # type: ignore

    def title_slide(self, title: str, sub_title: str = "") -> None:
        """
        Set a slides title and sub title.

        Args:
            title (str): Title.
            sub_title (str): Sub Title.

        Raises:
            DrawError: If error setting Slide.

        Returns:
            None:
        """
        mDraw.Draw.title_slide(self.component, title, sub_title)  # type: ignore

    def set_master_page(self, page: XDrawPage) -> None:
        """
        Sets The master page.

        Args:
            page (XDrawPage): Page to set as master.

        Raises:
            DrawError: If unable to remove master page.

        Returns:
            None:
        """
        mDraw.Draw.set_master_page(slide=self.component, page=page)  # type: ignore
