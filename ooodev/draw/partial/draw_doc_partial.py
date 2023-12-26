from __future__ import annotations
from typing import List, overload, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.frame import XModel

from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.adapter.drawing.draw_pages_comp import DrawPagesComp
from ooodev.office import draw as mDraw
from ooodev.utils import gui as mGUI
from ooodev.utils import lo as mLo
from ooodev.utils.kind.drawing_layer_kind import DrawingLayerKind
from ooodev.utils.kind.shape_comb_kind import ShapeCombKind
from ooodev.utils.kind.zoom_kind import ZoomKind
from ooodev.proto.component_proto import ComponentT

from .. import draw_page as mDrawPage
from .. import master_draw_page as mMasterDrawPage
from ..shapes import draw_shape as mDrawShape

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
    from com.sun.star.drawing import XDrawView
    from com.sun.star.drawing import XLayer
    from com.sun.star.drawing import XLayerManager
    from com.sun.star.drawing import XShapes
    from com.sun.star.frame import XController
    from com.sun.star.lang import XComponent
    from ooodev.utils.data_type.size import Size


_T = TypeVar("_T", bound="ComponentT")


class DrawDocPartial(Generic[_T]):
    def __init__(self, owner: _T, component: XComponent) -> None:
        self.__owner = owner
        self.__component = component

    def add_slide(self) -> mDrawPage.DrawPage[_T]:
        """
        Add a slide to the end of the document.

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPage: The slide that was inserted at the end of the document.
        """
        result = mDraw.Draw.add_slide(doc=self.__component)
        return mDrawPage.DrawPage(self.__owner, result)

    def add_layer(self, lm: XLayerManager, layer_name: str) -> XLayer:
        """
        Adds a layer

        Args:
            lm (XLayerManager): Layer Manager
            layer_name (str): Layer Name

        Raises:
            DrawError: If error occurs.

        Returns:
            XLayer: Newly added layer.
        """
        return mDraw.Draw.add_layer(lm, layer_name)

    def build_play_list(self, custom_name: str, *slide_idxs: int) -> NameContainerComp:
        """
        Build a named play list container of  slides from doc.
        The name of the play list is ``custom_name``.

        Args:
            custom_name (str): Name for play list
            slide_idxs (int): One or more index's of existing slides to add to play list.

        Raises:
            DrawError: If error occurs.

        Returns:
            NameContainerComp: Name Container.
        """
        return NameContainerComp(mDraw.Draw.build_play_list(self.__component, custom_name, *slide_idxs))

    def close_doc(self, deliver_ownership: bool = False) -> bool:
        """
        Closes text document.

        Args:
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException. Default ``False``.

        Returns:
            bool: False if DOC_CLOSING event is canceled, Other

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CLOSING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.DOC_CLOSED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc``.

        Attention:
            :py:meth:`Lo.close <.utils.lo.Lo.close>` method is called along with any of its events.
        """

        result = mLo.Lo.close(closeable=self.component, deliver_ownership=deliver_ownership)  # type: ignore
        return result

    def combine_shape(self, shapes: XShapes, combine_op: ShapeCombKind) -> mDrawShape.DrawShape[_T]:
        """
        Combines one or more shapes.

        Args:
            doc (XComponent): Document
            shapes (XShapes): Shapes to combine
            combine_op (ShapeCompKind): Combine Operation.

        Raises:
            ShapeError: If error occurs.

        Returns:
            DrawShape: New combined shape.
        """
        result = mDraw.Draw.combine_shape(self.__component, shapes, combine_op)
        return mDrawShape.DrawShape(self.__owner, result)

    def delete_slide(self, idx: int) -> bool:
        """
        Deletes a slide

        Args:
            idx (int): Index

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        return mDraw.Draw.delete_slide(self.__component, idx)

    def duplicate(self, idx: int) -> mDrawPage.DrawPage[_T]:
        """
        Duplicates a slide

        Args:
            idx (int): Index of slide to duplicate.

        Raises:
            DrawError If unable to create duplicate.

        Returns:
            DrawPage: Duplicated slide.
        """
        page = mDraw.Draw.duplicate(self.__component, idx)
        return mDrawPage.DrawPage(self.__owner, page)

    def find_master_page(self, style: str) -> mMasterDrawPage.MasterDrawPage[_T]:
        """
        Finds master page

        Args:
            style (str): Style of master page

        Raises:
            DrawPageMissingError: If unable to match ``style``.
            DrawPageError: if any other error occurs.

        Returns:
            DrawPage: Master page as Draw Page if found.
        """
        page = mDraw.Draw.find_master_page(self.__component, style)
        return mMasterDrawPage.MasterDrawPage(self.__owner, page)

    def find_slide_idx_by_name(self, name: str) -> int:
        """
        Gets a slides index by its name

        Args:
            name (str): Slide Name

        Returns:
            int: Zero based index if found; Otherwise ``-1``
        """
        return mDraw.Draw.find_slide_idx_by_name(self.__component, name)

    def get_controller(self) -> XController:
        """
        Gets controller from document.

        Returns:
            XController: Controller.
        """
        model = mLo.Lo.qi(XModel, self.__component, True)
        return model.getCurrentController()

    def get_handout_master_page(self) -> mDrawPage.DrawPage[_T]:
        """
        Gets handout master page

        Raises:
            DrawError: If unable to get hand-out master page.
            DrawPageMissingError: If Draw Page is ``None``.

        Returns:
            DrawPage: Draw Page
        """
        page = mDraw.Draw.get_handout_master_page(self.__component)
        return mDrawPage.DrawPage(self.__owner, page)

    def get_layer(self, layer_name: DrawingLayerKind | str) -> XLayer:
        """
        Gets layer from layer name

        Args:
            layer_name (str): Layer Name

        Raises:
            NameError: If ``layer_name`` does not exist.
            DrawError: If unable to get layer

        Returns:
            XLayer: Found Layer
        """
        return mDraw.Draw.get_layer(self.__component, layer_name)

    def get_layer_manager(self) -> XLayerManager:
        """
        Gets Layer manager for document.

        Args:
            doc (XComponent): Document

        Raises:
            DrawError: If error occurs.

        Returns:
            XLayerManager: Layer Manager
        """
        return mDraw.Draw.get_layer_manager(self.__component)

    def get_master_page(self, idx: int) -> mMasterDrawPage.MasterDrawPage[_T]:
        """
        Gets master page by index

        Args:
            idx (int): Index of master page

        Raises:
            DrawPageMissingError: If unable to get master page.
            DrawPageError: If any other error occurs.

        Returns:
            MasterDrawPage: Master page as Draw Page.
        """
        page = mDraw.Draw.get_master_page(doc=self.__component, idx=idx)
        return mMasterDrawPage.MasterDrawPage(self.__owner, page)

    def get_master_page_count(self) -> int:
        """
        Gets master page count

        Raises:
            DrawError: If error occurs.

        Returns:
            int: Master Page Count.
        """
        return mDraw.Draw.get_master_page_count(self.__component)

    def get_notes_page_by_index(self, idx: int) -> mDrawPage.DrawPage[_T]:
        """
        Gets notes page by index.

        Each draw page has a notes page.

        Args:
            idx (int): Index

        Raises:
            DrawPageError: If error occurs.

        Returns:
            DrawPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page`
        """
        page = mDraw.Draw.get_notes_page_by_index(self.__component, idx)
        return mDrawPage.DrawPage(self.__owner, page)

    def get_ordered_shapes(self) -> List[mDrawShape.DrawShape[_T]]:
        """
        Gets ordered shapes

        Returns:
            List[DrawShape[_T]]: List of Ordered Shapes.

        See Also:
            :py:meth:`~.draw.Draw.get_shapes`
        """
        shapes = mDraw.Draw.get_ordered_shapes(doc=self.__component)
        return [mDrawShape.DrawShape(self.__owner, shape) for shape in shapes]

    def get_play_list(self) -> NameContainerComp:
        """
        Gets Play list

        Raises:
            DrawError: If error occurs.

        Returns:
            NameContainerComp: Name Container
        """
        return NameContainerComp(mDraw.Draw.get_play_list(self.__component))

    def get_shapes(self) -> List[mDrawShape.DrawShape[_T]]:
        """
        Gets shapes

        Args:
            doc (XComponent): Document

        Raises:
            DrawError: If error occurs.

        Returns:
            List[DrawShape]: List of Shapes.
        """
        shapes = mDraw.Draw.get_shapes(doc=self.__component)
        return [mDrawShape.DrawShape(self.__owner, shape) for shape in shapes]

    def get_shapes_text(self) -> str:
        """
        Gets the text from inside all the document shapes

        Returns:
            str: Shapes text.

        See Also:
            - :py:meth:`~.draw.Draw.get_shapes`
            - :py:meth:`~.draw.Draw.get_ordered_shapes`
        """
        return mDraw.Draw.get_shapes_text(doc=self.__component)

    # region get_slide()
    @overload
    def get_slide(self) -> mDrawPage.DrawPage[_T]:
        """
        Gets draw page by page at index ``0``.

        Returns:
            DrawPage: Draw Page.
        """
        ...

    @overload
    def get_slide(self, *, idx: int) -> mDrawPage.DrawPage[_T]:
        """
        Gets draw page by page index

        Args:
            idx (int): Index of draw page. Default ``0``

        Returns:
            DrawPage: Draw Page.
        """
        ...

    @overload
    def get_slide(self, *, slides: XDrawPages) -> mDrawPage.DrawPage[_T]:
        """
        Gets draw page at index ``0`` from ``slides``.

        Args:
            slides (XDrawPages): Draw Pages

        Returns:
            DrawPage: Draw Page.
        """
        ...

    @overload
    def get_slide(self, *, slides: XDrawPages, idx: int) -> mDrawPage.DrawPage[_T]:
        """
        Gets draw page by page index from ``slides``.

        Args:
            slides (XDrawPages): Draw Pages
            idx (int): Index of slide. Default ``0``

        Returns:
            DrawPage: Slide as Draw Page.
        """
        ...

    def get_slide(self, **kwargs) -> mDrawPage.DrawPage[_T]:
        """
        Gets slide

        Args:
            slides (XDrawPages): Draw Pages
            idx (int): Index of slide. Default ``0``

        Raises:
            IndexError: If ``idx`` is out of bounds
            DrawError: If any other error occurs.

        Returns:
            DrawPage: Slide as Draw Page.
        """
        if not kwargs:
            result = mDraw.Draw.get_slide(doc=self.__component)
            return mDrawPage.DrawPage(self.__owner, result)
        if not "slides" in kwargs:
            kwargs["doc"] = self.__component
        result = mDraw.Draw.get_slide(**kwargs)
        return mDrawPage.DrawPage(self.__owner, result)

    # endregion get_slide()

    def get_slide_number(self, xdraw_view: XDrawView) -> int:
        """
        Gets slide number.

        Args:
            xdraw_view (XDrawView): Draw View.

        Raises:
            DrawError: If error occurs.

        Returns:
            int: Slide Number.
        """
        return mDraw.Draw.get_slide_number(xdraw_view=xdraw_view)

    def get_slide_size(self) -> Size:
        """
        Gets size of the given slide page (in mm units)

        Args:
            slide (XDrawPage): Slide

        Raises:
            SizeError: If unable to get size.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size struct.
        """
        return mDraw.Draw.get_slide_size(self.__component)  # type: ignore

    def get_slides(self) -> DrawPagesComp:
        """
        Gets the draw pages of a document.

        Args:
            doc (XComponent): Document.

        Raises:
            DrawPageMissingError: If there are no draw pages.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPagesComp: Draw Pages.
        """
        pages = mDraw.Draw.get_slides(self.__component)
        return DrawPagesComp(pages)

    def get_slides_count(self) -> int:
        """
        Gets the slides count.

        Returns:
            int: Number of slides.
        """
        return mDraw.Draw.get_slides_count(self.__component)

    def get_slides_list(self) -> List[mDrawPage.DrawPage[_T]]:
        """
        Gets all the slides as a list of XDrawPage

        Returns:
            List[DrawPage[_T]]: List of pages
        """
        slides = mDraw.Draw.get_slides_list(self.__component)
        return [mDrawPage.DrawPage(self.__owner, slide) for slide in slides]

    def get_viewed_page(self) -> mDrawPage.DrawPage[_T]:
        """
        Gets viewed page

        Raises:
            DrawPageError: If error occurs.

        Returns:
            DrawPage: Draw Page
        """
        page = mDraw.Draw.get_viewed_page(self.__component)
        return mDrawPage.DrawPage(self.__owner, page)

    def goto_page(self, page: XDrawPage) -> None:
        """
        Go to page.

        Args:
            page (XDrawPage): Page.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.goto_page(doc=self.__component, page=page)

    def insert_master_page(self, idx: int) -> mMasterDrawPage.MasterDrawPage[_T]:
        """
        Inserts a master page

        Args:
            idx (int): Index used to insert page

        Raises:
            DrawPageError: If unable to insert master page.

        Returns:
            MasterDrawPage: The newly inserted draw page.
        """
        page = mDraw.Draw.insert_master_page(doc=self.__component, idx=idx)
        return mMasterDrawPage.MasterDrawPage(self.__owner, page)

    def insert_slide(self, idx: int) -> mDrawPage.DrawPage[_T]:
        """
        Inserts a slide at the given position in the document

        Args:
            idx (int): Index

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPage: New slide that was inserted.
        """
        slide = mDraw.Draw.insert_slide(doc=self.__component, idx=idx)
        return mDrawPage.DrawPage(self.__owner, slide)

    def set_visible(self, visible: bool = True) -> None:
        """
        Set window visibility.

        Args:
            visible (bool, optional): If ``True`` window is set visible; Otherwise, window is set invisible. Default ``True``

        Returns:
            None:
        """
        mGUI.GUI.set_visible(doc=self.__component, visible=visible)

    def zoom(self, type: ZoomKind = ZoomKind.ENTIRE_PAGE) -> None:
        """
        Zooms document to a specific view.

        Args:
            type (ZoomKind, optional): Type of Zoom to set. Defaults to ``ZoomKind.ZOOM_100_PERCENT``.
        """

        def zoom_val(value: int) -> None:
            mGUI.GUI.zoom(view=ZoomKind.BY_VALUE, value=value)

        if type in (
            ZoomKind.ENTIRE_PAGE,
            ZoomKind.OPTIMAL,
            ZoomKind.PAGE_WIDTH,
            ZoomKind.PAGE_WIDTH_EXACT,
        ):
            mGUI.GUI.zoom(view=type)
        elif type == ZoomKind.ZOOM_200_PERCENT:
            zoom_val(200)
        elif type == ZoomKind.ZOOM_150_PERCENT:
            zoom_val(150)
        elif type == ZoomKind.ZOOM_100_PERCENT:
            zoom_val(100)
        elif type == ZoomKind.ZOOM_75_PERCENT:
            zoom_val(75)
        elif type == ZoomKind.ZOOM_50_PERCENT:
            zoom_val(50)

    def zoom_value(self, value: int = 100) -> None:
        """
        Sets the zoom level of the Document

        Args:
            value (int, optional): Value to set zoom. e.g. 160 set zoom to 160%. Default ``100``.
        """
        mGUI.GUI.zoom_value(value=value)
