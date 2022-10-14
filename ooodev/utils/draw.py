# region Imports
from __future__ import annotations
from enum import IntEnum, Enum
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple, cast, overload
import math

from com.sun.star.awt import XButton
from com.sun.star.awt import XControlModel
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNamed
from com.sun.star.drawing import XControlShape
from com.sun.star.drawing import XDrawPage
from com.sun.star.drawing import XDrawPageDuplicator
from com.sun.star.drawing import XDrawPages
from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.drawing import XDrawView
from com.sun.star.drawing import XGluePointsSupplier
from com.sun.star.drawing import XGraphicExportFilter
from com.sun.star.drawing import XLayer
from com.sun.star.drawing import XLayerManager
from com.sun.star.drawing import XLayerSupplier
from com.sun.star.drawing import XMasterPagesSupplier
from com.sun.star.drawing import XMasterPageTarget
from com.sun.star.drawing import XShape
from com.sun.star.drawing import XShapes
from com.sun.star.frame import XComponentLoader
from com.sun.star.frame import XController
from com.sun.star.frame import XModel
from com.sun.star.lang import XComponent
from com.sun.star.presentation import XHandoutMasterSupplier
from com.sun.star.presentation import XPresentationPage
from com.sun.star.text import XText
from com.sun.star.text import XTextRange
from com.sun.star.view import XSelectionSupplier
from com.sun.star.container import XNameContainer
from com.sun.star.style import XStyle

from . import lo as mLo
from . import info as mInfo
from . import file_io as mFileIO
from . import gui as mGui
from . import props as mProps
from . import color as mColor
from . import images_lo as mImgLo
from ..cfg.config import Config  # singleton class.
from .type_var import PathOrStr
from . import table_helper as mTblHelper
from .kind.drawing_shape_kind import DrawingShapeKind
from .kind.form_control_kind import FormControlKind
from .kind.presentation_kind import PresentationKind
from .kind.graphic_style_kind import GraphicStyleKind
from .kind.drawing_gradient import DrawingGradientKind
from .kind.drawing_hatching_kind import DrawingHatchingKind
from .kind.drawing_bitmap_kind import DrawingBitmapKind
from .data_type.intensity import Intensity
from .data_type.image_offset import ImageOffset
from .data_type.angle import Angle

from ooo.dyn.awt.point import Point
from ooo.dyn.awt.size import Size
from ooo.dyn.drawing.connector_type import ConnectorType
from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.glue_point2 import GluePoint2
from ooo.dyn.drawing.line_dash import LineDash
from ooo.dyn.drawing.line_style import LineStyle
from ooo.dyn.drawing.poly_polygon_bezier_coords import PolyPolygonBezierCoords
from ooo.dyn.drawing.polygon_flags import PolygonFlags
from ooo.dyn.lang.illegal_argument_exception import IllegalArgumentException
from ooo.dyn.awt.gradient import Gradient
from ooo.dyn.awt.gradient_style import GradientStyle
from ooo.dyn.drawing.homogen_matrix3 import HomogenMatrix3

# endregion Imports


class Draw:
    # region Enums
    class GluePointsKind(IntEnum):
        TOP = 0
        RIGHT = 1
        BOTTOM = 2
        LEFT = 3

    class LayoutKind(IntEnum):
        TITLE_SUB = 0
        TITLE_BULLETS = 1
        TITLE_CHART = 2
        TITLE_2CONTENT = 3
        TITLE_CONTENT_CHART = 4
        TITLE_CONTENT_CLIP = 6
        TITLE_CHART_CONTENT = 7
        TITLE_TABLE = 8
        TITLE_CLIP_CONTENT = 9
        TITLE_CONTENT_OBJECT = 10
        TITLE_OBJECT = 11
        TITLE_CONTENT_2CONTENT = 12
        TITLE_OBJECT_CONTENT = 13
        TITLE_CONTENT_OVER_CONTENT = 14
        TITLE_2CONTENT_CONTENT = 15
        TITLE_2CONTENT_OVER_CONTENT = 16
        TITLE_CONTENT_OVER_OBJECT = 17
        TITLE_4OBJECT = 18
        TITLE_ONLY = 19
        BLANK = 20
        VTITLE_VTEXT_CHART = 27
        VTITLE_VTEXT = 28
        TITLE_VTEXT = 29
        TITLE_VTEXT_CLIP = 30
        CENTERED_TEXT = 32
        TITLE_4CONTENT = 33
        TITLE_6CONTENT = 34

    class NameSpaceKind(str, Enum):
        BULLETS_TEXT = "com.sun.star.presentation.OutlinerShape"
        SHAPE_TYPE_FOOTER = "com.sun.star.presentation.FooterShape"
        SHAPE_TYPE_NOTES = "com.sun.star.presentation.NotesShape"
        SHAPE_TYPE_PAGE = "com.sun.star.presentation.PageShape"
        SUBTITLE_TEXT = "com.sun.star.presentation.SubtitleShape"
        TITLE_TEXT = "com.sun.star.presentation.TitleTextShape"

        def __str__(self) -> str:
            return self.value

    class ShapeCompKind(IntEnum):
        MERGE = 0
        INTERSECT = 1
        SUBTRACT = 2
        COMBINE = 3

    class SlideShowKind(IntEnum):
        """DrawPage slide show change constants"""

        CLICK_ALL_CHANGE = 0
        """a mouse-click triggers the next animation effect or page change"""
        AUTO_CHANGE = 1
        """everything (page change, animation effects) is automatic"""
        CLICK_PAGE_CHANGE = 2
        """animation effects run automatically, but the user must click on the page to change it"""

    # endregion Enums

    # region Constants
    POLY_RADIUS = 20
    # endregion Constants

    # region open, create, save draw/impress doc
    @staticmethod
    def is_shapes_based(doc: XComponent) -> bool:
        """
        Gets if the document is suipports Draw or Impress

        Args:
            doc (XComponent): Document

        Returns:
            bool: ``True`` if supports Draw or Impress; Otherwise, ``False``.
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW) or mInfo.Info.is_doc_type(
            obj=doc, doc_type=mLo.Lo.Service.IMPRESS
        )

    @staticmethod
    def is_draw(doc: XComponent) -> bool:
        """
        Gets if the document is a Draw document

        Args:
            doc (XComponent): Document

        Returns:
            bool: ``True`` if is Draw document; Otherwise, ``False``.
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW)

    @staticmethod
    def is_impress(doc: XComponent) -> bool:
        """
        Gets if the document is an Impress document

        Args:
            doc (XComponent): Document

        Returns:
            bool: ``True`` if is Impress document; Otherwise, ``False``
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.IMPRESS)

    @staticmethod
    def create_draw_doc(loader: XComponentLoader) -> XComponent:
        """
        Creates a new Draw document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            XComponent: Component representing document
        """
        return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW, loader=loader)

    @staticmethod
    def create_impress_doc(loader: XComponentLoader) -> XComponent:
        """
        Creates a new Impress document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            XComponent: Component representing document
        """
        return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.IMPRESS, loader=loader)

    @staticmethod
    def get_slide_template_path() -> str:
        """
        Gets Slide template directory.

        Returns:
            str: Path as string
        """
        p = Path(mInfo.Info.get_office_dir(), Config().slide_template_path)
        return str(p)

    @staticmethod
    def save_page(page: XDrawPage, fnm: PathOrStr, mime_type: str) -> None:
        """
        Saves a Draw page to file.

        Args:
            page (XDrawPage): Page to save
            fnm (PathOrStr): Path to save page as
            mime_type (str): Mime Type of page to save as
        """
        save_file_url = mFileIO.FileIO.fnm_to_url(fnm)
        mLo.Lo.print(f'Saving page in "{fnm}"')

        # create graphics exporter
        gef = mLo.Lo.create_instance_mcf(
            XGraphicExportFilter, "com.sun.star.drawing.GraphicExportFilter", raise_err=True
        )

        # set the output 'document' to be specified page
        doc = mLo.Lo.qi(XComponent, page, True)
        # link exporter to the document
        gef.setSourceDocument(doc)

        # export the page by converting to the specified mime type
        props = mProps.Props.make_props(MediaType=mime_type, URL=save_file_url)

        gef.filter(props)
        mLo.Lo.print("Export Complete")

    # endregion open, create, save draw/impress doc

    # region methods related to document/multiple slides/pages
    @staticmethod
    def get_slides(doc: XComponent) -> XDrawPages:
        """
        Gets the draw pages of a document

        Args:
            doc (XComponent): Document

        Returns:
            XDrawPages | None: Draw Pages if found; Otherwise, ``None``.
        """
        supplier = mLo.Lo.qi(XDrawPagesSupplier, doc)
        if supplier is None:
            return None
        return supplier.getDrawPages()

    @classmethod
    def get_slides_count(cls, doc: XComponent) -> int:
        """
        Gets the slides count

        Args:
            doc (XComponent): Document

        Returns:
            int: _description_
        """
        slides = cls.get_slides(doc)
        if slides is None:
            return 0
        return slides.getCount()

    @classmethod
    def get_slides_list(cls, doc: XComponent) -> List[XDrawPage]:
        """
        Gets all the slides as a list of XDrawPage

        Args:
            doc (XComponent): Document

        Returns:
            List[XDrawPage]: List of pages
        """
        slides = cls.get_slides(doc)
        if slides is None:
            return []
        num_slides = slides.getCount()
        results: List[XDrawPage] = []
        for i in range(num_slides):
            results.append(mLo.Lo.qi(XDrawPage, slides.getByIndex(i)))
        return results

    # region get_slide()

    @classmethod
    def _get_slide_doc(cls, doc: XComponent, idx: int) -> XDrawPage | None:
        return cls._get_slide_slides(cls.get_slides(doc), idx)

    @staticmethod
    def _get_slide_slides(slides: XDrawPages, idx: int) -> XDrawPage | None:
        slide = None
        try:
            slide = mLo.Lo.qi(XDrawPage, slides.getByIndex(idx))
        except Exception:
            mLo.Lo.print(f"Could not get slide: {idx}")
        return slide

    @overload
    @classmethod
    def get_slide(cls, doc: XComponent, idx: int) -> XDrawPage | None:
        ...

    @overload
    @classmethod
    def get_slide(cls, slides: XDrawPages, idx: int) -> XDrawPage | None:
        ...

    @classmethod
    def get_slide(cls, *args, **kwargs) -> XDrawPage | None:
        """
        Gets slide by page index

        Args:
            doc (XComponent): Document
            slides (XDrawPages): Draw Pages
            idx (int): Index of slide

        Returns:
            XDrawPage | None: Slide if found; Otherwise ``None``.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "slides", "idx")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_slide() got an unexpected keyword argument")
            keys = ("doc", "slides")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            ka[2] = ka.get("idx", None)
            return ka

        if count != 2:
            raise TypeError("get_slide() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mLo.Lo.is_uno_interfaces(kargs[1], XDrawPage):
            return cls._get_slide_slides(kargs[1], kargs[2])
        return cls._get_slide_doc(kargs[1], kargs[2])

    # endregion get_slide()

    @classmethod
    def find_slide_idx_by_name(cls, doc: XComponent, name: str) -> int:
        """
        Gets a slides index by its name

        Args:
            doc (XComponent): Document
            name (str): Slide Name

        Returns:
            int: Zero based index if found; Otherwise ``-1``
        """
        slide_name = name.casefold()
        num_slides = cls.get_slides_count
        for i in range(num_slides):
            slide = cls.get_slide(doc, i)
            nm = str(mProps.Props.get_property(obj=slide, name="LinkDisplayName")).casefold()
            if nm == slide_name:
                return i
        return -1

    # region get_shapes()

    @classmethod
    def _get_shapes_doc(cls, doc: XComponent) -> List[XShape]:
        slides = cls.get_slides_list(doc)
        if not slides:
            return []
        shapes: List[XShape] = []
        for slide in slides:
            shapes.extend(cls._get_shapes_slide(slide))
        return shapes

    @classmethod
    def _get_shapes_slide(cls, slide: XDrawPage) -> List[XShape]:
        if slide is None:
            mLo.Lo.print("Slide is nulll")
            return []
        if slide.getCount() == 0:
            mLo.Lo.print("Slide does not contain any shapes")
            return []

        shapes: List[XShape] = []
        try:
            for i in range(slide.getCount()):
                shapes.append(mLo.Lo.qi(XShape, slide.getByIndex(i), True))
        except Exception as e:
            mLo.Lo.print("Shapes extraction error in slide")
            mLo.Lo.print(f"  {e}")

        return shapes

    @overload
    @classmethod
    def get_shapes(cls, doc: XComponent) -> List[XShape]:
        ...

    @overload
    @classmethod
    def get_shapes(cls, slide: XDrawPage) -> List[XShape]:
        ...

    @classmethod
    def get_shapes(cls, *args, **kwargs) -> List[XShape]:
        """
        Gets Shapes

        Args:
            doc (XComponent): Document
            slide (XDrawPage): Slide

        Returns:
            List[XShape]: List of shpaes
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "slide")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_shapes() got an unexpected keyword argument")
            keys = ("doc", "slide")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_shapes() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mLo.Lo.is_uno_interfaces(kargs[1], XDrawPage):
            return cls._get_shapes_slide(mLo.Lo.qi(XDrawPage, kargs[1], True))

        return cls._get_shapes_doc(mLo.Lo.qi(XComponent, kargs[1], True))

    # endregion get_shapes()

    # region get_ordered_shapes()
    @classmethod
    def _get_ordered_shapes_doc(cls, doc: XComponent) -> List[XShape]:
        # get all the shapes in all the pages of the doc, in z-order per slide
        slides = cls.get_slides_list(doc)
        if not slides:
            return []
        shapes: List[XShape] = []
        for slide in slides:
            shapes.extend(cls._get_ordered_shapes_slide(slide))
        return shapes

    @classmethod
    def _get_ordered_shapes_slide(cls, slide: XDrawPage) -> List[XShape]:
        def sorter(obj: XShape) -> int:
            return cls.get_zorder(obj)

        shapes = cls._get_shapes_slide(slide)
        sorted_shapes = sorted(shapes, key=sorter, reverse=False)
        return sorted_shapes

    @overload
    @classmethod
    def get_ordered_shapes(cls, doc: XComponent) -> List[XShape]:
        ...

    @overload
    @classmethod
    def get_ordered_shapes(cls, slide: XDrawPage) -> List[XShape]:
        ...

    @classmethod
    def get_ordered_shapes(cls, *args, **kwargs) -> List[XShape]:
        """
        Gets ordered shapes

        Args:
            doc (XComponent): Document
            slide (XDrawPage): Slide

        Returns:
            List[XShape]: List of Ordered Shapes.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "slide")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_ordered_shapes() got an unexpected keyword argument")
            keys = ("doc", "slide")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_ordered_shapes() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mLo.Lo.is_uno_interfaces(kargs[1], XDrawPage):
            return cls._get_ordered_shapes_slide(mLo.Lo.qi(XDrawPage, kargs[1], True))

        return cls._get_ordered_shapes_doc(mLo.Lo.qi(XComponent, kargs[1], True))

    # endregion get_ordered_shapes()

    @classmethod
    def get_shapes_text(cls, doc: XComponent) -> str:
        """
        Gets the text from inside all the document shapes

        Args:
            doc (XComponent): _description_

        Returns:
            str: _description_
        """
        sb: List[str] = []
        shapes = cls._get_ordered_shapes_doc(doc)
        for shape in shapes:
            text = cls.get_shape_text(shape)
            sb.append(text)
        return "\n".join(sb)

    @classmethod
    def add_slide(cls, doc: XComponent) -> XDrawPage:
        """
        Add a slede to the end of the document.

        Args:
            doc (XComponent): Document

        Returns:
            XDrawPage: The slide that was inserted at the end of the document.
        """
        mLo.Lo.print("Adding a slide")
        slides = cls.get_slides(doc)
        num_slides = slides.getCount()
        return slides.insertNewByIndex(num_slides)

    @classmethod
    def insert_slide(cls, doc: XComponent, idx: int) -> XDrawPage:
        """
        Inserts a slide at the given position in the document

        Args:
            doc (XComponent): Document
            idx (int): Index

        Returns:
            XDrawPage: New slide that was inserted.
        """
        mLo.Lo.print(f"Inserting a slide at postion: {idx}")
        slides = cls.get_slides(doc)
        return slides.insertNewByIndex(idx)

    @classmethod
    def delete_slide(cls, doc: XComponent, idx: int) -> bool:
        """
        Deletes a slide

        Args:
            doc (XComponent): Docuent
            idx (int): Index

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        mLo.Lo.print(f"Deleting a slide as postion: {idx}")
        slides = cls.get_slides(doc)
        slide = None
        try:
            slide = mLo.Lo.qi(XDrawPage, slides.getByIndex(idx), True)
        except Exception as e:
            mLo.Lo.print(f"could not find slide: {idx}")
            return False
        slides.remove(slide)
        return True

    @classmethod
    def duplicate(cls, doc: XComponent, idx: int) -> XDrawPage | None:
        """
        Duplicates a slide

        Args:
            doc (XComponent): Document
            idx (int): Index of slide to duplicate.

        Returns:
            XDrawPage | None: Duplicated slide if created; Otherwise, ``None``
        """
        dup = mLo.Lo.qi(XDrawPageDuplicator, doc, True)
        from_slide = cls._get_slide_doc(doc, idx)
        if from_slide is None:
            return None
        # places copy after original
        return dup.duplicate(from_slide)

    # endregion methods related to document/multiple slides/pages

    # region Layer Management
    @staticmethod
    def get_layer(doc: XComponent, layer_name: str) -> XLayer | None:
        """
        Gets layer from layer name

        Args:
            doc (XComponent): Document
            layer_name (str): Layer Name

        Returns:
            XLayer | None: Layer if found; Otherwise, ``None``
        """
        layer_supplier = mLo.Lo.qi(XLayerSupplier, doc, True)
        xname_access = layer_supplier.getLayerManager()
        try:
            return mLo.Lo.qi(XLayer, xname_access.getByName(layer_name), True)
        except Exception as e:
            mLo.Lo.print(f'Could not find the layer "{layer_name}"')
            mLo.Lo.print(f"  {e}")
        return None

    @staticmethod
    def add_layer(lm: XLayerManager, layer_name: str) -> XLayer | None:
        """
        Adds a layer

        Args:
            lm (XLayerManager): Layer Manager
            layer_name (str): Layer Name

        Returns:
            XLayer | None: Newly added layer on success; Otherwise, ``None``
        """
        layer = None
        try:
            layer = lm.insertNewByIndex(lm.getCount())
            props = mLo.Lo.qi(XPropertySet, layer, True)
            props.setPropertyValue("Name", layer_name)
            props.setPropertyValue("IsVisible", True)
            props.setPropertyValue("IsLocked", False)
        except Exception as e:
            mLo.Lo.print(f'Could not add the layer "{layer_name}"')
            mLo.Lo.print(f"  {e}")
        return layer

    # endregion Layer Management

    # region view page

    # region goto_page()
    @classmethod
    def _goto_page_doc(cls, doc: XComponent, page: XDrawPage) -> None:
        ctl = mGui.GUI.get_current_controller(doc)
        cls._goto_page_ctl(ctl, page)

    @staticmethod
    def _goto_page_ctl(ctl: XController, page: XDrawPage) -> None:
        xdraw_view = mLo.Lo.qi(XDrawView, ctl)
        xdraw_view.setCurrentPage(page)

    @overload
    @classmethod
    def goto_page(cls, doc: XComponent, page: XDrawPage) -> None:
        ...

    @overload
    @classmethod
    def goto_page(cls, ctl: XController, page: XDrawPage) -> None:
        ...

    @classmethod
    def goto_page(cls, *args, **kwargs) -> None:
        """
        Goto page

        Args:
            doc (XComponent): Document
            ctl (XController): Controller
            page (XDrawPage): Page
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "ctl", "page")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("goto_page() got an unexpected keyword argument")
            keys = ("doc", "ctl")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            ka[2] = ka.get("page", None)
            return ka

        if count != 2:
            raise TypeError("goto_page() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        obj = kwargs[1]
        ctl = mLo.Lo.qi(XController, obj)
        if ctl is None:
            cls._goto_page_doc(obj, kargs[2])
        else:
            cls._goto_page_ctl(ctl, kargs[2])

    # endregion goto_page()

    @staticmethod
    def get_viewd_page(doc: XComponent) -> XDrawPage:
        """
        Gets viewed page

        Args:
            doc (XComponent): Document

        Returns:
            XDrawPage: Draw Page
        """
        ctl = mGui.GUI.get_current_controller(doc)
        xdraw_view = mLo.Lo.qi(XDrawView, ctl, True)
        return xdraw_view.getCurrentPage()

    # region get_slide_number()

    @classmethod
    def _get_slide_number_draw_view(cls, xdraw_view: XDrawView) -> int:
        """
        Gets Drawview slide number

        Args:
            xdraw_view (XDrawView): Draw View

        Returns:
            int: Draw View page number.
        """
        curr_page = xdraw_view.getCurrentPage()
        return cls.get_slide_number(curr_page)

    @classmethod
    def _get_slide_number_draw_page(cls, slide: XDrawPage) -> int:
        return int(mProps.Props.get_property(obj=slide, name="Number"))

    @overload
    @classmethod
    def get_slide_number(cls, xdraw_view: XDrawView) -> int:
        ...

    @overload
    @classmethod
    def get_slide_number(cls, slide: XDrawPage) -> int:
        ...

    @classmethod
    def get_slide_number(cls, *args, **kwargs) -> int:
        """
        Gets slide number

        Args:
            xdraw_view (XDrawView): Draw View
            slide (XDrawPage): Slide

        Returns:
            int: Slide Number
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("xdraw_view", "slide")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_slide_number() got an unexpected keyword argument")
            keys = ("xdraw_view", "slide")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_slide_number() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        page = mLo.Lo.qi(XDrawPage, kargs[1])
        if page is None:
            return cls._get_slide_number_draw_view(kargs[1])
        else:
            return cls._get_slide_number_draw_page(kargs[1])

    # endregion get_slide_number()

    # endregion view page

    # region master page methods
    @staticmethod
    def get_master_page_count(doc: XComponent) -> int:
        """
        Gets master page count

        Args:
            doc (XComponent): Document

        Returns:
            int: Master Page Count.
        """
        mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc)
        pgs = mp_supp.getMasterPages()
        return pgs.getCount()

    # region get_master_page()
    @staticmethod
    def _get_master_page_idx(doc: XComponent, idx: int) -> XDrawPage | None:
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc)
            pgs = mp_supp.getMasterPages()
            return mLo.Lo.qi(XDrawPage, pgs.getByIndex(idx), True)
        except Exception:
            mLo.Lo.print(f"Could not find master slide for index: {idx}")
        return None

    @staticmethod
    def _get_master_page_slide(slide: XDrawPage) -> XDrawPage | None:
        mp_target = mLo.Lo.qi(XMasterPageTarget, slide)
        return mp_target.getMasterPage()

    @overload
    @classmethod
    def get_master_page(cls, doc: XComponent, idx: int) -> XDrawPage | None:
        ...

    @overload
    @classmethod
    def get_master_page(cls, slide: XDrawPage) -> XDrawPage | None:
        ...

    @classmethod
    def get_master_page(cls, *args, **kwargs) -> XDrawPage | None:
        """
        Gets master page

        Args:
            doc (XComponent): Document
            idx (int): Index of page
            slide (XDrawPage): Slide to get master page from.

        Returns:
            XDrawPage | None: Draw page if found; Otherwise ``None``
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "idx", "slide")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_master_page() got an unexpected keyword argument")
            keys = ("doc", "slide")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = ka.get("idx", None)
            return ka

        if count not in (1, 2):
            raise TypeError("get_master_page() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return cls._get_master_page_idx(kargs[1], kargs[2])

        return cls._get_master_page_slide(kargs[1])

    # endregion get_master_page()

    @staticmethod
    def insert_master_page(doc: XComponent, idx: int) -> XDrawPage:
        """
        Inserts a master page

        Args:
            doc (XComponent): Document
            idx (int): Index used to insert page

        Returns:
            XDrawPage: The newly inserted draw page.
        """
        mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
        pgs = mp_supp.getMasterPages()
        pgs.insertNewByIndex(idx)

    @staticmethod
    def remove_master_page(doc: XComponent, slide: XDrawPage) -> None:
        """
        Removes a master page

        Args:
            doc (XComponent): Document
            slide (XDrawPage): Draw page to remove
        """
        mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
        pgs = mp_supp.getMasterPages()
        pgs.remove(slide)

    @staticmethod
    def set_master_page(slide: XDrawPage, page: XDrawPage) -> None:
        """
        Sets master page

        Args:
            slide (XDrawPage): Slide
            page (XDrawPage): Page to set as master
        """
        mp_target = mLo.Lo.qi(XMasterPageTarget, slide, True)
        mp_target.setMasterPage(page)

    @staticmethod
    def get_handout_master_page(doc: XComponent) -> XDrawPage:
        """
        Gets handout master page

        Args:
            doc (XComponent): Document

        Returns:
            XDrawPage: Draw Page
        """
        hm_supp = mLo.Lo.qi(XHandoutMasterSupplier, doc, True)
        return hm_supp.getHandoutMasterPage()

    @staticmethod
    def find_master_page(doc: XComponent, style: str) -> XDrawPage | None:
        """
        Finds master page

        Args:
            doc (XComponent): Document
            style (str): Style of master page

        Returns:
            XDrawPage | None: Master page as Draw Page if found; Otherwise, ``None``.
        """
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
            master_pgs = mp_supp.getMasterPages()

            for i in range(master_pgs.getCount()):
                pg = mLo.Lo.qi(XDrawPage, master_pgs.getByIndex(i), True)
                nm = str(mProps.Props.get_property(obj=pg, name="LinkDisplayName"))
                if style == nm:
                    return pg
            mLo.Lo.print(f'Could not find master slide with style of: "{style}"')
            return None
        except Exception as e:
            mLo.Lo.print("Could not find master slide")
            mLo.Lo.print(f"  {e}")

        return None

    # endregion master page methods

    # region slide/page methods
    @classmethod
    def show_shapes_info(cls, slide: XDrawPage) -> None:
        """
        Prints info for shapes to console

        Args:
            slide (XDrawPage): Slide
        """
        print("Draw Page shapes:")
        shapes = cls._get_shapes_slide(slide)
        for shape in shapes:
            cls.show_shape_info(shape)

    @classmethod
    def get_slide_title(cls, slide: XDrawPage) -> str | None:
        shape = cls.find_shape_by_type(slide=slide, shape_type=str(Draw.NameSpaceKind.TITLE_TEXT))
        if shape is None:
            return None
        return cls._get_shape_text_shape(shape)

    @staticmethod
    def get_slide_size(slide: XDrawPage) -> Size | None:
        """
        Gets slide size

        Args:
            slide (XDrawPage): Slide

        Returns:
            Size | None: Size object on success; Otherwise, ``None``.
        """
        try:
            props = mLo.Lo.qi(XPropertySet, slide)
            if props is None:
                mLo.Lo.print("No slide properties found")
                return None
            width = int(props.getPropertyValue("Width"))
            height = int(props.getPropertyValue("Height"))
            return Size(width, height)
        except Exception as e:
            mLo.Lo.print("Could not get page dimensions")
            mLo.Lo.print(f"  {e}")
        return None

    @staticmethod
    def set_name(slide: XDrawPage, name: str) -> None:
        """
        Sets the name of a slide

        Args:
            slide (XDrawPage): Slide
            name (str): Name
        """
        xpage_name = mLo.Lo.qi(XNamed, slide, True)
        xpage_name.setName(name)

    @classmethod
    def title_slide(cls, slide: XDrawPage, title: str, sub_title: str) -> None:
        """
        Set a slides title and sub title

        Args:
            slide (XDrawPage): Slide
            title (str): Title
            sub_title (str): Sub Title
        """
        # Add text to the slide page by treating it as a title page, which
        # has two text shapes: one for the title, the other for a subtitle
        mProps.Props.set_property(obj=slide, name="Layout", value=int(Draw.LayoutKind.TITLE_SUB))

        # add the title text to the title shape
        xs = cls.find_shape_by_type(slide=slide, shape_type=Draw.NameSpaceKind.TITLE_TEXT)
        txt_field = mLo.Lo.qi(XText, xs, True)
        txt_field.setString(title)

        # add the subtitle text to the subtitle shape
        xs = cls.find_shape_by_type(slide=slide, shape_type=Draw.NameSpaceKind.SUBTITLE_TEXT)
        txt_field = mLo.Lo.qi(XText, xs, True)
        txt_field.setString(sub_title)

    @classmethod
    def bullets_slide(cls, slide: XDrawPage, title: str) -> XText:
        """
        Add text to the slide page by treating it as a bullet page, which
        has two text shapes: one for the title, the other for a sequence of
        bullet points; add the title text but return a reference to the bullet
        text area

        Args:
            slide (XDrawPage): Slide
            title (str): Title

        Returns:
            XText: Text Object
        """
        # Add text to the slide page by treating it as a bullet page, which
        # has two text shapes: one for the title, the other for a sequence of
        # bullet points; add the title text but return a reference to the bullet
        # text area

        mProps.Props.set_property(obj=slide, name="Layout", value=int(Draw.LayoutKind.TITLE_BULLETS))

        # add the title text to the title shape
        xs = cls.find_shape_by_type(slide=slide, shape_type=Draw.NameSpaceKind.TITLE_TEXT)
        txt_field = mLo.Lo.qi(XText, xs, True)
        txt_field.setString(title)

        # return a reference to the bullet text area
        xs = cls.find_shape_by_type(slide=slide, shape_type=Draw.NameSpaceKind.BULLETS_TEXT)
        return mLo.Lo.qi(XText, xs, True)

    @staticmethod
    def add_bullet(bulls_txt: XText, level: int, text: str) -> None:
        """
        Add bullet text to the end of the bullets text area, specifying
        the nesting of the bullet using a numbering level value
        (numbering starts at 0).

        Args:
            bulls_txt (XText): Text object
            level (int): Bullet Level
            text (str): Bullet Text
        """
        # access the end of the bullets text
        bulls_txt_end = mLo.Lo.qi(XTextRange, bulls_txt, True).getEnd()

        # set the bullet's level
        mProps.Props.set_property(obj=bulls_txt_end, name="NumberingLevel", value=level)

        # add the text
        bulls_txt_end.setString(f"{text}\n")

    @classmethod
    def title_only_slide(cls, slide: XDrawPage, header: str) -> None:
        """
        Creates a slide with only a title

        Args:
            slide (XDrawPage): Slide
            header (str): Header text.
        """
        mProps.Props.set_property(obj=slide, name="Layout", value=int(Draw.LayoutKind.TITLE_ONLY))

        # add the text to the title shape
        xs = cls.find_shape_by_type(slide=slide, shape_type=Draw.NameSpaceKind.TITLE_TEXT)
        txt_field = mLo.Lo.qi(XText, xs, True)
        txt_field.setString(header)

    @staticmethod
    def blank_slide(slide: XDrawPage) -> None:
        """
        Inserts a blank slide

        Args:
            slide (XDrawPage): Slide
        """
        mProps.Props.set_property(obj=slide, name="Layout", value=str(Draw.LayoutKind.BLANK))

    @staticmethod
    def get_notes_page(slide: XDrawPage) -> XDrawPage | None:
        """
        Gets the notes page of a slide.

        Each draw page has a notes page.

        Args:
            slide (XDrawPage): Slide

        Returns:
            XDrawPage | None: Notes Page if found; Otherwise, ``None``.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page_by_index`
        """
        pres_page = mLo.Lo.qi(XPresentationPage, slide)
        if pres_page is None:
            mLo.Lo.print("This is not a presentation slide, so no notes page is available")
            return None
        return pres_page.getNotesPage()

    @classmethod
    def get_notes_page_by_index(cls, doc: XComponent, idx: int) -> XDrawPage | None:
        """
        Gets notes page by index.

        Each draw page has a notes page.

        Args:
            doc (XComponent): Document
            idx (int): Index

        Returns:
            XDrawPage | None: Notes Page if found; Otherwise, ``None``.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page`
        """
        slide = cls._get_slide_doc(doc, idx)
        return cls.get_notes_page(slide)

    # endregion slide/page methods

    # region shape methods
    @classmethod
    def show_shape_info(cls, shape: XShape) -> None:
        """
        Prints shape info to console.

        Args:
            shape (XShape): Shape
        """
        print(f"  Shape service: {shape.getShapeType()}; z-order: {cls.get_zorder(shape)}")

    # region get_shape_text()
    @classmethod
    def _get_shape_text_shape(cls, shape: XShape) -> str:
        xtext = mLo.Lo.qi(XText, shape, True)

        xtext_cursor = xtext.createTextCursor()
        xtext_rng = mLo.Lo.qi(XTextRange, xtext_cursor, True)
        text = xtext_rng.getString()
        return text

    @classmethod
    def _get_shape_text_slide(cls, slide: XDrawPage) -> str:
        sb = List[str] = []
        shapes = cls._get_ordered_shapes_slide(slide)
        for shape in shapes:
            text = cls._get_shape_text_shape(shape)
            sb.append(text)
        return "\n".join(sb)

    @overload
    @classmethod
    def get_shape_text(cls, shape: XShape) -> str:
        ...

    @overload
    @classmethod
    def get_shape_text(cls, slide: XDrawPage) -> str:
        ...

    @classmethod
    def get_shape_text(cls, *args, **kwargs) -> str:
        """
        Gets the text from inside a shape

        Args:
            shape (XShape): Shape
            slide (XDrawPage): Slide

        Returns:
            str: Shape text
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "slide")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_shape_text() got an unexpected keyword argument")
            keys = ("shape", "slide")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_shape_text() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        shape = mLo.Lo.qi(XShape, kargs[1])
        if shape is None:
            return cls._get_shape_text_slide(kargs[1])
        else:
            return cls._get_shape_text_shape(shape)

    # endregion get_shape_text()

    @classmethod
    def find_shape_by_type(cls, slide: XDrawPage, shape_type: Draw.NameSpaceKind | str) -> XShape | None:
        """
        Finds a shape by its type

        Args:
            slide (XDrawPage): Slide
            shape_type (str): Shape Type

        Returns:
            XShape | None: Shape if Found; Otherwise, ``None``
        """
        shapes = cls.get_shapes(slide)
        if not shapes:
            mLo.Lo.print("No shapes were found in the draw page")
            return None

        st = str(shape_type)

        for shape in shapes:
            if st == shape.getShapeType():
                return shape
        mLo.Lo.print(f'No shape found for "{st}"')
        return None

    @classmethod
    def find_shape_by_name(cls, slide: XDrawPage, shape_name: str) -> XShape | None:
        """
        Finds a shape by its name.

        Args:
            slide (XDrawPage): Slide
            shape_name (str): Shape Name

        Returns:
            XShape | None: Shape if Found; Otherwise, ``None``
        """
        shapes = cls.get_shapes(slide)
        sn = shape_name.casefold()
        if not shapes:
            mLo.Lo.print("No shapes were found in the draw page")
            return None

        for shape in shapes:
            nm = str(mProps.Props.get_property(obj=shape, name="Name")).casefold()
            if nm == sn:
                return shape

        mLo.Lo.print(f'No shape named "{shape_name}"')
        return None

    @classmethod
    def copy_shape_contents(cls, slide: XDrawPage, old_shape: XShape) -> XShape:
        """
        Copies a shapes contents from old shape into new shape.

        Args:
            slide (XDrawPage): Slide
            old_shape (XShape): Old shape

        Returns:
            XShape: New shape with contents of old shpae copied.
        """
        shape = cls.copy_shape(slide, old_shape)
        mLo.Lo.print(f'Shape type: "{old_shape.getShapeType()}"')
        cls.add_text(shape, cls.get_shape_text(old_shape))

        return shape

    @classmethod
    def copy_shape(slide: XDrawPage, old_shape: XShape) -> XShape | None:
        """
        Copies a shape

        Args:
            slide (XDrawPage): Slide
            old_shape (XShape): Old Shape

        Returns:
            XShape: Newly Copied shape if can be copied; Otherwise, ``None``
        """
        # parameters are in 1/100 mm units
        pt = old_shape.getPosition()
        sz = old_shape.getSize()

        shape = None
        try:
            shape = mLo.Lo.create_instance_msf(XShape, old_shape.getShapeType(), raise_err=True)
            mLo.Lo.print(f"Copying: {old_shape.getShapeType()}")
            shape.setPosition(pt)
            shape.setSize(sz)
            slide.add(shape)
        except Exception:
            mLo.Lo.print("Unable to copy shape")
            pass
        return shape

    @staticmethod
    def set_zorder(shape: XShape, order: int) -> None:
        """
        Sets the z-order of a shape

        Args:
            shape (XShape): Shape
            order (int): Z-Order
        """
        mProps.Props.set_property(obj=shape, name="ZOrder", value=order)

    @staticmethod
    def get_zorder(shape: XShape) -> int:
        """
        Gets the z-order of a shape

        Args:
            shape (XShape): Shape

        Returns:
            int: Z-Order
        """
        return int(mProps.Props.get_property(obj=shape, name="ZOrder"))

    @classmethod
    def move_to_top(cls, slide: XDrawPage, shape: XShape) -> None:
        """
        Moves the z-order of a shape to the top.

        Args:
            slide (XDrawPage): Slide
            shape (XShape): Shape
        """
        max_zo = cls.find_biggest_zorder(slide)
        cls.set_zorder(shape, max_zo + 1)

    @classmethod
    def find_biggest_zorder(cls, slide: XDrawPage) -> int:
        """
        Findes the shape with the largest z-order.

        Args:
            slide (XDrawPage): Slide

        Returns:
            int: Z-Order
        """
        return cls.get_zorder(cls.find_top_shape(slide))

    @classmethod
    def find_top(cls, slide: XDrawPage) -> XShape | None:
        """
        Gets the topmost shape of a slide.

        Args:
            slide (XDrawPage): Slide

        Returns:
            XShape | None: Topmost shape if found; Otherwise; ``None``
        """
        shapes = cls.get_shapes(slide)
        if not shapes:
            mLo.Lo.print("No shapes found")
            return None
        max_zorder = 0
        top = None
        for shape in shapes:
            zo = cls.get_zorder(shape)
            if zo > max_zorder:
                max_zorder = zo
                top = shape
        return top

    @classmethod
    def move_to_bottom(cls, slide: XDrawPage, shape: XShape) -> None:
        """
        Moves a shape to the bottom of the z-order

        Args:
            slide (XDrawPage): Slide
            shape (XShape): Shape
        """
        shapes = cls.get_shapes(slide)
        if not shapes:
            mLo.Lo.print("No shapes found")
            return None

        min_zorder = 999
        for sh in shapes:
            zo = cls.get_zorder(sh)
            if zo < min_zorder:
                min_zorder = zo
            cls.set_zorder(sh, zo + 1)
        cls.set_zorder(shape, min_zorder)

    # endregion shape methods

    # region draw/add shape to a page
    @staticmethod
    def make_shape(shape_type: DrawingShapeKind | str, x: int, y: int, width: int, height: int) -> XShape | None:
        """
        Creates a shape

        Args:
            shape_type (DrawingShapeKind | str): Shape type.
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``

        See Also:
            :py:meth:`~.draw.Draw.add_shape`
        """
        # parameters are in mm units
        shape = None
        try:
            shape = mLo.Lo.create_instance_msf(XShape, f"com.sun.star.drawing.{shape_type}", raise_err=True)
            shape.setPosition(Point(x * 100, y * 100))
            shape.setSize(Size(width * 100, height * 100))
        except Exception as e:
            mLo.Lo.print(f'Unable to create shape "{shape_type}"')
            mLo.Lo.print(f"  {e}")
            return None
        return shape

    @classmethod
    def warns_position(cls, slide: XDrawPage, x: int, y: int) -> None:
        """
        Warns via console if a ``x`` or ``y`` is not on the page.

        Args:
            slide (XDrawPage): Slide
            x (int): X Position
            y (int): Y Position

        Note:
            This method uses :py:meth:`.LO.print`. and if those ``print()`` commands are
            suppressed then this method will not be effective.
        """
        slide_size = cls.get_slide_size(slide)
        if slide_size is None:
            mLo.Lo.print("No slide size found")
            return
        slide_width = slide_size.Width
        slide_height = slide_size.Height

        if x < 0:
            mLo.Lo.print("x < 0")
        elif x > slide_width - 1:
            mLo.Lo.print("x position off right hand side of the slide")

        if y < 0:
            mLo.Lo.print("y < 0")
        elif y > slide_height - 1:
            mLo.Lo.print("y position off bottom of the slide")

    @classmethod
    def add_shape(
        cls, slide: XDrawPage, shape_type: DrawingShapeKind | str, x: int, y: int, width: int, height: int
    ) -> XShape | None:
        """
        Adds a shape to a slide.

        Args:
            slide (XDrawPage): Slide
            shape_type (DrawingShapeKind | str): Shape type.
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``

        See Also:
            - :py:meth:`~.draw.Draw.warns_position`
            - :py:meth:`~.draw.Draw.make_shape`
        """
        cls.warns_position(slide=slide, x=x, y=y)
        shape = cls.make_shape(shape_type=shape_type, x=x, y=y, width=width, height=height)
        if shape is None:
            return None
        slide.add(shape)
        return shape

    @classmethod
    def draw_rectangle(cls, slide: XDrawPage, x: int, y: int, width: int, height: int) -> XShape | None:
        """
        Gets a rectangle

        Args:
            slide (XDrawPage): Slide
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        return cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.RECTANGLE_SHAPE, x=x, y=y, width=width, height=height
        )

    @classmethod
    def draw_circle(cls, slide: XDrawPage, x: int, y: int, radius: int) -> XShape | None:
        """
        Gets a rectangle

        Args:
            slide (XDrawPage): Slide
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            radius (int): Shape radius in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        return cls.add_shape(
            slide=slide,
            shape_type=DrawingShapeKind.ELLIPSE_SHAPE,
            x=x - radius,
            y=y - radius,
            width=radius * 2,
            height=radius * 2,
        )

    @classmethod
    def draw_ellipse(cls, slide: XDrawPage, x: int, y: int, width: int, height: int) -> XShape | None:
        """
        Gets a rectangle

        Args:
            slide (XDrawPage): Slide
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        return cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.ELLIPSE_SHAPE, x=x, y=y, width=width, height=height
        )

    @classmethod
    def draw_polygon(cls, slide: XDrawPage, x: int, y: int, sides: int, radius: int | None = None) -> XShape | None:
        """
        Gets a rectangle

        Args:
            slide (XDrawPage): Slide
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            radius (int): Shape radius in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        if radius is None:
            radius = Draw.POLY_RADIUS
        polygon = cls.add_shape(
            slide=slide,
            shape_type=DrawingShapeKind.POLY_POLYGON_SHAPE,
            x=0,
            y=0,
            width=0,
            height=0,
        )
        pts = cls.gen_polygon_points(x=x, y=y, radius=radius, sides=sides)
        # could be many polygons pts in this 2D array
        polys = [pts]
        mProps.Props.set_property(polygon, name="PolyPolygon", value=polys)

    @staticmethod
    def gen_polygon_points(x: int, y: int, radius: int, sides: int) -> Tuple[Point, ...]:
        """
        Generates a list of polygon points

        Args:
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            radius (int): Shape radius in mm units.
            sides (int): Number of Polygon sides from 3 to 30.

        Returns:
            List[Point]: List of points.
        """
        if sides < 3:
            mLo.Lo.print("Too few sides; must be 3 or more. Setting to 3")
            sides = 3
        elif sides > 30:
            mLo.Lo.print("Too many sides; must be 30 or less. Setting to 30")
            sides = 30

        pts: List[Point] = []
        angle_step = math.pi / sides
        for i in range(sides):
            pt = Point(
                int(round(((x * 100) + (radius * 100)) * math.cos(i * 2 * angle_step))),
                int(round(((y * 100) + (radius * 100)) * math.sin(i * 2 * angle_step))),
            )
            pts.append(pt)
        return tuple(pts)

    @classmethod
    def draw_bezier(
        cls, slide: XDrawPage, pts: List[Point], flags: List[PolygonFlags], is_open: bool
    ) -> XShape | None:
        """
        Draws a bezier curve.

        Args:
            slide (XDrawPage): Slide
            pts (List[Point]): Points
            flags (List[PolygonFlags]): Falgs
            is_open (bool): Determines if an open or closed bezier is drawn.

        Raises:
            IndexError: If ``pts`` and ``flags`` do not have the same number of elements.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        if len(pts) != len(flags):
            raise IndexError("pts and flags must be the same length")

        bezier_type = "OpenBezierShape" if is_open else "ClosedBezierShape"
        bezier_poly = cls.add_shape(slide=slide, shape_type=bezier_type, x=0, y=0, width=0, height=0)
        if bezier_poly is None:
            mLo.Lo.print("Failded to create bezier")
            return None
        # create space for one bezier shape
        coords = PolyPolygonBezierCoords()
        #       for shapes formed by one *or more* bezier polygons
        coords.Coordinates = mTblHelper.TableHelper.make_2d_array(1, 1, Point())
        coords.Flags = mTblHelper.TableHelper.make_2d_array(1, 1, PolygonFlags())
        coords.Coordinates[0] = pts
        coords.Flags[0] = flags

        mProps.Props.set_property(obj=bezier_poly, name="PolyPolygonBezier", value=coords)
        return bezier_poly

    @classmethod
    def draw_line(cls, slide: XDrawPage, x1: int, y1: int, x2: int, y2: int) -> XShape | None:
        """
        Draws a line

        Args:
            slide (XDrawPage): Slide
            x1 (int): Line start X position
            y1 (int): Line start Y position
            x2 (int): Line end X position
            y2 (int): Line end Y position

        Raises:
            ValueError: IF is a point and not a line.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        # make sure size is non-zero
        if (x1 == x2) and (y1 == y2):
            raise ValueError("Cannot create a line from a point")

        width = x2 - x1  # may be negative
        height = y2 - y1  # may be negative
        return cls.add_shape(
            slide=slide,
            shape_type=DrawingShapeKind.LINE_SHAPE,
            x=x1,
            y=y1,
            width=width,
            height=height,
        )

    @classmethod
    def draw_polar_line(cls, slide: XDrawPage, x: int, y: int, degrees: int, distance: int) -> XShape | None:
        """
        Draw a line from ``x``, ``y`` in the direction of degrees, for the specified distance
        degrees is measured clockwise from x-axis

        Args:
            slide (XDrawPage): Slide
            x (int): Shape X position
            y (int): Shape Y position
            degrees (int): Direction of degrees
            distance (int): Distance of line.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        xdist = round(math.cos(math.radians(degrees)) * distance)
        ydist = round(math.sin(math.radians(degrees)) * distance) * -1  # convert to negative
        return cls.draw_line(slide=slide, x1=x, y1=y, x2=x + xdist, y2=y + ydist)

    @classmethod
    def draw_lines(cls, slide: XDrawPage, xs: Sequence[int], ys: Sequence[int]) -> XShape | None:
        """
        Draw lines

        Args:
            slide (XDrawPage): Slide
            xs (Sequence[int]): Sequence of X positions in mm units.
            ys (Sequence[int]): Sequence of Y positions in mm units.

        Raises:
            IndexError: If ``xs`` and ``xy`` do not have the same number of elements.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        num_points = len(xs)
        if num_points != len(ys):
            raise IndexError("xs and ys must be the same length")

        pts: List[Point] = []
        for x, y in zip(xs, ys):
            # in 1/100 mm units
            pts.append(Point(x * 100, y * 100))

        # an array of Point arrays, one Point array for each line path
        line_paths = mTblHelper.TableHelper.make_2d_array(1, 1)
        line_paths[0] = pts

        # for a shape formed by from multiple connected lines
        poly_line = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.POLY_LINE_SHAPE, x=0, y=0, width=0, height=0
        )
        if poly_line is None:
            mLo.Lo.print("Failed to create PolyLineShape")
            return None
        mProps.Props.set_property(obj=poly_line, name="PolyPolygon", value=line_paths)
        return poly_line

    # region draw_text()
    @overload
    @classmethod
    def draw_text(cls, slide: XDrawPage, msg: str, x: int, y: int, width: int, height: int) -> XShape | None:
        ...

    @overload
    @classmethod
    def draw_text(
        cls, slide: XDrawPage, msg: str, x: int, y: int, width: int, height: int, font_size: int
    ) -> XShape | None:
        ...

    @classmethod
    def draw_text(
        cls, slide: XDrawPage, msg: str, x: int, y: int, width: int, height: int, font_size: int = 0
    ) -> XShape | None:
        """
        Draws Text

        Args:
            slide (XDrawPage): Slide
            msg (str): Text to draw
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.
            font_size (int, optional): Font size of text.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        shape = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.TEXT_SHAPE, x=x, y=y, width=width, height=height
        )
        if shape is None:
            mLo.Lo.print("Failed to create TextShape")
            return None
        cls.add_text(shape=shape, msg=msg, font_size=font_size)
        return shape

    # endregion draw_text()

    @classmethod
    def add_text(cls, shape: XShape, msg: str, font_size: int = 0, **props) -> None:
        """
        Add text to a shape

        Args:
            shape (XShape): Shape
            msg (str): Text to add
            font_size (int, optional): Font size.
            props (Any, optional): Any extra properties that will be applied to cursor (font) such as ``CharUnderline=1``
        """
        txt = mLo.Lo.qi(XText, shape, True)
        cursor = txt.createTextCursor()
        cursor.gotoEnd(False)
        if font_size > 0:
            mProps.Props.set_property(obj=cursor, name="CharHeight", value=font_size)

        for k, v in props.items():
            mProps.Props.set_property(obj=cursor, name=k, value=v)

        rng = mLo.Lo.qi(XTextRange, cursor, True)
        rng.setString(msg)

    @classmethod
    def add_connector(
        cls,
        slide: XDrawPage,
        shape1: XShape,
        shape2: XShape,
        start_conn: Draw.GluePointsKind | None = None,
        end_conn: Draw.GluePointsKind | None = None,
    ) -> XShape | None:
        """
        Add connector

        Args:
            slide (XDrawPage): Slide
            shape1 (XShape): First Shape to add connector to.
            shape2 (XShape): Second Shape to add connector to.
            start_conn (GluePointsKind | None, optional): Start connector kind. Defaluts to right.
            end_conn (GluePointsKind | None, optional): End connector kind. Defaults to left.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.

        Note:
            Properties for shape can be added or changed by using :py:meth"~.draw.Draw.set_shape_props`.

            For instance the default value is ``EndShape=ConnectorType.STANDARD``.
            This could be changed.

            ..code-block:: python

                Draw.set_shape_props(shape, EndShape=ConnectorType.CURVE)
        """
        if start_conn is None:
            start_conn = Draw.GluePointsKind.RIGHT
        if end_conn is None:
            end_conn = Draw.GluePointsKind.LEFT

        xconnector = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.CONNECTOR_SHAPE, x=0, y=0, width=0, height=0
        )
        if xconnector is None:
            mLo.Lo.print("Failed to create ConnectorShape")
            return None

        prop_set = mLo.Lo.qi(XPropertySet, xconnector, True)
        try:
            prop_set.setPropertyValue("StartShape", shape1)
            prop_set.setPropertyValue("StartGluePointIndex", int(start_conn))

            prop_set.setPropertyValue("EndShape", shape2)
            prop_set.setPropertyValue("EndGluePointIndex", int(end_conn))

            prop_set.setPropertyValue("EndShape", ConnectorType.STANDARD)
        except Exception as e:
            mLo.Lo.print("Could not connect the shapes")
            mLo.Lo.print(f"  {e}")

        return xconnector

    @staticmethod
    def get_glue_points(shape: XShape) -> Tuple[GluePoint2, ...] | None:
        """
        Gets Glue Points

        Args:
            shape (XShape): Shape

        Returns:
            Tuple[GluePoint2, ...] | None: Glue Points on success; Otherwise, ``None``.
        """
        gp_supp = mLo.Lo.qi(XGluePointsSupplier, shape, True)
        glue_pts = gp_supp.getGluePoints()

        num_gps = glue_pts.getCount()  # should be 4 by default
        if num_gps == 0:
            mLo.Lo.print("No glue points for this shape")
            return None

        gps: List[GluePoint2] = []
        for i in range(num_gps):
            try:
                # the original java was using qi for some unknown reason.
                # this makes no sense here because you can not query a struct for an interface
                # gps.append(mLo.Lo.qi(GluePoint2, glue_pts.getByIndex(i), True))

                gps.append(glue_pts.getByIndex(i))
            except Exception as e:
                mLo.Lo.print(f"Could not access glue point: {i}")
                mLo.Lo.print(f"  {e}")

        return tuple(gps)

    @classmethod
    def get_chart_shape(cls, slide: XDrawPage, x: int, y: int, width: int, height: int) -> XShape | None:
        """
        Gets a chart shape

        Args:
            slide (XDrawPage): Slide
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        shape = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.OLE2_SHAPE, x=x, y=y, width=width, height=height
        )
        if shape is None:
            mLo.Lo.print("Error getting shape for OLE2Shape")
            return None
        mProps.Props.set_property(obj=shape, name="CLSID", value=str(mLo.Lo.CLSID.CHART))  # a chart
        return shape

    @classmethod
    def draw_formula(cls, slide: XDrawPage, formula: str, x: int, y: int, width: int, height: int) -> XShape | None:
        """
        Draws a formula

        Args:
            slide (XDrawPage): Slide
            formula (str): Formula as string to draw/
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        shape = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.OLE2_SHAPE, x=x, y=y, width=width, height=height
        )
        if shape is None:
            mLo.Lo.print("Error getting shape for OLE2Shape")
            return None
        cls.set_shape_props(shape, name="CLSID", value=str(mLo.Lo.CLSID.MATH))  # a formula

        model = mLo.Lo.qi(XModel, mProps.Props.get_property(shape, "Model"), True)
        # mInfo.Info.show_services(obj_name="OLE2Shape Model", obj=model)
        mProps.Props.set_property(obj=model, name="Formula", value=formula)
        return shape

    @classmethod
    def draw_media(cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int, width: int, height: int) -> XShape | None:
        shape = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.MEDIA_SHAPE, x=x, y=y, width=width, height=height
        )

        # mProps.Props.show_obj_props(prop_kind="Shape", obj=shape)
        mLo.Lo.print(f'Loading media: "{fnm}"')
        cls.set_shape_props(shape, Loop=True, MediaURL=mFileIO.FileIO.fnm_to_url(fnm))

    @staticmethod
    def is_group(shape: XShape) -> bool:
        """
        Gets if a shape is a Group Shape

        Args:
            shape (XShape): Shape

        Returns:
            bool: ``True`` if shape is a group; Otherwise; ``False``.
        """
        if shape is None:
            mLo.Lo.print("Shape is none. Not able to check for group")
            return False
        return shape.getShapeType() == "com.sun.star.drawing.GroupShape"

    @staticmethod
    def combine_shape(doc: XComponent, shapes: XShapes, combine_op: Draw.ShapeCompKind) -> XShape | None:
        """
        Combines one or more shapes.

        Args:
            doc (XComponent): Document
            shapes (XShapes): Shapes to combine
            combine_op (ShapeCompKind): Combine Operation.

        Returns:
            XShape | None: New combined shape on success; Otherwise, ``None``.
        """
        # select the shapes for the dispatches to apply to
        sel_supp = mLo.Lo.qi(XSelectionSupplier, mGui.GUI.get_current_controller(doc), True)
        sel_supp.select(shapes)

        if combine_op == Draw.ShapeCompKind.INTERSECT:
            mLo.Lo.dispatch_cmd("Intersect")
        elif combine_op == Draw.ShapeCompKind.SUBTRACT:
            mLo.Lo.dispatch_cmd("Substract")  # misspelt!
        elif combine_op == Draw.ShapeCompKind.COMBINE:
            mLo.Lo.dispatch_cmd("Combine")
        else:
            mLo.Lo.dispatch_cmd("Merge")

        mLo.Lo.delay(500)  # give time for dispatches to arrive and be processed

        # extract the new single shape from the modified selection
        xs = mLo.Lo.qi(XShapes, sel_supp.getSelection(), True)
        try:
            combined_shape = mLo.Lo.qi(XShapes, xs.getByIndex(0))
        except Exception as e:
            mLo.Lo.print("Could not get combined shape")
            mLo.Lo.print(f"  {e}")
            return None
        return combined_shape

    @staticmethod
    def create_control_shape(
        label: str, x: int, y: int, width: int, height: int, shape_kind: FormControlKind | str, **props
    ) -> XControlShape:
        """
        Creates a control shape

        Args:
            label (str): Label to apply.
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.
            shape_kind (FormControlKind | str): The kind of control to create
            props (Any, optional): Any extra key value options to set on the Model of the control being created
                such as ``FontHeight=18.0, Name="BLA"``

        Returns:
            XControlShape: _description_
        """
        cshape = mLo.Lo.create_instance_msf(XControlShape, "com.sun.star.drawing.ControlShape", raise_err=True)
        cshape.setSize(Size(width * 100, height * 100))
        cshape.setPosition(Point(x * 100, y * 100))

        cmodel = mLo.Lo.create_instance_msf(XControlModel, f"com.sun.star.form.control.{shape_kind}")

        prop_set = mLo.Lo.qi(XPropertySet, cmodel, True)
        prop_set.setPropertyValue("DefaultControl", f"com.sun.star.form.control.{shape_kind}")
        prop_set.setPropertyValue("Name", "XXX")
        prop_set.setPropertyValue("Label", label)

        prop_set.setPropertyValue("FontHeight", 18.0)
        prop_set.setPropertyValue("FontName", mInfo.Info.get_font_general_name())

        for k, v in props.items():
            prop_set.setPropertyValue(k, v)

        cshape.setControl(cmodel)

        xbtn = mLo.Lo.qi(XButton, cmodel)
        if xbtn is None:
            mLo.Lo.print("XButton is None")

        # mProps.Props.show_props(title="Control model props", props=props)
        return cshape

    # endregion draw/add shape to a page

    # region custom shape addition using dispatch and JNA

    # Two methods not include here from java. addDispatchShape and createDispatchShape
    # These were omitted because the require third party Libs that all for automatic screen click and screen moused selecting

    # endregion

    # region presentation shapes
    @classmethod
    def set_master_footer(cls, master: XDrawPage, text: str) -> None:
        """
        Sets master footer text

        Args:
            master (XDrawPage): Master Draw Page
            text (str): Footer text
        """
        footer_shape = cls.find_shape_by_type(slide=master, shape_type=Draw.NameSpaceKind.SHAPE_TYPE_FOOTER)
        if footer_shape is None:
            mLo.Lo.print(f'Unable to find "{Draw.NameSpaceKind.SHAPE_TYPE_FOOTER}"')
            return
        txt_field = mLo.Lo.qi(XText, footer_shape, True)
        txt_field.setString(text)

    @classmethod
    def add_slide_number(cls, slide: XDrawPage) -> XShape | None:
        """
        Adds slide number to a slide

        Args:
            slide (XDrawPage): Slide

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        sz = cls.get_slide_size(slide)
        if sz is None:
            mLo.Lo.print("Unable to get slide size")
            return None
        width = 60
        height = 15
        return cls.add_pres_shape(
            slide=slide,
            shape_type=PresentationKind.SLIDE_NUMBER_SHAPE,
            x=sz.Width - width - 12,
            y=sz.Height - height - 4,
            width=width,
            height=height,
        )

    @classmethod
    def add_pres_shape(
        cls, slide: XDrawPage, shape_type: PresentationKind, x: int, y: int, width: int, height: int
    ) -> XShape | None:
        """
        Creates a shape from the "com.sun.star.presentation" package:

        Args:
            slide (XDrawPage): Slide
            shape_type (PresentationKind): Kind of presentation package to create.
            x (int): Shape X position in mm units.
            y (int): Shape Y position in mm units.
            width (int): Shape width in mm units.
            height (int): Shape height in mm units.

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        cls.warns_position(slide=slide, x=x, y=y)
        shape = mLo.Lo.create_instance_msf(XShape, shape_type.to_namespace())
        if shape is not None:
            slide.add(shape)
            cls.set_position(shape, x, y)
            cls.set_size(shape, width, height)
        return shape

    # endregion presentation shapes

    # region get/set drawing properties
    @staticmethod
    def get_position(shape: XShape) -> Point:
        """
        Gets position in mm units

        Args:
            shape (XShape): Shape

        Returns:
            Point: Position as Point in mm units
        """
        pt = shape.getPosition()
        # convert to mm
        return Point(round(pt.X / 100, round(pt.Y / 100)))

    @staticmethod
    def get_size(shape: XShape) -> Size:
        """
        Gets Size in mm units

        Args:
            shape (XShape): Shape

        Returns:
            Size: Size in mm units
        """
        sz = shape.getSize()
        # convert to mm
        return Size(round(sz.Width / 100), round(sz.Height / 100))

    @staticmethod
    def print_point(pt: Point) -> None:
        """
        Prints point to console in mm units

        Args:
            pt (Point): Point object
        """
        print(f"  Point (mm): [{round(pt.X/100)}, {round(pt.Y/100)}]")

    @staticmethod
    def print_size(sz: Size) -> None:
        """
        Prints size to console in mm units

        Args:
            sz (Size): Size object
        """
        print(f"  Size (mm): [{round(sz.Width/100)}, {round(sz.Height/100)}]")

    @classmethod
    def report_pos_size(cls, shape: XShape) -> None:
        """
        Prints shape information to the console

        Args:
            shape (XShape): Shape
        """
        if shape is None:
            print("The shape is null")
            return
        print(f'Shape Name: {mProps.Props.get_property(obj=shape, name="Name")}')
        print(f"  Type: {shape.getShapeType()}")
        cls.print_point(shape.getPosition())
        cls.print_size(shape.getSize())

    # region set_position()

    @overload
    @staticmethod
    def set_position(shape: XShape, pt: Point) -> None:
        ...

    @overload
    @staticmethod
    def set_position(shape: XShape, x: int, y: int) -> None:
        ...

    @staticmethod
    def set_position(*args, **kwargs) -> None:
        """
        Sets Position of shape

        Args:
            shape (XShape): Shape
            pt (point): Point that contains x and y positions.
            x (int): X position
            y (int): Y Position

        Note:
            Positions are NOT in mm units when passed into this method.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "pt", "x", "y")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_position() got an unexpected keyword argument")
            ka[1] = kwargs.get("shape", None)
            keys = ("pt", "x")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = ka.get("y", None)
            return ka

        if not count in (2, 3):
            raise TypeError("set_position() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if count == 2:
            # def set_position(shape: XShape, pt: Point)
            pt_in = cast(Point, kargs[2])
            pt = Point(pt_in.X * 100, pt_in.Y * 100)
        else:
            # def set_position(shape: XShape, x:int, y: int)
            pt = Point(kargs[2] * 100, kargs[3] * 100)
        cast(XShape, kargs[1]).setPosition(pt)

    # endregion set_position()

    # region set_size()

    @overload
    @staticmethod
    def set_size(shape: XShape, sz: Size) -> None:
        ...

    @overload
    @staticmethod
    def set_size(shape: XShape, width: int, height: int) -> None:
        ...

    @staticmethod
    def set_size(*args, **kwargs) -> None:
        """
        Sets set_size of shape

        Args:
            shape (XShape): Shape
            sz (Size): Size that contains width and height positions.
            width (int): Width position
            height (int): Height Position

        Note:
            Positions are NOT in mm units when passed into this method.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "sz", "width", "height")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_size() got an unexpected keyword argument")
            ka[1] = kwargs.get("shape", None)
            keys = ("sz", "width")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = ka.get("height", None)
            return ka

        if not count in (2, 3):
            raise TypeError("set_size() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if count == 2:
            # def set_size(shape: XShape, sz: Size)
            sz_in = cast(Size, kargs[2])
            sz = Size(sz_in.Width * 100, sz_in.Height * 100)
        else:
            # def set_size(shape: XShape, width:int, height: int)
            sz = Size(kargs[2] * 100, kargs[3] * 100)
        cast(XShape, kargs[1]).setSize(sz)

    # endregion set_size()

    @staticmethod
    def set_style(shape: XShape, graphic_styles: XNameContainer, style_name: GraphicStyleKind | str) -> None:
        """
        Set the graphic style for a shape

        Args:
            shape (XShape): Shape
            graphic_styles (XNameContainer): Graphic styles
            style_name (GraphicStyleKind | str): Graphic Style Name
        """
        try:
            style = mLo.Lo.qi(XStyle, graphic_styles.getByName(str(style_name)), True)
            mProps.Props.set_property(obj=shape, name="Style", value=style)
        except Exception as e:
            mLo.Lo.print(f'Could not set the style to "{style_name}"')

    @staticmethod
    def get_text_properties(shape: XShape) -> XPropertySet:
        """
        Gets the properties associated with the text area inside the shape.

        Args:
            shape (XShape): Shape

        Returns:
            XPropertySet: Property Set
        """
        xtxt = mLo.Lo.qi(XText, shape, True)
        cursor = xtxt.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        xrng = mLo.Lo.qi(XTextRange, cursor, True)
        return mLo.Lo.qi(XPropertySet, xrng, True)

    @staticmethod
    def get_line_color(shape: XShape) -> mColor.Color | None:
        """
        Gets the line color of a shpae.

        Args:
            shape (XShape): Shape

        Returns:
            mColor.Color | None: Color on success; Otherwise, ``None``.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            c = mColor.Color(int(props.getPropertyValue("LineColor")))
            return c
        except Exception as e:
            mLo.Lo.print("Could not access line color")
            mLo.Lo.print(f"  {e}")
        return None

    @staticmethod
    def set_dashed_line(shape: XShape, is_dashed: bool) -> None:
        """
        Set a dashed line

        Args:
            shape (XShape): Shape
            is_dashed (bool): Determines if line is to be dashed or solid.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            if is_dashed:
                ld = LineDash()
                ld.Dots = 0
                ld.DotLen = 100
                ld.Dashes = 5
                ld.DashLen = 200
                ld.Distance = 200
                props.setPropertyValue("LineStyle", LineStyle.DASH)
                props.setPropertyValue("LineDash", ld)
            else:
                # switch to solid line
                props.setPropertyValue("LineStyle", LineStyle.SOLID)
        except Exception as e:
            mLo.Lo.print("Could not set dashed line property")
            mLo.Lo.print(f"  {e}")

    @staticmethod
    def get_line_thickness(shape: XShape) -> int:
        """
        Gets line thickness of a shape.

        Args:
            shape (XShape): Shape

        Returns:
            int: Line Thickness on success; Otherwise, ``0``.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            return int(props.getPropertyValue("LineWidth"))
        except Exception as e:
            mLo.Lo.print("Could not access line thickness")
            mLo.Lo.print(f"  {e}")
        return 0

    @staticmethod
    def get_fill_color(shape: XShape) -> mColor.Color | None:
        """
        Gets the fill color of a shpae.

        Args:
            shape (XShape): Shape

        Returns:
            mColor.Color | None: Color on success; Otherwise, ``None``.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            c = mColor.Color(int(props.getPropertyValue("FillColor")))
            return c
        except Exception as e:
            mLo.Lo.print("Could not access fill color")
            mLo.Lo.print(f"  {e}")
        return None

    @staticmethod
    def set_transparency(shape: XShape, level: Intensity) -> None:
        """
        Sets the transparency level for the shape.
        Higher level means more transparent.

        Args:
            shape (XShape): Shape
            level (Intensity): Transparency value
        """
        mProps.Props.set_property(obj=shape, name="FillTransparence", value=level.Value)

    # region set_gradient_color()

    @staticmethod
    def _set_gradient_color_name(shape: XShape, name: DrawingGradientKind | str) -> None:
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            props.setPropertyValue("FillStyle", FillStyle.GRADIENT)
            props.setPropertyValue("FillGradientName", str(name))
        except IllegalArgumentException:
            mLo.Lo.print(f'"{name}" is not a recognized gradient color name')
        except Exception as e:
            mLo.Lo.print(f'Could not set gradient color to "{name}"')
            mLo.Lo.print(f"  {e}")

    @staticmethod
    def _set_gradient_color_colors(
        shape: XShape, start_color: mColor.Color, end_color: mColor.Color, angle: Angle
    ) -> None:
        props = mLo.Lo.qi(XPropertySet, shape, True)

        grad = Gradient()
        grad.Style = GradientStyle.LINEAR
        grad.StartColor = start_color
        grad.EndColor = end_color

        grad.Angle = angle.Value * 10  # in 1/10 degree units
        grad.Border = 0
        grad.XOffset = 0
        grad.YOffset = 0
        grad.StartIntensity = 100
        grad.EndIntensity = 100
        grad.StepCount = 10

        try:
            props.setPropertyValue("FillStyle", FillStyle.GRADIENT)
            props.setPropertyValue("FillGradient", grad)
        except Exception as e:
            mLo.Lo.print("Could not set gradient colors")
            mLo.Lo.print(f"  {e}")

    @overload
    @classmethod
    def set_gradient_color(cls, shape: XShape, name: DrawingGradientKind | str) -> None:
        ...

    @overload
    @classmethod
    def set_gradient_color(cls, shape: XShape, start_color: mColor.Color, end_color: mColor.Color) -> None:
        ...

    @overload
    @classmethod
    def set_gradient_color(
        cls, shape: XShape, start_color: mColor.Color, end_color: mColor.Color, angle: Angle
    ) -> None:
        ...

    @classmethod
    def set_gradient_color(cls, *args, **kwargs) -> None:
        """
        Set the gradient color of the shape

        Args:
            shape (XShape): Shape
            name (DrawingGradientKind | str): Gradient color name.
            start_color (mColor.Color): Start Color
            end_color (mColor.Color): End Color
            angle (Angle): Angle

        Note:
            When using Graident Name.

            Getting the gradient color name can be a bit challenging.
            ``DrawingGradientKind`` contains name displayed in the Graident color menu of Draw.

            The Easies way to get the colors is to open Draw and see what gradient color names are available
            on your system.
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "name", "start_color", "end_color", "angle")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_gradient_color() got an unexpected keyword argument")
            ka[1] = kwargs.get("shape", None)
            keys = ("name", "start_color")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = ka.get("end_color", None)
            if count == 3:
                return ka
            ka[4] = ka.get("angle", None)
            return ka

        if not count in (2, 3, 4):
            raise TypeError("set_gradient_color() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            cls._set_gradient_color_name(kargs[1], kargs[2])
            return

        if count == 3:
            angle = Angle(0)
        else:
            angle = cast(Angle, kargs[4])
        cls._set_gradient_color_colors(shape=kargs[1], start_color=kargs[2], end_color=kargs[3], angle=angle)

    # endregion set_gradient_color()

    @staticmethod
    def set_hatch_color(shape: XShape, name: DrawingHatchingKind | str) -> None:
        """
        Set hatching color of a shape.

        Args:
            shape (XShape): Shape
            name (DrawingHatchingKind | str): Hatching Name

        Note:
            Getting the hatching color name can be a bit challenging.
            ``DrawingHatchingKind`` contains name displayed in the Hatching color menu of Draw.

            The Easies way to get the colors is to open Draw and see what gradient color names are available
            on your system.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            props.setPropertyValue("FillStyle", FillStyle.HATCH)
            props.setPropertyValue("FillHatchName", str(name))
        except IllegalArgumentException:
            mLo.Lo.print(f'"{name}" is not a recognized hatching name')
        except Exception as e:
            mLo.Lo.print(f'Could not set hatching color to "{name}"')
            mLo.Lo.print(f"  {e}")
        return None

    @staticmethod
    def set_bitmap_color(shape: XShape, name: DrawingBitmapKind | str) -> None:
        """
        Set bitmap color of a shape.

        Args:
            shape (XShape): Shape
            name (DrawingBitmapKind | str): Bitmap Name

        Note:
            Getting the bitmap color name can be a bit challenging.
            ``DrawingBitmapKind`` contains name displayed in the Bitmap color menu of Draw.

            The Easies way to get the colors is to open Draw and see what bitmap color names are available
            on your system.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            props.setPropertyValue("FillStyle", FillStyle.BITMAP)
            props.setPropertyValue("FillBitmapName", str(name))
        except IllegalArgumentException:
            mLo.Lo.print(f'"{name}" is not a recognized bitmap name')
        except Exception as e:
            mLo.Lo.print(f'Could not set bitmap color to "{name}"')
            mLo.Lo.print(f"  {e}")
        return None

    @staticmethod
    def set_bitmap_file_color(shape: XShape, fnm: PathOrStr) -> None:
        """
        Set bitmap color from file.

        Args:
            shape (XShape): Shape
            fnm (PathOrStr): path to file.
        """
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            props.setPropertyValue("FillStyle", FillStyle.BITMAP)
            props.setPropertyValue("FillBitmapURL", mFileIO.FileIO.fnm_to_url(fnm))
        except Exception as e:
            mLo.Lo.print(f'Could not set bitmap color using  "{fnm}"')
            mLo.Lo.print(f"  {e}")
        return None

    @classmethod
    def set_line_style(cls, shape: XShape, style: LineStyle) -> None:
        """
        Set the line style for a shape

        Args:
            shape (XShape): Shape
            style (LineStyle): Line Style
        """
        cls.set_shape_props(shape=shape, LineStyle=style)

    @classmethod
    def set_visible(cls, shape: XShape, is_visible: bool) -> None:
        """
        Set the line style for a shape

        Args:
            shape (XShape): Shape
            is_visible (bool): Set is shape is visible or not.
        """
        cls.set_shape_props(shape=shape, Visible=is_visible)

    # "RotateAngle" is deprecated but is much simpler
    # than the matrix approach, and works correctly
    # for rotations around the center

    @classmethod
    def set_angle(cls, shape: XShape, angle: Angle) -> None:
        """
        Set the line style for a shape

        Args:
            shape (XShape): Shape
            is_visible (bool): Set is shape is visibel or not.
        """
        cls.set_shape_props(shape=shape, RotateAngle=angle.Value * 100)

    @staticmethod
    def get_rotation(shape: XShape) -> Angle:
        """
        Gets the rotation of a shape

        Args:
            shape (XShape): Shape

        Returns:
            Angle: Rotation angle.
        """
        r_angle = int(mProps.Props.get_property(shape, "RotateAngle"))
        return Angle(round(r_angle / 100))

    @staticmethod
    def get_transformation(shape: XShape) -> HomogenMatrix3:
        """
        Gets a transformation matrix which seems to represent a clockwise rotation.

        Homogeneous matrix has three homogeneous lines

        Args:
            shape (XShape): Shape

        Returns:
            HomogenMatrix3: Matrix
        """
        #     Returns a transformation matrix, which seems to
        #     represent a clockwise rotation:
        #     cos(t)  sin(t) x
        #    -sin(t)  cos(t) y
        #       0       0    1
        return mProps.Props.get_property(shape, "Transformation")

    @staticmethod
    def print_matrix(mat: HomogenMatrix3) -> None:
        """
        Prints matrix to console

        Args:
            mat (HomogenMatrix3): Matrix
        """
        print("Transformation Matrix:")
        print(f"\t{mat.Line1.Column1:10.2f}\t{mat.Line1.Column2:10.2f}\t{mat.Line1.Column3:10.2f}")
        print(f"\t{mat.Line2.Column1:10.2f}\t{mat.Line2.Column2:10.2f}\t{mat.Line2.Column3:10.2f}")
        print(f"\t{mat.Line3.Column1:10.2f}\t{mat.Line3.Column2:10.2f}\t{mat.Line3.Column3:10.2f}")

        rad_angle = math.atan2(mat.Line2.Column1, mat.Line1.Column1)
        #       sin(t), cos(t)
        curr_angle = round(math.degrees(rad_angle))
        print(f"  Current angle: {curr_angle}")
        print()

    # endregion get/set drawing properties

    # region draw an image
    # region draw_image()
    @classmethod
    def _draw_image_path(cls, slide: XDrawPage, fnm: PathOrStr) -> XShape | None:
        slide_size = cls.get_slide_size(slide)
        if slide_size is None:
            mLo.Lo.print("Unable to get slide size")
            return None
        try:
            im_size = mImgLo.ImagesLo.get_size_100mm(fnm)
        except Exception as e:
            mLo.Lo.print(f'Could not calculate size of "{fnm}"')
            return None
        im_width = round(im_size.Width / 100)  # in mm units
        im_height = round(im_size.Height / 100)
        x = round((slide_size.Width - im_width) / 2)
        y = round((slide_size.Height - im_height) / 2)
        return cls._draw_image_path_x_y_w_h(slide=slide, fnm=fnm, x=x, y=y, width=im_width, height=im_height)

    @classmethod
    def _draw_image_path_x_y(cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int) -> XShape | None:
        try:
            im_size = mImgLo.ImagesLo.get_size_100mm(fnm)
        except Exception as e:
            mLo.Lo.print(f'Could not calculate size of "{fnm}"')
            return None
        return cls._draw_image_path_x_y_w_h(
            slide=slide, fnm=fnm, x=x, y=y, width=round(im_size.Width / 100), height=round(im_size.Height / 100)
        )

    @classmethod
    def _draw_image_path_x_y_w_h(
        cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int, width: int, height: int
    ) -> XShape | None:
        # units in mm's
        mLo.Lo.print(f'Adding the picture "{fnm}"')
        im_shape = cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.GRAPHIC_OBJECT_SHAPE, x=x, y=y, width=width, height=height
        )
        if im_shape is None:
            mLo.Lo.print("Unable to add shape")
            return None
        cls.set_image(im_shape, fnm)
        cls.set_line_style(shape=im_shape, style=LineStyle.NONE)
        return im_shape

    @overload
    @classmethod
    def draw_image(cls, slide: XDrawPage, fnm: PathOrStr) -> XShape | None:
        ...

    @overload
    @classmethod
    def draw_image(cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int) -> XShape | None:
        ...

    @overload
    @classmethod
    def draw_image(cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int, width: int, height: int) -> XShape | None:
        ...

    @classmethod
    def draw_image(cls, *args, **kwargs) -> XShape | None:
        """
        Draws an image

        Args:
            slide (XDrawPage): Slide
            fnm (PathOrStr): Path to image
            x (int): Shape X position
            y (int): Shape Y position
            width (int): Shape width
            height (int): Shape height

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``
        """
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("slide", "fnm", "x", "y", "width", "height")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("draw_image() got an unexpected keyword argument")
            ka[1] = kwargs.get("slide", None)
            ka[2] = kwargs.get("fnm", None)
            if count == 2:
                return ka
            ka[3] = kwargs.get("x", None)
            ka[4] = kwargs.get("y", None)
            if count == 4:
                return ka
            ka[5] = kwargs.get("width", None)
            ka[6] = kwargs.get("height", None)
            return ka

        if not count in (2, 4, 6):
            raise TypeError("draw_image() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return cls._draw_image_path(slide=kargs[1], fnm=kargs[2])
        if count == 4:
            return cls._draw_image_path_x_y(slide=kargs[1], fnm=kargs[2], x=kargs[3], y=kargs[4])
        return cls._draw_image_path_x_y_w_h(
            slide=kargs[1], fnm=kargs[2], x=kargs[3], y=kargs[4], width=kargs[5], height=kargs[6]
        )

    # endregion draw_image()

    @staticmethod
    def set_image(shape: XShape, fnm: PathOrStr) -> None:
        """
        Sets the image of a shape

        Args:
            shape (XShape): Shape
            fnm (PathOrStr): Path to image.
        """
        # GraphicURL is Deprecated using Graphic instead.
        # https://tinyurl.com/2qaqs2nr#a6312a2da62e2c67c90d5576502117906
        graphic = mImgLo.ImagesLo.load_graphic_file(fnm)

        mProps.Props.set_property(shape, "Graphic", graphic)

    @classmethod
    def draw_image_offset(
        cls, slide: XDrawPage, fnm: PathOrStr, xoffset: ImageOffset, yoffset: ImageOffset
    ) -> XShape | None:
        """
        Insert the specified picture onto the slide page in the doc
        presentation document. Use the supplied (x, y) offsets to locate the
        top-left of the image.

        Args:
            slide (XDrawPage): Slide
            fnm (PathOrStr): Path to image.
            xoffset (ImageOffset): X Offset
            yoffset (ImageOffset): Y Offset

        Returns:
            XShape | None: Shape on success; Otherwise, ``None``.
        """
        slide_size = cls.get_slide_size(slide)
        if slide_size is None:
            mLo.Lo.print("Unalble to get slide size")
            return None
        x = round(slide_size.Width * xoffset.Value)  # in mm units
        y = round(slide_size.Height * yoffset.Value)

        max_width = slide_size.Width - x
        max_height = slide_size.Height - y

        im_size = mImgLo.ImagesLo.calc_scale(fnm=fnm, max_width=max_width, max_height=max_height)
        if im_size is None:
            mLo.Lo.print(f'Unalbe to calc image size for "{fnm}"')
            return None
        return cls._draw_image_path_x_y_w_h(slide=slide, fnm=fnm, x=x, y=y, width=im_size.Width, height=im_size.Height)

    @staticmethod
    def is_image(shape: XShape) -> bool:
        """
        Gets if a shape is an image (GraphicObjectShape).

        Args:
            shape (XShape): Shape

        Returns:
            bool: ``True`` if shape is image; Otherwise, ``False``.
        """
        if shape is None:
            return False
        return shape.getShapeType() == "com.sun.star.drawing.GraphicObjectShape"

    # endregion draw an image

    # region helper
    @staticmethod
    def set_shape_props(shape: XShape, **props) -> None:
        """
        Sets properties on a shape

        Args:
            shape (XShape): Shape object
            props (Any): Key value pairs of property name and property value

        Example:
            .. code-block:: python

                Draw.set_shape_props(shape, Loop=True, MediaURL=FileIO.fnm_to_url(fnm))
        """
        if shape is None:
            return
        for k, v in props.items():
            mProps.Props.set_property(obj=shape, name=k, value=v)

    # endregion helper
