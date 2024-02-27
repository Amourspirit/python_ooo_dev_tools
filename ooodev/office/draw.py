# region Imports
from __future__ import annotations
import contextlib
from pathlib import Path
import time
from typing import List, Sequence, Tuple, cast, overload, TYPE_CHECKING, Union
import math

import uno

# from com.sun.star.awt import XButton
from com.sun.star.animations import XAnimationNode
from com.sun.star.animations import XAnimationNodeSupplier
from com.sun.star.awt import XControlModel
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XIndexContainer
from com.sun.star.container import XNameContainer
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
from com.sun.star.form import XFormsSupplier
from com.sun.star.frame import XComponentLoader
from com.sun.star.frame import XController
from com.sun.star.frame import XModel
from com.sun.star.graphic import XGraphic
from com.sun.star.lang import XComponent
from com.sun.star.lang import XSingleServiceFactory
from com.sun.star.presentation import XCustomPresentationSupplier
from com.sun.star.presentation import XHandoutMasterSupplier
from com.sun.star.presentation import XPresentation2
from com.sun.star.presentation import XPresentationPage
from com.sun.star.presentation import XPresentationSupplier
from com.sun.star.presentation import XSlideShowController
from com.sun.star.style import XStyle
from com.sun.star.text import XText
from com.sun.star.text import XTextRange
from com.sun.star.view import XSelectionSupplier

from ooo.dyn.awt.gradient import Gradient as Gradient
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.awt.point import Point as Point
from ooo.dyn.awt.size import Size as UnoSize
from ooo.dyn.drawing.connector_type import ConnectorType as ConnectorType
from ooo.dyn.drawing.fill_style import FillStyle as FillStyle
from ooo.dyn.drawing.glue_point2 import GluePoint2 as GluePoint2
from ooo.dyn.drawing.homogen_matrix3 import HomogenMatrix3 as HomogenMatrix3
from ooo.dyn.drawing.line_dash import LineDash as LineDash
from ooo.dyn.drawing.line_style import LineStyle as LineStyle
from ooo.dyn.drawing.poly_polygon_bezier_coords import PolyPolygonBezierCoords as PolyPolygonBezierCoords
from ooo.dyn.drawing.polygon_flags import PolygonFlags as PolygonFlags
from ooo.dyn.lang.illegal_argument_exception import IllegalArgumentException
from ooo.dyn.presentation.animation_speed import AnimationSpeed as AnimationSpeed
from ooo.dyn.presentation.fade_effect import FadeEffect as FadeEffect
from ooo.dyn.container.no_such_element_exception import NoSuchElementException
from ooo.dyn.lang.index_out_of_bounds_exception import IndexOutOfBoundsException


from ooodev.cfg.config import Config  # singleton class.
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.draw_named_event import DrawNamedEvent
from ooodev.events.event_singleton import _Events
from ooodev.exceptions import ex as mEx
from ooodev.units.angle import Angle as Angle
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.unit_pt import UnitPT
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import color as mColor
from ooodev.utils import file_io as mFileIO
from ooodev.utils import gen_util as gUtil
from ooodev.utils import gui as mGui
from ooodev.utils import images_lo as mImgLo
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils import table_helper as mTableHelper
from ooodev.utils.data_type.image_offset import ImageOffset as ImageOffset
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.poly_sides import PolySides as PolySides
from ooodev.utils.data_type.size import Size
from ooodev.utils.dispatch.shape_dispatch_kind import ShapeDispatchKind as ShapeDispatchKind
from ooodev.loader.inst import lo_inst as mLoInst
from ooodev.utils.kind.drawing_bitmap_kind import DrawingBitmapKind as DrawingBitmapKind
from ooodev.utils.kind.drawing_gradient_kind import DrawingGradientKind as DrawingGradientKind
from ooodev.utils.kind.drawing_hatching_kind import DrawingHatchingKind as DrawingHatchingKind
from ooodev.utils.kind.drawing_layer_kind import DrawingLayerKind as DrawingLayerKind
from ooodev.utils.kind.drawing_name_space_kind import DrawingNameSpaceKind as DrawingNameSpaceKind
from ooodev.utils.kind.drawing_shape_kind import DrawingShapeKind as DrawingShapeKind
from ooodev.utils.kind.drawing_slide_show_kind import DrawingSlideShowKind as DrawingSlideShowKind
from ooodev.utils.kind.form_control_kind import FormControlKind as FormControlKind
from ooodev.utils.kind.glue_points_kind import GluePointsKind as GluePointsKind
from ooodev.utils.kind.graphic_style_kind import GraphicStyleKind as GraphicStyleKind
from ooodev.utils.kind.presentation_kind import PresentationKind as PresentationKind
from ooodev.utils.kind.presentation_layout_kind import PresentationLayoutKind as PresentationLayoutKind
from ooodev.utils.kind.shape_comb_kind import ShapeCombKind as ShapeCombKind
from ooodev.utils.type_var import PathOrStr

if TYPE_CHECKING:
    from ooodev.proto.dispatch_shape import DispatchShape
    from ooodev.proto.size_obj import SizeObj
    from ooodev.units.unit_obj import UnitT

# endregion Imports


class Draw:
    """Draw Class"""

    # region Constants
    POLY_RADIUS: int = 20
    """Default Poly Radius"""

    # endregion Constants
    @staticmethod
    def _create_name(name: str, gen_len: int = 10) -> str:
        """
        Creates a name.

        Make a unique string by appending a number to the supplied name.

        |lo_safe|

        Args:
            name (str): Name to prepend.
            gen_len (int): Length of random string to append. Default ``10``

        Returns:
            str: a name not in container.
        """
        return f"{name}_{gUtil.Util.generate_random_string(gen_len)}"

    @staticmethod
    def _get_unit_mm_int(value: UnitT | float, min_val: int = -9999) -> int:
        """Gets the value in mm as an int.

        |lo_safe|

        Args:
            value (UnitT | float): UnitT or float
            min_val (int, optional): The min value allowed.
                If not != -9999 then value must be greater than or equal to ``min_val``.
                Defaults to -9999.

        Raises:
            ValueError: If min_val != -9999 and value < min_val

        Returns:
            int: mm units as int.
        """
        with contextlib.suppress(AttributeError):
            result = value.get_value_mm()  # type: ignore
            i = round(result)
            if min_val != -9999 and i < min_val:
                raise ValueError("Value must be positive")
            return i
        i = round(value)  # type: ignore
        if min_val != -9999 and i < min_val:
            raise ValueError("Value must be positive")
        return i

    @staticmethod
    def _get_unit_pt(value: UnitT | float) -> float:
        """Lo Safe Method."""
        with contextlib.suppress(AttributeError):
            return value.get_value_pt()  # type: ignore
        return value  # type: ignore

    @staticmethod
    def _get_unit_mm_float(value: UnitT | float) -> float:
        """Lo Safe Method."""
        with contextlib.suppress(AttributeError):
            result = value.get_value_mm()  # type: ignore
            return round(result)
        return round(value)  # type: ignore

    @staticmethod
    def _get_mm100_obj_from_mm(value: UnitT | float, min_value: int = -9999) -> UnitMM100:
        """
        Gets a UnitMM100 object from mm.

        |lo_safe|

        Args:
            value (UnitT | float): Units in mm or UnitT
            min_value (int, optional): The min value in ``1/100 mm`.
                If not != -9999 then value must be greater than or equal to ``min_val``.
                Defaults to -9999.

        Raises:
            ValueError: If min_val != -9999 and value < min_val

        Returns:
            UnitMM100: MM 100 units.
        """
        with contextlib.suppress(AttributeError):
            result = value.get_value_mm100()  # type: ignore
            if min_value != -9999 and result < min_value:
                raise ValueError("Value must be positive")
            return UnitMM100(result)
        result = UnitMM100.from_mm(value)  # type: ignore
        if min_value != -9999 and result.value < min_value:
            raise ValueError("Value must be positive")
        return result

    @staticmethod
    def _get_pt_obj_from_pt(value: UnitT | float) -> UnitPT:
        """Lo Safe Method."""
        with contextlib.suppress(AttributeError):
            result = value.get_value_pt()  # type: ignore
            return UnitPT(result)
        return UnitPT.from_pt(value)  # type: ignore

    # region open, create, save draw/impress doc
    @staticmethod
    def is_shapes_based(doc: XComponent) -> bool:
        """
        Gets if the document is supports Draw or Impress.

        |lo_safe|

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
        Gets if the document is a Draw document.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Returns:
            bool: ``True`` if is Draw document; Otherwise, ``False``.
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW)

    @staticmethod
    def is_impress(doc: XComponent) -> bool:
        """
        Gets if the document is an Impress document.

        |lo_safe|

        Args:
            doc (XComponent): Document

        Returns:
            bool: ``True`` if is Impress document; Otherwise, ``False``
        """
        return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.IMPRESS)

    # region create_draw_doc()
    @overload
    @staticmethod
    def create_draw_doc() -> XComponent:
        """
        Creates a new Draw document.

        |lo_unsafe|

        Returns:
            XComponent: Component representing document
        """
        ...

    @overload
    @staticmethod
    def create_draw_doc(loader: XComponentLoader) -> XComponent:
        """
        Creates a new Draw document.

        |lo_unsafe|

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            XComponent: Component representing document
        """
        ...

    @overload
    @staticmethod
    def create_draw_doc(lo_inst: mLoInst.LoInst) -> XComponent:
        """
        Creates a new Draw document.

        |lo_unsafe|

        Args:
            lo_inst (LoInst): Lo Instance. Usually a new instance generated with :py:meth:`Lo.create_lo_instance() <ooodev.utils.lo.Lo.create_lo_instance>`.

        Returns:
            XComponent: Component representing document.
        """
        ...

    @staticmethod
    def create_draw_doc(*args, **kwargs) -> XComponent:
        """
        Creates a new Draw document.

        |lo_unsafe|

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            XComponent: Component representing document
        """
        # 0 or 1 args
        arguments = list(args)
        arguments.extend(kwargs.values())
        count = len(arguments)
        if count == 0:
            return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW)
        if count == 1:
            arg = arguments[0]
            if mLo.Lo.is_uno_interfaces(arg, XComponentLoader):
                return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW, loader=arg)
            if isinstance(arg, mLoInst.LoInst):
                return arg.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW)
        raise TypeError("create_draw_doc() got an unexpected argument")

    # endregion create_draw_doc()

    # region create_impress_doc()
    @overload
    @staticmethod
    def create_impress_doc() -> XComponent:
        """
        Creates a new Impress document.

        |lo_unsafe|

        Returns:
            XComponent: Component representing document.
        """
        ...

    @overload
    @staticmethod
    def create_impress_doc(loader: XComponentLoader) -> XComponent:
        """
        Creates a new Impress document.

        |lo_unsafe|

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            XComponent: Component representing document.
        """
        ...

    @overload
    @staticmethod
    def create_impress_doc(lo_inst: mLoInst.LoInst) -> XComponent:
        """
        Creates a new Impress document.

        |lo_unsafe|

        Args:
            lo_inst (LoInst): Lo Instance. Usually a new instance generated with :py:meth:`Lo.create_lo_instance() <ooodev.utils.lo.Lo.create_lo_instance>`.

        Returns:
            XComponent: Component representing document.
        """
        ...

    @staticmethod
    def create_impress_doc(*args, **kwargs) -> XComponent:
        """
        Creates a new Impress document.

        |lo_unsafe|

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            XComponent: Component representing document
        """
        arguments = list(args)
        arguments.extend(kwargs.values())
        count = len(arguments)
        if count == 0:
            return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.IMPRESS)
        if count == 1:
            arg = arguments[0]
            if mLo.Lo.is_uno_interfaces(arg, XComponentLoader):
                return mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.IMPRESS, loader=arg)
            if isinstance(arg, mLoInst.LoInst):
                return arg.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW)
        raise TypeError("create_impress_doc() got an unexpected argument")

    # endregion create_impress_doc()

    @staticmethod
    def get_slide_template_path() -> str:
        """
        Gets Slide template directory.

        |lo_unsafe|

        Returns:
            str: Path as string.

        See Also:
            :py:class:`~.cfg.config.Config`
        """
        # Config meta class allows it to be called without args.
        p = Path(mInfo.Info.get_office_dir(), Config().slide_template_path)  # type: ignore
        return str(p)

    @staticmethod
    def save_page(page: XDrawPage, fnm: PathOrStr, mime_type: str, filter_data: dict | None = None) -> None:
        """
        Saves a Draw page to file.

        |lo_unsafe|

        Args:
            page (XDrawPage): Page to save
            fnm (PathOrStr): Path to save page as
            mime_type (str): Mime Type of page to save as such as ``image/jpeg`` or ``image/png``.
            filter_data (dict, optional): Filter data. Defaults to ``None``.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:

        See Also:
            - :ref:`ooodev.draw.filter.export_png`
            - :ref:`ooodev.draw.filter.export_jpg`
            - :py:meth:`ooodev.utils.images_lo.ImagesLo.change_to_mime`.
            - :py:meth:`ooodev.utils.images_lo.ImagesLo.get_dpi_width_height`.

        .. versionchanged:: 0.21.3
            Added `filter_data` parameter.
        """
        try:
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

            if filter_data is None:
                # export the page by converting to the specified mime type
                props = mProps.Props.make_props(MediaType=mime_type, URL=save_file_url)
            else:
                # set the filter data
                filter_props = mProps.Props.make_props(**filter_data)
                props = mProps.Props.make_props(
                    MediaType=mime_type,
                    URL=save_file_url,
                    FilterData=uno.Any("[]com.sun.star.beans.PropertyValue", filter_props),  # type: ignore
                )

            gef.filter(props)
            mLo.Lo.print("Export Complete")
        except Exception as e:
            raise mEx.DrawError("Error saving page") from e

    # endregion open, create, save draw/impress doc

    # region methods related to document/multiple slides/pages
    @staticmethod
    def get_slides(doc: XComponent) -> XDrawPages:
        """
        Gets the draw pages of a document.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawPageMissingError: If there are no draw pages.
            DrawPageError: If any other error occurs.

        Returns:
            XDrawPages: Draw Pages.
        """
        try:
            supplier = mLo.Lo.qi(XDrawPagesSupplier, doc, True)
            pages = supplier.getDrawPages()
            if pages is None:
                raise mEx.DrawPageMissingError("Draw page supplier returned no pages")
            return pages
        except mEx.DrawPageMissingError:
            raise
        except Exception as e:
            raise mEx.DrawPageError("Error getting slides") from e

    @classmethod
    def get_slides_count(cls, doc: XComponent) -> int:
        """
        Gets the slides count.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Returns:
            int: Number of slides.
        """
        slides = cls.get_slides(doc)
        return 0 if slides is None else slides.getCount()

    @classmethod
    def get_slides_list(cls, doc: XComponent) -> List[XDrawPage]:
        """
        Gets all the slides as a list of XDrawPage.

        |lo_safe|

        Args:
            doc (XComponent): Document

        Returns:
            List[XDrawPage]: List of pages
        """
        slides = cls.get_slides(doc)
        if slides is None:
            return []
        num_slides = slides.getCount()
        results: List[XDrawPage] = [mLo.Lo.qi(XDrawPage, slides.getByIndex(i), True) for i in range(num_slides)]
        return results

    # region get_slide()

    @classmethod
    def _get_slide_doc(cls, doc: XComponent, idx: int) -> XDrawPage:
        """Lo Safe Method."""
        return cls._get_slide_slides(cls.get_slides(doc), idx)

    @staticmethod
    def _get_slide_slides(slides: XDrawPages, idx: int) -> XDrawPage:
        """Lo Safe Method."""
        try:
            return mLo.Lo.qi(XDrawPage, slides.getByIndex(idx), True)
        except IndexOutOfBoundsException as e:
            raise IndexError(f"Index out of bounds: {idx}") from e
        except Exception as e:
            raise mEx.DrawError(f"Could not get slide: {idx}") from e

    @overload
    @classmethod
    def get_slide(cls, doc: XComponent) -> XDrawPage:
        """
        Gets slide at index ``0``.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Returns:
            XDrawPage: Slide as Draw Page.
        """
        ...

    @overload
    @classmethod
    def get_slide(cls, doc: XComponent, idx: int) -> XDrawPage:
        """
        Gets slide by page index.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            idx (int): Index of slide. Default ``0``.

        Returns:
            XDrawPage: Slide as Draw Page.
        """
        ...

    @overload
    @classmethod
    def get_slide(cls, slides: XDrawPages) -> XDrawPage:
        """
        Gets slide at index ``0``.

        |lo_safe|

        Args:
            slides (XDrawPages): Draw Pages.

        Returns:
            XDrawPage: Slide as Draw Page.
        """
        ...

    @overload
    @classmethod
    def get_slide(cls, slides: XDrawPages, idx: int) -> XDrawPage:
        """
        Gets slide by page index.

        |lo_safe|

        Args:
            slides (XDrawPages): Draw Pages.
            idx (int): Index of slide. Default ``0``.

        Returns:
            XDrawPage: Slide as Draw Page.
        """
        ...

    @classmethod
    def get_slide(cls, *args, **kwargs) -> XDrawPage:
        """
        Gets slide by page index.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            slides (XDrawPages): Draw Pages.
            idx (int): Index of slide. Default ``0``.

        Raises:
            IndexError: If ``idx`` is out of bounds.
            DrawError: If any other error occurs.

        Returns:
            XDrawPage: Slide as Draw Page.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "slides", "idx")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_slide() got an unexpected keyword argument")
            keys = ("doc", "slides")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("idx", None)
            return ka

        if count not in (1, 2):
            raise TypeError("get_slide() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        idx = 0 if count == 1 else cast(int, kargs[2])
        if mLo.Lo.is_uno_interfaces(kargs[1], XDrawPages):
            return cls._get_slide_slides(kargs[1], idx)
        return cls._get_slide_doc(kargs[1], idx)

    # endregion get_slide()

    @classmethod
    def find_slide_idx_by_name(cls, doc: XComponent, name: str) -> int:
        """
        Gets a slides index by its name.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            name (str): Slide Name.

        Returns:
            int: Zero based index if found; Otherwise ``-1``.
        """
        slide_name = name.casefold()
        num_slides = cls.get_slides_count(doc)
        for i in range(num_slides):
            slide = cls._get_slide_doc(doc, i)
            nm = str(mProps.Props.get(slide, "LinkDisplayName")).casefold()
            if nm == slide_name:
                return i
        return -1

    # region get_shapes()

    @classmethod
    def _get_shapes_doc(cls, doc: XComponent) -> List[XShape]:
        """LO Safe Method."""
        slides = cls.get_slides_list(doc)
        if not slides:
            return []
        shapes: List[XShape] = []
        for slide in slides:
            shapes.extend(cls._get_shapes_slide(slide))
        return shapes

    @classmethod
    def _get_shapes_slide(cls, slide: XDrawPage) -> List[XShape]:
        """LO Safe Method."""
        if slide.getCount() == 0:
            mLo.Lo.print("Slide does not contain any shapes")
            return []

        shapes: List[XShape] = []
        for i in range(slide.getCount()):
            try:
                shapes.append(mLo.Lo.qi(XShape, slide.getByIndex(i), True))
            except Exception as e:
                cargs = CancelEventArgs(Draw.get_shapes.__qualname__)
                cargs.event_data = {"index": i, "raise_error": False}
                mLo.Lo.print("Shapes extraction error in slide")
                mLo.Lo.print(f"  {e}")
                mLo.Lo.print(f"Raising Event: {DrawNamedEvent.GET_SHAPES_ERROR}")
                _Events().trigger(DrawNamedEvent.GET_SHAPES_ERROR, cargs)
                if cargs.event_data.get("raise_error", False):
                    raise mEx.ShapeError(f"Error getting slide shape for index: {i}") from e
                if not cargs.cancel:
                    continue
                mLo.Lo.print("Breaking from getting shapes due to event cancel")
                break
        return shapes

    @overload
    @classmethod
    def get_shapes(cls, doc: XComponent) -> List[XShape]:
        """
        Gets Shapes for drawpage at index ``0``.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Returns:
            List[XShape]: List of shapes.
        """
        ...

    @overload
    @classmethod
    def get_shapes(cls, slide: XDrawPage) -> List[XShape]:
        """
        Gets Shapes for specified drawpage.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            slide (XDrawPage): Slide.

        Returns:
            List[XShape]: List of shapes.
        """
        ...

    @classmethod
    def get_shapes(cls, *args, **kwargs) -> List[XShape]:
        """
        Gets Shapes.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            slide (XDrawPage): Slide.

        Raises:
            ShapeError: Conditionally if event :py:attr:`.DrawNamedEvent.GET_SHAPES_ERROR` is subscribe to. See Note.

        Returns:
            List[XShape]: List of shapes.

        Note:
            By default ``get_shapes`` will ignore shapes that fail to load.
            This behavior can be overridden by subscribing to :py:attr:`.DrawNamedEvent.GET_SHAPES_ERROR`
            event.

            The event is called with :py:class:`~.cancel_event_args.CancelEventArgs`.

            If ``event.cancel`` is ``True`` then a list of shapes is returned containing all shapes
            that were added before first error occurred.

            If ``Event.event_data["raise_error] = True`` then a ``ShapeError`` is raised.

            .. code-block:: python

                from ooodev.events.args.cancel_event_args import CancelEventArgs
                from ooodev.events.lo_events import LoEvents
                from ooodev.events.draw_named_event import DrawNamedEvent
                from ooodev.exceptions.ex import ShapeError

                def on_shapes_error(source: Any, e: CancelEventArgs) -> None:
                    e.event_data["raise_error] = True

                LoEvents().on(DrawNamedEvent.GET_SHAPES_ERROR, on_shapes_error)

                def do_something(doc: XComponent) -> None:
                    try:
                        shapes = Draw.get_shapes(doc)
                    except ShapeError:
                        # this error only occurred because event raise_error was set in on_shapes_error
                        # handle error
                        ...
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "slide")
            check = all(key in valid_keys for key in kwargs)
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
        """Lo Safe Method."""
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
        """Lo Safe Method."""

        def sorter(obj: XShape) -> int:
            return cls.get_zorder(obj)

        shapes = cls._get_shapes_slide(slide)
        sorted_shapes = sorted(shapes, key=sorter, reverse=False)
        return sorted_shapes

    @overload
    @classmethod
    def get_ordered_shapes(cls, doc: XComponent) -> List[XShape]:
        """
        Gets ordered shapes for drawpage at index ``0``.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Returns:
            List[XShape]: List of Ordered Shapes.
        """
        ...

    @overload
    @classmethod
    def get_ordered_shapes(cls, slide: XDrawPage) -> List[XShape]:
        """
        Gets ordered shapes for specified drawpage.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.

        Returns:
            List[XShape]: List of Ordered Shapes.
        """
        ...

    @classmethod
    def get_ordered_shapes(cls, *args, **kwargs) -> List[XShape]:
        """
        Gets ordered shapes.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            slide (XDrawPage): Slide.

        Returns:
            List[XShape]: List of Ordered Shapes.

        See Also:
            :py:meth:`~.draw.Draw.get_shapes`
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "slide")
            check = all(key in valid_keys for key in kwargs)
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
        Gets the text from inside all the document shapes.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Returns:
            str: Shapes text.

        See Also:
            - :py:meth:`~.draw.Draw.get_shapes`
            - :py:meth:`~.draw.Draw.get_ordered_shapes`
        """
        sb: List[str] = []
        shapes = cls._get_ordered_shapes_doc(doc)
        for shape in shapes:
            text = cls._get_shape_text_shape(shape)
            sb.append(text)
        return "\n".join(sb)

    @classmethod
    def add_slide(cls, doc: XComponent) -> XDrawPage:
        """
        Add a slide to the end of the document.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            XDrawPage: The slide that was inserted at the end of the document.
        """
        try:
            mLo.Lo.print("Adding a slide")
            slides = cls.get_slides(doc)
            num_slides = slides.getCount()
            return slides.insertNewByIndex(num_slides)
        except mEx.DrawPageMissingError:
            raise
        except mEx.DrawPageError:
            raise
        except Exception as e:
            raise mEx.DrawPageError("Error adding slide to document") from e

    @classmethod
    def insert_slide(cls, doc: XComponent, idx: int) -> XDrawPage:
        """
        Inserts a slide at the given position in the document.

        |lo_safe|

        Args:
            doc (XComponent): Document
            idx (int): Index

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            XDrawPage: New slide that was inserted.
        """
        try:
            mLo.Lo.print(f"Inserting a slide at position: {idx}")
            slides = cls.get_slides(doc)
            return slides.insertNewByIndex(idx)
        except mEx.DrawPageMissingError:
            raise
        except mEx.DrawPageError:
            raise
        except Exception as e:
            raise mEx.DrawPageError("Error inserting slide in document") from e

    @classmethod
    def delete_slide(cls, doc: XComponent, idx: int) -> bool:
        """
        Deletes a slide.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            idx (int): Index.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        mLo.Lo.print(f"Deleting a slide as position: {idx}")
        slides = cls.get_slides(doc)
        slide = None
        try:
            slide = mLo.Lo.qi(XDrawPage, slides.getByIndex(idx), True)
        except Exception as e:
            mLo.Lo.print(f"Error: Could not delete slide: {idx}: {e}")
            return False
        slides.remove(slide)
        return True

    @classmethod
    def duplicate(cls, doc: XComponent, idx: int) -> XDrawPage:
        """
        Duplicates a slide.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            idx (int): Index of slide to duplicate.

        Raises:
            DrawError If unable to create duplicate.

        Returns:
            XDrawPage: Duplicated slide.
        """
        dup = mLo.Lo.qi(XDrawPageDuplicator, doc, True)
        from_slide = cls._get_slide_doc(doc, idx)
        # places copy after original
        duplicate = dup.duplicate(from_slide)
        if duplicate is None:
            raise mEx.DrawError("Unable to create duplicate")
        return duplicate

    # endregion methods related to document/multiple slides/pages

    # region Layer Management
    @staticmethod
    def get_layer_manager(doc: XComponent) -> XLayerManager:
        """
        Gets Layer manager for document.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawError: If error occurs.

        Returns:
            XLayerManager: Layer Manager.
        """
        try:
            layer_supp = mLo.Lo.qi(XLayerSupplier, doc, True)
            name_acc = layer_supp.getLayerManager()
            return mLo.Lo.qi(XLayerManager, name_acc, True)
        except Exception as e:
            raise mEx.DrawError("Error getting XLayerManager") from e

    @staticmethod
    def get_layer(doc: XComponent, layer_name: DrawingLayerKind | str) -> XLayer:
        """
        Gets layer from layer name.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            layer_name (str): Layer Name.

        Raises:
            NameError: If ``layer_name`` does not exist.
            DrawError: If unable to get layer.

        Returns:
            XLayer: Found Layer
        """
        layer_supplier = mLo.Lo.qi(XLayerSupplier, doc, True)
        name_access = layer_supplier.getLayerManager()
        try:
            return mLo.Lo.qi(XLayer, name_access.getByName(str(layer_name)), True)
        except NoSuchElementException as e:
            raise NameError(f'"{layer_name}" does not exist') from e
        except Exception as e:
            raise mEx.DrawError(f'Could not find the layer "{layer_name}"') from e

    @staticmethod
    def add_layer(lm: XLayerManager, layer_name: str) -> XLayer:
        """
        Adds a layer.

        |lo_safe|

        Args:
            lm (XLayerManager): Layer Manager.
            layer_name (str): Layer Name.

        Raises:
            DrawError: If error occurs.

        Returns:
            XLayer: Newly added layer.
        """
        layer = None
        try:
            layer = lm.insertNewByIndex(lm.getCount())
            props = mLo.Lo.qi(XPropertySet, layer, True)
            props.setPropertyValue("Name", layer_name)
            props.setPropertyValue("IsVisible", True)
            props.setPropertyValue("IsLocked", False)
        except Exception as e:
            raise mEx.DrawError(f'Could not add the layer "{layer_name}"') from e
        return layer

    # endregion Layer Management

    # region view page

    # region goto_page()
    @classmethod
    def _goto_page_doc(cls, doc: XComponent, page: XDrawPage) -> None:
        """Lo Safe Method."""
        try:
            ctl = mGui.GUI.get_current_controller(doc)
            cls._goto_page_ctl(ctl, page)
        except mEx.DrawError:
            raise
        except Exception as e:
            raise mEx.DrawError("Error while trying to go to page") from e

    @staticmethod
    def _goto_page_ctl(ctl: XController, page: XDrawPage) -> None:
        """LO Safe Method."""
        try:
            x_draw_view = mLo.Lo.qi(XDrawView, ctl, True)
            x_draw_view.setCurrentPage(page)
        except Exception as e:
            raise mEx.DrawError("Error while trying to go to page") from e

    @overload
    @classmethod
    def goto_page(cls, doc: XComponent, page: XDrawPage) -> None:
        """
        Go to page.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            page (XDrawPage): Page.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def goto_page(cls, ctl: XController, page: XDrawPage) -> None:
        """
        Go to page.

        |lo_safe|

        Args:
            ctl (XController): Controller.
            page (XDrawPage): Page.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        ...

    @classmethod
    def goto_page(cls, *args, **kwargs) -> None:
        """
        Go to page.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            ctl (XController): Controller.
            page (XDrawPage): Page.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "ctl", "page")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("goto_page() got an unexpected keyword argument")
            keys = ("doc", "ctl")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            ka[2] = kwargs.get("page", None)
            return ka

        if count != 2:
            raise TypeError("goto_page() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        obj = kargs[1]
        ctl = mLo.Lo.qi(XController, obj)
        if ctl is None:
            cls._goto_page_doc(obj, kargs[2])
        else:
            cls._goto_page_ctl(ctl, kargs[2])

    # endregion goto_page()

    @staticmethod
    def get_viewed_page(doc: XComponent) -> XDrawPage:
        """
        Gets viewed page.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawPageError: If error occurs.

        Returns:
            XDrawPage: Draw Page.
        """
        try:
            ctl = mGui.GUI.get_current_controller(doc)
            draw_view = mLo.Lo.qi(XDrawView, ctl, True)
            return draw_view.getCurrentPage()
        except Exception as e:
            raise mEx.DrawPageError("Error getting Viewed page") from e

    # region get_slide_number()

    @classmethod
    def _get_slide_number_draw_view(cls, xdraw_view: XDrawView) -> int:
        """
        Gets Draw view slide number.

        |lo_safe|

        Args:
            xdraw_view (XDrawView): Draw View.

        Raises:
            DrawError: If error occurs.

        Returns:
            int: page number.
        """
        try:
            curr_page = xdraw_view.getCurrentPage()
            return cls._get_slide_number_draw_page(curr_page)
        except mEx.DrawError:
            raise
        except Exception as e:
            raise mEx.DrawError("Error getting slide number") from e

    @staticmethod
    def _get_slide_number_draw_page(slide: XDrawPage) -> int:
        """
        Gets slide page number.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawError: If error occurs.

        Returns:
            int: Page number
        """
        try:
            return int(mProps.Props.get(slide, "Number"))
        except Exception as e:
            raise mEx.DrawError("Error getting slide number") from e

    @overload
    @classmethod
    def get_slide_number(cls, xdraw_view: XDrawView) -> int:
        """
        Gets slide number.

        |lo_safe|

        Args:
            xdraw_view (XDrawView): Draw View.

        Returns:
            int: Slide Number.
        """
        ...

    @overload
    @classmethod
    def get_slide_number(cls, slide: XDrawPage) -> int:
        """
        Gets slide number.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.

        Returns:
            int: Slide Number.
        """
        ...

    @classmethod
    def get_slide_number(cls, *args, **kwargs) -> int:
        """
        Gets slide number.

        |lo_safe|

        Args:
            xdraw_view (XDrawView): Draw View.
            slide (XDrawPage): Slide.

        Raises:
            DrawError: If error occurs.

        Returns:
            int: Slide Number.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("xdraw_view", "slide")
            check = all(key in valid_keys for key in kwargs)
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
        Gets master page count.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawError: If error occurs.

        Returns:
            int: Master Page Count.
        """
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
            pgs = mp_supp.getMasterPages()
            return pgs.getCount()
        except Exception as e:
            raise mEx.DrawError("Error getting master page count") from e

    # region get_master_page()
    @staticmethod
    def _get_master_page_idx(doc: XComponent, idx: int) -> XDrawPage:
        """Lo Safe Method."""
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
            pgs = mp_supp.getMasterPages()
            return mLo.Lo.qi(XDrawPage, pgs.getByIndex(idx), True)
        except IndexOutOfBoundsException as e:
            raise IndexError(f'Index "{idx}" is out of range') from e
        except Exception as e:
            raise mEx.DrawPageError(f"Could not find master slide for index: {idx}") from e

    @staticmethod
    def _get_master_page_slide(slide: XDrawPage) -> XDrawPage:
        """LO Safe Method."""
        try:
            mp_target = mLo.Lo.qi(XMasterPageTarget, slide, True)
            return mp_target.getMasterPage()
        except Exception as e:
            raise mEx.DrawPageError("Unable to get master page") from e

    @overload
    @classmethod
    def get_master_page(cls, doc: XComponent, idx: int) -> XDrawPage:
        """
        Gets master page.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            idx (int): Index of page.

        Returns:
            XDrawPage: Draw page.
        """
        ...

    @overload
    @classmethod
    def get_master_page(cls, slide: XDrawPage) -> XDrawPage:
        """
        Gets master page.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide to get master page from.

        Raises:
            DrawPageError: if unable to get master page.
            IndexError: if ``idx`` is out of bounds.

        Returns:
            XDrawPage: Draw page.
        """
        ...

    @classmethod
    def get_master_page(cls, *args, **kwargs) -> XDrawPage:
        """
        Gets master page.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            idx (int): Index of page.
            slide (XDrawPage): Slide to get master page from.

        Raises:
            DrawPageError: if unable to get master page.
            IndexError: if ``idx`` is out of bounds.

        Returns:
            XDrawPage: Draw page.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "idx", "slide")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_master_page() got an unexpected keyword argument")
            keys = ("doc", "slide")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("idx", None)
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
        Inserts a master page.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            idx (int): Index used to insert page.

        Raises:
            DrawPageError: If unable to insert master page.

        Returns:
            XDrawPage: The newly inserted draw page.
        """
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
            pgs = mp_supp.getMasterPages()
            result = pgs.insertNewByIndex(idx)
            if result is None:
                raise mEx.NoneError("None Value: insertNewByIndex() return None")
            return result
        except Exception as e:
            raise mEx.DrawPageError("Unable to insert master page") from e

    @staticmethod
    def remove_master_page(doc: XComponent, slide: XDrawPage) -> None:
        """
        Removes a master page.

        |lo_safe|

        Args:
            doc (XComponent): Document.
            slide (XDrawPage): Draw page to remove.

        Raises:
            DrawError: If unable to remove master page.

        Returns:
            None:
        """
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
            pgs = mp_supp.getMasterPages()
            pgs.remove(slide)
        except Exception as e:
            raise mEx.DrawError("Unable to remove master page") from e

    @staticmethod
    def set_master_page(slide: XDrawPage, page: XDrawPage) -> None:
        """
        Sets master page.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            page (XDrawPage): Page to set as master.

        Raises:
            DrawError: If unable to set master page.

        Returns:
            None:
        """
        try:
            mp_target = mLo.Lo.qi(XMasterPageTarget, slide, True)
            mp_target.setMasterPage(page)
        except Exception as e:
            raise mEx.DrawError("Unable to set master page") from e

    @staticmethod
    def get_handout_master_page(doc: XComponent) -> XDrawPage:
        """
        Gets handout master page for an impress document.

        |lo_safe|

        Args:
            doc (XComponent): Impress Document.

        Raises:
            DrawError: If unable to get hand-out master page.
            DrawPageMissingError: If Draw Page is ``None``.

        Returns:
            XDrawPage: Draw Page
        """
        try:
            hm_supp = mLo.Lo.qi(XHandoutMasterSupplier, doc, True)
            result = hm_supp.getHandoutMasterPage()
        except Exception as e:
            raise mEx.DrawError("Unable to get hand-out master page") from e
        if result is None:
            raise mEx.DrawPageMissingError("Unable to get hand-out master page")
        return result

    @staticmethod
    def find_master_page(doc: XComponent, style: str) -> XDrawPage:
        """
        Finds master page.

        Args:
            doc (XComponent): Document.
            style (str): Style of master page.

        |lo_safe|

        Raises:
            DrawPageMissingError: If unable to match ``style``.
            DrawPageError: if any other error occurs.

        Returns:
            XDrawPage: Master page as Draw Page if found.
        """
        try:
            mp_supp = mLo.Lo.qi(XMasterPagesSupplier, doc, True)
            master_pgs = mp_supp.getMasterPages()

            for i in range(master_pgs.getCount()):
                pg = mLo.Lo.qi(XDrawPage, master_pgs.getByIndex(i), True)
                nm = str(mProps.Props.get(pg, "LinkDisplayName"))
                if style == nm:
                    return pg
            raise mEx.DrawPageMissingError(f'Could not find master slide with style of: "{style}"')
        except mEx.DrawPageMissingError:
            raise
        except Exception as e:
            raise mEx.DrawPageError("Could not find master slide") from e

    # endregion master page methods

    # region slide/page methods
    @classmethod
    def show_shapes_info(cls, slide: XDrawPage) -> None:
        """
        Prints info for shapes to console.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide/

        Returns:
            None:
        """
        print("Draw Page shapes:")
        shapes = cls._get_shapes_slide(slide)
        for shape in shapes:
            cls.show_shape_info(shape)

    @classmethod
    def get_slide_title(cls, slide: XDrawPage) -> str | None:
        """
        Gets slide title if it exist.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawError: If error occurs.

        Returns:
            str | None: Slide Title on success; Otherwise, ``None``.
        """
        try:
            shape = cls.find_shape_by_type(slide=slide, shape_type=str(DrawingNameSpaceKind.TITLE_TEXT))
            return None if shape is None else cls._get_shape_text_shape(shape)
        except mEx.DrawError:
            raise
        except Exception as e:
            raise mEx.DrawError("Error getting slide title") from e

    @staticmethod
    def get_slide_size(slide: XDrawPage) -> Size:
        """
        Gets size of the given slide page (in mm units).

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            SizeError: If unable to get size.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size struct.
        """
        try:
            props = mLo.Lo.qi(XPropertySet, slide, True)
            width = int(props.getPropertyValue("Width"))  # type: ignore
            height = int(props.getPropertyValue("Height"))  # type: ignore
            width_mm = UnitMM.from_mm100(width).value
            height_mm = UnitMM.from_mm100(height).value
            return Size(round(width_mm), round(height_mm))
        except Exception as e:
            raise mEx.SizeError("Could not get shape size") from e

    @staticmethod
    def get_name(slide: XDrawPage) -> str:
        """
        Gets the name of a slide.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.

        Raises:
            DrawError: If error occurs getting name.

        Returns:
            str: Slide name.

        .. versionadded:: 0.17.13
        """
        try:
            page_name = mLo.Lo.qi(XNamed, slide, True)
            return page_name.getName()
        except Exception as e:
            raise mEx.DrawError("Unable to get Name") from e

    @staticmethod
    def set_name(slide: XDrawPage, name: str) -> None:
        """
        Sets the name of a slide.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            name (str): Name.

        Raises:
            DrawError: If error occurs setting name.

        Returns:
            None:
        """
        try:
            page_name = mLo.Lo.qi(XNamed, slide, True)
            page_name.setName(name)
        except Exception as e:
            raise mEx.DrawError("Unable to set Name") from e

    # region title_slide()
    @overload
    @classmethod
    def title_slide(cls, slide: XDrawPage, title: str) -> None:
        """
        Set a slides title and sub title.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            title (str): Title.

        Raises:
            DrawError: If error setting Slide.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def title_slide(cls, slide: XDrawPage, title: str, sub_title: str) -> None:
        """
        Set a slides title and sub title.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            title (str): Title.
            sub_title (str): Sub Title.

        Raises:
            DrawError: If error setting Slide.

        Returns:
            None:
        """
        ...

    @classmethod
    def title_slide(cls, slide: XDrawPage, title: str, sub_title: str = "") -> None:
        """
        Set a slides title and sub title.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            title (str): Title.
            sub_title (str): Sub Title.

        Raises:
            DrawError: If error setting Slide.

        Returns:
            None:
        """
        try:
            # Add text to the slide page by treating it as a title page, which
            # has two text shapes: one for the title, the other for a subtitle
            mProps.Props.set(slide, Layout=PresentationLayoutKind.TITLE_SUB.value)

            # add the title text to the title shape
            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.TITLE_TEXT)
            txt_field = mLo.Lo.qi(XText, xs, True)
            txt_field.setString(title)

            # add the subtitle text to the subtitle shape
            if sub_title:
                xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.SUBTITLE_TEXT)
                txt_field = mLo.Lo.qi(XText, xs, True)
                txt_field.setString(sub_title)
        except Exception as e:
            raise mEx.DrawError("Error setting Slide") from e

    # endregion title_slide()

    @classmethod
    def bullets_slide(cls, slide: XDrawPage, title: str) -> XText:
        """
        Add text to the slide page by treating it as a bullet page, which
        has two text shapes: one for the title, the other for a sequence of
        bullet points; add the title text but return a reference to the bullet
        text area.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide
            title (str): Title

        Raises:
            DrawError: If error setting slide.

        Returns:
            XText: Text Object.
        """
        try:
            mProps.Props.set(slide, Layout=PresentationLayoutKind.TITLE_BULLETS.value)

            # add the title text to the title shape
            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.TITLE_TEXT)
            txt_field = mLo.Lo.qi(XText, xs, True)
            txt_field.setString(title)

            # return a reference to the bullet text area
            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.BULLETS_TEXT)
            return mLo.Lo.qi(XText, xs, True)
        except Exception as e:
            raise mEx.DrawError("Error occurred setting bullets slide") from e

    @staticmethod
    def add_bullet(bulls_txt: XText, level: int, text: str) -> None:
        """
        Add bullet text to the end of the bullets text area, specifying
        the nesting of the bullet using a numbering level value
        (numbering starts at 0).

        |lo_safe|

        Args:
            bulls_txt (XText): Text object
            level (int): Bullet Level
            text (str): Bullet Text

        Raises:
            DrawError: If error adding bullet.

        Returns:
            None:
        """
        try:
            # access the end of the bullets text
            bulls_txt_end = mLo.Lo.qi(XTextRange, bulls_txt, True).getEnd()

            # set the bullet's level
            mProps.Props.set(bulls_txt_end, NumberingLevel=level)

            # add the text
            bulls_txt_end.setString(f"{text}\n")
        except Exception as e:
            raise mEx.DrawError("Error adding bullet") from e

    @classmethod
    def title_only_slide(cls, slide: XDrawPage, header: str) -> None:
        """
        Creates a slide with only a title.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide
            header (str): Header text.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        try:
            mProps.Props.set(slide, Layout=PresentationLayoutKind.TITLE_ONLY.value)

            # add the text to the title shape
            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.TITLE_TEXT)
            txt_field = mLo.Lo.qi(XText, xs, True)
            txt_field.setString(header)
        except Exception as e:
            raise mEx.DrawError("Error creating title only slide") from e

    @staticmethod
    def blank_slide(slide: XDrawPage) -> None:
        """
        Inserts a blank slide.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        try:
            mProps.Props.set(slide, Layout=PresentationLayoutKind.BLANK.value)
        except Exception as e:
            raise mEx.DrawError("Error inserting blank slide") from e

    @staticmethod
    def get_notes_page(slide: XDrawPage) -> XDrawPage:
        """
        Gets the notes page of a slide.

        Each draw page has a notes page.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawPageMissingError: If notes page is ``None``.
            DrawPageError: If any other error occurs.

        Returns:
            XDrawPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page_by_index`
        """
        try:
            pres_page = mLo.Lo.qi(XPresentationPage, slide)
            if pres_page is None:
                raise mEx.MissingInterfaceError(
                    XPresentationPage, "This is not a presentation slide, so no notes page is available"
                )

            result = pres_page.getNotesPage()
            if result is None:
                raise mEx.DrawPageMissingError("Unable to get notes page")
            return result
        except mEx.DrawPageMissingError:
            raise
        except Exception as e:
            raise mEx.DrawPageError("An error occurred getting notes page") from e

    @classmethod
    def get_notes_page_by_index(cls, doc: XComponent, idx: int) -> XDrawPage:
        """
        Gets notes page by index.

        Each draw page has a notes page.

        |lo_safe|

        Args:
            doc (XComponent): Document
            idx (int): Index

        Raises:
            DrawPageError: If error occurs.

        Returns:
            XDrawPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page`
        """
        try:
            slide = cls._get_slide_doc(doc, idx)
            return cls.get_notes_page(slide)
        except mEx.DrawPageError:
            raise
        except Exception as e:
            raise mEx.DrawPageError(f"Error get notes page from index: {idx}") from e

    # endregion slide/page methods

    # region shape methods
    @classmethod
    def show_shape_info(cls, shape: XShape) -> None:
        """
        Prints shape info to console.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Returns:
            None:
        """
        print(f"  Shape service: {shape.getShapeType()}; z-order: {cls.get_zorder(shape)}")

    # region get_shape_text()
    @classmethod
    def _get_shape_text_shape(cls, shape: XShape) -> str:
        """Lo Safe Method."""
        try:
            xtext = mLo.Lo.qi(XText, shape, True)

            xtext_cursor = xtext.createTextCursor()
            xtext_rng = mLo.Lo.qi(XTextRange, xtext_cursor, True)
            return xtext_rng.getString()
        except Exception as e:
            raise mEx.DrawError("Error getting shape text from shape") from e

    @classmethod
    def _get_shape_text_slide(cls, slide: XDrawPage) -> str:
        """Lo Safe Method."""
        try:
            sb: List[str] = []
            shapes = cls._get_ordered_shapes_slide(slide)
            for shape in shapes:
                text = cls._get_shape_text_shape(shape)
                sb.append(text)
            return "\n".join(sb)
        except Exception as e:
            raise mEx.DrawError("Error getting shape text from slide") from e

    @overload
    @classmethod
    def get_shape_text(cls, shape: XShape) -> str: ...

    @overload
    @classmethod
    def get_shape_text(cls, slide: XDrawPage) -> str: ...

    @classmethod
    def get_shape_text(cls, *args, **kwargs) -> str:
        """
        Gets the text from inside a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            slide (XDrawPage): Slide.

        Raises:
            DrawError: If error occurs getting shape text.

        Returns:
            str: Shape text.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "slide")
            check = all(key in valid_keys for key in kwargs)
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
    def find_shape_by_type(cls, slide: XDrawPage, shape_type: DrawingNameSpaceKind | str) -> XShape:
        """
        Finds a shape by its type.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide
            shape_type (DrawingNameSpaceKind | str): Shape Type

        Raise:
            ShapeMissingError: If shape is not found.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape
        """
        try:
            shapes = cls.get_shapes(slide)
            if not shapes:
                raise mEx.ShapeMissingError("No shapes were found in the draw page")

            st = str(shape_type)

            for shape in shapes:
                if st == shape.getShapeType():
                    return shape
            raise mEx.ShapeMissingError(f'No shape found for "{st}"')
        except mEx.ShapeMissingError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error occurred while looking for shape") from e

    @classmethod
    def find_shape_by_name(cls, slide: XDrawPage, shape_name: str) -> XShape:
        """
        Finds a shape by its name.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            shape_name (str): Shape Name.

        Raise:
            ShapeMissingError: If shape is not found.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape.
        """
        try:
            shapes = cls.get_shapes(slide)
            sn = shape_name.casefold()
            if not shapes:
                raise mEx.ShapeMissingError("No shapes were found in the draw page")

            for shape in shapes:
                nm = str(mProps.Props.get(shape, "Name")).casefold()
                if nm == sn:
                    return shape

            raise mEx.ShapeMissingError(f'No shape named "{shape_name}"')
        except mEx.ShapeMissingError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error occurred while looking for shape") from e

    @classmethod
    def copy_shape_contents(cls, slide: XDrawPage, old_shape: XShape) -> XShape:
        """
        Copies a shapes contents from old shape into new shape.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            old_shape (XShape): Old shape.

        Raises:
            ShapeError: If unable to copy shape contents.

        Returns:
            XShape: New shape with contents of old shape copied.
        """
        try:
            shape = cls.copy_shape(slide, old_shape)
            mLo.Lo.print(f'Shape type: "{old_shape.getShapeType()}"')
            cls.add_text(shape, cls.get_shape_text(old_shape))
            return shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to copy shape contents") from e

    @classmethod
    def copy_shape(cls, slide: XDrawPage, old_shape: XShape) -> XShape:
        """
        Copies a shape.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide
            old_shape (XShape): Old Shape

        Raises:
            ShapeError: If unable to copy shape.

        Returns:
            XShape: Newly Copied shape.
        """
        try:
            # parameters are in 1/100 mm units
            pt = old_shape.getPosition()
            sz = old_shape.getSize()
            shape = mLo.Lo.create_instance_msf(XShape, old_shape.getShapeType(), raise_err=True)
            mLo.Lo.print(f"Copying: {old_shape.getShapeType()}")
            shape.setPosition(pt)
            shape.setSize(sz)
            slide.add(shape)
            return shape
        except Exception as ex:
            raise mEx.ShapeError("Unable to copy Shape") from ex

    @staticmethod
    def set_zorder(shape: XShape, order: int) -> None:
        """
        Sets the z-order of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            order (int): Z-Order.

        Raises:
            DrawError: If unable to set z-order.

        Returns:
            None:
        """
        try:
            mProps.Props.set(shape, ZOrder=order)
        except Exception as e:
            raise mEx.DrawError("Failed to set z-order") from e

    @staticmethod
    def get_zorder(shape: XShape) -> int:
        """
        Gets the z-order of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            DrawError: If unable to get z-order.

        Returns:
            int: Z-Order
        """
        try:
            return int(mProps.Props.get(shape, "ZOrder"))
        except Exception as e:
            raise mEx.DrawError("Failed to get z-order") from e

    @classmethod
    def move_to_top(cls, slide: XDrawPage, shape: XShape) -> None:
        """
        Moves the z-order of a shape to the top.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide
            shape (XShape): Shape

        Raises:
            DrawError: If unable to move shape to top.

        Returns:
            None:
        """
        try:
            max_zo = cls.find_biggest_zorder(slide)
            cls.set_zorder(shape, max_zo + 1)
        except mEx.DrawError:
            raise
        except Exception as e:
            raise mEx.DrawError("Error moving shape to top.") from e

    @classmethod
    def find_biggest_zorder(cls, slide: XDrawPage) -> int:
        """
        Finds the shape with the largest z-order.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawError: If unable to find biggest z-order.

        Returns:
            int: Z-Order
        """
        try:
            return cls.get_zorder(cls.find_top_shape(slide))
        except mEx.DrawError:
            raise
        except Exception as e:
            raise mEx.DrawError("Error finding biggest z-order") from e

    @classmethod
    def find_top_shape(cls, slide: XDrawPage) -> XShape:
        """
        Gets the top most shape of a slide.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            ShapeMissingError: If there are no shapes for slide or unable to find top shape.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Top most shape.
        """
        try:
            shapes = cls.get_shapes(slide)
            if not shapes:
                raise mEx.ShapeMissingError("No shapes found")
            max_zorder = -1
            top = None
            for shape in shapes:
                zo = cls.get_zorder(shape)
                if zo > max_zorder:
                    max_zorder = zo
                    top = shape
            if top is None:
                raise mEx.ShapeMissingError("Top shape not found")
            return top
        except mEx.ShapeMissingError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error finding top shape") from e

    @classmethod
    def move_to_bottom(cls, slide: XDrawPage, shape: XShape) -> None:
        """
        Moves a shape to the bottom of the z-order.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.
            shape (XShape): Shape.

        Raises:
            ShapeMissingError: If unable to find shapes for slide.
            ShapeError: If any other error occurs.

        Returns:
            None:
        """
        try:
            shapes = cls.get_shapes(slide)
            if not shapes:
                raise mEx.ShapeMissingError("No shapes found")

            min_zorder = 999
            for sh in shapes:
                zo = cls.get_zorder(sh)
                if zo < min_zorder:
                    min_zorder = zo
                cls.set_zorder(sh, zo + 1)
            cls.set_zorder(shape, min_zorder)
        except mEx.ShapeMissingError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error moving shape to bottom") from e

    # endregion shape methods

    # region draw/add shape to a page
    @classmethod
    def make_shape(
        cls,
        shape_type: DrawingShapeKind | str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
    ) -> XShape:
        """
        Creates a shape.

        |lo_unsafe|

        Args:
            shape_type (DrawingShapeKind | str): Shape type.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: New Shape

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        See Also:
            :py:meth:`~.draw.Draw.add_shape`

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        try:
            x = cls._get_mm100_obj_from_mm(x).value
            y = cls._get_mm100_obj_from_mm(y).value
            width = cls._get_mm100_obj_from_mm(width).value
            height = cls._get_mm100_obj_from_mm(height).value
            shape = mLo.Lo.create_instance_msf(XShape, f"com.sun.star.drawing.{shape_type}", raise_err=True)
            if x > 0 and y > 0:
                shape.setPosition(Point(x, y))
            if width > 0 and height > 0:
                shape.setSize(UnoSize(width, height))
            return shape
        except Exception as e:
            raise mEx.ShapeError(f'Unable to create shape "{shape_type}"') from e

    @classmethod
    def warns_position(cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT) -> None:
        """
        Warns via console if a ``x`` or ``y`` is not on the page.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide
            x (int, UnitT): X Position in mm units or UnitT
            y (int, UnitT): Y Position int mm units or UnitT

        Returns:
            None:

        Note:
            This method uses :py:meth:`.Lo.print`. and if those ``print()`` commands are
            suppressed then this method will not be effective.
        """
        try:
            slide_size = cls.get_slide_size(slide)
        except mEx.SizeError:
            mLo.Lo.print("No slide size found")
            return
        slide_width = slide_size.width
        slide_height = slide_size.height

        x = cls._get_unit_mm_int(x)
        y = cls._get_unit_mm_int(y)

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
        cls,
        slide: XDrawPage,
        shape_type: DrawingShapeKind | str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
    ) -> XShape:
        """
        Adds a shape to a slide.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide
            shape_type (DrawingShapeKind | str): Shape type.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Newly added Shape.

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        See Also:
            - :py:meth:`~.draw.Draw.warns_position`
            - :py:meth:`~.draw.Draw.make_shape`

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        # shape are position from center.
        # a square of 20x20 would be x=10, y=10 for top right corer to be at 0x0
        try:
            cls.warns_position(slide=slide, x=x, y=y)
            shape = cls.make_shape(shape_type=shape_type, x=x, y=y, width=width, height=height)
            slide.add(shape)
            return shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error adding shape") from e

    @classmethod
    def draw_rectangle(
        cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> XShape:
        """
        Gets a rectangle.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Rectangle Shape.

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        return cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.RECTANGLE_SHAPE, x=x, y=y, width=width, height=height
        )

    @classmethod
    def draw_circle(cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, radius: int | UnitT) -> XShape:
        """
        Gets a circle.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            radius (int, UnitT): Shape radius in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Circle Shape.
        """
        # get the mm value of the units
        x = cls._get_unit_mm_int(x, 1)
        y = cls._get_unit_mm_int(y, 1)
        radius = cls._get_unit_mm_int(radius, 1)
        return cls.add_shape(
            slide=slide,
            shape_type=DrawingShapeKind.ELLIPSE_SHAPE,
            x=x - radius,
            y=y - radius,
            width=radius * 2,
            height=radius * 2,
        )

    @classmethod
    def draw_ellipse(
        cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> XShape:
        """
        Gets an ellipse.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Ellipse Shape.
        """
        return cls.add_shape(
            slide=slide, shape_type=DrawingShapeKind.ELLIPSE_SHAPE, x=x, y=y, width=width, height=height
        )

    # region draw_polygon()
    @overload
    @classmethod
    def draw_polygon(cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, sides: PolySides | int) -> XShape:
        """
        Gets a polygon.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            sides (PolySides | int): Polygon Sides value from ``3`` to ``30``.

        Returns:
            XShape: Polygon Shape.
        """
        ...

    @overload
    @classmethod
    def draw_polygon(
        cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, sides: PolySides | int, radius: int
    ) -> XShape:
        """
        Gets a polygon.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            sides (PolySides | int): Polygon Sides value from ``3`` to ``30``.
            radius (int, optional): Shape radius in mm units. Defaults to the value of :py:attr:`.Draw.POLY_RADIUS`.

        Returns:
            XShape: Polygon Shape.
        """
        ...

    @classmethod
    def draw_polygon(
        cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, sides: PolySides | int, radius: int = POLY_RADIUS
    ) -> XShape:
        """
        Gets a polygon.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            sides (PolySides | int): Polygon Sides value from ``3`` to ``30``.
            radius (int, optional): Shape radius in mm units. Defaults to the value of :py:attr:`.Draw.POLY_RADIUS`.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Polygon Shape.
        """
        try:
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
            polys = (pts,)
            prop_set = mLo.Lo.qi(XPropertySet, polygon, raise_err=True)
            poly_seq = uno.Any("[][]com.sun.star.awt.Point", polys)  # type: ignore
            uno.invoke(prop_set, "setPropertyValue", ("PolyPolygon", poly_seq))  # type: ignore
            return polygon
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error drawing polygon shape") from e

    # endregion draw_polygon()

    @classmethod
    def gen_polygon_points(
        cls, x: int | UnitT, y: int | UnitT, radius: int | UnitT, sides: PolySides | int
    ) -> Tuple[Point, ...]:
        """
        Generates a list of polygon points.

        |lo_safe|

        Args:
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            radius (int, UnitT): Shape radius in mm units or UnitT.
            sides (PolySides, int): Number of Polygon sides from 3 to 30.

        Raises:
            DrawError: If Error occurs

        Returns:
            Tuple[Point, ...]: Tuple of points.
        """
        try:
            x = cls._get_mm100_obj_from_mm(x, 1).value
            y = cls._get_mm100_obj_from_mm(y, 1).value
            radius = cls._get_mm100_obj_from_mm(radius, 1).value
            sides = PolySides(int(sides))
            pts: List[Point] = []
            angle_step = math.pi / sides.value
            for i in range(sides.value):
                pt = Point(
                    round((x + radius * math.cos(i * 2 * angle_step))),
                    round((y + radius * math.sin(i * 2 * angle_step))),
                )
                pts.append(pt)
            return tuple(pts)
        except Exception as e:
            raise mEx.DrawError("Unable to generate polygon points") from e

    @classmethod
    def draw_bezier(
        cls,
        slide: XDrawPage,
        pts: Sequence[Point] | Sequence[Sequence[Point]],
        flags: Sequence[PolygonFlags] | Sequence[Sequence[PolygonFlags]],
        is_open: bool,
    ) -> XShape:
        """
        Draws a bezier curve.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            pts (Sequence[Point]): Points.
            flags (Sequence[PolygonFlags]): Flags.
            is_open (bool): Determines if an open or closed bezier is drawn.

        Raises:
            IndexError: If ``pts`` and ``flags`` do not have the same number of elements.
            ShapeError: If unable to create Bezier Shape.

        Returns:
            XShape: Bezier Shape.
        """
        if len(pts) == 0:
            raise mEx.ShapeError("No points were provided")
        if len(pts) != len(flags):
            raise IndexError("pts and flags must be the same length")

        try:
            point0 = pts[0]
            # find out if the first point is indeed a point or if it is a sequence of points,
            coords = PolyPolygonBezierCoords()
            if mInfo.Info.is_struct(point0):
                coords.Coordinates = (pts,)  # type: ignore
                coords.Flags = (flags,)  # type: ignore
            else:
                # assume that we have sequences of sequences of points and flags.

                coords.Coordinates = mTableHelper.TableHelper.to_2d_tuple(pts)  # type: ignore
                coords.Flags = mTableHelper.TableHelper.to_2d_tuple(flags)  # type: ignore
            # type: ignore

            bezier_type = "OpenBezierShape" if is_open else "ClosedBezierShape"
            bezier_poly = cls.add_shape(slide=slide, shape_type=bezier_type, x=0, y=0, width=0, height=0)
            # create space for one bezier shape

            mProps.Props.set(bezier_poly, PolyPolygonBezier=coords)
            return bezier_poly
        except IndexError:
            raise
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to create bezier shape.") from e

    @classmethod
    def draw_line(cls, slide: XDrawPage, x1: int | UnitT, y1: int | UnitT, x2: int | UnitT, y2: int | UnitT) -> XShape:
        """
        Draws a line.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x1 (int, UnitT): Line start X position in mm units or UnitT.
            y1 (int, UnitT): Line start Y position mm units or UnitT.
            x2 (int, UnitT): Line end X position mm units or UnitT.
            y2 (int, UnitT): Line end Y position mm units or UnitT.

        Raises:
            ValueError: If x values and y values are a point and not a line.
            ShapeError: If unable to create Line.

        Returns:
            XShape: Line Shape.
        """
        x1 = cls._get_unit_mm_int(x1, 1)
        y1 = cls._get_unit_mm_int(y1, 1)
        x2 = cls._get_unit_mm_int(x2, 1)
        y2 = cls._get_unit_mm_int(y2, 1)
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
    def draw_polar_line(
        cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, degrees: int, distance: int | UnitT
    ) -> XShape:
        """
        Draw a line from ``x``, ``y`` in the direction of degrees, for the specified distance
        degrees is measured clockwise from x-axis.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            degrees (int): Direction of degrees.
            distance (int, UnitT): Distance of line in mm units or UnitT.

        Raises:
            ShapeError: If unable to create Polar Line Shape.

        Returns:
            XShape: Polar Line Shape.
        """
        try:
            distance = cls._get_unit_mm_int(distance, 1)
            x = cls._get_unit_mm_int(x, 1)
            y = cls._get_unit_mm_int(y, 1)
            x_dist = round(math.cos(math.radians(degrees)) * distance)
            y_dist = round(math.sin(math.radians(degrees)) * distance) * -1  # convert to negative
            return cls.draw_line(slide=slide, x1=x, y1=y, x2=x + x_dist, y2=y + y_dist)
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to create polar line shape.") from e

    @classmethod
    def draw_lines(cls, slide: XDrawPage, xs: Sequence[Union[int, UnitT]], ys: Sequence[Union[int, UnitT]]) -> XShape:
        """
        Draw lines.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide
            xs (Sequence[Union[int, UnitT]): Sequence of X positions in mm units or UnitT.
            ys (Sequence[Union[int, UnitT]): Sequence of Y positions in mm units or UnitT.

        Raises:
            IndexError: If ``xs`` and ``xy`` do not have the same number of elements.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Lines Shape.

        Note:
            The number of points must be the same for both ``xs`` and ``ys``.
        """
        xs = [cls._get_unit_mm_int(x, 1) for x in xs]
        ys = [cls._get_unit_mm_int(y, 1) for y in ys]
        num_points = len(xs)
        if num_points != len(ys):
            raise IndexError("xs and ys must be the same length")

        try:
            pts: List[Point] = []
            for x, y in zip(xs, ys):
                # in 1/100 mm units
                pts.append(Point(x * 100, y * 100))

            # an array of Point arrays, one Point array for each line path
            line_paths = (tuple(pts),)

            # for a shape formed by from multiple connected lines
            poly_line = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.POLY_LINE_SHAPE, x=0, y=0, width=0, height=0
            )
            prop_set = mLo.Lo.qi(XPropertySet, poly_line, raise_err=True)
            seq = uno.Any("[][]com.sun.star.awt.Point", line_paths)  # type: ignore
            uno.invoke(prop_set, "setPropertyValue", ("PolyPolygon", seq))  # type: ignore
            return poly_line
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error occurred while drawing lines.") from e

    # region draw_text()
    @overload
    @classmethod
    def draw_text(cls, slide: XDrawPage, msg: str, x: int, y: int, width: int, height: int) -> XShape:
        """
        Draws Text.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            msg (str): Text to draw.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Returns:
            XShape: Shape
        """
        ...

    @overload
    @classmethod
    def draw_text(cls, slide: XDrawPage, msg: str, x: int, y: int, width: int, height: int, font_size: int) -> XShape:
        """
        Draws Text.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            msg (str): Text to draw.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.
            font_size (float, UnitT, optional): Font size of text in Points or UnitT.

        Returns:
            XShape: Shape
        """
        ...

    @classmethod
    def draw_text(
        cls,
        slide: XDrawPage,
        msg: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        font_size: float | UnitT = 0,
    ) -> XShape:
        """
        Draws Text.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            msg (str): Text to draw.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.
            font_size (float, UnitT, optional): Font size of text in Points or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Shape

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        try:
            shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.TEXT_SHAPE, x=x, y=y, width=width, height=height
            )
            cls.add_text(shape=shape, msg=msg, font_size=font_size)
            return shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error occurred while drawing text.") from e

    # endregion draw_text()

    @classmethod
    def add_text(cls, shape: XShape, msg: str, font_size: float | UnitT = 0, **props) -> None:
        """
        Add text to a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape
            msg (str): Text to add
            font_size (float, optional): Font size in Points or UnitT.
            props (Any, optional): Any extra properties that will be applied to cursor (font) such as ``CharUnderline=1``

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        try:
            font_size = cls._get_pt_obj_from_pt(font_size).value
            txt = mLo.Lo.qi(XText, shape, True)
            cursor = txt.createTextCursor()
            cursor.gotoEnd(False)
            if font_size > 0:
                mProps.Props.set(cursor, CharHeight=font_size)

            mProps.Props.set(cursor, **props)

            rng = mLo.Lo.qi(XTextRange, cursor, True)
            rng.setString(msg)
        except Exception as e:
            raise mEx.ShapeError("Error occurred while adding text to shape.") from e

    @classmethod
    def add_connector(
        cls,
        slide: XDrawPage,
        shape1: XShape,
        shape2: XShape,
        start_conn: GluePointsKind | None = None,
        end_conn: GluePointsKind | None = None,
    ) -> XShape:
        """
        Add connector.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            shape1 (XShape): First Shape to add connector to.
            shape2 (XShape): Second Shape to add connector to.
            start_conn (GluePointsKind | None, optional): Start connector kind. Defaults to right.
            end_conn (GluePointsKind | None, optional): End connector kind. Defaults to left.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Connector Shape.

        Note:
            Properties for shape can be added or changed by using :py:meth:`~.draw.Draw.set_shape_props`.

            For instance the default value is ``EndShape=ConnectorType.STANDARD``.
            This could be changed.

            .. code-block:: python

                Draw.set_shape_props(shape, EndShape=ConnectorType.CURVE)
        """
        if start_conn is None:
            start_conn = GluePointsKind.RIGHT
        if end_conn is None:
            end_conn = GluePointsKind.LEFT

        try:
            x_connector = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.CONNECTOR_SHAPE, x=0, y=0, width=0, height=0
            )
            prop_set = mLo.Lo.qi(XPropertySet, x_connector, True)
            prop_set.setPropertyValue("StartShape", shape1)
            prop_set.setPropertyValue("StartGluePointIndex", int(start_conn))

            prop_set.setPropertyValue("EndShape", shape2)
            prop_set.setPropertyValue("EndGluePointIndex", int(end_conn))

            prop_set.setPropertyValue("EdgeKind", ConnectorType.STANDARD)
            return x_connector
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Could not connect the shapes") from e

    @staticmethod
    def get_glue_points(shape: XShape) -> Tuple[GluePoint2, ...]:
        """
        Gets Glue Points.

        |lo_safe|

        Args:
            shape (XShape): Shape.

        Raises:
            DrawError: If error occurs.

        Returns:
            Tuple[GluePoint2, ...]: Glue Points.

        Note:
            If a glue point can not be accessed then it is ignored.
        """
        try:
            gp_supp = mLo.Lo.qi(XGluePointsSupplier, shape, True)
            glue_pts = gp_supp.getGluePoints()

            num_gps = glue_pts.getCount()  # should be 4 by default
            if num_gps == 0:
                mLo.Lo.print("No glue points for this shape")
                return ()

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
        except Exception as e:
            raise mEx.DrawError("Error getting glue points.") from e

    @classmethod
    def get_chart_shape(
        cls, slide: XDrawPage, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> XShape:
        """
        Gets a chart shape.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If Error occurs.

        Returns:
            XShape: Chart Shape.
        """
        try:
            shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.OLE2_SHAPE, x=x, y=y, width=width, height=height
            )
            mProps.Props.set(shape, CLSID=mLo.Lo.CLSID.CHART.value)  # a chart
            return shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to get chart shape") from e

    @classmethod
    def draw_formula(
        cls, slide: XDrawPage, formula: str, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> XShape:
        """
        Draws a formula.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            formula (str): Formula as string to draw.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, , UnitT): Shape Y position in mm units or UnitT.
            width (int, , UnitT): Shape width in mm units or UnitT.
            height (int, , UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Formula Shape.

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        try:
            x = cls._get_unit_mm_int(x)
            y = cls._get_unit_mm_int(y)
            width = cls._get_unit_mm_int(width)
            height = cls._get_unit_mm_int(height)
            shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.OLE2_SHAPE, x=x, y=y, width=width, height=height
            )
            cls.set_shape_props(shape, CLSID=str(mLo.Lo.CLSID.MATH))  # a formula

            model = mLo.Lo.qi(XModel, mProps.Props.get(shape, "Model"), True)
            # mInfo.Info.show_services(obj_name="OLE2Shape Model", obj=model)
            mProps.Props.set(model, Formula=formula)

            # for some reason setting model Formula here cause the shape size to be blown out.
            # resetting size and position corrects the issue.
            cls.set_size(shape, width, height)
            cls.set_position(shape, Point(x, y))
            return shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to get Draw formula") from e

    @classmethod
    def draw_media(
        cls, slide: XDrawPage, fnm: PathOrStr, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> XShape:
        """
        Draws media.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            fnm (PathOrStr): Path to Media file.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Media shape.

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.

        .. versionchanged:: 0.17.14
            Now does not set size and/or position unless the values are greater than ``0``.
        """
        # could not find MediaShape in api.
        # https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html
        # however it can be found in examples.
        # https://ask.libreoffice.org/t/how-to-add-video-to-impress-with-python/33050/2
        try:
            shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.MEDIA_SHAPE, x=x, y=y, width=width, height=height
            )

            # mProps.Props.show_obj_props(prop_kind="Shape", obj=shape)
            mLo.Lo.print(f'Loading media: "{fnm}"')
            cls.set_shape_props(shape, Loop=True, MediaURL=mFileIO.FileIO.fnm_to_url(fnm))
            return shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to get Draw media") from e

    @staticmethod
    def is_group(shape: XShape) -> bool:
        """
        Gets if a shape is a Group Shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Returns:
            bool: ``True`` if shape is a group; Otherwise; ``False``.
        """
        if shape is None:
            mLo.Lo.print("Shape is None. Not able to check for group")
            return False
        return shape.getShapeType() == "com.sun.star.drawing.GroupShape"

    @staticmethod
    def combine_shape(doc: XComponent, shapes: XShapes, combine_op: ShapeCombKind) -> XShape:
        """
        Combines one or more shapes using a dispatch command.

        |lo_unsafe|

        Args:
            doc (XComponent): Document.
            shapes (XShapes): Shapes to combine.
            combine_op (ShapeCompKind): Combine Operation.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: New combined shape.
        """
        # select the shapes for the dispatches to apply to
        try:
            sel_supp = mLo.Lo.qi(XSelectionSupplier, mGui.GUI.get_current_controller(doc), True)
            sel_supp.select(shapes)

            if combine_op == ShapeCombKind.INTERSECT:
                mLo.Lo.dispatch_cmd("Intersect")
            elif combine_op == ShapeCombKind.SUBTRACT:
                mLo.Lo.dispatch_cmd("Substract")  # misspelt!
            elif combine_op == ShapeCombKind.COMBINE:
                mLo.Lo.dispatch_cmd("Combine")
            else:
                mLo.Lo.dispatch_cmd("Merge")

            mLo.Lo.delay(500)  # give time for dispatches to arrive and be processed

            # extract the new single shape from the modified selection
            xs = mLo.Lo.qi(XShapes, sel_supp.getSelection(), True)
            return mLo.Lo.qi(XShape, xs.getByIndex(0), True)
        except Exception as e:
            raise mEx.ShapeError("Unable to combine shapes") from e

    @classmethod
    def create_control_shape(
        cls,
        label: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        shape_kind: FormControlKind | str,
        **props,
    ) -> XControlShape:
        """
        Creates a control shape.

        |lo_unsafe|

        Args:
            label (str): Label to apply.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.
            shape_kind (FormControlKind | str): The kind of control to create.
            props (Any, optional): Any extra key value options to set on the Model of the control being created
                such as ``FontHeight=18.0, Name="BLA"``.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XControlShape: Control Shape.

        .. versionchanged:: 0.17.0
            No longer sets ``FontHeight`` or ``FontName`` properties.
        """
        try:
            x = cls._get_mm100_obj_from_mm(x).value
            y = cls._get_mm100_obj_from_mm(y).value
            width = cls._get_mm100_obj_from_mm(width).value
            height = cls._get_mm100_obj_from_mm(height).value
            c_shape = mLo.Lo.create_instance_msf(XControlShape, "com.sun.star.drawing.ControlShape", raise_err=True)
            c_shape.setSize(UnoSize(width, height))
            c_shape.setPosition(Point(x, y))

            c_model = mLo.Lo.create_instance_msf(
                XControlModel, f"com.sun.star.form.control.{shape_kind}", raise_err=True
            )

            prop_set = mLo.Lo.qi(XPropertySet, c_model, True)
            prop_set.setPropertyValue("DefaultControl", f"com.sun.star.form.control.{shape_kind}")
            prop_set.setPropertyValue("Name", cls._create_name("Control"))
            if label:
                prop_set.setPropertyValue("Label", label)

            # prop_set.setPropertyValue("FontHeight", 18.0)
            # prop_set.setPropertyValue("FontName", mInfo.Info.get_font_general_name())

            for k, v in props.items():
                prop_set.setPropertyValue(k, v)

            c_shape.setControl(c_model)

            # x_btn = mLo.Lo.qi(XButton, c_model)
            # if x_btn is None:
            #     mLo.Lo.print("XButton is None")

            # mProps.Props.show_props(title="Control model props", props=props)
            return c_shape
        except Exception as e:
            raise mEx.ShapeError("Unable to create control shape") from e

    # endregion draw/add shape to a page

    # region custom shape addition using dispatch and JNA

    # Two methods not include here from java. addDispatchShape and createDispatchShape
    # These were omitted because the require third party Libs that all for automatic screen click and screen moused selecting

    @classmethod
    def add_dispatch_shape(
        cls,
        slide: XDrawPage,
        shape_dispatch: ShapeDispatchKind | str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        fn: DispatchShape,
    ) -> XShape:
        """
        Adds a shape to a Draw slide via a dispatch command.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            shape_dispatch (ShapeDispatchKind | str): Dispatch Command.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.
            fn (DispatchShape): Function that is responsible for running the dispatch command and returning the shape.

        Raises:
            NoneError: If adding a dispatch fails.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape

        See Also:
            - :py:protocol:`~.proto.dispatch_shape.DispatchShape`
            - Example - `Impress Make Slides <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_make_slides>`__
        """
        cls.warns_position(slide, x, y)
        try:
            shape = fn(slide, str(shape_dispatch))
            if shape is None:
                raise mEx.NoneError(f'Failed to add shape for dispatch command "{shape_dispatch}"')
            cls.set_position(shape=shape, x=x, y=y)
            cls.set_size(shape=shape, width=width, height=height)
            return shape
        except mEx.NoneError:
            raise
        except Exception as e:
            raise mEx.ShapeError(
                f'Error occurred adding dispatch shape for dispatch command "{shape_dispatch}"'
            ) from e

    @staticmethod
    def create_dispatch_shape(slide: XDrawPage, shape_dispatch: ShapeDispatchKind | str, fn: DispatchShape) -> XShape:
        """
        Creates a shape via a dispatch command.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            shape_dispatch (ShapeDispatchKind | str): Dispatch Command.
            fn (DispatchShape): Function that is responsible for running the dispatch command and returning the shape.

        Raises:
            NoneError: If adding a dispatch fails.
            ShapeError: If any other error occurs.

        Returns:
            XShape: Shape

        See Also:
            Example - `Impress Make Slides <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_make_slides>`__
        """
        try:
            shape = fn(slide, str(shape_dispatch))
            if shape is None:
                raise mEx.NoneError(f'Failed to add shape for dispatch command "{shape_dispatch}"')
            return shape
        except mEx.NoneError:
            raise
        except Exception as e:
            raise mEx.ShapeError(
                f'Error occurred creating dispatch shape for dispatch command "{shape_dispatch}"'
            ) from e

    # endregion

    # region presentation shapes
    @classmethod
    def set_master_footer(cls, master: XDrawPage, text: str) -> None:
        """
        Sets master footer text.

        |lo_safe|

        Args:
            master (XDrawPage): Master Draw Page.
            text (str): Footer text.

        Raises:
            ShapeMissingError: If unable to find footer shape.
            DrawPageError: If any other error occurs.

        Returns:
            None:
        """
        try:
            footer_shape = cls.find_shape_by_type(slide=master, shape_type=DrawingNameSpaceKind.SHAPE_TYPE_FOOTER)
            txt_field = mLo.Lo.qi(XText, footer_shape, True)
            txt_field.setString(text)
        except mEx.ShapeMissingError:
            raise
        except Exception as e:
            raise mEx.DrawPageError("Unable to set master footer") from e

    @classmethod
    def add_slide_number(cls, slide: XDrawPage) -> XShape:
        """
        Adds slide number to a slide.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Slide number shape.
        """
        try:
            sz = cls.get_slide_size(slide)
            width = 60
            height = 15
            return cls.add_pres_shape(
                slide=slide,
                shape_type=PresentationKind.SLIDE_NUMBER_SHAPE,
                x=sz.width - width - 12,
                y=sz.height - height - 4,
                width=width,
                height=height,
            )
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to add slide number") from e

    @classmethod
    def add_pres_shape(
        cls,
        slide: XDrawPage,
        shape_type: PresentationKind,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
    ) -> XShape:
        """
        Creates a shape from the "com.sun.star.presentation" package.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            shape_type (PresentationKind): Kind of presentation package to create.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Presentation Shape.
        """
        try:
            cls.warns_position(slide=slide, x=x, y=y)
            shape = mLo.Lo.create_instance_msf(XShape, shape_type.to_namespace(), raise_err=True)
            if shape is not None:
                slide.add(shape)
                cls.set_position(shape, x, y)
                cls.set_size(shape, width, height)
            return shape
        except Exception as e:
            raise mEx.ShapeError("Unable to add slide number") from e

    # endregion presentation shapes

    # region get/set drawing properties
    @staticmethod
    def get_position(shape: XShape) -> Point:
        """
        Gets position in mm units.

        |lo_safe|

        Args:
            shape (XShape): Shape.

        Raises:
            PointError: If error occurs.

        Returns:
            Point: Position as Point in mm units.
        """
        try:
            pt = shape.getPosition()
            # convert to mm
            x = UnitMM.from_mm100(pt.X).value
            y = UnitMM.from_mm100(pt.Y).value
            return Point(round(x), round(y))
        except Exception as e:
            raise mEx.PointError("Error getting position") from e

    @staticmethod
    def get_size(shape: XShape) -> Size:
        """
        Gets Size in mm units.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            SizeError: If error occurs.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size in mm units
        """
        try:
            sz = shape.getSize()
            # convert to mm
            width = UnitMM.from_mm100(sz.Width).value
            height = UnitMM.from_mm100(sz.Height).value
            return Size(round(width), round(height))
        except Exception as e:
            raise mEx.SizeError("Error getting Size") from e

    @staticmethod
    def print_point(pt: Point) -> None:
        """
        Prints point to console in mm units.

        |lo_safe|

        Args:
            pt (Point): Point object.

        Returns:
            None:
        """
        print(f"  Point (mm): [{round(pt.X/100)}, {round(pt.Y/100)}]")

    @staticmethod
    def print_size(sz: SizeObj) -> None:
        """
        Prints size to console in mm units.

        |lo_safe|

        Args:
            sz (SizeObj): Size object.

        Returns:
            None:
        """
        print(f"  Size (mm): [{round(sz.Width/100)}, {round(sz.Height/100)}]")

    @classmethod
    def report_pos_size(cls, shape: XShape) -> None:
        """
        Prints shape information to the console.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Returns:
            None:
        """
        if shape is None:
            print("The shape is null")
            return
        try:
            print(f'Shape Name: {mProps.Props.get(shape, "Name")}')
        except mEx.PropertyNotFoundError:
            print("Shapes does not have a name property")

        print(f"  Type: {shape.getShapeType()}")
        # not all shapes have size and position such as a FrameShape
        with contextlib.suppress(Exception):
            cls.print_point(shape.getPosition())
        with contextlib.suppress(Exception):
            cls.print_size(shape.getSize())

    # region set_position()

    @overload
    @classmethod
    def set_position(cls, shape: XShape, pt: Point) -> None:
        """
        Sets Position of shape.

        |lo_safe|

        Args:
            shape (XShape): Shape
            pt (point): Point that contains x and y positions in mm units.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_position(cls, shape: XShape, x: int | UnitT, y: int | UnitT) -> None:
        """
        Sets Position of shape.

        |lo_safe|

        Args:
            shape (XShape): Shape
            x (int, UnitT): X position in mm units or UnitT.
            y (int, UnitT): Y Position in mm units or UnitT.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_position(cls, *args, **kwargs) -> None:
        """
        Sets Position of shape.

        |lo_safe|

        Args:
            shape (XShape): Shape
            pt (point): Point that contains x and y positions in mm units.
            x (int, UnitT): X position in mm units or UnitT.
            y (int, UnitT): Y Position in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:

        Note:
            If ``x`` or ``y`` is negative or ``0`` then the shape position will not be set.

        .. versionchanged:: 0.17.14
            Now does not set position unless the values are greater than ``0``.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "pt", "x", "y")
            check = all(key in valid_keys for key in kwargs)
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
            ka[3] = kwargs.get("y", None)
            return ka

        if count not in (2, 3):
            raise TypeError("set_position() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        try:
            if count == 2:
                # def set_position(shape: XShape, pt: Point)
                pt_in = cast(Point, kargs[2])
                x = cls._get_mm100_obj_from_mm(pt_in.X).value
                y = cls._get_mm100_obj_from_mm(pt_in.Y).value
            else:
                # def set_position(shape: XShape, x:int, y: int)
                x = cls._get_mm100_obj_from_mm(kargs[2]).value
                y = cls._get_mm100_obj_from_mm(kargs[3]).value
            if x > 0 and y > 0:
                pt = Point(x, y)
                cast(XShape, kargs[1]).setPosition(pt)
        except Exception as e:
            raise mEx.ShapeError("Error setting position") from e

    # endregion set_position()

    # region set_size()

    @overload
    @classmethod
    def set_size(cls, shape: XShape, sz: SizeObj) -> None:
        """
        Sets set_size of shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            sz (~ooodev.utils.data_type.size.Size): Size that contains width and height positions in mm units.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_size(cls, shape: XShape, width: int | UnitT, height: int | UnitT) -> None:
        """
        Sets set_size of shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            width (int, UnitT): Width position in mm units or UnitT.
            height (int, UnitT): Height position in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_size(cls, *args, **kwargs) -> None:
        """
        Sets set_size of shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            sz (~ooodev.utils.data_type.size.Size): Size that contains width and height positions in mm units.
            width (int, UnitT): Width position in mm units or UnitT.
            height (int, UnitT): Height position in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:

        Note:
            If ``width`` or ``height`` is negative or ``0`` then the shape size will not be set.


        .. versionchanged:: 0.17.14
            Now does not set size unless the values are greater than ``0``.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "sz", "width", "height")
            check = all(key in valid_keys for key in kwargs)
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
            ka[3] = kwargs.get("height", None)
            return ka

        if count not in (2, 3):
            raise TypeError("set_size() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        try:
            if count == 2:
                # def set_size(shape: XShape, sz: Size)
                sz_in = cast(SizeObj, kargs[2])
                width = cls._get_mm100_obj_from_mm(sz_in.Width).value
                height = cls._get_mm100_obj_from_mm(sz_in.Height).value
            else:
                # def set_size(shape: XShape, width:int, height: int)
                width = cls._get_mm100_obj_from_mm(kargs[2]).value
                height = cls._get_mm100_obj_from_mm(kargs[3]).value
            if width <= 0 or height <= 0:
                return
            sz = Size(width, height)
            cast(XShape, kargs[1]).setSize(sz.get_uno_size())
        except Exception as e:
            raise mEx.ShapeError("Error setting size") from e

    # endregion set_size()

    @staticmethod
    def set_style(shape: XShape, graphic_styles: XNameContainer, style_name: GraphicStyleKind | str) -> None:
        """
        Set the graphic style for a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            graphic_styles (XNameContainer): Graphic styles.
            style_name (GraphicStyleKind | str): Graphic Style Name.

        Raises:
            DrawError: If error occurs.

        Returns:
            None:
        """
        try:
            style = mLo.Lo.qi(XStyle, graphic_styles.getByName(str(style_name)), True)
            mProps.Props.set(shape, Style=style)
        except Exception as e:
            raise mEx.DrawError(f'Could not set the style to "{style_name}"') from e

    @staticmethod
    def get_text_properties(shape: XShape) -> XPropertySet:
        """
        Gets the properties associated with the text area inside the shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            PropertySetError: If error occurs.

        Returns:
            XPropertySet: Property Set
        """
        try:
            x_txt = mLo.Lo.qi(XText, shape, True)
            cursor = x_txt.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            x_rng = mLo.Lo.qi(XTextRange, cursor, True)
            return mLo.Lo.qi(XPropertySet, x_rng, True)
        except Exception as e:
            raise mEx.PropertySetError("Error getting text properties") from e

    @staticmethod
    def get_line_color(shape: XShape) -> mColor.Color:
        """
        Gets the line color of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            ColorError: If error occurs.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
            return mColor.Color(int(props.getPropertyValue("LineColor")))
        except Exception as e:
            raise mEx.ColorError("Error getting line color") from e

    @staticmethod
    def set_dashed_line(shape: XShape, is_dashed: bool) -> None:
        """
        Set a dashed line.

        |lo_safe|

        Args:
            shape (XShape): Shape
            is_dashed (bool): Determines if line is to be dashed or solid.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        # shape is also com.sun.star.drawing.FillProperties service
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
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
            raise mEx.ShapeError("Error setting dashed line property") from e

    @staticmethod
    def get_line_thickness(shape: XShape) -> int:
        """
        Gets line thickness of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Returns:
            int: Line Thickness on success; Otherwise, ``0``.
        """
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
            return int(props.getPropertyValue("LineWidth"))
        except Exception as e:
            mLo.Lo.print("Could not access line thickness")
            mLo.Lo.print(f"  {e}")
        return 0

    @staticmethod
    def get_fill_color(shape: XShape) -> mColor.Color:
        """
        Gets the fill color of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            ColorError: If error occurs.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
            return mColor.Color(int(props.getPropertyValue("FillColor")))
        except Exception as e:
            raise mEx.ColorError("Error getting fill color") from e

    @staticmethod
    def set_transparency(shape: XShape, level: Intensity | int) -> None:
        """
        Sets the transparency level for the shape.
        Higher level means more transparent.

        |lo_safe|

        Args:
            shape (XShape): Shape
            level (Intensity, int): Transparency value. Represents a intensity value from ``0`` to ``100``.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        try:
            level = Intensity(int(level))
            mProps.Props.set(shape, FillTransparence=level.value)
        except Exception as e:
            raise mEx.ShapeError("Error setting transparency") from e

    @staticmethod
    def set_gradient_properties(shape: XShape, grad: Gradient) -> None:
        """
        Sets shapes gradient properties.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            grad (~com.sun.star.awt.Gradient): Gradient properties to set.

        Returns:
            None:

        See Also:
            :py:meth:`~.Draw.set_gradient_color`
        """
        # shape is also com.sun.star.drawing.FillProperties service
        mProps.Props.set(shape, FillStyle=FillStyle.GRADIENT, FillGradient=grad)

    # region set_gradient_color()

    @staticmethod
    def _set_gradient_color_name(shape: XShape, name: DrawingGradientKind | str) -> Gradient:
        """Lo Safe Method"""
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
            props.setPropertyValue("FillStyle", FillStyle.GRADIENT)
            props.setPropertyValue("FillGradientName", str(name))
            return props.getPropertyValue("FillGradient")
        except IllegalArgumentException as e:
            raise NameError(f'"{name}" is not a recognized gradient name') from e
        except Exception as e:
            raise mEx.ShapeError(f"Unable to set shape gradient color: {name}") from e

    @classmethod
    def _set_gradient_color_colors(
        cls, shape: XShape, start_color: mColor.Color, end_color: mColor.Color, angle: Angle
    ) -> Gradient:
        """Lo Safe Method."""
        try:
            grad = Gradient()
            grad.Style = GradientStyle.LINEAR  # type: ignore
            grad.StartColor = start_color  # type: ignore
            grad.EndColor = end_color  # type: ignore

            grad.Angle = angle.value * 10  # in 1/10 degree units
            grad.Border = 0
            grad.XOffset = 0
            grad.YOffset = 0
            grad.StartIntensity = 100
            grad.EndIntensity = 100
            grad.StepCount = 10

            cls.set_gradient_properties(shape, grad)

            return mProps.Props.get(shape, "FillGradient")

        except Exception as e:
            # f-string = is python >= 3.8
            raise mEx.ShapeError(
                f"Unable to set shape gradient color: start_color={start_color}, end_color={end_color}, {angle}"
            ) from e

    @overload
    @classmethod
    def set_gradient_color(cls, shape: XShape, name: DrawingGradientKind | str) -> Gradient:
        """
        Set the gradient color of the shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            name (DrawingGradientKind | str): Gradient color name.

        Returns:
            Gradient: Gradient instance that just had properties set.
        """
        ...

    @overload
    @classmethod
    def set_gradient_color(cls, shape: XShape, start_color: mColor.Color, end_color: mColor.Color) -> Gradient:
        """
        Set the gradient color of the shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            start_color (~ooodev.utils.color.Color): Start Color.
            end_color (~ooodev.utils.color.Color): End Color.

        Returns:
            Gradient: Gradient instance that just had properties set.
        """
        ...

    @overload
    @classmethod
    def set_gradient_color(
        cls, shape: XShape, start_color: mColor.Color, end_color: mColor.Color, angle: Angle | int
    ) -> Gradient:
        """
        Set the gradient color of the shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            start_color (~ooodev.utils.color.Color): Start Color.
            end_color (~ooodev.utils.color.Color): End Color.
            angle (Angle, int): Gradient angle.

        Returns:
            Gradient: Gradient instance that just had properties set.
        """
        ...

    @classmethod
    def set_gradient_color(cls, *args, **kwargs) -> Gradient:
        """
        Set the gradient color of the shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            name (DrawingGradientKind | str): Gradient color name.
            start_color (~ooodev.utils.color.Color): Start Color.
            end_color (~ooodev.utils.color.Color): End Color.
            angle (Angle, int): Gradient angle.

        Raises:
            NameError: If ``name`` is not recognized.
            ShapeError: If any other error occurs.

        Returns:
            ~com.sun.star.awt.Gradient: Gradient instance that just had properties set.

        Note:
            When using Gradient Name.

            Getting the gradient color name can be a bit challenging.
            ``DrawingGradientKind`` contains name displayed in the Gradient color menu of Draw.

            The Easiest way to get the colors is to open Draw and see what gradient color names are available
            on your system.

        See Also:
            :py:meth:`~.Draw.set_gradient_properties`
        """
        # shape is also com.sun.star.drawing.FillProperties service
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("shape", "name", "start_color", "end_color", "angle")
            check = all(key in valid_keys for key in kwargs)
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
            ka[3] = kwargs.get("end_color", None)
            if count == 3:
                return ka
            ka[4] = kwargs.get("angle", None)
            return ka

        if count not in (2, 3, 4):
            raise TypeError("set_gradient_color() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return cls._set_gradient_color_name(kargs[1], kargs[2])

        angle = Angle(0) if count == 3 else Angle(int(kargs[4]))
        return cls._set_gradient_color_colors(shape=kargs[1], start_color=kargs[2], end_color=kargs[3], angle=angle)

    # endregion set_gradient_color()

    @staticmethod
    def set_hatch_color(shape: XShape, name: DrawingHatchingKind | str) -> None:
        """
        Set hatching color of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            name (DrawingHatchingKind, str): Hatching Name.

        Raises:
            NameError: If ``name`` is not recognized.
            ShapeError: If any other error occurs.

        Returns:
            None:

        Note:
            Getting the hatching color name can be a bit challenging.
            ``DrawingHatchingKind`` contains name displayed in the Hatching color menu of Draw.

            The Easiest way to get the colors is to open Draw and see what gradient color names are available
            on your system.
        """
        # shape is also com.sun.star.drawing.FillProperties service
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
            props.setPropertyValue("FillStyle", FillStyle.HATCH)
            props.setPropertyValue("FillHatchName", str(name))
        except IllegalArgumentException as e:
            raise NameError(f'"{name}" is not a recognized hatching name') from e
        except Exception as e:
            raise mEx.ShapeError("Error setting hatch color") from e
        return None

    @staticmethod
    def set_bitmap_color(shape: XShape, name: DrawingBitmapKind | str) -> None:
        """
        Set bitmap color of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape
            name (DrawingBitmapKind, str): Bitmap Name

        Raises:
            NameError: If ``name`` is not recognized.
            ShapeError: If any other error occurs.

        Returns:
            None:

        Note:
            Getting the bitmap color name can be a bit challenging.
            ``DrawingBitmapKind`` contains name displayed in the Bitmap color menu of Draw.

            The Easiest way to get the colors is to open Draw and see what bitmap color names are available
            on your system.
        """
        # shape is also com.sun.star.drawing.FillProperties service
        props = mLo.Lo.qi(XPropertySet, shape, True)
        try:
            props.setPropertyValue("FillStyle", FillStyle.BITMAP)
            props.setPropertyValue("FillBitmapName", str(name))
        except IllegalArgumentException as e:
            raise NameError(f'"{name}" is not a recognized bitmap name') from e
        except Exception as e:
            raise mEx.ShapeError(f'Error setting bitmap color  to "{name}"') from e

    @staticmethod
    def set_bitmap_file_color(shape: XShape, fnm: PathOrStr) -> None:
        """
        Set bitmap color from file.

        |lo_safe|

        Args:
            shape (XShape): Shape
            fnm (PathOrStr): path to file.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        # shape is also com.sun.star.drawing.FillProperties service
        try:
            props = mLo.Lo.qi(XPropertySet, shape, True)
            props.setPropertyValue("FillStyle", FillStyle.BITMAP)
            props.setPropertyValue("FillBitmapURL", mFileIO.FileIO.fnm_to_url(fnm))
        except Exception as e:
            raise mEx.ShapeError(f'Could not set bitmap color using  "{fnm}"') from e

    @classmethod
    def set_line_style(cls, shape: XShape, style: LineStyle) -> None:
        """
        Set the line style for a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            style (LineStyle): Line Style.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        try:
            cls.set_shape_props(shape=shape, LineStyle=style)
        except Exception as e:
            raise mEx.ShapeError("Error setting line style") from e

    @classmethod
    def set_visible(cls, shape: XShape, is_visible: bool) -> None:
        """
        Set the line style for a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            is_visible (bool): Set is shape is visible or not.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        try:
            cls.set_shape_props(shape=shape, Visible=is_visible)
        except Exception as e:
            raise mEx.ShapeError("Error setting shape visibility") from e

    # "RotateAngle" is deprecated but is much simpler
    # than the matrix approach, and works correctly
    # for rotations around the center

    @classmethod
    def set_angle(cls, shape: XShape, angle: Angle | int) -> None:
        """
        Set the line style for a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            angle (Angle | int): Angle to set.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        try:
            angle = Angle(int(angle))
            cls.set_shape_props(shape=shape, RotateAngle=angle.value * 100)
        except Exception as e:
            raise mEx.ShapeError(f"Error setting shape angle: {angle}") from e

    @staticmethod
    def get_rotation(shape: XShape) -> Angle:
        """
        Gets the rotation of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            ShapeError: If error occurs.

        Returns:
            Angle: Rotation angle.
        """
        try:
            r_angle = int(mProps.Props.get(shape, "RotateAngle"))
            return Angle(round(r_angle / 100))
        except Exception as e:
            raise mEx.ShapeError("Error getting shape rotation") from e

    @staticmethod
    def set_rotation(shape: XShape, angle: Angle | int) -> None:
        """
        Set the rotation of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            angle (Angle | int): Angle or int. An angle value from ``0`` to ``359``.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        try:
            angle = Angle(int(angle))
            mProps.Props.set(shape, RotateAngle=angle.value * 100)
        except Exception as e:
            raise mEx.ShapeError("Error setting shape rotation") from e

    @staticmethod
    def get_transformation(shape: XShape) -> HomogenMatrix3:
        """
        Gets a transformation matrix which seems to represent a clockwise rotation.

        Homogeneous matrix has three homogeneous lines

        |lo_safe|

        Args:
            shape (XShape): Shape

        Raises:
            ShapeError: If error occurs.

        Returns:
            HomogenMatrix3: Matrix
        """
        #     Returns a transformation matrix, which seems to
        #     represent a clockwise rotation:
        #     cos(t)  sin(t) x
        #    -sin(t)  cos(t) y
        #       0       0    1
        try:
            return mProps.Props.get(shape, "Transformation")
        except Exception as e:
            raise mEx.ShapeError("Error getting shape transformation") from e

    @staticmethod
    def print_matrix(mat: HomogenMatrix3) -> None:
        """
        Prints matrix to console

        |lo_safe|

        Args:
            mat (HomogenMatrix3): Matrix

        Returns:
            None:
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
    def _draw_image_path(cls, slide: XDrawPage, fnm: PathOrStr) -> XShape:
        """LO UN-Safe Method."""
        try:
            slide_size = cls.get_slide_size(slide)
            try:
                im_size = mImgLo.ImagesLo.get_size_100mm(fnm)
            except Exception as ex:
                raise RuntimeError(f'Could not calculate size of "{fnm}"') from ex
            im_width = round(im_size.width / 100)  # in mm units
            im_height = round(im_size.height / 100)
            x = round((slide_size.width - im_width) / 2)
            y = round((slide_size.height - im_height) / 2)
            return cls._draw_image_path_x_y_w_h(slide=slide, fnm=fnm, x=x, y=y, width=im_width, height=im_height)
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError(f'Error getting shape for draw image for file: "{fnm}"') from e

    @classmethod
    def _draw_image_path_x_y(cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int) -> XShape:
        """LO UN-Safe Method."""
        try:
            try:
                im_size = mImgLo.ImagesLo.get_size_100mm(fnm)
            except Exception as ex:
                raise RuntimeError(f'Could not calculate size of "{fnm}"') from ex
            return cls._draw_image_path_x_y_w_h(
                slide=slide, fnm=fnm, x=x, y=y, width=round(im_size.width / 100), height=round(im_size.height / 100)
            )
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError(f'Error getting shape for draw image for file: "{fnm}"') from e

    @classmethod
    def _draw_image_path_x_y_w_h(
        cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int, width: int, height: int
    ) -> XShape:
        """LO UN-Safe Method."""
        # units in mm's
        mLo.Lo.print(f'Adding the picture "{fnm}"')
        try:
            im_shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.GRAPHIC_OBJECT_SHAPE, x=x, y=y, width=width, height=height
            )
            cls.set_image(im_shape, fnm)
            cls.set_line_style(shape=im_shape, style=LineStyle.NONE)
            return im_shape
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError(f'Error getting shape for draw image for file: "{fnm}"') from e

    @overload
    @classmethod
    def draw_image(cls, slide: XDrawPage, fnm: PathOrStr) -> XShape:
        """
        Draws an image.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            fnm (PathOrStr): Path to image.

        Returns:
            XShape: Shape.
        """
        ...

    @overload
    @classmethod
    def draw_image(cls, slide: XDrawPage, fnm: PathOrStr, x: int | UnitT, y: int | UnitT) -> XShape:
        """
        Draws an image.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            fnm (PathOrStr): Path to image.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.

        Returns:
            XShape: Shape.
        """
        ...

    @overload
    @classmethod
    def draw_image(
        cls, slide: XDrawPage, fnm: PathOrStr, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT
    ) -> XShape:
        """
        Draws an image.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            fnm (PathOrStr): Path to image.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Returns:
            XShape: Shape.
        """
        ...

    @classmethod
    def draw_image(cls, *args, **kwargs) -> XShape:
        """
        Draws an image.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide.
            fnm (PathOrStr): Path to image.
            x (int, UnitT): Shape X position in mm units or UnitT.
            y (int, UnitT): Shape Y position in mm units or UnitT.
            width (int, UnitT): Shape width in mm units or UnitT.
            height (int, UnitT): Shape height in mm units or UnitT.

        Raises:
            ShapeError: If error occurs.

        Returns:
            XShape: Shape.
        """
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("slide", "fnm", "x", "y", "width", "height")
            check = all(key in valid_keys for key in kwargs)
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

        if count not in (2, 4, 6):
            raise TypeError("draw_image() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return cls._draw_image_path(slide=kargs[1], fnm=kargs[2])
        x = cls._get_unit_mm_int(kargs[3])
        y = cls._get_unit_mm_int(kargs[4])
        if count == 4:
            return cls._draw_image_path_x_y(slide=kargs[1], fnm=kargs[2], x=x, y=y)
        width = cls._get_unit_mm_int(kargs[5])
        height = cls._get_unit_mm_int(kargs[6])
        return cls._draw_image_path_x_y_w_h(slide=kargs[1], fnm=kargs[2], x=x, y=y, width=width, height=height)

    # endregion draw_image()

    @staticmethod
    def set_image(shape: XShape, fnm: PathOrStr) -> None:
        """
        Sets the image of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            fnm (PathOrStr): Path to image.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        # GraphicURL is Deprecated using Graphic instead.
        # https://tinyurl.com/2qaqs2nr#a6312a2da62e2c67c90d5576502117906
        try:
            graphic = mImgLo.ImagesLo.load_graphic_file(fnm)

            mProps.Props.set(shape, Graphic=graphic)
        except Exception as e:
            raise mEx.ShapeError(f'Error setting shape image to: "{fnm}"') from e

    @staticmethod
    def set_image_graphic(shape: XShape, graphic: XGraphic) -> None:
        """
        Sets the image of a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape.
            graphic (XGraphic): Graphic.

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        # GraphicURL is Deprecated using Graphic instead.
        # https://tinyurl.com/2qaqs2nr#a6312a2da62e2c67c90d5576502117906
        try:
            mProps.Props.set(shape, Graphic=graphic)
        except Exception as e:
            raise mEx.ShapeError("Error setting shape graphic") from e

    @classmethod
    def draw_image_offset(
        cls, slide: XDrawPage, fnm: PathOrStr, xoffset: ImageOffset | float, yoffset: ImageOffset | float
    ) -> XShape | None:
        """
        Insert the specified picture onto the slide page in the doc
        presentation document. Use the supplied (x, y) offsets to locate the
        top-left of the image.

        |lo_unsafe|

        Args:
            slide (XDrawPage): Slide
            fnm (PathOrStr): Path to image.
            xoffset (ImageOffset, float): X Offset with value between ``0.0`` and ``1.0``
            yoffset (ImageOffset, float): Y Offset with value between ``0.0`` and ``1.0``

        Returns:
            XShape | None: Shape on success, ``None`` otherwise.
        """
        xoffset = ImageOffset(float(xoffset))
        yoffset = ImageOffset(float(yoffset))
        try:
            slide_size = cls.get_slide_size(slide)
            x = round(slide_size.width * xoffset.value)  # in mm units
            y = round(slide_size.height * yoffset.value)

            max_width = slide_size.width - x
            max_height = slide_size.height - y

            im_size = mImgLo.ImagesLo.calc_scale(fnm=fnm, max_width=max_width, max_height=max_height)
            if im_size is None:
                mLo.Lo.print(f'Unable to calc image size for "{fnm}"')
                return None
            return cls._draw_image_path_x_y_w_h(
                slide=slide, fnm=fnm, x=x, y=y, width=im_size.width, height=im_size.height
            )
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Error in draw image offset") from e

    @staticmethod
    def is_image(shape: XShape) -> bool:
        """
        Gets if a shape is an image (GraphicObjectShape).

        |lo_safe|

        Args:
            shape (XShape): Shape

        Returns:
            bool: ``True`` if shape is image; Otherwise, ``False``.
        """
        if shape is None:
            return False
        return shape.getShapeType() == "com.sun.star.drawing.GraphicObjectShape"

    # endregion draw an image

    # region form manipulation
    @staticmethod
    def get_form_container(slide: XDrawPage) -> XIndexContainer | None:
        """
        Gets form container.
        The first form in slide is returned if found.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide

        Raises:
            DrawError: If error occurs.

        Returns:
            XIndexContainer | None: Form Container on success, None otherwise.
        """
        try:
            x_supp_forms = mLo.Lo.qi(XFormsSupplier, slide)
            if x_supp_forms is None:
                raise mEx.MissingInterfaceError(XFormsSupplier, "Could not access forms supplier")

            x_forms_con = x_supp_forms.getForms()
            if x_forms_con is None:
                mLo.Lo.print("Could not access forms container")
                return None

            x_forms = mLo.Lo.qi(XIndexContainer, x_forms_con, True)
            return mLo.Lo.qi(XIndexContainer, x_forms.getByIndex(0), True)
        except Exception as e:
            raise mEx.DrawError("Could not find a form") from e

    # endregion form manipulation

    # region slide show related
    @staticmethod
    def get_show(doc: XComponent) -> XPresentation2:
        """
        Gets Slide show Presentation.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawError: If error occurs.

        Returns:
            XPresentation2: Slide Show Presentation.
        """
        try:
            ps = mLo.Lo.qi(XPresentationSupplier, doc, True)
            return mLo.Lo.qi(XPresentation2, ps.getPresentation(), True)
        except Exception as e:
            raise mEx.DrawError("Unable to get Presentation") from e

    @staticmethod
    def get_show_controller(show: XPresentation2) -> XSlideShowController:
        """
        Gets slide show controller.

        |lo_safe|

        Args:
            show (XPresentation2): Slide Show Presentation.

        Raises:
            DrawError: If error occurs.

        Returns:
            XSlideShowController: Slide Show Controller.

        Note:
            It may take a little bit for the slides show to start.
            For this reason this method will wait up to five seconds.
        """
        try:
            sc = show.getController()
            # may return None if executed too quickly after start of show
            if sc is not None:
                return sc
            timeout = 5.0  # wait time in seconds
            try_sleep = 0.5
            end_time = time.time() + timeout
            while end_time > time.time():
                time.sleep(try_sleep)  # give slide show time to start
                sc = show.getController()
                if sc is not None:
                    break
        except Exception as e:
            raise mEx.DrawError("Error getting slide show controller") from e
        if sc is None:
            raise mEx.DrawError(f"Could obtain slide show controller after {timeout:.1f} seconds")
        return sc

    @staticmethod
    def wait_ended(sc: XSlideShowController) -> None:
        """
        Wait for until the slide is ended, which occurs when he user exits the slide show.

        |lo_safe|

        Args:
            sc (XSlideShowController): Slide Show Controller

        Returns:
            None:
        """
        while True:
            curr_index = sc.getCurrentSlideIndex()
            if curr_index == -1:
                break
            mLo.Lo.delay(500)

        mLo.Lo.print("End of presentation detected")

    @staticmethod
    def wait_last(sc: XSlideShowController, delay: int) -> None:
        """
        Wait for delay milliseconds when the last slide is shown before returning.

        |lo_safe|

        Args:
            sc (XSlideShowController): Slide Show Controller
            delay (int): Delay in milliseconds

        Returns:
            None:
        """
        # sourcery skip: remove-unnecessary-cast
        wait = int(delay)
        num_slides = sc.getSlideCount()
        # print(f"Number of slides: {num_slides}")
        while True:
            curr_index = sc.getCurrentSlideIndex()
            # print(f"Current index: {curr_index}")
            if curr_index == -1:
                break
            if curr_index >= num_slides - 1:
                break
            mLo.Lo.delay(500)

        if wait > 0:
            mLo.Lo.delay(wait)

    @staticmethod
    def set_transition(
        slide: XDrawPage,
        fade_effect: FadeEffect,
        speed: AnimationSpeed,
        change: DrawingSlideShowKind,
        duration: int,
    ) -> None:
        """
        Sets the transition for a slide.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide
            fade_effect (FadeEffect): Fade Effect
            speed (AnimationSpeed): Animation Speed
            change (SlideShowKind): Slide show kind
            duration (int): Duration of slide. Only used when ``change=SlideShowKind.AUTO_CHANGE``

        Raises:
            DrawPageError: If error occurs.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1DrawPage-members.html

        try:
            ps = mLo.Lo.qi(XPropertySet, slide, True)
            ps.setPropertyValue("Effect", fade_effect)
            ps.setPropertyValue("Speed", speed)
            ps.setPropertyValue("Change", int(change))
            # if change is SlideShowKind.AUTO_CHANGE
            # then Duration is used
            ps.setPropertyValue("Duration", abs(duration))  # in seconds
        except Exception as e:
            raise mEx.DrawPageError("Could not set slide transition") from e

    @staticmethod
    def get_play_list(doc: XComponent) -> XNameContainer:
        """
        Gets Play list.

        |lo_safe|

        Args:
            doc (XComponent): Document.

        Raises:
            DrawError: If error occurs.

        Returns:
            XNameContainer: Name Container.
        """
        try:
            cp_supp = mLo.Lo.qi(XCustomPresentationSupplier, doc, True)
            return cp_supp.getCustomPresentations()
        except Exception as e:
            raise mEx.DrawError("Error getting play list") from e

    @classmethod
    def build_play_list(cls, doc: XComponent, custom_name: str, *slide_idxs: int) -> XNameContainer:
        """
        Build a named play list container of  slides from doc.
        The name of the play list is ``custom_name``.

        |lo_safe|

        Args:
            doc (XComponent): Document
            custom_name (str): Name for play list
            slide_idxs (int): One or more index's of existing slides to add to play list.

        Raises:
            DrawError: If error occurs.

        Returns:
            XNameContainer: Name Container.
        """
        # get a named container for holding the custom play list
        play_list = cls.get_play_list(doc)
        try:
            # create an indexed container for the play list
            factory = mLo.Lo.qi(XSingleServiceFactory, play_list, True)
            slides_con = mLo.Lo.qi(XIndexContainer, factory.createInstance(), True)

            # container holds slide references whose indices come from slideIdxs
            mLo.Lo.print("Building play list using:")
            j = 0
            for i in slide_idxs:
                try:
                    slide = cls._get_slide_doc(doc, i)
                except IndexError as ex:
                    mLo.Lo.print(f"  Error getting slide for playlist. Skipping index {i}")
                    mLo.Lo.print(f"    {ex}")
                    continue
                slides_con.insertByIndex(j, slide)
                j += 1
                mLo.Lo.print(f"  Slide No. {i+1}, index: {i}")

            # store the play list under the custom name
            play_list.insertByName(custom_name, slides_con)
            mLo.Lo.print(f'Play list stored under the name: "{custom_name}"')
            return play_list
        except Exception as e:
            raise mEx.DrawError("Unable to build play list.") from e

    @staticmethod
    def get_animation_node(slide: XDrawPage) -> XAnimationNode:
        """
        Gets Animation Node.

        |lo_safe|

        Args:
            slide (XDrawPage): Slide.

        Raises:
            DrawPageError: If error occurs.

        Returns:
            XAnimationNode: Animation Node.
        """
        try:
            node_supp = mLo.Lo.qi(XAnimationNodeSupplier, slide, True)
            result = node_supp.getAnimationNode()
            if result is None:
                raise mEx.UnKnownError("None Value: getAnimationNode() returned None Value")
            return result
        except Exception as e:
            raise mEx.DrawPageError("Error getting animation node") from e

    # endregion slide show related

    # region helper
    @staticmethod
    def set_shape_props(shape: XShape, **props) -> None:
        """
        Sets properties on a shape.

        |lo_safe|

        Args:
            shape (XShape): Shape object
            props (Any): Key value pairs of property name and property value

        Raises:
            MissingInterfaceError: if obj does not implement XPropertySet interface
            MultiError: If unable to set a property

        Returns:
            None:

        Example:
            .. code-block:: python

                Draw.set_shape_props(shape, Loop=True, MediaURL=FileIO.fnm_to_url(fnm))
        """
        mProps.Props.set(shape, **props)

    # endregion helper


__all__ = ("Draw",)
