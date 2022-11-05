# coding: utf-8
# region Imports
from __future__ import annotations
from logging import exception
from typing import TYPE_CHECKING, Iterable, overload, Any
from enum import Enum, IntEnum
import uno

from com.sun.star.accessibility import XAccessible
from com.sun.star.accessibility import XAccessibleContext
from com.sun.star.awt import PosSize  # const
from com.sun.star.awt import WindowAttribute  # const
from com.sun.star.awt import VclWindowPeerAttribute  # const
from com.sun.star.awt import XExtendedToolkit
from com.sun.star.awt import XMenuBar
from com.sun.star.awt import XMessageBox
from com.sun.star.awt import XSystemDependentWindowPeer
from com.sun.star.awt import XToolkit
from com.sun.star.awt import XTopWindow
from com.sun.star.awt import XTopWindow2
from com.sun.star.awt import XUserInputInterception
from com.sun.star.awt import XWindow
from com.sun.star.awt import XWindow2
from com.sun.star.awt import XWindowPeer
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XIndexContainer
from com.sun.star.frame import XDispatchProviderInterception
from com.sun.star.frame import XLayoutManager
from com.sun.star.frame import XFrame
from com.sun.star.frame import XFramesSupplier
from com.sun.star.frame import XModel
from com.sun.star.lang import SystemDependent  # const
from com.sun.star.lang import XComponent
from com.sun.star.view import XControlAccess
from com.sun.star.view import XSelectionSupplier
from com.sun.star.ui import UIElementType  # const
from com.sun.star.ui import XImageManager
from com.sun.star.ui import XUIConfigurationManagerSupplier
from com.sun.star.ui import XUIConfigurationManager

if TYPE_CHECKING:
    from com.sun.star.frame import XController
    from com.sun.star.ui import XUIElement

from ..dialog import input as mInput
from ..utils import lo as mLo
from ..utils import props as mProps
from ..utils import info as mInfo
from ..utils import file_io as mFileIO
from ..utils import sys_info as m_sys_info
from ..exceptions import ex as mEx

from ooo.dyn.awt.rectangle import Rectangle
from ooo.dyn.awt.window_descriptor import WindowDescriptor
from ooo.dyn.awt.window_class import WindowClass
from ooo.dyn.view.document_zoom_type import DocumentZoomTypeEnum

# endregion Imports

SysInfo = m_sys_info.SysInfo


class GUI:

    # region Class Enums
    # view settings zoom constants
    # ZoomEnum = DocumentZoomTypeEnum
    class ZoomEnum(IntEnum):
        OPTIMAL = 0
        """
        The page content width (excluding margins) at the current selection is fit into the view.
        """
        PAGE_WIDTH = 1
        """
        The page width at the current selection is fit into the view.
        """
        ENTIRE_PAGE = 2
        """
        A complete page of the document is fit into the view.
        """
        BY_VALUE = 3
        """
        The zoom is relative and is to be set via the property ViewSettings.ZoomValue.
        """
        PAGE_WIDTH_EXACT = 4
        """
        The page width at the current selection is fit into the view, with the view ends exactly at the end of the page.
        """
        ZOOM_50_PERCENT = 1000
        """Zoom 50%"""
        ZOOM_75_PERCENT = 1001
        """Zoom 75%"""
        ZOOM_100_PERCENT = 1002
        """Zoom 100%"""
        ZOOM_150_PERCENT = 1003
        """Zoom 150%"""
        ZOOM_200_PERCENT = 1004
        """Zoom 2000%"""

    # endregion Class Enums

    # region class Constants
    MENU_BAR = "private:resource/menubar/menubar"
    STATUS_BAR = "private:resource/statusbar/statusbar"
    FIND_BAR = "private:resource/toolbar/findbar"
    STANDARD_BAR = "private:resource/toolbar/standardbar"
    TOOL_BAR = "private:resource/toolbar/toolbar"

    class ToolBarName(str, Enum):
        THREE_D_OBJECTS_BAR = "3dobjectsbar"
        ADDON_LIBRE_LOGO_OFFICE_TOOL = "addon_LibreLogo.OfficeToolBar"
        ALIGNMENT_BAR = "alignmentbar"
        ARROWS_BAR = "arrowsbar"
        ARROW_SHAPES = "arrowshapes"
        BASIC_SHAPES = "basicshapes"
        BEZIER_OBJECT_BAR = "bezierobjectbar"
        CALL_OUT_SHAPES = "calloutshapes"
        CHANGES = "changes"
        CHOOSE_MODE_BAR = "choosemodebar"
        COLOR_BAR = "colorbar"
        COMMENTS_BAR = "commentsbar"
        COMMON_TASK_BAR = "commontaskbar"
        CONNECTORS_BAR = "connectorsbar"
        CUSTOM_TOOLBAR_1 = "custom_toolbar_1"
        DATASTREAMS = "datastreams"
        DESIGN_OBJECT_BAR = "designobjectbar"
        DIALOG_BAR = "dialogbar"
        DRAW_BAR = "drawbar"
        DRAWING_OBJECT_BAR = "drawingobjectbar"
        DRAW_OBJECT_BAR = "drawobjectbar"
        DRAW_TEXT_OBJECT_BAR = "drawtextobjectbar"
        ELLIPSES_BAR = "ellipsesbar"
        EXTRUSION_OBJECT_BAR = "extrusionobjectbar"
        FIND_BAR = "findbar"
        FLOW_CHART_SHAPES = "flowchartshapes"
        FONT_WORK_OBJECT_BAR = "fontworkobjectbar"
        FONT_WORK_SHAPE_TYPE = "fontworkshapetype"
        FORMAT_OBJECT_BAR = "formatobjectbar"
        FORMATTING = "Formatting"
        FORM_CONTROLS = "formcontrols"
        FORM_CONTROLS_BAR = "formcontrolsbar"
        FORM_DESIGN = "formdesign"
        FORM_OBJECT_BAR = "formobjectbar"
        FORMS_FILTER_BAR = "formsfilterbar"
        FORMS_NAVIGATION_BAR = "formsnavigationbar"
        FORM_TEXT_OBJECT_BAR = "formtextobjectbar"
        FRAME_OBJECT_BAR = "frameobjectbar"
        FULL_SCREEN_BAR = "fullscreenbar"
        GLUE_POINTS_OBJECT_BAR = "gluepointsobjectbar"
        GRAF_FILTER_BAR = "graffilterbar"
        GRAPH_ICO_OBJECT_BAR = "graphicobjectbar"
        INSERT_BAR = "insertbar"
        INSERT_CELLS_BAR = "insertcellsbar"
        INSERT_CONTROLS_BAR = "insertcontrolsbar"
        INSERT_OBJECT_BAR = "insertobjectbar"
        LINES_BAR = "linesbar"
        MACRO_BAR = "macrobar"
        MASTER_VIEW_TOOLBAR = "masterviewtoolbar"
        MEDIA_OBJECT_BAR = "mediaobjectbar"
        MORE_FORM_CONTROLS = "moreformcontrols"
        NAVIGATION_OBJECT_BAR = "navigationobjectbar"
        NUM_OBJECT_BAR = "numobjectbar"
        OLE_OBJECT_BAR = "oleobjectbar"
        OPTIMIZE_TABLE_BAR = "optimizetablebar"
        OPTIONS_BAR = "optionsbar"
        OUTLINE_TOOLBAR = "outlinetoolbar"
        POSITION_BAR = "positionbar"
        PREVIEW_BAR = "previewbar"
        PREVIEW_OBJECT_BAR = "previewobjectbar"
        QUERY_OBJECT_BAR = "queryobjectbar"
        RECTANGLES_BAR = "rectanglesbar"
        REPORT_CONTROLS = "reportcontrols"
        REPORT_OBJECT_BAR = "reportobjectbar"
        RESIZE_BAR = "resizebar"
        SECTION_ALIGNMENT_BAR = "sectionalignmentbar"
        SECTION_SHRINK_BAR = "sectionshrinkbar"
        SLIDE_VIEW_OBJECT_BAR = "slideviewobjectbar"
        SLIDE_VIEW_TOOL_BAR = "slideviewtoolbar"
        SQL_OBJECT_BAR = "sqlobjectbar"
        STANDARD_BAR = "standardbar"
        STAR_SHAPES = "starshapes"
        SYMBOL_SHAPES = "symbolshapes"
        TABLE_OBJECT_BAR = "tableobjectbar"
        TEXT_BAR = "textbar"
        TEXT_OBJECT_BAR = "textobjectbar"
        TOOLBAR = "toolbar"
        TRANSLATION_BAR = "translationbar"
        VIEWER_BAR = "viewerbar"
        ZOOM_BAR = "zoombar"

        def __str__(self) -> str:
            return self.value

    class SpecialWindows(str, Enum):
        BASIC_IDE = "BASICIDE"
        WELCOME_SCREEN = "WELCOMESCREEN"

        def __str__(self) -> str:
            return self.value

    class WindowSubtypes(str, Enum):
        BASE_TABLE = "BASETABLE"
        BASE_QUERY = "BASEQUERY"
        BASE_REPORT = "BASEREPORT"
        BASE_DIAGRAM = "BASEDIAGRAM"

        def __str__(self) -> str:
            return self.value

    class WindowInfo:
        __slots__ = (
            "component",
            "component",
            "frame",
            "window_name",
            "window_title",
            "window_file_name",
            "document_type",
        )

        def __init__(self) -> None:
            self.component: XComponent | None = None  # com.sun.star.lang.XComponent
            self.frame: XFrame | None = None  # com.sun.star.comp.framework.Frame
            self.window_name: str = ""  # Object Name
            self.window_title: str = ""  # Only mean to identify new documents
            self.window_file_name: str = ""  # URL of file name
            self.document_type: mLo.Lo.DocType = mLo.Lo.DocType.UNKNOWN  # Writer, Calc, ...

    # endregion class Constants

    # region ---------------- toolbar addition -------------------------

    @classmethod
    def get_toobar_resource(cls, name: GUI.ToolBarName) -> str:
        """
        Get toolbar resource for name

        Args:
            name (ToolBarName): Name of resource

        Returns:
            str: A formatted resource string such as ``private:resource/toolbar/zoombar``
        """
        resource = f"private:resource/toolbar/{name}"
        return resource

    @classmethod
    def add_item_to_toolbar(cls, doc: XComponent, toolbar_name: str, item_name: str, im_fnm: str) -> None:
        """
        Add a user-defined icon and command to the start of the specified toolbar.

        Args:
            doc (XComponent): office document
            toolbar_name (str): toolbar name
            item_name (str): item name
            im_fnm (str): image file path

        """
        from com.sun.star.graphic import XGraphicProvider

        def load_graphic_file(im_fnm: str):
            # this method is also in Images module.
            # images module currently does not run as macro.
            # Pillow not needed for this method so make it local
            gprovider = mLo.Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider")
            if gprovider is None:
                return None

            file_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(im_fnm))
            return gprovider.queryGraphic(file_props)

        try:
            cmd = mLo.Lo.make_uno_cmd(item_name)
            conf_man: XUIConfigurationManager = cls.get_ui_config_manager_doc(doc)
            image_man = mLo.Lo.qi(XImageManager, conf_man.getImageManager())
            if image_man is None:
                raise mEx.MissingInterfaceError(XImageManager)
            cmds = (cmd,)
            img = load_graphic_file(im_fnm)
            if img is None:
                mLo.Lo.print(f"Unable to load graphics file: '{im_fnm}'")
                return
            pics = (img,)
            image_man.insertImages(0, cmds, pics)

            # add item to toolbar
            settings = conf_man.getSettings(toolbar_name, True)
            con_settings = mLo.Lo.qi(XIndexContainer, settings)
            if con_settings is None:
                raise mEx.MissingInterfaceError(XIndexContainer)
            item_props = mProps.Props.make_bar_item(cmd, item_name)
            con_settings.insertByIndex(0, item_props)
            conf_man.replaceSettings(toolbar_name, con_settings)
        except Exception as e:
            mLo.Lo.print(e)

    # endregion ------------- toolbar addition -------------------------

    # region ---------------- floating frame, message box --------------

    @staticmethod
    def create_floating_frame(title: str, x: int, y: int, width: int, height: int) -> XFrame:
        """
        Create a floating XFrame at the given position and size

        Args:
            title (str): Floating frame title
            x (int): Frame x position
            y (int): Frame y position
            width (int): Frame width
            height (int): Frame Height

        Raises:
            MissingInterfaceError: If required interface can not be obtained.

        Returns:
            XFrame: Floating frame
        """
        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        if xtoolkit is None:
            raise mEx.MissingInterfaceError(XToolkit)
        desc = WindowDescriptor(Type=WindowClass.TOP, WindowServiceName="modelessdialog", ParentIndex=-1)

        desc.Bounds = Rectangle(x, y, width, height)
        desc.WindowAttributes = (
            WindowAttribute.BORDER
            + WindowAttribute.MOVEABLE
            + WindowAttribute.CLOSEABLE
            + WindowAttribute.SIZEABLE
            + VclWindowPeerAttribute.CLIPCHILDREN
        )

        xwindow_peer = xtoolkit.createWindow(desc)
        window = mLo.Lo.qi(XWindow, xwindow_peer)
        if window is None:
            raise mEx.MissingInterfaceError(XWindow)

        xframe = mLo.Lo.create_instance_mcf(XFrame, "com.sun.star.frame.Frame")
        if xframe is None:
            raise mEx.MissingInterfaceError(XFrame)

        xframe.setName(title)
        xframe.initialize(window)

        xframes_sup = mLo.Lo.qi(XFramesSupplier, mLo.Lo.get_desktop())
        if xframes_sup is None:
            raise mEx.MissingInterfaceError(XFramesSupplier)

        xframes = xframes_sup.getFrames()
        if xframes is None:
            raise mEx.MissingInterfaceError(XFramesSupplier, "No desktop frames found")
        else:
            xframes.append(xframe)

        window.setVisible(True)
        return xframe

    @classmethod
    def show_message_box(cls, title: str, message: str) -> None:
        """
        Shows a message box

        Args:
            title (str): Messagebox Title
            message (str): Message to display

        Raises:
            MissingInterfaceError: If required interface is not present.
        """
        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        xwindow = cls.get_window()
        if xtoolkit is None or xwindow is None:
            return None
        xpeer = mLo.Lo.qi(XWindowPeer, xwindow)
        if xpeer is None:
            raise mEx.MissingInterfaceError(XWindowPeer)
        desc = WindowDescriptor(
            Type=WindowClass.MODALTOP,
            WindowServiceName="infobox",
            ParentIndex=-1,
            Parent=xpeer,
            Bounds=Rectangle(0, 0, 300, 200),
            WindowAttributes=WindowAttribute.BORDER | WindowAttribute.MOVEABLE | WindowAttribute.CLOSEABLE,
        )

        desc_peer = xtoolkit.createWindow(desc)
        if desc_peer is None:
            msg_box = mLo.Lo.qi(XMessageBox, desc_peer)
            if msg_box is None:
                raise mEx.MissingInterfaceError(XMessageBox)
            msg_box.CaptionText = title
            msg_box.MessageText = message
            msg_box.execute()

    @staticmethod
    def get_password(title: str, input_msg: str) -> str:
        """
        Prompts for a password.

        If inside of a Office window then an Office Dialog is displayed prompting  for a password.
        If no Office window is available and attempt is made to create a ``tkinter`` dialog window for password input.

        ``tkinter`` does not ship with integrated python in LibreOffice. For this reason it may not be possible to display a ``tkinter`` dialog.
        It will depend on how your virtual environment is set up. In most cases this will work on Linux but not on windows.

        Args:
            title (str): Title of input box
            input_msg (str): Message to display

        Raises:
            Exception if unable to build a dialog password input.

        Returns:
            str: password as string or empty string if password is not given.

        See Also:
            :py:class:`~.input.Input`
        """
        try:
            result = mInput.Input.get_input(title=title, msg=input_msg, is_password=True)
            return result
        except Exception:
            # may not be in a LibreOffice window
            pass

        # try a tkinter dialog. Not available in macro mode.
        # this also means may not work on windows when virtual environment
        # is set to LibreOffice python.exe
        try:
            from ..dialog.tk_input import Window

            pass_inst = Window(title=title, input_msg=input_msg, is_password=True)
            return pass_inst.get_input()
        except ImportError:
            pass
        raise Exception("Unable to access a GUI to create a password dialog box")

    # endregion ------------- floating frame, message box --------------

    # region ---------------- controller and frame ---------------------

    @staticmethod
    def get_current_controller(odoc: object) -> XController:
        """
        Gets controller from document

        Args:
            odoc (object): office document

        Raises:
            MissingInterfaceError: If required interface is not present.

        Returns:
            XController: controller
        """
        doc = mLo.Lo.qi(XComponent, odoc)
        if doc is None:
            raise mEx.MissingInterfaceError(XComponent)
        model = mLo.Lo.qi(XModel, doc)
        if model is None:
            raise mEx.MissingInterfaceError(XComponent)
        return model.getCurrentController()

    @classmethod
    def get_frame(cls, doc: XComponent) -> XFrame:
        """
        Gets frame from doc

        Args:
            doc (XComponent): office document

        Returns:
            XFrame: document frame.
        """
        xcontroler = cls.get_current_controller(doc)
        return xcontroler.getFrame()

    @classmethod
    def get_control_access(cls, doc: XComponent) -> XControlAccess:
        """
        Get control access from office document

        Args:
            doc (XComponent): office document

        Raises:
            MissingInterfaceError: If doc does not implement XControlAccess interface.

        Returns:
            XControlAccess: control access
        """
        ca = mLo.Lo.qi(XControlAccess, cls.get_current_controller(doc))
        if ca is None:
            raise mEx.MissingInterfaceError(XControlAccess)
        return ca

    @staticmethod
    def get_uii(doc: XComponent) -> XUserInputInterception:
        """
        Gets user input interception

        Args:
            doc (XComponent): office document

        Raises:
            MissingInterfaceError: If doc does not implement XUserInputInterception interface.

        Returns:
            XUserInputInterception: user input interception
        """
        result = mLo.Lo.qi(XUserInputInterception, doc)
        if result is None:
            raise mEx.MissingInterfaceError(XUserInputInterception)
        return result

    @classmethod
    def get_selection_supplier(cls, odoc: object) -> XSelectionSupplier:
        """
        Gets selection supplier

        Args:
            odoc (object): office document

        Raises:
            MissingInterfaceError: if odoc does not implement XComponent interface.
            MissingInterfaceError: if XSelectionSupplier interface instance is not obtained.

        Returns:
            XSelectionSupplier: Selection supplier
        """
        doc = mLo.Lo.qi(XComponent, odoc)
        if doc is None:
            raise mEx.MissingInterfaceError(XComponent, "Not an office document")
        xcontroler = cls.get_current_controller(doc)
        result = mLo.Lo.qi(XSelectionSupplier, xcontroler)
        if result is None:
            raise mEx.MissingInterfaceError(XSelectionSupplier)
        return result

    @classmethod
    def get_dpi(cls, doc: XComponent) -> XDispatchProviderInterception:
        """
        Gets Dispatch provider interception

        Args:
            doc (XComponent): office document

        Raises:
            MissingInterfaceError: if XDispatchProviderInterception interface instance is not obtained.

        Returns:
            XDispatchProviderInterception: Dispatch provider interception
        """
        xframe = cls.get_frame(doc)
        result = mLo.Lo.qi(XDispatchProviderInterception, xframe)
        if result is None:
            raise mEx.MissingInterfaceError(XDispatchProviderInterception)
        return result

    # endregion ---------------- controller and frame ------------------

    # region ---------------- Office container window ------------------
    @staticmethod
    def get_window_idenity(obj: Any) -> GUI.WindowInfo:
        """
        Gets Identity Info for window of an object

        Args:
            obj (Any): object that implements XComponent

        Returns:
            GUI.Window: Window Info
        """
        # Credits to Scriptforge for the inspiration of this method.
        win = GUI.WindowInfo()
        component = mLo.Lo.qi(XComponent, obj)
        if component is None:
            return win
        win.component = component
        try:
            implementation = mInfo.Info.get_implementation_name(obj=obj)
        except Exception as e:
            mLo.Lo.print("Unable to get implementation name:")
            mLo.Lo.print(f"  {e}")
            return win
        try:
            identifier = mInfo.Info.get_identifier(obj)
        except Exception as e:
            mLo.Lo.print(f"Unable to get identity for implementation: {implementation}")
            mLo.Lo.print(f"  {e}")
            return win
        if implementation == "com.sun.star.comp.basic.BasicIDE":
            win.window_name = str(GUI.SpecialWindows.BASIC_IDE)
        elif implementation == "com.sun.star.comp.dba.ODatabaseDocument":  # No identifier
            model = mLo.Lo.qi(XModel, obj, True)
            win.window_file_name = str(mProps.Props.get_value(name="URL", props=model.getArgs()))
            if len(win.window_file_name) > 0:
                win.window_name = str(mFileIO.FileIO.url_to_path(win.window_file_name))
            win.document_type = mLo.Lo.DocType.BASE
        elif implementation in (
            "org.openoffice.comp.dbu.ODatasourceBrowser",
            "org.openoffice.comp.dbu.OTableDesign",
            "org.openoffice.comp.dbu.OQueryDesign",
            "org.openoffice.comp.dbu.ORelationDesign",
            "com.sun.star.comp.sfx2.BackingComp",
        ):
            win.frame = component.Frame
            win.window_name = str(GUI.SpecialWindows.WELCOME_SCREEN)
        else:
            if len(identifier) > 0:
                # Do not use URL : it contains the TemplateFile when new documents are created from a template
                win.window_file_name = component.Location
                if len(win.window_file_name) > 0:
                    win.window_name = str(mFileIO.FileIO.url_to_path(win.window_file_name))
                if hasattr(obj, "Title"):
                    win.window_title = obj.Title
                if identifier in (
                    "com.sun.star.sdb.FormDesign",
                    "com.sun.star.sdb.TextReportDesign",
                    "com.sun.star.text.TextDocument",
                ):
                    win.document_type = mLo.Lo.DocType.WRITER
                elif identifier == "com.sun.star.sheet.SpreadsheetDocument":
                    win.document_type = mLo.Lo.DocType.CALC
                elif identifier == "com.sun.star.presentation.PresentationDocument":
                    win.document_type = mLo.Lo.DocType.IMPRESS
                elif identifier == "com.sun.star.drawing.DrawingDocument":
                    win.document_type = mLo.Lo.DocType.DRAW
                elif identifier == "com.sun.star.formula.FormulaProperties":
                    win.document_type = mLo.Lo.DocType.MATH
        if win.frame is None:
            if hasattr(obj, "CurrentController"):
                win.frame = obj.CurrentController.Frame

        return win

    @classmethod
    def get_active_window(cls, obj: Any) -> str:
        """
        Gets Active window as string

        Args:
            obj (Any): doc like object

        Returns:
            str: Active window as string of found; Otherwise, an empty string.

        See Also:
            :py:meth:`~.gui.GUI.get_active_window`
        """
        win = cls.get_window_idenity(obj)
        if len(win.window_file_name) > 0:
            return str(mFileIO.FileIO.url_to_path(win.window_file_name))
        if len(win.window_name) > 0:
            return win.window_name
        if len(win.window_title) > 0:
            return win.window_title
        return ""

    @classmethod
    def activate(cls, window: str | XComponent) -> None:
        """
        Activates window

        Args:
            window (str | XComponent): Window name as str or doc as XComponent

        See Also:
            - :py:meth:`~.gui.GUI.get_active_window`
            - :py:meth:`~.gui.GUI.minimize`
            - :py:meth:`~.gui.GUI.maximize`
        """
        if not isinstance(window, str):
            swin = cls.get_active_window(window)
        else:
            swin = window
        desktop = mLo.Lo.XSCRIPTCONTEXT.getDesktop()
        enm = desktop.getComponents().createEnumeration()
        while enm.hasMoreElements():
            o_comp = enm.nextElement()
            win = cls.get_window_idenity(o_comp)
            if (
                (len(win.window_file_name) > 0 and win.window_file_name == mFileIO.FileIO.fnm_to_url(swin))
                or (len(win.window_name) > 0 and win.window_name == swin)
                or (len(win.window_title) > 0 and win.window_title == swin)
            ):
                if win.frame is None:
                    break
                container = win.frame.getContainerWindow()
                if container is None:
                    break
                xwin2 = mLo.Lo.qi(XWindow2, container)
                if xwin2 is None:
                    break
                top2 = mLo.Lo.qi(XTopWindow2, container)
                if top2 is None:
                    break
                if not xwin2.isVisible():
                    xwin2.setVisible(True)
                top2.IsMinimized = False
                xwin2.setFocus()
                top2.toFront()
                break

    @overload
    @classmethod
    def get_window(cls) -> XWindow:
        """
        Gets window

        Returns:
            XWindow: window instance
        """
        ...

    @overload
    @classmethod
    def get_window(cls, doc: XComponent) -> XWindow:
        """
        Gets window

        Args:
            doc (XComponent): Ofice document

        Returns:
            XWindow: window instance
        """
        ...

    @classmethod
    def get_window(cls, doc: XComponent = None) -> XWindow | None:
        """
        Gets window

        Args:
            doc (XComponent): Office document

        Returns:
            XWindow: window instance
        """
        if doc is None:
            desktop = mLo.Lo.get_desktop()
            if desktop is None:
                return None
            frame = desktop.getCurrentFrame()
            if frame is None:
                return None
            return frame.getContainerWindow()
        else:
            xcontroller = cls.get_current_controller(doc)
            return xcontroller.getFrame().getContainerWindow()

    @overload
    @classmethod
    def set_visible(cls, is_visible: bool) -> None:
        """
        Set window visibility.

        Args:
            is_visible (bool): If True window is set visible; Otherwise, window is set invisible.
        """
        ...

    @overload
    @classmethod
    def set_visible(cls, is_visible: bool, odoc: object) -> None:
        """
        Set window visibility.

        Args:
            is_visible (bool): If True window is set visible; Otherwise, window is set invisible.
            odoc (object): office document
        """
        ...

    @classmethod
    def set_visible(cls, is_visible: bool, odoc: object = None) -> None:
        """
        Set window visibility.

        Args:
            is_visible (bool): If True window is set visible; Otherwise, window is set invisible.
            odoc (object): office document
        """
        if odoc is None:
            xwindow = cls.get_window()
        else:
            doc = mLo.Lo.qi(XComponent, odoc)
            if doc is None:
                return
            xwindow = cls.get_frame(doc).getContainerWindow()

        if xwindow is not None:
            xwindow.setVisible(is_visible)
            xwindow.setFocus()

    @classmethod
    def set_size_window(cls, doc: XComponent, width: int, height: int) -> None:
        """
        Sets window size

        Args:
            doc (XComponent): office document
            width (int): Width of window
            height (int): Height of window
        """
        xwindow = cls.get_window(doc)
        rect = xwindow.getPosSize()
        xwindow.setPosSize(rect.X, rect.Y, width, height - 30, PosSize.POSSIZE)

    @classmethod
    def set_pos_size(cls, doc: XComponent, x: int, y: int, width: int, height: int) -> None:
        """
        Sets window position and size

        Args:
            doc (XComponent): office document
            x (int): Window X position
            y (int): Window Y Position
            width (int): Window Width
            height (int): Window Height
        """
        xwindow = cls.get_window(doc)
        xwindow.setPosSize(x, y, width, height, PosSize.POSSIZE)

    @classmethod
    def get_pos_size(cls, doc: XComponent) -> Rectangle:
        """
        Gets window position and Size

        Args:
            doc (XComponent): office document

        Returns:
            Rectangle: Rectangle representing position and size
        """
        xwindow = cls.get_window(doc)
        return xwindow.getPosSize()

    @staticmethod
    def get_top_window() -> XTopWindow:
        """
        Gets top window

        Raises:
            MissingInterfaceError: If XExtendedToolkit interface can not be obtained
            MissingInterfaceError: If XTopWindow interface can not be obtained

        Returns:
            XTopWindow: top window
        """
        tk = mLo.Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
        if tk is None:
            raise mEx.MissingInterfaceError(XExtendedToolkit)
        top_win = tk.getActiveTopWindow()
        if top_win is None:
            raise mEx.MissingInterfaceError(XTopWindow)
        return top_win

    @classmethod
    def get_title_bar(cls) -> str:
        """
        Gets title bar from top window

        Raises:
            MissingInterfaceError: If XAccessible interface can not be obtained
            MissingInterfaceError: If XAccessibleContext interface can not be obtained

        Returns:
            str: title bar text
        """
        top_win = cls.get_top_window()
        acc = mLo.Lo.qi(XAccessible, top_win)
        if acc is None:
            raise mEx.MissingInterfaceError(XAccessible, "Top window not accessible")
        acc_content = acc.getAccessibleContext()
        if acc_content is None:
            raise mEx.MissingInterfaceError(XAccessibleContext)
        return acc_content.getAccessibleName()

    @staticmethod
    def get_screen_size() -> Rectangle:
        """
        Get the work area as Rectangle

        Raises:
            mEx.MissingInterfaceError: If XToolkit interface can not be obtained

        Returns:
            Rectangle: Work Area.

        Note:
            Original java method used java to get area.
            Original method seemed to return effective size. (i.e. without Windows' taskbar)

            This implementation calls ``Toolkit.getWorkArea()``.

        See also:
            `Toolkit <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1Toolkit.html>`_
        """
        tk = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        if tk is None:
            raise mEx.MissingInterfaceError(XToolkit)
        return tk.getWorkArea()

    @staticmethod
    def print_rect(r: Rectangle) -> None:
        """
        Prints a rectangle to the console

        Args:
            r (Rectangle): Rectangle to print
        """
        mLo.Lo.print(f"Rectangle: ({r.X}, {r.Y}), {r.Width} -- {r.Height}")

    @classmethod
    def get_window_handle(cls, doc: XComponent) -> int | None:
        """
        Gets handle to a window

        Args:
            doc (XComponent): document to get window handle for.

        Returns:
            int | None: handle as int on success; Otherwise, None.

        Note:
            This method was part of original java lib but was only set to work with windows.
            An attempt is made to support Linux and Mac; However, not tested at this point.

            Use this method at your own risk.
        """
        win = cls.get_window(doc)
        win_peer = mLo.Lo.qi(XSystemDependentWindowPeer, win)
        pid = tuple([0 for _ in range(8)])  # tuple of zero's
        info = SysInfo.get_platform()
        if info == SysInfo.PlatformEnum.WINDOWS:
            system_type = SystemDependent.SYSTEM_WIN32
        elif info == SysInfo.PlatformEnum.MAC:
            system_type = SystemDependent.SYSTEM_MAC
        elif info == SysInfo.PlatformEnum.LINUX:
            system_type = SystemDependent.SYSTEM_XWINDOW
        else:
            mLo.Lo.print("Unable to support, don't know this system.")
            return None
        handel = int(win_peer.getWindowHandle(pid, system_type))
        return handel

    @staticmethod
    def set_look_feel() -> None:
        """
        This method is not supported. Part of Original java lib.

        Raises:
            NotImplementedError: Not supported
        """
        raise NotImplementedError

    # endregion ------------- Office container window ------------------

    # region ---------------- min/max ----------------------------------
    @classmethod
    def maximize(cls, odoc: object) -> None:
        """
        Maximizes Office window

        Args:
            odoc (object): Office document

        See Also:
            - :py:meth:`~.gui.GUI.minimize`
            - :py:meth:`~.gui.GUI.activate`
            - :py:meth:`~.gui.GUI.get_active_window`
        """
        try:
            cls.activate(odoc)
            top_win = cls.get_top_window()
        except mEx.MissingInterfaceError as e:
            mLo.Lo.print("Unable to get top window")
            print(f"  {e}")
            return
        top2 = mLo.Lo.qi(XTopWindow2, top_win)
        if top2 is None:
            mLo.Lo.print("Unable to get top window (2)")
            return
        top2.IsMaximized = True

    @classmethod
    def minimize(cls, odoc: object) -> None:
        """
        Minimizes Office window

        Args:
            odoc (object): Office document

        See Also:
            - :py:meth:`~.gui.GUI.maximize`
            - :py:meth:`~.gui.GUI.activate`
            - :py:meth:`~.gui.GUI.get_active_window`
        """
        try:
            top_win = cls.get_top_window()
        except mEx.MissingInterfaceError as e:
            mLo.Lo.print("Unable to get top window")
            print(f"  {e}")
            return
        top2 = mLo.Lo.qi(XTopWindow2, top_win)
        if top2 is None:
            mLo.Lo.print("Unable to get top window (2)")
            return
        if top2.IsMinimized == False:
            cls.set_visible(is_visible=True, odoc=odoc)
            top2.IsMinimized = True

    # endregion ------------- min/max ----------------------------------

    # region ---------------- zooming ----------------------------------
    @overload
    @classmethod
    def zoom(cls, view: GUI.ZoomEnum) -> None:
        ...

    @overload
    @classmethod
    def zoom(cls, view: GUI.ZoomEnum, value: int) -> None:
        ...

    @overload
    @classmethod
    def zoom(cls, *, value: int = 0) -> None:
        ...

    @classmethod
    def zoom(cls, view: GUI.ZoomEnum = ZoomEnum.BY_VALUE, value: int = 0) -> None:
        """
        Sets document zoom level.

        Args:
            view (ZoomEnum): Zoom value
            value (int): The amount to zoom. :abbreviation:`eg:` 160 zooms 160%
                ``value`` has a min value of 1 and a max value of 3000. If value is out of range then 100% is used.
        """
        if view == GUI.ZoomEnum.OPTIMAL:
            mLo.Lo.dispatch_cmd("ZoomOptimal")
        elif view == GUI.ZoomEnum.PAGE_WIDTH or view == GUI.ZoomEnum.PAGE_WIDTH_EXACT:
            mLo.Lo.dispatch_cmd("ZoomPageWidth")
        elif view == GUI.ZoomEnum.ENTIRE_PAGE:
            mLo.Lo.dispatch_cmd("ZoomPage")
        elif view == GUI.ZoomEnum.ZOOM_50_PERCENT:
            mLo.Lo.dispatch_cmd("Zoom50Percent")
        elif view == GUI.ZoomEnum.ZOOM_75_PERCENT:
            mLo.Lo.dispatch_cmd("Zoom75Percent")
        elif view == GUI.ZoomEnum.ZOOM_100_PERCENT:
            mLo.Lo.dispatch_cmd("Zoom100Percent")
        elif view == GUI.ZoomEnum.ZOOM_150_PERCENT:
            mLo.Lo.dispatch_cmd("Zoom150Percent")
        elif view == GUI.ZoomEnum.ZOOM_200_PERCENT:
            mLo.Lo.dispatch_cmd("Zoom200Percent")
        elif view == GUI.ZoomEnum.BY_VALUE:
            if value <= 0 or value > 3000:
                value = 100
            p_dic = {"Zoom.Value": value, "Zoom.ValueSet": 28703, "Zoom.Type": 0}
            props = mProps.Props.make_props(**p_dic)
            mLo.Lo.dispatch_cmd(cmd="Zoom", props=props)
        else:
            mLo.Lo.print("Zoom not reconized. Aborting zoom.")
            return
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Zooming

        mLo.Lo.delay(500)

    @overload
    @classmethod
    def zoom_value(cls, value: int) -> None:
        """
        Sets document custom zoom.

        Args:
            value (int): The amount to zoom. :abbreviation:`eg:` 160 zooms 160%
        """
        ...

    @overload
    @classmethod
    def zoom_value(cls, value: int, view: GUI.ZoomEnum) -> None:
        """
        Sets document custom zoom.

        Args:
            value (int): The amount to zoom. :abbreviation:`eg:` 160 zooms 160%
                ``value`` has a min value of 1 and a max value of 3000. If value is out of range then 100% is used.
            view (ZoomEnum): Type of zoom. If ``view`` is not ``ZoomEnum.BY_VALUE`` then ``value`` is ignored. Defaults to ``ZoomEnum.BY_VALUE``.
        """
        ...

    @classmethod
    def zoom_value(cls, value: int, view: GUI.ZoomEnum = ZoomEnum.BY_VALUE) -> None:
        """
        Sets document custom zoom.

        Args:
            value (int): The amount to zoom. :abbreviation:`eg:` 160 zooms 160%
            view (ZoomEnum): Type of zoom. If ``view`` is not ``ZoomEnum.BY_VALUE`` then ``value`` is ignored. Defaults to ``ZoomEnum.BY_VALUE``.
        """
        cls.zoom(view=view, value=value)

    # endregion ------------- zooming ----------------------------------

    # region ---------------- UI config manager ------------------------

    @staticmethod
    def get_ui_config_manager(doc: XComponent) -> XUIConfigurationManager:
        """
        Gets ui config manager

        Args:
            doc (XComponent): office document

        Raises:
            MissingInterfaceError: If XModel interface can not be obtained
            MissingInterfaceError: If XUIConfigurationManagerSupplier interface can not be obtained

        Returns:
            XUIConfigurationManager: ui config manager
        """
        xmodel = mLo.Lo.qi(XModel, doc)
        if xmodel is None:
            raise mEx.MissingInterfaceError(XModel)
        xsupplier = mLo.Lo.qi(XUIConfigurationManagerSupplier, xmodel)
        if xsupplier is None:
            raise mEx.MissingInterfaceError(XUIConfigurationManagerSupplier)
        return xsupplier.getUIConfigurationManager()

    @staticmethod
    def get_ui_config_manager_doc(doc: XComponent) -> XUIConfigurationManager:
        """
        Gets ui config manager base upon doc type reported by :py:meth:`.Info.doc_type_service`.

        Args:
            doc (XComponent): office document

        Raises:
            MissingInterfaceError: If XModel interface can not be obtained
            MissingInterfaceError: If XUIConfigurationManagerSupplier interface can not be obtained
            Exception: If unable to get XUIConfigurationManager from XUIConfigurationManagerSupplier instance.

        Returns:
            XUIConfigurationManager: ui config manager
        """
        doc_type = mInfo.Info.doc_type_service(doc)

        xmodel = mLo.Lo.qi(XModel, doc)
        if xmodel is None:
            raise mEx.MissingInterfaceError(XModel)
        xsupplier = mLo.Lo.qi(XUIConfigurationManagerSupplier, xmodel)
        if xsupplier is None:
            raise mEx.MissingInterfaceError(XUIConfigurationManagerSupplier)
        try:
            return xsupplier.getUIConfigurationManager(str(doc_type))
        except Exception as e:
            raise Exception(f"Could not create a config manager using '{doc_type}'") from e

    # region print_ui_cmds()

    @overload
    @classmethod
    def print_ui_cmds(cls, ui_elem_name: str, config_man: XUIConfigurationManager) -> None:
        """
        Prints ui elements matching ``ui_elem_name`` to console.

        Args:
            ui_elem_name (str): Name of ui element
            config_man (XUIConfigurationManager): configuration manager
        """
        ...

    @overload
    @classmethod
    def print_ui_cmds(cls, ui_elem_name: str, doc: XComponent) -> None:
        """
        Prints ui elements matching ``ui_elem_name`` to console.

        Args:
            ui_elem_name (str): Name of ui element
            doc (XComponent): office document
        """
        ...

    @classmethod
    def print_ui_cmds(cls, *args, **kwargs) -> None:
        """
        Prints ui elements matching ``ui_elem_name`` to console.

        Args:
            ui_elem_name (str): Name of ui element
            config_man (XUIConfigurationManager): configuration manager
            doc (XComponent): office document
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("ui_elem_name", "config_man", "doc")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("print_ui_cmds() got an unexpected keyword argument")
            ka[1] = kwargs.get("ui_elem_name", None)
            keys = ("config_man", "doc")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("print_ui_cmds() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        obj = mLo.Lo.qi(XUIConfigurationManager, kargs[2])
        if obj is None:
            cls._print_ui_cmds2(ui_elem_name=kargs[1], doc=kargs[2])
        else:
            cls._print_ui_cmds1(ui_elem_name=kargs[1], config_man=kargs[2])

    @staticmethod
    def _print_ui_cmds1(
        ui_elem_name: str,
        config_man: XUIConfigurationManager,
    ) -> None:
        """print every command used by the toolbar whose resource name is uiElemName"""
        # see Also: https://wiki.openoffice.org/wiki/Documentation/DevGuide/ProUNO/Properties
        try:
            settings = config_man.getSettings(ui_elem_name, True)
            num_settings = settings.getCount()
            print(f"No. of slements in '{ui_elem_name}' toolbar: {num_settings}")

            for i in range(num_settings):
                # line from java
                # PropertyValue[] settingProps =  Lo.qi(PropertyValue[].class, settings.getByIndex(i));
                setting_props = mLo.Lo.qi(XPropertySet, settings.getByIndex(i))
                val = mProps.Props.get_value(name="CommandURL", props=setting_props)
                print(f"{i}) {mProps.Props.prop_value_to_string(val)}")
            print()
        except exception as e:
            print(e)

    @classmethod
    def _print_ui_cmds2(cls, ui_elem_name: str, doc: XComponent) -> None:
        config_man = cls.get_ui_config_manager(doc)
        if config_man is None:
            print("Cannot create configuration manager")
            return
        cls.print_ui_cmds()

    # endregion print_ui_cmds()

    # endregion ------------- UI config manager ------------------------

    # region ---------------- layout manager ---------------------------

    # region    get_layout_manager()
    @overload
    @classmethod
    def get_layout_manager(cls) -> XLayoutManager:
        """
        Gets layout manager

        Args:
            doc (XComponent): office document

        Returns:
            XLayoutManager: Layout manager
        """
        ...

    @overload
    @classmethod
    def get_layout_manager(cls, doc: XComponent) -> XLayoutManager:
        """
        Gets layout manager

        Args:
            doc (XComponent): office document

        Raises:
            Exception: If unable to get layout manager

        Returns:
            XLayoutManager: Layout manager
        """
        ...

    @classmethod
    def get_layout_manager(cls, doc: XComponent = None) -> XLayoutManager:
        """
        Gets layout manager

        Args:
            doc (XComponent): office document

        Raises:
            Exception: If unable to get layout manager

        Returns:
            XLayoutManager: Layout manager
        """
        try:
            if doc is None:
                desktop = mLo.Lo.get_desktop()
                frame = desktop.getCurrentFrame()
            else:
                frame = cls.get_frame(doc)

            if frame is None:
                raise Exception("No current frame")

            lm = None
            prop_set = mLo.Lo.qi(XPropertySet, frame)
            lm = mLo.Lo.qi(XLayoutManager, prop_set.getPropertyValue("LayoutManager"))
            if lm is None:
                raise mEx.MissingInterfaceError(XLayoutManager)
            return lm
        except Exception as e:
            raise Exception("Could not access layout manager") from e

    # endregion    get_layout_manager()

    # region    print_u_is()
    @overload
    @classmethod
    def print_u_is(cls) -> None:
        """print the resource names of every toolbar used by desktop"""
        ...

    @overload
    @classmethod
    def print_u_is(cls, doc: XComponent) -> None:
        """
        Print to console the resource names of every toolbar used by doc

        Args:
            doc (XComponent): office document
        """
        ...

    @overload
    @classmethod
    def print_u_is(cls, lm: XLayoutManager) -> None:
        """
        Print to console the resource names of every toolbar used by layout manager

        Args:
            lm (XLayoutManager): Layout manager
        """
        ...

    @classmethod
    def print_u_is(cls, *args, **kwargs) -> None:
        """
        Print to console the resource names of every toolbar used by doc

        Args:
            lm (XLayoutManager): Layout manager
            doc (XComponent): office document
        """
        ordered_keys = ("first",)
        kargs = {}
        if "doc" in kwargs:
            kargs["first"] = kwargs["doc"]
        elif "lm" in kwargs:
            kargs["first"] = kwargs["lm"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)
        if k_len > 1:
            print("invalid number of arguments for print_u_is()")
            return
        if k_len == 0:
            lm = cls.get_layout_manager()
        else:
            obj = mLo.Lo.qi(XLayoutManager, kargs["first"])
            if obj is None:
                lm = cls.get_layout_manager(kargs["first"])
            else:
                lm = kargs["first"]
        if lm is None:
            print("No layout manager found")
            return
        ui_elems = lm.getElements()
        print(f"No. of UI Elemtnts: {len(ui_elems)}")
        for el in ui_elems:
            print(f"  {el.ResourceURL}; {cls.get_ui_element_type_str(el.Type)}")
        print()

    # endregion print_u_is()

    @staticmethod
    def get_ui_element_type_str(t: int) -> str:
        """
        Converts constant value to element type string

        `UIElementType <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1ui_1_1UIElementType.html>`_
        determines the type of a user interface element which is controlled by a layout manager.

        Args:
            t (int): UIElementType constant Value from 0 to 8

        Raises:
            TypeError: If t is not a int
            ValueError: If t is not a valid UIElementType constant.

        Returns:
            str: element type string
        """
        if not isinstance(t, int):
            raise TypeError("'t' is not an int")
        if t == UIElementType.UNKNOWN:
            return "unknown"
        if t == UIElementType.MENUBAR:
            return "menubar"
        if t == UIElementType.POPUPMENU:
            return "popup menu"
        if t == UIElementType.TOOLBAR:
            return "toolbar"
        if t == UIElementType.STATUSBAR:
            return "status bar"
        if t == UIElementType.FLOATINGWINDOW:
            return "floating window"
        if t == UIElementType.PROGRESSBAR:
            return "progress bar"
        if t == UIElementType.TOOLPANEL:
            return "tool panel"
        if t == UIElementType.DOCKINGWINDOW:
            return "docking window"
        if t == UIElementType.COUNT:
            return "count"

        raise ValueError("'t' is is not a valid UIElementType value")

    @classmethod
    def printAllUICommands(cls, doc: XComponent) -> None:
        """
        Prints all ui commands to console

        Args:
            doc (XComponent): office document
        """
        conf_man = cls.get_ui_config_manager_doc(doc)
        if conf_man is None:
            print("No configuration manager found")
            return
        lm = cls.get_layout_manager(doc)
        if lm is None:
            print("No layout manager found")
            return
        ui_elmes = lm.getElements()
        print(f"No. of UI Elements: {len(ui_elmes)}")
        for el in ui_elmes:
            name = el.ResourceURL
            print(f"--- {name} ---")
            cls._print_ui_cmds1(ui_elem_name=name, config_man=conf_man)

    @classmethod
    def show_one(cls, doc: XComponent, show_elem: str) -> None:
        """
        Leave only the single specified toolbar visible

        Args:
            doc (XComponent): office document
            show_elem (str): name of element to show only.
        """
        show_elems = [show_elem]
        cls.show_only(doc=doc, show_elems=[show_elems])

    @classmethod
    def show_only(cls, doc: XComponent, show_elems: Iterable[str]) -> None:
        """
        Leave only the specified toolbars visible

        Raises:
            Exception: if unable to get layout manager from doc

        Args:
            doc (XComponent): office document
            show_elems (Iterable[str]): Elements to show
        """
        lm = cls.get_layout_manager(doc)
        ui_elmes = lm.getElements()
        cls.hide_except(lm=lm, ui_elms=ui_elmes, show_elems=show_elems)

        for el_name in show_elems:  # these elems are not in lm
            lm.createElement(el_name)  # so need to be created & shown
            lm.showElement(el_name)
            mLo.Lo.print(f"{el_name} made visible")

    @staticmethod
    def hide_except(lm: XLayoutManager, ui_elms: Iterable[XUIElement], show_elms: Iterable[str]) -> None:
        """
        Hide all of ``ui_elms``, except ones in ``show_elms``;
        delete any strings that match in ``show_elms``.

        Args:
            lm (XLayoutManager): Layout Manager
            ui_elms (Iterable[XUIElement]): Elements
            show_elms (Iterable[str]): elements to show
        """
        for ui_elm in ui_elms:
            el_name = ui_elm.ResourceURL
            to_hide = True
            # show_elms_lst = list(show_elms)
            for el in show_elms:
                if el == el_name:
                    show_elms.remove(el)  # this elem is in lm so remove from show_elems
                    to_hide = False
                    break
            if to_hide:
                lm.hideElement(el_name)
                mLo.Lo.print(f"{el_name} hidden")

    @classmethod
    def show_none(cls, doc: XComponent) -> None:
        """
        Make all the toolbars invisible

        Raises:
            Exception: if unable to get layout manager from doc

        Args:
            doc (XComponent): office document.
        """
        lm = cls.get_layout_manager(doc)
        if lm is None:
            mLo.Lo.print("No layout manager found")
            return
        ui_elms = lm.getElements()
        for ui_elm in ui_elms:
            elem_name = ui_elm.ResourceURL
            lm.hideElement(elem_name)
            mLo.Lo.print(f"{elem_name} hidden")

    # endregion ------------- layout manager ---------------------------

    # region ---------------- menu bar ---------------------------------

    @classmethod
    def get_menubar(cls, lm: XLayoutManager) -> XMenuBar:
        """
        Get menu bar

        Args:
            lm (XLayoutManager): layout manager

        Raises:
            TypeError: If lm is None
            MissingInterfaceError: If a required interface cannot be obtained.

        Returns:
            XMenuBar: menu bar
        """
        if lm is None:
            raise TypeError("'lm' is None. No layout manager available for menu discovery")

        omenu_bar = lm.getElement(cls.MENU_BAR)
        props = mLo.Lo.qi(XPropertySet, omenu_bar)
        if bar is None:
            raise mEx.MissingInterfaceError(XPropertySet)

        bar = mLo.Lo.qi(XMenuBar, props.getPropertyValue("XMenuBar"))
        # the XMenuBar reference is a property of the menubar UI
        if bar is None:
            raise mEx.MissingInterfaceError(XMenuBar)

    @classmethod
    def get_menu_max_id(cls, bar: XMenuBar) -> int:
        """
        Scan through the IDs used by all the items in
        this menubar, and return the largest ID encountered.

        Args:
            bar (XMenuBar): Menu bar

        Returns:
            int: Largest menu bar id if found; Otherwise, -1
        """
        if bar is None:
            return -1

        item_count = bar.getItemCount()
        max_id = -1
        for i in range(item_count):
            id = bar.getItemId(i)
            if id > max_id:
                max_id = id

        return max_id

    # endregion ------------- menu bar ---------------------------------
