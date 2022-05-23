# coding: utf-8
# region Imports
from __future__ import annotations
from logging import exception
import sys
from typing import TYPE_CHECKING, Iterable, List, cast, overload
from enum import IntEnum
import uno

from com.sun.star.accessibility import XAccessible
from com.sun.star.awt import PosSize # const
from com.sun.star.awt import Rectangle # struct
from com.sun.star.awt import WindowAttribute # const
from com.sun.star.awt import VclWindowPeerAttribute # const
from com.sun.star.awt import WindowDescriptor # struct
from com.sun.star.awt import XExtendedToolkit
from com.sun.star.awt import XMenuBar
from com.sun.star.awt import XMessageBox
from com.sun.star.awt import XSystemDependentWindowPeer
from com.sun.star.awt import XToolkit
from com.sun.star.awt import XUserInputInterception
from com.sun.star.awt import XWindow
from com.sun.star.awt import XWindowPeer
from com.sun.star.awt.WindowClass import TOP as WC_TOP, MODALTOP as WC_MODALTOP
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XIndexContainer
from com.sun.star.frame import XDispatchProviderInterception
from com.sun.star.frame import XLayoutManager
from com.sun.star.frame import XFrame
from com.sun.star.frame import XFramesSupplier
from com.sun.star.frame import XModel
from com.sun.star.lang import SystemDependent # const
from com.sun.star.lang import XComponent
from com.sun.star.view import DocumentZoomType # const
from com.sun.star.view import XControlAccess
from com.sun.star.view import XSelectionSupplier
from com.sun.star.ui import UIElementType # const
from com.sun.star.ui import XImageManager
from com.sun.star.ui import XUIConfigurationManagerSupplier
from com.sun.star.ui import XUIConfigurationManager

if TYPE_CHECKING:
    # from com.sun.star.awt import Toolkit
    from com.sun.star.frame import XController
    from com.sun.star.awt import XTopWindow
    from com.sun.star.ui import XUIElement

from ..utils import lo as mLo
from ..utils import images as mImages
from ..utils import props as mProps
from ..utils import info as mInfo
from ..utils import sys_info as m_sys_info
# endregion Imports

SysInfo = m_sys_info.SysInfo

# if sys.version_info >= (3, 10):
#     from typing import Union
# else:
#     from typing_extensions import Union


class GUI:

    # region Class Enums
    # view settings zoom constants
    class ZoomEnum(IntEnum):
        OPTIMAL = DocumentZoomType.OPTIMAL
        PAGE_WIDTH = DocumentZoomType.PAGE_WIDTH
        ENTIRE_PAGE = DocumentZoomType.ENTIRE_PAGE
        BY_VALUE = DocumentZoomType.BY_VALUE
        PAGE_WIDTH_EXACT = DocumentZoomType.PAGE_WIDTH_EXACT
    # endregion Class Enums

    # region class Constants
    MENU_BAR = "private:resource/menubar/menubar"
    STATUS_BAR = "private:resource/statusbar/statusbar"
    FIND_BAR = "private:resource/toolbar/findbar"
    STANDARD_BAR = "private:resource/toolbar/standardbar"
    TOOL_BAR = "private:resource/toolbar/toolbar"

    TOOBAR_NMS = (
        "3dobjectsbar",
        "addon_LibreLogo.OfficeToolBar",
        "alignmentbar",
        "arrowsbar",
        "arrowshapes",
        "basicshapes",
        "bezierobjectbar",
        "calloutshapes",
        "changes",
        "choosemodebar",
        "colorbar",
        "commentsbar",
        "commontaskbar",
        "connectorsbar",
        "custom_toolbar_1",
        "datastreams",
        "designobjectbar",
        "dialogbar",
        "drawbar",
        "drawingobjectbar",
        "drawobjectbar",
        "drawtextobjectbar",
        "ellipsesbar",
        "extrusionobjectbar",
        "findbar",
        "flowchartshapes",
        "fontworkobjectbar",
        "fontworkshapetype",
        "formatobjectbar",
        "Formatting",
        "formcontrols",
        "formcontrolsbar",
        "formdesign",
        "formobjectbar",
        "formsfilterbar",
        "formsnavigationbar",
        "formtextobjectbar",
        "frameobjectbar",
        "fullscreenbar",
        "gluepointsobjectbar",
        "graffilterbar",
        "graphicobjectbar",
        "insertbar",
        "insertcellsbar",
        "insertcontrolsbar",
        "insertobjectbar",
        "linesbar",
        "macrobar",
        "masterviewtoolbar",
        "mediaobjectbar",
        "moreformcontrols",
        "navigationobjectbar",
        "numobjectbar",
        "oleobjectbar",
        "optimizetablebar",
        "optionsbar",
        "outlinetoolbar",
        "positionbar",
        "previewbar",
        "previewobjectbar",
        "queryobjectbar",
        "rectanglesbar",
        "reportcontrols",
        "reportobjectbar",
        "resizebar",
        "sectionalignmentbar",
        "sectionshrinkbar",
        "slideviewobjectbar",
        "slideviewtoolbar",
        "sqlobjectbar",
        "standardbar",
        "starshapes",
        "symbolshapes",
        "tableobjectbar",
        "textbar",
        "textobjectbar",
        "toolbar",
        "translationbar",
        "viewerbar",
        "zoombar",
    )

    # endregion class Constants

    # region ---------------- toolbar addition -------------------------

    @classmethod
    def get_toobar_resource(cls, nm: str) -> str | None:
        for res_nm in cls.TOOBAR_NMS:
            if res_nm in nm:
                resource = f"private:resource/toolbar/{nm}"
                print(f"Matched {nm} to {resource}")
                return resource
        return None

    @classmethod
    def add_item_to_toolbar(
        cls, doc: XComponent, toolbar_name: str, item_name: str, im_fnm: str
    ) -> None:
        """
        Add a user-defined icon and command to the start of the specified toolbar.
        """
        cmd = mLo.Lo.make_uno_cmd(item_name)
        conf_man: XUIConfigurationManager = cls.get_ui_config_manager_doc(doc)
        if conf_man is None:
            print("Cannot create configuration manager")
            return
        try:
            image_man = mLo.Lo.qi(XImageManager, conf_man.getImageManager())
            cmds = (cmd,)
            img = mImages.Images.load_graphic_file(im_fnm)
            if img is None:
                print(f"Unable to load graphics file: '{im_fnm}'")
                return
            pics = (img,)
            image_man.insertImages(0, cmds, pics)

            # add item to toolbar
            settings = conf_man.getSettings(toolbar_name, True)
            con_settings = mLo.Lo.qi(XIndexContainer, settings)
            if con_settings is None:
                raise TypeError("con_settings is None")
            item_props = mProps.Props.make_bar_item(cmd, item_name)
            con_settings.insertByIndex(0, item_props)
            conf_man.replaceSettings(toolbar_name, con_settings)
        except Exception as e:
            print(e)

    # endregion ------------- toolbar addition -------------------------

    # region ---------------- floating frame, message box --------------

    @staticmethod
    def create_floating_frame(
        title: str, x: int, y: int, width: int, height: int
    ) -> XFrame | None:
        """create a floating XFrame at the given position and size"""
        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        if xtoolkit is None:
            return None
        desc = WindowDescriptor()
        desc.Type = WC_TOP
        desc.WindowServiceName = "modelessdialog"
        desc.ParentIndex = -1

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
            print("Could not create window")
            return None

        xframe = mLo.Lo.create_instance_mcf(XFrame, "com.sun.star.frame.Frame")
        if xframe is None:
            print("Could not create frame")
            return None

        xframe.setName(title)
        xframe.initialize(window)

        xframes_sup = mLo.Lo.qi(XFramesSupplier, mLo.Lo.get_desktop())
        xframes = xframes_sup.getFrames()
        if xframes is None:
            print("Mo desktop frames found")
        else:
            xframes.append(xframe)

        window.setVisible(True)
        return xframe

    @classmethod
    def show_message_box(cls, title: str, message: str) -> None:
        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        xwindow = cls.get_window()
        if xtoolkit is None or xwindow is None:
            return None
        xpeer = mLo.Lo.qi(XWindowPeer, xwindow)
        desc = WindowDescriptor()
        desc.Type = WC_MODALTOP
        desc.WindowServiceName = "infobox"
        desc.ParentIndex = -1
        desc.Parent = xpeer
        desc.Bounds = Rectangle(0, 0, 300, 200)
        desc.WindowAttributes = (
            WindowAttribute.BORDER
            | WindowAttribute.MOVEABLE
            | WindowAttribute.CLOSEABLE
        )

        desc_peer = xtoolkit.createWindow(desc)
        if desc_peer is None:
            msg_box = mLo.Lo.qi(XMessageBox, desc_peer)
            if msg_box is not None:
                msg_box.CaptionText = title
                msg_box.MessageText = message
                msg_box.execute()

    @staticmethod
    def get_password(title: str, input_msg: str) -> str:
        """
        Prompts for a password.

        Currently Not Implemented.

        Args:
            title (str): Title of input box
            input_msg (str): Message to display

        Raises:
            NotImplementedError: Not yet implemented

        Returns:
            str: password as string.

        ToDo:
            Implement the get_password method.
        """
        raise NotImplementedError
        # in original java this was done by creating a input box with a password field
        # this could likely be done with LibreOffice API, create input box and set input as password field

    # endregion ------------- floating frame, message box --------------

    # region ---------------- controller and frame ---------------------

    @staticmethod
    def get_current_controller(odoc: object) -> XController | None:
        doc = mLo.Lo.qi(XComponent, odoc)
        model = mLo.Lo.qi(XModel, doc)
        if model is None:
            print("Document has no data model")
            return None
        return model.getCurrentController()

    @classmethod
    def get_frame(cls, doc: XComponent) -> XFrame | None:
        xcontroler = cls.get_current_controller(doc)
        if xcontroler is None:
            return None
        return xcontroler.getFrame()

    @staticmethod
    def get_control_access(doc: XComponent) -> XControlAccess | None:
        return mLo.Lo.qi(XControlAccess, doc)

    @staticmethod
    def get_uii(doc: XComponent) -> XUserInputInterception | None:
        return mLo.Lo.qi(XUserInputInterception, doc)

    @classmethod
    def get_selection_supplier(cls, odoc: object) -> XSelectionSupplier | None:
        doc = mLo.Lo.qi(XComponent, odoc)
        if doc is None:
            return None
        xcontroler = cls.get_current_controller(doc)
        if xcontroler is None:
            return None
        return mLo.Lo.qi(XSelectionSupplier, xcontroler)

    @classmethod
    def get_dpi(cls, doc: XComponent) -> XDispatchProviderInterception | None:
        xframe = cls.get_frame(doc)
        if xframe is None:
            return None
        return mLo.Lo.qi(XDispatchProviderInterception, xframe)

    # endregion ---------------- controller and frame ------------------

    # region ---------------- Office container window ------------------
    @overload
    @staticmethod
    def get_window() -> XWindow | None:
        ...

    @overload
    @staticmethod
    def get_window(doc: XComponent) -> XWindow | None:
        ...

    @classmethod
    def get_window(cls, doc: XComponent = None) -> XWindow | None:
        if doc is None:
            desktop = mLo.Lo.get_desktop()
            frame = desktop.getCurrentFrame()
            if frame is None:
                print("No current frame")
                return None
            return frame.getContainerWindow()
        else:
            xcontroller = cls.get_current_controller(doc)
            if xcontroller is None:
                return None
            return xcontroller.getFrame().getContainerWindow()

    @overload
    @staticmethod
    def set_visible(is_visible: bool) -> None:
        ...

    @overload
    @staticmethod
    def set_visible(is_visible: bool, odoc: object) -> None:
        ...

    @classmethod
    def set_visible(cls, is_visible: bool, odoc: object = None) -> None:
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
        xwindow = cls.get_window(doc)
        if xwindow is None:
            return
        rect = xwindow.getPosSize()
        xwindow.setPosSize(rect.X, rect.Y, width, height - 30, PosSize.POSSIZE)

    @classmethod
    def set_pos_size(
        cls, doc: XComponent, x: int, y: int, width: int, height: int
    ) -> None:
        xwindow = cls.get_window(doc)
        if xwindow is None:
            return
        xwindow.setPosSize(x, y, width, height, PosSize.POSSIZE)

    @classmethod
    def get_pos_size(cls, doc: XComponent) -> Rectangle | None:
        xwindow = cls.get_window(doc)
        if xwindow is None:
            return
        return xwindow.getPosSize()

    @staticmethod
    def get_top_window() -> XTopWindow | None:
        tk = mLo.Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
        if tk is None:
            print("Toolkit not found")
            return None
        top_win = tk.getActiveTopWindow()
        if top_win is None:
            print("Could not find top window")
            return None
        return top_win

    @classmethod
    def get_title_bar(cls) -> str | None:
        top_win = cls.get_top_window()
        if top_win is None:
            return None
        acc = mLo.Lo.qi(XAccessible, top_win)
        if acc is None:
            print("Top window not accessible")
            return None
        acc_content = acc.getAccessibleContext()
        if acc_content is None:
            return None
        return acc_content.getAccessibleName()

    @staticmethod
    def get_screen_size() -> Rectangle:
        """
        Get the work area as Rectangle

        Returns:
            Rectangle: Work Area.

        Notes:
            Original java method used java to get area.
            Original method seemed to return effective size. (i.e. without Windows' taskbar)

            This implemention calls Toolkit.getWorkArea().

            See also: `Toolkit <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1Toolkit.html>`_
        """
        tk = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        if tk is None:
            print("Toolkit not found")
            return None
        return tk.getWorkArea()

    @staticmethod
    def print_rect(r: Rectangle) -> None:
        print(f"Rectangle: ({r.X}, {r.Y}), {r.Width} -- {r.Height}")

    @classmethod
    def get_window_handle(cls, doc: XComponent) -> int | None:
        """
        Gets Handel to a window

        Args:
            doc (XComponent): document to get window handel for.

        Returns:
            int | None: handel as int on success; Otherwise, None.

        Notes:
            This method was part of original java lib but was only set to work with windows.
            An attemp is made to support Linux and Mac; However, not tested at this point.

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
            print("Unable to support, don't know this system.")
            return None
        handel = int(win_peer.getWindowHandle(pid, system_type))

    @staticmethod
    def set_look_feel() -> None:
        """
        This method is not supported. Part of Original java lib.

        Raises:
            NotImplementedError: Not supported
        """
        raise NotImplementedError

    # endregion ------------- Office container window ------------------

    # region ---------------- zooming ----------------------------------

    @classmethod
    def zoom(cls, view: int | ZoomEnum) -> None:
        """
        Sets document zoom level.

        Args:
            view (int | ZoomEnum): Zoom value
        """
        if view == cls.ZoomEnum.OPTIMAL:
            mLo.Lo.dispatch_cmd("ZoomOptimal")
        elif view == cls.ZoomEnum.PAGE_WIDTH:
            mLo.Lo.dispatch_cmd("ZoomPageWidth")
        elif view == cls.ZoomEnum.ENTIRE_PAGE:
            mLo.Lo.dispatch_cmd("ZoomPage")
        else:
            print(f"Did not recognize zoom view: {view}; using optimal")
            mLo.Lo.dispatch_cmd("ZoomOptimal")
        mLo.Lo.delay(500)

    @classmethod
    def zoom_value(cls, value: int, view: int | ZoomEnum = ZoomEnum.BY_VALUE) -> None:
        """
        Sets document custom zoom.

        Args:
            value (int): The amount to zoom. Eg: 160 zooms 160%
            view (int | ZoomEnum, optional): Type of zoom. If 'view' is not 'ZoomEnum.BY_VALUE' then 'value' is ignored. Defaults to ZoomEnum.BY_VALUE.
        """
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Zooming
        p_dic = {"Zoom.Value": 0, "Zoom.ValueSet": 28703, "Zoom.Type": int(view)}
        if view == cls.ZoomEnum.BY_VALUE:
            p_dic["Zoom.Value"] = value

        props = mProps.Props.make_props(**p_dic)
        mLo.Lo.dispatch_cmd(cmd="Zoom", props=props)
        mLo.Lo.delay(500)

    # endregion ------------- zooming ----------------------------------

    # region ---------------- UI config manager ------------------------

    @staticmethod
    def get_ui_config_manager(doc: XComponent) -> XUIConfigurationManager | None:
        xmodel = mLo.Lo.qi(XModel, doc)
        if xmodel is None:
            print("No XModel interface")
            return None
        xsupplier = mLo.Lo.qi(XUIConfigurationManagerSupplier, xmodel)
        if xsupplier is None:
            print("No XUIConfigurationManagerSupplier interface")
            return None
        return xsupplier.getUIConfigurationManager()

    @staticmethod
    def get_ui_config_manager_doc(doc: XComponent) -> XUIConfigurationManager | None:
        doc_type = mInfo.Info.doc_type_string(doc)

        xmodel = mLo.Lo.qi(XModel, doc)
        if xmodel is None:
            print("No XModel interface")
            return None
        xsupplier = mLo.Lo.qi(XUIConfigurationManagerSupplier, xmodel)
        if xsupplier is None:
            print("No XUIConfigurationManagerSupplier interface")
            return None
        config_man = None
        try:
            config_man = xsupplier.getUIConfigurationManager(doc_type)
        except Exception as e:
            print(f"Could not create a config manager using '{doc_type}'")
            print(f"    {e}")
        return config_man

    # region print_ui_cmds()

    @overload
    @staticmethod
    def print_ui_cmds(ui_elem_name: str, config_man: XUIConfigurationManager) -> None:
        ...

    @overload
    @staticmethod
    def print_ui_cmds(ui_elem_name: str, doc: XComponent) -> None:
        ...

    @classmethod
    def print_ui_cmds(cls, *args, **kwargs) -> None:
        ordered_keys = ("first", "second")
        kargs = {}
        kargs["first"] = kwargs.get("ui_elem_name", None)
        if "config_man" in kwargs:
            kargs["second"] = kwargs["config_man"]
        elif "doc" in kwargs:
            kargs["second"] = kargs["doc"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if len(kargs) != 2:
            print("invalid number of arguments for print_ui_cmds()")
            return
        obj = mLo.Lo.qi(XUIConfigurationManager, kargs["second"])
        if obj is None:
            cls._print_ui_cmds2(ui_elem_name=kargs["first"], doc=kargs["second"])
        else:
            cls._print_ui_cmds1(ui_elem_name=kargs["first"], config_man=kargs["second"])

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
    @staticmethod
    def get_layout_manager() -> XLayoutManager | None:
        ...

    @overload
    @staticmethod
    def get_layout_manager(doc: XComponent) -> XLayoutManager | None:
        ...

    @classmethod
    def get_layout_manager(cls, doc: XComponent = None) -> XLayoutManager | None:
        if doc is None:
            desktop = mLo.Lo.get_desktop()
            frame = desktop.getCurrentFrame()
        else:
            frame = cls.get_frame(doc)

        if frame is None:
            print("No current frame")
            return None

        lm = None
        try:
            prop_set = mLo.Lo.qi(XPropertySet, frame)
            lm = mLo.Lo.qi(XLayoutManager, prop_set.getPropertyValue("LayoutManager"))
        except Exception:
            print("Could not access layout manager")
        return lm

    # endregion    get_layout_manager()

    # region    print_u_is()
    @overload
    @staticmethod
    def print_u_is() -> None:
        """print the resource names of every toolbar used by desktop"""
        ...

    @overload
    @staticmethod
    def print_u_is(doc: XComponent) -> None:
        """print the resource names of every toolbar used by doc"""
        ...

    @overload
    @staticmethod
    def print_u_is(lm: XLayoutManager) -> None:
        """print the resource names of every toolbar"""
        ...

    @classmethod
    def print_u_is(cls, *args, **kwargs) -> None:
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
        default = "??"
        if not isinstance(t, int):
            return default
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

        return default

    @classmethod
    def printAllUICommands(cls, doc: XComponent) -> None:
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
        """leave only the single specified toolbar visible"""
        show_elems = [show_elem]
        cls.show_only(doc=doc, show_elems=show_elems)

    @classmethod
    def show_only(cls, doc: XComponent, show_elems: List[str]) -> None:
        """leave only the specified toolbars visible"""
        lm = cls.get_layout_manager(doc)
        if lm is None:
            print("No layout manager found")
            return
        ui_elmes = lm.getElements()
        cls.hide_except(lm=lm, ui_elms=ui_elmes, show_elems=show_elems)

        for el_name in show_elems:  # these elems are not in lm
            lm.createElement(el_name)  # so need to be created & shown
            lm.showElement(el_name)
            print(f"{el_name} made visible")

    @staticmethod
    def hide_except(
        lm: XLayoutManager, ui_elms: Iterable[XUIElement], show_elms: List[str]
    ) -> None:
        """
        hide all of uiElems, except ones in show_elms;
        delete any strings that match in show_elms
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
                print(f"{el_name} hidden")

    @classmethod
    def show_none(cls, doc: XComponent) -> None:
        """make all the toolbars invisible"""
        lm = cls.get_layout_manager(doc)
        if lm is None:
            print("No layout manager found")
            return
        ui_elms = lm.getElements()
        for ui_elm in ui_elms:
            elem_name = ui_elm.ResourceURL
            lm.hideElement(elem_name)
            print(f"{elem_name} hidden")

    # endregion ------------- layout manager ---------------------------

    # region ---------------- menu bar ---------------------------------

    @classmethod
    def get_menubar(cls, lm: XLayoutManager) -> XMenuBar | None:
        if lm is None:
            print("No layout manager available for menu discovery")
            return None

        bar = None
        try:
            omenu_bar = lm.getElement(cls.MENU_BAR)
            props = mLo.Lo.qi(XPropertySet, omenu_bar)

            bar = mLo.Lo.qi(XMenuBar, props.getPropertyValue("XMenuBar"))
            # the XMenuBar reference is a property of the menubar UI
            if bar is None:
                print("Menubar reference not found")
        except Exception:
            print("Could not access menubar")
        return bar

    @classmethod
    def get_menu_max_id(cls, bar: XMenuBar) -> int:
        """
        Scan through the IDs used by all the items in
        this menubar, and return the largest ID encountered.
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