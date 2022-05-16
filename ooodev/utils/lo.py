# coding: utf-8
# Python conversion of Lo.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

from __future__ import annotations
import sys
from datetime import datetime
import time
from typing import TYPE_CHECKING, Iterable, Optional, List, Tuple, overload, TypeVar, Type
from urllib.parse import urlparse
import uno
from .connect import ConnectBase, LoPipeStart, LoSocketStart
from com.sun.star.lang import DisposedException
from com.sun.star.util import CloseVetoException
from com.sun.star.io import IOException
from com.sun.star.document import MacroExecMode
from com.sun.star.util import XCloseable
from contextlib import contextmanager


T = TypeVar('T')
# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from . import props as mProps
from . import file_io as mFileIO
from . import xml_util as mXML
from . import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.beans import XPropertySet
    from com.sun.star.beans import XIntrospection
    from com.sun.star.container import XChild
    from com.sun.star.container import XIndexAccess
    from com.sun.star.container import XNamed
    from com.sun.star.frame import XDesktop
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.frame import XDispatchHelper
    from com.sun.star.frame import XFrame
    from com.sun.star.frame import XStorable
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.lang import XComponent
    from com.sun.star.lang import XMultiServiceFactory
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.script.provider import XScriptContext
    from com.sun.star.uno import XComponentContext
    from com.sun.star.uno import XInterface

if sys.version_info >= (3, 10):
    from typing import Union
else:
    from typing_extensions import Union

class Lo:
    class Loader:
        """
        Context Manager for Loader
        
        Example:
        
            .. code::
            
                with Lo.Loader() as loader:
                    doc = Write.create_doc(loader)
                    ...
        """
        def __init__(self, using_pipes=False):
            """
            Create a connection to office

            Args:
                using_pipes (bool, optional): If True the pipes are used to connect to office; Otherwise, connection is made with 'host' and 'port'. Defaults to False.
            """
            self.loader = Lo.load_office(using_pipes=using_pipes)
        
        def __enter__(self) -> XComponentLoader:
            return self.loader
        
        def __exit__(self, exc_type, exc_val, exc_tb):
            Lo.close_office()
    # region docType ints
    UNKNOWN = 0
    WRITER = 1
    BASE = 2
    CALC = 3
    DRAW = 4
    IMPRESS = 5
    MATH = 6
    # endregion docType ints

    # region docType strings
    UNKNOWN_STR = "unknown"
    WRITER_STR = "swriter"
    BASE_STR = "sbase"
    CALC_STR = "scalc"
    DRAW_STR = "sdraw"
    IMPRESS_STR = "simpress"
    MATH_STR = "smath"
    # endregion docType strings

    # region docType service names
    UNKNOWN_SERVICE = "com.sun.frame.XModel"
    WRITER_SERVICE = "com.sun.star.text.TextDocument"
    BASE_SERVICE = "com.sun.star.sdb.OfficeDatabaseDocument"
    CALC_SERVICE = "com.sun.star.sheet.SpreadsheetDocument"
    DRAW_SERVICE = "com.sun.star.drawing.DrawingDocument"
    IMPRESS_SERVICE = "com.sun.star.presentation.PresentationDocument"
    MATH_SERVICE = "com.sun.star.formula.FormulaProperties"
    # endregion docType service names

    # region port connect to locally running Office via port 8100
    SOCKET_PORT = 8100
    # endregion port

    # region CLSIDs for Office documents
    # defined in https://github.com/LibreOffice/core/blob/master/officecfg/registry/data/org/openoffice/Office/Embedding.xcu
    # https://opengrok.libreoffice.org/xref/core/officecfg/registry/data/org/openoffice/Office/Embedding.xcu
    WRITER_CLSID = "8BC6B165-B1B2-4EDD-aa47-dae2ee689dd6"
    CALC_CLSID = "47BBB4CB-CE4C-4E80-a591-42d9ae74950f"
    DRAW_CLSID = "4BAB8970-8A3B-45B3-991c-cbeeac6bd5e3"
    IMPRESS_CLSID = "9176E48A-637A-4D1F-803b-99d9bfac1047"
    MATH_CLSID = "078B7ABA-54FC-457F-8551-6147e776a997"
    CHART_CLSID = "12DCAE26-281F-416F-a234-c3086127382e"

    # unsure about these:
    #
    # chart2 "80243D39-6741-46C5-926E-069164FF87BB"
    #       service: com.sun.star.chart2.ChartDocument

    #  applet "970B1E81-CF2D-11CF-89CA-008029E4B0B1"
    #       service: com.sun.star.comp.sfx2.AppletObject

    #  plug-in "4CAA7761-6B8B-11CF-89CA-008029E4B0B1"
    #        service: com.sun.star.comp.sfx2.PluginObject

    #  frame "1A8A6701-DE58-11CF-89CA-008029E4B0B1"
    #        service: com.sun.star.comp.sfx2.IFrameObject

    #  XML report chart "D7896D52-B7AF-4820-9DFE-D404D015960F"
    #        service: com.sun.star.report.ReportDefinition
    # endregion CLSIDs for Office documents

    xcc: XComponentContext = None
    """remote component context"""
    xdesktop: XDesktop = None
    """remote desktop UNO service"""

    mc_factory: XMultiComponentFactory = None
    ms_factory: XMultiServiceFactory = None

    bridge_component: XComponent = None
    """this is only set if office is opened via a socket"""
    is_office_terminated: bool = False

    lo_inst: ConnectBase = None
    
    @staticmethod
    def qi(atype:Type[T], obj: XTypeProvider) -> T | None:
        """
        Generic method that test if an object implements an interface.

        Args:
            atype (T): Interface type such as XInterface
            o (object): Object to test for interface.

        Returns:
            T | None: Return obj if interface is supported: Otherwise, None
        """
        if uno.isInterface(atype):
            uno_t = uno.getTypeByName(atype.__pyunointerface__)
            q_obj = obj.queryInterface(uno_t)
            if q_obj:
                return q_obj
            # try:
            #     types = obj.getTypes()
            # except Exception:
            #     return None
            # for t in types:
            #     if t == uno_t:
            #         return obj
        return None

    @classmethod
    def get_context(cls) -> XComponentContext:
        return cls.xcc

    @classmethod
    def get_desktop(cls) -> XDesktop:
        return cls.xdesktop

    @classmethod
    def get_component_factory(cls) -> XMultiComponentFactory:
        return cls.mc_factory

    @classmethod
    def get_service_factory(cls) -> XMultiServiceFactory:
        return cls.bridge_component

    @classmethod
    def set_ooo_bean(cls, conn: ConnectBase) -> None:
        try:
            cls.xcc = conn.ctx
            cls.mc_factory = conn.service_manager
            cls.xdesktop = conn.desktop
        except Exception as e:
            print(f"Couldn't initialize LO using OOoBean: {e}")

    # region interface object creation
    @classmethod
    def create_instance_msf(cls, service_name: str) -> XInterface | None:
        if cls.ms_factory is None:
            print("No document found")
            return None
        try:
            obj = cls.ms_factory.createInstance(service_name)
        except Exception as e:
            print(f"Couldn't create interface for '{service_name}': {e}")
            return None
        return obj

    @overload
    @classmethod
    def create_instance_mcf(cls, service_name: str) -> XInterface | None:
        """
        Create an interface object of class aType from the named service

        Args:
            service_name (str): Service Name

        Returns:
            XInterface | None: Interface if creation is successful; Otherwise None.
        """

    @overload
    @classmethod
    def create_instance_mcf(
        cls, service_name: str, args: Tuple[object, ...]
    ) -> XInterface | None:
        """
        Create an interface object of class aType from the named service

        Args:
            service_name (str): Service Name
            args (Tuple[object, ...]): Args

        Returns:
            XInterface | None: Interface if creation is successful; Otherwise None.
        """

    @classmethod
    def create_instance_mcf(
        cls, service_name: str, args: Optional[Tuple[object, ...]] = None
    ) -> XInterface | None:
        #  create an interface object of class aType from the named service;
        #  uses XComponentContext and 'new' XMultiComponentFactory
        #  so only a bridge to office is needed
        if cls.xcc is None or cls.mc_factory is None:
            print("No office connection found")
            return None
        try:
            if args is not None:
                obj = cls.mc_factory.createInstanceWithArgumentsAndContext(
                    service_name, args, cls.xcc
                )
            else:
                obj = cls.mc_factory.createInstanceWithContext(service_name, cls.xcc)
        except Exception as e:
            print(f"Couldn't create interface for '{service_name}': {e}")
            return None
        return obj

    # endregion interface object creation

    @classmethod
    def get_parent(a_component: XChild) -> XInterface:
        """
        Retrieves the parent of the given object

        Args:
            a_component (XChild): component to get parent of.

        Returns:
            XInterface: parent component.
        """
        return a_component.getParent()


    # region Start Office
    @overload
    @staticmethod
    def load_office() -> XComponentLoader:
        ...

    @overload
    @staticmethod
    def load_office(using_pipes: bool) -> XComponentLoader:
        ...

    @classmethod
    def load_office(cls, using_pipes: bool = True) -> XComponentLoader:
        # Creation sequence: remote component content (xcc) -->
        #                     remote service manager (mcFactory) -->
        #                     remote desktop (xDesktop) -->
        #                     component loader (XComponentLoader)
        # Once we have a component loader, we can load a document.
        # xcc, mcFactory, and xDesktop are stored as static globals.
        print("Loading Office...")
        if using_pipes:
            try:
                cls.lo_inst = LoPipeStart()
                cls.lo_inst.connect()
                cls.xcc = cls.lo_inst.ctx
            except Exception as e:
                print("Office context could not be created")
                raise SystemExit(1)
        else:
            try:
                cls.lo_inst = LoSocketStart()
                cls.lo_inst.connect()
                cls.xcc = cls.lo_inst.ctx
            except Exception as e:
                print("Office context could not be created")
                raise SystemExit(1)
        cls.mc_factory = cls.xcc.getServiceManager()
        if cls.mc_factory is None:
            print("Office Service Manager is unavailable")
            raise SystemExit(1)
        cls.xdesktop = cls.create_instance_mcf("com.sun.star.frame.Desktop")
        if cls.xdesktop is None:
            print("Could not create a desktop service")
            raise SystemExit(1)

        return cls.xdesktop

    @classmethod
    def load_socket_office(cls) -> XComponentLoader:
        return cls.load_office(False)

    # endregion Start Office

    # region office shutdown
    @classmethod
    def close_office(cls) -> None:
        print("Closing Office")
        if cls.xdesktop is None:
            print("No office connection found")
            return

        if cls.is_office_terminated:
            print("Office has already been requested to terminate")
            return
        num_tries = 1
        while cls.is_office_terminated is False and num_tries < 4:
            time.sleep(0.2)
            cls.is_office_terminated = cls.try_to_terminate(num_tries)
            num_tries += 1

    @classmethod
    def try_to_terminate(cls, num_tries: int) -> bool:
        try:
            is_dead = cls.xdesktop.terminate()
            if is_dead:
                if num_tries > 1:
                    print(f"{num_tries}. Office terminated")
                else:
                    print("Office terminated")
            else:
                print(f"{num_tries}. Office failed to terminate")
            return is_dead
        except DisposedException as e:
            print("Office link disposed")
            return True
        except Exception as e:
            print(f"Termination exception: {e}")
            return False

    @classmethod
    def kill_office(cls) -> None:
        if cls.lo_inst is None:
            print("No instance to kill")
            return
        try:
            cls.lo_inst.kill_soffice()
            print("Killed Office")
        except Exception as e:
            print(f"Unbale to kill Office: {e}")

    # endregion office shutdown

    # region document opening
    @classmethod
    def open_flat_doc(
        cls, fnm: str, doc_type: str, loader: XComponentLoader
    ) -> XComponent | None:
        nn = mXML.XML.get_flat_fiter_name(doc_type=doc_type)
        print(f"Flat filter Name: {nn}")
        return cls.open_doc(fnm, loader, mProps.Props.make_props(Hidden=True))

    @overload
    @classmethod
    def open_doc(cls, fnm: str, loader: XComponentLoader) -> XComponent | None:
        ...

    @overload
    @classmethod
    def open_doc(
        cls, fnm: str, loader: XComponentLoader, props: Iterable[PropertyValue]
    ) -> XComponent | None:
        ...

    @classmethod
    def open_doc(
        cls,
        fnm: str,
        loader: XComponentLoader,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent | None:
        if fnm is None:
            print("Filename is null")
            return None

        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        open_file_url = None
        if not mFileIO.FileIO.is_openable(fnm):
            if cls.is_url(fnm):
                print(f"Will treat filename as a URL: '{fnm}'")
                open_file_url = fnm
            else:
                return None
        else:
            print(f"Opening {fnm}")
            open_file_url = mFileIO.FileIO.fnm_to_url(fnm)
            if open_file_url is None:
                return None

        doc: XComponent = None
        try:
            doc = loader.loadComponentFromURL(open_file_url, "_blank", 0, props)
        except Exception:
            print("Unable to open the document")
        return doc

    @classmethod
    def open_readonly_doc(cls, fnm: str, loader: XComponentLoader) -> XComponent:
        return cls.open_doc(fnm, loader, mProps.Props.make_props(Hidden=True, ReadOnly=True))

    # ======================== document creation ==============

    @classmethod
    def ext_to_doc_type(cls, ext: str) -> str:
        e = ext.casefold()
        if e == "":
            print("Empty string: Using writer")
            return cls.WRITER_STR
        if e == "odt":
            return cls.WRITER_STR
        elif e == "odp":
            return cls.IMPRESS_STR
        elif e == "odg":
            return cls.DRAW_STR
        elif e == "ods":
            return cls.CALC_STR
        elif e == "odb":
            return cls.BASE_STR
        elif e == "odf":
            return cls.MATH_STR
        else:
            print(f"Do not recognize extension '{ext}'; using writer")
            return cls.WRITER_STR

    @classmethod
    def doc_type_str(cls, doc_type_val: int) -> str:
        i = doc_type_val
        if i == cls.WRITER:
            return cls.WRITER_STR
        elif i == cls.IMPRESS:
            return cls.IMPRESS_STR
        elif i == cls.DRAW:
            return cls.DRAW_STR
        elif i == cls.CALC:
            return cls.CALC_STR
        elif i == cls.BASE:
            return cls.BASE_STR
        elif i == cls.MATH:
            return cls.MATH_STR
        else:
            print(f"Do not recognize extension '{i}'; using writer")
            return cls.WRITER_STR

    @overload
    @staticmethod
    def create_doc(doc_type: str, loader: XComponentLoader) -> XComponent | None:
        ...

    @overload
    @staticmethod
    def create_doc(
        doc_type: str, loader: XComponentLoader, props: Iterable[PropertyValue]
    ) -> XComponent | None:
        ...

    @staticmethod
    def create_doc(
        doc_type: str,
        loader: XComponentLoader,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent | None:
        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        print(f"Creating Office document {doc_type}")
        doc: XMultiServiceFactory = None
        try:
            doc = loader.loadComponentFromURL(
                f"private:factory/{doc_type}", "_blank", 0, props
            )
        except Exception:
            print("Could not create a document")
        return doc

    @classmethod
    def create_macro_doc(
        cls, doc_type: str, loader: XComponentLoader
    ) -> XComponent | None:
        return cls.create_doc(
            doc_type=doc_type,
            loader=loader,
            props=mProps.Props.make_props(
                Hidden=False, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE
            ),
        )

    @staticmethod
    def create_doc_from_template(
        template_path: str, loader: XComponentLoader
    ) -> XComponent | None:
        if not mFileIO.FileIO.is_openable(template_path):
            print(f"Template file can not be opened: '{template_path}'")
            return None
        print(f"Opening template: '{template_path}'")
        template_url = mFileIO.FileIO.fnm_to_url(fnm=template_path)
        if template_url is None:
            return None

        props = mProps.Props.make_props(Hidden=True, AsTemplate=True)
        doc = None
        try:
            doc = loader.loadComponentFromURL(template_url, "_blank", 0, props)
        except Exception as e:
            print(f"Could not create document from template: {e}")
        return doc

    # ======================== document saving ==============

    @staticmethod
    def save(doc: XStorable) -> None:
        try:
            doc.store()
            print("Saved the document by overwriting")
        except IOException as e:
            print(f"Could not save the document")
            print(f"    {e}")

    @overload
    @staticmethod
    def save_doc(doc: XStorable, fnm: str) -> None:
        ...

    @overload
    @staticmethod
    def save_doc(doc: XStorable, fnm: str, password: str) -> None:
        ...

    @overload
    @staticmethod
    def save_doc(doc: XStorable, fnm: str, password: str, format: str) -> None:
        ...

    @classmethod
    def save_doc(
        cls, doc: XStorable, fnm: str, password: str = None, format: str = None
    ) -> None:
        doc_type = mInfo.Info.report_doc_type(doc)
        kargs = {"fnm": fnm, "store": doc, "doc_type": doc_type}
        if password is not None:
            kargs["password"] = password
        if format is None:
            cls.store_doc(**kargs)
        else:
            kargs["format"] = format
            cls.store_doc_format(**kargs)

    @overload
    @staticmethod
    def store_doc(store: XStorable, doc_type: int, fnm: str) -> None:
        ...

    @overload
    @staticmethod
    def store_doc(store: XStorable, doc_type: int, fnm: str, password: str) -> None:
        ...

    @classmethod
    def store_doc(
        cls, store: XStorable, doc_type: int, fnm: str, password: Optional[str] = None
    ) -> None:
        ext = mInfo.Info.get_ext(fnm)
        frmt = "Text"
        if ext is None:
            "Assuming a text format"
        else:
            frmt = cls.ext_to_format(ext=ext, doc_type=doc_type)
        if password is None:
            cls.store_doc_format(store=store, fnm=fnm, format=frmt)
        else:
            cls.store_doc_format(store=store, fnm=fnm, format=frmt, password=password)

    @overload
    @staticmethod
    def ext_to_format(ext: str) -> str:
        ...

    @overload
    @staticmethod
    def ext_to_format(ext: str, doc_type: int) -> str:
        """
        convert the extension string into a suitable office format string.
        The formats were chosen based on the fact that they
        are being used to save (or export) a document.

        Args:
            ext (str): document extension
            doc_type (int, optional): Type of document.

        Returns:
            str: format of ext.

        Note:
            ``doc_type``is used to distinguish between the various meanings of the PDF ext.
            This could be a lot more extensive.

            Use ``Info.getFilterNames()`` to get the filter names for your Office.
        """
        ...

    @classmethod
    def ext_to_format(cls, ext: str, doc_type: int = UNKNOWN) -> str:
        i = doc_type
        s = ext.lower()
        if s == "doc":
            return "MS Word 97"
        elif s == "docx":
            return "Office Open XML Text"  # MS Word 2007 XML
        elif s == "rtf":
            if i == cls.CALC:
                return "Rich Text Format (StarCalc)"
            else:
                return "Rich Text Format"
        elif s == "odt":
            return "writer8"
        elif s == "ott":
            return "writer8_template"
        elif s == "pdf":
            if i == cls.WRITER:
                return "writer_pdf_Export"
            elif i == cls.IMPRESS:
                return "impress_pdf_Export"
            elif i == cls.DRAW:
                return "draw_pdf_Export"
            elif i == cls.CALC:
                return "calc_pdf_Export"
            elif i == cls.MATH:
                return "math_pdf_Export"
            else:
                return "writer_pdf_Export"  # assume we are saving a writer doc
        elif s == "txt":
            return "Text"
        elif s == "ppt":
            return "MS PowerPoint 97"
        elif s == "pptx":
            return "Impress MS PowerPoint 2007 XML"
        elif s == "odp":
            return "impress8"
        elif s == "odg":
            return "draw8"
        elif s == "jpg":
            if i == cls.IMPRESS:
                return "impress_jpg_Export"
            else:
                return "draw_jpg_Export"
        elif s == "png":
            if i == cls.IMPRESS:
                return "impress_png_Export"
            else:
                return "draw_png_Export"
        elif s == "xls":
            return "MS Excel 97"
        elif s == "xlsx":
            return "Calc MS Excel 2007 XML"
        elif s == "csv":
            return "Text - txt - csv (StarCalc)"  # "Text CSV"
        elif s == "ods":
            return "calc8"
        elif s == "odb":
            return "StarOffice XML (Base)"
        elif s == "htm" or s == "html":
            if i == cls.WRITER:
                return "HTML (StarWriter)"
            elif i == cls.IMPRESS:
                return "impress_html_Export"
            elif i == cls.DRAW:
                return "draw_html_Export"
            elif i == cls.CALC:
                return "HTML (StarCalc)"
            else:
                return "HTML"
        elif s == "xhtml":
            if i == cls.WRITER:
                return "XHTML Writer File"
            elif i == cls.IMPRESS:
                return "XHTML Impress File"
            elif i == cls.DRAW:
                return "XHTML Draw File"
            elif i == cls.CALC:
                return "XHTML Calc File"
            else:
                return "XHTML Writer File"
        elif s == "xml":
            if i == cls.WRITER:
                return "OpenDocument Text Flat XML"
            elif i == cls.IMPRESS:
                return "OpenDocument Presentation Flat XML"
            elif i == cls.DRAW:
                return "OpenDocument Drawing Flat XML"
            elif i == cls.CALC:
                return "OpenDocument Spreadsheet Flat XML"
            else:
                return "OpenDocument Text Flat XML"

        else:
            print(f"Do not recognize extension '{ext}'; using text")
            return "Text"

    @overload
    @staticmethod
    def store_doc_format(store: XStorable, fnm: str, format: str) -> None:
        ...

    @overload
    @staticmethod
    def store_doc_format(
        store: XStorable, fnm: str, format: str, password: str
    ) -> None:
        ...

    @staticmethod
    def store_doc_format(
        store: XStorable, fnm: str, format: str, password: str = None
    ) -> None:
        print(f"Saving the document in '{fnm}'")
        print(f"Using format {format}")

        try:
            save_file_url = mFileIO.FileIO.fnm_to_url(fnm)
            if save_file_url is None:
                return
            if password is None:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=format)
            else:
                store_props = mProps.Props.make_props(
                    Overwrite=True, FilterName=format, Password=password
                )
            store.storeToURL(save_file_url, store_props)
        except IOException as e:
            print(f"Could not save '{fnm}': {e}")

    # ======================== document closing ==============

    @overload
    @staticmethod
    def close(closeable: XCloseable) -> None:
        """
        Closes a document.

        Args:
            closeable (XCloseable): Object that implements XCloseable interface.
        """
        ...
    
    @overload
    @staticmethod
    def close(closeable: XCloseable, deliver_ownership: bool) -> None:
        """
        Closes a document.

        Args:
            closeable (XCloseable): Object that implements XCloseable interface.
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException.
                This new owner has to close the closing object again if his still running
                processes will be finished.
                False let the ownership at the original one which called the close() method.
                They must react for possible CloseVetoExceptions such as when document needs saving
                and try it again at a later time. This can be useful for a generic UI handling.
        """
        ...

    @staticmethod
    def close(closeable: XCloseable, deliver_ownership=False) -> None:
        if closeable is None:
            return
        print("Closing the document")
        try:
            closeable.close(deliver_ownership)
        except CloseVetoException:
            print("Close was vetoed")


    @overload
    @staticmethod
    def close_doc(doc: object) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Closeable doccument
        """
        ...

    @overload
    @staticmethod
    def close_doc(doc: object, deliver_ownership: bool) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Closeable doccument
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException.
                This new owner has to close the closing object again if his still running
                processes will be finished.
                False let the ownership at the original one which called the close() method.
                They must react for possible CloseVetoExceptions such as when document needs saving
                and try it again at a later time. This can be useful for a generic UI handling.
        """
        ...

    @classmethod
    def close_doc(cls, doc: object, deliver_ownership=False) -> None:
        try:
            closeable = cls.qi(XCloseable, doc)
            cls.close(closeable)
        except DisposedException:
            print("Document close failed since Office link disposed")

    # ================= initialization via Addon-supplied context ====================

    @staticmethod
    def addon_initialize(addon_xcc: XComponentContext) -> XComponent | None:
        xcc = addon_xcc
        if xcc is None:
            print("Could not access component context")
            return None
        mc_factory = xcc.getServiceManager()
        if mc_factory is None:
            print("Office Service Manager is unavailable")
            return None

        try:
            xdesktop: XDesktop = mc_factory.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xcc
            )
        except Exception:
            print("Could not access desktop")
            return None
        doc = xdesktop.getCurrentComponent()
        if doc is None:
            print("Could not access document")
            return None
        return doc

    # ============= initialization via script context ======================

    @staticmethod
    def script_initialize(sc: XScriptContext) -> XComponent | None:
        if sc is None:
            print("Script Context is null")
            return None
        xcc = sc.getComponentContext()
        if xcc is None:
            print("Could not access component context")
            return None
        mc_factory = xcc.getServiceManager()
        if mc_factory is None:
            print("Office Service Manager is unavailable")
            return None
        xdesktop = sc.getDesktop()
        if xdesktop is None:
            print("Could not access desktop")
            return None
        doc = xdesktop.getCurrentComponent()
        if doc is None:
            print("Could not access document")
            return None
        return doc

    # ==================== dispatch ===============================
    # see https://wiki.documentfoundation.org/Development/DispatchCommands

    @overload
    @staticmethod
    def dispatch_cmd(cmd: str) -> bool:
        ...

    @overload
    @staticmethod
    def dispatch_cmd(cmd: str, props: Iterable[PropertyValue]) -> bool:
        ...

    @overload
    @staticmethod
    def dispatch_cmd(cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> bool:
        ...

    @classmethod
    def dispatch_cmd(
        cls, cmd: str, props: Iterable[PropertyValue] = None, frame: XFrame = None
    ) -> bool:
        if props is None:
            props = ()
        if frame is None:
            frame = cls.xdesktop.getCurrentFrame()

        helper: XDispatchHelper = cls.create_instance_mcf(
            "com.sun.star.frame.DispatchHelper"
        )
        if helper is None:
            print(f"Could not create dispatch helper for command {cmd}")
            return False
        try:
            helper.executeDispatch(frame, f".uno:{cmd}", "", 0, props)
            return True
        except Exception as e:
            print(f"Could not dispatch '{cmd}'")
            print(f"    {e}")
        return False

    # ================= Uno cmds =========================

    @staticmethod
    def make_uno_cmd(item_name: str) -> str:
        return f"vnd.sun.star.script:Foo/Foo.{item_name}?language=Java&location=share"

    @staticmethod
    def extract_item_name(uno_cmd: str) -> str | None:
        try:
            foo_pos = uno_cmd.index("Foo.")
        except ValueError:
            print(f"Could not find Foo header in command: '{uno_cmd}'")
            return None
        try:
            lang_pos = uno_cmd.index("?language")
        except ValueError:
            print(f"Could not find language header in command: '{uno_cmd}'")
            return None
        start = foo_pos + 4
        return uno_cmd[start:lang_pos]

    # ======================== use Inspector extensions ====================

    @classmethod
    def inspect(cls, obj: object) -> None:
        if cls.xcc is None or cls.mc_factory is None:
            print("No office connection found")
            return
        try:
            ts = mInfo.Info.get_interface_types(obj)
            title = "Object"
            if ts is not None and len(ts) > 0:
                title = ts[0].getTypeName() + " " + title
            inspector = cls.mc_factory.createInstanceWithContext(
                "org.openoffice.InstanceInspector", cls.xcc
            )
            #       hands on second use
            if inspector is None:
                print("Inspector Service could not be instantiated")
                return
            print("Inspector Service instantiated")
            intro: XIntrospection = cls.create_instance_mcf(
                "com.sun.star.beans.Introspection"
            )
            intro_acc = intro.inspect(inspector)
            method = intro_acc.getMethod("inspect", -1)
            print(f"inspect() method was found: {method is not None}")
            params = [[obj, title]]
            method.invoke(inspector, params)
        except Exception as e:
            print("Could not access Inspector:")
            print(f"    {e}")

    @classmethod
    def mri_inspect(cls, obj: object) -> None:
        """call MRI's inspect()"""
        # Available from http://extensions.libreoffice.org/extension-center/mri-uno-object-inspection-tool
        #               or http://extensions.services.openoffice.org/en/project/MRI
        #  Docs: https://github.com/hanya/MRI/wiki
        #  Forum tutorial: https://forum.openoffice.org/en/forum/viewtopic.php?f=74&t=49294
        xi: XIntrospection = cls.create_instance_mcf("mytools.Mri")
        if xi is None:
            print("MRI Inspector Service could not be instantiated")
            return
        print("MRI Inspector Service instantiated")
        xi.inspect(obj)

    # ------------------ color methods ---------------------
    # section intentionally left out.

    # ================== other utils =============================

    @staticmethod
    def delay(ms: int) -> None:
        """
        Delay execution for a given number of milli-seconds.

        Args:
            ms (int): Number of milli-seconds to delay
        """
        if ms <= 0:
            print("Ms must be greater then zero")
            return
        sec = ms / 1000
        time.sleep(sec)

    wait = delay

    @staticmethod
    def is_null_or_empty(s: str) -> bool:
        """
        Gets is a string is None or Empyt

        Args:
            s (str): String to check.

        Returns:
            bool: True if None or empty string; Otherwise, False
        """
        return s == None or len(s) == 0

    @staticmethod
    def wait_enter() -> None:
        """
        Terminal dispalays Press Enter to continue...
        """
        input("Press Enter to continue...")

    @staticmethod
    def get_time_stamp() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def is_url(fnm: str) -> bool:
        # https://stackoverflow.com/questions/7160737/how-to-validate-a-url-in-python-malformed-or-not
        try:
            result = urlparse(fnm)
            return all([result.scheme, result.netloc])
        except ValueError:
            return False

    # endregion document opening

    @staticmethod
    def capitalize(s: str) -> str | None:
        if s is None or len(s) == 0:
            return None
        return s.capitalize()

    @staticmethod
    def parse_int(s: str) -> int:
        if s is None:
            return 0
        try:
            return int(s)
        except ValueError:
            print(f"{s} could not be parsed as an int; using 0")
        return 0

    @overload
    @staticmethod
    def print_names(names: Iterable[str]) -> None:
        ...

    @overload
    @staticmethod
    def print_names(names: Iterable[str], num_per_line: int) -> None:
        if names is None:
            print("  No names found")
            return
        sorted_list = sorted(names, key=str.casefold)
        print(f"No. of names: {len(sorted_list)}")
        nl_count = 0
        for name in sorted_list:
            print(f"  '{name}'")
            nl_count += 1
        if nl_count % num_per_line == 0:
            print()
            nl_count = 0
        print("\n\n")

    # ------------------- container manipulation --------------------

    @staticmethod
    def print_table(name: str, table: List[List[str]]) -> None:
        print(f"-- {name} ----------------")
        for row in table:
            col_str = "  ".join([str(el) for el in row])
            print(col_str)
        print()

    @staticmethod
    def get_container_names(con: XIndexAccess) -> List[str] | None:
        if con is None:
            print("Container is null")
            return None
        num_el = con.getCount()
        if num_el == 0:
            print("No elements in the container")
            return None

        names_list = []
        for i in range(num_el):
            named = con.getByIndex(i)
            names_list.append(named.getName())

        if len(names_list) == 0:
            print("No element names found in the container")
            return None
        return names_list

    @staticmethod
    def find_container_props(con: XIndexAccess, nm: str) -> XPropertySet | None:
        if con is None:
            print("Container is null")
            return None
        for i in range(con.getCount()):
            try:
                named: XNamed = con.getByIndex(i)
                if named.getName() == nm:
                    return named
            except Exception:
                print(f"Could not access element {i}")
        print(f"Could not find a '{nm}' property set in the container")
        return None
