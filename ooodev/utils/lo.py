# coding: utf-8
# Python conversion of Lo.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

from __future__ import annotations
from datetime import datetime
import time
from typing import TYPE_CHECKING, Iterable, Optional, List, Tuple, overload, TypeVar, Type
from urllib.parse import urlparse
import uno
from enum import IntEnum, Enum
from ..meta.static_meta import StaticProperty, classproperty
from .connect import ConnectBase, LoPipeStart, LoSocketStart
from com.sun.star.beans import XPropertySet
from com.sun.star.beans import XIntrospection
from com.sun.star.container import XNamed
from com.sun.star.document import MacroExecMode  # const
from com.sun.star.frame import XDesktop
from com.sun.star.frame import XDispatchHelper
from com.sun.star.lang import DisposedException
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.io import IOException
from com.sun.star.util import CloseVetoException
from com.sun.star.util import XCloseable
from com.sun.star.util import XNumberFormatsSupplier
from com.sun.star.frame import XComponentLoader

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XChild
    from com.sun.star.container import XIndexAccess
    from com.sun.star.frame import XFrame
    from com.sun.star.frame import XStorable
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.lang import XComponent
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.script.provider import XScriptContext
    from com.sun.star.uno import XComponentContext
    from com.sun.star.uno import XInterface

T = TypeVar("T")
UnoInterface = TypeVar("UnoInterface")

# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from . import props as mProps
from . import file_io as mFileIO
from . import xml_util as mXML
from . import info as mInfo
from ..exceptions import ex as mEx


class Lo(metaclass=StaticProperty):
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
    class DocType(IntEnum):
        UNKNOWN = 0
        WRITER = 1
        BASE = 2
        CALC = 3
        DRAW = 4
        IMPRESS = 5
        MATH = 6
        
        def __str__(self) -> str:
            return str(self.value)


    # region docType strings
    class DocTypeStr(str, Enum):
        UNKNOWN = "unknown"
        WRITER = "swriter"
        BASE = "sbase"
        CALC = "scalc"
        DRAW = "sdraw"
        IMPRESS = "simpress"
        MATH = "smath"
    
        def __str__(self) -> str:
            return self.value
    
    # region docType service names
    class Service(str, Enum):
        UNKNOWN = "com.sun.frame.XModel"
        WRITER = "com.sun.star.text.TextDocument"
        BASE = "com.sun.star.sdb.OfficeDatabaseDocument"
        CALC = "com.sun.star.sheet.SpreadsheetDocument"
        DRAW = "com.sun.star.drawing.DrawingDocument"
        IMPRESS = "com.sun.star.presentation.PresentationDocument"
        MATH = "com.sun.star.formula.FormulaProperties"
        
        def __str__(self) -> str:
            return self.value
    # endregion docType service names

   

    # region CLSIDs for Office documents
    # defined in https://github.com/LibreOffice/core/blob/master/officecfg/registry/data/org/openoffice/Office/Embedding.xcu
    # https://opengrok.libreoffice.org/xref/core/officecfg/registry/data/org/openoffice/Office/Embedding.xcu
    class CLSID(str, Enum):
        WRITER = "8BC6B165-B1B2-4EDD-aa47-dae2ee689dd6"
        CALC = "47BBB4CB-CE4C-4E80-a591-42d9ae74950f"
        DRAW = "4BAB8970-8A3B-45B3-991c-cbeeac6bd5e3"
        IMPRESS = "9176E48A-637A-4D1F-803b-99d9bfac1047"
        MATH = "078B7ABA-54FC-457F-8551-6147e776a997"
        CHART = "12DCAE26-281F-416F-a234-c3086127382e"
        
        def __str__(self) -> str:
            return self.value

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

     # region port connect to locally running Office via port 8100
    _SOCKET_PORT = 8100
    # endregion port

    _xcc: XComponentContext = None
    _doc: XComponent = None
    """remote component context"""
    _xdesktop: XDesktop = None
    """remote desktop UNO service"""

    _mc_factory: XMultiComponentFactory = None
    _ms_factory: XMultiServiceFactory = None

    _bridge_component: XComponent = None
    """this is only set if office is opened via a socket"""
    _is_office_terminated: bool = False

    _lo_inst: ConnectBase = None

    @staticmethod
    def qi(atype: Type[T], obj: XTypeProvider) -> T | None:
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface
            obj (object): Object that implements interface.

        Returns:
            T | None: instance of interface if supported; Otherwise, None
    
        Example:
                
            .. code-block:: python
                :emphasize-lines: 3
    
                from com.sun.star.util import XSearchable
                cell_range = ...
                srch = Lo.qi(XSearchable, cell_range)
                sd = srch.createSearchDescriptor()
        """
        if uno.isInterface(atype) and hasattr(obj, "queryInterface"):
            uno_t = uno.getTypeByName(atype.__pyunointerface__)
            return obj.queryInterface(uno_t)
        return None

    @classmethod
    def get_context(cls) -> XComponentContext:
        """
        Gets current LO Component Context
        """
        return cls._xcc

    @classmethod
    def get_desktop(cls) -> XDesktop:
        """
        Gets current LO Desktop
        """
        return cls._xdesktop

    @classmethod
    def get_component_factory(cls) -> XMultiComponentFactory:
        """Gets current multi component factory"""
        return cls._mc_factory

    @classmethod
    def get_service_factory(cls) -> XMultiServiceFactory:
        """Gets current multi service factory"""
        return cls._bridge_component

    # region interface object creation

    # region    create_instance_msf()
    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str) -> T:
        """
        Creates in instace of a service using multi service factory.

        Args:
            atype (Type[T]): Type of interface to return
            service_name (str): Service name

        Raises:
            Exception: if unable to create instance

        Returns:
            T: Instance of interface for the service name.
        """
        ...

    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str, msf: XMultiServiceFactory) -> T:
        """
        Creates in instace of a service using multi service factory.

        Args:
            atype (Type[T]): Type of interface to return
            service_name (str): Service name
            msf (XMultiServiceFactory): Multi service factory used to create instance

        Raises:
            Exception: if unable to create instance

        Returns:
            T: Instance of interface for the service name.
        """
        ...

    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str, msf: XMultiServiceFactory = None) -> T:
        """
        Creates in instace of a service using multi service factory.

        Args:
            atype (Type[T]): Type of interface to return
            service_name (str): Service name
            msf (XMultiServiceFactory): Multi service factory used to create instance

        Raises:
            Exception: if unable to create instance

        Returns:
            T: Instance of interface for the service name.

        Example:
            In the following example ``src_con`` is an instance of ``XSheetCellRangeContainer``

            .. code-block:: python

                from com.sun.star.sheet import XSheetCellRangeContainer
                src_con = Lo.create_instance_msf(XSheetCellRangeContainer, "com.sun.star.sheet.SheetCellRanges")
                
        """
        if cls._ms_factory is None:
            raise Exception("No document found")
        try:
            if msf is None:
                obj = cls._ms_factory.createInstance(service_name)
            else:
                obj = msf.createInstance(service_name)
            interface_obj = cls.qi(atype=atype, obj=obj)
            if interface_obj is None:
                raise mEx.MissingInterfaceError(atype)
        except Exception as e:
            raise Exception(f"Couldn't create interface for '{service_name}'") from e
        return interface_obj

    # endregion create_instance_msf()

    # region    create_instance_mcf()
    @overload
    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str) -> T:
        """
        Create an interface object of class atype from the named service

        Args:
            service_name (str): Service Name

        Raises:
            Exception: if unable to create instance

        Returns:
            T: Instance of interface for the service name.
        """

    @overload
    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str, args: Tuple[object, ...]) -> T:
        """
        Create an interface object of class atype from the named service

        Args:
            service_name (str): Service Name
            args (Tuple[object, ...]): Args

        Raises:
            Exception: if unable to create instance

        Returns:
            T: Instance of interface for the service name.
        """

    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str, args: Optional[Tuple[object, ...]] = None) -> T:
        """
        Create an interface object of class atype from the named service

        Args:
            service_name (str): Service Name
            args (Tuple[object, ...]): Args

        Raises:
            Exception: if unable to create instance

        Returns:
            T: Instance of interface for the service name.

        Example:
            In the following example ``tk`` is an instance of ``XExtendedToolkit``

            .. code-block:: python

                from com.sun.star.awt import XExtendedToolkit
                tk = Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
                
        """
        #  create an interface object of class atype from the named service;
        #  uses XComponentContext and XMultiComponentFactory
        #  so only a bridge to office is needed
        if cls._xcc is None or cls._mc_factory is None:
            raise Exception("No office connection found")
        try:
            if args is not None:
                obj = cls._mc_factory.createInstanceWithArgumentsAndContext(service_name, args, cls._xcc)
            else:
                obj = cls._mc_factory.createInstanceWithContext(service_name, cls._xcc)
            interface_obj = cls.qi(atype=atype, obj=obj)
            if interface_obj is None:
                raise mEx.MissingInterfaceError(atype)
        except Exception as e:
            raise Exception(f"Couldn't create interface for '{service_name}'") from e
        return interface_obj

    # endregion create_instance_mcf()

    # endregion interface object creation

    @classmethod
    def get_parent(cls, a_component: XChild) -> XInterface:
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
    @classmethod
    def load_office(cls) -> XComponentLoader:
        """
        Loads Office
        
        If running outside of office then a bridge is created that connects to office using pipes.
        
        If running from inside of office e.g. in a macro, then XSCRIPTCONTEXT is used

        Returns:
            XDesktop: Desktop object
        """
        ...

    @overload
    @classmethod
    def load_office(cls, using_pipes: bool) -> XComponentLoader:
        """
        Loads Office
        
        If running outside of office then a bridge is created that connects to office.
        
        If running from inside of office e.g. in a macro, then XSCRIPTCONTEXT is used.
        ``using_pipes`` is ignored with running inside office.

        Args:
            using_pipes (bool): determines if bridge connection is made with pipes or host, port.

        Returns:
            XDesktop: Desktop object.
        """
        ...

    @classmethod
    def load_office(cls, using_pipes: bool = True) -> XComponentLoader:
        """
        Loads Office
        
        If running outside of office then a bridge is created that connects to office.
        
        If running from inside of office e.g. in a macro, then XSCRIPTCONTEXT is used.
        ``using_pipes`` is ignored with running inside office.

        Args:
            using_pipes (bool): determines if bridge connection is made with pipes or host, port.

        Returns:
            XDesktop: Desktop object.
        """
        # Creation sequence: remote component content (xcc) -->
        #                     remote service manager (mcFactory) -->
        #                     remote desktop (xDesktop) -->
        #                     component loader (XComponentLoader)
        # Once we have a component loader, we can load a document.
        # xcc, mcFactory, and xDesktop are stored as static globals.
        if "XSCRIPTCONTEXT" in globals():
            return cls._load()
        print("Loading Office...")
        if using_pipes:
            try:
                cls._lo_inst = LoPipeStart()
                cls._lo_inst.connect()
            except Exception as e:
                print("Office context could not be created")
                raise SystemExit(1)
        else:
            try:
                cls._lo_inst = LoSocketStart()
                cls._lo_inst.connect()
            except Exception as e:
                print("Office context could not be created")
                raise SystemExit(1)
        # cls.set_ooo_bean(conn=cls.lo_inst)
        cls._xcc = cls._lo_inst.ctx
        cls._mc_factory = cls._xcc.getServiceManager()
        if cls._mc_factory is None:
            print("Office Service Manager is unavailable")
            raise SystemExit(1)
        cls._xdesktop = cls.create_instance_mcf(XDesktop, "com.sun.star.frame.Desktop")
        if cls._xdesktop is None:
            print("Could not create a desktop service")
            raise SystemExit(1)
        loader = cls.qi(XComponentLoader, cls._xdesktop)
        if loader is None:
            raise print("Unable to access XComponentLoader")
            SystemExit(1)
        return loader
        # return cls.xdesktop

    @classmethod
    def _load(cls) -> XComponentLoader:
        if not "XSCRIPTCONTEXT" in globals():
            raise Exception("XSCRIPTCONTEXT not found")
        XSCRIPTCONTEXT = globals()['XSCRIPTCONTEXT']
        cls._xcc = XSCRIPTCONTEXT.ctx
        cls._mc_factory = cls._xcc.getServiceManager()
        if cls._mc_factory is None:
            raise Exception("Office Service Manager is unavailable")
        cls._xdesktop = XSCRIPTCONTEXT.getDesktop()
        if cls._xdesktop is None:
            Exception("Could not create a desktop service")
        loader = cls.qi(XComponentLoader, cls._xdesktop)
        if loader is None:
            raise Exception("Unable to access XComponentLoader")
        return loader

        
    # endregion Start Office

    # region office shutdown
    @classmethod
    def close_office(cls) -> None:
        """Closes the ofice connection."""
        print("Closing Office")
        cls._doc = None
        if cls._xdesktop is None:
            print("No office connection found")
            return

        if cls._is_office_terminated:
            print("Office has already been requested to terminate")
            return
        num_tries = 1
        while cls._is_office_terminated is False and num_tries < 4:
            time.sleep(0.2)
            cls._is_office_terminated = cls._try_to_terminate(num_tries)
            num_tries += 1

    @classmethod
    def _try_to_terminate(cls, num_tries: int) -> bool:
        try:
            is_dead = cls._xdesktop.terminate()
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
        """
        Kills the office connection.
        
        See Also:
            :py:meth:`~Lo.close_office`
        """
        if cls._lo_inst is None:
            print("No instance to kill")
            return
        try:
            cls._lo_inst.kill_soffice()
            print("Killed Office")
        except Exception as e:
            raise Exception(f"Unbale to kill Office") from e

    # endregion office shutdown

    # region document opening
    @classmethod
    def open_flat_doc(cls, fnm: str, doc_type: DocType, loader: XComponentLoader) -> XComponent:
        """
        Opens a flat document

        Args:
            fnm (str): path of xml documenet
            doc_type (DocType): Type of document to open
            loader (XComponentLoader): Component loader

        Returns:
            XComponent: Document

        See Also:
            * :py:meth:`~Lo.open_doc`
            * :py:meth:`~Lo.open_readonly_doc`
        """
        nn = mXML.XML.get_flat_fiter_name(doc_type=doc_type)
        print(f"Flat filter Name: {nn}")
        return cls.open_doc(fnm, loader, mProps.Props.make_props(Hidden=True))

    @overload
    @classmethod
    def open_doc(cls, fnm: str, loader: XComponentLoader) -> XComponent:
        """
        Open a office document

        Args:
            fnm (str): path of document to open
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: if unable to open document

        Returns:
            XComponent: Document
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: str, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        """
        Open a office document

        Args:
            fnm (str): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader

        Raises:
            Exception: if unable to open document

        Returns:
            XComponent: Document
        """
        ...

    @classmethod
    def open_doc(
        cls,
        fnm: str,
        loader: XComponentLoader,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        """
        Open a office document

        Args:
            fnm (str): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader

        Raises:
            Exception: if unable to open document

        Returns:
            XComponent: Document

        See Also:
            * :py:meth:`~Lo.open_readonly_doc`
            * :py:meth:`~Lo.open_flat_doc`
        """
        if fnm is None:
            raise Exception("Filename is null")
        fnm = str(fnm)

        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        open_file_url = None
        if not mFileIO.FileIO.is_openable(fnm):
            if cls.is_url(fnm):
                print(f"Will treat filename as a URL: '{fnm}'")
                open_file_url = fnm
            else:
                raise Exception(f"Unable to get url from file: {fnm}")
        else:
            print(f"Opening {fnm}")
            open_file_url = mFileIO.FileIO.fnm_to_url(fnm)

        try:
            doc = loader.loadComponentFromURL(open_file_url, "_blank", 0, props)
            cls._ms_factory = cls.qi(XMultiServiceFactory, doc)
            cls._doc = doc
            return doc
        except Exception as e:
            raise Exception("Unable to open the document") from e

    @classmethod
    def open_readonly_doc(cls, fnm: str, loader: XComponentLoader) -> XComponent:
        """
        Open a office document as readonly

        Args:
            fnm (str): path of document to open
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: if unable to open document

        Returns:
            XComponent: Document

        See Also:
            * :py:meth:`~Lo.open_doc`
            * :py:meth:`~Lo.open_flat_doc`
        """
        return cls.open_doc(fnm, loader, mProps.Props.make_props(Hidden=True, ReadOnly=True))

    # ======================== document creation ==============

    @classmethod
    def ext_to_doc_type(cls, ext: str) -> DocTypeStr:
        """
        Gets doctype from extension

        Args:
            ext (str): extension used for lookup

        Returns:
            DocTypeStr: DocTypeStr enum. If not match if found defaults to 'DocTypeStr.WRITER'
        """
        e = ext.casefold().lstrip('.')
        if e == "":
            print("Empty string: Using writer")
            return cls.DocTypeStr.WRITER
        if e == "odt":
            return cls.DocTypeStr.WRITER
        elif e == "odp":
            return cls.DocTypeStr.IMPRESS
        elif e == "odg":
            return cls.DocTypeStr.DRAW
        elif e == "ods":
            return cls.DocTypeStr.CALC
        elif e == "odb":
            return cls.DocTypeStr.BASE
        elif e == "odf":
            return cls.DocTypeStr.MATH
        else:
            print(f"Do not recognize extension '{ext}'; using writer")
            return cls.DocTypeStr.WRITER

    @classmethod
    def doc_type_str(cls, doc_type_val: DocType) -> DocTypeStr:
        """
        Converts a doc type into a :py:class:`~Lo.DocTypeStr` representation.

        Args:
            doc_type_val (DocType): Doc type as int

        Returns:
            DocTypeStr: doc type as string.
        """
        if doc_type_val == cls.DocType.WRITER:
            return cls.DocTypeStr.WRITER
        elif doc_type_val == cls.DocType.IMPRESS:
            return cls.DocTypeStr.IMPRESS
        elif doc_type_val == cls.DocType.DRAW:
            return cls.DocTypeStr.DRAW
        elif doc_type_val == cls.DocType.CALC:
            return cls.DocTypeStr.CALC
        elif doc_type_val == cls.DocType.BASE:
            return cls.DocTypeStr.BASE
        elif doc_type_val == cls.DocType.MATH:
            return cls.DocTypeStr.MATH
        else:
            print(f"Do not recognize extension '{doc_type_val}'; using writer")
            return cls.DocTypeStr.WRITER

    @overload
    @classmethod
    def create_doc(csl, doc_type: DocTypeStr, loader: XComponentLoader) -> XComponent:
        """
        Creates a document

        Args:
            doc_type (DocTypeStr): Document type
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: If unable to create document.

        Returns:
            XComponent: document as component.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, doc_type: DocTypeStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        """
        Creates a document

        Args:
            doc_type (DocTypeStr): Document type
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Property values

        Raises:
            Exception: If unable to create document.

        Returns:
            XComponent: document as component.
        """
        ...

    @classmethod
    def create_doc(
        cls,
        doc_type: DocTypeStr,
        loader: XComponentLoader,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        """
        Creates a document

        Args:
            doc_type (DocTypeStr): Document type
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Property values

        Raises:
            Exception: If unable to create document.

        Returns:
            XComponent: document as component.
        """
        dtype = cls.DocTypeStr(doc_type)
        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        print(f"Creating Office document {dtype}")
        try:
            doc = loader.loadComponentFromURL(f"private:factory/{dtype}", "_blank", 0, props)
            cls._ms_factory = cls.qi(XMultiServiceFactory, doc)
            if cls._ms_factory is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            cls._doc = doc
            return cls._doc
        except Exception as e:
            raise Exception("Could not create a document") from e

    @classmethod
    def create_macro_doc(cls, doc_type: DocTypeStr, loader: XComponentLoader) -> XComponent:
        """
        Create a document that allows executing of macros

        Args:
            doc_type (DocTypeStr): Document type
            loader (XComponentLoader): Component Loader

        Returns:
            XComponent: document as component.
        """
        return cls.create_doc(
            doc_type=doc_type,
            loader=loader,
            props=mProps.Props.make_props(Hidden=False, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE),
        )

    @classmethod
    def create_doc_from_template(cls, template_path: str, loader: XComponentLoader) -> XComponent:
        """
        Create a document form a template

        Args:
            template_path (str): path to template file
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: If unable to create document.

        Returns:
            XComponent: document as component.
        """
        if not mFileIO.FileIO.is_openable(template_path):
            raise Exception(f"Template file can not be opened: '{template_path}'")
        print(f"Opening template: '{template_path}'")
        template_url = mFileIO.FileIO.fnm_to_url(fnm=template_path)

        props = mProps.Props.make_props(Hidden=True, AsTemplate=True)
        try:
            cls._doc = loader.loadComponentFromURL(template_url, "_blank", 0, props)
            cls._ms_factory = cls.qi(XMultiServiceFactory, cls._doc)
            if cls._ms_factory is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            return cls._doc
        except Exception as e:
            raise Exception(f"Could not create document from template") from e


    @classmethod
    def get_document(cls) -> XComponent:
        """
        Gets current document from ``XSCRIPTCONTEXT``.
        
        This method should be used in macro's

        Raises:
            Exception: if XSCRIPTCONTEXT is not found

        Returns:
            XComponent: current document
        """
        if not "XSCRIPTCONTEXT" in globals():
            raise Exception("XSCRIPTCONTEXT not found")
        XSCRIPTCONTEXT = globals()['XSCRIPTCONTEXT']
        cls._doc = XSCRIPTCONTEXT.getDocument()
        cls._ms_factory = cls.qi(XMultiServiceFactory, cls._doc)
        return cls._doc
    # ======================== document saving ==============

    @staticmethod
    def save(doc: XStorable) -> None:
        """
        Save as document

        Args:
            doc (XStorable): Office document

        Raises:
            Exception: If unable to save document
        """
        try:
            doc.store()
            print("Saved the document by overwriting")
        except IOException as e:
            raise Exception(f"Could not save the document") from e

    @overload
    @classmethod
    def save_doc(cls, doc: XStorable, fnm: str) -> None:
        """
        Save document

        Args:
            doc (XStorable): Office document
            fnm (str): file path to save as
        """
        ...

    @overload
    @classmethod
    def save_doc(cls, doc: XStorable, fnm: str, password: str) -> None:
        """
        Save document

        Args:
            doc (XStorable): Office document
            fnm (str): file path to save as
            password (str): Optional password
        """
        ...

    @overload
    @classmethod
    def save_doc(cls, doc: XStorable, fnm: str, password: str, format: str) -> None:
        """
        Save document

        Args:
            doc (XStorable): Office document
            fnm (str): file path to save as
            password (str): Optional password
            format (str): _description_. Defaults to None.
        """
        ...

    @classmethod
    def save_doc(cls, doc: XStorable, fnm: str, password: str = None, format: str = None) -> None:
        """
        Save document

        Args:
            doc (XStorable): Office document
            fnm (str): file path to save as
            password (str): password
            format (str): document format such as 'odt' or 'xml'
        """
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
    @classmethod
    def store_doc(cls, store: XStorable, doc_type: DocType, fnm: str) -> None:
        """
        Stores/Saves a document

        Args:
            store (XStorable): instance that implements XStorable interface.
            doc_type (DocType): Document type
            fnm (str): Path to save document as. If extension is absent then text (.txt) is assumed.
        """
        ...

    @overload
    @classmethod
    def store_doc(cls, store: XStorable, doc_type: DocType, fnm: str, password: str) -> None:
        """
        Stores/Saves a document

        Args:
            store (XStorable): instance that implements XStorable interface.
            doc_type (DocType): Document type
            fnm (str): Path to save document as. If extension is absent then text (.txt) is assumed.
            password (str): Password for document.
        """
        ...

    @classmethod
    def store_doc(cls, store: XStorable, doc_type: DocType, fnm: str, password: Optional[str] = None) -> None:
        """
        Stores/Saves a document

        Args:
            store (XStorable): instance that implements XStorable interface.
            doc_type (DocType): Document type
            fnm (str): Path to save document as. If extension is absent then text (.txt) is assumed.
            password (str): Password for document.
        """
        ext = mInfo.Info.get_ext(fnm)
        frmt = "Text"
        if ext is not None:
            frmt = cls.ext_to_format(ext=ext, doc_type=doc_type)
        else:
            print("Assuming a text format")
        if password is None:
            cls.store_doc_format(store=store, fnm=fnm, format=frmt)
        else:
            cls.store_doc_format(store=store, fnm=fnm, format=frmt, password=password)

    @overload
    @classmethod
    def ext_to_format(cls, ext: str) -> str:
        """
        Convert the extension string into a suitable office format string.
        The formats were chosen based on the fact that they
        are being used to save (or export) a document.

        Args:
            ext (str): document extension

        Returns:
            str: format of ext such as 'text', 'rtf', 'odt', 'pdf', 'jpg' etc...
            Defaults to 'text' if conversion is unknown.
        """
        ...

    @overload
    @classmethod
    def ext_to_format(cls, ext: str, doc_type: DocType) -> str:
        """
        Convert the extension string into a suitable office format string.
        The formats were chosen based on the fact that they
        are being used to save (or export) a document.

        Args:
            ext (str): document extension
            doc_type (DocType): Type of document.

        Returns:
            str: format of ext such as 'text', 'rtf', 'odt', 'pdf', 'jpg' etc...
            Defaults to 'text' if conversion is unknown.
        """
        ...

    @classmethod
    def ext_to_format(cls, ext: str, doc_type: DocType = DocType.UNKNOWN) -> str:
        """
        Convert the extension string into a suitable office format string.
        The formats were chosen based on the fact that they
        are being used to save (or export) a document.

        Args:
            ext (str): document extension
            doc_type (DocType): Type of document.

        Returns:
            str: format of ext such as 'text', 'rtf', 'odt', 'pdf', 'jpg' etc...
            Defaults to 'text' if conversion is unknown.

        Note:
            ``doc_type`` is used to distinguish between the various meanings of the PDF ext.
            This could be a lot more extensive.

            Use :py:meth:`Info.getFilterNames` to get the filter names for your Office.
        """
        dtype = cls.DocType(doc_type)
        s = ext.lower()
        if s == "doc":
            return "MS Word 97"
        elif s == "docx":
            return "Office Open XML Text"  # MS Word 2007 XML
        elif s == "rtf":
            if dtype == cls.DocType.CALC:
                return "Rich Text Format (StarCalc)"
            else:
                return "Rich Text Format"
        elif s == "odt":
            return "writer8"
        elif s == "ott":
            return "writer8_template"
        elif s == "pdf":
            if dtype == cls.DocType.WRITER:
                return "writer_pdf_Export"
            elif dtype == cls.DocType.IMPRESS:
                return "impress_pdf_Export"
            elif dtype == cls.DocType.DRAW:
                return "draw_pdf_Export"
            elif dtype == cls.DocType.CALC:
                return "calc_pdf_Export"
            elif dtype == cls.DocType.MATH:
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
            if dtype == cls.DocType.IMPRESS:
                return "impress_jpg_Export"
            else:
                return "draw_jpg_Export"
        elif s == "png":
            if dtype == cls.DocType.IMPRESS:
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
            if dtype == cls.DocType.WRITER:
                return "HTML (StarWriter)"
            elif dtype == cls.DocType.IMPRESS:
                return "impress_html_Export"
            elif dtype == cls.DocType.DRAW:
                return "draw_html_Export"
            elif dtype == cls.DocType.CALC:
                return "HTML (StarCalc)"
            else:
                return "HTML"
        elif s == "xhtml":
            if dtype == cls.DocType.WRITER:
                return "XHTML Writer File"
            elif dtype == cls.DocType.IMPRESS:
                return "XHTML Impress File"
            elif dtype == cls.DocType.DRAW:
                return "XHTML Draw File"
            elif dtype == cls.DocType.CALC:
                return "XHTML Calc File"
            else:
                return "XHTML Writer File"
        elif s == "xml":
            if dtype == cls.DocType.WRITER:
                return "OpenDocument Text Flat XML"
            elif dtype == cls.DocType.IMPRESS:
                return "OpenDocument Presentation Flat XML"
            elif dtype == cls.DocType.DRAW:
                return "OpenDocument Drawing Flat XML"
            elif dtype == cls.DocType.CALC:
                return "OpenDocument Spreadsheet Flat XML"
            else:
                return "OpenDocument Text Flat XML"

        else:
            print(f"Do not recognize extension '{ext}'; using text")
            return "Text"

    @overload
    @staticmethod
    def store_doc_format(store: XStorable, fnm: str, format: str) -> None:
        """
        Store document as format.

        Args:
            store (XStorable): instance that implements XStorable interface.
            fnm (str): Path to save document as.
            format (str): document format such as 'odt' or 'xml'

        Raises:
            Exception: If unable to save document
        """
        ...

    @overload
    @staticmethod
    def store_doc_format(store: XStorable, fnm: str, format: str, password: str) -> None:
        """
        Store document as format.

        Args:
            store (XStorable): instance that implements XStorable interface.
            fnm (str): Path to save document as.
            format (str): document format such as 'odt' or 'xml'
            password (str): Password for document.

        Raises:
            Exception: If unable to save document
        """
        ...

    @staticmethod
    def store_doc_format(store: XStorable, fnm: str, format: str, password: str = None) -> None:
        """
        Store document as format.

        Args:
            store (XStorable): instance that implements XStorable interface.
            fnm (str): Path to save document as.
            format (str): document format such as 'odt' or 'xml'
            password (str): Password for document.

        Raises:
            Exception: If unable to save document
        """
        print(f"Saving the document in '{fnm}'")
        print(f"Using format {format}")

        try:
            save_file_url = mFileIO.FileIO.fnm_to_url(fnm)
            if password is None:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=format)
            else:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=format, Password=password)
            store.storeToURL(save_file_url, store_props)
        except IOException as e:
            raise Exception(f"Could not save '{fnm}'") from e

    # ======================== document closing ==============

    @overload
    @classmethod
    def close(cls, closeable: XCloseable) -> None:
        """
        Closes a document.

        Args:
            closeable (XCloseable): Object that implements XCloseable interface.
        """
        ...

    @overload
    @classmethod
    def close(cls, closeable: XCloseable, deliver_ownership: bool) -> None:
        """
        Closes a document.

        Args:
            closeable (XCloseable): Object that implements XCloseable interface.
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException.
        """
        ...

    @classmethod
    def close(cls, closeable: XCloseable, deliver_ownership=False) -> None:
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
        if closeable is None:
            return
        print("Closing the document")
        try:
            closeable.close(deliver_ownership)
            cls._doc = None
        except CloseVetoException as e:
            raise Exception("Close was vetoed") from e

    @overload
    @classmethod
    def close_doc(cls, doc: object) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Closeable doccument
        """
        ...

    @overload
    @classmethod
    def close_doc(cls, doc: object, deliver_ownership: bool) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Closeable doccument
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException.

        Raises:
            MissingInterfaceError: if doc does not have XCloseable interface
        """
        ...

    @classmethod
    def close_doc(cls, doc: object, deliver_ownership=False) -> None:
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

        Raises:
            MissingInterfaceError: if doc does not have XCloseable interface
        """
        try:
            closeable = cls.qi(XCloseable, doc)
            if closeable is None:
                raise mEx.MissingInterfaceError(XCloseable)
            cls.close(closeable=closeable, deliver_ownership=deliver_ownership)
            cls._doc = None
        except DisposedException as e:
            raise Exception("Document close failed since Office link disposed") from e

    # ================= initialization via Addon-supplied context ====================

    @classmethod
    def addon_initialize(cls, addon_xcc: XComponentContext) -> XComponent:
        """
        Initalize and addon

        Args:
            addon_xcc (XComponentContext): Addon component context

        Raises:
            TypeError: If addon_xcc is None
            Exception: If unable to get service manager from addon_xcc 
            Exception: If unable to access desktop
            Exception: If unable to access document
            MissingInterfaceError: if unable to get XMultiServiceFactory interface instance

        Returns:
            XComponent: addon as component
        """
        xcc = addon_xcc
        if xcc is None:
            raise TypeError("'addon_xcc' is null. Could not access component context")
        mc_factory = xcc.getServiceManager()
        if mc_factory is None:
            raise Exception("Office Service Manager is unavailable")

        try:
            xdesktop: XDesktop = mc_factory.createInstanceWithContext("com.sun.star.frame.Desktop", xcc)
        except Exception:
            raise Exception("Could not access desktop")
        doc = xdesktop.getCurrentComponent()
        if doc is None:
            raise Exception("Could not access document")
        cls._ms_factory = cls.qi(XMultiServiceFactory, doc)
        if cls._ms_factory in None:
            raise mEx.MissingInterfaceError(XMultiServiceFactory)
        cls._doc = doc
        return doc

    # ============= initialization via script context ======================

    @classmethod
    def script_initialize(cls, sc: XScriptContext) -> XComponent:
        """
        Initialize script

        Args:
            sc (XScriptContext): Script context

        Raises:
            TypeError: If sc is None
            Exception: if unable to get Component Context from sc
            Exception: If unable to get service manager
            Exception: If unable to access desktop
            Exception: If unable to access document
            MissingInterfaceError: if unable to get XMultiServiceFactory interface instance
        Returns:
            XComponent: script component
        """
        if sc is None:
            raise TypeError("Script Context is null")
        xcc = sc.getComponentContext()
        if xcc is None:
            raise Exception("Could not access component context")
        mc_factory = xcc.getServiceManager()
        if mc_factory is None:
            raise Exception("Office Service Manager is unavailable")
        xdesktop = sc.getDesktop()
        if xdesktop is None:
            raise Exception("Could not access desktop")
        doc = xdesktop.getCurrentComponent()
        if doc is None:
            raise Exception("Could not access document")
        cls._ms_factory = cls.qi(XMultiServiceFactory, doc)
        if cls._ms_factory in None:
            raise mEx.MissingInterfaceError(XMultiServiceFactory)
        cls._doc = doc
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
        """
        Dispacches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as 'GoToCell'. Note: cmd does not containd '.uno:' prefix.
            props (PropertyValue): properties for dispatch

        Raises:
            MissingInterfaceError: If unable to obtain XDispatchHelper instance.
            Exception: If error occurs dispacthing command

        Returns:
            bool: True on success.
        """
        ...

    @overload
    @staticmethod
    def dispatch_cmd(cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> bool:
        """
        Dispacches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as 'GoToCell'. Note: cmd does not containd '.uno:' prefix.
            props (PropertyValue): properties for dispatch
            frame (XFrame): Frame to dispatch to.

        Raises:
            MissingInterfaceError: If unable to obtain XDispatchHelper instance.
            Exception: If error occurs dispacthing command

        Returns:
            bool: True on success.
        """
        ...

    @classmethod
    def dispatch_cmd(cls, cmd: str, props: Iterable[PropertyValue] = None, frame: XFrame = None) -> bool:
        """
        Dispacches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as 'GoToCell'. Note: cmd does not containd '.uno:' prefix.
            props (PropertyValue): properties for dispatch
            frame (XFrame): Frame to dispatch to.

        Raises:
            MissingInterfaceError: If unable to obtain XDispatchHelper instance.
            Exception: If error occurs dispacthing command

        Returns:
            bool: True on success.

        See Also:
            `LibreOffice Dispatch Commands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_
        """
        if props is None:
            props = ()
        if frame is None:
            frame = cls._xdesktop.getCurrentFrame()

        helper = cls.create_instance_mcf(XDispatchHelper, "com.sun.star.frame.DispatchHelper")
        if helper is None:
            raise mEx.MissingInterfaceError(XDispatchHelper, f"Could not create dispatch helper for command {cmd}")
        try:
            helper.executeDispatch(frame, f".uno:{cmd}", "", 0, props)
            return True
        except Exception as e:
            raise Exception(f"Could not dispatch '{cmd}'") from e

    # ================= Uno cmds =========================

    @staticmethod
    def make_uno_cmd(item_name: str) -> str:
        """
        Make a uno command that can be used with :py:meth:`~Lo.extract_item_name`

        Args:
            item_name (str): command item name

        Returns:
            str: uno command string
        """
        return f"vnd.sun.star.script:Foo/Foo.{item_name}?language=Java&location=share"

    @staticmethod
    def extract_item_name(uno_cmd: str) -> str:
        """
        Extract a uno command from a string that was created with :py:meth:`~Lo.make_uno_cmd`

        Args:
            uno_cmd (str): uno command

        Raises:
            ValueError: If unable to extract command

        Returns:
            str: uno command
        """
        try:
            foo_pos = uno_cmd.index("Foo.")
        except ValueError:
            raise ValueError(f"Could not find Foo header in command: '{uno_cmd}'")
        try:
            lang_pos = uno_cmd.index("?language")
        except ValueError:
            raise ValueError(f"Could not find language header in command: '{uno_cmd}'")
        start = foo_pos + 4
        return uno_cmd[start:lang_pos]

    # ======================== use Inspector extensions ====================

    @classmethod
    def inspect(cls, obj: object) -> None:
        """
        Inspects object using ``org.openoffice.InstanceInspector`` inspector.

        Args:
            obj (object): object to inspect.
        """
        if cls._xcc is None or cls._mc_factory is None:
            print("No office connection found")
            return
        try:
            ts = mInfo.Info.get_interface_types(obj)
            title = "Object"
            if ts is not None and len(ts) > 0:
                title = ts[0].getTypeName() + " " + title
            inspector = cls._mc_factory.createInstanceWithContext("org.openoffice.InstanceInspector", cls._xcc)
            #       hands on second use
            if inspector is None:
                print("Inspector Service could not be instantiated")
                return
            print("Inspector Service instantiated")
            intro = cls.create_instance_mcf(XIntrospection, "com.sun.star.beans.Introspection")
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
        """
        call MRI's inspect() to inspect obj.

        Args:
            obj (object): obj to inspect

        Raises:
            Exception: If MRI serivce could not be instantiated.

        See Also:
            `MRI - UNO Object Inspection Tool <https://extensions.libreoffice.org/en/extensions/show/mri-uno-object-inspection-tool>`_
        """
        # Available from http://extensions.libreoffice.org/extension-center/mri-uno-object-inspection-tool
        #               or http://extensions.services.openoffice.org/en/project/MRI
        #  Docs: https://github.com/hanya/MRI/wiki
        #  Forum tutorial: https://forum.openoffice.org/en/forum/viewtopic.php?f=74&t=49294
        xi = cls.create_instance_mcf(XIntrospection, "mytools.Mri")
        if xi is None:
            raise Exception("MRI Inspector Service could not be instantiated")
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
    def is_none_or_empty(s: str) -> bool:
        """
        Gets is a string is None or Empty

        Args:
            s (str): String to check.

        Returns:
            bool: True if None or empty string; Otherwise, False
        """
        return s == None or len(s) == 0

    is_null_or_empty = is_none_or_empty

    @staticmethod
    def wait_enter() -> None:
        """
        Terminal dispalays Press Enter to continue...
        """
        input("Press Enter to continue...")

    @staticmethod
    def is_url(fnm: str) -> bool:
        """
        Gets if a string is a url format.

        Args:
            fnm (str): string to check.

        Returns:
            bool: True if Url format; Otherwise, False
        """
        # https://stackoverflow.com/questions/7160737/how-to-validate-a-url-in-python-malformed-or-not
        try:
            result = urlparse(fnm)
            return all([result.scheme, result.netloc])
        except ValueError:
            return False

    # endregion document opening

    @staticmethod
    def capitalize(s: str) -> str:
        """
        Capitalizes a string

        Args:
            s (str): String to capitalize

        Returns:
            str: string capitalized.
        """
        return s.capitalize()

    @staticmethod
    def parse_int(s: str) -> int:
        """
        Convets string into int.

        Args:
            s (str): string to parse

        Returns:
            int: String as int. If unable to convert s to int then 0 is returned.
        """
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
        """
        Prints names to console

        Args:
            names (Iterable[str]): names to print
        """
        ...

    @overload
    @staticmethod
    def print_names(names: Iterable[str], num_per_line: int) -> None:
        """
        Prints names to console

        Args:
            names (Iterable[str]): names to print
            num_per_line (int): Number of names per line.
        """
        ...

    @staticmethod
    def print_names(names: Iterable[str], num_per_line: int = 0) -> None:
        """
        Prints names to console

        Args:
            names (Iterable[str]): names to print
            num_per_line (int): Number of names per line.
        """
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
    def print_table(name: str, table: Iterable[Iterable[object]]) -> None:
        """
        Prints table to console

        Args:
            name (str): Name of table
            table (List[List[str]]): Table Data
        """
        print(f"-- {name} ----------------")
        for row in table:
            col_str = "  ".join([str(el) for el in row])
            print(col_str)
        print()

    @staticmethod
    def get_container_names(con: XIndexAccess) -> List[str] | None:
        """
        Gets container names

        Args:
            con (XIndexAccess): container

        Returns:
            List[str] | None: Containor name is found; Otherwise, None
        """
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

    @classmethod
    def find_container_props(cls, con: XIndexAccess, nm: str) -> XPropertySet | None:
        """
        Find as Property Set in a container

        Args:
            con (XIndexAccess): Container to search
            nm (str): Name of property to searh for

        Raises:
            TypeError: if con is None

        Returns:
            XPropertySet | None: Found property set; Otherwise, None
        """
        if con is None:
            raise TypeError("Container is null")
        for i in range(con.getCount()):
            try:
                el = con.getByIndex(i)
                named = cls.qi(XNamed, el)
                if named and named.getName() == nm:
                    return cls.qi(XPropertySet, el)
            except Exception:
                print(f"Could not access element {i}")
        print(f"Could not find a '{nm}' property set in the container")
        return None

    @classmethod
    def is_uno_interfaces(cls, component:object, *args:str | UnoInterface) -> bool:
        """
        Gets if an object contains interface(s)

        Args:
            component (object): object to check for supplied interfaces
            args (str | UnoInterface): one or more strings such as 'com.sun.star.uno.XInterface'
                or Any uno interface that Starts with X such has XEnumTypeDescription

        Returns:
            bool: True if component contains all supplied interfaces; Otherwise, False
        """
        if len(args) == 0:
            return False
        result = True
        for arg in args:
            try:
                if isinstance(arg, str):
                    t = uno.getTypeByName(arg)
                else:
                    t = arg
                obj = cls.qi(t, component)
                if obj is None:
                    result = False
                    break
            except Exception:
                result = False
                break
        return result

    # ------------------- date --------------------
    @staticmethod
    def get_time_stamp() -> str:
        """Gets a time stamp"""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @classproperty
    def null_date(cls) -> datetime:
        """
        Gets Value of Null Date

        Returns:
            datetime: Null Date on sucess; Otherwise, None

        Note:
            If Lo has no document to determine date from then a
            default date of 1889/12/30 is returned.
        """
        # https://tinyurl.com/2pdrt5z9#NullDate
        try:
            return cls.__null_date
        except AttributeError:
            if cls._doc is None:
                print("No document found returning 1889/12/30")
                return datetime(year=1889, month=12, day=30)
            n_supplier = cls.qi(XNumberFormatsSupplier, cls._doc)
            number_settings = n_supplier.getNumberFormatSettings()
            d = number_settings.getPropertyValue("NullDate")
            cls.__null_date = datetime(d.Year, d.Month, d.Day)
        return cls.__null_date

    @null_date.setter
    def null_date(cls, value) -> None:
        # raise error on set. Not really neccesary but gives feedback.
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.name)
