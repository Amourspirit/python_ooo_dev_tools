# coding: utf-8
# Python conversion of Lo.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

from __future__ import annotations
from datetime import datetime, timezone
import time
import types
from typing import TYPE_CHECKING, Any, Iterable, Optional, List, Tuple, cast, overload, Type
from urllib.parse import urlparse
import uno
from enum import IntEnum, Enum

from ..mock import mock_g

from ..events.event_singleton import _Events
from ..events.lo_named_event import LoNamedEvent
from ..events.gbl_named_event import GblNamedEvent
from ..events.args.event_args import EventArgs
from ..events.args.cancel_event_args import CancelEventArgs
from ..events.args.dispatch_args import DispatchArgs
from ..events.args.dispatch_cancel_args import DispatchCancelArgs
from ..meta.static_meta import StaticProperty, classproperty
from ..conn.connect import ConnectBase, LoPipeStart, LoSocketStart, LoDirectStart
from ..conn import connectors
from ..conn import cache as mCache
from ..listeners.x_event_adapter import XEventAdapter

from com.sun.star.lang import XComponent

# if not mock_g.DOCS_BUILDING:
# not importing for doc building just result in short import name for
# args that use these.
# this is also true becuase docs/conf.py ignores com import for autodoc
from com.sun.star.beans import XPropertySet
from com.sun.star.beans import XIntrospection
from com.sun.star.container import XNamed
from com.sun.star.frame import XDesktop
from com.sun.star.frame import XDispatchHelper
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.io import IOException
from com.sun.star.util import XCloseable
from com.sun.star.util import XNumberFormatsSupplier
from com.sun.star.frame import XComponentLoader
from com.sun.star.frame import XModel
from com.sun.star.frame import XStorable

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XChild
    from com.sun.star.container import XIndexAccess
    from com.sun.star.frame import XFrame
    from com.sun.star.lang import EventObject
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.lang import XComponent
    from com.sun.star.script.provider import XScriptContext
    from com.sun.star.uno import XComponentContext
    from com.sun.star.uno import XInterface


from ooo.dyn.document.macro_exec_mode import MacroExecMode  # const
from ooo.dyn.lang.disposed_exception import DisposedException
from ooo.dyn.util.close_veto_exception import CloseVetoException

# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from . import script_context

from . import props as mProps
from . import file_io as mFileIO
from . import xml_util as mXML
from . import info as mInfo

from ..exceptions import ex as mEx
from .type_var import PathOrStr, UnoInterface, T

# PathOrStr = type_var.PathOrStr
# """Path like object or string"""


class Lo(metaclass=StaticProperty):
    class ControllerLock:
        """
        Context manager for Locking Controller

        In the following example ControllerLock is called using ``with``.

        All code inside the ``with Lo.ControllerLock() as xdoc`` block is written to **Writer**
        with controller locked. This means the ui will not update until the block is done.
        A soon as the block is processed the controller is unlocked and the ui is updated.

        Can be useful for large writes in document. Will give a speed improvement.

        Example:

            .. code::

                with Lo.ControllerLock() as xdoc:
                    cursor = Write.get_cursor(xdoc)
                    Write.append(cursor=cursor, text="Some examples of simple text ")
                    # do a bunch or work.
                    ...
        """

        def __init__(self):
            self.component = Lo.this_component
            Lo.lock_controllers()

        def __enter__(self) -> XComponent:
            return self.component

        def __exit__(self, exc_type, exc_val, exc_tb):
            Lo.unlock_controllers()

    class Loader:
        """
        Context Manager for Loader

        Example:

            .. code::

                with Lo.Loader(Lo.ConnectSocket()) as loader:
                    doc = Write.create_doc(loader)
                    ...

        See Also:
            :ref:`ch02`
        """

        def __init__(
            self,
            connector: connectors.ConnectPipe | connectors.ConnectSocket | None,
            cache_obj: mCache.Cache | None = None,
        ):
            """
            Create a connection to office

            Args:
                connector (connectors.ConnectPipe | connectors.ConnectSocket | None): Connection information. Ignore for macros.
                cache_obj (mCache.Cache | None, optional): Cache instance that determines if LibreOffice profile is to be copied and cached
                    Ignore for macros. Defaults to None.
            """
            self.loader = Lo.load_office(connector=connector, cache_obj=cache_obj)

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

    # endregion docType ints

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

    # endregion docType strings

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
    # endregion port

    ConnectPipe = connectors.ConnectPipe
    """Alias of connectors.ConnectPipe"""
    ConnectSocket = connectors.ConnectSocket
    """Alias of connectors.ConnectSocket"""

    _xcc: XComponentContext = None
    _doc: XComponent = None
    """remote component context"""
    _xdesktop: XDesktop = None
    """remote desktop UNO service"""

    _mc_factory: XMultiComponentFactory = None
    _ms_factory: XMultiServiceFactory = None

    _is_office_terminated: bool = False

    _lo_inst: ConnectBase = None

    # region    qi()

    @overload
    @staticmethod
    def qi(atype: Type[T], obj: XTypeProvider) -> T | None:
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface
            obj (object): Object that implements interface.

        Returns:
            T | None: instance of interface if supported; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def qi(atype: Type[T], obj: XTypeProvider, raise_err: bool) -> T | None:
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface
            obj (object): Object that implements interface.
            raise_err (bool, optional): If True then raises MissingInterfaceError if result is None. Default False

        Raises:
            MissingInterfaceError: If 'raise_err' is 'True' and result is None

        Returns:
            T | None: instance of interface if supported; Otherwise, None
        """
        ...

    @staticmethod
    def qi(atype: Type[T], obj: XTypeProvider, raise_err: bool = False) -> T | None:
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type to query obj for. Any Uno class that starts with 'X' such as XInterface
            obj (object): Object that implements interface.
            raise_err (bool, optional): If True then raises MissingInterfaceError if result is None. Default False

        Raises:
            MissingInterfaceError: If 'raise_err' is 'True' and result is None

        Returns:
            T | None: instance of interface if supported; Otherwise, None

        Note:
            When ``raise_err=True`` return value will never be ``None``.

        Example:

            .. code-block:: python
                :emphasize-lines: 3

                from com.sun.star.util import XSearchable
                cell_range = ...
                srch = Lo.qi(XSearchable, cell_range)
                sd = srch.createSearchDescriptor()
        """
        result = None
        if uno.isInterface(atype) and hasattr(obj, "queryInterface"):
            uno_t = uno.getTypeByName(atype.__pyunointerface__)
            result = obj.queryInterface(uno_t)
        if raise_err is True and result is None:
            raise mEx.MissingInterfaceError(atype)
        return result

    # endregion qi()

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
    @classmethod
    def create_instance_msf(
        cls, atype: Type[T], service_name: str, msf: XMultiServiceFactory = None, raise_err: bool = False
    ) -> T:
        """
        Creates an instance classified by the specified service name and
        optionally passes arguments to that instance.

        The interface specified by ``atype`` is returned from created instance.

        Args:
            atype (Type[T]): Type of interface to return from created service.
                Any Uno class that starts with 'X' such as XInterface
            service_name (str): Service name
            msf (XMultiServiceFactory, optional): Multi service factory used to create instance
            raise_err (bool, optional): If True then can raise CreateInstanceMsfError or MissingInterfaceError

        Raises:
            CreateInstanceMsfError: If ``raise_err`` is ``True`` and no instance was created
            MissingInterfaceError: If ``raise_err`` is ``True`` and instance was created but does not implement ``atype`` interface.
            Exception: if unable to create instance for any other reason

        Returns:
            T: Instance of interface for the service name.

        Note:
            When ``raise_err=True`` return value will never be ``None``.

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
            if raise_err is True and obj is None:
                mEx.CreateInstanceMsfError(atype, service_name)
            interface_obj = cls.qi(atype=atype, obj=obj)
            if raise_err is True and interface_obj is None:
                raise mEx.MissingInterfaceError(atype)
            return interface_obj
        except mEx.CreateInstanceMsfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception(f"Couldn't create interface for '{service_name}'") from e

    # endregion create_instance_msf()

    # region    create_instance_mcf()
    @classmethod
    def create_instance_mcf(
        cls, atype: Type[T], service_name: str, args: Tuple[object, ...] | None = None, raise_err: bool = False
    ) -> T:
        """
        Creates an instance of a component which supports the services specified by the factory,
        and optionally initializes the new instance with the given arguments and context.

        The interface specified by ``atype`` is returned from created instance.

        Args:
            atype (Type[T]): Type of interface to return from created instance.
                Any Uno class that starts with ``X`` such as ``XInterface``
            service_name (str): Service Name
            args (Tuple[object, ...], Optional): Args
            raise_err (bool, optional): If True then can raise CreateInstanceMcfError or MissingInterfaceError

        Raises:
            CreateInstanceMcfError: If ``raise_err`` is ``True`` and no instance was created
            MissingInterfaceError: If ``raise_err`` is ``True`` and instance was created but does not implement ``atype`` interface.
            Exception: if unable to create instance for any other reason

        Note:
            When ``raise_err=True`` return value will never be ``None``.

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
            if raise_err is True and obj is None:
                mEx.CreateInstanceMcfError(atype, service_name)
            interface_obj = cls.qi(atype=atype, obj=obj)
            if raise_err is True and interface_obj is None:
                raise mEx.MissingInterfaceError(atype)
            return interface_obj
        except mEx.CreateInstanceMcfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception(f"Couldn't create interface for '{service_name}'") from e

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

    @classmethod
    def load_office(
        cls,
        connector: connectors.ConnectPipe | connectors.ConnectSocket | None = None,
        cache_obj: mCache.Cache | None = None,
    ) -> XComponentLoader:
        """
        Loads Office

        Not available in a macro.

        If running outside of office then a bridge is created that connects to office.

        If running from inside of office e.g. in a macro, then ``Lo.XSCRIPTCONTEXT`` is used.
        ``using_pipes`` is ignored with running inside office.

        Args:
            connector (connectors.ConnectPipe | connectors.ConnectSocket | None): Connection information. Ignore for macros.
            cache_obj (Cache | None, optional): Cache instance that determines of LibreOffice profile is to be copied and cached
                Ignore for macros. Defaults to None.


        Raises:
            CancelEventError: If office_loading event is canceled
            Exception: If run outside a macro
            Exception: If unable to get access to XComponentLoader.

        Returns:
            XComponentLoader: component loader

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.OFFICE_LOADING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.OFFICE_LOADED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            - :py:meth:`open_doc`
            - :py:class:`.Lo.Loader`
            - :ref:`ch02`

        Example:

            .. code::

                loader =  Lo.Loader(Lo.ConnectSocket()):
                doc = Write.create_doc(loader)
                ...
        """
        if mock_g.DOCS_BUILDING:
            # some component call this method and are triggered during docs building.
            # by adding this block this method will be exited if docs are building.
            return None

        # Creation sequence: remote component content (xcc) -->
        #                     remote service manager (mcFactory) -->
        #                     remote desktop (xDesktop) -->
        #                     component loader (XComponentLoader)
        # Once we have a component loader, we can load a document.
        # xcc, mcFactory, and xDesktop are stored as static globals.

        cargs = CancelEventArgs(Lo.load_office.__qualname__)

        cargs.event_data = {
            "connector": connector,
        }

        eargs = EventArgs.from_args(cargs)
        _Events().trigger(LoNamedEvent.RESET, eargs)

        _Events().trigger(LoNamedEvent.OFFICE_LOADING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        b_connector = cargs.event_data["connector"]

        Lo.print("Loading Office...")
        if b_connector is None:
            try:
                cls._lo_inst = LoDirectStart()
                cls._lo_inst.connect()
            except Exception as e:
                Lo.print("Office context could not be created. A connector must be supplied if not running as a macro")
                Lo.print(f"    {e}")
                raise SystemExit(1)
        elif isinstance(b_connector, connectors.ConnectPipe):
            try:
                cls._lo_inst = LoPipeStart(connector=b_connector, cache_obj=cache_obj)
                cls._lo_inst.connect()
            except Exception as e:
                Lo.print("Office context could not be created")
                Lo.print(f"    {e}")
                raise SystemExit(1)
        elif isinstance(b_connector, connectors.ConnectSocket):
            try:
                cls._lo_inst = LoSocketStart(connector=b_connector, cache_obj=cache_obj)
                cls._lo_inst.connect()
            except Exception as e:
                Lo.print("Office context could not be created")
                Lo.print(f"    {e}")
                raise SystemExit(1)
        else:
            Lo.print("Invalid Connector type. Fatal Error.")
            raise SystemExit(1)

        cls._xcc = cls._lo_inst.ctx
        cls._mc_factory = cls._xcc.getServiceManager()
        if cls._mc_factory is None:
            Lo.print("Office Service Manager is unavailable")
            raise SystemExit(1)
        cls._xdesktop = cls.create_instance_mcf(XDesktop, "com.sun.star.frame.Desktop")
        if cls._xdesktop is None:
            # OPTIMIZE: Perhaps system exit is not the best what to handle no desktop service
            Lo.print("Could not create a desktop service")
            raise SystemExit(1)
        loader = cls.qi(XComponentLoader, cls._xdesktop)
        if loader is None:
            Lo.print("Unable to access XComponentLoader")
            SystemExit(1)
        _Events().trigger(LoNamedEvent.OFFICE_LOADED, eargs)
        return loader
        # return cls.xdesktop

    # endregion Start Office

    # region office shutdown
    @classmethod
    def close_office(cls) -> bool:
        """
        Closes the office connection.

        Returns:
            bool: True if office is closed; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.OFFICE_CLOSING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.OFFICE_CLOSED` :eventref:`src-docs-event`
        """
        Lo.print("Closing Office")

        cargs = CancelEventArgs(Lo.close_office.__qualname__)
        _Events().trigger(LoNamedEvent.OFFICE_CLOSING, cargs)
        if cargs.cancel:
            return False

        cls._doc = None
        if cls._xdesktop is None:
            cls.print("No office connection found")
            return True

        if cls._is_office_terminated:
            cls.print("Office has already been requested to terminate")
            return cls._is_office_terminated
        num_tries = 1
        start = time.time()
        elapsed = 0
        seconds = 10
        while cls._is_office_terminated is False and elapsed < seconds:
            elapsed = time.time() - start
            cls._is_office_terminated = cls._try_to_terminate(num_tries)
            time.sleep(0.5)
            num_tries += 1
        if cls._is_office_terminated:
            eargs = EventArgs.from_args(cargs)
            _Events().trigger(LoNamedEvent.OFFICE_CLOSED, eargs)
            _Events().trigger(LoNamedEvent.RESET, eargs)
        return cls._is_office_terminated

    @classmethod
    def _try_to_terminate(cls, num_tries: int) -> bool:
        try:
            is_dead = cls._xdesktop.terminate()
            if is_dead:
                if num_tries > 1:
                    cls.print(f"{num_tries}. Office terminated")
                else:
                    cls.print("Office terminated")
            else:
                cls.print(f"{num_tries}. Office failed to terminate")
            return is_dead
        except DisposedException as e:
            cls.print("Office link disposed")
            return True
        except Exception as e:
            cls.print(f"Termination exception: {e}")
            return False

    @classmethod
    def kill_office(cls) -> None:
        """
        Kills the office connection.

        See Also:
            :py:meth:`~Lo.close_office`
        """

        if cls._lo_inst is None:
            cls.print("No instance to kill")
            return
        try:
            # raised a NotImplementedError when cls._lo_inst is direct (macro mode)
            cls._lo_inst.kill_soffice()
            cls._is_office_terminated = True
            eargs = EventArgs(Lo.kill_office.__qualname__)
            _Events().trigger(LoNamedEvent.OFFICE_CLOSED, eargs)
            _Events().trigger(LoNamedEvent.RESET, eargs)
            cls.print("Killed Office")
        except Exception as e:
            raise Exception(f"Unbale to kill Office") from e

    # endregion office shutdown

    # region document opening
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr, doc_type: Lo.DocType, loader: XComponentLoader) -> XComponent:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader): Component loader

        Returns:
            XComponent: Document

        See Also:
            - :py:meth:`~Lo.open_doc`
            - :py:meth:`~Lo.open_readonly_doc`
            - :ref:`ch02_open_doc`

        Attention:
            :py:meth:`~.utils.lo.Lo.open_doc` method is called along with any of its events.
        """
        nn = mXML.XML.get_flat_fiter_name(doc_type=doc_type)
        Lo.print(f"Flat filter Name: {nn}")
        # do not set Hidden=True property here.
        # there is a strange error that pops up conditionally and it seems
        # to be remedied by not seting Hidden=True
        # see comments in tests.text_xml.test_in_filters.test_transform_clubs()
        return cls.open_doc(fnm, loader, mProps.Props.make_props(FilterName=nn))

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: if unable to open document

        Returns:
            XComponent: Document
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
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
        fnm: PathOrStr,
        loader: XComponentLoader,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader

        Raises:
            Exception: if unable to open document
            CancelEventError: if DOC_OPENING event is canceled.

        Returns:
            XComponent: Document

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            - :py:meth:`~Lo.open_readonly_doc`
            - :py:meth:`~Lo.open_flat_doc`
            - :py:meth:`load_office`
            - :ref:`ch02_open_doc`

        Example:
            .. code-block:: python

                from ooodev.utils.lo import Lo

                # connect to office
                with Lo.Loader() as loader:
                    doc = Lo.open_doc("/home/user/fancy.odt", loader)
                    ...
        """
        # Props and FileIO are called this method so triger global_reset first.
        cargs = CancelEventArgs(Lo.open_doc.__qualname__)
        cargs.event_data = {
            "fnm": fnm,
            "loader": loader,
            "props": props,
        }
        eargs = EventArgs.from_args(cargs)
        _Events().trigger(LoNamedEvent.RESET, eargs)
        _Events().trigger(LoNamedEvent.DOC_OPENING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        fnm = cargs.event_data["fnm"]

        if fnm is None:
            raise Exception("Filename is null")
        pth = mFileIO.FileIO.get_absolute_path(fnm)

        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        open_file_url = None
        if not mFileIO.FileIO.is_openable(pth):
            if cls.is_url(pth):
                Lo.print(f"Will treat filename as a URL: '{pth}'")
                open_file_url = pth
            else:
                raise Exception(f"Unable to get url from file: {pth}")
        else:
            Lo.print(f"Opening {pth}")
            open_file_url = mFileIO.FileIO.fnm_to_url(pth)

        try:
            doc = loader.loadComponentFromURL(open_file_url, "_blank", 0, props)
            cls._ms_factory = cls.qi(XMultiServiceFactory, doc)
            cls._doc = doc
            _Events().trigger(LoNamedEvent.DOC_OPENED, eargs)
            return doc
        except Exception as e:
            raise Exception("Unable to open the document") from e

    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: if unable to open document

        Returns:
            XComponent: Document

        See Also:
            - :py:meth:`~Lo.open_doc`
            - :py:meth:`~Lo.open_flat_doc`
            - :ref:`ch02_open_doc`

        Attention:
            :py:meth:`~.utils.lo.Lo.open_doc` method is called along with any of its events.
        """
        return cls.open_doc(fnm, loader, mProps.Props.make_props(Hidden=True, ReadOnly=True))

    # ======================== document creation ==============

    @classmethod
    def ext_to_doc_type(cls, ext: str) -> Lo.DocTypeStr:
        """
        Gets document type from extension

        Args:
            ext (str): extension used for lookup

        Returns:
            DocTypeStr: DocTypeStr enum. If not match if found defaults to ``DocTypeStr.WRITER``

        See Also:
            :ref:`ch02_save_doc`
        """
        e = ext.casefold().lstrip(".")
        if e == "":
            Lo.print("Empty string: Using writer")
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
            Lo.print(f"Do not recognize extension '{ext}'; using writer")
            return cls.DocTypeStr.WRITER

    @classmethod
    def doc_type_str(cls, doc_type_val: Lo.DocType) -> Lo.DocTypeStr:
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
            Lo.print(f"Do not recognize extension '{doc_type_val}'; using writer")
            return cls.DocTypeStr.WRITER

    @overload
    @classmethod
    def create_doc(cls, doc_type: DocTypeStr, loader: XComponentLoader) -> XComponent:
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
        doc_type: Lo.DocTypeStr,
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

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_CREATING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_CREATED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            :ref:`ch02_create_doc`
        """
        # Props is called in this method so trigger global_reset first
        cargs = CancelEventArgs(Lo.create_doc.__qualname__)
        cargs.event_data = {
            "doc_type": doc_type,
            "loader": loader,
            "props": props,
        }
        eargs = EventArgs.from_args(cargs)
        _Events().trigger(LoNamedEvent.RESET, eargs)
        _Events().trigger(LoNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        dtype = Lo.DocTypeStr(cargs.event_data["doc_type"])
        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        Lo.print(f"Creating Office document {dtype}")
        try:
            doc = loader.loadComponentFromURL(f"private:factory/{dtype}", "_blank", 0, props)
            cls._ms_factory = cls.qi(XMultiServiceFactory, doc)
            if cls._ms_factory is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            cls._doc = doc
            _Events().trigger(LoNamedEvent.DOC_CREATED, eargs)
            return cls._doc
        except Exception as e:
            raise Exception("Could not create a document") from e

    @classmethod
    def create_macro_doc(cls, doc_type: Lo.DocTypeStr, loader: XComponentLoader) -> XComponent:
        """
        Create a document that allows executing of macros

        Args:
            doc_type (DocTypeStr): Document type
            loader (XComponentLoader): Component Loader

        Returns:
            XComponent: document as component.

        Attention:
            :py:meth:`~.utils.lo.Lo.create_doc` method is called along with any of its events.

        See Also:
            :ref:`ch02_create_doc`
        """
        return cls.create_doc(
            doc_type=doc_type,
            loader=loader,
            props=mProps.Props.make_props(Hidden=False, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE),
        )

    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr, loader: XComponentLoader) -> XComponent:
        """
        Create a document form a template

        Args:
            template_path (PathOrStr): path to template file
            loader (XComponentLoader): Component Loader

        Raises:
            Exception: If unable to create document.

        Returns:
            XComponent: document as component.
        """
        cargs = CancelEventArgs(Lo.create_doc_from_template.__qualname__)
        _Events().trigger(LoNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        if not mFileIO.FileIO.is_openable(template_path):
            raise Exception(f"Template file can not be opened: '{template_path}'")
        Lo.print(f"Opening template: '{template_path}'")
        template_url = mFileIO.FileIO.fnm_to_url(fnm=template_path)

        props = mProps.Props.make_props(Hidden=True, AsTemplate=True)
        try:
            cls._doc = loader.loadComponentFromURL(template_url, "_blank", 0, props)
            cls._ms_factory = cls.qi(XMultiServiceFactory, cls._doc)
            if cls._ms_factory is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            _Events().trigger(LoNamedEvent.DOC_CREATED, EventArgs.from_args(cargs))
            return cls._doc
        except Exception as e:
            raise Exception(f"Could not create document from template") from e

    # ======================== document saving ==============

    @classmethod
    def save(cls, doc: object) -> bool:
        """
        Save as document

        Args:
            doc (object): Office document

        Raises:
            Exception: If unable to save document
            MissingInterfaceError: If doc does not implement XStorable interface

        Returns:
            bool: False if DOC_SAVING event is canceled; Otherwise, True

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_SAVING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_SAVED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing ``doc``.
        """
        cargs = CancelEventArgs(Lo.save.__qualname__)
        cargs.event_data = {"doc": doc}
        _Events().trigger(LoNamedEvent.DOC_SAVING, cargs)
        if cargs.cancel:
            return False

        store = cls.qi(XStorable, doc, True)
        try:
            store.store()
            cls.print("Saved the document by overwriting")
        except IOException as e:
            raise Exception(f"Could not save the document") from e

        _Events().trigger(LoNamedEvent.DOC_SAVED, EventArgs.from_args(cargs))
        return True

    # region    save_doc()

    @overload
    @classmethod
    def save_doc(cls, doc: object, fnm: PathOrStr) -> bool:
        """
        Save document

        Args:
            doc (object): Office document
            fnm (PathOrStr): file path to save as

        Returns:
            bool: False if DOC_SAVING event is canceled; Otherwise, True
        """
        ...

    @overload
    @classmethod
    def save_doc(cls, doc: object, fnm: PathOrStr, password: str) -> bool:
        """
        Save document

        Args:
            doc (object): Office document
            fnm (PathOrStr): file path to save as
            password (str): Optional password


        Returns:
            bool: False if DOC_SAVING event is canceled; Otherwise, True
        """
        ...

    @overload
    @classmethod
    def save_doc(cls, doc: object, fnm: PathOrStr, password: str, format: str) -> bool:
        """
        Save document

        Args:
            doc (object): Office document
            fnm (PathOrStr): file path to save as
            password (str): Optional password
            format (str): _description_. Defaults to None.

        Returns:
            bool: False if DOC_SAVING event is canceled; Otherwise, True
        """
        ...

    @classmethod
    def save_doc(cls, doc: object, fnm: PathOrStr, password: str = None, format: str = None) -> bool:
        """
        Save document

        Args:
            doc (object): Office document
            fnm (PathOrStr): file path to save as
            password (str): password
            format (str): document format such as 'odt' or 'xml'

        Raises:
            MissingInterfaceError: If doc does not implement XStorable interface

        Returns:
            bool: False if DOC_SAVING event is canceled; Otherwise, True

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_SAVING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_SAVED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.

        Attention:
            :py:meth:`~.utils.lo.Lo.store_doc` method is called along with any of its events.

        See Also:
            :ref:`ch02_save_doc`
        """
        cargs = CancelEventArgs(Lo.save_doc.__qualname__)
        cargs.event_data = {
            "doc": doc,
            "fnm": fnm,
            "password": password,
            "format": format,
        }

        fnm = cargs.event_data["fnm"]
        password = cargs.event_data["password"]
        format = cargs.event_data["format"]

        _Events().trigger(LoNamedEvent.DOC_SAVING, cargs)
        if cargs.cancel:
            return False
        store = cls.qi(XStorable, doc, True)
        doc_type = mInfo.Info.report_doc_type(doc)
        kargs = {"fnm": fnm, "store": store, "doc_type": doc_type}
        if password is not None:
            kargs["password"] = password
        if format is None:
            result = cls.store_doc(**kargs)
        else:
            kargs["format"] = format
            result = cls.store_doc_format(**kargs)
        if result:
            _Events().trigger(LoNamedEvent.DOC_SAVED, EventArgs.from_args(cargs))
        return result

    # endregion save_doc()

    # region    store_doc()

    @overload
    @classmethod
    def store_doc(cls, store: XStorable, doc_type: DocType, fnm: PathOrStr) -> bool:
        """
        Stores/Saves a document

        Args:
            store (XStorable): instance that implements XStorable interface.
            doc_type (DocType): Document type
            fnm (PathOrStr): Path to save document as. If extension is absent then text (.txt) is assumed.

        Returns:
            bool: True if document is saved; Otherwise False
        """
        ...

    @overload
    @classmethod
    def store_doc(cls, store: XStorable, doc_type: DocType, fnm: PathOrStr, password: str) -> bool:
        """
        Stores/Saves a document

        Args:
            store (XStorable): instance that implements XStorable interface.
            doc_type (DocType): Document type
            fnm (PathOrStr): Path to save document as. If extension is absent then text (.txt) is assumed.
            password (str): Password for document.

        Returns:
            bool: True if document is saved; Otherwise False
        """
        ...

    @classmethod
    def store_doc(cls, store: XStorable, doc_type: Lo.DocType, fnm: PathOrStr, password: Optional[str] = None) -> bool:
        """
        Stores/Saves a document

        Args:
            store (XStorable): instance that implements XStorable interface.
            doc_type (DocType): Document type
            fnm (PathOrStr): Path to save document as. If extension is absent then text ``.txt`` is assumed.
            password (str): Password for document.

        Returns:
            bool: True if document is saved; Otherwise False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_STORING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_STORED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            - :py:meth:`~.Lo.store_doc_format`
            - :ref:`ch02_save_doc`
        """
        cargs = CancelEventArgs(Lo.store_doc.__qualname__)
        cargs.event_data = {
            "store": store,
            "doc_type": doc_type,
            "fnm": fnm,
            "password": password,
        }
        _Events().trigger(LoNamedEvent.DOC_STORING, cargs)
        if cargs.cancel:
            return False
        ext = mInfo.Info.get_ext(fnm)
        frmt = "Text"
        if ext is None:
            Lo.print("Assuming a text format")
        else:
            frmt = cls.ext_to_format(ext=ext, doc_type=doc_type)
        if password is None:
            cls.store_doc_format(store=store, fnm=fnm, format=frmt)
        else:
            cls.store_doc_format(store=store, fnm=fnm, format=frmt, password=password)
        _Events().trigger(LoNamedEvent.DOC_STORED, EventArgs.from_args(cargs))
        return True

    # endregion  store_doc()

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
    def ext_to_format(cls, ext: str, doc_type: Lo.DocType = DocType.UNKNOWN) -> str:
        """
        Convert the extension string into a suitable office format string.
        The formats were chosen based on the fact that they
        are being used to save (or export) a document.

        Args:
            ext (str): document extension
            doc_type (DocType): Type of document.

        Returns:
            str: format of ext such as ``text``, ``rtf``, ``odt``, ``pdf``, ``jpg`` etc...
            Defaults to ``text`` if conversion is unknown.

        Note:
            ``doc_type`` is used to distinguish between the various meanings of the ``PDF`` ext.
            This could be a lot more extensive.

            Use :py:meth:`Info.getFilterNames` to get the filter names for your Office.

        See Also:
            - :ref:`ch02_save_doc`
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
            Lo.print(f"Do not recognize extension '{ext}'; using text")
            return "Text"

    # region    store_doc_format()

    @overload
    @classmethod
    def store_doc_format(cls, store: XStorable, fnm: PathOrStr, format: str) -> bool:
        """
        Store document as format.

        Args:
            store (XStorable): instance that implements XStorable interface.
            fnm (PathOrStr): Path to save document as.
            format (str): document format such as 'odt' or 'xml'

        Raises:
            Exception: If unable to save document

        Returns:
            bool: True if document is stored; Otherwise False
        """
        ...

    @overload
    @classmethod
    def store_doc_format(cls, store: XStorable, fnm: PathOrStr, format: str, password: str) -> bool:
        """
        Store document as format.

        Args:
            store (XStorable): instance that implements XStorable interface.
            fnm (PathOrStr): Path to save document as.
            format (str): document format such as 'odt' or 'xml'
            password (str): Password for document.

        Raises:
            Exception: If unable to save document

        Returns:
            bool: True if document is stored; Otherwise False
        """
        ...

    @classmethod
    def store_doc_format(cls, store: XStorable, fnm: PathOrStr, format: str, password: str = None) -> bool:
        """
        Store document as format.

        Args:
            store (XStorable): instance that implements XStorable interface.
            fnm (PathOrStr): Path to save document as.
            format (str): document format such as 'odt' or 'xml'
            password (str): Password for document.

        Raises:
            Exception: If unable to save document

        Returns:
            bool: True if document is stored; Otherwise False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_STORING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_STORED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            :py:meth:`~.Lo.store_doc`
        """
        cargs = CancelEventArgs(Lo.store_doc_format.__qualname__)
        cargs.event_data = {
            "store": store,
            "format": format,
            "fnm": fnm,
            "password": password,
        }
        _Events().trigger(LoNamedEvent.DOC_STORING, cargs)
        if cargs.cancel:
            return False
        pth = mFileIO.FileIO.get_absolute_path(cargs.event_data["fnm"])
        fmt = str(cargs.event_data["format"])
        Lo.print(f"Saving the document in '{pth}'")
        Lo.print(f"Using format {fmt}")

        try:
            save_file_url = mFileIO.FileIO.fnm_to_url(pth)
            if password is None:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=fmt)
            else:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=fmt, Password=password)
            store.storeToURL(save_file_url, store_props)
        except IOException as e:
            raise Exception(f"Could not save '{pth}'") from e
        _Events().trigger(LoNamedEvent.DOC_STORED, EventArgs.from_args(cargs))
        return True

    # endregion store_doc_format()

    # ======================== document closing ==============

    @overload
    @classmethod
    def close(cls, closeable: XCloseable) -> bool:
        """
        Closes a document.

        Args:
            closeable (XCloseable): Object that implements XCloseable interface.

        Returns:
            bool: True if Closed; Otherwise, False
        """
        ...

    @overload
    @classmethod
    def close(cls, closeable: XCloseable, deliver_ownership: bool) -> bool:
        """
        Closes a document.

        Args:
            closeable (XCloseable): Object that implements XCloseable interface.
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException.

        Returns:
            bool: True if Closed; Otherwise, False
        """
        ...

    @classmethod
    def close(cls, closeable: XCloseable, deliver_ownership=False) -> bool:
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

        Returns:
            bool: True if Closed; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_CLOSING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_CLOSED` :eventref:`src-docs-event`
        """
        cargs = CancelEventArgs(Lo.close.__qualname__)
        cargs.event_data = deliver_ownership
        _Events().trigger(LoNamedEvent.DOC_CLOSING, cargs)
        if cargs.cancel:
            return False
        if closeable is None:
            return
        cls.print("Closing the document")
        try:
            closeable.close(cargs.event_data)
            cls._doc = None
            _Events().trigger(LoNamedEvent.DOC_CLOSED, EventArgs.from_args(cargs))
        except CloseVetoException as e:
            raise Exception("Close was vetoed") from e

    @overload
    @classmethod
    def close_doc(cls, doc: object) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Closeable document
        """
        ...

    @overload
    @classmethod
    def close_doc(cls, doc: object, deliver_ownership: bool) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Closeable document
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
            doc (XCloseable): Close-able document
            deliver_ownership (bool): True delegates the ownership of this closing object to
                anyone which throw the CloseVetoException.
                This new owner has to close the closing object again if his still running
                processes will be finished.
                False let the ownership at the original one which called the close() method.
                They must react for possible CloseVetoExceptions such as when document needs saving
                and try it again at a later time. This can be useful for a generic UI handling.

        Raises:
            MissingInterfaceError: if doc does not have XCloseable interface

        Attention:
            :py:meth:`~.utils.lo.Lo.close` method is called along with any of its events.
        """
        try:
            closeable = cls.qi(XCloseable, doc, True)
            cls.close(closeable=closeable, deliver_ownership=deliver_ownership)
            cls._doc = None
        except DisposedException as e:
            raise Exception("Document close failed since Office link disposed") from e

    # ================= initialization via Addon-supplied context ====================

    @classmethod
    def addon_initialize(cls, addon_xcc: XComponentContext) -> XComponent:
        """
        Initialize and ad-don

        Args:
            addon_xcc (XComponentContext): Add-on component context

        Raises:
            TypeError: If ``addon_xcc`` is None
            Exception: If unable to get service manager from ``addon_xcc``
            Exception: If unable to access desktop
            Exception: If unable to access document
            MissingInterfaceError: If unable to get ``XMultiServiceFactory`` interface instance
            CancelEventError: If ``DOC_OPENING`` is canceled

        Returns:
            XComponent: add-on as component

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.
        """
        cargs = CancelEventArgs(Lo.addon_initialize.__qualname__)
        cargs.event_data = {"addon_xcc": addon_xcc}
        eargs = EventArgs.from_args(cargs)
        _Events().trigger(LoNamedEvent.RESET, eargs)
        _Events().trigger(LoNamedEvent.DOC_OPENING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
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
        _Events().trigger(LoNamedEvent.DOC_OPENED, eargs)
        return doc

    # ============= initialization via script context ======================

    @classmethod
    def script_initialize(cls, sc: XScriptContext) -> XComponent:
        """
        Initialize script

        Args:
            sc (XScriptContext): Script context

        Raises:
            TypeError: If ``sc`` is None
            Exception: if unable to get Component Context from ``sc``
            Exception: If unable to get service manager
            Exception: If unable to access desktop
            Exception: If unable to access document
            MissingInterfaceError: if unable to get XMultiServiceFactory interface instance

        Returns:
            XComponent: script component

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
           Event args ``event_data`` is a dictionary containing all method parameters.
        """
        cargs = CancelEventArgs(Lo.script_initialize.__qualname__)
        cargs.event_data = {"sc": sc}
        eargs = EventArgs.from_args(cargs)
        _Events().trigger(LoNamedEvent.RESET, eargs)
        _Events().trigger(LoNamedEvent.DOC_OPENING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
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
        _Events().trigger(LoNamedEvent.DOC_OPENED, eargs)
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
        Dispatches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue): properties for dispatch

        Raises:
            MissingInterfaceError: If unable to obtain XDispatchHelper instance.
            Exception: If error occurs dispatching command

        Returns:
            bool: True on success.
        """
        ...

    @overload
    @staticmethod
    def dispatch_cmd(cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> bool:
        """
        Dispatches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue): properties for dispatch
            frame (XFrame): Frame to dispatch to.

        Raises:
            MissingInterfaceError: If unable to obtain XDispatchHelper instance.
            Exception: If error occurs dispatching command

        Returns:
            bool: True on success.
        """
        ...

    @classmethod
    def dispatch_cmd(cls, cmd: str, props: Iterable[PropertyValue] = None, frame: XFrame = None) -> bool:
        """
        Dispatches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue): properties for dispatch
            frame (XFrame): Frame to dispatch to.

        Raises:
            MissingInterfaceError: If unable to obtain XDispatchHelper instance.
            Exception: If error occurs dispatching command

        Returns:
            bool: True on success.

         :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DISPATCHING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DISPATCHED` :eventref:`src-docs-event`

        See Also:
            `LibreOffice Dispatch Commands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_
        """
        cargs = DispatchCancelArgs(Lo.dispatch_cmd.__qualname__, cmd)
        _Events().trigger(LoNamedEvent.DISPATCHING, cargs)
        if cargs.cancel:
            return False

        if props is None:
            props = ()
        if frame is None:
            frame = cls._xdesktop.getCurrentFrame()

        helper = cls.create_instance_mcf(XDispatchHelper, "com.sun.star.frame.DispatchHelper")
        if helper is None:
            raise mEx.MissingInterfaceError(XDispatchHelper, f"Could not create dispatch helper for command {cmd}")
        try:
            helper.executeDispatch(frame, f".uno:{cmd}", "", 0, props)
            _Events().trigger(LoNamedEvent.DISPATCHED, DispatchArgs.from_args(cargs))
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
            Lo.print("No office connection found")
            return
        try:
            ts = mInfo.Info.get_interface_types(obj)
            title = "Object"
            if ts is not None and len(ts) > 0:
                title = ts[0].getTypeName() + " " + title
            inspector = cls._mc_factory.createInstanceWithContext("org.openoffice.InstanceInspector", cls._xcc)
            #       hands on second use
            if inspector is None:
                Lo.print("Inspector Service could not be instantiated")
                return
            Lo.print("Inspector Service instantiated")
            intro = cls.create_instance_mcf(XIntrospection, "com.sun.star.beans.Introspection")
            intro_acc = intro.inspect(inspector)
            method = intro_acc.getMethod("inspect", -1)
            Lo.print(f"inspect() method was found: {method is not None}")
            params = [[obj, title]]
            method.invoke(inspector, params)
        except Exception as e:
            Lo.print("Could not access Inspector:")
            Lo.print(f"    {e}")

    @classmethod
    def mri_inspect(cls, obj: object) -> None:
        """
        call MRI's inspect() to inspect obj.

        Args:
            obj (object): obj to inspect

        Raises:
            Exception: If MRI service could not be instantiated.

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
        Lo.print("MRI Inspector Service instantiated")
        xi.inspect(obj)

    # ------------------ color methods ---------------------
    # section intentionally left out.

    # ================== other utils =============================

    @staticmethod
    def delay(ms: int) -> None:
        """
        Delay execution for a given number of milliseconds.

        Args:
            ms (int): Number of milliseconds to delay
        """
        if ms <= 0:
            Lo.print("Ms must be greater then zero")
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
        Console displays Press Enter to continue...
        """
        input("Press Enter to continue...")

    @staticmethod
    def is_url(fnm: PathOrStr) -> bool:
        """
        Gets if a string is a URL format.

        Args:
            fnm (PathOrStr): string to check.

        Returns:
            bool: True if URL format; Otherwise, False
        """
        # https://stackoverflow.com/questions/7160737/how-to-validate-a-url-in-python-malformed-or-not
        try:
            pth = mFileIO.FileIO.get_absolute_path(fnm)
            result = urlparse(str(pth))
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
        Converts string into int.

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
            Lo.print(f"{s} could not be parsed as an int; using 0")
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
    def print_names(names: Iterable[str], num_per_line: int = 4) -> None:
        """
        Prints names to console

        Args:
            names (Iterable[str]): names to print
            num_per_line (int): Number of names per line. Default 4
        """
        if names is None:
            print("  No names found")
            return
        sorted_list = sorted(names, key=str.casefold)
        print(f"No. of names: {len(sorted_list)}")
        nl_count = 0
        for name in sorted_list:
            print(f"  '{name}'", end="")
            if num_per_line <= 0:
                print()
                continue
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
            List[str] | None: Container name is found; Otherwise, None
        """
        if con is None:
            Lo.print("Container is null")
            return None
        num_el = con.getCount()
        if num_el == 0:
            Lo.print("No elements in the container")
            return None

        names_list = []
        for i in range(num_el):
            named = con.getByIndex(i)
            names_list.append(named.getName())

        if len(names_list) == 0:
            Lo.print("No element names found in the container")
            return None
        return names_list

    @classmethod
    def find_container_props(cls, con: XIndexAccess, nm: str) -> XPropertySet | None:
        """
        Find as Property Set in a container

        Args:
            con (XIndexAccess): Container to search
            nm (str): Name of property to search for

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
                cls.print(f"Could not access element {i}")
        cls.print(f"Could not find a '{nm}' property set in the container")
        return None

    @classmethod
    def is_uno_interfaces(cls, component: object, *args: str | UnoInterface) -> bool:
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
                    t = uno.getClass(arg)
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

    @classmethod
    def get_frame(cls) -> XFrame:
        """
        Gets XFrame for current LibreOffice instance

        Returns:
            XFrame: frame
        """
        if cls.star_desktop is None:
            raise Exception("No desktop found")
        return cls.XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
        # return cast(XDesktop, cls.star_desktop).getCurrentFrame()

    @classmethod
    def get_model(cls) -> XModel:
        """
        Gets XModel

        Returns:
            XModel: Gets model for current LibreOffice instance
        """
        return cls.XSCRIPTCONTEXT.getDocument()
        # return cls.qi(XModel, cls._doc)

    @classmethod
    def lock_controllers(cls) -> bool:
        """
        Suspends some notifications to the controllers which are used for display updates.

        The calls to :py:meth:`~.lo.Lo.lock_controllers` and :py:meth:`~.lo.Lo.unlock_controllers`
        may be nested and even overlapping, but they must be in pairs.
        While there is at least one lock remaining, some notifications for
        display updates are not broadcast.

        Raises:
            MissingInterfaceError: If unable to obtain XModel interface.

        Returns:
            bool: False if ``CONTROLERS_LOCKING`` event is canceled; Otherwise, True

         :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.CONTROLERS_LOCKING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.CONTROLERS_LOCKED` :eventref:`src-docs-event`

        See Also:
            :py:class:`.Lo.ControllerLock`

        """
        # much faster updates as screen is basically suspended
        cargs = CancelEventArgs(Lo.lock_controllers.__qualname__)
        _Events().trigger(LoNamedEvent.CONTROLERS_LOCKING, cargs)
        if cargs.cancel:
            return False
        xmodel = cls.qi(XModel, cls._doc, True)
        xmodel.lockControllers()
        _Events().trigger(LoNamedEvent.CONTROLERS_LOCKED, EventArgs(cls))
        return True

    @classmethod
    def unlock_controllers(cls) -> uno.Bool:
        """
        Resumes the notifications which were suspended by :py:meth:`~.lo.Lo.lock_controllers`.

        The calls to :py:meth:`~.lo.Lo.lock_controllers` and :py:meth:`~.lo.Lo.unlock_controllers`
        may be nested and even overlapping, but they must be in pairs.
        While there is at least one lock remaining, some notifications for
        display updates are not broadcast.

        Raises:
            MissingInterfaceError: If unable to obtain XModel interface.

        Returns:
            bool: False if CONTROLERS_UNLOCKING event is canceled; Otherwise, True

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.CONTROLERS_UNLOCKING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.CONTROLERS_UNLOCKED` :eventref:`src-docs-event`

        See Also:
            :py:class:`.Lo.ControllerLock`
        """
        cargs = CancelEventArgs(Lo.unlock_controllers.__qualname__)
        _Events().trigger(LoNamedEvent.CONTROLERS_UNLOCKING, cargs)
        if cargs.cancel:
            return False
        xmodel = cls.qi(XModel, cls._doc, True)
        if xmodel.hasControllersLocked():
            xmodel.unlockControllers()
        _Events().trigger(LoNamedEvent.CONTROLERS_UNLOCKED, EventArgs.from_args(cargs))
        return True

    @classmethod
    def has_controllers_locked(cls) -> bool:
        """
        Determines if there is at least one lock remaining.

        While there is at least one lock remaining, some notifications for display
        updates are not broadcast to the controllers.

        Returns:
            bool: True if any lock exist; Otherwise, False

        See Also:
            :py:class:`.Lo.ControllerLock`
        """
        xmodel = cls.qi(XModel, cls._doc)
        return xmodel.hasControllersLocked()

    @staticmethod
    def print(*args, **kwargs) -> None:
        """
        Utility function that passes to actual print.

        If :py:attr:`GblNamedEvent.PRINTING <.events.gbl_named_event.GblNamedEvent.PRINTING>`
        event is canceled the this method will not print.

        :events:
           .. include:: ../../resources/global/printing_events.rst

        Note:
            .. include:: ../../resources/global/printing_note.rst
        """
        cargs = CancelEventArgs(Lo.print.__qualname__)
        _Events().trigger(GblNamedEvent.PRINTING, cargs)
        if cargs.cancel:
            return
        print(*args, **kwargs)

    @classproperty
    def null_date(cls) -> datetime:
        """
        Gets Value of Null Date in UTC

        Returns:
            datetime: Null Date on success; Otherwise, None

        Note:
            If Lo has no document to determine date from then a
            default date of 1889/12/30 is returned.
        """
        # https://tinyurl.com/2pdrt5z9#NullDate
        try:
            return cls.__null_date
        except AttributeError:
            cls.__null_date = datetime(year=1889, month=12, day=30, tzinfo=timezone.utc)
            if cls._doc is None:
                return cls.__null_date
            n_supplier = cls.qi(XNumberFormatsSupplier, cls._doc)
            if n_supplier is None:
                # this is not always a XNumberFormatsSupplier such as *.odp documents
                return cls.__null_date
            number_settings = n_supplier.getNumberFormatSettings()
            d = number_settings.getPropertyValue("NullDate")
            cls.__null_date = datetime(d.Year, d.Month, d.Day, tzinfo=timezone.utc)
        return cls.__null_date

    @null_date.setter
    def null_date(cls, value) -> None:
        # raise error on set. Not really necessary but gives feedback.
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.__name__)

    @classproperty
    def is_loaded(cls) -> bool:
        """
        Gets office is currently loaded

        Returns:
            bool: True if load_office has been called; Othwriwse, False
        """

        return not cls._lo_inst is None

    @is_loaded.setter
    def is_loaded(cls, value) -> None:
        # raise error on set. Not really necessary but gives feedback.
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.__name__)

    @classproperty
    def is_macro_mode(cls) -> bool:
        """
        Gets if currently running scripts inside of LO (macro) or standalone

        Returns:
            bool: True if running as a macro; Otherwise, False
        """

        try:
            return cls._is_macro_mode
        except AttributeError:
            if cls._lo_inst is None:
                return False
            cls._is_macro_mode = isinstance(cls._lo_inst, LoDirectStart)
        return cls._is_macro_mode

    @is_macro_mode.setter
    def is_macro_mode(cls, value) -> None:
        # raise error on set. Not really necessary but gives feedback.
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.__name__)

    @classproperty
    def star_desktop(cls) -> XDesktop:
        """Get current desktop"""
        return cls._xdesktop

    StarDesktop, stardesktop = star_desktop, star_desktop

    @classproperty
    def this_component(cls) -> XComponent:
        """
        When the current component is the Basic IDE, the ThisComponent object returns
        in Basic the component owning the currently run user script.
        Above behavior cannot be reproduced in Python.

        When running in a macro this property can be access directly to get the current document.

        When not in a macro then load_office() must be called first

        Returns:
            the current component or None when not a document
        """
        try:
            return cls._this_component
        except AttributeError:
            if mock_g.DOCS_BUILDING:
                cls._this_component = None
                return cls._this_component
            if cls.is_loaded is False:
                # attempt to connect direct
                # failure will result in script error and then exit
                cls.load_office()

            # comp = cls.star_desktop.getCurrentComponent()
            desktop = cls.get_desktop()
            if desktop is None:
                return None
            if cls._doc is None:
                cls._doc = desktop.getCurrentComponent()
            if cls._doc is None:
                return None
            impl = cls._doc.ImplementationName
            if impl in ("com.sun.star.comp.basic.BasicIDE", "com.sun.star.comp.sfx2.BackingComp"):
                return None  # None when Basic IDE or welcome screen
            cls._this_component = cls._doc
            return cls._this_component

    ThisComponent, thiscomponent = this_component, this_component

    @classproperty
    def xscript_context(cls) -> XScriptContext:
        """
        a substitute to `XSCRIPTCONTEXT` (Libre|Open)Office built-in

        Returns:
            XScriptContext: XScriptContext instance
        """
        try:
            return cls._xscript_context
        except AttributeError:
            ctx = cls.get_context()
            if ctx is None:
                # attempt to connect direct
                # failure will result in script error and then exit
                cls.load_office()
                ctx = cls.get_context()

            desktop = cls.get_desktop()
            model = cls.qi(XModel, cls._doc)
            cls._xscript_context = script_context.ScriptContext(ctx=ctx, desktop=desktop, doc=model)
        return cls._xscript_context

    XSCRIPTCONTEXT = xscript_context

    @classproperty
    def bridge(cls) -> XComponent:
        """
        Gets connection bridge component

        Returns:
            XComponent: bridge component
        """
        try:
            return cls._bridge_component
        except AttributeError:
            try:
                # when running as macro cls._lo_inst will not have bridge_component
                cls._bridge_component = cls._lo_inst.bridge_component
            except AttributeError:
                cls._bridge_component = None
            return cls._bridge_component


class _LoManager(metaclass=StaticProperty):
    """Manages clearing and resetting for Lo static class"""

    @staticmethod
    def del_cache_attrs(source: object, event: EventArgs) -> None:
        # clears Lo Attributes that are dynamically created
        dattrs = ("_xscript_context", "_is_macro_mode", "_this_component", "_bridge_component", "__null_date")
        for attr in dattrs:
            if hasattr(Lo, attr):
                delattr(Lo, attr)

    @staticmethod
    def disposing_bridge(src: XEventAdapter, event: EventObject) -> None:
        # do not try and exit script here.
        # for some reason when office triggers this method calls such as:
        # raise SystemExit(1)
        # does not exit the script
        _Events().trigger(LoNamedEvent.BRIDGE_DISPOSED, EventArgs(_LoManager.disposing_bridge.__qualname__))

    @staticmethod
    def on_disposed(source: Any, event: EventObject) -> None:
        Lo.print("Office bridge has gone!!")
        dattrs = ("_xcc", "_doc", "_mc_factory", "_ms_factory", "_lo_inst", "_xdesktop")
        dvals = (None, None, None, None, None, None)
        for attr, val in zip(dattrs, dvals):
            setattr(Lo, attr, val)

    @staticmethod
    def on_loading(source: Any, event: CancelEventArgs) -> None:
        try:
            bridge = cast(XComponent, Lo._lo_inst.bridge_component)
            bridge.removeEventListener(_LoManager.event_adapter)
        except Exception:
            pass

    @staticmethod
    def on_loaded(source: Any, event: EventObject) -> None:
        if Lo.bridge is not None:
            Lo.bridge.addEventListener(_LoManager.event_adapter)

    @classproperty
    def event_adapter(cls) -> XEventAdapter:
        try:
            return cls._event_adapter
        except AttributeError:
            bridge_listen = XEventAdapter()
            bridge_listen.disposing = types.MethodType(cls.disposing_bridge, bridge_listen)
            cls._event_adapter = bridge_listen
        return cls._event_adapter


_Events().on(LoNamedEvent.RESET, _LoManager.del_cache_attrs)
_Events().on(LoNamedEvent.OFFICE_LOADING, _LoManager.on_loading)
_Events().on(LoNamedEvent.OFFICE_LOADED, _LoManager.on_loaded)
_Events().on(LoNamedEvent.BRIDGE_DISPOSED, _LoManager.on_disposed)


__all__ = ("Lo",)
