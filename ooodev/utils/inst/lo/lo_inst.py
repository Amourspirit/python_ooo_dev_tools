# coding: utf-8
# Python conversion of Lo.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

from __future__ import annotations
from datetime import datetime, timezone
import time
from typing import TYPE_CHECKING, Any, Iterable, Optional, List, Sequence, Tuple, cast, overload, Type
from urllib.parse import urlparse
import uno

from ooodev.mock import mock_g

# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from ooodev.utils import file_io as mFileIO
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps
from ooodev.utils import script_context
from ooodev.utils import table_helper as mThelper
from ooodev.utils import xml_util as mXML
from ooodev.conn import cache as mCache
from ooodev.conn import connectors
from ooodev.conn.connect import ConnectBase, LoPipeStart, LoSocketStart, LoDirectStart, LoBridgeCommon
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.dispatch_args import DispatchArgs
from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.event_singleton import _Events
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.formatters.formatter_table import FormatterTable
from ooodev.utils.type_var import PathOrStr, UnoInterface, T, Table
from ooodev.utils.inst.lo.options import Options as LoOptions
from ooodev.utils.inst.lo.doc_type import DocType as LoDocType, DocTypeStr as LoDocTypeStr
from ooodev.adapter.lang.event_listener import EventListener

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
    try:
        from typing import Literal  # Py >= 3.8
    except ImportError:
        from typing_extensions import Literal
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XChild
    from com.sun.star.container import XIndexAccess
    from com.sun.star.frame import XFrame
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.lang import XComponent
    from com.sun.star.script.provider import XScriptContext
    from com.sun.star.uno import XComponentContext
    from com.sun.star.uno import XInterface
    from com.sun.star.lang import EventObject

from ooo.dyn.document.macro_exec_mode import MacroExecMode  # const
from ooo.dyn.lang.disposed_exception import DisposedException
from ooo.dyn.util.close_veto_exception import CloseVetoException


# PathOrStr = type_var.PathOrStr
# """Path like object or string"""


class LoInst:
    """Lo instance class. This is the main class for interacting with LO."""

    def __init__(self, opt: LoOptions | None = None) -> None:
        # self._events = Events(self)
        self._xcc: XComponentContext = None
        self._doc: XComponent = None
        """remote component context"""
        self._xdesktop: XDesktop = None
        """remote desktop UNO service"""
        self._mc_factory: XMultiComponentFactory = None
        self._ms_factory: XMultiServiceFactory = None
        self._is_office_terminated: bool = False
        self._lo_inst: ConnectBase = None
        self._loader = None
        self._disposed = True

        if opt is None:
            self._opt = LoOptions()
        else:
            self._opt = opt

        self._events = _Events()

        self._allow_print = opt.verbose
        self._set_lo_events()

    # region Events
    def _set_lo_events(self) -> None:
        def _on_lo_del_cache_attrs(source: object, event: EventArgs) -> None:
            self.on_lo_del_cache_attrs(source, event)

        def _on_lo_loading(source: Any, event: CancelEventArgs) -> None:
            self.on_lo_loading(source, event)

        def _on_lo_loaded(source: Any, event: EventObject) -> None:
            self.on_lo_loaded(source, event)

        def _on_lo_disposed(source: Any, event: EventObject) -> None:
            self.on_lo_disposed(source, event)

        def _on_lo_disposing_bridge(src: EventListener, event: EventObject) -> None:
            self.on_lo_disposing_bridge(src, event)

        self._fn_on_lo_del_cache_attrs = _on_lo_del_cache_attrs
        self._fn_on_lo_loading = _on_lo_loading
        self._fn_on_lo_loaded = _on_lo_loaded
        self._fn_on_lo_disposed = _on_lo_disposed
        self._fn_on_lo_disposing_bridge = _on_lo_disposing_bridge

        self._event_listener = EventListener()
        self._event_listener.on("disposing", _on_lo_disposing_bridge)

        _Events().on(LoNamedEvent.RESET, _on_lo_del_cache_attrs)
        _Events().on(LoNamedEvent.OFFICE_LOADING, _on_lo_loading)
        _Events().on(LoNamedEvent.OFFICE_LOADED, _on_lo_loaded)
        _Events().on(LoNamedEvent.BRIDGE_DISPOSED, _on_lo_disposed)

    def on_lo_del_cache_attrs(self, source: object, event: EventArgs) -> None:
        # clears Lo Attributes that are dynamically created
        dattrs = ("_xscript_context", "_is_macro_mode", "_this_component", "_bridge_component", "_null_date")
        for attr in dattrs:
            if hasattr(self, attr):
                delattr(self, attr)

    def on_lo_disposing_bridge(self, src: EventListener, event: EventObject) -> None:
        _Events().trigger(LoNamedEvent.BRIDGE_DISPOSED, EventArgs(self.on_lo_disposing_bridge.__qualname__))

    def on_lo_disposed(self, source: Any, event: EventObject) -> None:
        self.print("Office bridge has gone!!")
        dattrs = ("_xcc", "_doc", "_mc_factory", "_ms_factory", "_lo_inst", "_xdesktop", "_loader")
        dvals = (None, None, None, None, None, None, None)
        for attr, val in zip(dattrs, dvals):
            setattr(self, attr, val)
        setattr(self, "_disposed", True)

    def on_lo_loaded(self, source: Any, event: EventObject) -> None:
        if self.bridge is not None:
            self.bridge.addEventListener(self._event_listener)

    def on_lo_loading(self, source: Any, event: CancelEventArgs) -> None:
        try:
            if hasattr(self, "_bridge_component") and self.bridge is not None:
                self.bridge.removeEventListener(self._event_listener)
        except Exception:
            pass

    def on_lo_del_cache_attrs(self, source: object, event: EventArgs) -> None:
        # clears Attributes that are dynamically created
        dattrs = ("_xscript_context", "_is_macro_mode", "_this_component", "_bridge_component", "__null_date")
        for attr in dattrs:
            if hasattr(self, attr):
                delattr(self, attr)

    # endregion Events

    def _init_from_ctx(self, script_ctx: XScriptContext, connector: LoBridgeCommon) -> None:
        self._disposed = False
        self._lo_inst = connector
        self._xcc = self._lo_inst.ctx
        self._mc_factory = self._xcc.getServiceManager()
        if self._mc_factory is None:
            raise RuntimeError("Could not get service manager from component context")
        self._xdesktop = script_ctx.getDesktop()
        if self._xdesktop is None:
            raise RuntimeError("Could not get desktop from script context")
        self._loader = self.qi(XComponentLoader, self._xdesktop)
        if self._loader is None:
            raise RuntimeError("Could not get loader from desktop")

    # region    qi()

    @overload
    def qi(self, atype: Type[T], obj: Any) -> T | None:
        ...

    @overload
    def qi(self, atype: Type[T], obj: Any, raise_err: Literal[True]) -> T:
        ...

    @overload
    def qi(self, atype: Type[T], obj: Any, raise_err: Literal[False]) -> T | None:
        ...

    def qi(self, atype: Type[T], obj: XTypeProvider, raise_err: bool = False) -> T | None:
        result = None
        if uno.isInterface(atype) and hasattr(obj, "queryInterface"):
            uno_t = uno.getTypeByName(atype.__pyunointerface__)
            result = obj.queryInterface(uno_t)
        if raise_err is True and result is None:
            raise mEx.MissingInterfaceError(atype)
        return result

    # endregion qi()

    def get_context(self) -> XComponentContext:
        """
        Gets current LO Component Context
        """
        return self._xcc

    def get_desktop(self) -> XDesktop:
        """
        Gets current LO Desktop
        """
        return self._xdesktop

    def get_component_factory(self) -> XMultiComponentFactory:
        """Gets current multi component factory"""
        return self._mc_factory

    def get_service_factory(self) -> XMultiServiceFactory:
        """Gets current multi service factory"""
        # return cls._bridge_component
        return self._ms_factory

    # region interface object creation

    # region    create_instance_msf()
    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str) -> T | None:
        ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, msf: Any | None) -> T | None:
        ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, *, raise_err: Literal[True]) -> T:
        ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, *, raise_err: Literal[False]) -> T | None:
        ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[True]) -> T:
        ...

    @overload
    def create_instance_msf(
        self, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[False]
    ) -> T | None:
        ...

    def create_instance_msf(
        self, atype: Type[T], service_name: str, msf: XMultiServiceFactory = None, raise_err: bool = False
    ) -> T:
        if self._ms_factory is None:
            raise Exception("No document found")
        try:
            if msf is None:
                obj = self._ms_factory.createInstance(service_name)
            else:
                obj = msf.createInstance(service_name)
            if raise_err is True and obj is None:
                mEx.CreateInstanceMsfError(atype, service_name)
            interface_obj = self.qi(atype=atype, obj=obj)
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
    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str) -> T | None:
        ...

    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None) -> T | None:
        ...

    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str, *, raise_err: Literal[True]) -> T:
        ...

    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str, *, raise_err: Literal[False]) -> T | None:
        ...

    @overload
    def create_instance_mcf(
        self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None, raise_err: Literal[True]
    ) -> T:
        ...

    @overload
    def create_instance_mcf(
        self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None, raise_err: Literal[False]
    ) -> T | None:
        ...

    def create_instance_mcf(
        self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None = None, raise_err: bool = False
    ) -> T | None:
        #  create an interface object of class atype from the named service;
        #  uses XComponentContext and XMultiComponentFactory
        #  so only a bridge to office is needed
        if self._xcc is None or self._mc_factory is None:
            raise Exception("No office connection found")
        try:
            if args is not None:
                obj = self._mc_factory.createInstanceWithArgumentsAndContext(service_name, args, self._xcc)
            else:
                obj = self._mc_factory.createInstanceWithContext(service_name, self._xcc)
            if raise_err is True and obj is None:
                mEx.CreateInstanceMcfError(atype, service_name)
            interface_obj = self.qi(atype=atype, obj=obj)
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

    def get_parent(self, a_component: XChild) -> XInterface:
        return a_component.getParent()

    # region Start Office

    def load_office(
        self,
        connector: connectors.ConnectPipe | connectors.ConnectSocket | None = None,
        cache_obj: mCache.Cache | None = None,
        opt: LoOptions | None = None,
    ) -> XComponentLoader:
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

        cargs = CancelEventArgs(self.load_office.__qualname__)

        cargs.event_data = {
            "connector": connector,
        }

        eargs = EventArgs.from_args(cargs)
        self._events.trigger(LoNamedEvent.RESET, eargs)

        self._events.trigger(LoNamedEvent.OFFICE_LOADING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        b_connector = cargs.event_data["connector"]

        if opt:
            self._opt = opt

        self.print("Loading Office...")
        if b_connector is None:
            try:
                self._lo_inst = LoDirectStart()
                self._lo_inst.connect()
            except Exception as e:
                self.print(
                    "Office context could not be created. A connector must be supplied if not running as a macro"
                )
                self.print(f"    {e}")
                raise SystemExit(1)
        elif isinstance(b_connector, connectors.ConnectPipe):
            try:
                self._lo_inst = LoPipeStart(connector=b_connector, cache_obj=cache_obj)
                self._lo_inst.connect()
            except Exception as e:
                self.print("Office context could not be created")
                self.print(f"    {e}")
                raise SystemExit(1)
        elif isinstance(b_connector, connectors.ConnectSocket):
            try:
                self._lo_inst = LoSocketStart(connector=b_connector, cache_obj=cache_obj)
                self._lo_inst.connect()
            except Exception as e:
                self.print("Office context could not be created")
                self.print(f"    {e}")
                raise SystemExit(1)
        else:
            self.print("Invalid Connector type. Fatal Error.")
            raise SystemExit(1)
        self._disposed = False
        self._xcc = self._lo_inst.ctx
        self._mc_factory = self._xcc.getServiceManager()
        if self._mc_factory is None:
            self.print("Office Service Manager is unavailable")
            raise SystemExit(1)
        self._xdesktop = self.create_instance_mcf(XDesktop, "com.sun.star.frame.Desktop")
        if self._xdesktop is None:
            # OPTIMIZE: Perhaps system exit is not the best what to handle no desktop service
            self.print("Could not create a desktop service")
            raise SystemExit(1)
        self._loader = self.qi(XComponentLoader, self._xdesktop)
        if self._loader is None:
            self.print("Unable to access XComponentLoader")
            SystemExit(1)

        self._events.trigger(LoNamedEvent.OFFICE_LOADED, eargs)
        return self._loader
        # return cls.xdesktop

    # endregion Start Office

    # region office shutdown
    def close_office(self) -> bool:
        self.print("Closing Office")

        cargs = CancelEventArgs(self.close_office.__qualname__)
        self._events.trigger(LoNamedEvent.OFFICE_CLOSING, cargs)
        if cargs.cancel:
            return False

        self._doc = None
        if self._xdesktop is None:
            self.print("No office connection found")
            return True

        if self._is_office_terminated:
            self.print("Office has already been requested to terminate")
            return self._is_office_terminated
        num_tries = 1
        start = time.time()
        elapsed = 0
        seconds = 10
        while self._is_office_terminated is False and elapsed < seconds:
            elapsed = time.time() - start
            self._is_office_terminated = self._try_to_terminate(num_tries)
            time.sleep(0.5)
            num_tries += 1
        if self._is_office_terminated:
            eargs = EventArgs.from_args(cargs)
            self._events.trigger(LoNamedEvent.OFFICE_CLOSED, eargs)
            self._events.trigger(LoNamedEvent.RESET, eargs)
        return self._is_office_terminated

    def _try_to_terminate(self, num_tries: int) -> bool:
        if self._disposed:
            return True
        try:
            is_dead = self._xdesktop.terminate()
            if is_dead:
                if num_tries > 1:
                    self.print(f"{num_tries}. Office terminated")
                else:
                    self.print("Office terminated")
            else:
                self.print(f"{num_tries}. Office failed to terminate")
            return is_dead
        except DisposedException as e:
            self.print("Office link disposed")
            return True
        except Exception as e:
            self.print(f"Termination exception: {e}")
            return False

    def kill_office(self) -> None:
        if self._lo_inst is None:
            self.print("No instance to kill")
            return
        try:
            # raised a NotImplementedError when self._lo_inst is direct (macro mode)
            self._lo_inst.kill_soffice()
            self._is_office_terminated = True
            eargs = EventArgs(self.kill_office.__qualname__)
            self._events.trigger(LoNamedEvent.OFFICE_CLOSED, eargs)
            self._events.trigger(LoNamedEvent.RESET, eargs)
            self.print("Killed Office")
        except Exception as e:
            raise Exception(f"Unbale to kill Office") from e

    # endregion office shutdown

    # region document opening

    # region open_flat_doc()
    @overload
    def open_flat_doc(self, fnm: PathOrStr, doc_type: LoDocType) -> XComponent:
        ...

    @overload
    def open_flat_doc(self, fnm: PathOrStr, doc_type: LoDocType, loader: XComponentLoader) -> XComponent:
        ...

    def open_flat_doc(
        self, fnm: PathOrStr, doc_type: LoDocType, loader: Optional[XComponentLoader] = None
    ) -> XComponent:
        if loader is None:
            loader = self._loader
        nn = mXML.XML.get_flat_fiter_name(doc_type=doc_type)
        self.print(f"Flat filter Name: {nn}")
        # do not set Hidden=True property here.
        # there is a strange error that pops up conditionally and it seems
        # to be remedied by not seting Hidden=True
        # see comments in tests.text_xml.test_in_filters.test_transform_clubs()
        return self.open_doc(fnm, loader, mProps.Props.make_props(FilterName=nn))

    # endregion open_flat_doc()

    # region open_doc()
    @overload
    def open_doc(self, fnm: PathOrStr) -> XComponent:
        ...

    @overload
    def open_doc(self, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        ...

    @overload
    def open_doc(self, fnm: PathOrStr, *, props: Iterable[PropertyValue]) -> XComponent:
        ...

    @overload
    def open_doc(self, fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        ...

    def open_doc(
        self,
        fnm: PathOrStr,
        loader: Optional[XComponentLoader] = None,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        # Props and FileIO are called this method so triger global_reset first.
        if loader is None:
            loader = self._loader
        cargs = CancelEventArgs(self.open_doc.__qualname__)
        cargs.event_data = {
            "fnm": fnm,
            "loader": loader,
            "props": props,
        }
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(LoNamedEvent.RESET, eargs)
        self._events.trigger(LoNamedEvent.DOC_OPENING, cargs)
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
            if self.is_url(pth):
                self.print(f"Will treat filename as a URL: '{pth}'")
                open_file_url = pth
            else:
                raise Exception(f"Unable to get url from file: {pth}")
        else:
            self.print(f"Opening {pth}")
            open_file_url = mFileIO.FileIO.fnm_to_url(pth)

        try:
            doc = loader.loadComponentFromURL(open_file_url, "_blank", 0, props)
            self._ms_factory = self.qi(XMultiServiceFactory, doc)
            self._doc = doc
            if self._doc is None:
                raise mEx.NoneError("loadComponentFromURL() returned None")
            self._events.trigger(LoNamedEvent.DOC_OPENED, eargs)
            return doc
        except Exception as e:
            raise Exception("Unable to open the document") from e

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    def open_readonly_doc(self, fnm: PathOrStr) -> XComponent:
        ...

    @overload
    def open_readonly_doc(self, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        ...

    def open_readonly_doc(self, fnm: PathOrStr, loader: Optional[XComponentLoader] = None) -> XComponent:
        if loader is None:
            loader = self._loader
        return self.open_doc(fnm, loader, mProps.Props.make_props(Hidden=True, ReadOnly=True))

    # endregion open_readonly_doc()

    # ======================== document creation ==============

    def ext_to_doc_type(self, ext: str) -> LoDocTypeStr:
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
            self.print("Empty string: Using writer")
            return LoDocTypeStr.WRITER
        if e == "odt":
            return LoDocTypeStr.WRITER
        elif e == "odp":
            return LoDocTypeStr.IMPRESS
        elif e == "odg":
            return LoDocTypeStr.DRAW
        elif e == "ods":
            return LoDocTypeStr.CALC
        elif e == "odb":
            return LoDocTypeStr.BASE
        elif e == "odf":
            return LoDocTypeStr.MATH
        else:
            self.print(f"Do not recognize extension '{ext}'; using writer")
            return LoDocTypeStr.WRITER

    def doc_type_str(self, doc_type_val: LoDocType) -> LoDocTypeStr:
        if doc_type_val == LoDocType.WRITER:
            return LoDocTypeStr.WRITER
        elif doc_type_val == LoDocType.IMPRESS:
            return LoDocTypeStr.IMPRESS
        elif doc_type_val == LoDocType.DRAW:
            return LoDocTypeStr.DRAW
        elif doc_type_val == LoDocType.CALC:
            return LoDocTypeStr.CALC
        elif doc_type_val == LoDocType.BASE:
            return LoDocTypeStr.BASE
        elif doc_type_val == LoDocType.MATH:
            return LoDocTypeStr.MATH
        else:
            self.print(f"Do not recognize extension '{doc_type_val}'; using writer")
            return LoDocTypeStr.WRITER

    # region create_doc()
    @overload
    def create_doc(self, doc_type: LoDocTypeStr) -> XComponent:
        ...

    @overload
    def create_doc(self, doc_type: LoDocTypeStr, loader: XComponentLoader) -> XComponent:
        ...

    @overload
    def create_doc(self, doc_type: LoDocTypeStr, *, props: Iterable[PropertyValue]) -> XComponent:
        ...

    @overload
    def create_doc(
        self, doc_type: LoDocTypeStr, loader: XComponentLoader, props: Iterable[PropertyValue]
    ) -> XComponent:
        ...

    def create_doc(
        self,
        doc_type: LoDocTypeStr,
        loader: Optional[XComponentLoader] = None,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        # Props is called in this method so trigger global_reset first
        if loader is None:
            loader = self._loader
        cargs = CancelEventArgs(self.create_doc.__qualname__)
        cargs.event_data = {
            "doc_type": doc_type,
            "loader": loader,
            "props": props,
        }
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(LoNamedEvent.RESET, eargs)
        self._events.trigger(LoNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        dtype = LoDocTypeStr(cargs.event_data["doc_type"])
        if props is None:
            props = mProps.Props.make_props(Hidden=True)
        self.print(f"Creating Office document {dtype}")
        try:
            doc = loader.loadComponentFromURL(f"private:factory/{dtype}", "_blank", 0, props)
            self._ms_factory = self.qi(XMultiServiceFactory, doc)
            if self._ms_factory is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            self._doc = doc
            self._events.trigger(LoNamedEvent.DOC_CREATED, eargs)
            return self._doc
        except Exception as e:
            raise Exception("Could not create a document") from e

    # endregion create_doc()

    # region create_macro_doc()
    @overload
    def create_macro_doc(self, doc_type: LoDocTypeStr) -> XComponent:
        ...

    @overload
    def create_macro_doc(self, doc_type: LoDocTypeStr, loader: XComponentLoader) -> XComponent:
        ...

    def create_macro_doc(self, doc_type: LoDocTypeStr, loader: Optional[XComponentLoader] = None) -> XComponent:
        if loader is None:
            loader = self._loader
        return self.create_doc(
            doc_type=doc_type,
            loader=loader,
            props=mProps.Props.make_props(Hidden=False, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE),
        )

    # endregion create_macro_doc()

    # region create_doc_from_template()

    @overload
    def create_doc_from_template(self, template_path: PathOrStr) -> XComponent:
        ...

    @overload
    def create_doc_from_template(self, template_path: PathOrStr, loader: XComponentLoader) -> XComponent:
        ...

    def create_doc_from_template(
        self, template_path: PathOrStr, loader: Optional[XComponentLoader] = None
    ) -> XComponent:
        if loader is None:
            loader = self._loader
        cargs = CancelEventArgs(self.create_doc_from_template.__qualname__)
        self._events.trigger(LoNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        if not mFileIO.FileIO.is_openable(template_path):
            raise Exception(f"Template file can not be opened: '{template_path}'")
        self.print(f"Opening template: '{template_path}'")
        template_url = mFileIO.FileIO.fnm_to_url(fnm=template_path)

        props = mProps.Props.make_props(Hidden=True, AsTemplate=True)
        try:
            self._doc = loader.loadComponentFromURL(template_url, "_blank", 0, props)
            self._ms_factory = self.qi(XMultiServiceFactory, self._doc)
            if self._ms_factory is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            self._events.trigger(LoNamedEvent.DOC_CREATED, EventArgs.from_args(cargs))
            return self._doc
        except Exception as e:
            raise Exception(f"Could not create document from template") from e

    # endregion create_doc_from_template()

    # ======================== document saving ==============

    def save(self, doc: object) -> bool:
        cargs = CancelEventArgs(self.save.__qualname__)
        cargs.event_data = {"doc": doc}
        self._events.trigger(LoNamedEvent.DOC_SAVING, cargs)
        if cargs.cancel:
            return False

        store = self.qi(XStorable, doc, True)
        try:
            store.store()
            self.print("Saved the document by overwriting")
        except IOException as e:
            raise Exception(f"Could not save the document") from e

        self._events.trigger(LoNamedEvent.DOC_SAVED, EventArgs.from_args(cargs))
        return True

    # region    save_doc()

    @overload
    def save_doc(self, doc: object, fnm: PathOrStr) -> bool:
        ...

    @overload
    def save_doc(self, doc: object, fnm: PathOrStr, password: str) -> bool:
        ...

    @overload
    def save_doc(self, doc: object, fnm: PathOrStr, password: str, format: str) -> bool:
        ...

    def save_doc(self, doc: object, fnm: PathOrStr, password: str = None, format: str = None) -> bool:
        cargs = CancelEventArgs(self.save_doc.__qualname__)
        cargs.event_data = {
            "doc": doc,
            "fnm": fnm,
            "password": password,
            "format": format,
        }

        fnm = cargs.event_data["fnm"]
        password = cargs.event_data["password"]
        format = cargs.event_data["format"]

        self._events.trigger(LoNamedEvent.DOC_SAVING, cargs)
        if cargs.cancel:
            return False
        store = self.qi(XStorable, doc, True)
        doc_type = mInfo.Info.report_doc_type(doc)
        kargs = {"fnm": fnm, "store": store, "doc_type": doc_type}
        if password is not None:
            kargs["password"] = password
        if format is None:
            result = self.store_doc(**kargs)
        else:
            kargs["format"] = format
            result = self.store_doc_format(**kargs)
        if result:
            self._events.trigger(LoNamedEvent.DOC_SAVED, EventArgs.from_args(cargs))
        return result

    # endregion save_doc()

    # region    store_doc()

    @overload
    def store_doc(self, store: XStorable, doc_type: LoDocType, fnm: PathOrStr) -> bool:
        ...

    @overload
    def store_doc(self, store: XStorable, doc_type: LoDocType, fnm: PathOrStr, password: str) -> bool:
        ...

    def store_doc(self, store: XStorable, doc_type: LoDocType, fnm: PathOrStr, password: Optional[str] = None) -> bool:
        cargs = CancelEventArgs(self.store_doc.__qualname__)
        cargs.event_data = {
            "store": store,
            "doc_type": doc_type,
            "fnm": fnm,
            "password": password,
        }
        self._events.trigger(LoNamedEvent.DOC_STORING, cargs)
        if cargs.cancel:
            return False
        ext = mInfo.Info.get_ext(fnm)
        frmt = "Text"
        if ext is None:
            self.print("Assuming a text format")
        else:
            frmt = self.ext_to_format(ext=ext, doc_type=doc_type)
        if password is None:
            self.store_doc_format(store=store, fnm=fnm, format=frmt)
        else:
            self.store_doc_format(store=store, fnm=fnm, format=frmt, password=password)
        self._events.trigger(LoNamedEvent.DOC_STORED, EventArgs.from_args(cargs))
        return True

    # endregion  store_doc()

    @overload
    def ext_to_format(self, ext: str) -> str:
        ...

    @overload
    def ext_to_format(self, ext: str, doc_type: LoDocType) -> str:
        ...

    def ext_to_format(self, ext: str, doc_type: LoDocType = LoDocType.UNKNOWN) -> str:
        dtype = LoDocType(doc_type)
        s = ext.lower()
        if s == "doc":
            return "MS Word 97"
        elif s == "docx":
            return "Office Open XML Text"  # MS Word 2007 XML
        elif s == "rtf":
            if dtype == LoDocType.CALC:
                return "Rich Text Format (StarCalc)"
            else:
                return "Rich Text Format"
        elif s == "odt":
            return "writer8"
        elif s == "ott":
            return "writer8_template"
        elif s == "pdf":
            if dtype == LoDocType.WRITER:
                return "writer_pdf_Export"
            elif dtype == LoDocType.IMPRESS:
                return "impress_pdf_Export"
            elif dtype == LoDocType.DRAW:
                return "draw_pdf_Export"
            elif dtype == LoDocType.CALC:
                return "calc_pdf_Export"
            elif dtype == LoDocType.MATH:
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
            if dtype == LoDocType.IMPRESS:
                return "impress_jpg_Export"
            else:
                return "draw_jpg_Export"
        elif s == "png":
            if dtype == LoDocType.IMPRESS:
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
            if dtype == LoDocType.WRITER:
                return "HTML (StarWriter)"
            elif dtype == LoDocType.IMPRESS:
                return "impress_html_Export"
            elif dtype == LoDocType.DRAW:
                return "draw_html_Export"
            elif dtype == LoDocType.CALC:
                return "HTML (StarCalc)"
            else:
                return "HTML"
        elif s == "xhtml":
            if dtype == LoDocType.WRITER:
                return "XHTML Writer File"
            elif dtype == LoDocType.IMPRESS:
                return "XHTML Impress File"
            elif dtype == LoDocType.DRAW:
                return "XHTML Draw File"
            elif dtype == LoDocType.CALC:
                return "XHTML Calc File"
            else:
                return "XHTML Writer File"
        elif s == "xml":
            if dtype == LoDocType.WRITER:
                return "OpenDocument Text Flat XML"
            elif dtype == LoDocType.IMPRESS:
                return "OpenDocument Presentation Flat XML"
            elif dtype == LoDocType.DRAW:
                return "OpenDocument Drawing Flat XML"
            elif dtype == LoDocType.CALC:
                return "OpenDocument Spreadsheet Flat XML"
            else:
                return "OpenDocument Text Flat XML"

        else:
            self.print(f"Do not recognize extension '{ext}'; using text")
            return "Text"

    # region    store_doc_format()

    @overload
    def store_doc_format(self, store: XStorable, fnm: PathOrStr, format: str) -> bool:
        ...

    @overload
    def store_doc_format(self, store: XStorable, fnm: PathOrStr, format: str, password: str) -> bool:
        ...

    def store_doc_format(self, store: XStorable, fnm: PathOrStr, format: str, password: str = None) -> bool:
        cargs = CancelEventArgs(self.store_doc_format.__qualname__)
        cargs.event_data = {
            "store": store,
            "format": format,
            "fnm": fnm,
            "password": password,
        }
        self._events.trigger(LoNamedEvent.DOC_STORING, cargs)
        if cargs.cancel:
            return False
        pth = mFileIO.FileIO.get_absolute_path(cargs.event_data["fnm"])
        fmt = str(cargs.event_data["format"])
        self.print(f"Saving the document in '{pth}'")
        self.print(f"Using format {fmt}")

        try:
            save_file_url = mFileIO.FileIO.fnm_to_url(pth)
            if password is None:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=fmt)
            else:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=fmt, Password=password)
            store.storeToURL(save_file_url, store_props)
        except IOException as e:
            raise Exception(f"Could not save '{pth}'") from e
        self._events.trigger(LoNamedEvent.DOC_STORED, EventArgs.from_args(cargs))
        return True

    # endregion store_doc_format()

    # ======================== document closing ==============

    @overload
    def close(self, closeable: XCloseable) -> bool:
        ...

    @overload
    def close(self, closeable: XCloseable, deliver_ownership: bool) -> bool:
        ...

    def close(self, closeable: XCloseable, deliver_ownership=False) -> bool:
        cargs = CancelEventArgs(self.close.__qualname__)
        cargs.event_data = deliver_ownership
        self._events.trigger(LoNamedEvent.DOC_CLOSING, cargs)
        if cargs.cancel:
            return False
        if closeable is None:
            return
        self.print("Closing the document")
        try:
            closeable.close(cargs.event_data)
            self._doc = None
            self._events.trigger(LoNamedEvent.DOC_CLOSED, EventArgs.from_args(cargs))
        except CloseVetoException as e:
            raise Exception("Close was vetoed") from e

    # region close_doc()
    @overload
    def close_doc(self, doc: object) -> None:
        ...

    @overload
    def close_doc(self, doc: object, deliver_ownership: bool) -> None:
        ...

    def close_doc(self, doc: object, deliver_ownership=False) -> None:
        if self._disposed:
            self._doc = None
            return
        try:
            closeable = self.qi(XCloseable, doc, True)
            self.close(closeable=closeable, deliver_ownership=deliver_ownership)
            self._doc = None
        except DisposedException as e:
            raise Exception("Document close failed since Office link disposed") from e

    # endregion close_doc()

    # ================= initialization via Addon-supplied context ====================

    def addon_initialize(self, addon_xcc: XComponentContext) -> XComponent:
        cargs = CancelEventArgs(self.addon_initialize.__qualname__)
        cargs.event_data = {"addon_xcc": addon_xcc}
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(LoNamedEvent.RESET, eargs)
        self._events.trigger(LoNamedEvent.DOC_OPENING, cargs)
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
        self._ms_factory = self.qi(XMultiServiceFactory, doc)
        if self._ms_factory in None:
            raise mEx.MissingInterfaceError(XMultiServiceFactory)
        self._doc = doc
        self._events.trigger(LoNamedEvent.DOC_OPENED, eargs)
        return doc

    # ============= initialization via script context ======================

    def script_initialize(self, sc: XScriptContext) -> XComponent:
        cargs = CancelEventArgs(self.script_initialize.__qualname__)
        cargs.event_data = {"sc": sc}
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(LoNamedEvent.RESET, eargs)
        self._events.trigger(LoNamedEvent.DOC_OPENING, cargs)
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
        self._ms_factory = self.qi(XMultiServiceFactory, doc)
        if self._ms_factory in None:
            raise mEx.MissingInterfaceError(XMultiServiceFactory)
        self._doc = doc
        self._events.trigger(LoNamedEvent.DOC_OPENED, eargs)
        return doc

    # ==================== dispatch ===============================
    # see https://wiki.documentfoundation.org/Development/DispatchCommands

    # region dispatch_cmd()
    @overload
    def dispatch_cmd(self, cmd: str) -> Any:
        ...

    @overload
    def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue]) -> Any:
        ...

    @overload
    def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> Any:
        ...

    @overload
    def dispatch_cmd(self, cmd: str, *, frame: XFrame) -> Any:
        ...

    def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue] = None, frame: XFrame = None) -> Any:
        if not cmd:
            raise mEx.DispatchError("cmd must not be empty or None")
        try:
            str_cmd = str(cmd)  # make sure and enum or other lookup did not get passed by mistake
            cargs = DispatchCancelArgs(self.dispatch_cmd.__qualname__, str_cmd)
            cargs.event_data = props
            self._events.trigger(LoNamedEvent.DISPATCHING, cargs)
            if cargs.cancel:
                raise mEx.CancelEventError(cargs, f'Dispatch Command "{str_cmd}" has been canceled')
            props = cargs.event_data
            if props is None:
                props = ()
            if frame is None:
                frame = self._xdesktop.getCurrentFrame()

            helper = self.create_instance_mcf(XDispatchHelper, "com.sun.star.frame.DispatchHelper")
            if helper is None:
                raise mEx.MissingInterfaceError(
                    XDispatchHelper, f"Could not create dispatch helper for command {str_cmd}"
                )
            result = helper.executeDispatch(frame, f".uno:{str_cmd}", "", 0, props)
            eargs = DispatchArgs.from_args(cargs)
            eargs.event_data = result
            self._events.trigger(LoNamedEvent.DISPATCHED, eargs)
            return result
        except mEx.CancelEventError:
            raise
        except Exception as e:
            raise mEx.DispatchError(f'Error dispatching "{cmd}"') from e

    # endregion dispatch_cmd()

    # ================= Uno cmds =========================

    @staticmethod
    def make_uno_cmd(self, item_name: str) -> str:
        return f"vnd.sun.star.script:Foo/Foo.{item_name}?language=Java&location=share"

    @staticmethod
    def extract_item_name(uno_cmd: str) -> str:
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

    def inspect(self, obj: object) -> None:
        if self._xcc is None or self._mc_factory is None:
            self.print("No office connection found")
            return
        try:
            ts = mInfo.Info.get_interface_types(obj)
            title = "Object"
            if ts is not None and len(ts) > 0:
                title = ts[0].getTypeName() + " " + title
            inspector = self._mc_factory.createInstanceWithContext("org.openoffice.InstanceInspector", self._xcc)
            #       hands on second use
            if inspector is None:
                self.print("Inspector Service could not be instantiated")
                return
            self.print("Inspector Service instantiated")
            intro = self.create_instance_mcf(XIntrospection, "com.sun.star.beans.Introspection")
            intro_acc = intro.inspect(inspector)
            method = intro_acc.getMethod("inspect", -1)
            self.print(f"inspect() method was found: {method is not None}")
            params = [[obj, title]]
            method.invoke(inspector, params)
        except Exception as e:
            self.print("Could not access Inspector:")
            self.print(f"    {e}")

    def mri_inspect(self, obj: object) -> None:
        # Available from http://extensions.libreoffice.org/extension-center/mri-uno-object-inspection-tool
        #               or http://extensions.services.openoffice.org/en/project/MRI
        #  Docs: https://github.com/hanya/MRI/wiki
        #  Forum tutorial: https://forum.openoffice.org/en/forum/viewtopic.php?f=74&t=49294
        xi = self.create_instance_mcf(XIntrospection, "mytools.Mri")
        if xi is None:
            raise Exception("MRI Inspector Service could not be instantiated")
        self.print("MRI Inspector Service instantiated")
        xi.inspect(obj)

    # ------------------ color methods ---------------------
    # section intentionally left out.

    # ================== other utils =============================

    def delay(self, ms: int) -> None:
        """
        Delay execution for a given number of milliseconds.

        Args:
            ms (int): Number of milliseconds to delay
        """
        if ms <= 0:
            self.print("Lo.delay(): Ms must be greater then zero")
            return
        sec = ms / 1000
        time.sleep(sec)

    wait = delay

    @staticmethod
    def is_none_or_empty(s: str) -> bool:
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
        return s.capitalize()

    def parse_int(self, s: str) -> int:
        if s is None:
            return 0
        try:
            return int(s)
        except ValueError:
            self.print(f"{s} could not be parsed as an int; using 0")
        return 0

    @overload
    @staticmethod
    def print_names(names: Iterable[str]) -> None:
        ...

    @overload
    @staticmethod
    def print_names(names: Sequence[str], num_per_line: int) -> None:
        ...

    @staticmethod
    def print_names(names: Sequence[str], num_per_line: int = 4) -> None:
        if not names:
            print("  No names found")
            return
        col_count = 1 if num_per_line < 1 else num_per_line

        lst_2d = mThelper.TableHelper.convert_1d_to_2d(
            seq_obj=sorted(names, key=str.casefold), col_count=col_count, empty_cell_val=""
        )
        longest = mThelper.TableHelper.get_largest_str(names)
        fmt_len = longest + 1
        if longest > 0:
            format_opt = FormatterTable(format=f"<{fmt_len}")
        else:
            format_opt = None

        indent = "  "
        print(f"No. of names: {len(names)}")
        if format_opt:
            acutal_count = len(lst_2d[0])
            if acutal_count > 1:
                # if this is more then on colum then print header
                #  -----------|-----------|-----------
                print(f"{indent}", end="")
                for i, _ in enumerate(range(acutal_count)):
                    print("-" * fmt_len, end="")
                    if i < acutal_count - 1:
                        print("|-", end="")
                print()
            for i, row in enumerate(lst_2d):
                col_str = format_opt.get_formatted(idx_row=i, row_data=row, join_str="| ")
                print(f"{indent}{col_str}")
        else:
            for row in lst_2d:
                for col in row:
                    print(f'{indent}"{col}"', end="")
                print()
        print("\n\n")

    # ------------------- container manipulation --------------------
    # region print_table()
    @overload
    @staticmethod
    def print_table(name: str, table: Table) -> None:
        ...

    @overload
    @staticmethod
    def print_table(name: str, table: Table, format_opt: FormatterTable) -> None:
        ...

    @staticmethod
    def print_table(name: str, table: Table, format_opt: FormatterTable | None = None) -> None:
        if format_opt:
            for i, row in enumerate(table):
                col_str = format_opt.get_formatted(idx_row=i, row_data=row)
                print(col_str)
        else:
            print(f"-- {name} ----------------")
            for row in table:
                col_str = "  ".join([str(el) for el in row])
                print(col_str)
        print()

    # endregion print_table()

    def get_container_names(self, con: XIndexAccess) -> List[str] | None:
        if con is None:
            self.print("Container is null")
            return None
        num_el = con.getCount()
        if num_el == 0:
            self.print("No elements in the container")
            return None

        names_list = []
        for i in range(num_el):
            named = con.getByIndex(i)
            names_list.append(named.getName())

        if len(names_list) == 0:
            self.print("No element names found in the container")
            return None
        return names_list

    def find_container_props(self, con: XIndexAccess, nm: str) -> XPropertySet | None:
        if con is None:
            raise TypeError("Container is null")
        for i in range(con.getCount()):
            try:
                el = con.getByIndex(i)
                named = self.qi(XNamed, el)
                if named and named.getName() == nm:
                    return self.qi(XPropertySet, el)
            except Exception:
                self.print(f"Could not access element {i}")
        self.print(f"Could not find a '{nm}' property set in the container")
        return None

    def is_uno_interfaces(self, component: object, *args: str | UnoInterface) -> bool:
        if len(args) == 0:
            return False
        result = True
        for arg in args:
            try:
                if isinstance(arg, str):
                    t = uno.getClass(arg)
                else:
                    t = arg
                obj = self.qi(t, component)
                if obj is None:
                    result = False
                    break
            except Exception:
                result = False
                break
        return result

    def get_frame(self) -> XFrame:
        if self.star_desktop is None:
            raise Exception("No desktop found")
        return self.XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
        # return cast(XDesktop, cls.star_desktop).getCurrentFrame()

    def get_model(self) -> XModel:
        return self.XSCRIPTCONTEXT.getDocument()

    def lock_controllers(self) -> bool:
        # much faster updates as screen is basically suspended
        cargs = CancelEventArgs(self.lock_controllers.__qualname__)
        self._events.trigger(LoNamedEvent.CONTROLERS_LOCKING, cargs)
        if cargs.cancel:
            return False
        xmodel = self.qi(XModel, self._doc, True)
        xmodel.lockControllers()
        self._events.trigger(LoNamedEvent.CONTROLERS_LOCKED, EventArgs(self))
        return True

    def unlock_controllers(self) -> bool:
        cargs = CancelEventArgs(self.unlock_controllers.__qualname__)
        self._events.trigger(LoNamedEvent.CONTROLERS_UNLOCKING, cargs)
        if cargs.cancel:
            return False
        xmodel = self.qi(XModel, self._doc, True)
        if xmodel.hasControllersLocked():
            xmodel.unlockControllers()
        self._events.trigger(LoNamedEvent.CONTROLERS_UNLOCKED, EventArgs.from_args(cargs))
        return True

    def has_controllers_locked(self) -> bool:
        xmodel = self.qi(XModel, self._doc)
        return xmodel.hasControllersLocked()

    def print(self, *args, **kwargs) -> None:
        if not self._allow_print:
            return
        cargs = CancelEventArgs(self.print.__qualname__)
        self._events.trigger(GblNamedEvent.PRINTING, cargs)
        if cargs.cancel:
            return
        print(*args, **kwargs)

    # region static methods
    @classmethod
    def from_existing(
        cls, script_ctx: XScriptContext, connector: LoBridgeCommon, opt: LoOptions | None = None
    ) -> LoInst:
        inst = LoInst(opt=opt)
        inst._init_from_ctx(script_ctx, connector)
        inst._xscript_context = script_ctx
        return inst

    # endregion static methods

    @property
    def null_date(self) -> datetime:
        # https://tinyurl.com/2pdrt5z9#NullDate
        try:
            return self._null_date
        except AttributeError:
            self._null_date = datetime(year=1889, month=12, day=30, tzinfo=timezone.utc)
            if self._doc is None:
                return self._null_date
            n_supplier = self.qi(XNumberFormatsSupplier, self._doc)
            if n_supplier is None:
                # this is not always a XNumberFormatsSupplier such as *.odp documents
                return self._null_date
            number_settings = n_supplier.getNumberFormatSettings()
            d = number_settings.getPropertyValue("NullDate")
            self._null_date = datetime(d.Year, d.Month, d.Day, tzinfo=timezone.utc)
        return self._null_date

    @property
    def is_loaded(self) -> bool:
        return not self._lo_inst is None

    @property
    def is_macro_mode(self) -> bool:
        try:
            return self._is_macro_mode
        except AttributeError:
            if self._lo_inst is None:
                return False
            self._is_macro_mode = isinstance(self._lo_inst, LoDirectStart)
        return self._is_macro_mode

    @property
    def star_desktop(self) -> XDesktop:
        """Get current desktop"""
        return self._xdesktop

    StarDesktop, stardesktop = star_desktop, star_desktop

    @property
    def this_component(self) -> XComponent:
        try:
            return self._this_component
        except AttributeError:
            if mock_g.DOCS_BUILDING:
                self._this_component = None
                return self._this_component
            if self.is_loaded is False:
                # attempt to connect direct
                # failure will result in script error and then exit
                self.load_office()

            # comp = self.star_desktop.getCurrentComponent()
            desktop = self.get_desktop()
            if desktop is None:
                return None
            if self._doc is None:
                self._doc = desktop.getCurrentComponent()
            if self._doc is None:
                return None
            impl = self._doc.ImplementationName
            if impl in ("com.sun.star.comp.basic.BasicIDE", "com.sun.star.comp.sfx2.BackingComp"):
                return None  # None when Basic IDE or welcome screen
            self._this_component = self._doc
            return self._this_component

    ThisComponent, thiscomponent = this_component, this_component

    @property
    def xscript_context(self) -> XScriptContext:
        try:
            return self._xscript_context
        except AttributeError:
            ctx = self.get_context()
            if ctx is None:
                # attempt to connect direct
                # failure will result in script error and then exit
                self.load_office()
                ctx = self.get_context()

            desktop = self.get_desktop()
            model = self.qi(XModel, self._doc)
            self._xscript_context = script_context.ScriptContext(ctx=ctx, desktop=desktop, doc=model)
        return self._xscript_context

    XSCRIPTCONTEXT = xscript_context

    @property
    def bridge(self) -> XComponent:
        try:
            return self._bridge_component
        except AttributeError:
            try:
                # when running as macro self._lo_inst will not have bridge_component
                self._bridge_component = self._lo_inst.bridge_component
            except AttributeError:
                self._bridge_component = None
            return self._bridge_component

    @property
    def loader_current(self) -> XComponentLoader:
        return self._loader

    @property
    def bridge_connector(self) -> LoBridgeCommon:
        return self._lo_inst

    @property
    def options(self) -> LoOptions:
        return self._opt


__all__ = ("LoInst",)
