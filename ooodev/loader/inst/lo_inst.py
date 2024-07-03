"""Lo Instance Module - General entry point for all LibreOffice related functionality as an instance."""

# Python conversion of Lo.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

# pylint: disable=too-many-lines
# pylint: disable=missing-function-docstring
# pylint: disable=broad-exception-raised
# pylint: disable=broad-exception-caught

from __future__ import annotations
import contextlib
from pathlib import Path
import threading
from datetime import datetime, timezone
import time
from typing import Any, cast, Iterable, List, Optional, overload, Sequence, Tuple, TYPE_CHECKING, Type
from urllib.parse import urlparse
import uno
from com.sun.star.beans import XIntrospection
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNamed
from com.sun.star.frame import XComponentLoader
from com.sun.star.frame import XDesktop
from com.sun.star.frame import XDispatchHelper
from com.sun.star.frame import XDispatchProvider
from com.sun.star.frame import XModel
from com.sun.star.frame import XStorable
from com.sun.star.io import IOException
from com.sun.star.lang import XComponent
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.lang import XServiceInfo
from com.sun.star.util import XCloseable
from com.sun.star.util import XNumberFormatsSupplier

from ooo.dyn.document.macro_exec_mode import MacroExecMode  # const
from ooo.dyn.lang.disposed_exception import DisposedException
from ooo.dyn.util.close_veto_exception import CloseVetoException


# if not mock_g.DOCS_BUILDING:
# not importing for doc building just result in short import name for
# args that use these.
# this is also true because docs/conf.py ignores com import for autodoc
import ooodev
from ooodev.mock import mock_g

# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from ooodev.loader.inst.lo_loader import LoLoader
from ooodev.adapter.lang.event_listener import EventListener
from ooodev.conn import cache as mCache
from ooodev.conn import connectors
from ooodev.conn.connect import ConnectBase, LoDirectStart
from ooodev.utils.data_type.generic_size_pos import GenericSizePos
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.dispatch_args import DispatchArgs
from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.event_singleton import _Events
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import Events
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.formatters.formatter_table import FormatterTable
from ooodev.loader.comp.the_desktop import TheDesktop
from ooodev.loader.comp.the_global_event_broadcaster import TheGlobalEventBroadcaster
from ooodev.loader.inst.doc_type import DocType as LoDocType, DocTypeStr as LoDocTypeStr
from ooodev.loader.inst.options import Options as LoOptions
from ooodev.utils import file_io as mFileIO
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps
from ooodev.utils import script_context
from ooodev.utils import table_helper as mThelper
from ooodev.utils.factory.doc_factory import doc_factory
from ooodev.utils.cache.lru_cache import LRUCache
from ooodev.io.log import logging as logger
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    try:
        from typing import Literal  # Py >= 3.8
    except ImportError:
        from typing_extensions import Literal
    from ooo.dyn.beans.property_value import PropertyValue
    from com.sun.star.container import XChild
    from com.sun.star.container import XIndexAccess
    from com.sun.star.frame import XFrame
    from com.sun.star.lang import EventObject
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.script.provider import XScriptContext
    from com.sun.star.uno import XComponentContext
    from com.sun.star.uno import XInterface
    from com.sun.star.util import Date as UnoDate
    from ooodev.proto.event_observer import EventObserver
    from ooodev.proto.office_document_t import OfficeDocumentT
    from com.sun.star.document import DocumentEvent
    from ooodev.utils.type_var import PathOrStr
    from ooodev.utils.type_var import UnoInterface
    from ooodev.utils.type_var import T
    from ooodev.utils.type_var import Table
    from ooodev.adapter.awt.unit_conversion_comp import UnitConversionComp
else:
    PathOrStr = Any
    UnoInterface = Any
    T = Any
    Table = Any
    # from ooodev.events.events_t import EventsT


class LoInst(EventsPartial):
    """
    Lo instance class. This is the main class for interacting with LO.

    This class is of advanced usage and is not intended for general use.

    In most cases, you should use the :py:class:`~ooodev.loader.Lo` static class instead.

    This class mirrors the properties and methods of the ``Lo`` class so for documentation see :py:class:`ooodev.utils.lo.Lo`,
    """

    def __init__(self, opt: LoOptions | None = None, events: EventObserver | None = None, **kwargs) -> None:
        """
        Constructor

        Args:
            opt (LoOptions, optional): Options
            events (EventObserver, optional): Event observer

        Hint:
            - ``LoOptions`` can be imported from ``ooodev.loader.inst.options``
        """
        self._singleton_instance = kwargs.get("is_singleton", False)
        self._opt = LoOptions() if opt is None else opt
        if self._singleton_instance:
            # only set the log level if thi instance is the ooodev.loader.lo.Lo instance
            logger.set_log_level(self._opt.log_level)
        if self._singleton_instance:
            self._logger = NamedLogger(f"{self.__class__.__name__} - Root")
        else:
            self._logger = NamedLogger(self.__class__.__name__)
        self._logger.debug("LoInst Init")
        self._is_default = False
        self._current_doc = None
        _events = Events(source=self) if events is None else events
        EventsPartial.__init__(self, _events)
        self._xcc: XComponentContext | None = None
        """remote component context"""
        self._xdesktop: TheDesktop | None = None
        self._glb_event_broadcaster: TheGlobalEventBroadcaster | None = None
        """remote desktop UNO service"""
        self._mc_factory: XMultiComponentFactory | None = None
        self._is_office_terminated: bool = False
        self._lo_inst: ConnectBase | None = None
        self._loader = None
        self._disposed = True
        self._lo_loader: LoLoader | None = None
        self._app_font_pixel_ratio = None
        self._sys_font_pixel_ratio = None
        self._cache = LRUCache(50)
        self._shared_cache = LRUCache(self._opt.lo_cache_size)

        self._allow_print = self._opt.verbose
        self._set_lo_events()
        self._logger.debug("LoInst created")

    # region Events
    def _set_lo_events(self) -> None:
        self._fn_on_lo_del_cache_attrs = self.on_lo_del_cache_attrs
        self._fn_on_lo_loading = self.on_lo_loading
        self._fn_on_lo_loaded = self.on_lo_loaded
        self._fn_on_lo_disposed = self.on_lo_disposed
        self._fn_on_lo_disposing_bridge = self.on_lo_disposing_bridge
        self._fn_on_document_event = self.on_global_document_event

        self._event_listener = EventListener()
        self._event_listener.on("disposing", self._fn_on_lo_disposing_bridge)

        self.subscribe_event(LoNamedEvent.RESET, self._fn_on_lo_del_cache_attrs)
        self.subscribe_event(LoNamedEvent.OFFICE_LOADING, self._fn_on_lo_loading)
        self.subscribe_event(LoNamedEvent.OFFICE_LOADED, self._fn_on_lo_loaded)
        self.subscribe_event(LoNamedEvent.BRIDGE_DISPOSED, self._fn_on_lo_disposed)
        _Events().on(GblNamedEvent.DOCUMENT_EVENT, self._fn_on_document_event)

    def on_reset(self, event_args: EventArgs) -> None:
        self._logger.debug("on_reset() Triggering RESET")
        self.trigger_event(LoNamedEvent.RESET, event_args)
        self._current_doc = None
        self._logger.debug("on_reset() RESET Triggered. Current Doc is None")

    def on_doc_creating(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_doc_creating() Triggering DOC_CREATING")
        self.trigger_event(LoNamedEvent.DOC_CREATING, event_args)
        self._current_doc = None
        self._logger.debug("on_doc_creating() Triggered DOC_CREATING. Current Doc is None")

    def on_doc_created(self, event_args: EventArgs) -> None:
        self._logger.debug("on_doc_created() Triggering DOC_CREATED")
        self.trigger_event(LoNamedEvent.DOC_CREATED, event_args)
        self._current_doc = None
        self._logger.debug("on_doc_created() Triggered DOC_CREATED. Current Doc is None")

    def on_doc_opening(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_doc_opening() Triggering DOC_OPENING")
        self.trigger_event(LoNamedEvent.DOC_OPENING, event_args)
        self._logger.debug("on_doc_opening() Triggered DOC_OPENING.")

    def on_doc_opened(self, event_args: EventArgs) -> None:
        self._logger.debug("on_doc_opened() Triggering DOC_OPENED")
        self.trigger_event(LoNamedEvent.DOC_OPENED, event_args)
        self._current_doc = None
        self._logger.debug("on_doc_opened() Triggered DOC_OPENED. Current Doc is None")

    def on_doc_closing(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_doc_closing() Triggering DOC_CLOSING")
        self.trigger_event(LoNamedEvent.DOC_CLOSING, event_args)
        self._logger.debug("on_doc_closing() Triggered DOC_CLOSING")

    def on_doc_closed(self, event_args: EventArgs) -> None:
        self._logger.debug("on_doc_closed() Triggering DOC_CLOSED")
        self.trigger_event(LoNamedEvent.DOC_CLOSED, event_args)
        self._current_doc = None
        self._logger.debug("on_doc_closed() Triggered DOC_CLOSED. Current Doc is None")

    def on_doc_saving(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_doc_saving() Triggering DOC_SAVING")
        self.trigger_event(LoNamedEvent.DOC_SAVING, event_args)
        self._logger.debug("on_doc_saving() Triggered DOC_SAVING")

    def on_doc_saved(self, event_args: EventArgs) -> None:
        self._logger.debug("on_doc_saved() Triggering DOC_SAVED")
        self.trigger_event(LoNamedEvent.DOC_SAVED, event_args)
        self._logger.debug("on_doc_saved() Triggered DOC_SAVED")

    def on_doc_storing(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_doc_storing() Triggering DOC_STORING")
        self.trigger_event(LoNamedEvent.DOC_STORING, event_args)
        self._logger.debug("on_doc_storing() Triggered DOC_STORING")

    def on_doc_stored(self, event_args: EventArgs) -> None:
        self._logger.debug("on_doc_stored() Triggering DOC_STORED")
        self.trigger_event(LoNamedEvent.DOC_STORED, event_args)
        self._logger.debug("on_doc_stored() Triggered DOC_STORED")

    def on_office_loading(self, event_args: EventArgs) -> None:
        self._logger.debug("on_office_loading() Triggering OFFICE_LOADING")
        self.trigger_event(LoNamedEvent.OFFICE_LOADING, event_args)
        self._logger.debug("on_office_loading() Triggered OFFICE_LOADING.")

    def on_office_loaded(self, event_args: EventArgs) -> None:
        self._logger.debug("on_office_loaded() Triggering OFFICE_LOADED")
        self.trigger_event(LoNamedEvent.OFFICE_LOADED, event_args)
        self._logger.debug("on_office_loaded() Triggered OFFICE_LOADED")

    def on_office_closing(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_office_closing() Triggering OFFICE_CLOSING")
        self.trigger_event(LoNamedEvent.OFFICE_CLOSING, event_args)
        if event_args.cancel:
            return
        self._clear_cache()
        self._glb_event_broadcaster = None
        self._current_doc = None
        self._mc_factory = None
        self._xdesktop = None
        self._xcc = None
        self._fn_on_document_event = None
        self._fn_on_lo_del_cache_attrs = None
        self._logger.debug("on_office_closing() Triggered OFFICE_CLOSING")

    def on_office_closed(self, event_args: EventArgs) -> None:
        self._logger.debug("on_office_closed() Triggering OFFICE_CLOSED")
        self.trigger_event(LoNamedEvent.OFFICE_CLOSED, event_args)
        self._logger.debug("on_office_closed() Triggered OFFICE_CLOSED")

    def on_component_loading(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_component_loading() Triggering COMPONENT_LOADING")
        self.trigger_event(LoNamedEvent.COMPONENT_LOADING, event_args)
        self._logger.debug("on_component_loading() Triggered COMPONENT_LOADING")

    def on_component_loaded(self, event_args: EventArgs) -> None:
        self._logger.debug("on_component_loaded() Triggering COMPONENT_LOADED")
        self.trigger_event(LoNamedEvent.COMPONENT_LOADED, event_args)
        self._logger.debug("on_component_loaded() Triggered COMPONENT_LOADED")

    def on_dispatching(self, event_args: DispatchCancelArgs) -> None:
        self._logger.debug("on_dispatching() Triggering DISPATCHING")
        self.trigger_event(LoNamedEvent.DISPATCHING, event_args)
        self._logger.debug("on_dispatching() Triggered DISPATCHING")

    def on_dispatched(self, event_args: DispatchArgs) -> None:
        self._logger.debug("on_dispatched() Triggering DISPATCHED")
        self.trigger_event(LoNamedEvent.DISPATCHED, event_args)
        self._logger.debug("on_dispatched() Triggered DISPATCHED")

    def on_controllers_locking(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_controllers_locking() Triggering CONTROLLERS_LOCKING")
        self.trigger_event(LoNamedEvent.CONTROLLERS_LOCKING, event_args)
        self._logger.debug("on_controllers_locking() Triggered CONTROLLERS_LOCKING")

    def on_controllers_locked(self, event_args: EventArgs) -> None:
        self._logger.debug("on_controllers_locked() Triggering CONTROLLERS_LOCKED")
        self.trigger_event(LoNamedEvent.CONTROLLERS_LOCKED, event_args)
        self._logger.debug("on_controllers_locked() Triggered CONTROLLERS_LOCKED")

    def on_controllers_unlocking(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_controllers_unlocking() Triggering CONTROLLERS_UNLOCKING")
        self.trigger_event(LoNamedEvent.CONTROLLERS_UNLOCKING, event_args)
        self._logger.debug("on_controllers_unlocking() Triggered CONTROLLERS_UNLOCKING")

    def on_controllers_unlocked(self, event_args: EventArgs) -> None:
        self._logger.debug("on_controllers_unlocked() Triggering CONTROLLERS_UNLOCKED")
        self.trigger_event(LoNamedEvent.CONTROLLERS_UNLOCKED, event_args)
        self._logger.debug("on_controllers_unlocked() Triggered CONTROLLERS_UNLOCKED")

    def on_printing(self, event_args: CancelEventArgs) -> None:
        self._logger.debug("on_printing() Triggering PRINTING")
        self.trigger_event(GblNamedEvent.PRINTING, event_args)
        self._logger.debug("on_printing() Triggered PRINTING")

    def on_lo_del_cache_attrs(self, source: object, event: EventArgs) -> None:  # pylint: disable=unused-argument
        # clears Lo Attributes that are dynamically created
        data_attrs = ("_xscript_context", "_is_macro_mode", "_this_component", "_bridge_component", "_null_date")
        for attr in data_attrs:
            if hasattr(self, attr):
                delattr(self, attr)

    # pylint: disable=unused-argument
    def on_lo_disposing_bridge(self, src: EventListener, event: EventObject) -> None:
        self.trigger_event(LoNamedEvent.BRIDGE_DISPOSED, EventArgs(self.on_lo_disposing_bridge.__qualname__))

    def on_lo_disposed(self, source: Any, event: EventObject) -> None:  # pylint: disable=unused-argument
        self._logger.debug("Office bridge has gone!!")
        data_attrs = ("_xcc", "_doc", "_mc_factory", "_ms_factory", "_lo_inst", "_xdesktop", "_loader")
        data_vals = (None, None, None, None, None, None, None)
        for attr, val in zip(data_attrs, data_vals):
            setattr(self, attr, val)
        setattr(self, "_disposed", True)

    def on_lo_loaded(self, source: Any, event: EventObject) -> None:  # pylint: disable=unused-argument
        if self.bridge is not None:
            self.bridge.addEventListener(self._event_listener)

    def on_lo_loading(self, source: Any, event: CancelEventArgs) -> None:  # pylint: disable=unused-argument
        with contextlib.suppress(Exception):
            if hasattr(self, "_bridge_component") and self.bridge is not None:
                self.bridge.removeEventListener(self._event_listener)

    def on_global_document_event(
        self, src: Any, event: EventArgs, *args, **kwargs
    ) -> None:  # pylint: disable=unused-argument
        with contextlib.suppress(Exception):
            doc_event = cast("DocumentEvent", event.event_data)
            name = doc_event.EventName
            if name == "OnUnfocus":
                self._clear_cache()
                self._logger.debug("on_global_document_event() Cleared Cache")

    # endregion Events

    # region    qi()

    # pylint: disable=invalid-name
    @overload
    def qi(self, atype: Type[T], obj: Any) -> T | None: ...

    # pylint: disable=invalid-name
    @overload
    def qi(self, atype: Type[T], obj: Any, raise_err: Literal[True]) -> T: ...

    @overload
    def qi(self, atype: Type[T], obj: Any, raise_err: Literal[False]) -> T | None: ...

    def qi(self, atype: Type[T], obj: XTypeProvider, raise_err: bool = False) -> T | None:
        """
        Generic method that get an interface instance from  an object.

        |lo_safe|

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
        """
        result = None
        if uno.isInterface(atype) and hasattr(obj, "queryInterface"):
            uno_t = uno.getTypeByName(atype.__pyunointerface__)  # type: ignore
            result = obj.queryInterface(uno_t)
        if raise_err and result is None:
            raise mEx.MissingInterfaceError(atype)
        if result is not None:
            result = cast(T, result)
        return result

    # endregion qi()

    # region    cache
    def _clear_cache(self) -> None:
        self._current_doc = None
        self._cache.clear()

    # endregion cache

    def _get_font_ratio(self):
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        key = "_get_font_ratio"
        if key in self._cache:
            return self._cache[key]
        from ooodev.adapter.awt.unit_conversion_comp import UnitConversionComp
        from ooo.dyn.awt.point import Point as UnoPoint
        from ooo.dyn.awt.size import Size as UnitSize

        comp = UnitConversionComp(self)

        def get_ratio(target_unit: int) -> GenericSizePos[float]:
            nonlocal comp
            uno_p = comp.convert_point_to_logic(UnoPoint(100_000, 100_000), target_unit)
            uno_sz = comp.convert_size_to_logic(UnitSize(100_000, 100_000), target_unit)
            r_x = uno_p.X / 100_000
            r_y = uno_p.Y / 100_000
            r_w = uno_sz.Width / 100_000
            r_h = uno_sz.Height / 100_000
            return GenericSizePos(r_x, r_y, r_w, r_h)

        self._cache[key] = get_ratio
        return self._cache[key]

    def get_context(self) -> XComponentContext:
        """
        Gets current LO Component Context.

        |lo_unsafe|
        """
        return self._xcc  # type: ignore

    def get_desktop(self) -> XDesktop:
        """
        Gets current LO Desktop.

        |lo_unsafe|
        """
        if self._opt.dynamic:
            return self.xscript_context.getDesktop()
        return self.desktop.component  # type: ignore

    def get_component_factory(self) -> XMultiComponentFactory:
        """
        Gets current multi component factory.

        |lo_unsafe|
        """
        return self._mc_factory  # type: ignore

    def get_service_factory(self) -> XMultiServiceFactory:
        """
        Gets current multi service factory.

        |lo_unsafe|
        """
        # return cls._bridge_component
        doc = self.desktop.get_current_component()
        return self.qi(XMultiServiceFactory, doc, True)

    def get_relative_doc(self) -> XComponent:
        """
        Gets current document.

        If the current options are set to dynamic,
        then the current document is returned from the script context.
        Otherwise, the current internal document is returned.
        The internal document is set when a new document is created via the Write, Calc, etc.

        In most instances the internal document is the same as the xscript context document.

        By Default the dynamic option is set to ``False``.

        |lo_unsafe|

        Raises:
            NoneError: If the document is ``None``.

        Returns:
            XComponent: Current Document
        """
        if self._opt.dynamic:
            doc = self.xscript_context.getDesktop().getCurrentComponent()
            if doc is None:
                raise mEx.NoneError("Document is None")
            return doc
        return self.desktop.get_current_component()

    # region interface object creation

    # region    create_instance_msf()
    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str) -> T | None: ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, msf: Any | None) -> T | None: ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, *, raise_err: Literal[True]) -> T: ...

    @overload
    def create_instance_msf(self, atype: Type[T], service_name: str, *, raise_err: Literal[False]) -> T | None: ...

    @overload
    def create_instance_msf(
        self, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[True]
    ) -> T: ...

    @overload
    def create_instance_msf(
        self, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[True], *args: Any
    ) -> T: ...

    @overload
    def create_instance_msf(
        self, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[False]
    ) -> T | None: ...

    @overload
    def create_instance_msf(
        self, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[False], *args: Any
    ) -> T | None: ...

    def create_instance_msf(
        self,
        atype: Type[T],
        service_name: str,
        msf: XMultiServiceFactory | None = None,
        raise_err: bool = False,
        *args: Any,
    ) -> T | None:  # sourcery skip: raise-specific-error
        # in a multi document environment, get the multi service factory from the current document
        if msf is None:
            doc = self.desktop.get_current_component()
            factory = self.qi(XMultiServiceFactory, doc)
            if factory is None:
                raise Exception("No document found")
        else:
            factory = msf

        try:
            if args:
                obj = factory.createInstanceWithArguments(service_name, args)
            else:
                obj = factory.createInstance(service_name)
            if raise_err and obj is None:
                raise mEx.CreateInstanceMsfError(atype, service_name)
            interface_obj = self.qi(atype=atype, obj=obj)
            if raise_err and interface_obj is None:
                raise mEx.MissingInterfaceError(atype)
            return interface_obj
        except mEx.CreateInstanceMsfError:
            raise
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            self._logger.exception(f"Couldn't create interface for '{service_name}'")
            raise Exception(f"Couldn't create interface for '{service_name}'") from e

    # endregion create_instance_msf()

    # region    create_instance_mcf()
    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str) -> T | None: ...

    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None) -> T | None: ...

    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str, *, raise_err: Literal[True]) -> T: ...

    @overload
    def create_instance_mcf(self, atype: Type[T], service_name: str, *, raise_err: Literal[False]) -> T | None: ...

    @overload
    def create_instance_mcf(
        self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None, raise_err: Literal[True]
    ) -> T: ...

    @overload
    def create_instance_mcf(
        self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None, raise_err: Literal[False]
    ) -> T | None: ...

    def create_instance_mcf(
        self, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None = None, raise_err: bool = False
    ) -> T | None:
        # sourcery skip: raise-specific-error
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
            if raise_err and obj is None:
                raise mEx.CreateInstanceMcfError(atype, service_name)
            interface_obj = self.qi(atype=atype, obj=obj)
            if raise_err and interface_obj is None:
                raise mEx.MissingInterfaceError(atype)
            return interface_obj
        except mEx.CreateInstanceMcfError:
            self._logger.error(f"CreateInstanceMcfError Error. Couldn't create instance for '{service_name}'")
            raise
        except mEx.MissingInterfaceError:
            self._logger.exception(f"MissingInterfaceError Couldn't get interface for '{service_name}'")
            raise
        except Exception as e:
            self._logger.exception(f"Couldn't create interface for '{service_name}'")
            raise Exception(f"Couldn't create interface for '{service_name}'") from e

    # endregion create_instance_mcf()

    # endregion interface object creation

    def get_parent(self, a_component: XChild) -> XInterface:
        return a_component.getParent()

    # region Start Office

    def load_office(
        self,
        connector: connectors.ConnectPipe | connectors.ConnectSocket | ConnectBase | None = None,
        cache_obj: mCache.Cache | None = None,
        opt: LoOptions | None = None,
    ) -> XComponentLoader:
        """
        Loads Office

        Not available in a macro.

        If running outside of office then a bridge is created that connects to office.

        If running from inside of office e.g. in a macro, then ``Lo.XSCRIPTCONTEXT`` is used.
        ``using_pipes`` is ignored with running inside office.

        |lo_unsafe|

        Args:
            connector (connectors.ConnectPipe, connectors.ConnectSocket, ConnectBase, None): Connection information. Ignore for macros.
            cache_obj (Cache, optional): Cache instance that determines of LibreOffice profile is to be copied and cached
                Ignore for macros. Defaults to None.
            opt (Options, optional): Extra Load options.

        Returns:
            XComponentLoader: component loader

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.OFFICE_LOADING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.OFFICE_LOADED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method parameters.

        Note:
            When using this class to connect to a running office instance,  the ``connector`` parameter should be set to ``Lo.bridge_connector``, the existing bridge connection.

            The following example creates a new instance, connects to a running office instance and creates a new calc document.

            .. code-block:: python

                lo = LoInst()
                lo.load_office(Lo.bridge_connector)
                lo_doc = lo.create_doc(DocTypeStr.CALC)

        See Also:
            - :ref:`ch02`
            - :ref:`ch02_multiple_docs`
            - :py:meth:`open_doc`
            - :py:class:`~ooodev.loader.Lo`
            - :py:class:`~ooodev.utils.lo.Lo.Loader`

        .. versionchanged:: 0.6.10
            Added ``opt`` parameter.

        .. versionchanged:: 0.9.8
            connector can now also be ``ConnectBase`` to allow for existing connections.
        """
        # sourcery skip: class-extract-method, extract-duplicate-method
        if mock_g.DOCS_BUILDING:
            # some component call this method and are triggered during docs building.
            # by adding this block this method will be exited if docs are building.
            return None  # type: ignore
        self._logger.debug("load_office()")
        # Creation sequence: remote component content (xcc) -->
        #                     remote service manager (mcFactory) -->
        #                     remote desktop (xDesktop) -->
        #                     component loader (XComponentLoader)
        # Once we have a component loader, we can load a document.
        # xcc, mcFactory, and xDesktop are stored as static globals.
        if self._is_default:
            raise mEx.LoadingError("Cannot set loader for default instance")

        cargs = CancelEventArgs(self.load_office.__qualname__)

        cargs.event_data = {
            "connector": connector,
        }

        eargs = EventArgs.from_args(cargs)
        self.on_reset(eargs)

        self.on_office_loading(eargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        b_connector = cargs.event_data["connector"]
        lo_loader = LoLoader(connector=b_connector, cache_obj=cache_obj, opt=opt)
        loader = self.load_from_lo_loader(lo_loader)
        if self._singleton_instance:
            self._is_default = True
        self.on_office_loaded(eargs)
        self._logger.debug("load_office() Loaded Office")
        return loader

    def load_from_lo_loader(self, loader: LoLoader) -> XComponentLoader:
        """
        Loads Office from a LoLoader instance.

        |lo_unsafe|

        Args:
            loader (LoLoader): LoLoader instance.

        Returns:
            XComponentLoader: component loader.

        .. versionadded:: 0.40.0
        """
        self._lo_loader = loader
        self._lo_inst = self._lo_loader.lo_inst
        self._opt = self._lo_loader.options
        self._disposed = False
        self._xcc = self._lo_inst.ctx
        self._load_from_context()
        if self._loader is None:
            self._logger.error("load_from_lo_loader() Unable to access XComponentLoader")
            raise mEx.LoadingError("Unable to access XComponentLoader")
        return self._loader

    def get_singleton(self, name: str) -> Any:
        """
        Gets a singleton object from the office default context.

        Args:
            name (str): Singleton name such as ``/singletons/com.sun.star.frame.theDesktop``

        Returns:
            Any: Singleton object or ``None`` if not found
        """
        if self._mc_factory is None:
            return None
        result = None
        with contextlib.suppress(Exception):
            result = self._mc_factory.DefaultContext.getByName(name)  # type: ignore
        return result

    def _load_from_context(self) -> None:
        if self._xcc is None:
            self._logger.error("No component context found")
            raise mEx.LoadingError("No component context found")
        self._mc_factory = self._xcc.getServiceManager()
        if self._mc_factory is None:
            self._logger.error("Office Service Manager is unavailable")
            raise mEx.LoadingError("Office Service Manager is unavailable")
        desktop: Any = self._mc_factory.DefaultContext.getByName("/singletons/com.sun.star.frame.theDesktop")  # type: ignore
        self._xdesktop = TheDesktop(desktop)
        gb = self._mc_factory.DefaultContext.getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")  # type: ignore
        self._glb_event_broadcaster = TheGlobalEventBroadcaster(gb)
        if self._xdesktop is None:
            self._logger.error("Could not create a desktop service")
            raise mEx.LoadingError("Could not create a desktop service")
        # self._xdesktop.add_event_frame_action(self._fn_on_context_changed)
        self._loader = self.qi(XComponentLoader, self._xdesktop.component)
        if self._loader is None:
            self._logger.error("Unable to access XComponentLoader")
            raise mEx.LoadingError("Unable to access XComponentLoader")

    def _reset_for_doc(self, doc: XComponent) -> None:
        """
        Gets the frame from the document and sets it as the active frame for the desktop.

        Sets the multi service factory for the document.

        Setting the active frame is necessary for the document to be accessible via ``desktop.get_current_component()``.

        Args:
            doc (XComponent): Document to reset for.
        """
        model = self.qi(XModel, doc, True)
        controller = model.getCurrentController()
        frm = controller.getFrame()
        self.desktop.set_active_frame(frm)

        # With some components such as Base, the multi service factory is not available
        self._ms_factory = self.qi(XMultiServiceFactory, doc)

    # endregion Start Office

    # region office shutdown
    def close_office(self) -> bool:
        self._logger.debug("Closing Office")

        cargs = CancelEventArgs(self.close_office.__qualname__)
        self.on_office_closing(cargs)
        if cargs.cancel:
            return False

        # if self._xdesktop is None:
        #     self._logger.debug("No office connection found")
        #     return True

        if self._is_office_terminated:
            self._logger.debug("Office has already been requested to terminate")
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
            self.on_office_closed(eargs)
            self.on_reset(eargs)
        return self._is_office_terminated

    def _try_to_terminate(self, num_tries: int) -> bool:
        if self._disposed:
            return True
        if self._xdesktop is None:
            return True
        try:
            is_dead = self._xdesktop.terminate()
            if is_dead:
                if num_tries > 1:
                    self._logger.debug(f"{num_tries}. Office terminated")
                else:
                    self._logger.debug("Office terminated")
            else:
                self._logger.debug(f"{num_tries}. Office failed to terminate")
            return is_dead
        except DisposedException:
            self._logger.debug("Office link disposed")
            return True
        except Exception as e:
            self._logger.debug(f"Termination exception: {e}")
            return False

    def kill_office(self) -> None:
        # sourcery skip: extract-method, raise-specific-error
        if self._lo_inst is None:
            self._logger.debug("No instance to kill")
            return
        try:
            # raised a NotImplementedError when self._lo_inst is direct (macro mode)
            self._lo_inst.kill_soffice()
            self._is_office_terminated = True
            eargs = EventArgs(self.kill_office.__qualname__)
            self.on_office_closed(eargs)
            self.on_reset(eargs)
            self._lo_inst = None
            # self._logger.debug("Killed Office")
        except Exception as e:
            raise Exception("Unable to kill Office") from e

    # endregion office shutdown

    # region document opening

    def load_component(self, component: XComponent) -> None:
        """
        Loads a component, can be a document or a sub-component such as a database sub-form

        Args:
            component (XComponent): Component to load

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.COMPONENT_LOADING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.COMPONENT_LOADED` :eventref:`src-docs-event`

        Note:
            When and instance of this class is created for the purpose of loading a component then
            the :py:class:`~ooodev.utils.inst.lo.options` should have it ``dynamic`` property set to ``False``.

        .. versionadded:: 0.9.8
        """
        if self._is_default:
            self._logger.error("Cannot set loader for default instance")
            raise mEx.LoadingError("Cannot set loader for default instance")
        cargs = CancelEventArgs(self.open_doc.__qualname__)
        cargs.event_data = component
        self.on_component_loading(cargs)
        if cargs.cancel:
            return
        eargs = EventArgs.from_args(cargs)
        self.on_reset(eargs)
        self._reset_for_doc(component)
        self.on_component_loaded(eargs)

    # region open_flat_doc()
    @overload
    def open_flat_doc(self, fnm: PathOrStr, doc_type: LoDocType) -> XComponent: ...

    @overload
    def open_flat_doc(self, fnm: PathOrStr, doc_type: LoDocType, loader: XComponentLoader) -> XComponent: ...

    def open_flat_doc(
        self, fnm: PathOrStr, doc_type: LoDocType, loader: Optional[XComponentLoader] = None
    ) -> XComponent:
        if loader is None:
            loader = cast(XComponentLoader, self._loader)
        nn = self.get_flat_filter_name(doc_type=doc_type.get_doc_type_str())
        self._logger.debug(f"Flat filter Name: {nn}")
        # do not set Hidden=True property here.
        # there is a strange error that pops up conditionally and it seems
        # to be remedied by not setting Hidden=True
        # see comments in tests.text_xml.test_in_filters.test_transform_clubs()
        return self.open_doc(fnm, loader, mProps.Props.make_props(FilterName=nn))

    # endregion open_flat_doc()

    # region open_doc()
    @overload
    def open_doc(self, fnm: PathOrStr) -> XComponent: ...

    @overload
    def open_doc(self, fnm: PathOrStr, loader: XComponentLoader) -> XComponent: ...

    @overload
    def open_doc(self, fnm: PathOrStr, *, props: Iterable[PropertyValue]) -> XComponent: ...

    @overload
    def open_doc(self, fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent: ...

    def open_doc(
        self,
        fnm: PathOrStr,
        loader: Optional[XComponentLoader] = None,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        # sourcery skip: extract-method, raise-specific-error
        # Props and FileIO are called this method so trigger global_reset first.
        if loader is None:
            loader = cast(XComponentLoader, self._loader)
        cargs = CancelEventArgs(self.open_doc.__qualname__)
        cargs.event_data = {
            "fnm": fnm,
            "loader": loader,
            "props": props,
        }
        eargs = EventArgs.from_args(cargs)
        self.on_reset(eargs)
        self.on_doc_opening(cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        fnm = cast("PathOrStr", cargs.event_data["fnm"])

        if fnm is None:
            raise Exception("Filename is null")
        pth = mFileIO.FileIO.get_absolute_path(fnm)

        if props is None:
            internal_props = mProps.Props.make_props(Hidden=True)
        else:
            internal_props = tuple(props)
        open_file_url = None
        if mFileIO.FileIO.is_openable(pth):
            self._logger.debug(f"Opening {pth}")
            open_file_url = mFileIO.FileIO.fnm_to_url(pth)

        elif self.is_url(pth):
            self._logger.debug(f"Will treat filename as a URL: '{pth}'")
            open_file_url = pth
        else:
            raise Exception(f"Unable to get url from file: {pth}")
        try:
            doc = loader.loadComponentFromURL(str(open_file_url), "_blank", 0, internal_props)  # type: ignore
            if doc is None:
                raise mEx.NoneError("loadComponentFromURL() returned None")
            self._reset_for_doc(doc)
            self.on_doc_opened(eargs)
            return doc
        except Exception as e:
            self._logger.exception(f"Unable to open document: {open_file_url}")
            raise Exception("Unable to open the document") from e

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    def open_readonly_doc(self, fnm: PathOrStr) -> XComponent: ...

    @overload
    def open_readonly_doc(self, fnm: PathOrStr, loader: XComponentLoader) -> XComponent: ...

    def open_readonly_doc(self, fnm: PathOrStr, loader: Optional[XComponentLoader] = None) -> XComponent:
        if loader is None:
            loader = cast(XComponentLoader, self._loader)
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
        if not e:
            self._logger.debug("Empty string: Using writer")
            return LoDocTypeStr.WRITER
        if e == "odb":
            return LoDocTypeStr.BASE
        elif e == "odf":
            return LoDocTypeStr.MATH
        elif e == "odg":
            return LoDocTypeStr.DRAW
        elif e == "odp":
            return LoDocTypeStr.IMPRESS
        elif e == "ods":
            return LoDocTypeStr.CALC
        elif e == "odt":
            return LoDocTypeStr.WRITER
        else:
            self._logger.debug(f"Do not recognize extension '{ext}'; using writer")
            return LoDocTypeStr.WRITER

    def doc_type_str(self, doc_type_val: LoDocType) -> LoDocTypeStr:
        return doc_type_val.get_doc_type_str()

    # region create_doc()
    @overload
    def create_doc(self, doc_type: LoDocTypeStr) -> XComponent: ...

    @overload
    def create_doc(self, doc_type: LoDocTypeStr, loader: XComponentLoader) -> XComponent: ...

    @overload
    def create_doc(self, doc_type: LoDocTypeStr, *, props: Iterable[PropertyValue]) -> XComponent: ...

    @overload
    def create_doc(
        self, doc_type: LoDocTypeStr, loader: XComponentLoader, props: Iterable[PropertyValue]
    ) -> XComponent: ...

    def create_doc(
        self,
        doc_type: LoDocTypeStr,
        loader: Optional[XComponentLoader] = None,
        props: Optional[Iterable[PropertyValue]] = None,
    ) -> XComponent:
        # sourcery skip: raise-specific-error
        # Props is called in this method so trigger global_reset first
        if loader is None:
            loader = cast(XComponentLoader, self._loader)
        cargs = CancelEventArgs(self.create_doc.__qualname__)
        cargs.event_data = {
            "doc_type": doc_type,
            "loader": loader,
            "props": props,
        }
        eargs = EventArgs.from_args(cargs)
        self.on_reset(eargs)
        self.on_doc_creating(cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        dtype = LoDocTypeStr(cargs.event_data["doc_type"])
        properties = cast("Iterable[PropertyValue]", cargs.event_data["props"])
        if properties is None:
            local_props = mProps.Props.make_props(Hidden=True)
        else:
            local_props = tuple(properties)
        self._logger.debug(f"Creating Office document {dtype}")
        try:
            doc = loader.loadComponentFromURL(f"private:factory/{dtype}", "_blank", 0, local_props)  # type: ignore
            self._reset_for_doc(doc)
            self.on_doc_created(eargs)
            return doc
        except Exception as e:
            self._logger.exception(f"Could not create document: {dtype}")
            raise Exception("Could not create a document") from e

    # endregion create_doc()

    # region create_macro_doc()
    @overload
    def create_macro_doc(self, doc_type: LoDocTypeStr) -> XComponent: ...

    @overload
    def create_macro_doc(self, doc_type: LoDocTypeStr, loader: XComponentLoader) -> XComponent: ...

    def create_macro_doc(self, doc_type: LoDocTypeStr, loader: Optional[XComponentLoader] = None) -> XComponent:
        if loader is None:
            loader = cast(XComponentLoader, self._loader)
        return self.create_doc(
            doc_type=doc_type,
            loader=loader,
            props=mProps.Props.make_props(Hidden=False, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE),
        )

    # endregion create_macro_doc()

    # region create_doc_from_template()

    @overload
    def create_doc_from_template(self, template_path: PathOrStr) -> XComponent: ...

    @overload
    def create_doc_from_template(self, template_path: PathOrStr, loader: XComponentLoader) -> XComponent: ...

    def create_doc_from_template(
        self, template_path: PathOrStr, loader: Optional[XComponentLoader] = None
    ) -> XComponent:
        # sourcery skip: raise-specific-error
        if loader is None:
            loader = cast(XComponentLoader, self._loader)
        cargs = CancelEventArgs(self.create_doc_from_template.__qualname__)
        self.on_doc_creating(cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        if not mFileIO.FileIO.is_openable(template_path):
            raise Exception(f"Template file can not be opened: '{template_path}'")

        self._logger.debug(f"Opening template: '{template_path}'")
        template_url = mFileIO.FileIO.fnm_to_url(fnm=template_path)

        props = mProps.Props.make_props(Hidden=True, AsTemplate=True)
        try:
            doc = loader.loadComponentFromURL(template_url, "_blank", 0, props)  # type: ignore
            self._reset_for_doc(doc)
            self.on_doc_created(EventArgs.from_args(cargs))
            return doc
        except Exception as e:
            self._logger.exception(f"Could not create document from template: '{template_path}'")
            raise Exception("Could not create document from template") from e

    # endregion create_doc_from_template()

    # ======================== document saving ==============

    def save(self, doc: object) -> bool:
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(self.save.__qualname__)
        cargs.event_data = {"doc": doc}
        self.on_doc_saving(cargs)
        if cargs.cancel:
            return False

        store = self.qi(XStorable, doc, True)
        try:
            store.store()
            self._logger.debug("Saved the document by overwriting")
        except IOException as e:
            self._logger.exception("Could not save the document")
            raise Exception("Could not save the document") from e

        self.on_doc_saved(EventArgs.from_args(cargs))
        return True

    # region    save_doc()

    @overload
    def save_doc(self, doc: object, fnm: PathOrStr) -> bool: ...

    @overload
    def save_doc(self, doc: object, fnm: PathOrStr, password: str) -> bool: ...

    @overload
    def save_doc(self, doc: object, fnm: PathOrStr, password: str, format: str) -> bool:  # pylint: disable=W0622
        ...

    # pylint: disable=W0622
    def save_doc(self, doc: object, fnm: PathOrStr, password: str | None = None, format: str | None = None) -> bool:
        # sourcery skip: avoid-builtin-shadow
        cargs = CancelEventArgs(self.save_doc.__qualname__)
        cargs.event_data = {
            "doc": doc,
            "fnm": fnm,
            "password": password,
            "format": format,
        }

        fnm = cast("PathOrStr", cargs.event_data["fnm"])
        password = cast(str, cargs.event_data["password"])
        format = cast(str, cargs.event_data["format"])

        self.on_doc_saving(cargs)
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
            self.on_doc_saved(EventArgs.from_args(cargs))
        return result

    # endregion save_doc()

    # region    store_doc()

    @overload
    def store_doc(self, store: XStorable, doc_type: LoDocType, fnm: PathOrStr) -> bool: ...

    @overload
    def store_doc(self, store: XStorable, doc_type: LoDocType, fnm: PathOrStr, password: str) -> bool: ...

    def store_doc(self, store: XStorable, doc_type: LoDocType, fnm: PathOrStr, password: Optional[str] = None) -> bool:
        cargs = CancelEventArgs(self.store_doc.__qualname__)
        cargs.event_data = {
            "store": store,
            "doc_type": doc_type,
            "fnm": fnm,
            "password": password,
        }
        self.on_doc_storing(cargs)
        if cargs.cancel:
            return False
        ext = mInfo.Info.get_ext(fnm)
        txt_format = "Text"
        if ext is None:
            self._logger.debug("Assuming a text format")
        else:
            txt_format = self.ext_to_format(ext=ext, doc_type=doc_type)
        if password is None:
            self.store_doc_format(store=store, fnm=fnm, format=txt_format)
        else:
            self.store_doc_format(store=store, fnm=fnm, format=txt_format, password=password)
        self.on_doc_stored(EventArgs.from_args(cargs))
        return True

    # endregion  store_doc()

    @overload
    def ext_to_format(self, ext: str) -> str: ...

    @overload
    def ext_to_format(self, ext: str, doc_type: LoDocType) -> str: ...

    def ext_to_format(self, ext: str, doc_type: LoDocType = LoDocType.UNKNOWN) -> str:
        # sourcery skip: collection-into-set, low-code-quality, merge-comparisons, merge-duplicate-blocks, remove-redundant-if
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
            self._logger.debug(f"Do not recognize extension '{ext}'; using text")
            return "Text"

    # region    store_doc_format()

    @overload
    def store_doc_format(self, store: XStorable, fnm: PathOrStr, format: str) -> bool:  # pylint: disable=W0622
        ...

    @overload
    def store_doc_format(self, store: XStorable, fnm: PathOrStr, format: str, password: str) -> bool: ...

    def store_doc_format(
        self, store: XStorable, fnm: PathOrStr, format: str, password: str | None = None
    ) -> bool:  # pylint: disable=W0622
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(self.store_doc_format.__qualname__)
        cargs.event_data = {
            "store": store,
            "format": format,
            "fnm": fnm,
            "password": password,
        }
        self.on_doc_storing(cargs)
        if cargs.cancel:
            return False
        pth = mFileIO.FileIO.get_absolute_path(cast("PathOrStr", cargs.event_data["fnm"]))
        fmt = str(cargs.event_data["format"])
        self._logger.debug(f"Saving the document in '{pth}'")
        self._logger.debug(f"Using format {fmt}")

        try:
            save_file_url = mFileIO.FileIO.fnm_to_url(pth)
            if password is None:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=fmt)
            else:
                store_props = mProps.Props.make_props(Overwrite=True, FilterName=fmt, Password=password)
            store.storeToURL(save_file_url, store_props)  # type: ignore
        except IOException as e:
            self._logger.exception(f"Could not save '{pth}'")
            raise Exception(f"Could not save '{pth}'") from e
        self.on_doc_stored(EventArgs.from_args(cargs))
        return True

    # endregion store_doc_format()

    # ======================== document closing ==============

    @overload
    def close(self, closeable: XCloseable) -> bool: ...

    @overload
    def close(self, closeable: XCloseable, deliver_ownership: bool) -> bool: ...

    def close(self, closeable: XCloseable, deliver_ownership=False) -> bool:
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(self.close.__qualname__)
        cargs.event_data = deliver_ownership
        self.on_doc_closing(cargs)
        if cargs.cancel:
            return False
        if closeable is None:
            return False
        self._logger.debug("Closing the document")
        try:
            closeable.close(cargs.event_data)
            self.on_doc_closed(EventArgs.from_args(cargs))
            return True
        except CloseVetoException as e:
            self._logger.error("close() Close was vetoed")
            raise Exception("Close was vetoed") from e

    # region close_doc()
    @overload
    def close_doc(self, doc: Any) -> None: ...

    @overload
    def close_doc(self, doc: Any, deliver_ownership: bool) -> None: ...

    def close_doc(self, doc: Any, deliver_ownership=False) -> None:
        # sourcery skip: raise-specific-error
        if self._disposed:
            return
        try:
            closeable = self.qi(XCloseable, doc, True)
            self.close(closeable=closeable, deliver_ownership=deliver_ownership)
        except DisposedException as e:
            self._logger.error("close_doc() Document close failed since Office link disposed")
            raise Exception("Document close failed since Office link disposed") from e

    # endregion close_doc()

    # ================= initialization via Addon-supplied context ====================

    def addon_initialize(self, addon_xcc: XComponentContext) -> XComponent:
        # this should most likely be put into a separate addon class
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(self.addon_initialize.__qualname__)
        cargs.event_data = {"addon_xcc": addon_xcc}
        eargs = EventArgs.from_args(cargs)
        self.on_reset(eargs)
        self.on_doc_opening(cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)

        xcc = addon_xcc
        if xcc is None:
            raise TypeError("'addon_xcc' is null. Could not access component context")
        mc_factory = xcc.getServiceManager()
        if mc_factory is None:
            raise Exception("Office Service Manager is unavailable")

        try:
            xdesktop = self.qi(XDesktop, mc_factory.createInstanceWithContext("com.sun.star.frame.Desktop", xcc), True)
        except Exception as e:
            raise Exception("Could not access desktop") from e
        doc = xdesktop.getCurrentComponent()
        if doc is None:
            raise Exception("Could not access document")
        # self._ms_factory = self.qi(XMultiServiceFactory, doc, True)
        self.on_doc_opened(eargs)
        return doc

    # ============= initialization via script context ======================

    def script_initialize(self, sc: XScriptContext) -> XComponent:
        # sourcery skip: raise-specific-error
        cargs = CancelEventArgs(self.script_initialize.__qualname__)
        cargs.event_data = {"sc": sc}
        eargs = EventArgs.from_args(cargs)
        self.on_reset(eargs)
        self.on_doc_opening(cargs)
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
        # self._ms_factory = self.qi(XMultiServiceFactory, doc, True)
        self.on_doc_opened(eargs)
        return doc

    # ==================== dispatch ===============================
    # see https://wiki.documentfoundation.org/Development/DispatchCommands

    def get_supported_dispatch_prefixes(self) -> Tuple[str, ...]:
        """
        Gets the prefixes are are supported by the local ``dispatch_cmd()`` method.

        Returns:
            Tuple[str, ...]: Tuple of supported dispatch prefixes.

        .. versionadded:: 0.40.0
        """
        return (
            ".uno:",
            "vnd.sun.star.",
            "service:",
        )

    # region dispatch_cmd()

    # @overload
    # def dispatch_cmd(self, cmd: str) -> Any: ...

    # @overload
    # def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue]) -> Any: ...

    # @overload
    # def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> Any: ...

    # @overload
    # def dispatch_cmd(self, cmd: str, *, frame: XFrame) -> Any: ...

    def dispatch_cmd(
        self,
        cmd: str,
        props: Iterable[PropertyValue] | None = None,
        frame: XFrame | None = None,
        in_thread: bool = False,
    ) -> Any:
        """
        Dispatches a LibreOffice command.

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not need to start with ``.uno:`` prefix.
            props (PropertyValue, optional): properties for dispatch.
            frame (XFrame, optional): Frame to dispatch to.
            in_thread (bool, optional): If ``True`` then dispatch is done in a separate thread.

        Raises:
            CancelEventError: If Dispatching is canceled via event.
            DispatchError: If any other error occurs.

        Returns:
            Any: A possible result of the executed internal dispatch. The information behind this any depends on the dispatch!

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DISPATCHING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DISPATCHED` :eventref:`src-docs-event`

        Note:
            There are many dispatch command constants that can be found in :ref:`utils_dispatch` Namespace

            | ``DISPATCHING`` Event args data contains any properties passed in via ``props``.
            | ``DISPATCHED`` Event args data contains any results from the dispatch commands.

        See Also:
            - :ref:`ch04_dispatching`
            - `LibreOffice Dispatch Commands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_

        .. versionchanged:: 0.40.0
            Now supports ``in_thread`` parameter.
        """
        # sourcery skip: assign-if-exp, extract-method, remove-unnecessary-cast
        if not cmd:
            raise mEx.DispatchError("cmd must not be empty or None")
        try:
            str_cmd = str(cmd).replace(".uno:", "")  # make sure and enum or other lookup did not get passed by mistake
            cargs = DispatchCancelArgs(self.dispatch_cmd.__qualname__, str_cmd)
            cargs.event_data = props
            self.on_dispatching(cargs)
            if cargs.cancel:
                raise mEx.CancelEventError(cargs, f'Dispatch Command "{str_cmd}" has been canceled')
            props = cargs.event_data
            if props is None:
                dispatch_props = ()
            else:
                dispatch_props = tuple(props)
            if frame is None:
                frame = self.get_frame()

            helper = self.create_instance_mcf(XDispatchHelper, "com.sun.star.frame.DispatchHelper")
            if helper is None:
                raise mEx.MissingInterfaceError(
                    XDispatchHelper, f"Could not create dispatch helper for command {str_cmd}"
                )
            provider = self.qi(XDispatchProvider, frame, True)

            prefix_key = "dispatch_prefixes"
            if prefix_key in self._cache:
                prefixes = cast(Tuple[str, ...], self._cache[prefix_key])
            else:
                set_pre = set(self.get_supported_dispatch_prefixes())
                set_pre.remove(".uno:")
                prefixes = tuple(set_pre)
                self._cache[prefix_key] = prefixes

            supported_prefix = False
            if str_cmd.startswith(prefixes):
                supported_prefix = True
            if supported_prefix:
                dispatch_cmd = str_cmd
            else:
                dispatch_cmd = f".uno:{str_cmd}"
            if in_thread:

                def worker(callback, cancel_args, helper, provider, cmd, props) -> None:
                    result = helper.executeDispatch(provider, cmd, "", 0, props)
                    callback(result, cancel_args)

                def callback(result: Any, cancel_args: DispatchCancelArgs) -> None:
                    # runs after the threaded dispatch is finished
                    eargs = DispatchArgs.from_args(cancel_args)
                    eargs.event_data = result
                    self.on_dispatched(eargs)
                    if self._logger.debug:
                        self._logger.debug(f"Finished Dispatched in thread: {str_cmd}")

                if self._logger.is_debug:
                    self._logger.debug(f"Dispatching in thread: {dispatch_cmd}")
                # t = threading.Thread(
                #     target=helper.executeDispatch, args=(provider, dispatch_cmd, "", 0, dispatch_props)
                # )
                t = threading.Thread(
                    target=worker, args=(callback, cargs, helper, provider, dispatch_cmd, dispatch_props)
                )
                t.start()
                return None
            else:
                if self._logger.is_debug:
                    self._logger.debug(f"Dispatching in main thread: {dispatch_cmd}")
                result = helper.executeDispatch(provider, dispatch_cmd, "", 0, dispatch_props)
            eargs = DispatchArgs.from_args(cargs)
            eargs.event_data = result
            self.on_dispatched(eargs)
            return result
        except mEx.CancelEventError:
            self._logger.info(f'Dispatch Command "{str_cmd}" has been canceled')
            raise
        except Exception as e:
            self._logger.error(f'Error dispatching "{cmd}"', exc_info=True)
            raise mEx.DispatchError(f'Error dispatching "{cmd}"') from e

    # endregion dispatch_cmd()

    # ================= Uno Commands =========================

    def make_uno_cmd(self, item_name: str) -> str:
        return f"vnd.sun.star.script:Foo/Foo.{item_name}?language=Java&location=share"

    def extract_item_name(self, uno_cmd: str) -> str:
        try:
            foo_pos = uno_cmd.index("Foo.")
        except ValueError as e:
            self._logger.exception(f"Could not find Foo header in command: '{uno_cmd}'")
            raise ValueError(f"Could not find Foo header in command: '{uno_cmd}'") from e
        try:
            lang_pos = uno_cmd.index("?language")
        except ValueError as exc:
            self._logger.exception(f"Could not find language header in command: '{uno_cmd}'")
            raise ValueError(f"Could not find language header in command: '{uno_cmd}'") from exc
        start = foo_pos + 4
        return uno_cmd[start:lang_pos]

    # ======================== use Inspector extensions ====================

    def inspect(self, obj: object) -> None:
        if self._xcc is None or self._mc_factory is None:
            self._logger.debug("No office connection found")
            return
        try:
            ts = mInfo.Info.get_interface_types(obj)
            title = "Object"
            if ts is not None and len(ts) > 0:
                title = f"{ts[0].getTypeName()} {title}"  # type: ignore
            inspector = self._mc_factory.createInstanceWithContext("org.openoffice.InstanceInspector", self._xcc)
            #       hands on second use
            if inspector is None:
                self._logger.debug("Inspector Service could not be instantiated")
                return
            self._logger.debug("Inspector Service instantiated")
            intro = self.create_instance_mcf(XIntrospection, "com.sun.star.beans.Introspection", raise_err=True)
            intro_acc = intro.inspect(inspector)
            method = intro_acc.getMethod("inspect", -1)
            self._logger.debug(f"inspect() method was found: {method is not None}")
            params = [[obj, title]]
            method.invoke(inspector, params)
        except Exception as e:
            self._logger.exception("Could not access Inspector")

    def mri_inspect(self, obj: object) -> None:
        # sourcery skip: raise-specific-error
        # Available from http://extensions.libreoffice.org/extension-center/mri-uno-object-inspection-tool
        #               or http://extensions.services.openoffice.org/en/project/MRI
        #  Docs: https://github.com/hanya/MRI/wiki
        #  Forum tutorial: https://forum.openoffice.org/en/forum/viewtopic.php?f=74&t=49294
        xi = self.create_instance_mcf(XIntrospection, "mytools.Mri")
        if xi is None:
            self._logger.error("MRI Inspector Service could not be instantiated")
            raise Exception("MRI Inspector Service could not be instantiated")
        self._logger.debug("MRI Inspector Service instantiated")
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
            self._logger.debug("Lo.delay(): Ms must be greater then zero")
            return
        sec = ms / 1000
        time.sleep(sec)

    wait = delay

    @staticmethod
    def is_none_or_empty(s: str) -> bool:
        return s is None or not s

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
            self._logger.debug(f"{s} could not be parsed as an int; using 0")
        return 0

    @overload
    @staticmethod
    def print_names(names: Sequence[str]) -> None: ...

    @overload
    @staticmethod
    def print_names(names: Sequence[str], num_per_line: int) -> None: ...

    @staticmethod
    def print_names(names: Sequence[str], num_per_line: int = 4) -> None:
        if not names:
            print("  No names found")
            return
        col_count = max(num_per_line, 1)

        lst_2d = mThelper.TableHelper.convert_1d_to_2d(
            seq_obj=sorted(names, key=str.casefold), col_count=col_count, empty_cell_val=""
        )
        longest = mThelper.TableHelper.get_largest_str(names)
        fmt_len = longest + 1
        format_opt = FormatterTable(format=f"<{fmt_len}") if longest > 0 else None
        indent = "  "
        print(f"No. of names: {len(names)}")
        if format_opt:
            actual_count = len(lst_2d[0])  # pylint: disable=unsubscriptable-object
            if actual_count > 1:
                # if this is more then on colum then print header
                #  -----------|-----------|-----------
                print(f"{indent}", end="")
                for i, _ in enumerate(range(actual_count)):
                    print("-" * fmt_len, end="")
                    if i < actual_count - 1:
                        print("|-", end="")
                print()
            for i, row in enumerate(lst_2d):
                col_str = format_opt.get_formatted(idx_row=i, row_data=row, join_str="| ")
                print(f"{indent}{col_str}")
        else:
            for row in lst_2d:  # pylint: disable=not-an-iterable
                for col in row:
                    print(f'{indent}"{col}"', end="")
                print()
        print("\n\n")

    # ------------------- container manipulation --------------------
    # region print_table()
    @overload
    @staticmethod
    def print_table(name: str, table: Table) -> None: ...

    @overload
    @staticmethod
    def print_table(name: str, table: Table, format_opt: FormatterTable) -> None: ...

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
            self._logger.debug("Container is null")
            return None
        num_el = con.getCount()
        if num_el == 0:
            self._logger.debug("No elements in the container")
            return None

        names_list = []
        for i in range(num_el):
            named = self.qi(XNamed, con.getByIndex(i), True)
            names_list.append(named.getName())

        if not names_list:
            self._logger.debug("No element names found in the container")
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
                self._logger.debug(f"Could not access element {i}")
        self._logger.debug(f"Could not find a '{nm}' property set in the container")
        return None

    def is_uno_interfaces(self, component: Any, *args: str | UnoInterface) -> bool:
        if not args:
            return False
        result = True
        for arg in args:
            try:
                t = uno.getClass(arg) if isinstance(arg, str) else arg
                obj = self.qi(t, component)  # type: ignore
                if obj is None:
                    result = False
                    break
            except Exception:
                result = False
                break
        return result

    def get_frame(self) -> XFrame:
        model = self.get_model()
        controller = model.getCurrentController()
        return controller.getFrame()

    def get_model(self) -> XModel:
        if self._opt.dynamic:
            return self.XSCRIPTCONTEXT.getDocument()
        component = self.desktop.get_current_component()
        model = self.qi(XModel, component, True)
        return model

    def lock_controllers(self) -> bool:
        # much faster updates as screen is basically suspended
        cargs = CancelEventArgs(self.lock_controllers.__qualname__)
        self.on_controllers_locking(cargs)
        if cargs.cancel:
            return False
        # doc = self.xscript_context.getDocument() if self._opt.dynamic else self.lo_component
        xmodel = self.get_model()
        xmodel.lockControllers()
        self.on_controllers_locked(EventArgs(self))
        return True

    def unlock_controllers(self) -> bool:
        cargs = CancelEventArgs(self.unlock_controllers.__qualname__)
        self.on_controllers_unlocking(cargs)
        if cargs.cancel:
            return False
        # doc = self.xscript_context.getDocument() if self._opt.dynamic else self.lo_component
        xmodel = self.get_model()
        if xmodel.hasControllersLocked():
            xmodel.unlockControllers()
            self.on_controllers_unlocked(EventArgs(self))
            return True
        return False

    def has_controllers_locked(self) -> bool:
        xmodel = self.get_model()
        return xmodel.hasControllersLocked()

    def print(self, *args, **kwargs) -> None:
        if not self._allow_print:
            return
        cargs = CancelEventArgs(self.print.__qualname__)
        self.on_printing(cargs)
        if cargs.cancel:
            return
        print(*args, **kwargs)

    # region XML
    def get_flat_filter_name(self, doc_type: LoDocTypeStr) -> str:
        """
        Gets the Flat XML filter name for the doc type.

        Args:
            doc_type (DocTypeStr): Document type.

        Returns:
            str: Flat XML filter name.

        .. versionadded:: 0.12.0
        """
        if doc_type == LoDocTypeStr.WRITER:
            return "OpenDocument Text Flat XML"
        elif doc_type == LoDocTypeStr.CALC:
            return "OpenDocument Spreadsheet Flat XML"
        elif doc_type == LoDocTypeStr.DRAW:
            return "OpenDocument Drawing Flat XML"
        elif doc_type == LoDocTypeStr.IMPRESS:
            return "OpenDocument Presentation Flat XML"
        else:
            self._logger.debug("No Flat XML filter for this document type; using Flat text")
            return "OpenDocument Text Flat XML"

    # endregion XML

    @property
    def null_date(self) -> datetime:
        # https://tinyurl.com/2pdrt5z9#NullDate
        try:
            return self._null_date
        except AttributeError:
            self._null_date = datetime(year=1889, month=12, day=30, tzinfo=timezone.utc)
            if self._xdesktop is None:
                return self._null_date
            doc = self.desktop.get_current_component()
            if doc is None:
                return self._null_date
            n_supplier = self.qi(XNumberFormatsSupplier, doc)
            if n_supplier is None:
                # this is not always a XNumberFormatsSupplier such as *.odp documents
                return self._null_date
            number_settings = n_supplier.getNumberFormatSettings()
            d = cast("UnoDate", number_settings.getPropertyValue("NullDate"))
            self._null_date = datetime(d.Year, d.Month, d.Day, tzinfo=timezone.utc)
        return self._null_date

    @property
    def is_loaded(self) -> bool:
        return self._lo_inst is not None

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
        return self.desktop.component  # type: ignore

    StarDesktop, stardesktop = star_desktop, star_desktop

    @property
    def lo_component(self) -> XComponent | None:
        """
        Gets the internal component.

        Unlike the :py:attr:`this_component` property this property will not autoload
        and it will not observe the dynamic option for this instance.

        This property will always return the current internal component document.

        In most cases the :py:attr:`this_component` property should be used instead of this property.

        Returns:
            XComponent | None: Component or None if not loaded.
        """
        if self._xdesktop is None:
            return None
        return self._xdesktop.get_current_component()

    @property
    def this_component(self) -> XComponent | None:
        """
        This component is similar to the ThisComponent in Basic.

        It is functionally the same as ``XSCRIPTCONTEXT.getDesktop().getCurrentComponent()``.

        When running in a macro this property can be access directly to get the current document.
        When not in a macro then load_office() must be called first

        This property differs from :py:attr:`lo_component` in the following ways.

        1. It will autoload if called in a macro.
        2. It will observe the dynamic option for this instance.

        When this class options are set to dynamic then this property will always return the current document
        from internal XSCRIPTCONTEXT; otherwise, it will return the current internal document.
        In most cases this property should be used instead of :py:attr:`lo_component`.
        Also in most cases this property will return the same component :py:attr:`lo_component`.

        Returns:
            XComponent | None: Component or None if not loaded.
        """
        if mock_g.DOCS_BUILDING:
            self._this_component = None
            return self._this_component
        if self.is_loaded is False:
            # attempt to connect direct
            # failure will result in script error and then exit
            self._logger.debug("this_component: Office not loaded. Calling load_office()")
            self.load_office()

        # comp = self.star_desktop.getCurrentComponent()
        if self.desktop is None:
            self._logger.debug("this_component: Could not access desktop. Returning NOne")
            return None
        desktop = self.desktop.component
        doc = desktop.getCurrentComponent()
        # if not self._opt.dynamic and self._doc is None or self._opt.dynamic:
        #     doc = desktop.getCurrentComponent()
        # else:
        #     doc = self._doc
        if doc is None:
            self._logger.debug("this_component: Could not access current document. Returning None")
            return None
        service_info = self.qi(XServiceInfo, doc, True)
        impl = service_info.getImplementationName()
        # com.sun.star.comp.sfx2.BackingComp is the Main Launcher App.
        if impl in ("com.sun.star.comp.basic.BasicIDE", "com.sun.star.comp.sfx2.BackingComp"):
            self._logger.debug("this_component: Basic IDE or Welcome screen. Returning None")
            return None  # None when Basic IDE or welcome screen
        return doc

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

            self._xscript_context = script_context.ScriptContext(ctx=ctx)
        return self._xscript_context

    XSCRIPTCONTEXT = xscript_context

    @property
    def bridge(self) -> XComponent | None:
        try:
            return self._bridge_component
        except AttributeError:
            try:
                # when running as macro self._lo_inst will not have bridge_component
                self._bridge_component = self._lo_inst.bridge_component  # type: ignore
            except AttributeError:
                self._bridge_component = None
            return self._bridge_component

    @property
    def loader_current(self) -> XComponentLoader:
        return self._loader  # type: ignore

    @property
    def bridge_connector(self) -> ConnectBase:
        return self._lo_inst  # type: ignore

    @property
    def options(self) -> LoOptions:
        return self._opt

    @property
    def events(self) -> EventObserver:
        # mangled name.
        # pylint: disable=no-member
        return self._EventsPartial__events  # type: ignore

    @property
    def lo_loader(self) -> LoLoader:
        """Get/Sets the loader for this instance"""
        if self._lo_loader is None:
            self._logger.error("lo_loader: Office not loaded")
            raise mEx.LoadingError("Office not loaded")
        return self._lo_loader

    @lo_loader.setter
    def lo_loader(self, value: LoLoader) -> None:
        if self._is_default:
            self._logger.error("lo_loader: Cannot set loader for default instance")
            raise mEx.LoadingError("Cannot set loader for default instance")
        if value is self._lo_loader:
            return
        _ = self.load_from_lo_loader(value)

    @property
    def is_default(self) -> bool:
        """Gets if the current instance is the default Lo instance."""
        return self._is_default

    @property
    def current_doc(self) -> OfficeDocumentT | None:
        """
        Get/Sets the current document.

        If there is no current document then an attempt is made to get the current document from the last active document from the desktop components.

        Note:
            This property does not require the use of the :py:class:`~ooodev.macro.MacroLoader` in macros.

            It is almost never necessary to set this property directly. When there are changes to the environment, such as when a new document gets focus,
            this class will automatically pick up the changes and update the current document.
            However, there may be exceptions this this such as when extension has a OnFocus job.
            In cases such as that the OnFocus may need to set the current document.


        .. versionchanged:: 0.47.4
            It is now possible to set the current document.

        .. versionchanged:: 0.45.5
            This property will now return the latest document from the desktop components if the current document is None.
        """
        if self._current_doc is None:
            self._logger.debug(
                "current_doc: Current document is None. Attempting to get via desktop.get_current_component()"
            )
            doc = self.desktop.get_current_component()
            if doc is None:
                self._logger.debug(
                    "current_doc: Could not access current document. Attempting to get from desktop current desktop.components"
                )
                for comp in self.desktop.components:
                    # if there is more then on component then the first match is used.
                    # It seems the last opened document is the first in the list.
                    self._logger.debug("current_doc: found a component to use.")
                    doc = comp
                    if self._logger.is_debug:
                        if hasattr(doc, "getImplementationName"):
                            self._logger.debug(f"current_doc: Component: {doc.getImplementationName()}")
                        if hasattr(doc, "getURL"):
                            self._logger.debug(f"current_doc: Component URL: {doc.getURL()}")
                        if hasattr(doc, "RuntimeUID"):
                            self._logger.debug(f"current_doc: Component RuntimeUID: {doc.RuntimeUID}")
                    break
            if doc is None:
                self._logger.debug("current_doc: Could not access current document. Returning None")
                return None  # type: ignore
            self._current_doc = doc_factory(doc=doc, lo_inst=self)
            # self._current_doc = doc_factory(doc=self.desktop.get_current_component(), lo_inst=self)
        return self._current_doc  # type: ignore

    @current_doc.setter
    def current_doc(self, value: OfficeDocumentT | XComponent) -> None:
        self._clear_cache()
        if hasattr(value, "DOC_TYPE"):
            self._current_doc = value
        else:
            try:
                self._current_doc = doc_factory(doc=value, lo_inst=self)
            except Exception:
                self._logger.exception("current_doc: Could not set current document")
                raise

    @property
    def desktop(self) -> TheDesktop:
        """Get the current desktop"""
        if self._xdesktop is None:
            self._logger.debug("desktop: Could not access desktop")
            if self.is_loaded is False:
                self._logger.debug("desktop: Office not loaded. Attempting to load office")
                # attempt to connect direct
                # failure will result in script error and then exit
                self.load_office()
        if self._xdesktop is None:
            self._logger.error("desktop: Could not access desktop")
            raise mEx.LoadingError("Could not access desktop")
        return self._xdesktop

    @property
    def global_event_broadcaster(self) -> TheGlobalEventBroadcaster:
        if self._glb_event_broadcaster is None:
            self._logger.debug("global_event_broadcaster:Global Event Broadcaster is None")
            if self.is_loaded is False:
                self._logger.debug("global_event_broadcaster: Office not loaded. Attempting to load office")
                # attempt to connect direct
                # failure will result in script error and then exit
                self.load_office()
        if self._glb_event_broadcaster is None:
            self._logger.error("global_event_broadcaster: Could not access the Global Event Broadcaster")
            raise mEx.LoadingError("Could not access the Global Event Broadcaster")
        return self._glb_event_broadcaster

    @property
    def app_font_pixel_ratio(self) -> GenericSizePos[float]:
        """
        Gets the ratio between App Font and Pixels.

        This is used to convert font sizes to pixels.
        This value will vary on different systems.

        Returns:
            GenericSizePos[float]: Ratios of how many pixels are in an app font.
        """
        if self._app_font_pixel_ratio is None:
            # 17 is AppFont
            try:
                self._app_font_pixel_ratio = self._get_font_ratio()(17)  # type: ignore
            except Exception:
                self._app_font_pixel_ratio = GenericSizePos(0.5, 0.5, 0.5, 0.5)  # best guess from window
        return self._app_font_pixel_ratio

    @property
    def sys_font_pixel_ratio(self) -> GenericSizePos[float]:
        """
        Gets the ratio between System Font and Pixels.

        This is used to convert font sizes to pixels.
        This value will vary on different systems.

        Returns:
            GenericSizePos[float]: Ratios of how many pixels are in an system font.
        """
        if self._sys_font_pixel_ratio is None:
            # 18 is SysFont
            try:
                self._sys_font_pixel_ratio = self._get_font_ratio()(18)  # type: ignore
            except Exception:
                self._sys_font_pixel_ratio = GenericSizePos(0.5, 0.5, 0.5, 0.5)  # best guess from windows
        return self._sys_font_pixel_ratio

    @property
    def cache(self) -> LRUCache:
        """
        Gets access to the a cache for the current instance.

        This is a Least Recently Used (LRU) Cache. This cache is also used within this framework.

        Returns:
            LRUCache: Cache instance.
        """
        return self._shared_cache

    @property
    def tmp_dir(self) -> Path:
        """
        Gets the LibreOffice temporary directory.

        Returns:
            Path: Temporary directory.

        .. versionadded:: 0.41.0
        """
        key = "tmp_dir"
        if key in self._cache:
            return self._cache[key]
        # pylint: disable=import-outside-toplevel
        from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp

        ps = ThePathSettingsComp.from_lo(lo_inst=self)
        tmp_dir = Path(uno.fileUrlToSystemPath(ps.temp[0]))
        self._cache[key] = tmp_dir
        return tmp_dir

    @property
    def version(self) -> Tuple[int, int, int]:
        """
        Gets the OooDev version.

        Returns:
            Tuple[int, int, int]: Version tuple.

        .. versionadded:: 0.45.0
        """
        key = "ooodev_version"
        if key in self._cache:
            return self._cache[key]
        s = ooodev.get_version()
        ver = s.split(".")
        ver_int = tuple([int(v) for v in ver])
        self._cache[key] = ver_int
        return ver_int  # type: ignore


__all__ = ("LoInst",)

if mock_g.FULL_IMPORT:
    from ooodev.adapter.awt.unit_conversion_comp import UnitConversionComp
    from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp
