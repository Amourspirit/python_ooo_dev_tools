# coding: utf-8
# Python conversion of Lo.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/

from __future__ import annotations
from datetime import datetime
import time
from typing import TYPE_CHECKING, Any, Iterable, Optional, List, Sequence, Tuple, overload, Type

import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.frame import XDesktop
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.util import XCloseable
from com.sun.star.frame import XComponentLoader
from com.sun.star.frame import XModel
from com.sun.star.frame import XStorable


# if not mock_g.DOCS_BUILDING:
# not importing for doc building just result in short import name for
# args that use these.
# this is also true becuase docs/conf.py ignores com import for autodoc
from ..mock import mock_g
from .inst.lo import lo_inst

# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from ..conn import cache as mCache
from ..conn import connectors
from ..conn.connect import LoBridgeCommon
from ..events.event_singleton import _Events
from ..events.lo_named_event import LoNamedEvent
from ..formatters.formatter_table import FormatterTable
from ..meta.static_meta import StaticProperty, classproperty
from .type_var import PathOrStr, UnoInterface, T, Table
from ooodev.utils.inst.lo.options import Options as LoOptions
from ooodev.utils.inst.lo.doc_type import DocType as LoDocType, DocTypeStr as LoDocTypeStr
from ooodev.utils.inst.lo.service import Service as LoService
from ooodev.utils.inst.lo.clsid import CLSID as LoClsid

from com.sun.star.lang import XComponent

if TYPE_CHECKING:
    try:
        from typing import Literal  # Py >= 3.8
    except ImportError:
        from typing_extensions import Literal
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


# PathOrStr = type_var.PathOrStr
# """Path like object or string"""


class Lo(metaclass=StaticProperty):
    """LibreOffice helper class"""

    Options = LoOptions

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
            opt: Lo.Options | None = None,
        ):
            """
            Create a connection to office

            Args:
                connector (connectors.ConnectPipe | connectors.ConnectSocket | None): Connection information. Ignore for macros.
                cache_obj (~ooodev.conn.cache.Cache | None, optional): Cache instance that determines if LibreOffice profile is to be copied and cached
                    Ignore for macros. Defaults to None.
                opt (~ooodev.utils.lo.Lo.Options, optional): Extra Load options.

            .. versionchanged:: 0.6.10

                Added ``opt`` parameter.
            """
            self.loader = Lo.load_office(connector=connector, cache_obj=cache_obj, opt=opt)

        def __enter__(self) -> XComponentLoader:
            return self.loader

        def __exit__(self, exc_type, exc_val, exc_tb):
            Lo.close_office()

    # region aliases
    DocType = LoDocType
    DocTypeStr = LoDocTypeStr
    Service = LoService
    CLSID = LoClsid

    # endregion aliases

    # region port connect to locally running Office via port 8100
    # endregion port

    ConnectPipe = connectors.ConnectPipe
    """Alias of connectors.ConnectPipe"""
    ConnectSocket = connectors.ConnectSocket
    """Alias of connectors.ConnectSocket"""

    _lo_inst: lo_inst.LoInst = None

    # region    qi()

    @overload
    @classmethod
    def qi(cls, atype: Type[T], obj: Any) -> T | None:
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
    @classmethod
    def qi(cls, atype: Type[T], obj: Any, raise_err: Literal[True]) -> T:
        """
        Generic method that get an interface instance from  an object.

        Args:
            atype (T): Interface type such as XInterface
            obj (object): Object that implements interface.
            raise_err (bool, optional): If True then raises MissingInterfaceError if result is None. Default False

        Raises:
            MissingInterfaceError: If 'raise_err' is 'True' and result is None

        Returns:
            T: instance of interface.
        """
        ...

    @overload
    @classmethod
    def qi(cls, atype: Type[T], obj: Any, raise_err: Literal[False]) -> T | None:
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

    @classmethod
    def qi(cls, atype: Type[T], obj: XTypeProvider, raise_err: bool = False) -> T | None:
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
                search = Lo.qi(XSearchable, cell_range)
                sd = search.createSearchDescriptor()
        """
        if raise_err:
            return cls._lo_inst.qi(atype, obj, raise_err)
        return cls._lo_inst.qi(atype, obj)

    # endregion qi()

    @classmethod
    def get_context(cls) -> XComponentContext:
        """
        Gets current LO Component Context
        """
        return cls._lo_inst.get_context()

    @classmethod
    def get_desktop(cls) -> XDesktop:
        """
        Gets current LO Desktop
        """
        return cls._lo_inst.get_desktop()

    @classmethod
    def get_component_factory(cls) -> XMultiComponentFactory:
        """Gets current multi component factory"""
        return cls._lo_inst.get_component_factory()

    @classmethod
    def get_service_factory(cls) -> XMultiServiceFactory:
        """Gets current multi service factory"""
        # return cls._bridge_component
        return cls._lo_inst.get_service_factory()

    # region interface object creation

    # region    create_instance_msf()
    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str) -> T | None:
        ...

    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str, msf: Any | None) -> T | None:
        ...

    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str, *, raise_err: Literal[True]) -> T:
        ...

    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str, *, raise_err: Literal[False]) -> T | None:
        ...

    @overload
    @classmethod
    def create_instance_msf(cls, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[True]) -> T:
        ...

    @overload
    @classmethod
    def create_instance_msf(
        cls, atype: Type[T], service_name: str, msf: Any | None, raise_err: Literal[False]
    ) -> T | None:
        ...

    @classmethod
    def create_instance_msf(
        cls, atype: Type[T], service_name: str, msf: XMultiServiceFactory | None = None, raise_err: bool = False
    ) -> T | None:
        """
        Creates an instance classified by the specified service name and
        optionally passes arguments to that instance.

        The interface specified by ``atype`` is returned from created instance.

        Args:
            atype (Type[T]): Type of interface to return from created service.
                Any Uno class that starts with 'X' such as XInterface
            service_name (str): Service name
            msf (XMultiServiceFactory, optional): Multi service factory used to create instance
            raise_err (bool, optional): If ``True`` then can raise CreateInstanceMsfError or MissingInterfaceError. Default is ``False``

        Raises:
            CreateInstanceMsfError: If ``raise_err`` is ``True`` and no instance was created
            MissingInterfaceError: If ``raise_err`` is ``True`` and instance was created but does not implement ``atype`` interface.
            Exception: if unable to create instance for any other reason

        Returns:
            T: Instance of interface for the service name or possibly ``None`` if ``raise_err`` is False.

        Note:
            When ``raise_err=True`` return value will never be ``None``.

        Example:
            In the following example ``src_con`` is an instance of ``XSheetCellRangeContainer``

            .. code-block:: python

                from com.sun.star.sheet import XSheetCellRangeContainer
                src_con = Lo.create_instance_msf(XSheetCellRangeContainer, "com.sun.star.sheet.SheetCellRanges")

        """
        return cls._lo_inst.create_instance_msf(atype, service_name, msf, raise_err)

    # endregion create_instance_msf()

    # region    create_instance_mcf()
    @overload
    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str) -> T | None:
        ...

    @overload
    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None) -> T | None:
        ...

    @overload
    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str, *, raise_err: Literal[True]) -> T:
        ...

    @overload
    @classmethod
    def create_instance_mcf(cls, atype: Type[T], service_name: str, *, raise_err: Literal[False]) -> T | None:
        ...

    @overload
    @classmethod
    def create_instance_mcf(
        cls, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None, raise_err: Literal[True]
    ) -> T:
        ...

    @overload
    @classmethod
    def create_instance_mcf(
        cls, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None, raise_err: Literal[False]
    ) -> T | None:
        ...

    @classmethod
    def create_instance_mcf(
        cls, atype: Type[T], service_name: str, args: Tuple[Any, ...] | None = None, raise_err: bool = False
    ) -> T | None:
        """
        Creates an instance of a component which supports the services specified by the factory,
        and optionally initializes the new instance with the given arguments and context.

        The interface specified by ``atype`` is returned from created instance.

        Args:
            atype (Type[T]): Type of interface to return from created instance.
                Any Uno class that starts with ``X`` such as ``XInterface``
            service_name (str): Service Name
            args (Tuple[Any, ...], Optional): Args
            raise_err (bool, optional): If ``True`` then can raise CreateInstanceMcfError or MissingInterfaceError. Default is ``False``

        Raises:
            CreateInstanceMcfError: If ``raise_err`` is ``True`` and no instance was created
            MissingInterfaceError: If ``raise_err`` is ``True`` and instance was created but does not implement ``atype`` interface.
            Exception: if unable to create instance for any other reason

        Note:
            When ``raise_err=True`` return value will never be ``None``.

        Returns:
            T | None: Instance of interface for the service name or possibly ``None`` if ``raise_err`` is False.

        Example:
            In the following example ``tk`` is an instance of ``XExtendedToolkit``

            .. code-block:: python

                from com.sun.star.awt import XExtendedToolkit
                tk = Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")

        """
        return cls._lo_inst.create_instance_mcf(atype, service_name, args, raise_err)

    # endregion create_instance_mcf()

    # endregion interface object creation

    @staticmethod
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

    @classmethod
    def load_office(
        cls,
        connector: connectors.ConnectPipe | connectors.ConnectSocket | None = None,
        cache_obj: mCache.Cache | None = None,
        opt: Lo.Options | None = None,
    ) -> XComponentLoader:
        """
        Loads Office

        Not available in a macro.

        If running outside of office then a bridge is created that connects to office.

        If running from inside of office e.g. in a macro, then ``Lo.XSCRIPTCONTEXT`` is used.
        ``using_pipes`` is ignored with running inside office.

        Args:
            connector (connectors.ConnectPipe, connectors.ConnectSocket, optional): Connection information. Ignore for macros.
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

        See Also:
            - :py:meth:`open_doc`
            - :py:class:`.Lo.Loader`
            - :ref:`ch02`

        Example:

            .. code::

                loader =  Lo.Loader(Lo.ConnectSocket()):
                doc = Write.create_doc(loader)
                ...

        .. versionchanged:: 0.6.10

                Added ``opt`` parameter.
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

        cls._lo_inst = lo_inst.LoInst(opt=opt, events=_Events())

        try:
            return cls._lo_inst.load_office(connector=connector, cache_obj=cache_obj)
        except Exception:
            raise SystemExit(1)

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
        return cls._lo_inst.close_office()

    @classmethod
    def kill_office(cls) -> None:
        """
        Kills the office connection.

        Returns:
            None:

        See Also:
            :py:meth:`~Lo.close_office`
        """
        cls._lo_inst.kill_office()

    # endregion office shutdown

    # region document opening

    # region open_flat_doc()
    @overload
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr, doc_type: Lo.DocType) -> XComponent:
        ...

    @overload
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr, doc_type: Lo.DocType, loader: XComponentLoader) -> XComponent:
        ...

    @classmethod
    def open_flat_doc(
        cls, fnm: PathOrStr, doc_type: Lo.DocType, loader: Optional[XComponentLoader] = None
    ) -> XComponent:
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
        return cls._lo_inst.open_flat_doc(fnm=fnm, doc_type=doc_type, loader=loader)

    # endregion open_flat_doc()

    # region open_doc()
    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr) -> XComponent:
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, *, props: Iterable[PropertyValue]) -> XComponent:
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        ...

    @classmethod
    def open_doc(
        cls,
        fnm: PathOrStr,
        loader: Optional[XComponentLoader] = None,
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
        return cls._lo_inst.open_doc(fnm=fnm, loader=loader, props=props)

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr) -> XComponent:
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        ...

    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: Optional[XComponentLoader] = None) -> XComponent:
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
        return cls._lo_inst.open_readonly_doc(fnm=fnm, loader=loader)

    # endregion open_readonly_doc()

    # ======================== document creation ==============

    @classmethod
    def ext_to_doc_type(cls, ext: str) -> LoDocTypeStr:
        """
        Gets document type from extension

        Args:
            ext (str): extension used for lookup

        Returns:
            DocTypeStr: DocTypeStr enum. If not match if found defaults to ``DocTypeStr.WRITER``

        See Also:
            :ref:`ch02_save_doc`
        """
        return cls._lo_inst.ext_to_doc_type(ext)

    @classmethod
    def doc_type_str(cls, doc_type_val: LoDocType) -> LoDocTypeStr:
        """
        Converts a doc type into a :py:class:`~Lo.DocTypeStr` representation.

        Args:
            doc_type_val (DocType): Doc type as int

        Returns:
            DocTypeStr: doc type as string.
        """
        return cls._lo_inst.doc_type_str(doc_type_val)

    # region create_doc()
    @overload
    @classmethod
    def create_doc(cls, doc_type: DocTypeStr) -> XComponent:
        ...

    @overload
    @classmethod
    def create_doc(cls, doc_type: DocTypeStr, loader: XComponentLoader) -> XComponent:
        ...

    @overload
    @classmethod
    def create_doc(cls, doc_type: DocTypeStr, *, props: Iterable[PropertyValue]) -> XComponent:
        ...

    @overload
    @classmethod
    def create_doc(cls, doc_type: DocTypeStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        ...

    @classmethod
    def create_doc(
        cls,
        doc_type: Lo.DocTypeStr,
        loader: Optional[XComponentLoader] = None,
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
        return cls._lo_inst.create_doc(doc_type=doc_type, loader=loader, props=props)

    # endregion create_doc()

    # region create_macro_doc()
    @overload
    @classmethod
    def create_macro_doc(cls, doc_type: Lo.DocTypeStr) -> XComponent:
        ...

    @overload
    @classmethod
    def create_macro_doc(cls, doc_type: Lo.DocTypeStr, loader: XComponentLoader) -> XComponent:
        ...

    @classmethod
    def create_macro_doc(cls, doc_type: Lo.DocTypeStr, loader: Optional[XComponentLoader] = None) -> XComponent:
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
        return cls._lo_inst.create_macro_doc(doc_type=doc_type, loader=loader)

    # endregion create_macro_doc()

    # region create_doc_from_template()

    @overload
    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr) -> XComponent:
        ...

    @overload
    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr, loader: XComponentLoader) -> XComponent:
        ...

    @classmethod
    def create_doc_from_template(
        cls, template_path: PathOrStr, loader: Optional[XComponentLoader] = None
    ) -> XComponent:
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
        return cls._lo_inst.create_doc_from_template(template_path=template_path, loader=loader)

    # endregion create_doc_from_template()

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
        return cls._lo_inst.save(doc)

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
        return cls._lo_inst.save_doc(doc=doc, fnm=fnm, password=password, format=format)

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
        return cls._lo_inst.store_doc(store=store, doc_type=doc_type, fnm=fnm, password=password)

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
    def ext_to_format(cls, ext: str, doc_type: LoDocType = LoDocType.UNKNOWN) -> str:
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
        return cls._lo_inst.ext_to_format(ext=ext, doc_type=doc_type)

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
        return cls._lo_inst.store_doc_format(store=store, fnm=fnm, format=format, password=password)

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
        return cls._lo_inst.close(closeable=closeable, deliver_ownership=deliver_ownership)

    # region close_doc()
    @overload
    @classmethod
    def close_doc(cls, doc: object) -> None:
        ...

    @overload
    @classmethod
    def close_doc(cls, doc: object, deliver_ownership: bool) -> None:
        ...

    @classmethod
    def close_doc(cls, doc: object, deliver_ownership=False) -> None:
        """
        Closes document.

        Args:
            doc (XCloseable): Close-able document
            deliver_ownership (bool): If ``True`` delegates the ownership of this closing object to
                anyone which throw the CloseVetoException. Default is ``False``.

        Raises:
            MissingInterfaceError: if doc does not have XCloseable interface

        Returns:
            None:

        Note:
            If ``deliver_ownership`` is ``True`` then new owner has to close the closing object again if his still running
            processes will be finished.

            ``False`` let the ownership at the original one which called the close() method.
            They must react for possible CloseVetoExceptions such as when document needs saving
            and try it again at a later time. This can be useful for a generic UI handling.

        Attention:
            :py:meth:`~.utils.lo.Lo.close` method is called along with any of its events.
        """
        cls._lo_inst.close_doc(doc=doc, deliver_ownership=deliver_ownership)

    # endregion close_doc()

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
        return cls._lo_inst.addon_initialize(addon_xcc=addon_xcc)

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
        return cls._lo_inst.script_initialize(sc=sc)

    # ==================== dispatch ===============================
    # see https://wiki.documentfoundation.org/Development/DispatchCommands

    # region dispatch_cmd()
    @overload
    @classmethod
    def dispatch_cmd(cls, cmd: str) -> Any:
        ...

    @overload
    @classmethod
    def dispatch_cmd(cls, cmd: str, props: Iterable[PropertyValue]) -> Any:
        ...

    @overload
    @classmethod
    def dispatch_cmd(cls, cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> Any:
        ...

    @overload
    @classmethod
    def dispatch_cmd(cls, cmd: str, *, frame: XFrame) -> Any:
        ...

    @classmethod
    def dispatch_cmd(cls, cmd: str, props: Iterable[PropertyValue] = None, frame: XFrame = None) -> Any:
        """
        Dispatches a LibreOffice command

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue, optional): properties for dispatch
            frame (XFrame, optonal): Frame to dispatch to.

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
        """
        return cls._lo_inst.dispatch_cmd(cmd=cmd, props=props, frame=frame)

    # endregion dispatch_cmd()

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

    @classmethod
    def extract_item_name(cls, uno_cmd: str) -> str:
        """
        Extract a uno command from a string that was created with :py:meth:`~Lo.make_uno_cmd`

        Args:
            uno_cmd (str): uno command

        Raises:
            ValueError: If unable to extract command

        Returns:
            str: uno command
        """
        return cls._lo_inst.extract_item_name(uno_cmd)

    # ======================== use Inspector extensions ====================

    @classmethod
    def inspect(cls, obj: object) -> None:
        """
        Inspects object using ``org.openoffice.InstanceInspector`` inspector.

        Args:
            obj (object): object to inspect.
        """
        cls._lo_inst.inspect(obj)

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
        cls._lo_inst.mri_inspect(obj)

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
            Lo.print("Lo.delay(): Ms must be greater then zero")
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

    @classmethod
    def is_url(cls, fnm: PathOrStr) -> bool:
        """
        Gets if a string is a URL format.

        Args:
            fnm (PathOrStr): string to check.

        Returns:
            bool: True if URL format; Otherwise, False
        """
        return cls._lo_inst.is_url(fnm)

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
    @classmethod
    def print_names(cls, names: Sequence[str]) -> None:
        ...

    @overload
    @classmethod
    def print_names(cls, names: Sequence[str], num_per_line: int) -> None:
        ...

    @classmethod
    def print_names(cls, names: Sequence[str], num_per_line: int = 4) -> None:
        """
        Prints names to console

        Args:
            names (Iterable[str]): names to print
            num_per_line (int): Number of names per line. Default ``4``
            format_opt (FormatterTable, optional): Optional format used to format values when printing to console such as ``FormatterTable(format=">2")``

        Returns:
            None:

        Example:
            Given a list of ``20`` names the output is similar to:

            ::

                No. of names: 20
                  ----------|-----------|-----------|-----------
                  Accent    | Accent 1  | Accent 2  | Accent 3
                  Bad       | Default   | Error     | Footnote
                  Good      | Heading   | Heading 1 | Heading 2
                  Hyperlink | Neutral   | Note      | Result
                  Result2   | Status    | Text      | Warning
        """
        cls._lo_inst.print_names(names=names, num_per_line=num_per_line)

    # ------------------- container manipulation --------------------
    # region print_table()
    @overload
    @classmethod
    def print_table(cls, name: str, table: Table) -> None:
        ...

    @overload
    @classmethod
    def print_table(cls, name: str, table: Table, format_opt: FormatterTable) -> None:
        ...

    @classmethod
    def print_table(cls, name: str, table: Table, format_opt: FormatterTable | None = None) -> None:
        """
        Prints a 2-Dimensional table to console

        Args:
            name (str): Name of table
            table (Table): Table Data
            format_opt (FormatterTable, optional): Optional format used to format values when printing to console such as ``FormatterTable(format=".2f")``

        Returns:
            None:

        See Also:
            - :ref:`ch21_format_data_console`
            - :py:data:`~.type_var.Table`

        .. versionchanged:: 0.6.7
            Added ``format_opt`` parameter
        """
        cls._lo_inst.print_table(name=name, table=table, format_opt=format_opt)

    # endregion print_table()

    @classmethod
    def get_container_names(cls, con: XIndexAccess) -> List[str] | None:
        """
        Gets container names

        Args:
            con (XIndexAccess): container

        Returns:
            List[str] | None: Container name is found; Otherwise, None
        """
        return cls._lo_inst.get_container_names(con)

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
        return cls._lo_inst.find_container_props(con=con, nm=nm)

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
        return cls._lo_inst.is_uno_interfaces(component, *args)

    @classmethod
    def get_frame(cls) -> XFrame:
        """
        Gets XFrame for current LibreOffice instance

        Returns:
            XFrame: frame
        """
        return cls._lo_inst.get_frame()

    @classmethod
    def get_model(cls) -> XModel:
        """
        Gets XModel

        Returns:
            XModel: Gets model for current LibreOffice instance
        """
        return cls._lo_inst.get_model()

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
        return cls._lo_inst.lock_controllers()

    @classmethod
    def unlock_controllers(cls) -> bool:
        """
        Resumes the notifications which were suspended by :py:meth:`~.lo.Lo.lock_controllers`.

        The calls to :py:meth:`~.lo.Lo.lock_controllers` and :py:meth:`~.lo.Lo.unlock_controllers`
        may be nested and even overlapping, but they must be in pairs.
        While there is at least one lock remaining, some notifications for
        display updates are not broadcast.

        Raises:
            MissingInterfaceError: If unable to obtain XModel interface.

        Returns:
            bool: False if ``CONTROLERS_UNLOCKING`` event is canceled; Otherwise, True

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.CONTROLERS_UNLOCKING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.CONTROLERS_UNLOCKED` :eventref:`src-docs-event`

        See Also:
            :py:class:`.Lo.ControllerLock`
        """
        return cls._lo_inst.unlock_controllers()

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
        return cls._lo_inst.has_controllers_locked()

    @classmethod
    def print(cls, *args, **kwargs) -> None:
        """
        Utility function that passes to actual print.

        If :py:attr:`GblNamedEvent.PRINTING <.events.gbl_named_event.GblNamedEvent.PRINTING>`
        event is canceled the this method will not print.

        :events:
           .. include:: ../../resources/global/printing_events.rst

        Note:
            .. include:: ../../resources/global/printing_note.rst
        """
        cls._lo_inst.print(*args, **kwargs)

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
        return cls._lo_inst.null_date

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
        return cls._lo_inst is not None

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
        return cls._lo_inst.is_macro_mode

    @is_macro_mode.setter
    def is_macro_mode(cls, value) -> None:
        # raise error on set. Not really necessary but gives feedback.
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.__name__)

    @classproperty
    def star_desktop(cls) -> XDesktop:
        """Get current desktop"""
        return cls._lo_inst.star_desktop

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
        return cls._lo_inst.this_component

    ThisComponent, thiscomponent = this_component, this_component

    @classproperty
    def xscript_context(cls) -> XScriptContext:
        """
        a substitute to `XSCRIPTCONTEXT` (Libre|Open)Office built-in

        Returns:
            XScriptContext: XScriptContext instance
        """
        return cls._lo_inst.xscript_context

    XSCRIPTCONTEXT = xscript_context

    @classproperty
    def bridge(cls) -> XComponent:
        """
        Gets connection bridge component

        Returns:
            XComponent: bridge component
        """
        return cls._lo_inst.bridge

    @classproperty
    def loader_current(cls) -> XComponentLoader:
        """
        Gets the current ``XComponentLoader`` instance.

        Returns:
            XComponentLoader: Component Loader
        """
        return cls._lo_inst.loader_current

    @classproperty
    def bridge_connector(cls) -> LoBridgeCommon:
        """
        Get the current Bride connection

        Returns:
            LoBridgeCommon: If not in macro mode; Otherwise, ``ConnectBase``
        """
        return cls._lo_inst.bridge_connector

    @classproperty
    def options(cls) -> LoOptions:
        """
        Gets the current options.

        Returns:
            LoOptions: Options
        """
        return cls._lo_inst._opt


def _on_bridge_disposed(source: Any, event: EventObject) -> None:
    setattr(Lo, "_lo_inst", None)


_Events().on(LoNamedEvent.BRIDGE_DISPOSED, _on_bridge_disposed)


__all__ = ("Lo",)
