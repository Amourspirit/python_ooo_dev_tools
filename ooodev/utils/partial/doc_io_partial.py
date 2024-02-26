from __future__ import annotations
from typing import Any, cast, overload, TYPE_CHECKING, TypeVar, Generic, ClassVar
import uno
from com.sun.star.frame import XComponentLoader
from com.sun.star.util import XCloseable
from com.sun.star.frame import XModule

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.loader.inst.doc_type import DocType
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps
from ooodev.utils.factory import doc_factory as mDocFactory
from ooodev.utils.type_var import PathOrStr

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from com.sun.star.lang import XComponent

_T = TypeVar("_T", bound="ComponentT")


class DocIoPartial(Generic[_T]):
    DOC_TYPE: ClassVar[DocType]

    def __init__(self, owner: _T, lo_inst: LoInst | None = None):
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__owner = owner

    @classmethod
    def get_doc_from_component(cls, doc: XComponent, lo_inst: LoInst | None) -> _T:
        """
        Gets a document.

        Args:
            doc (XComponent): Component to build Draw document from.

        Raises:
            Exception: If not a Draw document.

        Returns:
            _T : Document.
        """
        if not mInfo.Info.is_doc_type(doc_type=cls.DOC_TYPE.get_service(), obj=doc):
            raise Exception("Not a valid document")
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        doc_instance = cast(_T, mDocFactory.doc_factory(doc=doc, lo_inst=lo_inst))
        return doc_instance

    # region Close
    def close(self, deliver_ownership=True) -> bool:
        """
        Try to close the Document.

        Nobody can guarantee real closing of called object - because it can disagree with that if any still running processes can't be canceled yet.
        It's not allowed to block this call till internal operations will be finished here.

        Args:
            deliver_ownership (bool, optional): If ``True`` ownership is delivered to caller. Default ``True``.
                ``True`` delegates the ownership of this closing object to anyone which throw the CloseVetoException.
                This new owner has to close the closing object again if his still running processes will be finished.
                ``False`` let the ownership at the original one which called the close() method.
                They must react for possible CloseVetoExceptions such as when document needs saving and
                try it again at a later time. This can be useful for a generic UI handling.

        Raises:
            CancelEventError: If Saving event is canceled.

        Returns:
            bool: ``True`` if document was closed; Otherwise, ``False``.

        See Also:
            See LibreOffice API: `XCloseable.close() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XCloseable.html#af3d34677f1707b1904f8e07be4408592>`__
        """
        event_data = {"deliver_ownership": deliver_ownership}
        cargs = CancelEventArgs(self.close.__qualname__)
        cargs.event_data = event_data
        self._on_io_closing(cargs)
        if cargs.cancel and cargs.handled:
            raise mEx.CancelEventError(cargs, "Doc closing event was canceled")
        closable = self.__lo_inst.qi(XCloseable, self.__owner.component, True)
        result = self.__lo_inst.close(closable, deliver_ownership)
        self._on_io_closed(EventArgs.from_args(cargs))
        return result

    def _on_io_closing(self, event_args: CancelEventArgs) -> None:
        """
        Event called before document is closed.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        pass

    def _on_io_closed(self, event_args: EventArgs) -> None:
        """
        Event called after document is closed.

        Args:
            event_args (EventArgs): Event data.
        """
        pass

    # endregion Close

    # region save_doc

    @overload
    def save_doc(self, fnm: PathOrStr) -> bool:
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.

        Returns:
            bool: ``False`` if ``DOC_SAVING`` event is canceled; Otherwise, ``True``
        """
        ...

    @overload
    def save_doc(self, fnm: PathOrStr, password: str) -> bool:
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.
            password (str): Password to save document with.


        Returns:
            bool: ``False`` if ``DOC_SAVING`` event is canceled; Otherwise, ``True``
        """
        ...

    @overload
    def save_doc(self, fnm: PathOrStr, password: str, format: str) -> bool:  # pylint: disable=W0622
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.
            password (str): Password to save document with.
            format (str): document format such as 'odt' or 'xml'.

        Returns:
            bool: ``False`` if ``DOC_SAVING`` event is canceled; Otherwise, ``True``.
        """
        ...

    def save_doc(self, fnm: PathOrStr, password: str | None = None, format: str | None = None) -> bool:
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.
            password (str, optional): password to save document with.
            format (str, optional): document format such as 'odt' or 'xml'.

        Raises:
            CancelEventError: If Saving event is canceled.


        Returns:
            bool: ``False`` if DOC_SAVING event is canceled; Otherwise, ``True``

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
        event_data = {"fnm": fnm, "password": password, "format": format}

        cargs = CancelEventArgs(self.close.__qualname__)
        cargs.event_data = event_data
        self._on_io_saving(cargs)
        if cargs.cancel:
            if cargs.handled:
                raise mEx.CancelEventError(cargs, "Doc saving event was canceled")
        result = self.__lo_inst.save_doc(self.component, fnm, password, format)  # type: ignore
        self._on_io_saved(EventArgs.from_args(cargs))
        return result

    def _on_io_saving(self, event_args: CancelEventArgs) -> None:
        """
        Event called before document is saved.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        pass

    def _on_io_saved(self, event_args: EventArgs) -> None:
        """
        Event called after document is saved.

        Args:
            event_args (EventArgs): Event data.
        """
        pass

    # endregion save_doc

    # region Create Document
    @overload
    @classmethod
    def create_doc(cls) -> _T:
        """
        Creates a new document.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, *, lo_inst: LoInst) -> _T:
        """
        Creates a new document.

        Args:
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, *, visible: bool) -> _T:
        """
        Creates a new document.

        Args:
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible. Default ``False``.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, *, lo_inst: LoInst, visible: bool) -> _T:
        """
        Creates a new document.

        Args:
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible. Default ``False``.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, loader: XComponentLoader) -> _T:
        """
        Creates a new document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, loader: XComponentLoader, *, visible: bool) -> _T:
        """
        Creates a new document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible. Default ``False``.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, loader: XComponentLoader, lo_inst: LoInst) -> _T:
        """
        Creates a new document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc(cls, loader: XComponentLoader, lo_inst: LoInst, *, visible: bool) -> _T:
        """
        Creates a new document.

        Args:
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible. Default ``False``.

        Returns:
            _T: Class instance representing document.
        """
        ...

    # @staticmethod
    # def create_doc(*args, **kwargs) -> DrawDoc:
    @classmethod
    def create_doc(cls, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None, **kwargs: Any) -> _T:
        """
        Creates a new document.

        Args:
            loader (XComponentLoader, optional): Component Loader. Usually generated with :py:class:`~.lo.Lo`
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents. Defaults to None.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible. Default ``False``.

        Returns:
            _T: Class instance representing document.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        cargs = CancelEventArgs(cls.create_doc.__qualname__)

        event_data = {"loader": loader, "lo_inst": lo_inst}
        if "visible" in kwargs:
            event_data["visible"] = bool(kwargs.pop("visible"))

        cargs.event_data = event_data
        cls._on_io_creating_doc(cargs)

        props_dict = {"Hidden": True}
        if "visible" in cargs.event_data:
            visible = bool(cargs.event_data["visible"])
            props_dict["Hidden"] = not visible

        if kwargs:
            props_dict.update(kwargs)

        local_props = mProps.Props.make_props(**props_dict)
        if loader is None:
            doc = lo_inst.create_doc(doc_type=cls.DOC_TYPE.get_doc_type_str(), props=local_props)
        else:
            doc = lo_inst.create_doc(doc_type=cls.DOC_TYPE.get_doc_type_str(), loader=loader, props=local_props)
        result = cls.get_doc_from_component(doc=doc, lo_inst=lo_inst)
        cls._on_io_created_doc(EventArgs.from_args(cargs))
        return result

    @classmethod
    def _on_io_creating_doc(cls, event_args: CancelEventArgs) -> None:
        """
        Event called before document is Created.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        pass

    @classmethod
    def _on_io_created_doc(cls, event_args: EventArgs) -> None:
        """
        Event called after document is Created.

        Args:
            event_args (EventArgs): Event data.
        """
        pass

    # endregion Create Document

    # region create_doc_from_template()

    @overload
    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr) -> _T:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr, *, lo_inst: LoInst) -> _T:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr, loader: XComponentLoader) -> _T:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader): Component Loader.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_doc_from_template(cls, template_path: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> _T:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @classmethod
    def create_doc_from_template(
        cls, template_path: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> _T:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader, optional): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Raises:
            Exception: If unable to create document.

        Returns:
            _T: Class instance representing document.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if loader is None:
            doc = lo_inst.create_doc_from_template(template_path=template_path)
        else:
            doc = lo_inst.create_doc_from_template(template_path=template_path, loader=loader)
        return cls.get_doc_from_component(doc=doc, lo_inst=lo_inst)

    # endregion create_doc_from_template()

    # region from_current_doc()

    @classmethod
    def from_current_doc(cls) -> _T:
        """
        Get a document from the current component.

        This method is useful in macros where the access to current document is needed.
        This method does not require the use of the :py:class:`~ooodev.macro.MacroLoader` in macros.

        Args:
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.

        Returns:
            _T: Class instance representing document.

        Example:
            .. code-block:: python

                from ooodev.calc import CalcDoc
                doc = CalcDoc.from_current_doc()
                doc.sheets[0]["A1"].Value = "Hello World"

        See Also:
            :py:attr:`ooodev.utils.lo.Lo.current_doc`
        """
        cargs = CancelEventArgs(cls.from_current_doc.__qualname__)
        cargs.event_data = {"doc_type": None}
        cls._on_from_current_doc_loading(cargs)
        if cargs.cancel:
            if not cargs.handled:
                if "msg" in cargs.event_data:
                    msg = cargs.event_data["msg"]
                else:
                    msg = "from_current_doc event was canceled"
                raise mEx.CancelEventError(cargs, msg)

        doc_type = cast(DocType, cargs.event_data["doc_type"])
        doc = None
        if doc_type is None:
            doc = mLo.Lo.current_doc
        else:
            doc_service_name = str(doc_type.get_service())
            # if the current doc is a match the prefer it.
            comp = mLo.Lo.desktop.component.getCurrentComponent()
            module = mLo.Lo.qi(XModule, comp)
            if module is not None:
                identifier = module.getIdentifier()
                if identifier == doc_service_name:
                    doc = mDocFactory.doc_factory(doc=comp, lo_inst=mLo.Lo.current_lo)

            if doc is None:
                for comp in mLo.Lo.desktop.components:
                    # if there is more then on component then the first match is used.
                    # It seems the last opened document is the first in the list.
                    module = mLo.Lo.qi(XModule, comp, True)
                    identifier = module.getIdentifier()
                    if identifier == doc_service_name:
                        doc = mDocFactory.doc_factory(doc=comp, lo_inst=mLo.Lo.current_lo)
                        break
        if doc is None:
            raise mEx.NotSupportedDocumentError("No supported document found")
        args = EventArgs(cls.from_current_doc.__qualname__)
        args.event_data = {"doc": doc}
        cls._on_from_current_doc_loaded(args)
        return cast(_T, doc)

    @classmethod
    def _on_from_current_doc_loading(cls, event_args: CancelEventArgs) -> None:
        """
        Event called while from_current_doc loading.

        Args:
            event_args (EventArgs): Event data.

        Returns:
            None:

        Note:
            event_args.event_data is a dictionary and contains the document in a key named 'doc'.
        """
        pass

    @classmethod
    def _on_from_current_doc_loaded(cls, event_args: EventArgs) -> None:
        """
        Event called after from_current_doc is loaded.

        Args:
            event_args (EventArgs): Event data.

        Returns:
            None:

        Note:
            event_args.event_data is a dictionary and contains the document in a key named 'doc'.
        """
        pass

    # endregion from_current_doc()

    # region open_doc()
    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr) -> _T:
        """
        Open a office document.

        Args:
            fnm (PathOrStr): path of document to open.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr) -> _T:
        """
        Open a office document.

        Args:
            fnm (PathOrStr): path of document to open.

        Returns:
            DrawDoc: Document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, *, lo_inst: LoInst | None) -> _T:
        """
        Open a office document.

        Args:
            fnm (PathOrStr): path of document to open.
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents. Defaults to None.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, *, visible: bool) -> _T:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> _T:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader, *, visible: bool) -> _T:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader, *, lo_inst: LoInst) -> _T:
        """
        Open a office document.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents. Defaults to None.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader, *, lo_inst: LoInst, visible: bool) -> _T:
        """
        Open a office document.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents. Defaults to None.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @classmethod
    def open_doc(
        cls, fnm: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None, **kwargs: Any
    ) -> _T:
        """
        Open a office document.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader, optional): Component Loader.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible. Default ``False``.

        Raises:
            CancelEventError: if ``DOC_OPENING`` event is canceled.

        Returns:
            _T: Class instance representing document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            - :py:meth:`~Lo.open_doc`
            - :py:meth:`load_office`
            - :ref:`ch02_open_doc`

        Note:
            If connection it office is a remote server then File URL must be used,
            such as ``file:///home/user/fancy.odt``
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        cargs = CancelEventArgs(cls.open_doc.__qualname__)

        event_data = {"loader": loader, "lo_inst": lo_inst}
        if "visible" in kwargs:
            event_data["visible"] = bool(kwargs.pop("visible"))

        cargs.event_data = event_data
        cls._on_io_opening_doc(cargs)

        props_dict = {"Hidden": True}
        if "visible" in cargs.event_data:
            visible = bool(cargs.event_data["visible"])
            props_dict["Hidden"] = not visible

        if kwargs:
            props_dict.update(kwargs)
        local_props = mProps.Props.make_props(**props_dict)

        doc = lo_inst.open_doc(fnm=fnm, loader=loader, props=local_props)  # type: ignore
        result = cls.get_doc_from_component(doc=doc, lo_inst=lo_inst)
        cls._on_io_opened_doc(EventArgs.from_args(cargs))
        return result

    @classmethod
    def _on_io_opening_doc(cls, event_args: CancelEventArgs) -> None:
        """
        Event called before document is Opened.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        pass

    @classmethod
    def _on_io_opened_doc(cls, event_args: EventArgs) -> None:
        """
        Event called after document is Opened.

        Args:
            event_args (EventArgs): Event data.
        """
        pass

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr) -> _T:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, *, visible: bool) -> _T:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, *, lo_inst: LoInst) -> _T:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> _T:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader, *, visible: bool) -> _T:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> _T:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_readonly_doc(cls, fnm: PathOrStr, loader: XComponentLoader, lo_inst: LoInst, *, visible: bool) -> _T:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.
            visible (bool): If ``True`` document is visible; Otherwise, document is invisible.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @classmethod
    def open_readonly_doc(
        cls, fnm: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None, **kwargs: Any
    ) -> _T:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.

        See Also:
            - :ref:`ch02_open_doc`
        """
        key_args = kwargs.copy()
        key_args["ReadOnly"] = True
        return cls.open_doc(fnm=fnm, loader=loader, lo_inst=lo_inst, **key_args)  # type: ignore

    # endregion open_readonly_doc()

    # region open_flat_doc()
    @overload
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr) -> _T:
        """
        Opens a flat document.

        Args:
            fnm (PathOrStr): path of XML document.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr, *, lo_inst: LoInst) -> _T:
        """
        Opens a flat document.

        Args:
            fnm (PathOrStr): path of XML document.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> _T:
        """
        Opens a flat document.

        Args:
            fnm (PathOrStr): path of XML document.
            loader (XComponentLoader, optional): Component loader.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def open_flat_doc(cls, fnm: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> _T:
        """
        Opens a flat document.

        Args:
            fnm (PathOrStr): path of XML document.
            loader (XComponentLoader, optional): Component loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @classmethod
    def open_flat_doc(
        cls, fnm: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> _T:
        """
        Opens a flat document.

        Args:
            fnm (PathOrStr): path of XML document.
            loader (XComponentLoader, optional): Component loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Raises:
            Exception: if unable to open document.

        Returns:
            _T: Class instance representing document.

        See Also:
            - :py:meth:`~Lo.open_flat_doc`
            - :ref:`ch02_open_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if loader is None:
            doc = lo_inst.open_flat_doc(fnm=fnm, doc_type=cls.DOC_TYPE)
        else:
            doc = lo_inst.open_flat_doc(fnm=fnm, doc_type=cls.DOC_TYPE, loader=loader)
        return cls.get_doc_from_component(doc=doc, lo_inst=lo_inst)

    # endregion open_flat_doc()

    # region create_macro_doc()
    @overload
    @classmethod
    def create_macro_doc(cls, *, lo_inst: LoInst) -> _T:
        """
        Create a document that allows executing of macros.

        Args:
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_macro_doc(cls, loader: XComponentLoader) -> _T:
        """
        Create a document that allows executing of macros.

        Args:
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @overload
    @classmethod
    def create_macro_doc(cls, loader: XComponentLoader, lo_inst: LoInst) -> _T:
        """
        Create a document that allows executing of macros.

        Args:
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo Instance. Use when creating multiple documents.

        Returns:
            _T: Class instance representing document.
        """
        ...

    @classmethod
    def create_macro_doc(cls, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None) -> _T:
        """
        Create a document that allows executing of macros.

        Args:
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.

        Returns:
            _T: Class instance representing document.

        Attention:
            :py:meth:`~.utils.lo.Lo.create_doc` method is called along with any of its events.

        See Also:
            :ref:`ch02_create_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        doc_type = mLo.Lo.DocTypeStr.CALC
        if loader is None:
            doc = mLo.Lo.create_macro_doc(doc_type=doc_type)
        else:
            doc = mLo.Lo.create_macro_doc(doc_type=doc_type, loader=loader)
        return cls.get_doc_from_component(doc=doc, lo_inst=lo_inst)

    # endregion create_macro_doc()
