from __future__ import annotations
from typing import Iterable, TYPE_CHECKING, overload
import uno
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.loader import lo as mLo
from ooodev.loader.inst.doc_type import DocType


if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.beans import PropertyValue
    from ooodev.utils.type_var import PathOrStr


class LoOpenPartial:
    """Partial Class used for opening Office Documents."""

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        """
        Constructor.

        Args:
            lo_inst (LoInst, optional): Lo instance.
        """
        if lo_inst is None:
            self.__lo_inst = mLo.Lo.current_lo
        else:
            self.__lo_inst = lo_inst

    # region open_doc()
    @overload
    def open_doc(self, fnm: PathOrStr) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open

        Returns:
            XComponent: Document
        """
        ...

    @overload
    def open_doc(self, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Returns:
            XComponent: Document
        """
        ...

    @overload
    def open_doc(self, fnm: PathOrStr, *, props: Iterable[PropertyValue]) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            props (Iterable[PropertyValue]): Properties passed to component loader

        Returns:
            XComponent: Document
        """
        ...

    @overload
    def open_doc(self, fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> XComponent:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader


        Returns:
            XComponent: Document
        """
        ...

    def open_doc(
        self,
        fnm: PathOrStr,
        loader: XComponentLoader | None = None,
        props: Iterable[PropertyValue] | None = None,
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
            - :py:meth:`~Lo.open_doc`
            - :py:meth:`load_office`
            - :ref:`ch02_open_doc`

        Note:
            If connection it office is a remote server then File URL must be used,
            such as ``file:///home/user/fancy.odt``
        """
        return self.__lo_inst.open_doc(fnm=fnm, loader=loader, props=props)  # type: ignore

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    def open_readonly_doc(self, fnm: PathOrStr) -> XComponent:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open

        Returns:
            XComponent: Document
        """
        ...

    @overload
    def open_readonly_doc(self, fnm: PathOrStr, loader: XComponentLoader) -> XComponent:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Returns:
            XComponent: Document
        """
        ...

    def open_readonly_doc(self, fnm: PathOrStr, loader: XComponentLoader | None = None) -> XComponent:
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
            - :ref:`ch02_open_doc`
        """
        if loader is None:
            return self.__lo_inst.open_readonly_doc(fnm=fnm)
        return self.__lo_inst.open_readonly_doc(fnm=fnm, loader=loader)

    # endregion open_readonly_doc()

    # region open_flat_doc()
    @overload
    def open_flat_doc(self, fnm: PathOrStr, doc_type: DocType) -> XComponent: ...

    @overload
    def open_flat_doc(self, fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader) -> XComponent: ...

    def open_flat_doc(self, fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader | None = None) -> XComponent:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader): Component loader

        Returns:
            XComponent: Document

        See Also:
            - :py:meth:`~Lo.open_flat_doc`
            - :ref:`ch02_open_doc`
        """
        if loader is None:
            return self.__lo_inst.open_flat_doc(fnm=fnm, doc_type=doc_type)
        return self.__lo_inst.open_flat_doc(fnm=fnm, doc_type=doc_type, loader=loader)

    # endregion open_flat_doc()
