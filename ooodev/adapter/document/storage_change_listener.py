from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.document import XStorageChangeListener

from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from com.sun.star.embed import XStorage
    from com.sun.star.document import XStorageBasedDocument


class StorageChangeListener(AdapterBase, XStorageChangeListener):
    """
    allows to be notified when a document is switched to a new storage.

    See Also:
        `API XStorageChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XStorageChangeListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XStorageBasedDocument | None = None
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XStorageBasedDocument, optional): An UNO object that implements the ``XStorageBasedDocument`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addStorageChangeListener(self)

    def notifyStorageChange(self, document: XInterface, storage: XStorage) -> None:
        """
        Event is invoked when document switches to another storage.

        When this method is triggered it raises the ``notifyStorageChange`` event.
        The event data is a dictionary with the following keys:

        - ``document``: The document that is being switched to another storage.
        - ``storage``: The new storage that the document is being switched to.
        """
        args = EventArgs(self)
        args.event_data = {"document": document, "storage": storage}
        self._trigger_event("notifyStorageChange", args)
