from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.form.binding import XListEntryListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.form.binding import ListEntryEvent
    from com.sun.star.form.binding import XListEntrySource


class ListEntryListener(AdapterBase, XListEntryListener):
    """
    Specifies a listener for changes in a string entry list

    See Also:
        `API XListEntryListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1binding_1_1XListEntryListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XListEntrySource | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XListEntrySource, XWindow, optional): An UNO object that implements the ``XExtendedToolkit`` or ``XWindow`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addListEntryListener(self)

    # region XListEntryListener

    def allEntriesChanged(self, event: EventObject) -> None:
        """
        Notifies the listener that all entries of the list have changed.

        The listener should retrieve the complete new list by calling the
        ``XListEntrySource.getAllListEntries()`` method of the event source
        (which is denoted by com.sun.star.lang.EventObject.Source).
        """
        self._trigger_event("allEntriesChanged", event)

    def entryChanged(self, event: ListEntryEvent) -> None:
        """
        Notifies the listener that a single entry in the list has change,
        """
        self._trigger_event("entryChanged", event)

    def entryRangeInserted(self, event: ListEntryEvent) -> None:
        """
        Notifies the listener that a range of entries has been inserted into the list,
        """
        self._trigger_event("entryRangeInserted", event)

    def entryRangeRemoved(self, event: ListEntryEvent) -> None:
        """
        Notifies the listener that a range of entries has been removed from the list,
        """
        self._trigger_event("entryRangeRemoved", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)

    # endregion XListEntryListener
