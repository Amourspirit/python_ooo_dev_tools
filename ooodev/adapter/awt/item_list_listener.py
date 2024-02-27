from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt import XItemListListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import ItemListEvent
    from com.sun.star.awt import XItemList


class ItemListListener(AdapterBase, XItemListListener):
    """
    Describes a listener for changes in an item list

    See Also:
        `API XItemListListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XItemListListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XItemList | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XItemList, optional): An UNO object that implements the ``com.sun.star.form.XItemList`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addItemListListener(self)

    def allItemsRemoved(self, event: EventObject) -> None:
        """
        Event is invoked when the list has been completely cleared, i.e. after an invocation of ``XItemList.removeAllItems()``

        Args:
            event (EventObject): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("allItemsRemoved", event)

    def itemListChanged(self, event: EventObject) -> None:
        """
        Event is invoked when the changes to the item list which occurred are too complex to be notified in single events.

        Consumers of this event should discard their cached information about the current item list, and
        completely refresh it from the XItemList's current state.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("itemListChanged", event)

    def listItemInserted(self, event: ItemListEvent) -> None:
        """
        Event is invoked when an item is inserted into the list.

        Args:
            event (ItemListEvent): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("listItemInserted", event)

    def listItemModified(self, event: ItemListEvent) -> None:
        """
        Event is invoked when an item in the list is modified, i.e. its text or image changed.

        Args:
            event (ItemListEvent): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("listItemModified", event)

    def listItemRemoved(self, event: ItemListEvent) -> None:
        """
        Event is invoked when an item is removed from the list.

        Args:
            event (ItemListEvent): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("listItemRemoved", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets invoked when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            None:
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)
