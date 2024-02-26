from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt.tree import XTreeExpansionListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt.tree import TreeExpansionEvent
    from com.sun.star.awt.tree import XTreeControl


class TreeExpansionListener(AdapterBase, XTreeExpansionListener):
    """
    An instance of this interface can get notifications from a TreeControl when nodes are expanded or collapsed.

    See Also:
        `API XTreeExpansionListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1tree_1_1XTreeExpansionListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XTreeControl | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XTreeControl, optional): An UNO object that implements the ``XTreeControl`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addTreeExpansionListener(self)

    # region XTreeExpansionListener
    def requestChildNodes(self, event: TreeExpansionEvent) -> None:
        """
        Invoked when a node with children on demand is about to be expanded.

        This event is invoked before the treeExpanding() event.
        """
        self._trigger_event("requestChildNodes", event)

    def treeCollapsed(self, event: TreeExpansionEvent) -> None:
        """
        Called whenever a node in the tree has been successfully collapsed.
        """
        self._trigger_event("treeCollapsed", event)

    def treeCollapsing(self, event: TreeExpansionEvent) -> None:
        """
        Invoked whenever a node in the tree is about to be collapsed.

        Raises:
            ExpandVetoException: ``ExpandVetoException``
        """
        self._trigger_event("treeCollapsing", event)

    def treeExpanded(self, event: TreeExpansionEvent) -> None:
        """
        Invoked whenever a node in the tree has been successfully expanded.
        """
        self._trigger_event("treeExpanded", event)

    def treeExpanding(self, event: TreeExpansionEvent) -> None:
        """
        Invoked whenever a node in the tree is about to be expanded.

        Raises:
            ExpandVetoException: ``ExpandVetoException``
        """
        self._trigger_event("treeExpanding", event)

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

    # endregion XTreeExpansionListener
