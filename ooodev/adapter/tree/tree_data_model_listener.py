from __future__ import annotations

# pylint: disable=invalid-name, unused-import
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt.tree import XTreeDataModelListener
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt.tree import TreeDataModelEvent


class TreeDataModelListener(AdapterBase, XTreeDataModelListener):
    """
    An instance of this interface is used by the TreeControl to get notifications about data model changes.

    Usually you must not implement this interface yourself as it is already handled by the TreeControl,
    but you must notify it correctly if you implement the XTreeDataModel yourself.

    See Also:
        `API XTreeDataModelListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1tree_1_1XTreeDataModelListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XTreeDataModelListener

    def treeNodesChanged(self, event: TreeDataModelEvent) -> None:
        """
        Event is invoked after a node (or a set of siblings) has changed in some way.

        The node(s) have not changed locations in the tree or altered their children arrays,
        but other attributes have changed and may affect presentation.

        Example: the name of a file has changed, but it is in the same location in the file system.

        To indicate the root has changed, TreeDataModelEvent.Nodes will contain the root
        node and ``TreeDataModelEvent.ParentNode`` will be empty.
        """
        self._trigger_event("treeNodesChanged", event)

    def treeNodesInserted(self, event: TreeDataModelEvent) -> None:
        """
        Invoked after nodes have been inserted into the tree.

        Use ``TreeDataModelEvent.ParentNode`` to get the parent of the new node(s).
        ``TreeDataModelEvent.Nodes`` contains the new node(s).
        """
        self._trigger_event("treeNodesInserted", event)

    def treeNodesRemoved(self, event: TreeDataModelEvent) -> None:
        """
        Invoked after nodes have been removed from the tree.

        Note that if a subtree is removed from the tree, this method may only be invoked once for the root of the
        removed subtree, not once for each individual set of siblings removed.

        Use ``TreeDataModelEvent.ParentNode`` to get the former parent of the deleted node(s).
        ``TreeDataModelEvent.Nodes`` contains the removed node(s).
        """
        self._trigger_event("treeNodesRemoved", event)

    def treeStructureChanged(self, event: TreeDataModelEvent) -> None:
        """
        Invoked after the tree has drastically changed structure from a given node down.

        Use ``TreeDataModelEvent.ParentNode`` to get the node which structure has changed.
        ``TreeDataModelEvent.Nodes`` is empty.
        """
        self._trigger_event("treeStructureChanged", event)

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

    # endregion XTreeDataModelListener
