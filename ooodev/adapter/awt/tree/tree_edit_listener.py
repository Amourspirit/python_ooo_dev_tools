from __future__ import annotations
from typing import NamedTuple, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.awt.tree import XTreeEditListener
from com.sun.star.awt.tree import XTreeNode

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt.tree import XTreeControl


class NodeEditedArgs(NamedTuple):
    """Node Edit Args"""

    node: XTreeNode
    """Node being edited"""
    new_text: str
    """New Text"""


class TreeEditListener(AdapterBase, XTreeEditListener):
    """
    You can implement this interface and register with XTreeControl.addTreeEditListener() to get notifications when editing of a node starts and ends.

    You have to set the TreeControlModel.Editable property to TRUE before a tree supports editing.

    See Also:
        `API XTreeEditListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1tree_1_1XTreeEditListener.html>`_
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
            subscriber.addTreeEditListener(self)

    # region XTreeEditListener
    @override
    def nodeEdited(self, Node: XTreeNode, NewText: str) -> None:
        """
        This method is called from the TreeControl implementation when editing of Node is finished and was not canceled.

        Implementations that register a XTreeEditListener must update the display value at the Node.

        Note:
            ``EventArgs.event_data`` will contain a :py:class:`~.NodeEditedArgs` object.
        """
        args = NodeEditedArgs(node=Node, new_text=NewText)
        self._trigger_event("nodeEdited", args)

    @override
    def nodeEditing(self, Node: XTreeNode) -> None:
        """
        This method is invoked from the TreeControl implementation when editing of Node is requested by calling XTreeControl.startEditingAtNode().

        Raises:
            com.sun.star.util.VetoException: ``VetoException``
        """
        self._trigger_event("nodeEditing", Node)

    @override
    def disposing(self, Source: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", Source)

    # endregion XTreeEditListener
