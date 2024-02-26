from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt.grid import XGridColumnListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt.grid import GridColumnEvent


class GridColumnListener(AdapterBase, XGridColumnListener):
    """
    An instance of this interface is used by the XGridColumnModel to get notifications about column model changes.

    **since**

        OOo 3.3

    See Also:
        `API XGridColumnListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1grid_1_1XGridColumnListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XGridColumnListener
    def columnChanged(self, event: GridColumnEvent) -> None:
        """
        Invoked after a column was modified.
        """
        self._trigger_event("columnChanged", event)

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

    # endregion XGridColumnListener
