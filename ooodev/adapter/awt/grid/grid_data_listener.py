from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt.grid import XGridDataListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt.grid import GridDataEvent


class GridDataListener(AdapterBase, XGridDataListener):
    """
    An instance of this interface is used by the XGridDataModel to get notifications about data model changes.

    Usually you must not implement this interface yourself, but you must notify it correctly if you implement
    the XGridDataModel yourself.

    **since**

        OOo 3.3

    See Also:
        `API XGridDataListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1grid_1_1XGridDataListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XGridDataListener
    def dataChanged(self, event: GridDataEvent) -> None:
        """
        Event is invoked when existing data in a grid control's data model has been modified.
        """
        self._trigger_event("dataChanged", event)

    def rowHeadingChanged(self, event: GridDataEvent) -> None:
        """
        Event is invoked when the title of one or more rows changed.
        """
        self._trigger_event("rowHeadingChanged", event)

    def rowsInserted(self, event: GridDataEvent) -> None:
        """
        is called when one or more rows of data have been inserted into a grid control's data model.
        """
        self._trigger_event("rowsInserted", event)

    def rowsRemoved(self, event: GridDataEvent) -> None:
        """
        is called when one or more rows of data have been removed from a grid control's data model.
        """
        self._trigger_event("rowsRemoved", event)

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

    # endregion XGridDataListener
