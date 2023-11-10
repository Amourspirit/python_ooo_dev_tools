from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ...events.args.event_args import EventArgs as EventArgs
from ..adapter_base import AdapterBase, GenericArgs as GenericArgs

from com.sun.star.awt import XPaintListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import PaintEvent


class PaintListener(AdapterBase, XPaintListener):
    """
    makes it possible to receive paint events.

    See Also:
        `API XPaintListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XPaintListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XPaintListener
    def windowPaint(self, event: PaintEvent) -> None:
        """
        Is invoked when a region of the window became invalid, e.g.
        when another window has been moved away.
        """
        self._trigger_event("windowPaint", event)

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

    # endregion XPaintListener
