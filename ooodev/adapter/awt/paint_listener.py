from __future__ import annotations
from typing import TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.awt import XPaintListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import PaintEvent
    from com.sun.star.presentation import XSlideShowView
    from com.sun.star.awt import XWindow


class PaintListener(AdapterBase, XPaintListener):
    """
    makes it possible to receive paint events.

    See Also:
        `API XPaintListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XPaintListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XSlideShowView | XWindow | None = None
    ) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XSlideShowView, XWindow, optional): An UNO object that implements ``XSlideShowView`` or ``XWindow`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addPaintListener(self)

    # region XPaintListener
    @override
    def windowPaint(self, e: PaintEvent) -> None:
        """
        Is invoked when a region of the window became invalid, e.g.
        when another window has been moved away.
        """
        self._trigger_event("windowPaint", e)

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

    # endregion XPaintListener
