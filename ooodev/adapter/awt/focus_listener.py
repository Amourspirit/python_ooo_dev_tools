from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ...events.args.event_args import EventArgs as EventArgs
from ..adapter_base import AdapterBase, GenericArgs as GenericArgs

from com.sun.star.awt import XFocusListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import FocusEvent


class FocusListener(AdapterBase, XFocusListener):
    """
    Makes it possible to receive keyboard focus events.

    The window which has the keyboard focus is the window which gets the keyboard events.

    See Also:
        `API XFocusListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XFocusListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XFocusListener

    def focusGained(self, event: FocusEvent) -> None:
        """
        is invoked when a window gains the keyboard focus.
        """
        self._trigger_event("focusGained", event)

    def focusLost(self, event: FocusEvent) -> None:
        """
        is invoked when a window loses the keyboard focus.
        """
        self._trigger_event("focusGained", event)

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

    # endregion XFocusListener
