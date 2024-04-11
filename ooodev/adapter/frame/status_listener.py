from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XStatusListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.frame import XDispatch
    from com.sun.star.lang import EventObject
    from com.sun.star.frame import FeatureStateEvent


class StatusListener(AdapterBase, XStatusListener):
    """
    Makes it possible to receive events when the state of a feature changes.

    Nobody guarantee any notification. Use combination of XNotifyingDispatch and XDispatchResultListener for that.

    See Also:
        `API XStatusListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStatusListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    def statusChanged(self, event: FeatureStateEvent) -> None:
        """
        Event is invoked when the status of the feature changes.
        """
        self._trigger_event("statusChanged", event)

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
