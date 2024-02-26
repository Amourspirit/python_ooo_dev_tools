from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.util import XChangesListener
from com.sun.star.util import XChangesNotifier

from ooodev.loader import lo as mLo

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.util import ChangesEvent


class ChangesListener(AdapterBase, XChangesListener):
    """
    Receives events from batch change broadcaster objects.

    See Also:
        `API XChangesListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XChangesListener.html>`_

    .. versionadded:: 0.19.1
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XChangesNotifier | None = None, **kwargs
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XChangesNotifier, optional): An UNO object that implements the ``XChangesNotifier`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        # doc parameter was renamed to subscriber in version 0.13.6
        if subscriber is None:
            doc = kwargs.get("doc", None)
        else:
            doc = subscriber
        if doc is None:
            return

        mb = mLo.Lo.qi(XChangesNotifier, doc, True)
        mb.addChangesListener(self)

    def changesOccurred(self, event: ChangesEvent) -> None:
        """
        Is invoked when a batch of changes occurred.
        """
        self._trigger_event("changesOccurred", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets invoked when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)
