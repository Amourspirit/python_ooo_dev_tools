from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.util import XModifyListener
from com.sun.star.util import XModifyBroadcaster
from com.sun.star.lang import XComponent

from ooodev.utils import lo as mLo

from ..adapter_base import AdapterBase, GenericArgs as GenericArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class ModifyListener(AdapterBase, XModifyListener):
    """
    makes it possible to receive events when a model object changes.

    See Also:
        `API XModifyListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XModifyListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, doc: XComponent | None = None) -> None:
        """
        Constructor

        Args:
            doc (XComponent, Optional): Office Document. If document is passed then ``ModifyListener`` instance
                is automatically added.
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)
        if doc is None:
            return

        mb = mLo.Lo.qi(XModifyBroadcaster, doc, True)
        mb.addModifyListener(self)

    def modified(self, event: EventObject) -> None:
        """
        Is called when something changes in the object.

        Due to such an event, it may be necessary to update views or controllers.

        The source of the event may be the content of the object to which the listener
        is registered.
        """
        self._trigger_event("modified", event)

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
