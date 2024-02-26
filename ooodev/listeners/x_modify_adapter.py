from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.util import XModifyListener

from ooodev.mock import mock_g

if mock_g.DOCS_BUILDING:
    from ooodev.mock import unohelper
else:
    import unohelper

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class XModifyAdapter(unohelper.Base, XModifyListener):  # type: ignore
    """
    makes it possible to receive events when a model object changes.

    This class is meant a parent class.
    """

    def modified(self, event: EventObject) -> None:
        """
        is called when something changes in the object.

        Due to such an event, it may be necessary to update views or controllers.

        The source of the event may be the content of the object to which the listener
        is registered.
        """
        pass

    def disposing(self, event: EventObject) -> None:
        """
        gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including XComponent.removeEventListener() ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at XComponent.
        """
        # from com.sun.star.lang.XEventListener
        pass
