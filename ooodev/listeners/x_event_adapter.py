from __future__ import annotations
from typing import TYPE_CHECKING
from ..mock import mock_g

if mock_g.DOCS_BUILDING:
    from ..mock import unohelper
else:
    import unohelper

from com.sun.star.lang import XEventListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class XEventAdapter(unohelper.Base, XEventListener):
    """
    XEventListener implementation

    See Also:
        `API XEventListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XEventListener.html>`_
    """

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
