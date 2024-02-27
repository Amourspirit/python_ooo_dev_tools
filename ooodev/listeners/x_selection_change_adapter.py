from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.mock import mock_g

if mock_g.DOCS_BUILDING:
    from ooodev.mock import unohelper
else:
    import unohelper

from com.sun.star.view import XSelectionChangeListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class XSelectionChangeAdapter(unohelper.Base, XSelectionChangeListener):  # type: ignore
    """
    Makes it possible to receive an event when the current selection changes.

    This class is meant a parent class.
    """

    def selectionChanged(self, event: EventObject) -> None:
        """
        is called when the selection changes.

        You can get the new selection via XSelectionSupplier from ``com.sun.star.lang.EventObject.Source``.
        """
        pass

    def disposing(self, event: EventObject) -> None:
        """
        gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at XComponent.
        """
        # from com.sun.star.lang.XEventListener
        pass
