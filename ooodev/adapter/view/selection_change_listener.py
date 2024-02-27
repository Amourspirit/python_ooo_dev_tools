from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XModel
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view import XSelectionSupplier

from ooodev.loader import lo as mLo

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class SelectionChangeListener(AdapterBase, XSelectionChangeListener):
    """
    makes it possible to receive events when a model object changes.

    See Also:
        - :ref:`ch25_listen_cell_select`
        - `API XSelectionChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XSelectionChangeListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, doc: Any | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            doc (Any, Optional): Office Document. If document is passed then ``SelectionChangeListener`` instance
                is automatically added.
        """
        super().__init__(trigger_args=trigger_args)
        if doc is None:
            return

        model = mLo.Lo.qi(XModel, doc)
        if model is None:
            mLo.Lo.print("Could not get model for doc")
            return

        supp = mLo.Lo.qi(XSelectionSupplier, model.getCurrentController())
        if supp is None:
            mLo.Lo.print("Could not attach selection change listener")
            return
        supp.addSelectionChangeListener(self)

    def selectionChanged(self, event: EventObject) -> None:
        """
        Is called when the selection changes.

        You can get the new selection via XSelectionSupplier from ``com.sun.star.lang.EventObject.Source``.
        """
        self._trigger_event("selectionChanged", event)

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
