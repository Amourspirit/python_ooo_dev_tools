from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.awt import XAdjustmentListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import AdjustmentEvent


class AdjustmentListener(AdapterBase, XAdjustmentListener):
    """
    Makes it possible to receive action events.

    See Also:
        `API XAdjustmentListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XAdjustmentListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: Any = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (Any, optional): An UNO object that has a ``addAdjustmentListener()`` Method.
                If passed in then this listener instance is automatically added to it.
                Valid objects are: Scrollbar, SpinValue, or any other UNO object that has ``addAdjustmentListener()`` method.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            # several object such as Scrollbar and SpinValue
            # There is no common interface for this, so we have to try them all.
            with contextlib.suppress(AttributeError):
                subscriber.addAdjustmentListener(self)

    # region XAdjustmentListener
    def adjustmentValueChanged(self, event: AdjustmentEvent) -> None:
        """
        Event is invoked when the adjustment has changed.
        """
        self._trigger_event("adjustmentValueChanged", event)

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

    # endregion XAdjustmentListener
