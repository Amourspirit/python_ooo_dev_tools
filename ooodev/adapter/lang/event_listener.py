from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib

import uno  # pylint: disable=unused-import
from com.sun.star.lang import XEventListener

# pylint: disable=useless-import-alias
# pylint: disable=unused-import
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class EventListener(AdapterBase, XEventListener):
    """
    Implementation of base interface

    This listener must be attached manually.

    See Also:
        - :ref:`ch04_detect_shutdown_via_listener`
        - `API XEventListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XEventListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: Any = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (Any, optional): An UNO object that has a ``addAdjustmentListener()`` Method.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            # several object such as Scrollbar and SpinValue
            # There is no common interface for this, so we have to try them all.
            with contextlib.suppress(AttributeError):
                subscriber.addEventListener(self)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        self._trigger_event("disposing", event)
