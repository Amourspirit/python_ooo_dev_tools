from __future__ import annotations
from typing import TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.frame import XDispatchResultListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.frame import DispatchResultEvent


class DispatchResultListener(AdapterBase, XDispatchResultListener):
    """
    Listener for results of ``XNotifyingDispatch.dispatchWithNotification()``

    See Also:
        - `API XDispatchResultListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XDispatchResultListener.html>`_
        - :py:class:`ooodev.adapter.frame.notifying_dispatch_partial.NotifyingDispatchPartial`
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    @override
    def dispatchFinished(self, Result: DispatchResultEvent) -> None:
        """
        Indicates finished dispatch
        """
        self._trigger_event("dispatchFinished", Result)

    @override
    def disposing(self, Source: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", Source)
