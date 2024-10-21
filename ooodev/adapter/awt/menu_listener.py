from __future__ import annotations
from typing import TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.awt import XMenuListener
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import MenuEvent
    from com.sun.star.awt import XMenu


class MenuListener(AdapterBase, XMenuListener):
    """
    Makes it possible to receive menu events on a window.

    See Also:
        `API XMenuListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMenuListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XMenu | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XMenu, optional): An UNO object that implements the ``XMenu`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addMenuListener(self)

    # region XMenuListener
    @override
    def itemActivated(self, aEvent: MenuEvent) -> None:
        """
        Invoked when a menu is activated.
        """
        self._trigger_event("itemActivated", aEvent)

    @override
    def itemDeactivated(self, aEvent: MenuEvent) -> None:
        """
        Invoked when a menu is deactivated.
        """
        self._trigger_event("itemDeactivated", aEvent)

    @override
    def itemHighlighted(self, aEvent: MenuEvent) -> None:
        """
        Invoked when a menu item is highlighted.
        """
        self._trigger_event("itemHighlighted", aEvent)

    @override
    def itemSelected(self, aEvent: MenuEvent) -> None:
        """
        Invoked when a menu item is selected.
        """
        self._trigger_event("itemSelected", aEvent)

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

    # endregion XMenuListener
