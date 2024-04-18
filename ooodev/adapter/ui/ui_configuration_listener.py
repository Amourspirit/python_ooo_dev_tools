from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.ui import XUIConfigurationListener
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.ui import XUIConfiguration
    from com.sun.star.lang import EventObject
    from com.sun.star.ui import ConfigurationEvent


class UIConfigurationListener(AdapterBase, XUIConfigurationListener):
    """
    Supplies information about changes of a user interface configuration manager.

    See Also:
        `API XUIConfigurationListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ui_1_1XUIConfigurationListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XUIConfiguration | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XUIConfiguration, optional): An UNO object that implements ``com.sun.star.ui.XUIConfiguration``.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addConfigurationListener(self)

    # region XUIConfigurationListener

    def elementInserted(self, event: ConfigurationEvent) -> None:
        """
        Invoked when a configuration has inserted an user interface element.
        """
        self._trigger_event("elementInserted", event)

    def elementRemoved(self, event: ConfigurationEvent) -> None:
        """
        is invoked when a configuration has removed an user interface element.
        """
        self._trigger_event("elementRemoved", event)

    def elementReplaced(self, event: ConfigurationEvent) -> None:
        """
        is invoked when a configuration has replaced an user interface element.
        """
        self._trigger_event("elementReplaced", event)

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

    # endregion XUIConfigurationListener
