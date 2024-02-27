from __future__ import annotations
from typing import TYPE_CHECKING, Tuple

import uno
from com.sun.star.beans import XPropertiesChangeListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.beans import PropertyChangeEvent


class PropertiesChangeListener(AdapterBase, XPropertiesChangeListener):
    """
    Receives events which get fired whenever a bound property is changed.

    See Also:
        `API XPropertiesChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertiesChangeListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XPropertiesChangeListener
    def propertiesChange(self, event: Tuple[PropertyChangeEvent, ...]) -> None:
        """
        Gets called when bound properties are changed.
        """
        self._trigger_event("propertiesChange", event)

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

    # endregion XPropertiesChangeListener
