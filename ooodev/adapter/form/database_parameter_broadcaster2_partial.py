from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.form import XDatabaseParameterBroadcaster2
from ooodev.adapter.form.database_parameter_broadcaster_partial import DatabaseParameterBroadcasterPartial

if TYPE_CHECKING:
    from com.sun.star.form import XDatabaseParameterListener
    from ooodev.utils.type_var import UnoInterface


class DatabaseParameterBroadcaster2Partial(DatabaseParameterBroadcasterPartial):
    """
    Partial class for Forms Component.
    """

    def __init__(
        self,
        component: XDatabaseParameterBroadcaster2,
        interface: UnoInterface | None = XDatabaseParameterBroadcaster2,
    ) -> None:
        """
        Constructor

        Args:
            component (XDatabaseParameterBroadcaster2): UNO Component that implements ``com.sun.star.form.XDatabaseParameterBroadcaster2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDatabaseParameterBroadcaster2``.
        """
        DatabaseParameterBroadcasterPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XDatabaseParameterBroadcaster2
    def add_database_parameter_listener(self, listener: XDatabaseParameterListener) -> None:
        """
        Registers an ``XDatabaseParameterListener``

        This method behaves exactly as the XDatabaseParameterBroadcaster.addParameterListener() method inherited from the base interface.
        """
        self.__component.addDatabaseParameterListener(listener)

    def remove_database_parameter_listener(self, listener: XDatabaseParameterListener) -> None:
        """
        revokes an XDatabaseParameterListener

        This method behaves exactly as the XDatabaseParameterBroadcaster.removeParameterListener() method inherited from the base interface.
        """
        self.__component.removeDatabaseParameterListener(listener)

    # endregion XDatabaseParameterBroadcaster2
