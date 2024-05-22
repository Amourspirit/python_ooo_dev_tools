from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XCommandProcessor2

from .command_processor_partial import CommandProcessorPartial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class CommandProcessor2Partial(CommandProcessorPartial):
    """
    Partial Class XCommandProcessor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCommandProcessor2, interface: UnoInterface | None = XCommandProcessor2) -> None:
        """
        Constructor

        Args:
            component (XCommandProcessor2): UNO Component that implements ``com.sun.star.ucb.XCommandProcessor2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCommandProcessor2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        CommandProcessorPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XCommandProcessor2
    def release_command_identifier(self, cmd_id: int) -> None:
        """
        Releases a command identifier obtained through ``XCommandProcessor.createCommandIdentifier()`` when it is no longer used.

        After this call the command identifier cannot be used any longer in calls to ``XCommandProcessor.execute()`` and ``XCommandProcessor.abort()``.
        (But it can happen that a call to ``XCommandProcessor.createCommandIdentifier()`` reuses this identifier.)

        Args:
            cmd_id (int): A command identifier obtained through ``XCommandProcessor.createCommandIdentifier()``.
                If the identifier is zero, the request is silently ignored; but if the identifier is invalid
                (not obtained via ``XCommandProcessor.createCommandIdentifier()`` or already handed to ``XCommandProcessor2.releaseCommandIdentifier()`` before), the behavior is undefined.
        """
        self.__component.releaseCommandIdentifier(cmd_id)

    # endregion XCommandProcessor2
