from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XCommandProcessor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ucb import Command  # struct
    from com.sun.star.ucb import XCommandEnvironment
    from ooodev.utils.type_var import UnoInterface


class CommandProcessorPartial:
    """
    Partial Class XCommandProcessor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCommandProcessor, interface: UnoInterface | None = XCommandProcessor) -> None:
        """
        Constructor

        Args:
            component (XCommandProcessor): UNO Component that implements ``com.sun.star.ucb.XCommandProcessor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCommandProcessor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCommandProcessor
    def abort(self, cmd_id: int) -> None:
        """
        ends the command associated with the given id.

        Not every command can be aborted. It's up to the implementation to decide whether this method will actually end the processing of the command or simply do nothing.
        """
        self.__component.abort(cmd_id)

    def create_command_identifier(self) -> int:
        """
        Creates a unique identifier for a command.

        This identifier can be used to abort the execution of the command associated with that identifier.
        Note that it is generally not necessary to obtain a new id for each command, because commands are executed synchronously.
        So the id for a command is valid again after a command previously associated with this id has finished.
        In fact you only should get one identifier per thread and assign it to every command executed by that thread.

        Also, after a call to ``abort()``, an identifier should not be used any longer (and instead be released by a call
        to ``XCommandProcessor2.releaseCommandIdentifier()``), because it may well abort all further calls to ``execute()``.

        To avoid ever-increasing resource consumption, the identifier should be released via ``XCommandProcessor2.releaseCommandIdentifier()`` when it is no longer used.
        """
        return self.__component.createCommandIdentifier()

    def execute(self, cmd: Command, cmd_id: int, env: XCommandEnvironment) -> Any:
        """
        Executes a command.

        Common command definitions can be found in the specification of the service Content.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
            CommandAbortedException: ``CommandAbortedException``
        """
        return self.__component.execute(cmd, cmd_id, env)

    # endregion XCommandProcessor
