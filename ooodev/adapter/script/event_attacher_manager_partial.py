from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.script import XEventAttacherManager

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from com.sun.star.script import XScriptListener
    from com.sun.star.script import ScriptEventDescriptor  # struct
    from ooodev.utils.type_var import UnoInterface


class EventAttacherManagerPartial:
    """
    Partial class for XEventAttacherManager.
    """

    def __init__(
        self, component: XEventAttacherManager, interface: UnoInterface | None = XEventAttacherManager
    ) -> None:
        """
        Constructor

        Args:
            component (XEventAttacherManager): UNO Component that implements ``com.sun.star.container.XEventAttacherManager`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEventAttacherManager``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XEventAttacherManager
    def add_script_listener(self, listener: XScriptListener) -> None:
        """
        Adds an ``XScriptListener`` that will be notified when an event takes place.

        Args:
            listener (XScriptListener): The listener to be added.
        """
        self.__component.addScriptListener(listener)

    def attach(self, idx: int, obj: XInterface, helper: Any) -> None:
        """
        Attaches all the ScriptEvents which are registered for the given index to the given object.
        """
        self.__component.attach(idx, obj, helper)

    def detach(self, idx: int, obj: XInterface) -> None:
        """
        Detaches all the ScriptEvents which are registered for the given index from the given object.
        """
        self.__component.detach(idx, obj)

    def get_script_events(self, idx: int) -> Tuple[ScriptEventDescriptor, ...]:
        """
        Returns all the ScriptEvents which are registered for the given index.
        """
        return self.__component.getScriptEvents(idx)

    def insert_entry(self, idx: int) -> None:
        """
        Inserts a new entry into the event attacher manager.
        """
        self.__component.insertEntry(idx)

    def register_script_event(self, idx: int, event: ScriptEventDescriptor) -> None:
        """
        Registers a ScriptEvent for the given index.
        """
        self.__component.registerScriptEvent(idx, event)

    def register_script_events(self, idx: int, events: Tuple[ScriptEventDescriptor, ...]) -> None:
        """
        Registers a ScriptEvent for the given index.
        """
        self.__component.registerScriptEvents(idx, events)

    def remove_entry(self, idx: int) -> None:
        """
        Removes the entry at the given position.
        """
        self.__component.removeEntry(idx)

    def remove_script_listener(self, listener: XScriptListener) -> None:
        """
        Removes an ``XScriptListener`` that was added with ``addScriptListener``.
        """
        self.__component.removeScriptListener(listener)

    def revoke_script_event(self, idx: int, listen_type: str, event_method: str, remove_listener_param: str) -> None:
        """
        Revokes the registration of an event.

        The parameters ``listen_type`` and ``event_method`` are equivalent to the first two members of the ScriptEventDescriptor used to register events.
        """
        self.__component.revokeScriptEvent(idx, listen_type, event_method, remove_listener_param)

    def revoke_script_events(self, idx: int) -> None:
        """
        Revokes all events which are registered for the given index.

        If the events at this index have been attached to any object, they are detached automatically.
        """
        self.__component.revokeScriptEvents(idx)

    # endregion XEventAttacherManager
