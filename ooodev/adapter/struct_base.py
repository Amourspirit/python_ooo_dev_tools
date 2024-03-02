from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Generic, TypeVar
import uno
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.args.key_val_args import KeyValArgs

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.

_T = TypeVar("_T")


class StructBase(ComponentBase, Generic[_T]):
    """
    Struct Base Class

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: _T, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Struct object.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        ComponentBase.__init__(self, component)
        self._event_provider = event_provider
        self._prop_name = prop_name

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    def _get_on_changed_event_name(self) -> str:
        raise NotImplementedError

    def _get_on_changing_event_name(self) -> str:
        raise NotImplementedError

    def _get_prop_name(self) -> str:
        return self._prop_name

    def _on_property_changing(self, event_args: KeyValCancelArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event(self._get_on_changing_event_name(), event_args)

    def _on_property_changed(self, event_args: KeyValArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event(self._get_on_changed_event_name(), event_args)

    def _copy(self, src: _T | None = None) -> _T:
        raise NotImplementedError

    def copy(self) -> _T:
        """
        Makes a copy of the Struct.

        Returns:
            BorderLine: Copied Struct.
        """
        return self._copy()

    def _set_component(self, struct: _T, copy: bool = False) -> None:
        # pylint: disable=no-member
        if copy:
            self._ComponentBase__set_component(self._copy(src=struct))  # type: ignore
        else:
            self._ComponentBase__set_component(struct)  # type: ignore

    def _get_component(self, copy: bool = False) -> _T:
        # pylint: disable=no-member
        if copy:
            return cast(_T, self._copy(self._ComponentBase__get_component()))  # type: ignore
        else:
            return cast(_T, self._ComponentBase__get_component())  # type: ignore

    def __repr__(self) -> str:
        return f"{self.__class__.__name__} {repr(self.component)}"

    @property
    def component(self) -> _T:
        """Struct Component"""
        # pylint: disable=no-member
        return self._get_component()

    @component.setter
    def component(self, value: _T) -> None:
        # pylint: disable=no-member
        self._set_component(value, True)

    def _trigger_cancel_event(self, struct_prop_name: str, old_value: Any, value: Any) -> KeyValCancelArgs:
        """
        Triggers the Property Changing Event.

        Args:
            struct_prop_name (str): This must be the name of the struct property.
            old_value (Any): The value of the property before it is changed.
            value (Any): The new value to assign to the property if not canceled.

        Returns:
            KeyValCancelArgs: Event args for before the property is changed.
        """
        event_args = KeyValCancelArgs(
            source=self,
            key=struct_prop_name,
            value=value,
        )
        event_args.event_data = {
            "old_value": old_value,
            "prop_name": self._get_prop_name(),
        }
        self._on_property_changing(event_args)
        return event_args

    def _trigger_done_event(self, cargs: KeyValCancelArgs) -> KeyValArgs | None:
        """
        Triggers Property Changed Event.

        If the event is not cancelled, the property is changed and the event is triggered.

        This method also assigns the new value to the struct if the event is not cancelled.

        Args:
            cargs (KeyValCancelArgs): The event args for before the property is changed.

        Returns:
            KeyValArgs | None: Event args for after the property is changed or ``None`` if the event was cancelled.
        """
        if not cargs.cancel:
            struct = self._copy()
            setattr(struct, cargs.key, cargs.value)
            self._set_component(struct)
            kv_args = KeyValArgs.from_args(cargs)  # type: ignore
            self._on_property_changed(kv_args)
            return kv_args
        return None
