from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Tuple
import uno

from ooodev.mock.mock_g import DOCS_BUILDING

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from com.sun.star.beans import PropertyValue
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.events.args.event_args import EventArgs
    from ooodev.events.args.key_val_args import KeyValArgs
    from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
    from ooodev.format.inner.kind.format_kind import FormatKind
    from ooodev.proto.event_observer import EventObserver
    from ooodev.utils.type_var import EventCallback
else:
    Protocol = object
    PropertyValue = Any
    CancelEventArgs = Any
    EventArgs = Any
    KeyValArgs = Any
    KeyValCancelArgs = Any
    FormatKind = Any
    EventObserver = Any


class StyleT(Protocol):
    """Style base Protocol"""

    # region Update Methods
    def set_update_obj(self, obj: Any) -> None:
        """
        Sets the update object for the style instance.

        Args:
            obj (Any): Object used to apply style to when update is called.
        """
        ...

    def get_update_obj(self) -> Any:
        """
        Gets the update object for the style instance.
        """
        ...

    def has_update_obj(self) -> bool:
        """Gets if the update object is set. for the style instance."""
        ...

    @overload
    def update(self) -> bool:
        """
        Applies the style to the update object.

        Returns:
            bool: Returns ``True`` if the update object is set and the style is applied; Otherwise, ``False``.
        """
        ...

    @overload
    def update(self, **kwargs: Any) -> bool:
        """
        Applies the style to the update object.

        Returns:
            bool: Returns ``True`` if the update object is set and the style is applied; Otherwise, ``False``.
            **kwargs: Expandable list of key value pairs that may be used in child classes.
        """
        ...

    # endregion Update Methods

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        ...

    def apply(self, obj: Any, **kwargs: Any) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.
            **kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        ...

    def add_event_observer(self, *args: EventObserver) -> None:
        """
        Adds observers that gets their ``trigger`` method called when this class ``trigger`` method is called.

        Parameters:
            args (EventObserver): One or more observers to add.

        Returns:
            None:
        """
        ...

    def add_event_listener(self, event_name: str, callback: EventCallback) -> None:
        """
        Add an event listener to current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:

        See Also:
            - :py:class:`~.format_named_event.FormatNamedEvent`
            - :py:meth:`.remove_event_listener`

        Note:
            This method is generally only used in the context of child classes.
        """
        ...

    def remove_event_listener(self, event_name: str, callback: EventCallback) -> None:
        """
        Remove an event listener from current instance.

        Args:
            event_name (str): Event Name.
            callback (EventCallback): Callback of the event listener.

        Returns:
            None:

        See Also:
            - :py:class:`~.format_named_event.FormatNamedEvent`
            - :py:meth:`.add_event_listener`

        Note:
            This method is generally only used in the context of child classes.
        """
        ...

    def remove_event_observer(self, observer: EventObserver) -> bool:
        """
        Removes an observer.

        Args:
            observer (EventObserver): Observers to remove.

        Returns:
            bool: ``True`` if observer has been removed; Otherwise, ``False``.

        .. versionadded:: 0.30.1
        """
        ...

    def support_service(self, *service: str) -> bool:
        """
        Gets if service is supported.

        Args:
            service: expandable list of service names of UNO services such as ``com.sun.star.text.TextFrame``.

        Returns:
            bool: ``True`` if service is supported; Otherwise, ``False``.
        """
        ...

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attributes
        """
        ...

    def get_props(self) -> Tuple[PropertyValue, ...]:
        """
        Gets instance properties

        Returns:
            Tuple[PropertyValue, ...]: Tuple of properties.
        """
        ...

    @overload
    def copy(self) -> Any: ...

    @overload
    def copy(self, **kwargs) -> Any: ...

    def backup(self, obj: Any) -> None:
        """
        Backs up Attributes that are to be changed by apply.

        If used method should be called before apply.

        Args:
            obj (Any): Object to backup properties from.

        Returns:
            None:
        """
        ...

    def restore(self, obj: Any, clear: bool = ...) -> None:
        """
        Restores ``obj`` properties from backed up setting if any exist.

        Restore can only be effective if ``backup()`` has be run before calling this method.

        Args:
            obj (Any): Object to restore properties on.
            clear (bool): Determines if backup is cleared after restore. Default ``False``

        Returns:
            None:
        """
        ...

    def on_property_set(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that is set.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        ...

    def on_property_set_error(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that fails to set.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        ...

    def on_property_backing_up(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be backup up during backup.

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args.
        """
        ...

    def on_property_backed_up(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers before each property that is about to be set during backup.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        ...

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers before each property that is about to be set during restore

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args
        """
        ...

    def on_property_restore_set(self, source: Any, event_args: KeyValArgs) -> None:
        """
        Triggers for each property that has been set during restore

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        ...

    def on_applying(self, source: Any, event_args: CancelEventArgs) -> None:
        """
        Triggers before style/format is applied.

        Args:
            source (Any): Event Source.
            event_args (CancelEventArgs): Event Args
        """
        ...

    def on_applied(self, source: Any, event_args: EventArgs) -> None:
        """
        Triggers after style/format is applied.

        Args:
            source (Any): Event Source.
            event_args (KeyValArgs): Event Args
        """
        ...

    # region Properties

    @property
    def prop_has_attribs(self) -> bool:
        """Gets If instance has any attributes set."""
        ...

    @property
    def prop_has_backup(self) -> bool:
        """Gets If instance has backup data set."""
        ...

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        ...

    @property
    def prop_parent(self) -> Any:
        """Gets Parent Class"""
        ...

    # endregion Properties
