# coding: utf-8
"""
This module is for the purpose of sharing events between classes.
"""
from __future__ import annotations
import contextlib
from weakref import ref, ReferenceType
from typing import Any, Dict, List, NamedTuple, Generator, Callable, Union, Tuple, TYPE_CHECKING
from ooodev.events import event_singleton
from ooodev.events.args.event_args_t import EventArgsT
from ooodev.utils.type_var import EventCallback as EventCallback
from ooodev.events.args.generic_args import GenericArgs as GenericArgs

# pylint: disable=protected-access

if TYPE_CHECKING:
    from ooodev.proto.event_observer import EventObserver


class EventArg(NamedTuple):
    """
    Event Args for passing event to :py:func:`~.event_ctx`
    """

    name: str
    """Event name"""
    callback: EventCallback
    """
    Event Callback
    """


class _event_base(object):
    """Base events class"""

    def __init__(self) -> None:
        self._callbacks: Dict[str, List[ReferenceType[EventCallback]]] | None = None
        self._observers: Union[List[ReferenceType[EventObserver]], None] = None

    def on(self, event_name: str, callback: EventCallback):
        """
        Registers an event

        Args:
            event_name (str): Unique event name
            callback (Callable[[object, EventArgs], None]): Callback function
        """
        if self._callbacks is None:
            self._callbacks = {}

        if event_name not in self._callbacks:
            self._callbacks[event_name] = [ref(callback)]
        else:
            self._callbacks[event_name].append(ref(callback))

    def remove(self, event_name: str, callback: EventCallback) -> bool:
        """
        Removes an event callback

        Args:
            event_name (str): Unique event name
            callback (Callable[[object, EventArgs], None]): Callback function

        Returns:
            bool: ``True`` if callback has been removed; Otherwise, ``False``.
                False means the callback was not found.
        """
        if self._callbacks is None:
            return False
        result = False
        if event_name in self._callbacks:
            # cb = cast(Dict[str, List[EventCallback]], self._callbacks)
            with contextlib.suppress(ValueError):
                self._callbacks[event_name].remove(ref(callback))
                result = True
        return result

    def has_event_name(self, event_name: str) -> bool:
        """
        Gets if event exists.

        Args:
            event_name (str): Event name

        Returns:
            bool: True if event exists; Otherwise, False
        """
        return False if self._callbacks is None else event_name in self._callbacks

    def has_event(self, event_name: str, callback: EventCallback) -> bool:
        """
        Gets if event exists.

        Args:
            event_name (str): Event name
            callback (EventCallback): Callback function

        Returns:
            bool: True if event exists; Otherwise, False
        """
        if self._callbacks is None:
            return False
        if callback is None:
            return False
        if event_name not in self._callbacks:
            return False
        return ref(callback) in self._callbacks[event_name]

    def _clear(self) -> None:
        if self._callbacks is not None:
            self._callbacks.clear()
        if self._observers is not None:
            self._observers.clear()

    def _set_event_args(self, event_name: str, event_args: EventArgsT) -> None:
        if event_args is None:
            return
        event_args._event_name = event_name
        event_args._event_source = self  # type: ignore

    def trigger(self, event_name: str, event_args: EventArgsT, *args, **kwargs):
        """
        Trigger event(s) for a given name.

        Args:
            event_name (str): Name of event to trigger
            event_args (EventArgsT): Event args passed to the callback for trigger.
            args (Any, optional): Optional positional args to pass to callback
            kwargs (Any, optional): Optional keyword args to pass to callback

        Note:
            Events are removed automatically when they are out of scope.
        """
        # sourcery skip: last-if-guard

        if self._callbacks is not None and event_name in self._callbacks:
            cleanup = None
            for i, callback in enumerate(self._callbacks[event_name]):
                if callback() is None:
                    if cleanup is None:
                        cleanup = []
                    cleanup.append(i)
                    continue
                self._set_event_args(event_name=event_name, event_args=event_args)
                if callable(callback()):
                    try:
                        callback()(event_args.source, event_args, *args, **kwargs)  # type: ignore
                    except AttributeError:
                        # event_arg is None
                        callback()(self, None)  # type: ignore
            if cleanup:
                cleanup.reverse()
                for i in cleanup:
                    self._callbacks[event_name].pop(i)
                if len(self._callbacks[event_name]) == 0:
                    del self._callbacks[event_name]

    def add_observer(self, *args: EventObserver) -> None:
        """
        Adds observers that gets their ``trigger`` method called when this class ``trigger`` method is called.

        Parameters:
            args (EventObserver): One or more observers to add.

        Returns:
            None:

        Note:
            Observers are removed automatically when they are out of scope.
        """
        if self._observers is None:
            self._observers = []
        for observer in args:
            self._observers.append(ref(observer))

    def _update_observers(self, event_name: str, event_args: EventArgsT) -> None:
        # sourcery skip: last-if-guard
        if self._observers is not None:
            cleanup = None
            for i, observer in enumerate(self._observers):
                if observer() is None:
                    if cleanup is None:
                        cleanup = []
                    cleanup.append(i)
                    continue
                observer().trigger(event_name=event_name, event_args=event_args)  # type: ignore
            if cleanup:
                # reverse list to allow removing form highest to lowest to avoid errors
                cleanup.reverse()
                for i in cleanup:
                    _ = self._observers.pop(i)

    def remove_observer(self, observer: EventObserver) -> bool:
        """
        Removes an observer.

        Args:
            observer (EventObserver): Observers to remove.

        Returns:
            bool: ``True`` if observer has been removed; Otherwise, ``False``.
        """

        if self._observers is None:
            return False
        with contextlib.suppress(Exception):
            self._observers.remove(ref(observer))
            return True
        return False


class Events(_event_base):
    """
    Class for sharing events among classes and functions.

    Note:
        If an events source is ``None`` then it is set to ``source`` or ``Events`` instance if ``source`` is None.
    """

    # Dev Notes:
    # Event callbacks are assigned to this class as a weak ref.
    # This is necessary; However, a side effect is class method cannot be assigned
    # as an event from class __init__. It is possible to assign a class static method from
    # __init__ but not a class method. Attempting to assign class method result with method
    # being out of scope before trigger is called on it.
    # Making an Events class with strong ref ( no weak ref ) and then assigning a class method
    # as a callback result in the class method being triggered even after the class instance is set
    # to none. In other words python does not release the object or callback because the strong ref Events class
    # is still holding on to it.
    # In short, do not change this class!

    def __init__(self, source: Any | None = None, trigger_args: GenericArgs | None = None) -> None:
        """
        Construct for Events

        Args:
            source (Any | None, optional): Source can be class or any object.
                The value of ``source`` is the value assigned to the ``EventArgs.event_source`` property.
                Defaults to current instance of this class.
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__()
        self._source = source
        self._t_args = trigger_args

        # register wih LoEvents so this instance get triggered when LoEvents() are triggered.
        # LoEvents().add_observer(self)  # type: ignore

    def clear(self) -> None:
        """
        Clears all events.

        .. versionadded:: 0.13.7
        """
        super()._clear()

    def trigger(self, event_name: str, event_args: EventArgsT):
        if self._t_args is None:
            super().trigger(event_name=event_name, event_args=event_args)
        else:
            super().trigger(event_name, event_args, *self._t_args.args, **self._t_args.kwargs)

        self._update_observers(event_name, event_args)

    def _set_event_args(self, event_name: str, event_args: EventArgsT) -> None:
        if event_args is None:
            return
        event_args._event_name = event_name
        event_args._event_source = self if self._source is None else self._source  # type: ignore
        if event_args.source is None:
            event_args.source = self if self._source is None else self._source


class LoEvents(_event_base):
    """Singleton Class for ODEV global events."""

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(LoEvents, cls).__new__(cls, *args, **kwargs)
            cls._instance._callbacks = None
            # cls._instance._observers: List[ReferenceType[event_observer.EventObserver]] | None = None
            cls._instance._observers = None
            # register wih _Events so this instance get triggered when _Events() are triggered.
            event_singleton._Events().add_observer(cls._instance)  # type: ignore
        return cls._instance

    def __init__(self) -> None:
        self._observers: Union[List[ReferenceType[EventObserver]], None]

    def add_observer(self, *args: EventObserver) -> None:
        """
        Adds observers that gets their ``trigger`` method called when this class ``trigger`` method is called.

        Parameters:
            args (EventObserver): One or more observers to add.

        Returns:
            None:

        Note:
            Observers are removed automatically when they are out of scope.
        """
        super().add_observer(*args)

    def trigger(self, event_name: str, event_args: EventArgsT):
        super().trigger(event_name, event_args)
        self._update_observers(event_name, event_args)

    def _update_observers(self, event_name: str, event_args: EventArgsT) -> None:
        # sourcery skip: last-if-guard
        if self._observers is not None:
            cleanup = None
            for i, observer in enumerate(self._observers):
                if observer() is None:
                    if cleanup is None:
                        cleanup = []
                    cleanup.append(i)
                    continue
                observer().trigger(event_name=event_name, event_args=event_args)  # type: ignore
            if cleanup:
                # reverse list to allow removing form highest to lowest to avoid errors
                cleanup.reverse()
                for i in cleanup:
                    _ = self._observers.pop(i)


class DummyEvents:
    """Dummy events class for ignoring events."""

    def __init__(self, *args, **kwargs) -> None:
        pass

    def on(self, event_name: str, callback: EventCallback) -> None:
        pass

    def remove(self, event_name: str, callback: EventCallback) -> bool:
        return True

    def trigger(self, event_name: str, event_args: EventArgsT, *args, **kwargs) -> None:
        pass


@contextlib.contextmanager
def event_ctx(
    *args: EventArg | Tuple[str, EventCallback],
    source: Any = None,
    trigger_args: GenericArgs | None = None,
    lo_observe: bool = False,
) -> Generator[EventObserver, None, None]:
    """
    Event context manager.

    This manager adds and removes events.

    Parameters:
        args (EventArg, Tuple[str, EventCallback], optional): One or more EventArgs to add.
        source (Any, optional): Source can be class or any object.
        trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.

    Yields:
        Generator[EventObserver, None, None]: events
    """
    # pylint: disable=no-member
    e_obj = None
    lo_inst = None
    try:
        e_obj = Events(
            source=source, trigger_args=trigger_args
        )  # automatically adds itself as an observer to LoEvents()
        for arg in args:
            if isinstance(arg, tuple):
                e_obj.on(arg[0], arg[1])
            else:
                e_obj.on(arg.name, arg.callback)
        # a proxy is not able to be added to the LoEvents() observer list
        # yield proxy(e_obj)
        if lo_observe:
            # pylint: disable=import-outside-toplevel
            from ooodev.loader.lo import Lo

            lo_inst = Lo.current_lo
            lo_inst.add_event_observers(e_obj)
        yield e_obj
    except Exception:
        raise
    finally:
        if e_obj is not None:
            e_obj.clear()
            if lo_inst is not None:
                # strictly speaking this is not necessary because Events will automatically remove
                # stale observers; However, it is good practice to remove observers when they are
                # no longer needed.
                lo_inst.remove_event_observer(e_obj)
        e_obj = None
        _ = None  # just to make sure _ is not a ref to e_obj


@contextlib.contextmanager
def observe_events(
    observer: EventObserver,
    events: EventObserver | None = None,
) -> Generator[EventObserver, None, None]:
    """
    Event Observer context manager.

    This manager adds and removes event observer.

    Parameters:
        observer (EventObserver): Observer to add.
        events (EventObserver, optional): Events to add observer to. Defaults to ``LoEvents()`` Singleton

    Yields:
        Generator[EventObserver, None, None]: events
    """
    try:
        if events is None:
            events = LoEvents()
        events.add_observer(observer)
        yield observer
    except Exception:
        raise
    finally:
        if events is not None:
            events.remove_observer(observer)


def is_meth_event(source: str, meth: Callable) -> bool:
    """
    Gets if event source is the same as meth.
    This method for for core events.

    Args:
        source (str): source as str
        meth (callable): method to test.

    Returns:
        bool: True if event is raised by meth; Otherwise; False
    """
    with contextlib.suppress(Exception):
        return source == meth.__qualname__
    return False
