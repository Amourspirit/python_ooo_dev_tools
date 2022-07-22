# coding: utf-8
"""
This module is for the purpose of sharing events between classes.
"""
from __future__ import annotations
import contextlib
from weakref import ref, ReferenceType, proxy
from typing import Dict, List, NamedTuple, Generator
from .args.event_args import EventArgs

from ..utils.type_var import EventCallback as EventCallback
from ..proto import event_observer
from . import event_singleton


class EventArg(NamedTuple):
    """
    Event Arg for passing event to :py:func:`~.event_ctx`
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
        self._callbacks = None

    def on(self, event_name: str, callback: EventCallback):
        """
        Registers an event

        Args:
            event_name (str): Unique event name
            callback (Callable[[object, EventArgs], None]): Callback function
        """
        if self._callbacks is None:
            self._callbacks: Dict[str, List[ReferenceType[EventCallback]]] = {}

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
            bool: True if callback has been removed; Otherwise, False.
            False means the callback was not found.
        """
        if self._callbacks is None:
            return False
        result = False
        if event_name in self._callbacks:
            # cb = cast(Dict[str, List[EventCallback]], self._callbacks)
            try:
                self._callbacks[event_name].remove(ref(callback))
                result = True
            except ValueError:
                pass
        return result

    def trigger(self, event_name: str, event_args: EventArgs):
        """
        Trigger event(s) for a given name.

        Args:
            event_name (str): Name of event to trigger
            event_args (EventArgs): Event args passed to the callback for trigger.

        Note:
            Events are removed automatically when they are out of scope.
        """

        if self._callbacks is not None and event_name in self._callbacks:
            cleanup = None
            for i, callback in enumerate(self._callbacks[event_name]):
                if callback() is None:
                    if cleanup is None:
                        cleanup = []
                    cleanup.append(i)
                    continue
                if event_args is not None:
                    event_args._event_name = event_name
                if callable(callback()):
                    try:
                        callback()(event_args.source, event_args)
                    except AttributeError:
                        # event_arg is None
                        callback()(self, None)
            if cleanup is not None and len(cleanup) > 0:
                cleanup.reverse()
                for i in cleanup:
                    self._callbacks[event_name].pop(i)
                if len(self._callbacks[event_name]) == 0:
                    del self._callbacks[event_name]


class Events(_event_base):
    """Static Class for sharing events among classes and functions."""

    def __init__(self) -> None:
        super().__init__()
        # register wih LoEvents so this instance get triggered when LoEvents() are triggered.
        LoEvents().add_observer(self)


class LoEvents(_event_base):
    """Singleton Class for ODEV global events."""

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(LoEvents, cls).__new__(cls, *args, **kwargs)
            cls._instance._callbacks = None
            cls._instance._observers: List[ReferenceType[event_observer.EventObserver]] = None
            # register wih _Events so this instance get triggered when _Events() are triggered.
            event_singleton._Events().add_observer(cls._instance)
        return cls._instance

    def __init__(self) -> None:
        pass

    def add_observer(self, *args: event_observer.EventObserver) -> None:
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

    def trigger(self, event_name: str, event_args: EventArgs):
        super().trigger(event_name, event_args)
        self._update_observers(event_name, event_args)

    def _update_observers(self, event_name: str, event_args: EventArgs) -> None:
        if self._observers is not None:
            cleanup = None
            for i, observer in enumerate(self._observers):
                if observer() is None:
                    if cleanup is None:
                        cleanup = []
                    cleanup.append(i)
                    continue
                observer().trigger(event_name=event_name, event_args=event_args)
            if cleanup is not None and len(cleanup) > 0:
                # reverse list to allow removing form highest to lowest to avoid errors
                cleanup.reverse()
                for i in cleanup:
                    self._observers.pop(i)


@contextlib.contextmanager
def event_ctx(*args: EventArg) -> Generator[event_observer.EventObserver, None, None]:
    """
    Event context manager.

    This manager adds and removes events.

    Parameters:
        args (EventArg): One or more EventArgs to add.

    Yields:
        Generator[EventObserver, None, None]: events
    """
    try:
        # yields a weakref.proxy obj
        # wekaref is dead as soon as e_obj is set to none.
        e_obj = Events()  # automatically adds itself as an observer to LoEvents()
        for arg in args:
            e_obj.on(arg.name, arg.callback)
        yield proxy(e_obj)
    except Exception:
        raise
    finally:
        e_obj = None
        _ = None  # just to make sure _ is not a ref to e_obj


def is_meth_event(source: str, meth: callable) -> bool:
    """
    Gets if event source is the same as meth.
    This method for for core events.

    Args:
        source (str): source as str
        meth (callable): method to test.

    Returns:
        bool: True if event is rased by meth; Otherwise; False
    """
    try:
        return source == meth.__qualname__
    except Exception:
        pass
    return False
