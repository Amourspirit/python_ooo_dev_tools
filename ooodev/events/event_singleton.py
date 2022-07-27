# coding: utf-8
"""
Internal Module only! DO NOT use this module/class!

This module is for the purpose of sharing events between classes internally
"""
from __future__ import annotations
from weakref import ref, ReferenceType
from .args.event_args import EventArgs
from typing import List, Dict
from ..utils import type_var
from ..proto import event_observer


class _Events(object):
    """
    Singleton Class for sharing events among internal classes. DO NOT USE!

    Use: lo_events.LoEvents for global events. Use lo_events.Events for locally scoped events.
    """

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(_Events, cls).__new__(cls, *args, **kwargs)
            cls._instance._callbacks = None
            cls._instance._observers: List[ReferenceType[event_observer.EventObserver]] = None
        return cls._instance

    def on(self, event_name: str, callback: type_var.EventCallback):
        """
        Registers an event

        Args:
            event_name (str): Unique event name
            callback (Callable[[object, EventArgs], None]): Callback function
        """
        if self._callbacks is None:
            self._callbacks: Dict[str, List[ReferenceType[type_var.EventCallback]]] = {}

        if event_name not in self._callbacks:
            self._callbacks[event_name] = [ref(callback)]
        else:
            self._callbacks[event_name].append(ref(callback))

    def trigger(self, event_name: str, event_args: EventArgs, *args, **kwargs) -> None:
        """
        Trigger event(s) for a given name.

        Args:
            event_name (str): Name of event to trigger/
            event_args (EventArgs): Event args passed to the callback for trigger.
            args (Any, optional): Optional positional args to pass to callback
            kwargs (Any, optional): Optional keyword args to pass to callback
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
                    if event_args.event_source is None:
                        event_args._event_source = self
                if callable(callback()):
                    try:
                        callback()(event_args.source, event_args, *args, **kwargs)
                    except AttributeError:
                        # event_arg is None
                        callback()(self, None)
            if cleanup is not None and len(cleanup) > 0:
                # reverse list to allow removing form highest to lowest to avoid errors
                cleanup.reverse()
                for i in cleanup:
                    self._callbacks[event_name].pop(i)
                if len(self._callbacks[event_name]) == 0:
                    del self._callbacks[event_name]
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

    def add_observer(self, *args: event_observer.EventObserver) -> None:
        """
        Adds observers that gets their ``trigger`` method called when this class ``trigger`` method is called.
        """
        if self._observers is None:
            self._observers = []
        for observer in args:
            self._observers.append(ref(observer))
