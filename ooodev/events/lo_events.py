# coding: utf-8
"""
This module is for the purpose of sharing events between classes.
"""
from __future__ import annotations
from typing import Callable, cast, Dict, List
from .args.event_args import EventArgs

class LoEvents(object):
    """Static Class for sharing events among classes and functions."""
    _instance = None
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(LoEvents, cls).__new__(cls, *args, **kwargs)
            cls._instance._callbacks = None
        return cls._instance

    def on(self, event_name: str, callback:Callable[[object, EventArgs], None]):
        """
        Registers an event

        Args:
            event_name (str): Uniquie event name
            callback (Callable[[object, EventArgs], None]): Callback function
        """
        if self._callbacks is None:
            self._callbacks = {}

        if event_name not in self._callbacks:
            self._callbacks[event_name] = [callback]
        else:
            self._callbacks[event_name].append(callback)
    
    def remove(self, event_name: str, callback:Callable[[object, EventArgs], None]) -> bool:
        """
        Removes an event callback

        Args:
            event_name (str): Uniquie event name
            callback (Callable[[object, EventArgs], None]): Callback function

        Returns:
            bool: True if callback has been removed; Otherwise, False.
                False means the callback was not found.
        """
        if self._callbacks is None:
            return False
        result = False
        if event_name in self._callbacks:
            cb =cast(Dict[str, List[Callable[[object, EventArgs], None]]], self._callbacks)
            try:
                cb["event_name"].remove(callback)
                result = True
            except ValueError:
                pass
        return result


    def trigger(self, event_name: str, event_args: EventArgs):
        """
        Trigger event(s) for a given name.

        Args:
            event_name (str): Name of event to trigger/
            event_args (EventArgs): Event args passed to the callback for trigger.
        """
        if self._callbacks is not None and event_name in self._callbacks:
            for callback in self._callbacks[event_name]:
                if event_args is not None:
                    event_args._event_name = event_name
                callback(event_args.source, event_args)