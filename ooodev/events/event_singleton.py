# coding: utf-8
"""
Internal Module only! DO NOT use this module/class!

This module is for the purpose of sharing events between classes internally
"""
from __future__ import annotations
from typing import Callable
from .args.event_args import EventArgs

from . import lo_events as mLoEvents

class Events(object):
    """Static Class for sharing events among internal classes. DO NOT USE!"""
    _instance = None
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(Events, cls).__new__(cls, *args, **kwargs)
            cls._instance.callbacks = None
        return cls._instance

    def on(self, event_name: str, callback:Callable[[object, EventArgs], None]):
        """
        Registers an event

        Args:
            event_name (str): Uniquie event name
            callback (Callable[[object, EventArgs], None]): Callback function
        """
        if self.callbacks is None:
            self.callbacks = {}

        if event_name not in self.callbacks:
            self.callbacks[event_name] = [callback]
        else:
            self.callbacks[event_name].append(callback)

    def trigger(self, event_name: str, event_args: EventArgs):
        """
        Trigger event(s) for a given name.

        Args:
            event_name (str): Name of event to trigger/
            event_args (EventArgs): Event args passed to the callback for trigger.
        """
        if self.callbacks is not None and event_name in self.callbacks:
            for callback in self.callbacks[event_name]:
                if event_args is not None:
                    event_args.event_name = event_name
                callback(event_args.source, event_args)
        # Trigger events on class that is designed for end users to use.
        # LoEventsis the class that end users will subscripe to an not this Events() class.
        mLoEvents.LoEvents().trigger(event_name=event_name, event_args=event_args)