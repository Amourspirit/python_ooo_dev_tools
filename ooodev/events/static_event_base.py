# coding: utf-8
from typing import Callable, Any
from .event_args import EventArgs
class StaticEventBase:
    callbacks = None

    @classmethod
    def on(cls, event_name: str, callback:Callable[[object, EventArgs], None]):
        if cls.callbacks is None:
            cls.callbacks = {}

        if event_name not in cls.callbacks:
            cls.callbacks[event_name] = [callback]
        else:
            cls.callbacks[event_name].append(callback)

    @classmethod
    def trigger(cls, event_name: str, event_args: EventArgs):
        if cls.callbacks is not None and event_name in cls.callbacks:
            for callback in cls.callbacks[event_name]:
                if event_args is not None:
                    event_args.event_name = event_name
                callback(cls, event_args)