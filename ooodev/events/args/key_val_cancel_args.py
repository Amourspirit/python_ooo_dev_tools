# coding: utf-8
from __future__ import annotations
from typing import Any
from .key_val_args import AbstractKeyValArgs
from .cancel_event_args import AbstractCancelEventArgs


class AbstractKeyValueArgs(AbstractKeyValArgs, AbstractCancelEventArgs):
    __slots__ = ()

    def __init__(self, source: Any, key: str, value: Any, cancel=False) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            key (str): Key
            value (Any: Value
        """
        super().__init__(source=source, key=key, value=value)
        self.cancel = cancel


class KeyValCancelArgs(AbstractKeyValueArgs):
    """
    Key Value Cancel Args
    """

    __slots__ = ("key", "value", "source", "_event_name", "event_data", "cancel", "_event_source")

    @staticmethod
    def from_args(args: KeyValCancelArgs) -> KeyValCancelArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (KeyValCancelArgs): Existing Instance

        Returns:
            KeyValCancelArgs: args
        """
        eargs = KeyValCancelArgs(source=args.source, key=args.key, value=args.value)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        return eargs
