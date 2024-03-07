from __future__ import annotations
from typing import Any
from ooodev.events.args.key_val_args import AbstractKeyValArgs
from ooodev.events.args.cancel_event_args import AbstractCancelEventArgs

# pylint: disable=protected-access
# pylint: disable=assigning-non-slot


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
        self.default = False
        self.handled = False


class KeyValCancelArgs(AbstractKeyValueArgs):
    """
    Key Value Cancel Args

    .. versionadded:: 0.9.0
    """

    __slots__ = (
        "key",
        "value",
        "source",
        "_event_name",
        "event_data",
        "cancel",
        "handled",
        "_event_source",
        "_kv_data",
        "default",
    )

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
        eargs.default = args.default
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.cancel = args.cancel
        eargs.handled = args.handled
        return eargs
