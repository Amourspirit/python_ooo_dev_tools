from __future__ import annotations
from typing import Any
from ooodev.events.args.event_args import AbstractEvent


class AbstractListenerEventArgs(AbstractEvent):
    # https://stackoverflow.com/questions/472000/usage-of-slots
    __slots__ = ()

    def __init__(self, source: Any, trigger_name: str, is_add=True) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            trigger_name (str, optional): Event Trigger Name
        """
        super().__init__(source)
        self.trigger_name = trigger_name
        self.is_add = is_add

    trigger_name: str
    """Gets/Sets trigger name value"""
    is_add: bool
    """Gets/Sets if listener is being added or removed"""
    remove_callback: bool = False
    """Gets/Sets if callback should be removed after being triggered"""


class ListenerEventArgs(AbstractListenerEventArgs):
    """Cancel Event Arguments"""

    __slots__ = (
        "source",
        "_event_name",
        "event_data",
        "trigger_name",
        "is_add",
        "_event_source",
        "_kv_data",
        "remove_callback",
    )

    @staticmethod
    def from_args(args: AbstractListenerEventArgs) -> ListenerEventArgs:
        """
        Gets a new instance from existing instance

        Args:
            args (AbstractListenerEventArgs): Existing Instance

        Returns:
            ListenerEventArgs: args
        """
        # pylint: disable=protected-access
        eargs = ListenerEventArgs(source=args.source, trigger_name=args.trigger_name, is_add=args.is_add)
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.event_data = args.event_data
        eargs.remove_callback = args.remove_callback
        return eargs
