from __future__ import annotations
from typing import Any, Generic, TypeVar
from ooodev.utils.type_var import PathOrStr
from ooodev.events.args.event_args_export import EventArgsExport

_T = TypeVar("_T")
# pylint: disable=protected-access
# pylint: disable=assigning-non-slot


class CancelEventArgsExport(EventArgsExport[_T], Generic[_T]):
    """Cancel Event Arguments for Export events."""

    __slots__ = ("cancel", "handled")

    def __init__(
        self,
        source: Any,
        event_data: _T,
        fnm: PathOrStr = "",
        overwrite: bool = True,
        cancel=False,
    ) -> None:
        """
        Constructor

        Args:
            source (Any): Event Source
            cancel (bool, optional): Cancel value. Defaults to False.
        """
        super().__init__(source, event_data, fnm=fnm, overwrite=overwrite)
        self.cancel = cancel
        self.handled = False

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}: {self.source}, {self.event_data}, {self.cancel}>"

    cancel: bool
    """Gets/Sets cancel value"""
    handled: bool
    """Get/Set Handled value. Typically if set to ``True`` then ``cancel`` is ignored."""

    @staticmethod
    def from_args(args: CancelEventArgsExport) -> CancelEventArgsExport:
        """
        Gets a new instance from existing instance

        Args:
            args (CancelEventArgsExport): Existing Instance

        Returns:
            CancelEventArgsExport: args
        """
        eargs = CancelEventArgsExport(
            source=args.source, event_data=args.event_data, fnm=args.fnm, overwrite=args.overwrite, cancel=args.cancel
        )
        eargs._event_name = args.event_name
        eargs._event_source = args.event_source
        eargs.handled = args.handled
        if args._kv_data is not None:
            eargs._kv_data = args._kv_data.copy()
        return eargs
