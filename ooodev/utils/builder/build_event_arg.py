from __future__ import annotations
from typing import NamedTuple, Tuple
from ooodev.utils.builder.check_kind import CheckKind

# from dataclasses import dataclass


class BuildEventArg(NamedTuple):

    module_name: str
    """Ooodev name of the module such as ``ooodev.adapter.util.refresh_events``."""
    class_name: str
    """Ooodev class name such as ``RefreshEvents``."""
    callback_name: str
    """
    CallBack for the event. Usually ``on_lazy_cb``.

    Expected signature ``def on_lazy_cb(source: Any, event: ListenerEventArgs, comp: XRefreshable, inst: RefreshEvents)``
    """
    uno_name: Tuple[str]
    """Uno name of the class such as ``com.sun.star.container.XNameAccess``."""
    optional: bool
    """
    Specifies if the import is optional.
    If optional the a check is done to see if the component implements the interface.
    """
    check_kind: CheckKind = CheckKind.INTERFACE
    """Specifies the check kind. Defaults to ``CheckKind.INTERFACE``."""

    def __copy__(self) -> BuildEventArg:
        return BuildEventArg(
            module_name=self.module_name,
            class_name=self.class_name,
            callback_name=self.callback_name,
            uno_name=self.uno_name,
            optional=self.optional,
            check_kind=self.check_kind,
        )

    def __hash__(self):
        return hash((self.module_name, self.class_name, self.callback_name))

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, BuildEventArg):
            return NotImplemented
        return (
            self.module_name == other.module_name
            and self.class_name == other.class_name
            and self.callback_name == other.callback_name
        )
