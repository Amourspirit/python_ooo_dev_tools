from __future__ import annotations
from typing import NamedTuple, Tuple
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils.builder.check_kind import CheckKind

# from dataclasses import dataclass


class BuildImportArg(NamedTuple):

    ooodev_name: str
    """Ooodev name of the class such as ``ooodev.adapter.container.name_access_partial.NameAccessPartial``."""
    uno_name: Tuple[str]
    """Uno name of the class such as ``com.sun.star.container.XNameAccess``."""
    optional: bool
    """
    Specifies if the import is optional.
    If optional the a check is done to see if the component implements the interface.
    """
    init_kind: InitKind = InitKind.COMPONENT_INTERFACE
    """Specifies the import kind. Defaults to ``ImportKind.COMPONENT_INTERFACE``."""
    check_kind: CheckKind = CheckKind.INTERFACE
    """Specifies the check kind. Defaults to ``CheckKind.INTERFACE``."""

    def __copy__(self) -> BuildImportArg:
        return BuildImportArg(
            ooodev_name=self.ooodev_name,
            uno_name=self.uno_name,
            optional=self.optional,
            init_kind=self.init_kind,
            check_kind=self.check_kind,
        )

    def __hash__(self):
        return hash((self.ooodev_name))

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, BuildImportArg):
            return NotImplemented
        return self.ooodev_name == other.ooodev_name
