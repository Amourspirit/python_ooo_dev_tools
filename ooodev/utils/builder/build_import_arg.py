from __future__ import annotations
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils.builder.check_kind import CheckKind
from dataclasses import dataclass


@dataclass(frozen=False)
class BuildImportArg:
    uno_name: str
    """Uno name of the class such as ``com.sun.star.container.XNameAccess``."""
    ooodev_name: str
    """Ooodev name of the class such as ``ooodev.adapter.container.name_access_partial.NameAccessPartial``."""
    optional: bool
    """
    Specifies if the import is optional.
    If optional the a check is done to see if the component implements the interface.
    """
    import_kind: InitKind = InitKind.COMPONENT_INTERFACE
    """Specifies the import kind. Defaults to ``ImportKind.COMPONENT_INTERFACE``."""
    check_kind: CheckKind = CheckKind.INTERFACE
    """Specifies the check kind. Defaults to ``CheckKind.INTERFACE``."""

    def __copy__(self) -> BuildImportArg:
        return BuildImportArg(
            uno_name=self.uno_name,
            ooodev_name=self.ooodev_name,
            optional=self.optional,
            import_kind=self.import_kind,
            check_kind=self.check_kind,
        )
