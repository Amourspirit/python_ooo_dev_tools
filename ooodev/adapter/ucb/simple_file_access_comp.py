from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.ucb import XSimpleFileAccess3

from ooodev.adapter.component_prop import ComponentProp

from ooodev.adapter.ucb.simple_file_access3_partial import SimpleFileAccess3Partial

if TYPE_CHECKING:
    from com.sun.star.ucb import SimpleFileAccess  # service
    from ooodev.loader.inst.lo_inst import LoInst


class SimpleFileAccessComp(ComponentProp, SimpleFileAccess3Partial):
    """
    Class for managing SimpleFileAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSimpleFileAccess3) -> None:
        """
        Constructor

        Args:
            component (XSimpleFileAccess3): UNO Component that supports ``com.sun.star.ucb.SimpleFileAccess`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        SimpleFileAccess3Partial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.ucb.SimpleFileAccess",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> SimpleFileAccessComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            SimpleFileAccessComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XSimpleFileAccess3, "com.sun.star.ucb.SimpleFileAccess", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> SimpleFileAccess:
        """SimpleFileAccess Component"""
        # pylint: disable=no-member
        return cast("SimpleFileAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
