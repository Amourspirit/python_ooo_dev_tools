from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.io import XPipe


from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.pipe_partial import PipePartial
from ooodev.adapter.io.connectable_partial import ConnectablePartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class PipeComp(ComponentProp, PipePartial, ConnectablePartial):
    """
    Class for managing Pipe Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPipe) -> None:
        """
        Constructor

        Args:
            component (XPipe): UNO Component that supports ``com.sun.star.io.Pipe`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        PipePartial.__init__(self, component=component, interface=None)
        ConnectablePartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.io.Pipe",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> PipeComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            PipeComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XPipe, "com.sun.star.io.Pipe", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods
