from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.io import XDataOutputStream

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.data_output_stream_partial import DataOutputStreamPartial
from ooodev.adapter.io.active_data_source_partial import ActiveDataSourcePartial

if TYPE_CHECKING:
    from com.sun.star.io import DataOutputStream  # service
    from ooodev.loader.inst.lo_inst import LoInst


class DataOutputStreamComp(ComponentProp, DataOutputStreamPartial, ActiveDataSourcePartial):
    """
    Class for managing DataOutputStream Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDataOutputStream) -> None:
        """
        Constructor

        Args:
            component (XDataOutputStream): UNO Component that supports ``com.sun.star.io.DataOutputStream`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        DataOutputStreamPartial.__init__(self, component=component, interface=None)
        ActiveDataSourcePartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.io.DataOutputStream",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> DataOutputStreamComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            DataOutputStreamComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XDataOutputStream, "com.sun.star.io.DataOutputStream", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> DataOutputStream:
        """DataOutputStream Component"""
        # pylint: disable=no-member
        return cast("DataOutputStream", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
