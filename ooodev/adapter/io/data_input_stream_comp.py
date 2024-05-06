from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.io import XDataInputStream

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.data_input_stream_partial import DataInputStreamPartial
from ooodev.adapter.io.active_data_sink_partial import ActiveDataSinkPartial
from ooodev.adapter.io.connectable_partial import ConnectablePartial

if TYPE_CHECKING:
    from com.sun.star.io import DataInputStream  # service
    from ooodev.loader.inst.lo_inst import LoInst


class DataInputStreamComp(ComponentProp, DataInputStreamPartial, ActiveDataSinkPartial, ConnectablePartial):
    """
    Class for managing DataInputStream Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDataInputStream) -> None:
        """
        Constructor

        Args:
            component (XDataInputStream): UNO Component that supports ``com.sun.star.io.DataInputStream`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        DataInputStreamPartial.__init__(self, component=component, interface=None)
        ActiveDataSinkPartial.__init__(self, component=component, interface=None)  # type: ignore
        ConnectablePartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.io.DataInputStream",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> DataInputStreamComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            DataInputStreamComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XDataInputStream, "com.sun.star.io.DataInputStream", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> DataInputStream:
        """DataInputStream Component"""
        # pylint: disable=no-member
        return cast("DataInputStream", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
