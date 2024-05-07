from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.io import XTextOutputStream2

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.text_output_stream2_partial import TextOutputStream2Partial
from ooodev.adapter.io.active_data_source_partial import ActiveDataSourcePartial

if TYPE_CHECKING:
    from com.sun.star.io import TextOutputStream  # service
    from ooodev.loader.inst.lo_inst import LoInst


class TextOutputStreamComp(ComponentProp, TextOutputStream2Partial, ActiveDataSourcePartial):
    """
    Class for managing TextOutputStream Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextOutputStream2) -> None:
        """
        Constructor

        Args:
            component (XTextOutputStream2): UNO Component that supports ``com.sun.star.io.TextOutputStream`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        TextOutputStream2Partial.__init__(self, component=component, interface=None)
        ActiveDataSourcePartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.io.TextOutputStream",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TextOutputStreamComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TextOutputStreamComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XTextOutputStream2, "com.sun.star.io.TextOutputStream", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> TextOutputStream:
        """TextOutputStream Component"""
        # pylint: disable=no-member
        return cast("TextOutputStream", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
