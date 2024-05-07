from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.io import XTextInputStream2

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.text_input_stream2_partial import TextInputStream2Partial

if TYPE_CHECKING:
    from com.sun.star.io import TextInputStream  # service
    from ooodev.loader.inst.lo_inst import LoInst


class TextInputStreamComp(ComponentProp, TextInputStream2Partial):
    """
    Class for managing TextInputStream Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextInputStream2) -> None:
        """
        Constructor

        Args:
            component (XTextInputStream2): UNO Component that supports ``com.sun.star.io.TextInputStream`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        TextInputStream2Partial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.io.TextInputStream",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TextInputStreamComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TextInputStreamComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XTextInputStream2, "com.sun.star.io.TextInputStream", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> TextInputStream:
        """TextInputStream Component"""
        # pylint: disable=no-member
        return cast("TextInputStream", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
