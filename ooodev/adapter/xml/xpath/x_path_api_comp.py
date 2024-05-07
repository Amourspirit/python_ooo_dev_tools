from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.xml.xpath import XXPathAPI
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.xml.xpath.x_path_api_partial import XPathAPIPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class XPathAPIComp(ComponentProp, XPathAPIPartial):
    """
    Class for managing XXPathAPI Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XXPathAPI) -> None:
        """
        Constructor

        Args:
            component (XPathAPI): UNO Component that supports ``com.sun.star.xml.xpath.XPathAPI`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        XPathAPIPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.xml.xpath.XPathAPI",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> XPathAPIComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            XPathAPIComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XXPathAPI, "com.sun.star.xml.xpath.XPathAPI", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> XPathAPIComp:
        """XPathAPIComp Component"""
        # pylint: disable=no-member
        return cast("XPathAPIComp", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
