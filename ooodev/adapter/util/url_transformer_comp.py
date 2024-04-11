from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.util import XURLTransformer
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.url_transformer_partial import URLTransformerPartial
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial

if TYPE_CHECKING:
    from com.sun.star.util import URLTransformer
    from ooodev.loader.inst.lo_inst import LoInst


class URLTransformerComp(ComponentBase, URLTransformerPartial, ServiceInfoPartial):
    """
    Class for managing URLTransformer Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XURLTransformer) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that supports `com.sun.star.util.URLTransformer`` service.
        """
        # pylint: disable=no-member
        ComponentBase.__init__(self, component)
        URLTransformerPartial.__init__(self, component=component)
        ServiceInfoPartial.__init__(self, component=component, interface=None)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.util.URLTransformer",)

    # endregion Overrides

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> URLTransformerComp:
        """
        Creates the instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            MenuBarComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XURLTransformer, "com.sun.star.util.URLTransformer", raise_err=True)  # type: ignore
        return cls(inst)

    # region Properties

    @property
    def component(self) -> URLTransformer:
        """URLTransformer Component"""
        # pylint: disable=no-member
        return cast("URLTransformer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
