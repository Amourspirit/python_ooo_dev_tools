from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.util.path_settings_partial import PathSettingsPartial

if TYPE_CHECKING:
    from com.sun.star.util import thePathSettings  # singleton
    from ooodev.loader.inst.lo_inst import LoInst


class ThePathSettingsComp(ComponentBase, PathSettingsPartial):
    """
    Class for managing thePathSettings Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: thePathSettings) -> None:
        """
        Constructor

        Args:
            component (thePathSettings): UNO Component that implements ``com.sun.star.ui.thePathSettings`` service.
        """
        ComponentBase.__init__(self, component)
        PathSettingsPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.util.PathSettings",)

    # endregion Overrides
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ThePathSettingsComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ThePathSettingsComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.util.thePathSettings"
        if key in lo_inst.cache:
            return cast(ThePathSettingsComp, lo_inst.cache[key])

        factory = lo_inst.get_singleton("/singletons/com.sun.star.util.thePathSettings")  # type: ignore
        if factory is None:
            raise ValueError("Could not get thePathSettings singleton.")
        class_inst = cls(factory)
        lo_inst.cache[key] = class_inst
        return cast(ThePathSettingsComp, class_inst)

    # region Properties
    @property
    def component(self) -> thePathSettings:
        """thePathSettings Component"""
        # pylint: disable=no-member
        return cast("thePathSettings", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
