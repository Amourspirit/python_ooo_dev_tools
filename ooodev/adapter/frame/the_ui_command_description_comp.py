from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container import name_access_partial
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.ui.module_ui_command_description_comp import ModuleUICommandDescriptionComp

if TYPE_CHECKING:
    from com.sun.star.frame import theUICommandDescription  # singleton
    from com.sun.star.ui import ModuleUICommandDescription
    from ooodev.loader.inst.lo_inst import LoInst


class TheUICommandDescriptionComp(
    ComponentProp, name_access_partial.NameAccessPartial[ModuleUICommandDescriptionComp]
):
    """
    Class for managing theUICommandDescription Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: theUICommandDescription) -> None:
        """
        Constructor

        Args:
            component (theUICommandDescription): UNO Component that implements ``com.sun.star.frame.theUICommandDescription`` service.
        """
        ComponentProp.__init__(self, component)
        name_access_partial.NameAccessPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.frame.UICommandDescription",)

    # endregion Overrides

    # region NameAccessPartial Overrides
    def get_by_name(self, name: str) -> ModuleUICommandDescriptionComp:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            Any: The element with the specified name.
        """
        comp = self.component.getByName(name)
        if comp is None:
            return None  # type: ignore
        return ModuleUICommandDescriptionComp(comp)

    # endregion NameAccessPartial Overrides

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TheUICommandDescriptionComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TheUICommandDescriptionComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        key = "com.sun.star.frame.theUICommandDescription"
        if key in lo_inst.cache:
            return cast(TheUICommandDescriptionComp, lo_inst.cache[key])

        factory = lo_inst.get_singleton("/singletons/com.sun.star.frame.theUICommandDescription")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theUICommandDescription singleton.")
        inst = cls(factory)
        lo_inst.cache[key] = inst
        return inst

    # region Properties
    @property
    def component(self) -> theUICommandDescription:
        """theUICommandDescription Component"""
        # pylint: disable=no-member
        return cast("theUICommandDescription", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    return name_access_partial.get_builder(component)
