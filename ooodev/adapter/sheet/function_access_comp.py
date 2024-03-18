from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.sheet import XFunctionAccess

from ooodev.adapter.sheet.function_access_partial import FunctionAccessPartial
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.sheet import FunctionAccess  # service
    from ooodev.loader.inst.lo_inst import LoInst


class FunctionAccessComp(ComponentBase, FunctionAccessPartial):
    """
    Class for managing Calc FunctionAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, lo_inst: LoInst, component: XFunctionAccess | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.
            component (XFunctionAccess, optional): Function Access Component.

        Returns:
            None:

        Note:
            ``component`` is automatically created from ``lo_inst`` if it is not provided.
        """
        if component is None:
            component = lo_inst.create_instance_mcf(
                XFunctionAccess, "com.sun.star.sheet.FunctionAccess", raise_err=True
            )
        ComponentBase.__init__(self, component)  # type: ignore
        FunctionAccessPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.FunctionAccess",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> FunctionAccess:
        """FunctionAccess Component"""
        # pylint: disable=no-member
        return cast("FunctionAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
