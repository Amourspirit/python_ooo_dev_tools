from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.sheet import XNamedRange

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.sheet.named_range_partial import NamedRangePartial
from ooodev.adapter.sheet.cell_range_referrer_partial import CellRangeReferrerPartial


if TYPE_CHECKING:
    from com.sun.star.sheet import NamedRange  # service


class _NamedRangeComp(ComponentProp):

    def __init__(self, component: XNamedRange) -> None:
        """
        Constructor

        Args:
            component (XNamedRange): UNO Component that supports ``com.sun.star.sheet.NamedRanges`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        # ContentProviderPartial is init in __new__
        # ContentProviderPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.NamedRange",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> NamedRange:
        """NamedRange Component"""
        # pylint: disable=no-member
        return cast("NamedRange", self._ComponentBase__get_component())  # type: ignore

    @property
    def __class__(self):
        # pretend to be a NamedRangeComp class
        return NamedRangeComp

    @property
    def is_shared_formula(self) -> bool | None:
        """
        Gets/Sets if this defined name represents a shared formula.

        This special property shall not be used externally. It is used by import and export filters for compatibility with spreadsheet documents containing shared formulas. Shared formulas are shared by several cells to save memory and to decrease file size.

        A defined name with this property set will not appear in the user interface of Calc, and its name will not appear in cell formulas. A formula referring to this defined name will show the formula definition contained in the name instead.

        **optional**

        Returns:
            bool: The value or None if not available.
        """
        if hasattr(self.component, "IsSharedFormula"):
            return self.component.IsSharedFormula
        return None

    @is_shared_formula.setter
    def is_shared_formula(self, value: bool) -> None:
        if hasattr(self.component, "IsSharedFormula"):
            self.component.IsSharedFormula = value

    @property
    def token_index(self) -> int | None:
        """
        Gets the index used to refer to this name in token arrays.

        A token describing a defined name shall contain the op-code obtained from the FormulaMapGroupSpecialOffset.NAME offset and this index as data part.

        **optional**

        Returns:
            int: The index or None if not available.
        """
        if hasattr(self.component, "TokenIndex"):
            return self.component.TokenIndex
        return None

    # endregion Properties


class NamedRangeComp(_NamedRangeComp, NamedRangePartial, CellRangeReferrerPartial):
    """
    Class for managing Sheet Cell Ranges Component.
    """

    # pylint: disable=unused-argument
    def __new__(cls, component: XNamedRange, *args, **kwargs):

        new_class = type("NamedRangesComp", (_NamedRangeComp,), {})

        builder = get_builder(component)
        builder_helper.builder_add_comp_defaults(builder)
        clz = builder.get_class_type(
            name="ooodev.adapter.sheet.named_range_comp.NamedRangeComp",
            base_class=new_class,
            set_mod_name=True,
        )
        builder.init_class_properties(clz)

        result = super(new_class, new_class).__new__(clz, *args, **kwargs)  # type: ignore
        builder.init_classes(result)
        _NamedRangeComp.__init__(result, component)
        return result

    def __init__(self, component: XNamedRange) -> None:
        """
        Constructor

        Args:
            component (XNamedRange): UNO Sheet Cell Range Component
        """
        pass


def get_builder(component: Any) -> DefaultBuilder:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    builder.auto_interface()

    return builder
