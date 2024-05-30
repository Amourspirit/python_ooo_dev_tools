from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.sheet import XDatabaseRange

from ooodev.adapter.component_prop import ComponentProp

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.sheet.named_range_comp import NamedRangeComp
from ooodev.adapter.sheet.database_range_partial import DatabaseRangePartial
from ooodev.adapter.sheet.cell_range_referrer_partial import CellRangeReferrerPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.container.named_partial import NamedPartial
from ooodev.adapter.util.refreshable_partial import RefreshablePartial

if TYPE_CHECKING:
    from com.sun.star.sheet import DatabaseRange  # service
    from ooodev.utils.builder.default_builder import DefaultBuilder

    # class TransientDocumentsContentProviderComp(ComponentProp, ContentProviderPartial):


class _DatabaseRangeComp(ComponentProp):

    def __init__(self, component: XDatabaseRange) -> None:
        """
        Constructor

        Args:
            component (XNamedRanges): UNO Component that supports ``com.sun.star.sheet.DatabaseRange`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.DatabaseRange",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> DatabaseRange:
        """DatabaseRange Component"""
        # pylint: disable=no-member
        return cast("DatabaseRange", self._ComponentBase__get_component())  # type: ignore

    @property
    def __class__(self):
        # pretend to be a DatabaseRangeComp class
        return DatabaseRangeComp

    # endregion Properties


class DatabaseRangeComp(
    _DatabaseRangeComp,
    DatabaseRangePartial,
    CellRangeReferrerPartial,
    PropertySetPartial,
    NamedPartial,
    RefreshablePartial,
    CompDefaultsPartial,
):
    """
    Class for managing TransientDocumentsContentProvider Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional or different classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~ooodev.utils.partial.interface_partial.InterfacePartial.is_supported_interface` methods which is always available in this class.
    """

    # pylint: disable=unused-argument

    def __new__(cls, component: XDatabaseRange, *args, **kwargs):

        new_class = type("DatabaseRangeComp", (_DatabaseRangeComp,), {})

        builder = get_builder(component)
        builder_helper.builder_add_comp_defaults(builder)
        builder_helper.builder_add_property_change_implement(builder)
        builder_helper.builder_add_property_veto_implement(builder)
        clz = builder.get_class_type(
            name="ooodev.adapter.sheet.database_range_comp.DatabaseRangeComp",
            base_class=new_class,
            set_mod_name=True,
        )
        builder.init_class_properties(clz)

        result = super(new_class, new_class).__new__(clz, *args, **kwargs)  # type: ignore
        # result = super().__new__(clz, *args, **kwargs)  # type: ignore
        builder.init_classes(result)
        _DatabaseRangeComp.__init__(result, component)
        return result

    def __init__(self, component: XDatabaseRange) -> None:
        """
        Constructor

        Args:
            component (XDatabaseRange): UNO Component that supports ``com.sun.star.sheet.DatabaseRange`` service.
        """
        pass

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    builder.auto_interface()

    return builder
