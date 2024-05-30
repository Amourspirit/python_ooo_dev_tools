from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.sheet import XNamedRanges

from ooodev.adapter.component_prop import ComponentProp

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.sheet.named_ranges_partial import NamedRangesPartial
from ooodev.adapter.document.action_lockable_partial import ActionLockablePartial
from ooodev.adapter.sheet.named_range_comp import NamedRangeComp


if TYPE_CHECKING:
    from com.sun.star.sheet import NamedRanges  # service
    from ooodev.utils.builder.default_builder import DefaultBuilder

    # class TransientDocumentsContentProviderComp(ComponentProp, ContentProviderPartial):


class _NamedRangesComp(ComponentProp):

    def __init__(self, component: XNamedRanges) -> None:
        """
        Constructor

        Args:
            component (XNamedRanges): UNO Component that supports ``com.sun.star.sheet.NamedRanges`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        # ContentProviderPartial is init in __new__
        # ContentProviderPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.NamedRanges",)

    def get_by_name(self, name: str) -> NamedRangeComp:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            Any: The element with the specified name.
        """
        result = self.component.getByName(name)
        if result is None:
            return None  # type: ignore
        return NamedRangeComp(result)  # type: ignore

    def get_by_index(self, idx: int) -> NamedRangeComp:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element.

        Returns:
            Any: The element at the specified index.
        """
        result = self.component.getByIndex(idx)
        if result is None:
            return None  # type: ignore
        return NamedRangeComp(result)  # type: ignore

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> NamedRanges:
        """NamedRanges Component"""
        # pylint: disable=no-member
        return cast("NamedRanges", self._ComponentBase__get_component())  # type: ignore

    @property
    def __class__(self):
        # pretend to be a NamedRangesComp class
        return NamedRangesComp

    # endregion Properties


class NamedRangesComp(
    _NamedRangesComp,
    IndexAccessPartial[NamedRangeComp],
    EnumerationAccessPartial[NamedRangeComp],
    NamedRangesPartial,
    ActionLockablePartial,
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

    def __new__(cls, component: XNamedRanges, *args, **kwargs):

        new_class = type("NamedRangesComp", (_NamedRangesComp,), {})

        builder = get_builder(component)
        builder_helper.builder_add_comp_defaults(builder)
        clz = builder.get_class_type(
            name="ooodev.adapter.sheet.named_ranges_comp.NamedRangesComp",
            base_class=new_class,
            set_mod_name=True,
        )
        builder.init_class_properties(clz)

        result = super(new_class, new_class).__new__(clz, *args, **kwargs)  # type: ignore
        # result = super().__new__(clz, *args, **kwargs)  # type: ignore
        builder.init_classes(result)
        _NamedRangesComp.__init__(result, component)
        return result

    def __init__(self, component: XNamedRanges) -> None:
        """
        Constructor

        Args:
            component (XNamedRanges): UNO Component that supports ``com.sun.star.sheet.NamedRanges`` service.
        """
        pass

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    # builder.auto_add_interface("com.sun.star.sheet.XNamedRanges", False)
    # builder.auto_add_interface("com.sun.star.container.XIndexAccess", False)
    # builder.auto_add_interface("com.sun.star.container.XEnumerationAccess", False)
    # builder.auto_add_interface("com.sun.star.document.XActionLockable")
    builder.auto_interface()

    builder.set_omit("com.sun.star.container.XElementAccess", "com.sun.star.container.XNameAccess")

    return builder
