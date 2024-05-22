from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.ucb import XContent

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.beans.properties_change_notifier_partial import PropertiesChangeNotifierPartial
from ooodev.adapter.beans.property_container_partial import PropertyContainerPartial
from ooodev.adapter.beans.property_set_info_change_notifier_partial import PropertySetInfoChangeNotifierPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.child_partial import ChildPartial
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.adapter.ucb.command_info_change_notifier_partial import CommandInfoChangeNotifierPartial
from ooodev.adapter.ucb.command_processor2_partial import CommandProcessor2Partial
from ooodev.adapter.ucb.content_partial import ContentPartial
from ooodev.adapter.ucb.command_info_change_events import CommandInfoChangeEvents
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ucb import TransientDocumentsRootContent  # service
    from ooodev.utils.builder.default_builder import DefaultBuilder

    # class TransientDocumentsContentProviderComp(ComponentProp, ContentProviderPartial):


class _TransientDocumentsRootContentComp(ComponentProp):

    def __init__(self, component: XContent) -> None:
        """
        Constructor

        Args:
            component (XContent): UNO Component that supports ``com.sun.star.ucb.TransientDocumentsRootContent`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.ucb.TransientDocumentsRootContent",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> TransientDocumentsRootContent:
        """TransientDocumentsRootContent Component"""
        # pylint: disable=no-member
        return cast("TransientDocumentsRootContent", self._ComponentBase__get_component())  # type: ignore

    @property
    def __class__(self):
        # pretend to be a TransientDocumentsRootContentComp class
        return TransientDocumentsRootContentComp

    # endregion Properties


class TransientDocumentsRootContentComp(
    _TransientDocumentsRootContentComp,
    ComponentPartial,
    ContentPartial,
    CommandProcessor2Partial,
    PropertiesChangeNotifierPartial,
    PropertyContainerPartial,
    PropertySetInfoChangeNotifierPartial,
    CommandInfoChangeNotifierPartial,
    ChildPartial,
    CommandInfoChangeEvents,
    CompDefaultsPartial,
):
    """
    Class for managing TransientDocumentsRootContent Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional or different classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~ooodev.utils.partial.interface_partial.InterfacePartial.is_supported_interface` methods which is always available in this class.
    """

    # pylint: disable=unused-argument

    def __new__(cls, component: XContent, *args, **kwargs):

        new_class = type("TransientDocumentsRootContentComp", (_TransientDocumentsRootContentComp,), {})

        builder = get_builder(component)
        builder_helper.builder_add_comp_defaults(builder)
        clz = builder.get_class_type(
            name="ooodev.adapter.ucb.transient_documents_root_content_comp.TransientDocumentsRootContentComp",
            base_class=new_class,
            set_mod_name=True,
        )
        builder.init_class_properties(clz)

        result = super(new_class, new_class).__new__(clz, *args, **kwargs)  # type: ignore
        # result = super().__new__(clz, *args, **kwargs)  # type: ignore
        builder.init_classes(result)
        _TransientDocumentsRootContentComp.__init__(result, component)
        return result

    def __init__(self, component: XContent) -> None:
        """
        Constructor

        Args:
            component (XContent): UNO Component that supports ``com.sun.star.ucb.TransientDocumentsRootContent`` service.
        """
        pass

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    if mLo.Lo.is_uno_interfaces(component, "com.sun.star.ucb.XCommandProcessor2"):
        builder.set_omit("com.sun.star.ucb.XContentProvider")
    builder.auto_interface()
    builder.add_event(
        module_name="ooodev.adapter.ucb.command_info_change_events",
        class_name="CommandInfoChangeEvents",
        uno_name="com.sun.star.ucb.XCommandInfoChangeListener",
        optional=True,
    )

    return builder
