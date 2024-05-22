from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.ucb import XContent
from com.sun.star.io import XStream

from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.embed.encryption_protected_source_partial import EncryptionProtectedSourcePartial  # optional
from ooodev.adapter.io.seekable_partial import SeekablePartial  # optional
from ooodev.adapter.io.stream_partial import StreamPartial
from ooodev.adapter.lang.component_partial import ComponentPartial


if TYPE_CHECKING:
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.adapter._helper.builder import builder_helper
    from com.sun.star.embed import StorageStream  # service

    # class TransientDocumentsContentProviderComp(ComponentProp, ContentProviderPartial):


class _StorageStreamComp(ComponentProp):

    def __init__(self, component: XStream) -> None:
        """
        Constructor

        Args:
            component (XStream): UNO Component that supports ``com.sun.star.embed.StorageStream`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.embed.StorageStream",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> StorageStream:
        """StorageStream Component"""
        # pylint: disable=no-member
        return cast("StorageStream", self._ComponentBase__get_component())  # type: ignore

    @property
    def __class__(self):
        # pretend to be a StorageStreamComp class
        return StorageStreamComp

    # endregion Properties


class StorageStreamComp(
    _StorageStreamComp,
    StreamPartial,
    ComponentPartial,
    PropertySetPartial,
    SeekablePartial,
    EncryptionProtectedSourcePartial,
    CompDefaultsPartial,
):
    """
    Class for managing StorageStream Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional or different classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~ooodev.utils.partial.interface_partial.InterfacePartial.is_supported_interface` methods which is always available in this class.
    """

    # pylint: disable=unused-argument

    def __new__(cls, component: XStream, *args, **kwargs):

        new_class = type("StorageStreamComp", (_StorageStreamComp,), {})

        builder = get_builder(component)
        builder_helper.builder_add_comp_defaults(builder)
        clz = builder.get_class_type(
            name="ooodev.adapter.embed.storage_stream_comp.StorageStreamComp",
            base_class=new_class,
            set_mod_name=True,
        )
        builder.init_class_properties(clz)

        result = super(new_class, new_class).__new__(clz, *args, **kwargs)  # type: ignore
        # result = super().__new__(clz, *args, **kwargs)  # type: ignore
        builder.init_classes(result)
        _StorageStreamComp.__init__(result, component)
        return result

    def __init__(self, component: XStream) -> None:
        """
        Constructor

        Args:
            component (XStream): UNO Component that supports ``com.sun.star.embed.StorageStream`` service.
        """
        pass

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    # pylint: disable=import-outside-toplevel
    # pylint: disable=redefined-outer-name

    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.adapter._helper.builder import builder_helper

    builder = DefaultBuilder(component)

    builder.auto_interface()
    builder_helper.builder_add_property_change_implement(builder)
    builder_helper.builder_add_property_veto_implement(builder)

    return builder
