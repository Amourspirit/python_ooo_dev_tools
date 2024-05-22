from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.ucb import XContentProvider

from ooodev.adapter.component_prop import ComponentProp

from ooodev.adapter.ucb.content_provider_partial import ContentProviderPartial
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter.frame.transient_documents_document_content_factory_partial import (
    TransientDocumentsDocumentContentFactoryPartial,
)
from ooodev.adapter.frame.transient_documents_document_content_identifier_factory_partial import (
    TransientDocumentsDocumentContentIdentifierFactoryPartial,
)
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.ucb.transient_documents_document_content_comp import TransientDocumentsDocumentContentComp

if TYPE_CHECKING:
    from com.sun.star.ucb import TransientDocumentsContentProvider  # service
    from com.sun.star.frame import XModel
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.loader.inst.lo_inst import LoInst
    from typing_extensions import Self

    # class TransientDocumentsContentProviderComp(ComponentProp, ContentProviderPartial):


class _TransientDocumentsContentProviderComp(ComponentProp):

    def __init__(self, component: XContentProvider) -> None:
        """
        Constructor

        Args:
            component (XContentProvider): UNO Component that supports ``com.sun.star.ucb.TransientDocumentsContentProvider`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        # ContentProviderPartial is init in __new__
        # ContentProviderPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.ucb.TransientDocumentsContentProvider",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> Self:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TransientDocumentsContentProviderComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XContentProvider, "com.sun.star.ucb.TransientDocumentsContentProvider", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region XTransientDocumentsDocumentContentFactory Overrides
    def create_document_content(self, model: XModel) -> TransientDocumentsDocumentContentComp:
        """
        Creates a ``com.sun.star.ucb.TransientDocumentsDocumentContent`` based on a given ``com.sun.star.document.OfficeDocument``.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return TransientDocumentsDocumentContentComp(self.component.createDocumentContent(model))  # type: ignore

    # endregion XTransientDocumentsDocumentContentFactory Overrides

    # region Properties
    @property
    def component(self) -> TransientDocumentsContentProvider:
        """TransientDocumentsContentProvider Component"""
        # pylint: disable=no-member
        return cast("TransientDocumentsContentProvider", self._ComponentBase__get_component())  # type: ignore

    @property
    def __class__(self):
        # pretend to be a TransientDocumentsContentProviderComp class
        return TransientDocumentsContentProviderComp

    # endregion Properties


class TransientDocumentsContentProviderComp(
    _TransientDocumentsContentProviderComp,
    ContentProviderPartial,
    TransientDocumentsDocumentContentFactoryPartial,
    TransientDocumentsDocumentContentIdentifierFactoryPartial,
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

    def __new__(cls, component: XContentProvider, *args, **kwargs):

        new_class = type("TransientDocumentsContentProviderComp", (_TransientDocumentsContentProviderComp,), {})

        builder = get_builder(component)
        builder_helper.builder_add_comp_defaults(builder)
        clz = builder.get_class_type(
            name="ooodev.adapter.ucb.transient_documents_content_provider_comp.TransientDocumentsContentProviderComp",
            base_class=new_class,
            set_mod_name=True,
        )
        builder.init_class_properties(clz)

        result = super(new_class, new_class).__new__(clz, *args, **kwargs)  # type: ignore
        # result = super().__new__(clz, *args, **kwargs)  # type: ignore
        builder.init_classes(result)
        _TransientDocumentsContentProviderComp.__init__(result, component)
        return result

    def __init__(self, component: XContentProvider) -> None:
        """
        Constructor

        Args:
            component (XContentProvider): UNO Component that supports ``com.sun.star.ucb.TransientDocumentsContentProvider`` service.
        """
        pass

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    builder.auto_interface()

    return builder
