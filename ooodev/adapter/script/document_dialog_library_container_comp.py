from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.script import XStorageBasedLibraryContainer
from com.sun.star.container import XNameAccess

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container.name_access_partial import NameAccessPartial
from ooodev.adapter.script.library_container3_partial import LibraryContainer3Partial
from ooodev.adapter.script.library_query_executable_partial import LibraryQueryExecutablePartial
from ooodev.adapter.script.storage_based_library_container_partial import StorageBasedLibraryContainerPartial
from ooodev.adapter.script.vba.vba_compatibility_partial import VBACompatibilityPartial
from ooodev.adapter.script.vba.vba_script_events import VBAScriptEvents
from ooodev.utils.builder.check_kind import CheckKind

if TYPE_CHECKING:
    from com.sun.star.script import DocumentDialogLibraryContainer  # service
    from com.sun.star.document import XStorageBasedDocument
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.loader.inst.lo_inst import LoInst
    from typing_extensions import Self


class _DocumentDialogLibraryContainerComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.script.DocumentDialogLibraryContainer",)

    # region DocumentDialogLibraryContainer
    def create(self, document: XStorageBasedDocument) -> None:
        """
        Creates an instance of the DocumentDialogLibraryContainer, belonging to a document

        The current storage of the document will be set as initial root storage (see XPersistentLibraryContainer.RootStorage) of the container.

        Usually, you will only create a DocumentDialogLibraryContainer within the implementation of the document to which the container should belong.

        Args:
            document (XStorageBasedDocument): The document to which the container should belong to. Must not be ``None``.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.component.create(document)

    def create_with_url(self, url: str) -> None:
        """
        Creates an instance of the DocumentDialogLibraryContainer, belonging to a document

        Args:
            url (str): The URL of the document to which the container should belong to.
        """
        self.component.createWithURL(url)

    # endregion DocumentDialogLibraryContainer

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> Self:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            DocumentDialogLibraryContainerComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XStorageBasedLibraryContainer, "com.sun.star.script.DocumentDialogLibraryContainer", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def __class__(self):
        # pretend to be a DocumentScriptLibraryContainerComp class
        return DocumentDialogLibraryContainerComp

    # endregion Properties


class DocumentDialogLibraryContainerComp(
    _DocumentDialogLibraryContainerComp,
    StorageBasedLibraryContainerPartial,
    LibraryContainer3Partial,
    LibraryQueryExecutablePartial,
    VBACompatibilityPartial,
    VBAScriptEvents,
    PropertySetPartial,
    NameAccessPartial[XNameAccess],
    CompDefaultsPartial,
):
    """
    Service Class

    Defines a container of dialog libraries, which is to be made persistent in a sub storage of a document storage.

    See Also:
        `API DocumentDialogLibraryContainer <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1script_1_1DocumentDialogLibraryContainer.html>`_
    """

    # pylint: disable=unused-argument
    def __new__(cls, component: Any, *args, **kwargs):
        builder = get_builder(component=component)
        builder_helper.builder_add_comp_defaults(builder)
        # if builder.has_import("ooodev.utils.partial.qi_partial.QiPartial"):
        #     builder.remove_import("ooodev.utils.partial.qi_partial.QiPartial")

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.adapter.script.document_dialog_library_container_comp.DocumentDialogLibraryContainerComp",
            base_class=_DocumentDialogLibraryContainerComp,
        )
        return inst

    def __init__(self, component: XStorageBasedLibraryContainer) -> None:
        """
        Constructor

        Args:
            component (XStorageBasedLibraryContainer): UNO Component that supports ``com.sun.star.script.DocumentDialogLibraryContainer`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties
    @property
    def component(self) -> DocumentDialogLibraryContainer:
        """DocumentDialogLibraryContainer Component"""
        # pylint: disable=no-member
        return cast("DocumentDialogLibraryContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_interface()
    builder.auto_add_interface(uno_name="com.sun.star.uno.XInterface")
    # com.sun.star.script.vba.VBAScriptListener is not implemented but is supported
    if hasattr(component, "addVBAScriptListener"):
        builder.add_event(
            module_name="ooodev.adapter.script.vba.vba_script_events",
            class_name="VBAScriptEvents",
            uno_name="com.sun.star.script.vba.VBAScriptListener",
            optional=False,
            check_kind=CheckKind.NONE,
        )
    if hasattr(component, "hasByName"):
        # com.sun.star.container.XNameAccess is not implemented but is supported
        builder.add_import(
            name="ooodev.adapter.container.name_access_partial.NameAccessPartial",
            uno_name="com.sun.star.container.XNameAccess",
            optional=False,
            check_kind=CheckKind.NONE,
        )
    return builder
