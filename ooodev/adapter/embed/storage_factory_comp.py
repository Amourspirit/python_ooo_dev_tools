from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.lang import XSingleServiceFactory

from ooodev.adapter.component_prop import ComponentProp

from ooodev.adapter.lang.single_service_factory_partial import SingleServiceFactoryPartial

if TYPE_CHECKING:
    from com.sun.star.embed import StorageFactory  # service
    from com.sun.star.embed import XStorage
    from ooodev.loader.inst.lo_inst import LoInst


class StorageFactoryComp(ComponentProp, SingleServiceFactoryPartial):
    """
    Class for managing StorageFactory Component.

    The StorageFactory is a service that allows to create a storage based on either stream or URL.

    In case ``create_instance()`` call is used the result storage will be open in read-write mode based on an arbitrary medium.

    In case ``create_instance_with_arguments()`` call is used a sequence of the following parameters can be used:
    The parameters are optional, that means that sequence can be empty or contain only first parameter, or first and second one.
    In case no parameters are provided the call works the same way as ``create_instance_with_arguments()``.
    In case only first parameter is provided, the storage is opened in readonly mode.

    The opened root storage can support read access in addition to specified one.

    See Also:
        `API StorageFactory <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1embed_1_1StorageFactory.html>`_
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSingleServiceFactory) -> None:
        """
        Constructor

        Args:
            component (XSingleServiceFactory): UNO Component that supports ``com.sun.star.embed.StorageFactory`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        SingleServiceFactoryPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.embed.StorageFactory",)

    # endregion Overrides
    # com.sun.star.embed.StorageStream

    # region XSingleServiceFactory
    def create_instance(self) -> XStorage:
        """
        Creates an instance of a service implementation.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return super().create_instance()

    def create_instance_with_arguments(self, *args: Any) -> XStorage:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Args:
            args (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return super().create_instance_with_arguments(*args)

    def create_instance_with_prop_args(self, **kwargs: Any) -> XStorage:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Each Key, Value pair is converted to a ``PropertyValue`` before adding to the service arguments.

        Args:
            kwargs (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return super().create_instance_with_prop_args(**kwargs)

    # endregion XSingleServiceFactory

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> StorageFactoryComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            StorageFactoryComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XSingleServiceFactory, "com.sun.star.embed.StorageFactory", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Properties
    @property
    def component(self) -> StorageFactory:
        """StorageFactory Component"""
        # pylint: disable=no-member
        return cast("StorageFactory", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
