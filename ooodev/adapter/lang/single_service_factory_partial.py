from __future__ import annotations
from ast import Name
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.lang import XSingleServiceFactory
from ooo.dyn.beans.property_value import PropertyValue

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class SingleServiceFactoryPartial:
    """
    Partial class for ``XSingleServiceFactory``.
    """

    def __init__(
        self, component: XSingleServiceFactory, interface: UnoInterface | None = XSingleServiceFactory
    ) -> None:
        """
        Constructor

        Args:
            component (XSingleServiceFactory): UNO Component that implements ``com.sun.star.lang.XSingleServiceFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSingleServiceFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSingleServiceFactory
    def create_instance(self) -> Any:
        """
        Creates an instance of a service implementation.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstance()

    def create_instance_with_arguments(self, *args: Any) -> Any:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Args:
            args (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstanceWithArguments(args)

    def create_instance_with_prop_args(self, **kwargs: Any) -> Any:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Each Key, Value pair is converted to a ``PropertyValue`` before adding to the service arguments.

        Args:
            kwargs (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        # convert args to tuple
        pvs = []
        for key, value in kwargs.items():
            pv = PropertyValue(Name=key, Value=value)
            pvs.append(pv)
        if not pvs:
            return
        return self.__component.createInstanceWithArguments(tuple(pvs))

    # endregion XSingleServiceFactory


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.lang.XSingleServiceFactory", False)
    return builder
