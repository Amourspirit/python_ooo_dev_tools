from __future__ import annotations
from typing import Any, cast, Tuple, overload
import uno

from ooo.dyn.beans.property_value import PropertyValue
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.init_kind import InitKind
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs


class CombinedServiceFactoryPartial:
    """
    Partial class for ``XSingleServiceFactory``  and ``XMultiServiceFactory``.

    Because the two interface has the same method names, this class allows for overloads of the methods.

    The builder that is return by ``get_builder()`` for this module will only include the ``CombinedServiceFactoryPartial`` class if both ``XMultiServiceFactory`` and ``XSingleServiceFactory`` are supported.
    """

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any ): UNO Component that implements ``com.sun.star.lang.XMultiServiceFactory`` and ``com.sun.star.lang.XSingleServiceFactory`` interfaces.
        """

        def on_is_supported_interface(src: Any, event_args: CancelEventArgs) -> None:
            # search_term = str(event_args.event_data["search_term"])
            is_partial = bool(event_args.event_data["is_partial"])
            name = str(event_args.event_data["name"])
            is_match = bool(event_args.event_data["is_match"])
            if event_args.cancel:
                return
            if is_match:
                return
            if is_partial:
                if name.lower() in (
                    ".combined_service_factory_partial.combinedservicefactorypartial",
                    ".multi_service_factory_partial.multiservicefactorypartial",
                    ".single_service_factory_partial.singleservicefactorypartial",
                ):
                    event_args.event_data["is_match"] = True
            else:
                if name.lower() in (
                    "ooodev.adapter.lang.combined_service_factory_partial.combinedservicefactorypartial",
                    "ooodev.adapter.lang.multi_service_factory_partial.multiservicefactorypartial",
                    "ooodev.adapter.lang.single_service_factory_partial.singleservicefactorypartial",
                ):
                    event_args.event_data["is_match"] = True

        self.__component = component
        if isinstance(self, EventsPartial):
            self.__fn_on_is_supported_interface = on_is_supported_interface
            self.subscribe_event("interface_partial.is_supported_interface", self.__fn_on_is_supported_interface)

    # region XSingleServiceFactory
    def __create_instance_single(self) -> Any:
        """
        Creates an instance of a service implementation.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstance()

    def __create_instance_with_arguments_single(self, *args: Any) -> Any:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Args:
            args (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstanceWithArguments(args)

    def __create_instance_with_prop_args_single(self, **kwargs: Any) -> Any:
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

    # region XMultiServiceFactory
    def __create_instance(self, service_name: str) -> Any:
        """
        Creates an instance classified by the specified name.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstance(service_name)

    def __create_instance_with_arguments(self, *args: Any, service_name: str) -> Any:
        """
        Creates an instance classified by the specified name and passes the arguments to that instance.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.createInstanceWithArguments(service_name, args)

    def __create_instance_with_prop_args(self, service_name: str, **kwargs: Any) -> Any:
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
        return self.__component.createInstanceWithArguments(service_name, tuple(pvs))

    def get_available_service_names(self) -> Tuple[str, ...]:
        """
        Provides the available names of the factory to be used to create instances.
        """
        return self.__component.getAvailableServiceNames()

    # endregion XMultiServiceFactory

    # region create_instance()

    @overload
    def create_instance(self) -> Any:
        """
        Creates an instance of a service implementation.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        ...

    @overload
    def create_instance(self, service_name: str) -> Any:
        """
        Creates an instance classified by the specified name.

        Args:
            service_name (str): The name of the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        ...

    def create_instance(self, *args, **kwargs) -> Any:
        """
        Creates an instance of a service implementation.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        kargs_len = len(kwargs)
        count = len(args) + kargs_len
        if count == 0:
            return self.__create_instance_single()
        if count == 1:
            if args:
                return self.__create_instance(args[0])
            if "service_name" not in kwargs:
                raise TypeError("create_instance() got an unexpected keyword argument")
            return self.__create_instance(kwargs["service_name"])
        raise TypeError("create_instance() got an invalid number of arguments")

    # endregion create_instance()

    # region create_instance_with_arguments()

    @overload
    def create_instance_with_arguments(self, *args: Any) -> Any:
        """
        Creates an instance classified by the specified name and passes the arguments to that instance.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        ...

    @overload
    def create_instance_with_arguments(self, *args: Any, service_name: str) -> Any:
        """
        Creates an instance classified by the specified name and passes the arguments to that instance.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        ...

    def create_instance_with_arguments(self, *args: Any, service_name: Any = None) -> Any:
        """
        Creates an instance classified by the specified name and passes the arguments to that instance.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        if service_name is None:
            return self.__create_instance_with_arguments_single(*args)

        return self.__create_instance_with_arguments(*args, service_name=service_name)

    # endregion create_instance_with_arguments()

    @overload
    def create_instance_with_prop_args(self, **kwargs: Any) -> Any:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Each Key, Value pair is converted to a ``PropertyValue`` before adding to the service arguments.

        Args:
            kwargs (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        ...

    @overload
    def create_instance_with_prop_args(self, service_name: str, **kwargs: Any) -> Any:
        """
        Creates an instance of a service implementation initialized with some arguments.

        Each Key, Value pair is converted to a ``PropertyValue`` before adding to the service arguments.

        Args:
            service_name (str): The name of the service to create.
            kwargs (Any): One or more arguments to initialize the service.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        ...

    def create_instance_with_prop_args(self, service_name: str = "", **kwargs: Any) -> Any:
        if service_name == "":
            return self.__create_instance_with_prop_args_single(**kwargs)
        return self.__create_instance_with_prop_args(service_name, **kwargs)


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.add_import(
        name="ooodev.adapter.lang.combined_service_factory_partial.CombinedServiceFactoryPartial",
        uno_name=cast(
            Tuple[str], ("com.sun.star.lang.XSingleServiceFactory", "com.sun.star.lang.XMultiServiceFactory")
        ),
        init_kind=InitKind.COMPONENT,
        check_kind=CheckKind.INTERFACE_ALL,
    )
    return builder
