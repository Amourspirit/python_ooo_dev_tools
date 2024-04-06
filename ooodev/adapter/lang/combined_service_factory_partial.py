from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple, overload
import uno

from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.lang import XSingleServiceFactory
from ooo.dyn.beans.property_value import PropertyValue
from ooodev.adapter import builder_helper
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.init_kind import InitKind
from ooodev.adapter.lang import single_service_factory_partial
from ooodev.adapter.lang import multi_service_factory_partial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    pass


class _CombinedServiceFactoryPartial:

    pass


class CombinedServiceFactoryPartial:
    """
    Partial class for ``XSingleServiceFactory``  and ``XMultiServiceFactory``.

    Because the two interface has the same method names, this class allows for overloads of the methods.

    This class is also a factory class that can create a class based on the component for ``XSingleServiceFactory`` or ``XMultiServiceFactory``
    If only one interface is implemented, the class will be created with the appropriate interface.
    If no interface is implemented, an error will be raised.
    """

    def __new__(cls, component: Any, *args, **kwargs):
        has_single = mLo.Lo.is_uno_interfaces(component, XSingleServiceFactory)
        has_multi = mLo.Lo.is_uno_interfaces(component, XMultiServiceFactory)
        if has_single and has_multi:
            inst = super(CombinedServiceFactoryPartial, cls).__new__(cls)
            inst.__init__(component)
            return inst
        builder_only = kwargs.get("_builder_only", False)
        if has_single:
            builder = single_service_factory_partial.get_builder(component=component)
            builder_helper.builder_add_interface_defaults(builder)
            if builder_only:
                # cast to prevent type checker error
                return cast(Any, builder)
            inst = builder.build_class(
                name="ooodev.adapter.lang.single_service_factory_partial.SingleServiceFactoryPartial",
                base_class=_CombinedServiceFactoryPartial,
            )
            return inst
        elif has_multi:
            builder = multi_service_factory_partial.get_builder(component=component)
            builder_helper.builder_add_interface_defaults(builder)
            if builder_only:
                # cast to prevent type checker error
                return cast(Any, builder)
            inst = builder.build_class(
                name="ooodev.adapter.lang.multi_service_factory_partial.MultiServiceFactoryPartial",
                base_class=_CombinedServiceFactoryPartial,
            )
            return inst

        raise mEx.MissingInterfaceError((XSingleServiceFactory, XMultiServiceFactory))

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (XMultiServiceFactory  ): UNO Component that implements ``com.sun.star.lang.XMultiServiceFactory  `` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMultiServiceFactory  ``.
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
