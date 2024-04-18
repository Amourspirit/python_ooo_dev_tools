from __future__ import annotations
import contextlib
from typing import cast, Set
from ooodev.utils import gen_util as gUtil
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs


class InterfacePartial:
    """
    Partial Class used for getting service information from a UNO component.

    Note:
        If this class is used in class that inherit from ``EventsPartial`` then the following event will be triggered:

        Event ``interface_partial.is_supported_interface`` is triggered when the method ``is_supported_interface()`` is called.

        ``CancelEventArgs`` Event data: ``{"search_term": itm, "is_partial": is_partial, "name": name, "is_match": False}``
        If the event is cancelled, the method will return ``False``.
        If the event data ``is_match`` is set to ``True`` then the method will return ``True``.
    """

    def __init__(self) -> None:
        """
        Constructor.

        Args:
            component (Any): Any Uno Component that supports ``XServiceInfo`` interface.
            lo_inst (LoInst, optional): Lo instance.
        """
        self.__clazz_set = cast(Set[str], None)
        self.__has_events = isinstance(self, EventsPartial)

    def is_supported_interface(self, *interface: str) -> bool:
        """
        Gets if instance supports an interface in the OooDev Framework.

        This is a convenience method to check if the instance supports any of the passed in interfaces.

        Rather then import the class and check if the instance is an instance of the class, this method allows to check by name.

        Args:
            *interface (str): One or more case sensitive interface strings to test for.

        Example:

        .. code-block:: python

            from ooodev.adapter.beans.property_set_info_partial import PropertySetInfoPartial

            if instance(my_instance, PropertySetInfoPartial):
                prop = instance.get_property_by_name()
                # do something

        is functionally the same as:

        .. code-block:: python

            if instance.is_supported_interface("com.sun.star.beans.XPropertySetInfo"):
                prop = instance.get_property_by_name()
                # do something

        Short forms are also acceptable such as:

        .. code-block:: python

            if instance.is_supported_interface("XPropertySetInfo"):
                prop = instance.get_property_by_name()
                # do something

        or

        .. code-block:: python

            if instance.is_supported_interface("PropertySetInfoPartial"):
                prop = instance.get_property_by_name()
                # do something

        Returns:
            bool: ``True`` if instance supports any passed in service; Otherwise, ``False``
        """

        if self.__clazz_set is None:
            cs: Set[str] = set()
            classes = type.mro(self.__class__)
            for clz in classes:
                cs.add(f"{clz.__module__}.{clz.__name__}")
            self.__clazz_set = cs

        def has_partial_name(itm: str) -> bool:

            def contains_partial(name: str) -> bool:
                l_name = name.lower()
                for clz in self.__clazz_set:
                    if clz.lower().endswith(l_name):
                        return True
                return False

            return contains_partial(itm)

        def has_full_name(itm: str) -> bool:
            return itm in self.__clazz_set

        def get_name(itm: str) -> str:
            def convert_to_ooodev(name: str) -> str:
                # sample input: com.sun.star.beans.XExactName
                suffix = name.replace("com.sun.star.", "")  # beans.XExactName
                ns, class_name = suffix.rsplit(".", 1)  #  beans, XExactName
                if class_name.startswith("X"):
                    class_name = class_name[1:]  # ExactName
                odev_ns = f"ooodev.adapter.{ns.lower()}.{gUtil.Util.camel_to_snake(class_name)}_partial"
                odev_class = f"{class_name}Partial"
                return f"{odev_ns}.{odev_class}"

            name = ""
            with contextlib.suppress(Exception):
                name = convert_to_ooodev(itm)
            return name

        def get_uno_name(itm: str) -> str:
            def convert_from_partial_name(name: str) -> str:
                # sample input: ExactName
                return f".{gUtil.Util.to_snake_case(name)}.{name}"

            def convert_from_x_name(name: str) -> str:
                # sample input: XExactName
                if name.startswith("X"):
                    name = name[1:]
                p_name = f"{name}Partial"
                return convert_from_partial_name(p_name)

            name = ""

            if "." in itm:
                return ""
            if itm.lower().endswith("partial"):
                name = convert_from_partial_name(itm)
            else:
                name = convert_from_x_name(itm)
            return name

        for itm in interface:
            if not itm:
                continue
            if itm in self.__clazz_set:
                return True

            name = get_name(itm)
            is_partial = False
            if not name:
                is_partial = True
                name = get_uno_name(itm)
            if not name:
                continue
            if self.__has_events:
                cargs = CancelEventArgs(self)
                cargs.event_data = {"search_term": itm, "is_partial": is_partial, "name": name, "is_match": False}
                self.trigger_event("interface_partial.is_supported_interface", cargs)  # type: ignore
                if cargs.cancel:
                    return False
                if cargs.event_data["is_match"]:
                    return True
            if has_full_name(name):
                return True
            if has_partial_name(name):
                return True
        return False
