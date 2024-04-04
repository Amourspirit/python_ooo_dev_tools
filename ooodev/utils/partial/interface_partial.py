from __future__ import annotations
import contextlib
from typing import cast, Set
from ooodev.utils import gen_util as gUtil


class InterfacePartial:
    """Partial Class used for getting service information from a UNO component."""

    def __init__(self) -> None:
        """
        Constructor.

        Args:
            component (Any): Any Uno Component that supports ``XServiceInfo`` interface.
            lo_inst (LoInst, optional): Lo instance.
        """
        self.__clazz_set = cast(Set[str], None)

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
            # sample input: ExactNamePartial
            def convert_from_partial_name(name: str) -> str:
                # sample input: ExactName
                return f".{gUtil.Util.to_snake_case(name)}.{name}"

            def convert_from_x_name(name: str) -> str:
                # sample input: XExactName
                if name.startswith("X"):
                    name = name[1:]
                p_name = f"{name}Partial"
                return convert_from_partial_name(p_name)

            def contains_partial(name: str) -> bool:
                l_name = name.lower()
                for clz in self.__clazz_set:
                    if clz.lower().endswith(l_name):
                        return True
                return False

            if "." in itm:
                return False

            if itm.lower().endswith("partial"):
                nm = convert_from_partial_name(itm)
            else:
                nm = convert_from_x_name(itm)

            return contains_partial(nm)

        def has_full_name(itm: str) -> bool:
            def convert_to_ooodev(name: str) -> str:
                # sample input: com.sun.star.beans.XExactName
                suffix = name.replace("com.sun.star.", "")  # beans.XExactName
                ns, class_name = suffix.rsplit(".", 1)  #  beans, XExactName
                if class_name.startswith("X"):
                    class_name = class_name[1:]  # ExactName
                odev_ns = f"ooodev.adapter.{ns.lower()}.{gUtil.Util.to_snake_case(class_name)}_partial"
                odev_class = f"{class_name}Partial"
                return f"{odev_ns}.{odev_class}"

            with contextlib.suppress(Exception):
                if convert_to_ooodev(itm) in self.__clazz_set:
                    return True
            return False

        for itm in interface:
            if not itm:
                continue
            if itm in self.__clazz_set:
                return True
            if has_full_name(itm):
                return True
            if has_partial_name(itm):
                return True
        return False
