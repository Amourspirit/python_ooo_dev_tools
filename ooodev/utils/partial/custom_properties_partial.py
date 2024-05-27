from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno
import unohelper
from com.sun.star.lang import XComponent
from com.sun.star.container import XContainerListener
from com.sun.star.form import XForm
from ooo.dyn.beans.property_attribute import PropertyAttributeEnum
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.form.controls.form_ctl_hidden import FormCtlHidden
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils import props as mProps
from ooodev.utils.helper.dot_dict import DotDict


if TYPE_CHECKING:
    # from com.sun.star.form import Forms
    from com.sun.star.form.component import Form
    from com.sun.star.lang import EventObject
    from com.sun.star.container import ContainerEvent
    from com.sun.star.container import XContainer
    from ooodev.loader.inst.lo_inst import LoInst


class CustomPropertiesPartial:
    """
    Partial class to add custom properties to a class that implements XForm.

    Allows custom properties to be added to a document via its forms.

    Note:
        The following keys are forbidden: ``HiddenValue``, ``Name`` ``ClassId``, ``Tag``.

        Properties are stored in a hidden control on the form and are persisted with the document.

        Property value are basically ``int``, ``float``, ``str``, ``bool``. A Property value of ``None`` is not allowed; However, empty string is allowed.
    """

    class ContainerListener(unohelper.Base, XContainerListener):

        def __init__(
            self, form_name: str, cp: CustomPropertiesPartial, lo_inst: LoInst, subscriber: XContainer | None = None
        ) -> None:
            super().__init__()
            self._form_name = form_name
            self._cp = cp
            self._lo_inst = lo_inst
            if subscriber:
                subscriber.addContainerListener(self)

        def is_element_monitored_form(self, element: Any) -> bool:
            form = self.lo_inst.qi(XForm, element)
            if form is None:
                return False
            return form.Name == self._form_name  # type: ignore

        def reset(self) -> None:
            self._cp._CustomPropertiesPartial__reset()  # type: ignore

        # region XContainerListener
        def elementInserted(self, event: ContainerEvent) -> None:
            """
            Event is invoked when a container has inserted an element.
            """
            # replaced element should be a form
            if self.is_element_monitored_form(event.Element):
                self.reset()

        def elementRemoved(self, event: ContainerEvent) -> None:
            """
            Event is invoked when a container has removed an element.
            """
            if self.is_element_monitored_form(event.Element):
                self.reset()

        def elementReplaced(self, event: ContainerEvent) -> None:
            """
            Event is invoked when a container has replaced an element.
            """
            if self.is_element_monitored_form(event.ReplacedElement):
                self.reset()

        def disposing(self, event: EventObject) -> None:
            """
            Gets called when the broadcaster is about to be disposed.

            All listeners and all other objects, which reference the broadcaster
            should release the reference to the source. No method should be invoked
            anymore on this object ( including ``XComponent.removeEventListener()`` ).

            This method is called for every listener registration of derived listener
            interfaced, not only for registrations at ``XComponent``.
            """
            # from com.sun.star.lang.XEventListener
            self.reset()

        @property
        def lo_inst(self) -> LoInst:
            return self._lo_inst

    def __init__(self, forms: Any, form_name: str = "Form_CustomProperties", ctl_name: str = "CustomProperties"):
        """
        Constructor.

        Args:
            forms (Any): The component that implements ``XForm``.
            form_name (str, optional): The name to assign to the form. Defaults to "CustomProperties".
            ctl_name (str, optional): The name to assign to the hidden control. Defaults to "CustomProperties".
        """
        if isinstance(self, LoInstPropsPartial):
            lo_inst = self.lo_inst
        else:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__forbidden_keys = set(("HiddenValue", "Name", "ClassId", "Tag"))
        self.__component = forms
        self.__form_name = form_name
        self.__ctl_name = ctl_name
        self.__cache = {}
        # please the type checker
        self.__container_listener: CustomPropertiesPartial.ContainerListener
        self.__container_listener = CustomPropertiesPartial.ContainerListener(
            form_name=self.__form_name, cp=self, lo_inst=self.__lo_inst, subscriber=self.__component
        )

    def get_custom_property(self, name: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets a custom property.

        Args:
            name (str): The name of the property.
            default (Any, optional): The default value to return if the property does not exist.

        Raises:
            AttributeError: If the property is not found.

        Returns:
            Any: The value of the property.
        """
        ctl = self.__get_hidden_control()
        info = ctl.get_property_set_info()
        if info.hasPropertyByName(name):
            return ctl.get_property(name)
        if default is not NULL_OBJ:
            return default
        raise AttributeError(f"Property '{name}' not found.")

    def set_custom_property(self, name: str, value: Any):
        """
        Sets a custom property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.

        Raises:
            AttributeError: If the property is a forbidden key.
        """
        if name in self.__forbidden_keys:
            raise AttributeError(f"Property '{name}' is forbidden. Forbidden keys: {self.__forbidden_keys}")
        ctl = self.__get_hidden_control()
        info = ctl.get_property_set_info()
        if info.hasPropertyByName(name):
            ctl.remove_property(name)
        ctl.add_property(name, PropertyAttributeEnum.REMOVABLE, value)

    def get_custom_properties(self) -> DotDict:
        """
        Gets custom properties.

        Returns:
            DotDict: custom properties.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        ctl = self.__get_hidden_control()
        props = ctl.get_property_values()
        lst = []
        for prop in props:
            if prop.Name not in self.__forbidden_keys:
                lst.append(prop)
        return mProps.Props.props_to_dot_dict(lst)

    def set_custom_properties(self, properties: DotDict) -> None:
        """
        Sets custom properties.

        Args:
            properties (DotDict): custom properties to set.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        for name, value in properties.items():
            self.set_custom_property(name, value)

    def remove_custom_property(self, name: str) -> None:
        """
        Removes a custom property.

        Args:
            name (str): The name of the property to remove.

        Raises:
            AttributeError: If the property is a forbidden key.

        Returns:
            None:
        """
        if name in self.__forbidden_keys:
            raise AttributeError(f"Property '{name}' is forbidden. Forbidden keys: {self.__forbidden_keys}")
        ctl = self.__get_hidden_control()
        info = ctl.get_property_set_info()
        if info.hasPropertyByName(name):
            ctl.remove_property(name)

    def has_custom_property(self, name: str) -> bool:
        """
        Gets if a custom property exists.

        Args:
            name (str): The name of the property to check.

        Returns:
            bool: ``True`` if the property exists, otherwise ``False``.
        """
        ctl = self.__get_hidden_control()
        info = ctl.get_property_set_info()
        return info.hasPropertyByName(name)

    def __get_form(self) -> Form:

        key = f"{self.__form_name}"
        if key in self.__cache:
            return self.__cache[key]
        forms = self.__component
        if len(forms) == 0:
            # insert a default form1.
            # The reason for this is many users many working in forms[0].
            # This way there will be a from to work with that is not for properties.
            # This is not critical but it is a good practice.
            # If the user deletes Forms[0] it will not wipe the property forms.
            # Also if the user draws control on the spreadsheet or other document it will use this form.
            frm = cast(
                "Form", self.__lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True)
            )
            frm.Name = "Form1"
            forms.insertByName("Form1", frm)

        if forms.hasByName(key):
            frm = forms.getByName(key)
        else:
            frm = cast(
                "Form", self.__lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True)
            )
            frm.Name = key
            forms.insertByName(key, frm)
        self.__cache[key] = frm
        return frm

    def __get_hidden_control(self) -> FormCtlHidden:
        key = f"hidden_ctl_{self.__ctl_name}"
        if key in self.__cache:
            return cast(FormCtlHidden, self.__cache[key])
        frm = self.__get_form()
        if not frm.hasByName(self.__ctl_name):
            comp = self.__lo_inst.create_instance_mcf(
                XComponent, "com.sun.star.form.component.HiddenControl", raise_err=True
            )
            comp.HiddenValue = self.__class__.__qualname__  # type: ignore
            frm.insertByName(self.__ctl_name, comp)

        ctl = FormCtlHidden(frm.getByName(self.__ctl_name), self.__lo_inst)
        self.__cache[key] = ctl
        return ctl

    def __reset(self) -> None:
        self.__cache.clear()

    def __del__(self) -> None:
        with contextlib.suppress(Exception):
            if self.__container_listener and self.__component:
                self.__component.removeContainerListener(self.__container_listener)
            self.__container_listener = None  # type: ignore
