from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XFrame2

from ooodev.adapter.frame.layout_manager_comp import LayoutManagerComp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.adapter.frame import dispatch_provider_partial
from ooodev.adapter.frame import dispatch_information_provider_partial
from ooodev.adapter.frame import dispatch_provider_interception_partial
from ooodev.adapter.frame import frames_supplier_partial
from ooodev.adapter.task import status_indicator_supplier_partial

if TYPE_CHECKING:
    from com.sun.star.frame import XDispatchRecorderSupplier
    from com.sun.star.uno import XInterface
    from com.sun.star.container import XNameContainer
    from ooodev.utils.type_var import UnoInterface


class Frame2Partial(
    dispatch_provider_partial.DispatchProviderPartial,
    dispatch_information_provider_partial.DispatchInformationProviderPartial,
    dispatch_provider_interception_partial.DispatchProviderInterceptionPartial,
    frames_supplier_partial.FramesSupplierPartial,
    status_indicator_supplier_partial.StatusIndicatorSupplierPartial,
):
    """
    Partial class for XFrame2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFrame2, interface: UnoInterface | None = XFrame2) -> None:
        """
        Constructor

        Args:
            component (XFrame2 ): UNO Component that implements ``com.sun.star.frame.XFrame2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFrame2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        dispatch_provider_partial.DispatchProviderPartial.__init__(self, component=component, interface=None)
        dispatch_information_provider_partial.DispatchInformationProviderPartial.__init__(
            self, component=component, interface=None
        )
        dispatch_provider_interception_partial.DispatchProviderInterceptionPartial.__init__(
            self, component=component, interface=None
        )
        frames_supplier_partial.FramesSupplierPartial.__init__(self, component=component, interface=None)
        status_indicator_supplier_partial.StatusIndicatorSupplierPartial.__init__(
            self, component=component, interface=None  # type: ignore
        )
        self.__component = component

    # region XFrame2
    @property
    def dispatch_recorder_supplier(self) -> XDispatchRecorderSupplier | None:
        """
        Provides access to the dispatch recorder of the frame.

        Such recorder can be used to record dispatch requests.
        The supplier contains a dispatch recorder and provide the functionality to use it for any dispatch object from outside which supports the interface XDispatch.
        A supplier is available only, if recording was enabled.
        That means: if someone wishes to enable recoding on a frame he must set a supplier with a recorder object inside of it.
        Every user of dispatches has to check then if such supplier is available at this frame property.
        If value of this property is ``None`` they must call ``XDispatch.dispatch()`` on the original dispatch object.
        If it's a valid value he must use the supplier by calling the method ``XDispatchRecorderSupplier.dispatchAndRecord()`` with the original dispatch object as argument.

        Note:
            It's not recommended to cache an already gotten supplier.
            Because there exist no possibility to check for enabled/disabled recording then.

        **since**

            OOo 1.1.2
        """
        return self.__component.DispatchRecorderSupplier

    @dispatch_recorder_supplier.setter
    def dispatch_recorder_supplier(self, value: XDispatchRecorderSupplier) -> None:
        self.__component.DispatchRecorderSupplier = value

    @property
    def layout_manager(self) -> LayoutManagerComp:
        """
        Provides access to the LayoutManager of the frame.

        When setting can be instance of ``XLayoutManager2`` or ``LayoutManagerComp``.
        """

        lm = self.__component.LayoutManager
        if lm is None:
            return None  # type: ignore
        return LayoutManagerComp(lm)

    @layout_manager.setter
    def layout_manager(self, value: XInterface | LayoutManagerComp) -> None:
        if mInfo.Info.is_instance(value, LayoutManagerComp):
            self.__component.LayoutManager = value.component
        else:
            self.__component.LayoutManager = value  # type: ignore

    @property
    def title(self) -> str:
        """
        If possible it sets/gets the UI title on/from the frame container window

        It depends from the type of the frame container window.
        If it is a system task window all will be OK.
        Otherwise the title can't be set.
        Setting/getting of the pure value of this property must be possible in every case.
        Only showing on the UI can be fail.
        """
        return self.__component.Title

    @title.setter
    def title(self, value: str) -> None:
        self.__component.Title = value

    @property
    def user_defined_attributes(self) -> NameContainerComp | None:
        """
        Contains user defined attributes.

        When setting can be instance of ``XNameContainer`` or ``NameContainerComp``.

        Returns:
            NameContainerComp: User defined attributes.
        """
        attribs = self.__component.UserDefinedAttributes
        if attribs is None:
            return None
        return NameContainerComp(attribs)

    @user_defined_attributes.setter
    def user_defined_attributes(self, value: XNameContainer | NameContainerComp) -> None:
        if mInfo.Info.is_instance(value, NameContainerComp):
            self.__component.UserDefinedAttributes = value.component
        else:
            self.__component.UserDefinedAttributes = value  # type: ignore

    # endregion XFrame2


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel

    builder = DefaultBuilder(component)
    builder.merge(dispatch_provider_partial.get_builder(component))
    builder.merge(dispatch_information_provider_partial.get_builder(component))
    builder.merge(dispatch_provider_interception_partial.get_builder(component))
    builder.merge(frames_supplier_partial.get_builder(component))
    builder.merge(status_indicator_supplier_partial.get_builder(component))
    builder.auto_add_interface("com.sun.star.frame.XFrame2", False)
    return builder
