from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.awt import XTabControllerModel

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.awt import XControlModel


class TabControllerModelPartial:
    """
    Partial Class XTabControllerModel.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTabControllerModel, interface: UnoInterface | None = XTabControllerModel) -> None:
        """
        Constructor

        Args:
            component (XTabControllerModel): UNO Component that implements ``com.sun.star.io.XTabControllerModel`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTabControllerModel``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTabControllerModel
    def get_control_models(self) -> Tuple[XControlModel, ...]:
        """
        Returns the control models.

        Returns:
            tuple[XControlModel]: The controls.
        """
        return self.__component.getControlModels()

    # both getGroup and getGroupByName are not implemented here.
    # They both contain OUT arguments, which are not supported by Python.
    # I put a post on StackOverflow about this:
    # https://stackoverflow.com/questions/77749118/how-to-get-out-parameters-from-an-api-that-has-multiple-our-args-in-function-cal

    def get_group(self, idx: int) -> tuple:
        """
        Gets the group for the specified index.

        Args:
            idx (int): The group index.


        Returns:
            tuple: Results as a tuple of ``3`` elements. Element at index ``1`` is a tuple of ``XControlModel`` objects.
            Element ``2`` is a tuple of ``str`` objects. Element at index ``2`` is a ``str`` containing name.
            A tuple is returned even if the index is not found.

        Note:
            The API documentation shows a return value of ``void`` for ``getGroup()``.
            This is incorrect for python. The return value is a tuple with 3 elements.

            See the LibreOffice API documentation for `getGroup() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTabControllerModel.html#addcec8bffb58a80e157cf5b83feb8286>`__
            for more information.
        """
        # even though the API documentation shows a return value of ``void``, it is actually a tuple with 3 elements.
        # the out arguments Group and Name must be passed as None.
        return self.__component.getGroup(idx, None, None)  # type: ignore

    def get_group_by_name(self, name: str) -> tuple:
        """
        Gets the group for the specified name.

        Args:
            name (str): The name.

        Returns:
            tuple: Results as a tuple of ``2`` elements. Element at index ``1`` is a tuple of ``XControlModel`` objects.
            A tuple is returned even if the index is not found.

        Note:
            The API documentation shows a return value of ``void`` for ``getGroupByName()``.
            This is incorrect for python. The return value is a tuple with 2 elements.

            See the LibreOffice API documentation for `getGroupByName() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTabControllerModel.html#a5784398579a39405d834461d75264adb>`__
            for more information.
        """
        # even though the API documentation shows a return value of ``void``, it is actually a tuple with 2 elements.
        # the out argument Group must be passed as None.
        return self.__component.getGroupByName(name, None)  # type: ignore

    def get_group_control(self) -> bool:
        """
        Returns the group control.

        Returns:
            bool: The group control.
        """
        return self.__component.getGroupControl()

    def get_group_count(self) -> int:
        """
        Returns the group count.

        Returns:
            int: The group count.
        """
        return self.__component.getGroupCount()

    def set_control_models(self, controls: tuple[XControlModel, ...]) -> None:
        """
        Sets the control models.

        Args:
            controls (tuple[XControlModel]): The controls.
        """
        self.__component.setControlModels(controls)

    def set_group(self, group: tuple[XControlModel, ...], name: str) -> None:
        """
        Sets the group.

        Args:
            group (int): The group.
            name (str): The name.
        """
        self.__component.setGroup(group, name)

    # endregion XTabControllerModel
