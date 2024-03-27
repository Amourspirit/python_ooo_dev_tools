from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.view import XFormLayerAccess

from ooodev.adapter.view.control_access_partial import ControlAccessPartial
from ooodev.adapter.form.runtime.form_controller_comp import FormControllerComp

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.form import XForm


class FormLayerAccessPartial(ControlAccessPartial):
    """
    Partial class for XFormLayerAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFormLayerAccess, interface: UnoInterface | None = XFormLayerAccess) -> None:
        """
        Constructor

        Args:
            component (XFormLayerAccess): UNO Component that implements ``com.sun.star.view.XFormLayerAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFormLayerAccess``.
        """
        ControlAccessPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XFormLayerAccess
    def get_form_controller(self, form: XForm) -> FormControllerComp:
        """
        Gets the ``FormControllerComp`` instance which operates on a given form.

        A form controller is a component which controls the user interaction with the form layer, as long as the form is not in design mode.
        """
        return FormControllerComp(self.__component.getFormController(form))

    def is_form_design_mode(self) -> bool:
        """
        Gets whether the view's form layer is currently in design or alive mode

        Note: This is a convenience method. In the user interface, the design mode is coupled with the .uno:SwitchControlDesignMode feature (see com.sun.star.frame.XDispatchProvider), and asking for the current mode is the same as asking for the state of this feature.
        """
        return self.__component.isFormDesignMode()

    def set_form_design_mode(self, design_mode: bool) -> None:
        """
        Sets whether the view's form layer is currently in design or alive mode

        Note:
            This is a convenience method. In the user interface, the design mode is coupled with the ``.uno:SwitchControlDesignMode`` feature
            (see ``com.sun.star.frame.XDispatchProvider``), and changing the current mode is the same as dispatching this feature URL.
        """
        self.__component.setFormDesignMode(design_mode)

    # endregion XFormLayerAccess
