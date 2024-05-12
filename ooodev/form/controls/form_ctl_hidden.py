from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.form.form_component_type import FormComponentType

from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.loader import lo as mLo
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.beans.property_set_option_partial import PropertySetOptionPartial
from ooodev.adapter.beans.property_access_partial import PropertyAccessPartial
from ooodev.adapter.beans.property_container_partial import PropertyContainerPartial
from ooodev.adapter.container.named_partial import NamedPartial
from ooodev.adapter.io.persist_object_partial import PersistObjectPartial
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial
from ooodev.adapter.lang.type_provider_partial import TypeProviderPartial
from ooodev.adapter.util.cloneable_partial import CloneablePartial
from ooodev.adapter.beans.fast_property_set_partial import FastPropertySetPartial
from ooodev.adapter.beans.multi_property_set_partial import MultiPropertySetPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlHidden(
    LoInstPropsPartial,
    PropPartial,
    NamedPartial,
    PropertySetOptionPartial,
    PropertyAccessPartial,
    PropertyContainerPartial,
    FastPropertySetPartial,
    MultiPropertySetPartial,
    PersistObjectPartial,
    ServiceInfoPartial,
    TypeProviderPartial,
    PropertyChangeImplement,
    PropertiesChangeImplement,
    VetoableChangeImplement,
):
    """``com.sun.star.form.component.HiddenControl`` control"""

    def __init__(self, ctl: XComponent, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.HiddenControl`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            None:

        Note:
            If the :ref:`LoContext <ooodev.utils.context.lo_context.LoContext>` manager is use before this class is instantiated,
            then the Lo instance will be set using the current Lo instance. That the context manager has set.
            Generally speaking this means that there is no need to set ``lo_inst`` when instantiating this class.

        See Also:
            :ref:`ooodev.form.Forms`.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        self._ctl = ctl
        PropPartial.__init__(self, component=self._ctl, lo_inst=self.lo_inst)
        NamedPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        PropertySetOptionPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        PropertyAccessPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        PropertyContainerPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        FastPropertySetPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        MultiPropertySetPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        PersistObjectPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        ServiceInfoPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        TypeProviderPartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        CloneablePartial.__init__(self, component=self._ctl, interface=None)  # type: ignore
        trigger_args = self._get_generic_args()
        PropertyChangeImplement.__init__(self, component=self._ctl, trigger_args=trigger_args)  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self._ctl, trigger_args=trigger_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=self._ctl, trigger_args=trigger_args)  # type: ignore

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.HIDDEN_CONTROL

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    # region XCloneable
    def create_clone(self) -> FormCtlHidden:
        """
        Creates a clone of the object.

        Returns:
            FormCtlHidden: The clone.

        Note:
            The returned object will have the same Name as the original object and will not be inserted into the form.
        """
        cpy = self._ctl.createClone()  # type: ignore
        return type(self)(cpy, lo_inst=self.lo_inst)

    # endregion XCloneable

    # region Properties

    @property
    def component_type(self) -> int:
        """
        Gets the form component type.

        The return value is a ``com.sun.star.form.FormComponentType`` constant.

        Returns:
            int: Form component type

        .. versionadded:: 0.14.1
        """
        return FormComponentType.HIDDENCONTROL

    @property
    def hidden_value(self) -> str:
        """
        Gets the hidden value.

        Returns:
            str: Hidden value
        """
        return self._ctl.HiddenValue  # type: ignore

    @hidden_value.setter
    def hidden_value(self, value: str) -> None:
        """
        Sets the hidden value.

        Args:
            value (str): Hidden value
        """
        self._ctl.HiddenValue = value  # type: ignore

    @property
    def tag(self) -> str:
        """
        Gets the tag.

        Returns:
            str: Tag
        """
        return self._ctl.Tag  # type: ignore

    @tag.setter
    def tag(self, value: str) -> None:
        """
        Sets the tag.

        Args:
            value (str): Tag
        """
        self._ctl.Tag = value  # type: ignore

    # endregion Properties
