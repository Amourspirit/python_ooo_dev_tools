from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.frame import layout_manager2_partial
from ooodev.mock import mock_g
from ooodev.utils import props as mProps
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial


if TYPE_CHECKING:
    from com.sun.star.frame import LayoutManager as UnoLayoutManager  # service
    from ooodev.gui.menu.menu_bar import MenuBar
    from ooodev.loader.inst.lo_inst import LoInst


class _LayoutManager(ComponentProp):

    # region Dunder Methods
    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:  # type: ignore
            return True
        return self.component == other.component  # type: ignore

    # endregion Dunder Methods

    # region override base methods
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.frame.LayoutManager",)

    # endregion override base methods

    # region context manage
    def __enter__(self) -> _LayoutManager:
        # LayoutManagerPartial has lock and unlock
        self.lock()  # type: ignore
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.unlock()  # type: ignore

    def __instancecheck__(cls, instance):
        # Customize isinstance checks here
        return hasattr(instance, "my_custom_attribute")

    # endregion context manage

    # region Methods

    def get_menu_bar(self) -> MenuBar | None:
        """
        Get the MenuBar component.

        Returns:
            MenuBar: MenuBar component.

        See Also:
            - :ref:`help_common_gui_menus_menu_bar`
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=redefined-outer-name
        from ooodev.gui.menu.menu_bar import MenuBar

        mb_el = self.get_element(MenuBar.NODE)  # type: ignore
        if mb_el is None:
            return None
        mb = mProps.Props.get(mb_el, "XMenuBar", default=None)
        if mb is None:
            return None
        return MenuBar(component=mb, lo_inst=self.lo_inst)  # type: ignore

    # endregion Methods

    # region Properties
    @property
    def __class__(self):
        # pretend to be a LayoutManager class
        return LayoutManager

    # endregion Properties


class LayoutManager(
    _LayoutManager,
    layout_manager2_partial.LayoutManager2Partial,
    CompDefaultsPartial,
    LoInstPropsPartial,
    # child_partial.ChildPartial,
):
    """
    Class for managing LayoutManager Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~.InterfacePartial.is_supported_interface` methods which is always available in this class.

    Note:
        This component is also a context manager and will lock and unlock the component when used in a ``with`` statement.

        .. code-block:: python

            comp = doc.get_frame_comp()
            with comp.layout_manager:
                # do something with the layout manager while locked
    """

    # pylint: disable=unused-argument
    def __new__(cls, component: Any, **kwargs):
        lo_inst = kwargs.get("lo_inst", None)
        builder = get_builder(component=component)
        if lo_inst is not None:
            builder.lo_inst = lo_inst
        # builder.add_class_property("__class__", LayoutManager)
        builder_helper.builder_add_comp_defaults(builder)
        builder_helper.builder_add_property_change_implement(builder)
        builder_helper.builder_add_property_veto_implement(builder)
        builder_helper.builder_add_lo_inst_props_partial(builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.gui.comp.layout_manager.LayoutManager",
            base_class=_LayoutManager,
        )
        return inst

    def __init__(self, component: Any, *, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.ConfigurationAccess`` service.
            lo_inst (LoInst, optional): Instance of the Lo. Defaults to ``None``.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> UnoLayoutManager:
        """LayoutManager Component"""
        # pylint: disable=no-member
        return cast("UnoLayoutManager", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)
    builder_helper.builder_add_component_prop(builder)
    builder.merge(layout_manager2_partial.get_builder(component))
    builder.auto_add_interface("com.sun.star.awt.XWindowListener")
    builder.auto_add_interface("com.sun.star.beans.XPropertySet")
    builder.auto_add_interface("com.sun.star.beans.XMultiPropertySet")
    builder.auto_add_interface("com.sun.star.beans.XFastPropertySet")
    return builder


if mock_g.FULL_IMPORT:
    from ooodev.gui.menu.menu_bar import MenuBar
