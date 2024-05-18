from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_dialog_element_partial import UnoControlDialogElementPartial
from ooodev.units.unit_app_font_height import UnitAppFontHeight
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlModel
    from ooodev.units.unit_obj import UnitT


class ModelDialogElementPartial:

    def __new__(cls, model: UnoControlModel, *args, **kwargs):
        # model = kwargs.get("model", None)
        # if model is None:
        #     model = args[0]
        builder = _get_builder(model)
        clz = builder.get_class_type(
            name="ooodev.dialog.dl_control.model.model_dialog_element_partial.ModelDialogElementPartial",
            base_class=cls,
            set_mod_name=False,
        )
        return super().__new__(clz, *args, **kwargs)

    def __init__(self, model: UnoControlModel) -> None:
        """
        Constructor

        Args:
            component (UnoControlModel): UNO Component that implements ``com.sun.star.awt.UnoControlModel`` service.
        """
        if isinstance(self, UnoControlDialogElementPartial):
            UnoControlDialogElementPartial.__init__(self, model)

    # region Properties
    if TYPE_CHECKING:
        # The properties from this point are optional. The builder will add them if needed.
        # There are include here for type hinting only.
        @property
        def height(self) -> UnitAppFontHeight:
            """
            Gets/Sets the height of the control.

            When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

            Returns:
                UnitAppFontHeight: Height of the control.
            """
            ...

        @height.setter
        def height(self, value: int | UnitT) -> None: ...

        @property
        def name(self) -> str:
            """
            Gets/Sets the name of the control.
            """
            ...

        @name.setter
        def name(self, value: str) -> None: ...

        @property
        def x(self) -> UnitAppFontX:
            """
            Gets/Sets the horizontal position of the control.

            When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

            Returns:
                UnitAppFontX: Horizontal position of the control.
            """
            ...

        @x.setter
        def x(self, value: int | UnitT) -> None: ...
        @property
        def y(self) -> UnitAppFontY:
            """
            Gets/Sets the vertical position of the control.

            When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

            Returns:
                UnitAppFontY: Vertical position of the control.
            """
            ...

        @y.setter
        def y(self, value: int | UnitT) -> None: ...

        @property
        def step(self) -> int: ...
        @step.setter
        def step(self, value: int) -> None: ...

        @property
        def tab_index(self) -> int: ...

        @tab_index.setter
        def tab_index(self, value: int) -> None: ...

        @property
        def tag(self) -> str:
            """
            Gets/Sets the tag of the control.
            """
            ...

        @tag.setter
        def tag(self, value: str) -> None: ...

        @property
        def width(self) -> UnitAppFontWidth:
            """
            Gets/Sets the width of the control.

            When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

            Returns:
                UnitAppFontWidth: Width of the control.
            """
            ...

        @width.setter
        def width(self, value: int | UnitT) -> None: ...

    # endregion Properties


def _get_builder(model: UnoControlModel) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(model)
    if hasattr(model, "Name"):
        builder.add_import(
            "ooodev.adapter.awt.uno_control_dialog_element_partial.UnoControlDialogElementPartial",
            optional=False,
            init_kind=1,
            check_kind=0,
        )
    return builder
