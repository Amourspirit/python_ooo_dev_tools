"""
Module for Cell Properties Cell Back Color.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.loader import lo as mLo
from ooodev.utils import color as mColor
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.cell_background_color_props import CellBackgroundColorProps

# endregion Import

_TColor = TypeVar(name="_TColor", bound="Color")


class Color(StyleBase):
    """
    Class for Cell Properties Back Color.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_background`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(self, color: mColor.Color = mColor.Color(-1)) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Color such as ``CommonColor.LIGHT_BLUE``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_direct_cell_background`
        """
        super().__init__()
        self.prop_color = color

    # endregion Init

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.table.CellProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    # endregion overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TColor], obj: Any) -> _TColor: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TColor], obj: Any, **kwargs) -> _TColor: ...

    @classmethod
    def from_obj(cls: Type[_TColor], obj: Any, **kwargs) -> _TColor:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            BackColor: ``BackColor`` instance that represents ``obj`` Back Color properties.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        color = mProps.Props.get(obj, inst._props.color, None)
        if inst._props.is_transparent:
            bg = mProps.Props.get(obj, inst._props.is_transparent, None)
        else:
            bg = None
        if color is not None:
            inst._set(inst._props.color, int(color))
        if bg is not None:
            inst._set(inst._props.is_transparent, bool(bg))

        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CELL
        return self._format_kind_prop

    @property
    def prop_color(self) -> mColor.Color:
        """Gets/Sets color"""
        return self._get(self._props.color)

    @prop_color.setter
    def prop_color(self, value: mColor.Color):
        if value >= 0:
            self._set(self._props.color, value)
            if self._props.is_transparent:
                self._set(self._props.is_transparent, False)
        else:
            self._set(self._props.color, -1)
            if self._props.is_transparent:
                self._set(self._props.is_transparent, True)

    @property
    def _props(self) -> CellBackgroundColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellBackgroundColorProps(
                color="CellBackColor", is_transparent="IsCellBackgroundTransparent"
            )
        return self._props_internal_attributes

    @property
    def empty(self: _TColor) -> _TColor:  # type: ignore[misc]
        """Gets BackColor empty."""
        try:
            return self._empty_inst
        except AttributeError:
            self._empty_inst = self.__class__(_cattribs=self._get_internal_cattribs())  # type: ignore
            self._empty_inst._is_default_inst = True
        return self._empty_inst

    default = empty  # for protocol compatibility with other classes
    # endregion Properties
