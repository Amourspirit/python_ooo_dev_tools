"""
Module for Cell Properties Cell Back Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload, Type, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import props as mProps
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import color as mColor
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

_TColor = TypeVar(name="_TColor", bound="Color")


class Color(StyleBase):
    """
    Class for Cell Properties Back Color.

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: mColor.Color = -1) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Color such as ``CommonColor.LIGHT_BLUE``

        Returns:
            None:
        """

        init_vals = {}
        if color >= 0:
            init_vals["CellBackColor"] = color
            init_vals["IsCellBackgroundTransparent"] = False
        else:
            init_vals["CellBackColor"] = -1
            init_vals["IsCellBackgroundTransparent"] = True

        super().__init__(**init_vals)

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.table.CellProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.table.CellProperties",)

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TColor], obj: object) -> _TColor:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TColor], obj: object, **kwargs) -> _TColor:
        ...

    @classmethod
    def from_obj(cls: Type[_TColor], obj: object, **kwargs) -> _TColor:
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

        color = mProps.Props.get(obj, "CellBackColor", None)
        bg = mProps.Props.get(obj, "IsCellBackgroundTransparent", None)
        if not color is None:
            inst._set("CellBackColor", int(color))
        if not bg is None:
            inst._set("IsCellBackgroundTransparent", bool(bg))

        return inst

    # endregion from_obj()
    # region set styles

    # endregion set styles

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
        return self._get("CellBackColor")

    @prop_color.setter
    def prop_color(self, value: mColor.Color):
        if value >= 0:
            self._set("CellBackColor", value)
            self._set("IsCellBackgroundTransparent", False)
        else:
            self._set("CellBackColor", -1)
            self._set("IsCellBackgroundTransparent", True)

    @property
    def prop_is_bg_transparent(self) -> bool:
        """Gets/Sets if Background color is Transparent"""
        return self._get("IsCellBackgroundTransparent")

    @prop_is_bg_transparent.setter
    def prop_is_bg_transparent(self, value: bool):
        self._set("IsCellBackgroundTransparent", value)

    @static_prop
    def empty() -> Color:  # type: ignore[misc]
        """Gets BackColor empty. Static Property."""
        try:
            return Color._EMPTY_INST
        except AttributeError:
            Color._EMPTY_INST = Color()
            Color._EMPTY_INST._is_default_inst = True
        return Color._EMPTY_INST
