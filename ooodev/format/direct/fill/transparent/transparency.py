"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload
import uno
from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.intensity import Intensity as Intensity
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.transparent_transparency_props import TransparentTransparencyProps

_TTransparency = TypeVar(name="_TTransparency", bound="Transparency")


class Transparency(StyleBase):
    """
    Fill Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(self, value: Intensity | int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.
        """
        super().__init__()
        self.prop_value = value

    # region Internal Methods

    # endregion Internal Methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.text.TextContent",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
            )
        return self._supported_services_values

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def _is_valid_obj(self, obj: object) -> bool:
        return mProps.Props.has(obj, self._props.transparence)

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: object) -> _TTransparency:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: object, **kwargs) -> _TTransparency:
        ...

    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: object, **kwargs) -> _TTransparency:
        """
        Gets instance from object

        Args:
            obj (object): Object that implements ``com.sun.star.drawing.FillProperties`` service

        Returns:
            Gradient: Instance that represents Gradient color.
        """
        # this nu is only used to get Property Name

        nu = cls(value=0, **kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        tp = cast(int, mProps.Props.get(obj, nu._props.transparence, None))
        if tp is None:
            return nu
        else:
            return cls(value=tp, **kwargs)

    # endregion from_obj()
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL
        return self._format_kind_prop

    @property
    def prop_value(self) -> Intensity:
        """Gets/Sets Transparency value"""
        pv = cast(int, self._get(self._props.transparence))
        return Intensity(pv)

    @prop_value.setter
    def prop_value(self, value: Intensity | int) -> None:
        val = Intensity(int(value))
        self._set(self._props.transparence, val.value)

    @property
    def _props(self) -> TransparentTransparencyProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = TransparentTransparencyProps(transparence="FillTransparence")
        return self._props_internal_attributes

    @static_prop
    def default() -> Transparency:  # type: ignore[misc]
        """Gets Transparency Default. Static Property."""
        try:
            return Transparency._DEFAULT_INST
        except AttributeError:
            inst = Transparency(0)
            inst._is_default_inst = True
            Transparency._DEFAULT_INST = inst
        return Transparency._DEFAULT_INST
