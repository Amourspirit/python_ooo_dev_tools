"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar
import uno
from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.intensity import Intensity as Intensity
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

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
        value = Intensity(int(value))

        super().__init__(FillTransparence=value.value)

    # region Internal Methods

    # endregion Internal Methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.style.ParagraphStyle",
        )

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
        return mProps.Props.has(obj, "FillTransparence")

    # endregion Overrides

    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: object) -> _TTransparency:
        """
        Gets instance from object

        Args:
            obj (object): Object that implements ``com.sun.star.drawing.FillProperties`` service

        Returns:
            Gradient: Instance that represents Gradient color.
        """
        # this nu is only used to get Property Name

        nu = super(Transparency, cls).__new__(cls)
        nu.__init__()
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        tp = cast(int, mProps.Props.get(obj, "FillTransparence", None))
        inst = super(Transparency, cls).__new__(cls)
        if tp is None:
            inst.__init__(value=0)
        else:
            inst.__init__(value=tp)
        return inst

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL

    @property
    def prop_value(self) -> Intensity:
        """Gets/Sets Transparency value"""
        pv = cast(int, self._get("FillTransparence"))
        return Intensity(pv)

    @prop_value.setter
    def prop_value(self, value: Intensity | int) -> None:
        val = Intensity(int(value))
        self._set("FillTransparence", val.value)

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