"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.data_type.intensity import Intensity
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.transparent_transparency_props import TransparentTransparencyProps

_TTransparency = TypeVar(name="_TTransparency", bound="Transparency")


class Transparency(StyleBase):
    """
    Fill Transparency

    .. seealso::

        - :ref:`help_writer_format_direct_para_transparency`

    .. versionadded:: 0.9.0
    """

    def __init__(self, value: Intensity | int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_transparency`
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
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextContent",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def _is_valid_obj(self, obj: Any) -> bool:
        return mProps.Props.has(obj, self._props.transparence)

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: Any) -> _TTransparency: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: Any, **kwargs) -> _TTransparency: ...

    @classmethod
    def from_obj(cls: Type[_TTransparency], obj: Any, **kwargs) -> _TTransparency:
        """
        Gets instance from object

        Args:
            obj (object): Object that implements ``com.sun.star.drawing.FillProperties`` service

        Returns:
            Gradient: Instance that represents Gradient color.
        """
        # this nu is only used to get Property Name
        # pylint: disable=protected-access
        nu = cls(value=0, **kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        tp = cast(int, mProps.Props.get(obj, nu._props.transparence, None))
        result = nu if tp is None else cls(value=tp, **kwargs)
        result.set_update_obj(obj)
        return result

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

    @property
    def default(self: _TTransparency) -> _TTransparency:  # type: ignore[misc]
        """Gets Transparency Default."""
        try:
            return self._DEFAULT_INST
        except AttributeError:
            # pylint: disable=unexpected-keyword-arg
            inst = self.__class__(value=0, _cattribs=self._get_internal_cattribs())
            inst._is_default_inst = True
            self._DEFAULT_INST = inst
        return self._DEFAULT_INST
