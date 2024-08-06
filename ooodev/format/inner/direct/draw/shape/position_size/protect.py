from __future__ import annotations
from typing import overload
from typing import Tuple, Type, TypeVar
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

_TProtect = TypeVar("_TProtect", bound="Protect")


class Protect(StyleBase):
    """
    Shape Protection

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        position: bool | None = None,
        size: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            position (bool, optional): Specifies position protection.
            size (bool, optional): Specifies size protection.

        Returns:
            None:

        Note:
            If ``position`` is ``True``, ``size`` is ``True`` too.
            Setting ``size`` to ``False`` will be ignored if ``position`` is ``True``.

        .. versionadded:: 0.17.3
        """
        super().__init__()
        self.prop_size = size
        # position must be set last, because it may set size too
        self.prop_position = position

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.drawing.Shape",
                "com.sun.star.presentation.Shape",
            )
        return self._supported_services_values

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TProtect], obj: object) -> _TProtect: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TProtect], obj: object, **kwargs) -> _TProtect: ...

    @classmethod
    def from_obj(cls: Type[_TProtect], obj: object, **kwargs) -> _TProtect:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Protect: Instance that represents shape protection.
        """

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_size = mProps.Props.get(obj, "SizeProtect")
        inst.prop_position = mProps.Props.get(obj, "MoveProtect")
        return inst

    # endregion from_obj()
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.SHAPE | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_size(self) -> bool | None:
        """Gets/Sets size"""
        return self._get("SizeProtect")

    @prop_size.setter
    def prop_size(self, value: bool | None) -> None:
        if value is False:
            if self.prop_position is True:
                # size cannot be unprotected if position is protected
                return
        if value is None:
            self._remove("SizeProtect")
            return
        self._set("SizeProtect", value)

    @property
    def prop_position(self) -> bool | None:
        """Gets/Sets position"""
        return self._get("MoveProtect")

    @prop_position.setter
    def prop_position(self, value: bool | None) -> None:
        if value is None:
            self._remove("MoveProtect")
            return
        self._set("MoveProtect", value)
        if value:
            # if position is protected, size is protected too
            self._set("SizeProtect", value)
