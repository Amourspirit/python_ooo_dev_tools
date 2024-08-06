from __future__ import annotations
from typing import overload
from typing import Tuple, Type, TypeVar
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

_TAdapt = TypeVar("_TAdapt", bound="Adapt")


class Adapt(StyleBase):
    """
    Shape Adapt (only for Shape Text Boxes)

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        fit_height: bool | None = None,
        fit_width: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            fit_height (bool, optional): Expands the width of the object to the width of the text, if the object is smaller than the text.
            fit_width (bool, optional): Expands the height of the object to the height of the text, if the object is smaller than the text.

        Returns:
            None:

        Note:
            Adapt is only available for Text Boxes.

        .. versionadded:: 0.17.3
        """
        super().__init__()
        self.prop_fit_width = fit_width
        # position must be set last, because it may set size too
        self.prop_fit_height = fit_height

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.TextShape",)
        return self._supported_services_values

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAdapt], obj: object) -> _TAdapt: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAdapt], obj: object, **kwargs) -> _TAdapt: ...

    @classmethod
    def from_obj(cls: Type[_TAdapt], obj: object, **kwargs) -> _TAdapt:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Adapt: New instance.
        """

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_fit_width = mProps.Props.get(obj, "TextAutoGrowWidth")
        inst.prop_fit_height = mProps.Props.get(obj, "TextAutoGrowHeight")
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
    def prop_fit_width(self) -> bool | None:
        """Gets/Sets size"""
        return self._get("TextAutoGrowWidth")

    @prop_fit_width.setter
    def prop_fit_width(self, value: bool | None) -> None:
        if value is None:
            self._remove("TextAutoGrowWidth")
            return
        self._set("TextAutoGrowWidth", value)

    @property
    def prop_fit_height(self) -> bool | None:
        """Gets/Sets position"""
        return self._get("TextAutoGrowHeight")

    @prop_fit_height.setter
    def prop_fit_height(self, value: bool | None) -> None:
        if value is None:
            self._remove("TextAutoGrowHeight")
            return
        self._set("TextAutoGrowHeight", value)
