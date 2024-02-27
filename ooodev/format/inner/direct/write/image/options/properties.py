from __future__ import annotations
from typing import overload
from typing import Any, Tuple, Type, TypeVar
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.image_options_properties import ImageOptionsProperties

_TProperties = TypeVar(name="_TProperties", bound="Properties")


class Properties(StyleBase):
    """
    Image Options Properties.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        printable: bool = True,
    ) -> None:
        """
        Constructor

        Args:
            printable (bool, optional): Specifies if Frame can be printed. Default ``True``.
        """
        super().__init__()

        self.prop_printable = printable

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
            )
        return self._supported_services_values

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TProperties], obj: Any) -> _TProperties: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TProperties], obj: Any, **kwargs) -> _TProperties: ...

    @classmethod
    def from_obj(cls: Type[_TProperties], obj: Any, **kwargs) -> _TProperties:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Properties: Instance that represents Image Properties.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop_name in inst._props:
            if prop_name:
                inst._set(prop_name, mProps.Props.get(obj, prop_name))
        # prev, next not currently working
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
            self._format_kind_prop = FormatKind.IMAGE
        return self._format_kind_prop

    @property
    def prop_printable(self) -> bool | None:
        """Gets/Sets print value"""
        return self._get(self._props.printable)

    @prop_printable.setter
    def prop_printable(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.printable)
            return
        self._set(self._props.printable, value)

    @property
    def _props(self) -> ImageOptionsProperties:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImageOptionsProperties(printable="Print")
        return self._props_internal_attributes

    # endregion Properties
