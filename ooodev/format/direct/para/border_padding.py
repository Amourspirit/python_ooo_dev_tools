"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ...kind.format_kind import FormatKind
from ..common.abstract_padding import AbstractPadding
from ..common.border_props import BorderProps


class BorderPadding(AbstractPadding):
    """
    Paragraph BorderPadding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region methods

    @staticmethod
    def from_obj(obj: object) -> BorderPadding:
        """
        Gets BorderPadding instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            BorderPadding: BorderPadding that represents ``obj`` padding.
        """
        inst = BorderPadding()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        if inst._is_valid_obj(obj):
            inst._set(inst._border.left, int(mProps.Props.get(obj, inst._border.left)))
            inst._set(inst._border.right, int(mProps.Props.get(obj, inst._border.right)))
            inst._set(inst._border.top, int(mProps.Props.get(obj, inst._border.top)))
            inst._set(inst._border.bottom, int(mProps.Props.get(obj, inst._border.bottom)))
        else:
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])
        return inst

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def _border(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="LeftBorderDistance",
                top="TopBorderDistance",
                right="RightBorderDistance",
                bottom="BottomBorderDistance",
            )
        return self.__border_properties

    @static_prop
    def default() -> BorderPadding:  # type: ignore[misc]
        """Gets BorderPadding default. Static Property."""
        if BorderPadding._DEFAULT is None:
            inst = BorderPadding()
            inst._set(inst._border.bottom, 0)
            inst._set(inst._border.left, 0)
            inst._set(inst._border.right, 0)
            inst._set(inst._border.top, 0)
            BorderPadding._DEFAULT = inst

        return BorderPadding._DEFAULT

    # endregion properties
