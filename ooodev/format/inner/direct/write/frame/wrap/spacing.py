# region Import
from __future__ import annotations
from typing import Any, Tuple

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.common.abstract.abstract_padding import AbstractPadding

# endregion Import


class Spacing(AbstractPadding):
    """
    Frame Spacing

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.Style",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        """
        Gets if ``obj`` supports one of the services required by style class

        Args:
            obj (Any): UNO object that must have requires service

        Returns:
            bool: ``True`` if has a required service; Otherwise, ``False``
        """
        if super()._is_valid_obj(obj):
            return True
        # check if obj has matching property
        # Some objects such as 'com.sun.star.drawing.shape' sometime support this style.
        # Such is the case when a shape is added to a Writer drawing page.
        # Assume if on attribute matches then it all is a match.
        return hasattr(obj, self._props.left)

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="LeftMargin", top="TopMargin", right="RightMargin", bottom="BottomMargin"
            )
        return self._props_internal_attributes

    # endregion Overrides
