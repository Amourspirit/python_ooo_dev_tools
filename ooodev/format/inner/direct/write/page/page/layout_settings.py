"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload
import uno
from ooo.dyn.style.page_style_layout import PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind


_TLayoutSettings = TypeVar(name="_TLayoutSettings", bound="LayoutSettings")

# The LayoutSettings class is missing a few properties:
# Paper Tray options. I did not see much value in adding it.
#
# Gutter Position.
# I looked for a few hours. Not able to find what properties are being set for gutter position.
# There are likely three properties being set but I can find no trace of what these properties are.
# It seems the properties are not in the PageProperties or PageStyle.


class LayoutSettings(StyleBase):
    """
    Page Layout Setting Style

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        layout: PageStyleLayout | None = None,
        numbers: NumberingTypeEnum | None = None,
        ref_style: str | StyleParaKind | None = None,
        right_gutter: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            layout (PageStyleLayout, optional): Specifies the layout of the page.
            numbers (NumberingTypeEnum, optional): Specifies the default numbering type for this page.
            ref_style (str, StyleParaKind, optional): Specifies the name of the paragraph style that is used as reference of the register mode.
            right_gutter (bool, optional): Specifies that the page gutter shall be placed on the right side of the page.
        """

        super().__init__()
        self.prop_layout = layout
        self.prop_numbers = numbers
        self.prop_ref_style = ref_style
        self.prop_right_gutter = right_gutter

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")
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

    # endregion Overrides

    # region internal methods

    # endregion internal methods

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TLayoutSettings], obj: Any) -> _TLayoutSettings: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TLayoutSettings], obj: Any, **kwargs) -> _TLayoutSettings: ...

    @classmethod
    def from_obj(cls: Type[_TLayoutSettings], obj: Any, **kwargs) -> _TLayoutSettings:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Returns:
            Margins: Instance that represents object margins.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, clazz: LayoutSettings):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                clazz._set(key, val)

        set_prop("PageStyleLayout", inst)
        set_prop("NumberingType", inst)
        set_prop("RegisterParagraphStyle", inst)
        set_prop("RtlGutter", inst)
        return inst

    # endregion from_obj()

    # endregion Static Methods
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PAGE | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_layout(self) -> PageStyleLayout | None:
        """Gets/Sets Page Style Layout value"""
        return self._get("PageStyleLayout")

    @prop_layout.setter
    def prop_layout(self, value: PageStyleLayout | None) -> None:
        if value is None:
            self._remove("PageStyleLayout")
            return
        self._set("PageStyleLayout", value)

    @property
    def prop_numbers(self) -> NumberingTypeEnum | None:
        """Gets/Sets Page Numbering value"""
        pv = cast(int, self._get("NumberingType"))
        return None if pv is None else NumberingTypeEnum(pv)

    @prop_numbers.setter
    def prop_numbers(self, value: NumberingTypeEnum | None) -> None:
        if value is None:
            self._remove("NumberingType")
            return

        self._set("NumberingType", value.value)

    @property
    def prop_ref_style(self) -> str | None:
        """Gets/Sets the name of the paragraph style that is used as reference of the register mode."""
        return self._get("RegisterParagraphStyle")

    @prop_ref_style.setter
    def prop_ref_style(self, value: str | StyleParaKind | None) -> None:
        if value is None:
            self._remove("RegisterParagraphStyle")
            return

        self._set("RegisterParagraphStyle", str(value))

    @property
    def prop_right_gutter(self) -> bool | None:
        """Gets/Sets if the page gutter is placed on the right side of the page."""
        return self._get("RtlGutter")

    @prop_right_gutter.setter
    def prop_right_gutter(self, value: bool | None) -> None:
        if value is None:
            self._remove("RtlGutter")
            return

        self._set("RtlGutter", str(value))
