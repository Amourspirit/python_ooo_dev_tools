# region Import
from __future__ import annotations
from ast import Tuple
from typing import Any, cast
import uno
from ooo.dyn.style.page_style_layout import PageStyleLayout
from ooo.dyn.style.numbering_type import NumberingTypeEnum

from ooodev.format.inner.common.abstract.abstract_document import AbstractDocument
from ooodev.format.inner.direct.write.page.page.layout_settings import LayoutSettings as InnerLayoutSettings
from ooodev.format.inner.modify.write.page.page_style_base_multi import PageStyleBaseMulti
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

# endregion Import


class _GutterPosition(AbstractDocument):
    """
    Gutter Position

    Gutter Position is a document level setting.
    Therefore No matter what style it is set in will apply to all styles.
    """

    def __init__(self, pos_left: bool = True) -> None:
        super().__init__()
        # True means left, False means top
        self._pos_left = pos_left

    def _supported_services(self) -> Tuple[str, ...]:  # type: ignore
        try:
            return self._supported_services_values  # type: ignore
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.DocumentSettings",)
        return self._supported_services_values  # type: ignore

    def copy(self, **kwargs) -> _GutterPosition:
        cp = super().copy(**kwargs)
        cp.prop_pos_left = self.prop_pos_left
        return cp

    def apply(self, obj: Any, **kwargs) -> None:
        try:
            props = self.get_doc_settings()
        except Exception as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply() Failed to get document settings")
            mLo.Lo.print(f"  {e}")
            return
        super().apply(obj=props, override_dv={"GutterAtTop": not self._pos_left})

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _GutterPosition:
        inst = cls(**kwargs)
        props = inst.get_doc_settings()
        inst._pos_left = not bool(mProps.Props.get(props, "GutterAtTop", False))
        return inst

    @property
    def prop_pos_left(self) -> bool:
        return self._pos_left

    @prop_pos_left.setter
    def prop_pos_left(self, value: bool) -> None:
        self._pos_left = value


class LayoutSettings(PageStyleBaseMulti):
    """
    Page Layout Setting style

    .. seealso::

        - :ref:`help_writer_format_modify_page_page`

    .. versionadded:: 0.9.0

    .. versionchanged:: 0.9.7
        Added ``gutter_pos_left`` parameter.
    """

    def __init__(
        self,
        *,
        layout: PageStyleLayout | None = None,
        numbers: NumberingTypeEnum | None = None,
        ref_style: str | StyleParaKind | None = None,
        right_gutter: bool | None = None,
        gutter_pos_left: bool | None = None,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            layout (PageStyleLayout, optional): Specifies the layout of the page.
            numbers (NumberingTypeEnum, optional): Specifies the default numbering type for this page.
            ref_style (str, StyleParaKind, optional): Specifies the name of the paragraph style that is used as
                reference of the register mode.
            right_gutter (bool, optional): Specifies that the page gutter shall be placed on the right side of the page.
            gutter_pos_left (bool, optional): Specifies if the gutters position is Left or Top of page. If ``True`` then gutter is left, if ``False`` gutter is top.
            style_name (WriterStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_page`
        """

        direct = InnerLayoutSettings(layout=layout, numbers=numbers, ref_style=ref_style, right_gutter=right_gutter)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())
        if gutter_pos_left is not None:
            gutter_pos = _GutterPosition(pos_left=gutter_pos_left)
            self._set_style("gutter_pos", gutter_pos)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> LayoutSettings:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (WriterStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            LayoutSettings: ``LayoutSettings`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerLayoutSettings.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        inst._set_style("gutter_pos", _GutterPosition.from_obj(doc))
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerLayoutSettings:
        """Gets/Sets Inner Layout Settings instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerLayoutSettings, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerLayoutSettings) -> None:
        if not isinstance(value, InnerLayoutSettings):
            raise TypeError(f'Expected type of InnerLayoutSettings, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())

    @property
    def _inner_gutter_pos(self) -> _GutterPosition | None:
        """Gets/Sets Inner Layout Settings instance"""
        try:
            return self._gutter_pos
        except AttributeError:
            self._gutter_pos = cast(_GutterPosition, self._get_style_inst("gutter_pos"))
        return self._gutter_pos

    @_inner_gutter_pos.setter
    def _inner_gutter_pos(self, value: _GutterPosition | None) -> None:
        if value is None:
            self._del_attribs("_gutter_pos")
            self._remove_style("gutter_pos")
            return
        if not isinstance(value, _GutterPosition):
            raise TypeError(f'Expected type of _GutterPosition, got "{type(value).__name__}"')
        self._del_attribs("_gutter_pos")
        self._set_style("gutter_pos", value)

    @property
    def prop_gutter_pos_left(self) -> bool | None:
        """
        Gets/Sets if the gutters position is Left or Top of page. If ``True`` then gutter is left, if ``False`` gutter is top.
        """
        if self._inner_gutter_pos is None:
            return None
        return self._inner_gutter_pos.prop_pos_left

    @prop_gutter_pos_left.setter
    def prop_gutter_pos_left(self, value: bool | None) -> None:
        if value is None:
            self._inner_gutter_pos = None
            return
        self._inner_gutter_pos = _GutterPosition(pos_left=value)

    @property
    def prop_layout(self) -> PageStyleLayout | None:
        """Gets/Sets Page Style Layout value"""
        return self.prop_inner.prop_layout

    @prop_layout.setter
    def prop_layout(self, value: PageStyleLayout | None) -> None:
        self.prop_inner.prop_layout = value

    @property
    def prop_numbers(self) -> NumberingTypeEnum | None:
        """Gets/Sets Page Numbering value"""
        return self.prop_inner.prop_numbers

    @prop_numbers.setter
    def prop_numbers(self, value: NumberingTypeEnum | None) -> None:
        self.prop_inner.prop_numbers = value

    @property
    def prop_ref_style(self) -> str | None:
        """Gets/Sets the name of the paragraph style that is used as reference of the register mode."""
        return self.prop_inner.prop_ref_style

    @prop_ref_style.setter
    def prop_ref_style(self, value: str | StyleParaKind | None) -> None:
        self.prop_inner.prop_ref_style = value

    @property
    def prop_right_gutter(self) -> bool | None:
        """Gets/Sets if the page gutter is placed on the right side of the page."""
        return self.prop_inner.prop_right_gutter

    @prop_right_gutter.setter
    def prop_right_gutter(self, value: bool | None) -> None:
        self.prop_inner.prop_right_gutter = value
