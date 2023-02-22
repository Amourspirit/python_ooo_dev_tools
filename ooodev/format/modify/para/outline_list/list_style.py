from __future__ import annotations
from typing import Any, cast, Tuple, overload, Type, TypeVar
import uno
from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ....writer.style.lst import StyleListKind as StyleListKind
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti


_TInnerListStyle = TypeVar(name="_TInnerListStyle", bound="InnerListStyle")


class InnerListStyle(StyleBase):
    """
    Style Paragraph ListStyle

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, list_style: str | StyleListKind | None = None) -> None:
        """
        Constructor

        Args:
            list_style (str, StyleListKind, optional): List Style.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        init_vals = {}
        if not list_style is None:
            # if list_style is StyleListKind and it is StyleListKind.NONE then str will be empty string
            str_style = str(list_style)
            if str_style:
                init_vals["NumberingStyleName"] = str_style
            else:
                init_vals["NumberingStyleName"] = ""

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    @classmethod
    def from_obj(cls: Type[_TInnerListStyle], obj: object) -> _TInnerListStyle:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            DirectListStyle: ``DirectListStyle`` instance that represents ``obj`` properties.
        """
        inst = cls()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, o: InnerListStyle):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                o._set(key, val)

        set_prop("NumberingStyleName", inst)

        return inst

    # endregion methods

    # region Style Methods
    def fmt_list_style(self: _TInnerListStyle, value: str | StyleListKind | None) -> _TInnerListStyle:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (str, StyleListKind, None): List style value.

        Returns:
            DirectListStyle: List Style instance
        """
        cp = self.copy()
        cp.prop_list_style = value
        return cp

    # endregion Style Methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_list_style(self) -> str | None:
        """Gets/Sets List Style"""
        return self._get("NumberingStyleName")

    @prop_list_style.setter
    def prop_list_style(self, value: str | StyleListKind | None) -> None:
        if value is None:
            self._remove("NumberingStyleName")
            return
        str_val = str(value)
        if not str_val:
            # empty string
            self._remove("NumberingStyleName")
            return
        self._set("NumberingStyleName", value)

    # endregion properties


class ListStyle(ParaStyleBaseMulti):
    """
    Paragraph List Style

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        list_style: str | StyleListKind | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            list_style (str, StyleListKind, optional): List Style.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = InnerListStyle(list_style=list_style)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> ListStyle:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            ListStyle: ``ListStyle`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerListStyle.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerListStyle:
        """Gets/Sets Inner List Style instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerListStyle, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerListStyle) -> None:
        if not isinstance(value, InnerListStyle):
            raise TypeError(f'Expected type of InnerListStyle, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
