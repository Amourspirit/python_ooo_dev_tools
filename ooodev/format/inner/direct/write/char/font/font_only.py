"""
Module for managing character font.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
import contextlib
from typing import Any, Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING

from com.sun.star.beans import XPropertySet

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.font_only_props import FontOnlyProps
from ooodev.format.inner.direct.structs.locale_struct import LocaleStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo
from ooodev.meta.class_property_readonly import ClassPropertyReadonly
from ooodev.mock import mock_g
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_pt import UnitPT
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.format.proto.font.font_lang_t import FontLangT
else:
    FontLangT = Any
# endregion Import

if mock_g.DOCS_BUILDING:
    from ooodev.meta.static_prop import static_prop

_TFontOnly = TypeVar("_TFontOnly", bound="FontOnly")
_TFontLang = TypeVar("_TFontLang", bound="FontLang")


class FontLang(LocaleStruct):
    """
    Class for Character Language

    .. seealso::

        - :ref:`help_writer_format_direct_char_font_only`
        - :ref:`help_writer_format_direct_char_font_effects`
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "CharLocale"
        return self._property_name

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("Lang.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    if mock_g.DOCS_BUILDING:
        default: FontLang = None  # type: ignore
        """
        Gets ``Lang`` default.

        The language is determined by LibreOffice settings / language settings / languages / Local Setting

        Static Property.
        """

    else:

        @ClassPropertyReadonly
        @classmethod
        def default(cls: Type[_TFontLang]) -> FontLang:
            """
            Gets ``Lang`` default.

            The language is determined by LibreOffice settings / language settings / languages / Local Setting

            Static Property.
            """
            # sourcery skip: raise-from-previous-error
            try:
                return cls._DEFAULT_INSTANCE
            except AttributeError:
                inst = cls()
                s = (
                    cast(
                        str,
                        mInfo.Info.get_config(node_str="ooSetupSystemLocale", node_path="/org.openoffice.Setup/L10N"),
                    )
                    or cast(str, mInfo.Info.get_config(node_str="ooLocale", node_path="/org.openoffice.Setup/L10N"))
                    or "en-US"
                )

                parts = s.split("-")
                if len(parts) == 2:
                    inst._set("Language", parts[0])
                    inst._set("Country", parts[1])
                else:
                    inst._set("Language", "qlt")
                    country = next((part for part in parts if part.upper() == part), None)
                    if country is None:
                        raise RuntimeError(f'Unable to get country code from System locale "{s}"')
                    inst._set("Country", country)
                    inst._set("Variant", s)
                    inst._is_default_inst = True
                cls._DEFAULT_INSTANCE = inst
            return cls._DEFAULT_INSTANCE


class FontOnly(StyleMulti):
    """
    Character Font

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    .. seealso::

        - :ref:`help_writer_format_direct_char_font_only`
        - :ref:`help_writer_format_direct_char_font_effects`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        name: str | None = None,
        size: float | UnitT | None = None,
        font_style: str | None = None,
        lang: FontLangT | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            name (str, optional): This property specifies the name of the font style. It may contain more than one
                name separated by comma.
            size (float, UnitT, optional): This value contains the size of the characters in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            font_style (str, optional): Font style name such as ``Bold``.
            lang (Lang, optional): Font Language

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_char_font_only`
            - :ref:`help_writer_format_direct_char_font_effects`

        Hint:
            - ``FontLang`` can be imported from ``ooodev.format.inner.direct.write.char.font.font_only``
        """
        init_vals = {}
        if name is not None:
            init_vals[self._props.name] = name
        if font_style is not None:
            init_vals[self._props.style_name] = font_style

        super().__init__(**init_vals)
        if lang is not None:
            self._set_style("lang", lang, *lang.get_attrs())
        self._set_fd_style(name, font_style)
        if size is not None:
            self.prop_size = size

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.drawing.ControlShape",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Setting properties on a default instance is not allowed")
        return super()._on_modifying(source, event)

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == self._props.style_name and not event_args.value:
            event_args.default = True
        return super().on_property_setting(source, event_args)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    # endregion Overrides

    # region Internal Methods

    def _set_fd_style(self, name: str | None, style: str | None) -> None:
        """Set Font Descriptor Style"""
        # sourcery skip: extract-method
        if not name or not style:
            return None
        fd = mInfo.Info.get_font_descriptor(name, style)
        if fd is None:
            self._remove("CharFontFamily")
            self._remove("CharPosture")
            self._remove(self._props.style_name)
            self._remove("CharStrikeout")
            self._remove("CharRotation")
            self._remove("CharUnderline")
            self._remove("CharWeight")
            return None
        self._set("CharFontFamily", fd.Family)
        self._set("CharPosture", fd.Slant)
        self._set(self._props.style_name, fd.StyleName)
        self._set("CharStrikeout", fd.Strikeout)
        self._set("CharRotation", fd.Orientation)
        self._set("CharUnderline", fd.Underline)
        self._set("CharWeight", fd.Weight)

    # endregion Internal Methods

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TFontOnly], obj: Any) -> _TFontOnly: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TFontOnly], obj: Any, **kwargs) -> _TFontOnly: ...

    @classmethod
    def from_obj(cls: Type[_TFontOnly], obj: Any, **kwargs) -> _TFontOnly:
        """
        Gets Font Only instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            FontOnly: Font Only that represents ``obj`` Font.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, align: FontOnly):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                align._set(key, val)

        set_prop(inst._props.name, inst)
        set_prop(inst._props.size, inst)
        set_prop(inst._props.style_name, inst)
        with contextlib.suppress(mEx.PropertyNotFoundError):
            lang = FontLang.from_obj(obj)
            inst._set_style("lang", lang, *lang.get_attrs())
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region Format Methods

    def fmt_size(self: _TFontOnly, value: float | UnitT | None = None) -> _TFontOnly:
        """
        Get copy of instance with text size set.

        Args:
            value (float, UnitT, optional): The size of the characters in ``pt`` (point) units or :ref:`proto_unit_obj`.

        Returns:
            FontOnly: Font with style added
        """
        ft = self.copy()
        ft.prop_size = value
        return ft

    def fmt_name(self: _TFontOnly, value: str | None = None) -> _TFontOnly:
        """
        Get copy of instance with name set.

        Args:
            value (str, optional): The name of the font.

        Returns:
            FontOnly: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_name = value
        return ft

    def fmt_style_name(self: _TFontOnly, value: str | None = None) -> _TFontOnly:
        """
        Get copy of instance with style name set.

        Args:
            value (str, optional): The style name of the font.

        Returns:
            FontOnly: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_style_name = value
        return ft

    # endregion Format Methods

    # region Style Properties

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_size(self) -> UnitPT | None:
        """This value contains the size of the characters in point."""
        return UnitPT(self._get(self._props.size))

    @prop_size.setter
    def prop_size(self, value: float | UnitT | None) -> None:
        if value is None:
            self._remove(self._props.size)
            return
        try:
            self._set(self._props.size, value.get_value_pt())  # type: ignore
        except AttributeError:
            self._set(self._props.size, value)

    @property
    def prop_name(self) -> str | None:
        """This property specifies the name of the font style."""
        return self._get(self._props.name)

    @prop_name.setter
    def prop_name(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.name)
            return
        self._set(self._props.name, value)

    @property
    def prop_style_name(self) -> str | None:
        """This property specifies the style name of the font style."""
        return self._get(self._props.style_name)

    @prop_style_name.setter
    def prop_style_name(self, value: str | None) -> None:
        # style name will be added or removed in _set_fd_style()
        self._set_fd_style(self.prop_name, value)

    @property
    def _props(self) -> FontOnlyProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FontOnlyProps(
                name="CharFontName", size="CharHeight", style_name="CharFontStyleName"
            )
        return self._props_internal_attributes

    @property
    def prop_inner(self) -> FontLang:
        """Gets Lang instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(FontLang, self._get_style_inst("lang"))
        return self._direct_inner

    # endregion Prop Properties

    # region Static Properties

    if mock_g.DOCS_BUILDING:
        default: FontOnly = None  # type: ignore
        """
        Gets ``Font`` default.

        The language is determined by LibreOffice settings / Basic Fonts / Default

        Static Property.
        """

        default_caption: FontOnly = None  # type: ignore
        """
        Gets ``Font`` caption default.

        The language is determined by LibreOffice settings / Basic Fonts / Caption

        Static Property.
        """

        default_heading: FontOnly = None  # type: ignore
        """
        Gets ``Font`` heading default.

        The language is determined by LibreOffice settings / Basic Fonts / Heading

        Static Property.
        """

        default_list: FontOnly = None  # type: ignore
        """
        Gets ``Font`` list default.

        The language is determined by LibreOffice settings / Basic Fonts / List

        Static Property.
        """

        default_index: FontOnly = None  # type: ignore
        """
        Gets ``Font`` index default.

        The language is determined by LibreOffice settings / Basic Fonts / Index

        Static Property.
        """

    else:

        @ClassPropertyReadonly
        @classmethod
        def default(cls: Type[_TFontOnly]) -> _TFontOnly:  # type: ignore[misc]
            """
            Gets ``Font`` default.

            The language is determined by LibreOffice settings / Basic Fonts / Default

            Static Property."""
            try:
                return cls._DEFAULT_STANDARD
            except AttributeError:
                inst = cls()

                props = mLo.Lo.qi(
                    XPropertySet,
                    mInfo.Info.get_config(node_str="DefaultFont", node_path="/org.openoffice.Office.Writer/"),
                )
                font_name = mProps.Props.get(props, "Standard", "Liberation Serif")
                font_size = mProps.Props.get(props, "StandardHeight", None)
                if font_size is None:
                    font_size = 12
                else:
                    font_size = round(UnitConvert.convert_mm100_pt(font_size), 2)
                inst._set(inst._props.name, font_name)
                inst._set(inst._props.size, font_size)
                inst._set(inst._props.style_name, "")
                lang = FontLang.default.copy()
                inst._set_style("lang", lang, *lang.get_attrs())
                inst._is_default_inst = True
                cls._DEFAULT_STANDARD = inst
            return cls._DEFAULT_STANDARD

        @ClassPropertyReadonly
        @classmethod
        def default_caption(cls: Type[_TFontOnly]) -> _TFontOnly:  # type: ignore[misc]
            """
            Gets ``Font`` caption default.

            The language is determined by LibreOffice settings / Basic Fonts / Caption

            Static Property."""
            try:
                return cls._DEFAULT_CAPTION
            except AttributeError:
                inst = cls()

                props = mLo.Lo.qi(
                    XPropertySet,
                    mInfo.Info.get_config(node_str="DefaultFont", node_path="/org.openoffice.Office.Writer/"),
                )
                font_name = mProps.Props.get(props, "Caption", "Liberation Serif")
                font_size = mProps.Props.get(props, "CaptionHeight", None)
                if font_size is None:
                    font_size = 12
                else:
                    font_size = round(UnitConvert.convert_mm100_pt(font_size), 2)
                inst._set(inst._props.name, font_name)
                inst._set(inst._props.size, font_size)
                inst._set(inst._props.style_name, "")
                lang = FontLang.default.copy()
                inst._set_style("lang", lang, *lang.get_attrs())
                inst._is_default_inst = True
                cls._DEFAULT_CAPTION = inst
            return cls._DEFAULT_CAPTION

        @ClassPropertyReadonly
        @classmethod
        def default_heading(cls: Type[_TFontOnly]) -> _TFontOnly:  # type: ignore[misc]
            """
            Gets ``Font`` heading default.

            The language is determined by LibreOffice settings / Basic Fonts / Heading

            Static Property."""
            try:
                return cls._DEFAULT_HEADING
            except AttributeError:
                inst = cls()

                props = mLo.Lo.qi(
                    XPropertySet,
                    mInfo.Info.get_config(node_str="DefaultFont", node_path="/org.openoffice.Office.Writer/"),
                )
                font_name = mProps.Props.get(props, "Heading", "Liberation Sans")
                font_size = mProps.Props.get(props, "HeadingHeight", None)
                if font_size is None:
                    font_size = 14
                else:
                    font_size = round(UnitConvert.convert_mm100_pt(font_size), 2)
                inst._set(inst._props.name, font_name)
                inst._set(inst._props.size, font_size)
                inst._set(inst._props.style_name, "")
                lang = FontLang.default.copy()
                inst._set_style("lang", lang, *lang.get_attrs())
                inst._is_default_inst = True
                cls._DEFAULT_HEADING = inst
            return cls._DEFAULT_HEADING

        @ClassPropertyReadonly
        @classmethod
        def default_list(cls: Type[_TFontOnly]) -> _TFontOnly:  # type: ignore[misc]
            """
            Gets ``Font`` list default.

            The language is determined by LibreOffice settings / Basic Fonts / List

            Static Property."""
            try:
                return cls._DEFAULT_LIST
            except AttributeError:
                inst = cls()

                props = mLo.Lo.qi(
                    XPropertySet,
                    mInfo.Info.get_config(node_str="DefaultFont", node_path="/org.openoffice.Office.Writer/"),
                )
                font_name = mProps.Props.get(props, "List", "Liberation Serif")
                font_size = mProps.Props.get(props, "ListHeight", None)
                if font_size is None:
                    font_size = 12
                else:
                    font_size = round(UnitConvert.convert_mm100_pt(font_size), 2)
                inst._set(inst._props.name, font_name)
                inst._set(inst._props.size, font_size)
                inst._set(inst._props.style_name, "")
                lang = FontLang.default.copy()
                inst._set_style("lang", lang, *lang.get_attrs())
                inst._is_default_inst = True
                cls._DEFAULT_LIST = inst
            return cls._DEFAULT_LIST

        @ClassPropertyReadonly
        @classmethod
        def default_index(cls: Type[_TFontOnly]) -> _TFontOnly:  # type: ignore[misc]
            """
            Gets ``Font`` index default.

            The language is determined by LibreOffice settings / Basic Fonts / Index

            Static Property."""
            try:
                return cls._DEFAULT_INDEX
            except AttributeError:
                inst = cls()

                props = mLo.Lo.qi(
                    XPropertySet,
                    mInfo.Info.get_config(node_str="DefaultFont", node_path="/org.openoffice.Office.Writer/"),
                )
                font_name = mProps.Props.get(props, "Index", "Liberation Serif")
                font_size = mProps.Props.get(props, "IndexHeight", None)
                if font_size is None:
                    font_size = 12
                else:
                    font_size = round(UnitConvert.convert_mm100_pt(font_size), 2)
                inst._set(inst._props.name, font_name)
                inst._set(inst._props.size, font_size)
                inst._set(inst._props.style_name, "")
                lang = FontLang.default.copy()
                inst._set_style("lang", lang, *lang.get_attrs())
                inst._is_default_inst = True
                cls._DEFAULT_INDEX = inst
            return cls._DEFAULT_INDEX

    # endregion Static Properties
