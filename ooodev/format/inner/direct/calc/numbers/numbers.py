# region Import
from __future__ import annotations
from typing import overload
from typing import Any, Tuple, Type, TypeVar
import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.lang import XComponent
from com.sun.star.util import XNumberFormatsSupplier
from com.sun.star.util import XNumberFormatTypes

from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.util.number_format import NumberFormatEnum

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps

# com.sun.star.i18n.NumberFormatIndex
# endregion Import

_TNumbers = TypeVar(name="_TNumbers", bound="Numbers")

# https://wiki.documentfoundation.org/Documentation/DevGuide/Office_Development#Number_Formats


class Numbers(StyleBase):
    """
    Calc Numbers format.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
        component: XComponent | None = None,
    ) -> None:
        """
        Constructor

        Args:
            num_format (NumberFormatEnum, int, optional): Type of a number format.
                Use this to select a defatult format. Defaults to 0 (General Format).
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum | int, optional): Index of a number format.
                The enumeration values represent the built-in number formats. Defaults to ``-1``.
            lang_locale (Locale, optional): Locale of the number format. Defaults to ``None`` which used current Locale.
            component (XComponent, optional): Document such as Spreadsheet or Chart. If Omittet, the current document is used. Defaults to ``None``.
        """
        super().__init__()
        if lang_locale is None:
            # note the empty locale for default locale
            lang_locale = Locale()
        if component is None:
            component = mLo.Lo.this_component
        self._num_cat = int(num_format)
        if self._num_cat < 0:
            self._num_cat = 0
        self._num_format_index = int(num_format_index)
        self._lang_locale = lang_locale
        self._component = component

    # region internal methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "NumberFormat"
        return self._property_name

    def _get_format_supplier(self) -> XNumberFormatsSupplier:
        return mLo.Lo.qi(XNumberFormatsSupplier, self._component, True)

    def _query_key(self, nf_str: str) -> int:
        xfs = mLo.Lo.qi(XNumberFormatsSupplier, self._component, True)
        n_formats = xfs.getNumberFormats()
        key = int(n_formats.queryKey(nf_str, self._lang_locale, False))
        return key  # -1 means not found

    def _get_by_key_props(self) -> XPropertySet:
        xfs = mLo.Lo.qi(XNumberFormatsSupplier, self._component, True)
        n_formats = xfs.getNumberFormats()
        nf_props = n_formats.getByKey(self.prop_format_key)
        return nf_props

    def _get_format_index(self) -> int:
        xfs = mLo.Lo.qi(XNumberFormatsSupplier, self._component, True)
        n_formats = xfs.getNumberFormats()
        nft = mLo.Lo.qi(XNumberFormatTypes, n_formats, True)
        if self._num_format_index == -1:
            key = nft.getStandardFormat(self._num_cat, self._lang_locale)
            return key
        else:
            key = nft.getFormatIndex(self._num_format_index, self._lang_locale)
            return key

    # endregion internal methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def _is_valid_obj(self, obj: object) -> bool:
        return hasattr(obj, self._get_property_name())

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, **kwargs: Any) -> None:
        ...

    def apply(self, obj: object, **kwargs: Any) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Keyword Args:
            format_key (Any, optional): NumberFormat key, overrides ``prop_format_key`` property value.

        Returns:
            None:
        """
        # key = self._get_by_key_props(self._num_cat)
        key = kwargs.pop("format_key", -1)
        # key may be none.
        # This is fine for some child classes.
        if key == -1:
            key = self.prop_format_key

            if key == -1:
                mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property. NumberFormat not found")
                return

        props = kwargs.pop("override_dv", {})
        props.update({self._get_property_name(): key})
        super().apply(obj, override_dv=props, **kwargs)

    # endregion apply()
    # region Copy()
    @overload
    def copy(self: _TNumbers) -> _TNumbers:
        ...

    @overload
    def copy(self: _TNumbers, **kwargs) -> _TNumbers:
        ...

    def copy(self: _TNumbers, **kwargs) -> _TNumbers:
        """
        Creates a copy of the instance.

        Returns:
            Numbers: Copy of the instance.
        """
        inst = self.__class__(**kwargs)
        inst._format_key_prop = self.prop_format_key
        return inst

    # endregion Copy()

    # endregion Overrides

    # region Static Methods

    # region from string
    @classmethod
    def from_str(
        cls: Type[_TNumbers], nf_str: str, lang_locale: Locale | None = None, auto_add: bool = False, **kwargs
    ) -> _TNumbers:
        """
        Gets instance from format string

        Args:
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If True, format string will be added to document if not found. Defaults to ``False``.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        nu = cls(num_format=0, lang_locale=lang_locale, **kwargs)
        key = nu._query_key(nf_str)
        if key == -1:
            if auto_add:
                try:
                    xfs = nu._get_format_supplier()
                    n_formats = xfs.getNumberFormats()
                    key = n_formats.addNew(nf_str, nu._lang_locale)
                except Exception as e:
                    raise ValueError(f"Unable to add format string: {nf_str}") from e
            else:
                raise ValueError(f"Format string not found: {nf_str}")

        inst = cls(lang_locale=lang_locale, **kwargs)
        inst._format_key_prop = key
        return inst

    # endregion from string

    # region from index
    @classmethod
    def from_index(cls: Type[_TNumbers], index: int, lang_locale: Locale | None = None, **kwargs) -> _TNumbers:
        """
        Gets instance from number format index. This is the index that is assinged to the ``NumberFormat`` property of an object such as a cell.

        Args:
            index (int): Format (``NumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        inst = cls(lang_locale=lang_locale, **kwargs)
        inst._format_key_prop = index
        return inst

    # endregion from index

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TNumbers], obj: object) -> _TNumbers:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TNumbers], obj: object, **kwargs) -> _TNumbers:
        ...

    @classmethod
    def from_obj(cls: Type[_TNumbers], obj: object, **kwargs) -> _TNumbers:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        nu = cls(**kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        # Get the number format index key of the cell's properties
        nf = int(mProps.Props.get(obj, nu._get_property_name(), -1))
        if nf == -1:
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        locale = mProps.Props.get(obj, "CharLocale", None)
        inst = cls(lang_locale=locale, **kwargs)
        inst._format_key_prop = nf
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Instance Properties

    @property
    def general(self: _TNumbers) -> _TNumbers:
        """Gets general format"""
        return self.__class__(num_format=0, lang_locale=self._lang_locale, component=self._component)

    @property
    def date(self: _TNumbers) -> _TNumbers:
        """Gets date format"""
        return self.__class__(
            num_format=NumberFormatEnum.DATE, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def time(self: _TNumbers) -> _TNumbers:
        """Gets time format"""
        return self.__class__(
            num_format=NumberFormatEnum.TIME, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def currency(self: _TNumbers) -> _TNumbers:
        """Gets currency format"""
        return self.__class__(
            num_format=NumberFormatEnum.CURRENCY, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def number(self: _TNumbers) -> _TNumbers:
        """Gets number format"""
        return self.__class__(
            num_format=NumberFormatEnum.NUMBER, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def scientific(self: _TNumbers) -> _TNumbers:
        """Gets scientific format"""
        return self.__class__(
            num_format=NumberFormatEnum.SCIENTIFIC, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def fraction(self: _TNumbers) -> _TNumbers:
        """Gets fraction format"""
        return self.__class__(
            num_format=NumberFormatEnum.FRACTION, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def percent(self: _TNumbers) -> _TNumbers:
        """Gets percent format"""
        return self.__class__(
            num_format=NumberFormatEnum.PERCENT, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def text(self: _TNumbers) -> _TNumbers:
        """Gets text format"""
        return self.__class__(
            num_format=NumberFormatEnum.TEXT, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def datetime(self: _TNumbers) -> _TNumbers:
        """Gets datetime format"""
        return self.__class__(
            num_format=NumberFormatEnum.DATETIME, lang_locale=self._lang_locale, component=self._component
        )

    @property
    def boolean(self: _TNumbers) -> _TNumbers:
        """Gets boolean format"""
        return self.__class__(
            num_format=NumberFormatEnum.LOGICAL, lang_locale=self._lang_locale, component=self._component
        )

    # endregion Instance Properties

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CELL
        return self._format_kind_prop

    @property
    def prop_format_key(self) -> int:
        """Gets the format key"""
        try:
            return self._format_key_prop
        except AttributeError:
            self._format_key_prop = self._get_format_index()
        return self._format_key_prop

    @property
    def prop_format_str(self) -> str:
        """Gets the format string, e.g. ``#,##0.00``"""
        try:
            return self._format_str_prop
        except AttributeError:
            props = self._get_by_key_props()
            self._format_str_prop = str(mProps.Props.get(props, "FormatString"))
        return self._format_str_prop

    # endregion Properties
