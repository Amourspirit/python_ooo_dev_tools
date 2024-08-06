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
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

# com.sun.star.i18n.NumberFormatIndex
# endregion Import

_TNumbers = TypeVar("_TNumbers", bound="Numbers")

# https://wiki.documentfoundation.org/Documentation/DevGuide/Office_Development#Number_Formats


class Numbers(StyleBase):
    """
    Calc Numbers format.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_numbers`

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
                Use this to select a default format. Defaults to 0 (General Format).
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum, int, optional): Index of a number format.
                The enumeration values represent the built-in number formats. Defaults to ``-1``.
            lang_locale (Locale, optional): Locale of the number format. Defaults to ``None`` which used current Locale.
            component (XComponent, optional): Document such as Spreadsheet or Chart. If Omitted, the current document is used. Defaults to ``None``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_direct_cell_numbers`
            - `API NumberFormat <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html>`__
            - `API NumberFormatIndex <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html>`__
        """
        super().__init__()
        if lang_locale is None:
            # note the empty locale for default locale
            lang_locale = Locale()
        if component is None:
            component = mLo.Lo.this_component
        if component is None:
            # this is only likely if both headless and options dynamic is set for Lo class
            component = mLo.Lo.lo_component
        self._num_cat = int(num_format)
        self._num_cat = max(self._num_cat, 0)
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
        return int(n_formats.queryKey(nf_str, self._lang_locale, False))

    def _get_by_key_props(self) -> XPropertySet:
        xfs = mLo.Lo.qi(XNumberFormatsSupplier, self._component, True)
        n_formats = xfs.getNumberFormats()
        return n_formats.getByKey(self.prop_format_key)

    def _get_format_index(self) -> int:
        xfs = mLo.Lo.qi(XNumberFormatsSupplier, self._component, True)
        n_formats = xfs.getNumberFormats()
        nft = mLo.Lo.qi(XNumberFormatTypes, n_formats, True)
        if self._num_format_index == -1:
            return nft.getStandardFormat(self._num_cat, self._lang_locale)
        else:
            return nft.getFormatIndex(self._num_format_index, self._lang_locale)

    # endregion internal methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return hasattr(obj, self._get_property_name())

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs: Any) -> None: ...

    def apply(self, obj: Any, **kwargs: Any) -> None:
        # sourcery skip: hoist-if-from-if
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
    def copy(self: _TNumbers) -> _TNumbers: ...

    @overload
    def copy(self: _TNumbers, **kwargs) -> _TNumbers: ...

    def copy(self: _TNumbers, **kwargs) -> _TNumbers:
        """
        Creates a copy of the instance.

        Returns:
            Numbers: Copy of the instance.
        """
        # pylint: disable=protected-access
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
        # sourcery skip: hoist-similar-statement-from-if, remove-unnecessary-else, swap-if-else-branches, swap-nested-ifs
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
        # pylint: disable=protected-access
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
        Gets instance from number format index. This is the index that is assigned to the ``NumberFormat`` property of an object such as a cell.

        Args:
            index (int): Format (``NumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        # pylint: disable=protected-access
        inst = cls(lang_locale=lang_locale, **kwargs)
        inst._format_key_prop = index
        return inst

    # endregion from index

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TNumbers], obj: Any) -> _TNumbers: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TNumbers], obj: Any, **kwargs) -> _TNumbers: ...

    @classmethod
    def from_obj(cls: Type[_TNumbers], obj: Any, **kwargs) -> _TNumbers:
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
        # pylint: disable=protected-access
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
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Instance Properties

    @property
    def general(self: _TNumbers) -> _TNumbers:
        """Gets general format"""
        result = self.__class__(num_format=0, lang_locale=self._lang_locale, component=self._component)
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def date(self: _TNumbers) -> _TNumbers:
        """Gets date format"""
        result = self.__class__(
            num_format=NumberFormatEnum.DATE, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def time(self: _TNumbers) -> _TNumbers:
        """Gets time format"""
        result = self.__class__(
            num_format=NumberFormatEnum.TIME, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def currency(self: _TNumbers) -> _TNumbers:
        """Gets currency format"""
        result = self.__class__(
            num_format=NumberFormatEnum.CURRENCY, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def number(self: _TNumbers) -> _TNumbers:
        """Gets number format"""
        result = self.__class__(
            num_format=NumberFormatEnum.NUMBER, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def scientific(self: _TNumbers) -> _TNumbers:
        """Gets scientific format"""
        result = self.__class__(
            num_format=NumberFormatEnum.SCIENTIFIC, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def fraction(self: _TNumbers) -> _TNumbers:
        """Gets fraction format"""
        result = self.__class__(
            num_format=NumberFormatEnum.FRACTION, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def percent(self: _TNumbers) -> _TNumbers:
        """Gets percent format"""
        result = self.__class__(
            num_format=NumberFormatEnum.PERCENT, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def text(self: _TNumbers) -> _TNumbers:
        """Gets text format"""
        result = self.__class__(
            num_format=NumberFormatEnum.TEXT, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def datetime(self: _TNumbers) -> _TNumbers:
        """Gets datetime format"""
        result = self.__class__(
            num_format=NumberFormatEnum.DATETIME, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

    @property
    def boolean(self: _TNumbers) -> _TNumbers:
        """Gets boolean format"""
        result = self.__class__(
            num_format=NumberFormatEnum.LOGICAL, lang_locale=self._lang_locale, component=self._component
        )
        if self.has_update_obj():
            result.set_update_obj(self.get_update_obj())
        return result

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
