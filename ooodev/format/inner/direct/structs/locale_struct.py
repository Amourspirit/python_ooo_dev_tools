"""
Module for ``LocaleStruct`` struct.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Dict, Tuple, Type, cast, overload, TypeVar

import uno
from ooo.dyn.lang.locale import Locale

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import props as mProps


# endregion Import
# pylint: disable=unexpected-keyword-arg

_TLocaleStruct = TypeVar("_TLocaleStruct", bound="LocaleStruct")


class LocaleStruct(StructBase):
    """
    Paragraph Drop Cap

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, *, country: str = "US", language: str = "en", variant: str = "") -> None:
        """
        Constructor

        Args:
            country (str, optional): Specifies an ``ISO 3166`` Country Code.
                These codes are the upper-case two-letter codes as defined by ``ISO 3166-1``.
                You can find a full list of these codes at a number of sites, such as:
                `Wikipedia ISO 3166-1 alpha-2 <https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2>`__.
                If this field contains an empty string, the meaning depends on the context.
            language (str, optional): Specifies an ``ISO 639`` Language Code.
                These codes are preferably the lower-case two-letter codes as defined by ``ISO 639-1``, or
                three-letter codes as defined by ``ISO 639-3``.
                You can find a full list of these codes at a number of sites, such as:
                `ISO 639 Code Tables <https://iso639-3.sil.org/code_tables/639/data>`__.
                If this field contains an empty string, the meaning depends on the context.
            variant (str, optional): Specifies a ``BCP 47`` Language Tag.
                You can find BCP 47 language tag resources at:
                `Wikipedia IETF language tag <https://en.wikipedia.org/wiki/IETF_language_tag>`__ or
                `Language tags in HTML and XML <https://www.w3.org/International/articles/language-tags/>`__.

        Returns:
            None:
        """

        init_vals = {"Country": country.upper()[:2], "Language": language, "Variant": variant}

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "CharLocale"
        return self._property_name

    def _is_valid_obj(self, obj: Any) -> bool:
        return mProps.Props.has(obj, self._get_property_name())

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, keys: Dict[str, str]) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:  # type: ignore
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.
            keys (Dict[str, str], optional): Property key, value items that map properties.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLIED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        # sourcery skip: dict-assign-update-to-union
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_srv("apply")
            return

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        if cargs.cancel:
            return
        self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["prop"]
        dcf = self.get_uno_struct()
        mProps.Props.set(obj, **{key: dcf})
        eargs = EventArgs.from_args(cargs)
        self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

        # mProps.Props.set(obj, **{key: tuple(tss_lst)})

    # endregion apply()

    def get_uno_struct(self) -> Locale:
        """
        Gets UNO ``Locale`` from instance.

        Returns:
            Locale: ``Locale`` instance
        """
        return Locale(Language=self._get("Language"), Country=self._get("Country"), Variant=self._get("Variant"))

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TLocaleStruct], obj: Any) -> _TLocaleStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TLocaleStruct], obj: Any, **kwargs) -> _TLocaleStruct: ...

    @classmethod
    def from_obj(cls: Type[_TLocaleStruct], obj: Any, **kwargs) -> _TLocaleStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            DropCap: ``DropCap`` instance that represents ``obj`` Drop cap format properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            dcf = cast(Locale, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_uno_struct(dcf, **kwargs)

    # endregion from_obj()

    # region from_locale()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TLocaleStruct], locale: Locale) -> _TLocaleStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TLocaleStruct], locale: Locale, **kwargs) -> _TLocaleStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TLocaleStruct], locale: Locale, **kwargs) -> _TLocaleStruct:
        """
        Converts a ``Locale`` Stop instance to a ``LocaleStruct``

        Args:
            dcf (Locale): UNO locale

        Returns:
            LocaleStruct: ``LocaleStruct`` set with locale properties
        """
        return cls(country=locale.Country, language=locale.Language, variant=locale.Variant, **kwargs)

    # endregion from_locale()

    # endregion methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, LocaleStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.lang.Locale":
            obj2 = cast(Locale, oth)
        if not obj2 is None:
            obj1 = self.get_uno_struct()
            return obj1.Country == obj2.Country and obj1.Language == obj2.Language and obj1.Variant == obj2.Variant
        return NotImplemented

    # endregion dunder methods

    # region format methods
    def fmt_country(self: _TLocaleStruct, value: str) -> _TLocaleStruct:
        """
        Gets a copy of instance with country set.

        Args:
            value (int): Country value.

        Returns:
            LocaleStruct: ``LocaleStruct`` instance
        """
        cp = self.copy()
        cp.prop_country = value
        return cp

    def fmt_language(self: _TLocaleStruct, value: str) -> _TLocaleStruct:
        """
        Gets a copy of instance with language set.

        Args:
            value (int): Language value.

        Returns:
            LocaleStruct: ``LocaleStruct`` instance
        """
        cp = self.copy()
        cp.prop_language = value
        return cp

    def fmt_variant(self: _TLocaleStruct, value: str) -> _TLocaleStruct:
        """
        Gets a copy of instance with variant set.

        Args:
            value (int): Variant value.

        Returns:
            LocaleStruct: ``LocaleStruct`` instance
        """
        cp = self.copy()
        cp.prop_variant = value
        return cp

    # endregion format methods

    # region Style Properties
    @property
    def english_us(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English US locale"""
        return self.__class__(
            country="US",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def locale_none(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets no locale"""
        return self.__class__(
            country="",
            language="zxx",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_australia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Australia locale"""
        return self.__class__(
            country="AU",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_belize(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Belize locale"""
        return self.__class__(
            country="BZ",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_botswana(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Botswana locale"""
        return self.__class__(
            country="BW",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_canada(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Canada locale"""
        return self.__class__(
            country="CA",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_caribbean(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Caribbean locale"""
        return self.__class__(
            country="BS",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_gambia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Gambia locale"""
        return self.__class__(
            country="GM",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_ghana(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Ghana locale"""
        return self.__class__(
            country="GH",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_hong_kong(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Hong Kong locale"""
        return self.__class__(
            country="HK",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_india(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English India locale"""
        return self.__class__(
            country="IN",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_ireland(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Ireland locale"""
        return self.__class__(
            country="IE",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_israel(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Israel locale"""
        return self.__class__(
            country="IL",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_jamaica(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Jamaica locale"""
        return self.__class__(
            country="JM",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_kenya(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Kenya locale"""
        return self.__class__(
            country="KE",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_malawi(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Malawi locale"""
        return self.__class__(
            country="MW",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_malaysia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Malaysia locale"""
        return self.__class__(
            country="MY",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_mauritius(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Mauritius locale"""
        return self.__class__(
            country="MU",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_namibia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Namibia locale"""
        return self.__class__(
            country="NA",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_new_zealand(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English New Zealand locale"""
        return self.__class__(
            country="NZ",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_nigeria(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Nigeria locale"""
        return self.__class__(
            country="NG",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_philippines(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Philippines locale"""
        return self.__class__(
            country="PH",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_south_africa(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English South Africa locale"""
        return self.__class__(
            country="ZA",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_sri_lanka(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English ``Sri Lanka`` locale"""
        return self.__class__(
            country="LK",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_trinidad(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Trinidad locale"""
        return self.__class__(
            country="TT",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_uk(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English UK locale"""
        return self.__class__(
            country="GB",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_usa(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English USA locale"""
        return self.__class__(
            country="US",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_zambia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Zambia locale"""
        return self.__class__(
            country="ZM",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_zimbabwe(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English Zimbabwe locale"""
        return self.__class__(
            country="ZW",
            language="en",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def english_uk_ode(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets English ODE Spelling UK locale"""
        return self.__class__(
            country="GB",
            language="qlt",
            variant="en-GB-oxendict",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def german_austria(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets German Austria locale"""
        return self.__class__(
            country="AT",
            language="de",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def german_belgium(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets German Belgium locale"""
        return self.__class__(
            country="BE",
            language="de",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def german_germany(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets German Germany locale"""
        return self.__class__(
            country="DE",
            language="de",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def german_liechtenstein(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets German Liechtenstein locale"""
        return self.__class__(
            country="LI",
            language="de",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def german_luxembourg(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets German Luxembourg locale"""
        return self.__class__(
            country="LU",
            language="de",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def german_switzerland(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets German Switzerland locale"""
        return self.__class__(
            country="CH",
            language="de",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_belgium(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Belgium locale"""
        return self.__class__(
            country="BE",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_benin(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Benin locale"""
        return self.__class__(
            country="BJ",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_burkina_faso(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French ``Burkina Faso`` locale"""
        return self.__class__(
            country="BF",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_canada(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Canada locale"""
        return self.__class__(
            country="CA",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_cote_d_ivoire(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French ``Côte d'Ivoire`` locale"""
        return self.__class__(
            country="CI",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_france(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French France locale"""
        return self.__class__(
            country="FR",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_luxembourg(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Luxembourg locale"""
        return self.__class__(
            country="LU",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_mali(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Mali locale"""
        return self.__class__(
            country="ML",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_mauritius(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Mauritius locale"""
        return self.__class__(
            country="MU",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_monaco(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Monaco locale"""
        return self.__class__(
            country="MC",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_niger(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Niger locale"""
        return self.__class__(
            country="NE",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_senegal(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Senegal locale"""
        return self.__class__(
            country="SN",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_switzerland(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Switzerland locale"""
        return self.__class__(
            country="CH",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def french_togo(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets French Togo locale"""
        return self.__class__(
            country="TG",
            language="fr",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_argentina(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Argentina locale"""
        return self.__class__(
            country="AR",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_bolivia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Bolivia locale"""
        return self.__class__(
            country="BO",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_chile(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Chile locale"""
        return self.__class__(
            country="CL",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_colombia(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Colombia locale"""
        return self.__class__(
            country="CO",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_costa_rica(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish ``Costa Rica`` locale"""
        return self.__class__(
            country="CR",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_cuba(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Cuba locale"""
        return self.__class__(
            country="CU",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_dom_rep(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Dominion Republic locale"""
        return self.__class__(
            country="DO",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_ecuador(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Ecuador locale"""
        return self.__class__(
            country="EC",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_el_salvador(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish ``El Salvador`` locale"""
        return self.__class__(
            country="SV",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_equatorial_guinea(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Equatorial Guinea locale ``es-GQ``"""
        return self.__class__(
            country="GQ",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_guatemala(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Guatemala locale"""
        return self.__class__(
            country="GT",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_honduras(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Honduras locale"""
        return self.__class__(
            country="HN",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_mexico(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Mexico locale"""
        return self.__class__(
            country="MX",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_nicaragua(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Nicaragua locale"""
        return self.__class__(
            country="NI",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_panama(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Panama locale"""
        return self.__class__(
            country="PA",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_paraguay(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Paraguay locale"""
        return self.__class__(
            country="PY",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_peru(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Peru locale"""
        return self.__class__(
            country="PE",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_philippines(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Philippines locale ``es-PH``"""
        return self.__class__(
            country="PH",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_puerto_rico(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Puerto Rico locale"""
        return self.__class__(
            country="PR",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_spain(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Spain locale"""
        return self.__class__(
            country="ES",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_usa(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish USA locale ``es-US``"""
        return self.__class__(
            country="US",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_uruguay(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Uruguay locale"""
        return self.__class__(
            country="UY",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_venezuela(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish Venezuela locale"""
        return self.__class__(
            country="VE",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    @property
    def spanish_es(self: _TLocaleStruct) -> _TLocaleStruct:
        """Gets Spanish locale ``es``"""
        return self.__class__(
            country="",
            language="es",
            variant="",
            _cattribs=self._get_internal_cattribs(),
        )

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_country(self) -> str:
        """Gets/Sets the ``ISO 3166`` Country Code"""
        return self._get("Country")

    @prop_country.setter
    def prop_country(self, value: str) -> None:
        self._set("Country", value.upper()[:2])

    @property
    def prop_language(self) -> str:
        """Gets/Sets the ``ISO 639`` Language Code."""
        return self._get("Language")

    @prop_language.setter
    def prop_language(self, value: str) -> None:
        self._set("Language", value)

    @property
    def prop_variant(self) -> str:
        """Gets/Sets the ``BCP 47`` Language Tag."""
        return self._get("Variant")

    @prop_variant.setter
    def prop_variant(self, value: str) -> None:
        self._set("Variant", value)

    # endregion properties
