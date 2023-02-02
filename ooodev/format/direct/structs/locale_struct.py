"""
Module for ``DropCapFormat`` struct.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Dict, Tuple, cast, overload

import uno
from ....events.event_singleton import _Events
from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....utils.data_type.byte import Byte
from ....utils.type_var import T
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent


from ooo.dyn.lang.locale import Locale


class LocaleStruct(StyleBase):
    """
    Paragraph Drop Cap

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, country: str = "US", language: str = "en", variant: str = "") -> None:
        """
        Constructor

        Args:
            country (str, optional): Specifies an ``ISO 3166`` Country Code.
                These codes are the upper-case two-letter codes as defined by ``ISO 3166-1``. You can find a full list of these codes at a number of sites, such as:
                `Wikipedia ISO 3166-1 alpha-2 <https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2>`__.
                If this field contains an empty string, the meaning depends on the context.
            language (str, opttional): Specifies an ``ISO 639`` Language Code.
                These codes are preferably the lower-case two-letter codes as defined by ``ISO 639-1``, or three-letter codes as defined by ``ISO 639-3``.
                You can find a full list of these codes at a number of sites, such as:
                `ISO 639 Code Tables <https://iso639-3.sil.org/code_tables/639/data>`__.
                If this field contains an empty string, the meaning depends on the context.
            lines (str, optional): Specifies a ``BCP 47`` Language Tag.
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
        """
        Gets a tuple of supported services.
        This is an empty value for this class but may be different for child classes.

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ()

    def _get_property_name(self) -> str:
        return "CharLocale"

    def _is_valid_obj(self, obj: object) -> bool:
        return mProps.Props.has(obj, self._get_property_name())

    def copy(self: T) -> T:
        nu = super(LocaleStruct, self.__class__).__new__(self.__class__)
        nu.__init__()
        if self._dv:
            nu._update(self._dv)
        return nu

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, keys: Dict[str, str]) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.
            keys (Dict[str, str], optional): Property key, value items that map properties.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply")
            return

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["prop"]
        dcf = self.get_locale()
        mProps.Props.set(obj, **{key: dcf})
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

        # mProps.Props.set(obj, **{key: tuple(tss_lst)})

    # endregion apply()

    def get_locale(self) -> Locale:
        """
        Gets locale for instance

        Returns:
            Locale: ``Locale`` instance
        """
        return Locale(Language=self._get("Language"), Country=self._get("Country"), Variant=self._get("Variant"))

    @classmethod
    def from_obj(cls, obj: object) -> LocaleStruct:
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
        nu = super(LocaleStruct, cls).__new__(cls)
        nu.__init__()
        prop_name = nu._get_property_name()

        try:
            dcf = cast(Locale, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_locale(dcf)

    @classmethod
    def from_locale(cls, locale: Locale) -> LocaleStruct:
        """
        Converts a ``Locale`` Stop instance to a ``LocaleStruct``

        Args:
            dcf (Locale): UNO locale

        Returns:
            LocaleStruct: ``LocaleStruct`` set with locale properties
        """
        inst = super(LocaleStruct, cls).__new__(cls)
        inst.__init__(country=locale.Country, language=locale.Language, variant=locale.Variant)
        return inst

    # endregion methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, LocaleStruct):
            obj2 = oth.get_locale()
        if getattr(oth, "typeName", None) == "com.sun.star.lang.Locale":
            obj2 = cast(Locale, oth)
        if not obj2 is None:
            obj1 = self.get_locale()
            return obj1.Country == obj2.Country and obj1.Language == obj2.Language and obj1.Variant == obj2.Variant
        return NotImplemented

    # endregion dunder methods

    # region format methods
    def fmt_country(self, value: str) -> LocaleStruct:
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

    def fmt_language(self, value: str) -> LocaleStruct:
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

    def fmt_variant(self, value: str) -> LocaleStruct:
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
    def english_us(self) -> LocaleStruct:
        """Gets English US locale"""
        return LocaleStruct(country="US", language="en", variant="")

    @property
    def locale_none(self) -> LocaleStruct:
        """Gets no locale"""
        return LocaleStruct(country="", language="zxx", variant="")

    @property
    def english_australia(self) -> LocaleStruct:
        """Gets English Australia locale"""
        return LocaleStruct(country="AU", language="en", variant="")

    @property
    def english_belize(self) -> LocaleStruct:
        """Gets English Belize locale"""
        return LocaleStruct(country="BZ", language="en", variant="")

    @property
    def english_botswana(self) -> LocaleStruct:
        """Gets English Botswana locale"""
        return LocaleStruct(country="BW", language="en", variant="")

    @property
    def english_canada(self) -> LocaleStruct:
        """Gets English Canada locale"""
        return LocaleStruct(country="CA", language="en", variant="")

    @property
    def english_caribbean(self) -> LocaleStruct:
        """Gets English Caribbean locale"""
        return LocaleStruct(country="BS", language="en", variant="")

    @property
    def english_gambia(self) -> LocaleStruct:
        """Gets English Gambia locale"""
        return LocaleStruct(country="GM", language="en", variant="")

    @property
    def english_ghana(self) -> LocaleStruct:
        """Gets English Ghana locale"""
        return LocaleStruct(country="GH", language="en", variant="")

    @property
    def english_hong_kong(self) -> LocaleStruct:
        """Gets English Hong Kong locale"""
        return LocaleStruct(country="HK", language="en", variant="")

    @property
    def english_india(self) -> LocaleStruct:
        """Gets English India locale"""
        return LocaleStruct(country="IN", language="en", variant="")

    @property
    def english_ireland(self) -> LocaleStruct:
        """Gets English Ireland locale"""
        return LocaleStruct(country="IE", language="en", variant="")

    @property
    def english_ireland(self) -> LocaleStruct:
        """Gets English Ireland locale"""
        return LocaleStruct(country="IE", language="en", variant="")

    @property
    def english_israel(self) -> LocaleStruct:
        """Gets English Israel locale"""
        return LocaleStruct(country="IL", language="en", variant="")

    @property
    def english_jamaica(self) -> LocaleStruct:
        """Gets English Jamaica locale"""
        return LocaleStruct(country="JM", language="en", variant="")

    @property
    def english_kenya(self) -> LocaleStruct:
        """Gets English Kenya locale"""
        return LocaleStruct(country="KE", language="en", variant="")

    @property
    def english_malawi(self) -> LocaleStruct:
        """Gets English Malawi locale"""
        return LocaleStruct(country="MW", language="en", variant="")

    @property
    def english_malaysia(self) -> LocaleStruct:
        """Gets English Malaysia locale"""
        return LocaleStruct(country="MY", language="en", variant="")

    @property
    def english_mauritius(self) -> LocaleStruct:
        """Gets English Mauritius locale"""
        return LocaleStruct(country="MU", language="en", variant="")

    @property
    def english_namibia(self) -> LocaleStruct:
        """Gets English Namibia locale"""
        return LocaleStruct(country="NA", language="en", variant="")

    @property
    def english_new_zealand(self) -> LocaleStruct:
        """Gets English New Zealand locale"""
        return LocaleStruct(country="NZ", language="en", variant="")

    @property
    def english_nigeria(self) -> LocaleStruct:
        """Gets English Nigeria locale"""
        return LocaleStruct(country="NG", language="en", variant="")

    @property
    def english_philippines(self) -> LocaleStruct:
        """Gets English Philippines locale"""
        return LocaleStruct(country="PH", language="en", variant="")

    @property
    def english_south_africa(self) -> LocaleStruct:
        """Gets English South Africa locale"""
        return LocaleStruct(country="ZA", language="en", variant="")

    @property
    def english_sri_lanka(self) -> LocaleStruct:
        """Gets English Sri Lanka locale"""
        return LocaleStruct(country="LK", language="en", variant="")

    @property
    def english_trinidad(self) -> LocaleStruct:
        """Gets English Trinidad locale"""
        return LocaleStruct(country="TT", language="en", variant="")

    @property
    def english_uk(self) -> LocaleStruct:
        """Gets English UK locale"""
        return LocaleStruct(country="GB", language="en", variant="")

    @property
    def english_usa(self) -> LocaleStruct:
        """Gets English USA locale"""
        return LocaleStruct(country="US", language="en", variant="")

    @property
    def english_zambia(self) -> LocaleStruct:
        """Gets English Zambia locale"""
        return LocaleStruct(country="ZM", language="en", variant="")

    @property
    def english_zimbabwe(self) -> LocaleStruct:
        """Gets English Zimbabwe locale"""
        return LocaleStruct(country="ZW", language="en", variant="")

    @property
    def english_uk_ode(self) -> LocaleStruct:
        """Gets English ODE Spelling UK locale"""
        return LocaleStruct(country="GB", language="qlt", variant="en-GB-oxendict")

    @property
    def german_austria(self) -> LocaleStruct:
        """Gets German Austria locale"""
        return LocaleStruct(country="AT", language="de", variant="")

    @property
    def german_belgium(self) -> LocaleStruct:
        """Gets German Belgium locale"""
        return LocaleStruct(country="BE", language="de", variant="")

    @property
    def german_germany(self) -> LocaleStruct:
        """Gets German Germany locale"""
        return LocaleStruct(country="DE", language="de", variant="")

    @property
    def german_liechtenstein(self) -> LocaleStruct:
        """Gets German Liechtenstein locale"""
        return LocaleStruct(country="LI", language="de", variant="")

    @property
    def german_luxembourg(self) -> LocaleStruct:
        """Gets German Luxembourg locale"""
        return LocaleStruct(country="LU", language="de", variant="")

    @property
    def german_switzerland(self) -> LocaleStruct:
        """Gets German Switzerland locale"""
        return LocaleStruct(country="CH", language="de", variant="")

    @property
    def french_belgium(self) -> LocaleStruct:
        """Gets French Belgium locale"""
        return LocaleStruct(country="BE", language="fr", variant="")

    @property
    def french_benin(self) -> LocaleStruct:
        """Gets French Benin locale"""
        return LocaleStruct(country="BJ", language="fr", variant="")

    @property
    def french_benin(self) -> LocaleStruct:
        """Gets French Benin locale"""
        return LocaleStruct(country="BJ", language="fr", variant="")

    @property
    def french_burkina_faso(self) -> LocaleStruct:
        """Gets French Burkina Faso locale"""
        return LocaleStruct(country="BF", language="fr", variant="")

    @property
    def french_canada(self) -> LocaleStruct:
        """Gets French Canada locale"""
        return LocaleStruct(country="CA", language="fr", variant="")

    @property
    def french_cote_d_ivoire(self) -> LocaleStruct:
        """Gets French (Côte d'Ivoire) locale"""
        return LocaleStruct(country="CI", language="fr", variant="")

    @property
    def french_france(self) -> LocaleStruct:
        """Gets French France locale"""
        return LocaleStruct(country="FR", language="fr", variant="")

    @property
    def french_luxembourg(self) -> LocaleStruct:
        """Gets French Luxembourg locale"""
        return LocaleStruct(country="LU", language="fr", variant="")

    @property
    def french_mali(self) -> LocaleStruct:
        """Gets French Mali locale"""
        return LocaleStruct(country="ML", language="fr", variant="")

    @property
    def french_mauritius(self) -> LocaleStruct:
        """Gets French Mauritius locale"""
        return LocaleStruct(country="MU", language="fr", variant="")

    @property
    def french_monaco(self) -> LocaleStruct:
        """Gets French Monaco locale"""
        return LocaleStruct(country="MC", language="fr", variant="")

    @property
    def french_niger(self) -> LocaleStruct:
        """Gets French Niger locale"""
        return LocaleStruct(country="NE", language="fr", variant="")

    @property
    def french_senegal(self) -> LocaleStruct:
        """Gets French Senegal locale"""
        return LocaleStruct(country="SN", language="fr", variant="")

    @property
    def french_switzerland(self) -> LocaleStruct:
        """Gets French Switzerland locale"""
        return LocaleStruct(country="CH", language="fr", variant="")

    @property
    def french_togo(self) -> LocaleStruct:
        """Gets French Togo locale"""
        return LocaleStruct(country="TG", language="fr", variant="")

    @property
    def spanish_argentina(self) -> LocaleStruct:
        """Gets Spanish Argentina locale"""
        return LocaleStruct(country="AR", language="es", variant="")

    @property
    def spanish_bolivia(self) -> LocaleStruct:
        """Gets Spanish Bolivia locale"""
        return LocaleStruct(country="BO", language="es", variant="")

    @property
    def spanish_chile(self) -> LocaleStruct:
        """Gets Spanish Chile locale"""
        return LocaleStruct(country="CL", language="es", variant="")

    @property
    def spanish_colombia(self) -> LocaleStruct:
        """Gets Spanish Colombia locale"""
        return LocaleStruct(country="CO", language="es", variant="")

    @property
    def spanish_costa_rica(self) -> LocaleStruct:
        """Gets Spanish Costa Rica locale"""
        return LocaleStruct(country="CR", language="es", variant="")

    @property
    def spanish_cuba(self) -> LocaleStruct:
        """Gets Spanish Cuba locale"""
        return LocaleStruct(country="CU", language="es", variant="")

    @property
    def spanish_dom_rep(self) -> LocaleStruct:
        """Gets Spanish Dominion Republic locale"""
        return LocaleStruct(country="DO", language="es", variant="")

    @property
    def spanish_ecuador(self) -> LocaleStruct:
        """Gets Spanish Ecuador locale"""
        return LocaleStruct(country="EC", language="es", variant="")

    @property
    def spanish_el_salvador(self) -> LocaleStruct:
        """Gets Spanish El Salvador locale"""
        return LocaleStruct(country="SV", language="es", variant="")

    @property
    def spanish_equatorial_guinea(self) -> LocaleStruct:
        """Gets Spanish Equatorial Guinea locale ``es-GQ``"""
        return LocaleStruct(country="GQ", language="es", variant="")

    @property
    def spanish_guatemala(self) -> LocaleStruct:
        """Gets Spanish Guatemala locale"""
        return LocaleStruct(country="GT", language="es", variant="")

    @property
    def spanish_honduras(self) -> LocaleStruct:
        """Gets Spanish Honduras locale"""
        return LocaleStruct(country="HN", language="es", variant="")

    @property
    def spanish_mexico(self) -> LocaleStruct:
        """Gets Spanish Mexico locale"""
        return LocaleStruct(country="MX", language="es", variant="")

    @property
    def spanish_nicaragua(self) -> LocaleStruct:
        """Gets Spanish Nicaragua locale"""
        return LocaleStruct(country="NI", language="es", variant="")

    @property
    def spanish_panama(self) -> LocaleStruct:
        """Gets Spanish Panama locale"""
        return LocaleStruct(country="PA", language="es", variant="")

    @property
    def spanish_paraguay(self) -> LocaleStruct:
        """Gets Spanish Paraguay locale"""
        return LocaleStruct(country="PY", language="es", variant="")

    @property
    def spanish_peru(self) -> LocaleStruct:
        """Gets Spanish Peru locale"""
        return LocaleStruct(country="PE", language="es", variant="")

    @property
    def spanish_philippines(self) -> LocaleStruct:
        """Gets Spanish Philippines locale ``es-PH``"""
        return LocaleStruct(country="PH", language="es", variant="")

    @property
    def spanish_puerto_rico(self) -> LocaleStruct:
        """Gets Spanish Puerto Rico locale"""
        return LocaleStruct(country="PR", language="es", variant="")

    @property
    def spanish_spain(self) -> LocaleStruct:
        """Gets Spanish Spain locale"""
        return LocaleStruct(country="ES", language="es", variant="")

    @property
    def spanish_usa(self) -> LocaleStruct:
        """Gets Spanish USA locale ``es-US``"""
        return LocaleStruct(country="US", language="es", variant="")

    @property
    def spanish_uruguay(self) -> LocaleStruct:
        """Gets Spanish Uruguay locale"""
        return LocaleStruct(country="UY", language="es", variant="")

    @property
    def spanish_venezuela(self) -> LocaleStruct:
        """Gets Spanish Venezuela locale"""
        return LocaleStruct(country="VE", language="es", variant="")

    @property
    def spanish_es(self) -> LocaleStruct:
        """Gets Spanish locale ``es``"""
        return LocaleStruct(country="", language="es", variant="")

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

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
