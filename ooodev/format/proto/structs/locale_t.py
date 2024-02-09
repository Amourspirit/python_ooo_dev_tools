from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Dict
import uno

from ..style_t import StyleT

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class LocaleT(StyleT, Protocol):
    """Font Effect Protocol"""

    def __init__(self, *, country: str = ..., language: str = ..., variant: str = ...) -> None: ...

    @overload
    def apply(self, obj: Any, keys: Dict[str, str]) -> None:  # type: ignore
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> LocaleT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> LocaleT: ...

    # region Format Methods
    def fmt_country(self, value: str) -> LocaleT:
        """
        Gets a copy of instance with country set.

        Args:
            value (int): Country value.

        Returns:
            LocaleT: ``LocaleT`` instance
        """
        ...

    def fmt_language(self, value: str) -> LocaleT:
        """
        Gets a copy of instance with language set.

        Args:
            value (int): Language value.

        Returns:
            LocaleT: ``LocaleT`` instance
        """
        ...

    def fmt_variant(self, value: str) -> LocaleT:
        """
        Gets a copy of instance with variant set.

        Args:
            value (int): Variant value.

        Returns:
            LocaleT: ``LocaleT`` instance
        """
        ...

    # endregion Format Methods

    # region Style Properties
    @property
    def english_us(self) -> LocaleT:
        """Gets English US locale"""
        ...

    @property
    def locale_none(self) -> LocaleT:
        """Gets no locale"""
        ...

    @property
    def english_australia(self) -> LocaleT:
        """Gets English Australia locale"""
        ...

    @property
    def english_belize(self) -> LocaleT:
        """Gets English Belize locale"""
        ...

    @property
    def english_botswana(self) -> LocaleT:
        """Gets English Botswana locale"""
        ...

    @property
    def english_canada(self) -> LocaleT:
        """Gets English Canada locale"""
        ...

    @property
    def english_caribbean(self) -> LocaleT:
        """Gets English Caribbean locale"""
        ...

    @property
    def english_gambia(self) -> LocaleT:
        """Gets English Gambia locale"""
        ...

    @property
    def english_ghana(self) -> LocaleT:
        """Gets English Ghana locale"""
        ...

    @property
    def english_hong_kong(self) -> LocaleT:
        """Gets English Hong Kong locale"""
        ...

    @property
    def english_india(self) -> LocaleT:
        """Gets English India locale"""
        ...

    @property
    def english_ireland(self) -> LocaleT:
        """Gets English Ireland locale"""
        ...

    @property
    def english_israel(self) -> LocaleT:
        """Gets English Israel locale"""
        ...

    @property
    def english_jamaica(self) -> LocaleT:
        """Gets English Jamaica locale"""
        ...

    @property
    def english_kenya(self) -> LocaleT:
        """Gets English Kenya locale"""
        ...

    @property
    def english_malawi(self) -> LocaleT:
        """Gets English Malawi locale"""
        ...

    @property
    def english_malaysia(self) -> LocaleT:
        """Gets English Malaysia locale"""
        ...

    @property
    def english_mauritius(self) -> LocaleT:
        """Gets English Mauritius locale"""
        ...

    @property
    def english_namibia(self) -> LocaleT:
        """Gets English Namibia locale"""
        ...

    @property
    def english_new_zealand(self) -> LocaleT:
        """Gets English New Zealand locale"""
        ...

    @property
    def english_nigeria(self) -> LocaleT:
        """Gets English Nigeria locale"""
        ...

    @property
    def english_philippines(self) -> LocaleT:
        """Gets English Philippines locale"""
        ...

    @property
    def english_south_africa(self) -> LocaleT:
        """Gets English South Africa locale"""
        ...

    @property
    def english_sri_lanka(self) -> LocaleT:
        """Gets English ``Sri Lanka`` locale"""
        ...

    @property
    def english_trinidad(self) -> LocaleT:
        """Gets English Trinidad locale"""
        ...

    @property
    def english_uk(self) -> LocaleT:
        """Gets English UK locale"""
        ...

    @property
    def english_usa(self) -> LocaleT:
        """Gets English USA locale"""
        ...

    @property
    def english_zambia(self) -> LocaleT:
        """Gets English Zambia locale"""
        ...

    @property
    def english_zimbabwe(self) -> LocaleT:
        """Gets English Zimbabwe locale"""
        ...

    @property
    def english_uk_ode(self) -> LocaleT:
        """Gets English ODE Spelling UK locale"""
        ...

    @property
    def german_austria(self) -> LocaleT:
        """Gets German Austria locale"""
        ...

    @property
    def german_belgium(self) -> LocaleT:
        """Gets German Belgium locale"""
        ...

    @property
    def german_germany(self) -> LocaleT:
        """Gets German Germany locale"""
        ...

    @property
    def german_liechtenstein(self) -> LocaleT:
        """Gets German Liechtenstein locale"""
        ...

    @property
    def german_luxembourg(self) -> LocaleT:
        """Gets German Luxembourg locale"""
        ...

    @property
    def german_switzerland(self) -> LocaleT:
        """Gets German Switzerland locale"""
        ...

    @property
    def french_belgium(self) -> LocaleT:
        """Gets French Belgium locale"""
        ...

    @property
    def french_benin(self) -> LocaleT:
        """Gets French Benin locale"""
        ...

    @property
    def french_burkina_faso(self) -> LocaleT:
        """Gets French ``Burkina Faso`` locale"""
        ...

    @property
    def french_canada(self) -> LocaleT:
        """Gets French Canada locale"""
        ...

    @property
    def french_cote_d_ivoire(self) -> LocaleT:
        """Gets French ``Côte d'Ivoire`` locale"""
        ...

    @property
    def french_france(self) -> LocaleT:
        """Gets French France locale"""
        ...

    @property
    def french_luxembourg(self) -> LocaleT:
        """Gets French Luxembourg locale"""
        ...

    @property
    def french_mali(self) -> LocaleT:
        """Gets French Mali locale"""
        ...

    @property
    def french_mauritius(self) -> LocaleT:
        """Gets French Mauritius locale"""
        ...

    @property
    def french_monaco(self) -> LocaleT:
        """Gets French Monaco locale"""
        ...

    @property
    def french_niger(self) -> LocaleT:
        """Gets French Niger locale"""
        ...

    @property
    def french_senegal(self) -> LocaleT:
        """Gets French Senegal locale"""
        ...

    @property
    def french_switzerland(self) -> LocaleT:
        """Gets French Switzerland locale"""
        ...

    @property
    def french_togo(self) -> LocaleT:
        """Gets French Togo locale"""
        ...

    @property
    def spanish_argentina(self) -> LocaleT:
        """Gets Spanish Argentina locale"""
        ...

    @property
    def spanish_bolivia(self) -> LocaleT:
        """Gets Spanish Bolivia locale"""
        ...

    @property
    def spanish_chile(self) -> LocaleT:
        """Gets Spanish Chile locale"""
        ...

    @property
    def spanish_colombia(self) -> LocaleT:
        """Gets Spanish Colombia locale"""
        ...

    @property
    def spanish_costa_rica(self) -> LocaleT:
        """Gets Spanish ``Costa Rica`` locale"""
        ...

    @property
    def spanish_cuba(self) -> LocaleT:
        """Gets Spanish Cuba locale"""
        ...

    @property
    def spanish_dom_rep(self) -> LocaleT:
        """Gets Spanish Dominion Republic locale"""
        ...

    @property
    def spanish_ecuador(self) -> LocaleT:
        """Gets Spanish Ecuador locale"""
        ...

    @property
    def spanish_el_salvador(self) -> LocaleT:
        """Gets Spanish ``El Salvador`` locale"""
        ...

    @property
    def spanish_equatorial_guinea(self) -> LocaleT:
        """Gets Spanish Equatorial Guinea locale ``es-GQ``"""
        ...

    @property
    def spanish_guatemala(self) -> LocaleT:
        """Gets Spanish Guatemala locale"""
        ...

    @property
    def spanish_honduras(self) -> LocaleT:
        """Gets Spanish Honduras locale"""
        ...

    @property
    def spanish_mexico(self) -> LocaleT:
        """Gets Spanish Mexico locale"""
        ...

    @property
    def spanish_nicaragua(self) -> LocaleT:
        """Gets Spanish Nicaragua locale"""
        ...

    @property
    def spanish_panama(self) -> LocaleT:
        """Gets Spanish Panama locale"""
        ...

    @property
    def spanish_paraguay(self) -> LocaleT:
        """Gets Spanish Paraguay locale"""
        ...

    @property
    def spanish_peru(self) -> LocaleT:
        """Gets Spanish Peru locale"""
        ...

    @property
    def spanish_philippines(self) -> LocaleT:
        """Gets Spanish Philippines locale ``es-PH``"""
        ...

    @property
    def spanish_puerto_rico(self) -> LocaleT:
        """Gets Spanish Puerto Rico locale"""
        ...

    @property
    def spanish_spain(self) -> LocaleT:
        """Gets Spanish Spain locale"""
        ...

    @property
    def spanish_usa(self) -> LocaleT:
        """Gets Spanish USA locale ``es-US``"""
        ...

    @property
    def spanish_uruguay(self) -> LocaleT:
        """Gets Spanish Uruguay locale"""
        ...

    @property
    def spanish_venezuela(self) -> LocaleT:
        """Gets Spanish Venezuela locale"""
        ...

    @property
    def spanish_es(self) -> LocaleT:
        """Gets Spanish locale ``es``"""
        ...

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_country(self) -> str:
        """Gets/Sets the ``ISO 3166`` Country Code"""
        ...

    @prop_country.setter
    def prop_country(self, value: str) -> None: ...

    @property
    def prop_language(self) -> str:
        """Gets/Sets the ``ISO 639`` Language Code."""
        ...

    @prop_language.setter
    def prop_language(self, value: str) -> None: ...

    @property
    def prop_variant(self) -> str:
        """Gets/Sets the ``BCP 47`` Language Tag."""
        ...

    @prop_variant.setter
    def prop_variant(self, value: str) -> None: ...

    # endregion Prop Properties
