from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Dict
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
else:
    Protocol = object
    Self = Any


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
    def fmt_country(self, value: str) -> Self:
        """
        Gets a copy of instance with country set.

        Args:
            value (int): Country value.

        Returns:
            LocaleT: ``LocaleT`` instance
        """
        ...

    def fmt_language(self, value: str) -> Self:
        """
        Gets a copy of instance with language set.

        Args:
            value (int): Language value.

        Returns:
            LocaleT: ``LocaleT`` instance
        """
        ...

    def fmt_variant(self, value: str) -> Self:
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
    def english_us(self) -> Self:
        """Gets English US locale"""
        ...

    @property
    def locale_none(self) -> Self:
        """Gets no locale"""
        ...

    @property
    def english_australia(self) -> Self:
        """Gets English Australia locale"""
        ...

    @property
    def english_belize(self) -> Self:
        """Gets English Belize locale"""
        ...

    @property
    def english_botswana(self) -> Self:
        """Gets English Botswana locale"""
        ...

    @property
    def english_canada(self) -> Self:
        """Gets English Canada locale"""
        ...

    @property
    def english_caribbean(self) -> Self:
        """Gets English Caribbean locale"""
        ...

    @property
    def english_gambia(self) -> Self:
        """Gets English Gambia locale"""
        ...

    @property
    def english_ghana(self) -> Self:
        """Gets English Ghana locale"""
        ...

    @property
    def english_hong_kong(self) -> Self:
        """Gets English Hong Kong locale"""
        ...

    @property
    def english_india(self) -> Self:
        """Gets English India locale"""
        ...

    @property
    def english_ireland(self) -> Self:
        """Gets English Ireland locale"""
        ...

    @property
    def english_israel(self) -> Self:
        """Gets English Israel locale"""
        ...

    @property
    def english_jamaica(self) -> Self:
        """Gets English Jamaica locale"""
        ...

    @property
    def english_kenya(self) -> Self:
        """Gets English Kenya locale"""
        ...

    @property
    def english_malawi(self) -> Self:
        """Gets English Malawi locale"""
        ...

    @property
    def english_malaysia(self) -> Self:
        """Gets English Malaysia locale"""
        ...

    @property
    def english_mauritius(self) -> Self:
        """Gets English Mauritius locale"""
        ...

    @property
    def english_namibia(self) -> Self:
        """Gets English Namibia locale"""
        ...

    @property
    def english_new_zealand(self) -> Self:
        """Gets English New Zealand locale"""
        ...

    @property
    def english_nigeria(self) -> Self:
        """Gets English Nigeria locale"""
        ...

    @property
    def english_philippines(self) -> Self:
        """Gets English Philippines locale"""
        ...

    @property
    def english_south_africa(self) -> Self:
        """Gets English South Africa locale"""
        ...

    @property
    def english_sri_lanka(self) -> Self:
        """Gets English ``Sri Lanka`` locale"""
        ...

    @property
    def english_trinidad(self) -> Self:
        """Gets English Trinidad locale"""
        ...

    @property
    def english_uk(self) -> Self:
        """Gets English UK locale"""
        ...

    @property
    def english_usa(self) -> Self:
        """Gets English USA locale"""
        ...

    @property
    def english_zambia(self) -> Self:
        """Gets English Zambia locale"""
        ...

    @property
    def english_zimbabwe(self) -> Self:
        """Gets English Zimbabwe locale"""
        ...

    @property
    def english_uk_ode(self) -> Self:
        """Gets English ODE Spelling UK locale"""
        ...

    @property
    def german_austria(self) -> Self:
        """Gets German Austria locale"""
        ...

    @property
    def german_belgium(self) -> Self:
        """Gets German Belgium locale"""
        ...

    @property
    def german_germany(self) -> Self:
        """Gets German Germany locale"""
        ...

    @property
    def german_liechtenstein(self) -> Self:
        """Gets German Liechtenstein locale"""
        ...

    @property
    def german_luxembourg(self) -> Self:
        """Gets German Luxembourg locale"""
        ...

    @property
    def german_switzerland(self) -> Self:
        """Gets German Switzerland locale"""
        ...

    @property
    def french_belgium(self) -> Self:
        """Gets French Belgium locale"""
        ...

    @property
    def french_benin(self) -> Self:
        """Gets French Benin locale"""
        ...

    @property
    def french_burkina_faso(self) -> Self:
        """Gets French ``Burkina Faso`` locale"""
        ...

    @property
    def french_canada(self) -> Self:
        """Gets French Canada locale"""
        ...

    @property
    def french_cote_d_ivoire(self) -> Self:
        """Gets French ``Côte d'Ivoire`` locale"""
        ...

    @property
    def french_france(self) -> Self:
        """Gets French France locale"""
        ...

    @property
    def french_luxembourg(self) -> Self:
        """Gets French Luxembourg locale"""
        ...

    @property
    def french_mali(self) -> Self:
        """Gets French Mali locale"""
        ...

    @property
    def french_mauritius(self) -> Self:
        """Gets French Mauritius locale"""
        ...

    @property
    def french_monaco(self) -> Self:
        """Gets French Monaco locale"""
        ...

    @property
    def french_niger(self) -> Self:
        """Gets French Niger locale"""
        ...

    @property
    def french_senegal(self) -> Self:
        """Gets French Senegal locale"""
        ...

    @property
    def french_switzerland(self) -> Self:
        """Gets French Switzerland locale"""
        ...

    @property
    def french_togo(self) -> Self:
        """Gets French Togo locale"""
        ...

    @property
    def spanish_argentina(self) -> Self:
        """Gets Spanish Argentina locale"""
        ...

    @property
    def spanish_bolivia(self) -> Self:
        """Gets Spanish Bolivia locale"""
        ...

    @property
    def spanish_chile(self) -> Self:
        """Gets Spanish Chile locale"""
        ...

    @property
    def spanish_colombia(self) -> Self:
        """Gets Spanish Colombia locale"""
        ...

    @property
    def spanish_costa_rica(self) -> Self:
        """Gets Spanish ``Costa Rica`` locale"""
        ...

    @property
    def spanish_cuba(self) -> Self:
        """Gets Spanish Cuba locale"""
        ...

    @property
    def spanish_dom_rep(self) -> Self:
        """Gets Spanish Dominion Republic locale"""
        ...

    @property
    def spanish_ecuador(self) -> Self:
        """Gets Spanish Ecuador locale"""
        ...

    @property
    def spanish_el_salvador(self) -> Self:
        """Gets Spanish ``El Salvador`` locale"""
        ...

    @property
    def spanish_equatorial_guinea(self) -> Self:
        """Gets Spanish Equatorial Guinea locale ``es-GQ``"""
        ...

    @property
    def spanish_guatemala(self) -> Self:
        """Gets Spanish Guatemala locale"""
        ...

    @property
    def spanish_honduras(self) -> Self:
        """Gets Spanish Honduras locale"""
        ...

    @property
    def spanish_mexico(self) -> Self:
        """Gets Spanish Mexico locale"""
        ...

    @property
    def spanish_nicaragua(self) -> Self:
        """Gets Spanish Nicaragua locale"""
        ...

    @property
    def spanish_panama(self) -> Self:
        """Gets Spanish Panama locale"""
        ...

    @property
    def spanish_paraguay(self) -> Self:
        """Gets Spanish Paraguay locale"""
        ...

    @property
    def spanish_peru(self) -> Self:
        """Gets Spanish Peru locale"""
        ...

    @property
    def spanish_philippines(self) -> Self:
        """Gets Spanish Philippines locale ``es-PH``"""
        ...

    @property
    def spanish_puerto_rico(self) -> Self:
        """Gets Spanish Puerto Rico locale"""
        ...

    @property
    def spanish_spain(self) -> Self:
        """Gets Spanish Spain locale"""
        ...

    @property
    def spanish_usa(self) -> Self:
        """Gets Spanish USA locale ``es-US``"""
        ...

    @property
    def spanish_uruguay(self) -> Self:
        """Gets Spanish Uruguay locale"""
        ...

    @property
    def spanish_venezuela(self) -> Self:
        """Gets Spanish Venezuela locale"""
        ...

    @property
    def spanish_es(self) -> Self:
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
