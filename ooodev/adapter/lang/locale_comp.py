from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.lang.locale import Locale
from ooodev.adapter.struct_base import StructBase

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class LocaleComp(StructBase[Locale]):
    """
    Locale Struct

    Object represents a specific geographical, political, or cultural region.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_lang_Locale_changing``.
    The event raised after the property is changed is called ``com_sun_star_lang_Locale_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: Locale, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (Locale): Border Line.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_lang_Locale_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_lang_Locale_changed"

    def _copy(self, src: Locale | None = None) -> Locale:
        if src is None:
            src = self.component
        return Locale(
            Language=src.Language,
            Country=src.Country,
            Variant=src.Variant,
        )

    # endregion Overrides

    # region Properties

    @property
    def language(self) -> str:
        """
        Gets/Sets - an ISO 639 Language Code.

        These codes are preferably the lower-case two-letter codes as defined by ISO 639-1, or three-letter codes as defined by ISO 639-3.
        If this field contains an empty string, the meaning depends on the context.
        """
        return self.component.Language

    @language.setter
    def language(self, value: str) -> None:
        old_value = self.component.Language
        if old_value != value:
            event_args = self._trigger_cancel_event("Language", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def country(self) -> str:
        """
        Gets/Sets - an ISO 3166 Country Code.

        These codes are the upper-case two-letter codes as defined by ISO 3166-1.
        If this field contains an empty string, the meaning depends on the context.
        """
        return self.component.Country

    @country.setter
    def country(self, value: str) -> None:
        old_value = self.component.Country
        if old_value != value:
            event_args = self._trigger_cancel_event("Country", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def variant(self) -> str:
        """
        Gets/Sets - a BCP 47 Language Tag.
        """
        return self.component.Variant

    @variant.setter
    def variant(self, value: str) -> None:
        old_value = self.component.Variant
        if old_value != value:
            event_args = self._trigger_cancel_event("Variant", old_value, value)
            self._trigger_done_event(event_args)

    # endregion Properties
