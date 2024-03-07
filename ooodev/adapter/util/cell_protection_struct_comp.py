from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.util.cell_protection import CellProtection

from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.args.key_val_args import KeyValArgs

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class CellProtectionStructComp(ComponentBase):
    """
    Cell Protection Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_util_CellProtection_changing``.
    The event raised after the property is changed is called ``com_sun_star_util_CellProtection_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: CellProtection, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (CellProtection): Font Descriptor.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        ComponentBase.__init__(self, component)
        self._event_provider = event_provider
        self._prop_name = prop_name

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # PropertySetPartial will validate
        return ()

    # endregion Overrides

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_util_CellProtection_changed"

    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_util_CellProtection_changing"

    def _get_prop_name(self) -> str:
        return self._prop_name

    def _on_property_changing(self, event_args: KeyValCancelArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event(self._get_on_changing_event_name(), event_args)

    def _on_property_changed(self, event_args: KeyValArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event(self._get_on_changed_event_name(), event_args)

    def _copy(self, src: CellProtection | None = None) -> CellProtection:
        if src is None:
            src = self.component
        return CellProtection(
            IsLocked=src.IsLocked,
            IsFormulaHidden=src.IsFormulaHidden,
            IsHidden=src.IsHidden,
            IsPrintHidden=src.IsPrintHidden,
        )

    def copy(self) -> CellProtection:
        """
        Makes a copy of the Border Line.

        Returns:
            CellProtection: Copied Border Line.
        """
        return self._copy()

    # region Properties

    @property
    def component(self) -> CellProtection:
        """CellProtection Component"""
        # pylint: disable=no-member
        return cast("CellProtection", self._ComponentBase__get_component())  # type: ignore

    @component.setter
    def component(self, value: CellProtection) -> None:
        # pylint: disable=no-member
        self._ComponentBase__set_component(self._copy(src=value))  # type: ignore

    @property
    def is_locked(self) -> bool:
        """
        Gets/Sets if the cell is locked from modifications by the user.
        """
        return self.component.IsLocked  # type: ignore

    @is_locked.setter
    def is_locked(self, value: bool) -> None:
        old_value = self.component.IsLocked
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="is_locked",
                value=value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.IsLocked = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def is_formula_hidden(self) -> bool:
        """
        Gets/Sets if the formula is hidden from the user.
        """
        return self.component.IsFormulaHidden  # type: ignore

    @is_formula_hidden.setter
    def is_formula_hidden(self, value: bool) -> None:
        old_value = self.component.IsFormulaHidden
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="is_formula_hidden",
                value=value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.IsFormulaHidden = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def is_hidden(self) -> bool:
        """
        Gets/Sets if the cell is hidden from the user.
        """
        return self.component.IsHidden  # type: ignore

    @is_hidden.setter
    def is_hidden(self, value: bool) -> None:
        old_value = self.component.IsHidden
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="is_hidden",
                value=value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.IsHidden = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def is_print_hidden(self) -> bool:
        """
        Gets/Sets if the cell is hidden on printouts.
        """
        return self.component.IsPrintHidden  # type: ignore

    @is_print_hidden.setter
    def is_print_hidden(self, value: bool) -> None:
        old_value = self.component.IsPrintHidden
        if old_value != value:
            event_args = KeyValCancelArgs(
                source=self,
                key="is_print_hidden",
                value=value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.IsPrintHidden = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    # endregion Properties
