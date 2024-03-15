from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.drawing.glue_point2 import GluePoint2
from ooo.dyn.awt.point import Point

# from ooodev.adapter.awt.point_struct_comp import PointStructComp
from ooodev.adapter.awt.point_struct_generic_comp import PointStructGenericComp

from ooodev.adapter.struct_base import StructBase
from ooodev.events.events import Events
from ooodev.utils import info as mInfo
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint

if TYPE_CHECKING:
    from ooo.dyn.drawing.escape_direction import EscapeDirection
    from ooo.dyn.drawing.alignment import Alignment
    from ooodev.events.events_t import EventsT
    from ooodev.events.args.key_val_args import KeyValArgs


class GluePoint2StructComp(StructBase[GluePoint2]):
    """
    GluePoint2 Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_drawing_GluePoint2_changing``.
    The event raised after the property is changed is called ``com_sun_star_drawing_GluePoint2_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: GluePoint2, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (GluePoint2): Border Line.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)
        self._event_provider = Events(self)
        self._props = {}

        def on_comp_struct_changed(src: PointStructGenericComp[UnitMM100], event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            setattr(self, prop_name, src.component)

        self._fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self._event_provider.subscribe_event("generic_com_sun_star_awt_Point_changed", self._fn_on_comp_struct_changed)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_drawing_GluePoint2_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_drawing_GluePoint2_changed"

    def _copy(self, src: GluePoint2 | None = None) -> GluePoint2:
        def copy_point(point: Point) -> Point:
            return Point(X=point.X, Y=point.Y)

        if src is None:
            src = self.component
        return GluePoint2(
            Position=copy_point(src.Position),
            IsRelative=src.IsRelative,
            PositionAlignment=src.PositionAlignment,
            Escape=src.Escape,
        )

    # endregion Overrides

    # region Properties

    @property
    def position(self) -> PointStructGenericComp[UnitMM100]:
        """
        This is the position of this glue point.

        Depending on the flag ``is_relative``, this is either in ``1/100cm`` or in ``1/100%``.

        When setting this value, it can be set with a ``Point`` instance or a ``PointStructGenericComp`` instance.

        Returns:
            PointStructComp: Position
        """
        key = "position"
        prop = self._props.get(key, None)
        if prop is None:
            prop = PointStructGenericComp(self.component.Position, UnitMM100, key, self._event_provider)
            self._props[key] = prop
        return cast(PointStructGenericComp[UnitMM100], prop)

    @position.setter
    def position(self, value: Point | PointStructGenericComp[UnitMM100]) -> None:
        key = "position"
        if mInfo.Info.is_instance(value, PointStructGenericComp):
            self.component.Position = value.copy()
        else:
            self.component.Position = cast("Point", value)
        if key in self._props:
            del self._props[key]

    @property
    def is_relative(self) -> bool:
        """
        Gets/Sets - If this flag is set to ``True``, the position of this glue point is given in ``1/100%`` values instead of ``1/100cm``.
        """
        return self.component.IsRelative

    @is_relative.setter
    def is_relative(self, value: bool) -> None:
        old_value = self.component.IsRelative
        if old_value != value:
            event_args = self._trigger_cancel_event("IsRelative", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def position_alignment(self) -> Alignment:
        """
        Gets/Sets - if this glue points position is not relative, this enum specifies the vertical and horizontal alignment of this point.

        The alignment specifies how the glue point is moved if the shape is resized.

        Returns:
            Alignment: Position Alignment

        Hint:
            - ``Alignment`` can be imported from ``ooo.dyn.drawing.alignment``.
        """
        return self.component.PositionAlignment  # type: ignore

    @position_alignment.setter
    def position_alignment(self, value: Alignment) -> None:
        old_value = self.component.PositionAlignment
        if old_value != value:
            event_args = self._trigger_cancel_event("PositionAlignment", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def escape(self) -> EscapeDirection:
        """
        Gets/Sets the escape direction for a glue point.

        The escape direction is the direction the connecting line escapes the shape.

        Returns:
            EscapeDirection: Escape Direction

        Hint:
            - ``EscapeDirection`` can be imported from ``ooo.dyn.drawing.escape_direction``.
        """
        return self.component.Escape  # type: ignore

    @escape.setter
    def escape(self, value: EscapeDirection) -> None:
        old_value = self.component.Escape
        if old_value != value:
            event_args = self._trigger_cancel_event("Escape", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def is_user_defined(self) -> bool:
        """
        Gets/Sets - If this flag is set to ``False``, this is a default glue point.

        Some shapes may have default glue points attached to them which cannot be altered or removed.
        """
        return self.component.IsUserDefined

    @is_user_defined.setter
    def is_user_defined(self, value: bool) -> None:
        old_value = self.component.IsUserDefined
        if old_value != value:
            event_args = self._trigger_cancel_event("IsUserDefined", old_value, value)
            self._trigger_done_event(event_args)

    # endregion Properties
