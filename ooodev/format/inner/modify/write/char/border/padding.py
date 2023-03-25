# region Imports
from __future__ import annotations
from typing import cast
import uno
from ooodev.units import UnitObj
from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ooodev.format.inner.direct.write.char.border.padding import Padding as InnerPadding
from ..char_style_base_multi import CharStyleBaseMulti

# endregion Imports


class Padding(CharStyleBaseMulti):
    """
    Character Style Border padding.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitObj | None = None,
        right: float | UnitObj | None = None,
        top: float | UnitObj | None = None,
        bottom: float | UnitObj | None = None,
        all: float | UnitObj | None = None,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitObj, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitObj, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitObj, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitObj,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitObj, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`. If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            None:
        """

        direct = InnerPadding(left=left, right=right, top=top, bottom=bottom, all=all)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> Padding:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            Padding: ``Padding`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPadding.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCharKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerPadding:
        """Gets/Sets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPadding, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerPadding) -> None:
        if not isinstance(value, InnerPadding):
            raise TypeError(f'Expected type of InnerPadding, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
