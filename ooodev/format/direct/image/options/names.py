from __future__ import annotations
from typing import Tuple
from ...frame.options.names import Names as FrameNames
from ...common.props.image_options_names_props import ImageOptionsNamesProps
from ....kind.format_kind import FormatKind


class Names(FrameNames):
    """Image Options Names"""

    # region Init
    def __init__(
        self,
        *,
        name: str | None = None,
        desc: str | None = None,
        alt: str | None = None,
        prev: str | None = None,
        next: str | None = None,
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): Specifies name.
            desc (str, optional): Specifies description.
            alt (str, optional): Specifies alternative text.
            prev (str, optional): Specifies previous link.
            next (str, optional): Specifies next link.
        """
        # TODO: Implement prev and next on Frame options Names class.
        # see FrameNames base class.
        super().__init__(name=name, desc=desc, prev=prev, next=next)
        self.prop_alt = alt

    # endregion Init

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
            )
        return self._supported_services_values

    # endregion Overrides

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.IMAGE
        return self._format_kind_prop

    @property
    def prop_alt(self) -> SystemError | None:
        """Gets/Sets Alternative text"""
        return self._get(self._props.alt)

    @prop_alt.setter
    def prop_alt(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.alt)
            return
        self._set(self._props.alt, value)

    @property
    def _props(self) -> ImageOptionsNamesProps:
        try:
            return self._props_frame_opts_protect
        except AttributeError:
            self._props_frame_opts_protect = ImageOptionsNamesProps(
                name="Name",
                desc="Description",
                prev="",  # ChainPrevName not working
                next="",  # ChainNextName not working
                alt="Title",
            )
        return self._props_frame_opts_protect

    # endregion Properties
