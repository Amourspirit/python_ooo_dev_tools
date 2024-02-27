from __future__ import annotations
from ooodev.theme.theme import ThemeBase


class ThemeGeneral(ThemeBase):
    """
    Theme General Properties.

    The properties are populated from LibreOffice theme colors.

    Automatic color values are returned with a value of ``-1``.
    All other values are positive numbers.
    """

    # region Properties
    @property
    def background_color(self) -> int:
        """Background color."""
        try:
            return self._background_color
        except AttributeError:
            self._background_color = self._get_color("AppBackground")
        return self._background_color

    @property
    def font_color(self) -> int:
        """Font color."""
        try:
            return self._font_color
        except AttributeError:
            self._font_color = self._get_color("FontColor")
        return self._font_color

    @property
    def links_color(self) -> int:
        """Links color."""
        try:
            return self._links_color
        except AttributeError:
            self._links_color = self._get_color("Links")
        return self._links_color

    @property
    def links_visible(self) -> bool:
        """Links visible."""
        try:
            return self._links_visible
        except AttributeError:
            self._links_visible = self._get_visible("Links")
        return self._links_visible

    @property
    def links_visited_color(self) -> int:
        """Links visited color."""
        try:
            return self._links_visited_color
        except AttributeError:
            self._links_visited_color = self._get_color("LinksVisited")
        return self._links_visited_color

    @property
    def links_visited_visible(self) -> bool:
        """Links visited visible."""
        try:
            return self._links_visited_visible
        except AttributeError:
            self._links_visited_visible = self._get_visible("LinksVisited")
        return self._links_visited_visible

    @property
    def object_boundaries_color(self) -> int:
        """Object boundaries color."""
        try:
            return self._object_boundaries_color
        except AttributeError:
            self._object_boundaries_color = self._get_color("ObjectBoundaries")
        return self._object_boundaries_color

    @property
    def object_boundaries_visible(self) -> bool:
        """Object boundaries visible."""
        try:
            return self._object_boundaries_visible
        except AttributeError:
            self._object_boundaries_visible = self._get_visible("ObjectBoundaries")
        return self._object_boundaries_visible

    @property
    def shadow_color(self) -> int:
        """Shadow color."""
        try:
            return self._shadow_color
        except AttributeError:
            self._shadow_color = self._get_color("Shadow")
        return self._shadow_color

    @property
    def shadow_visible(self) -> bool:
        """Shadow visible."""
        try:
            return self._shadow_visible
        except AttributeError:
            self._shadow_visible = self._get_visible("Shadow")
        return self._shadow_visible

    @property
    def smart_tags_color(self) -> int:
        """Smart Tags color."""
        try:
            return self._smart_tags_color
        except AttributeError:
            self._smart_tags_color = self._get_color("SmartTags")
        return self._smart_tags_color

    @property
    def spell_color(self) -> int:
        """Spelling mistakes color."""
        try:
            return self._spell_color
        except AttributeError:
            self._spell_color = self._get_color("Spell")
        return self._spell_color

    @property
    def table_boundaries_color(self) -> int:
        """Table Boundaries color."""
        try:
            return self._table_boundaries_color
        except AttributeError:
            self._table_boundaries_color = self._get_color("TableBoundaries")
        return self._table_boundaries_color

    @property
    def table_boundaries_visible(self) -> bool:
        """Table Boundaries visible."""
        try:
            return self._table_boundaries_visible
        except AttributeError:
            self._table_boundaries_visible = self._get_visible("TableBoundaries")
        return self._table_boundaries_visible

    # endregion Properties
