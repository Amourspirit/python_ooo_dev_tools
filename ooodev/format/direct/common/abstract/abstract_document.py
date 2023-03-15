"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

import uno
from com.sun.star.text import XTextDocument
from com.sun.star.container import XNameAccess

from .....exceptions import ex as mEx
from .....utils import info as mInfo
from .....utils import lo as mLo
from .....utils.data_type.size import Size
from .....office import write as mWrite
from ....style_base import StyleBase
from ..props.border_props import BorderProps as BorderProps


class AbstractDocument(StyleBase):
    """
    Abstract Document.

    Contians method related to Document Specific methods

    .. versionadded:: 0.9.0
    """

    # region methods
    def _get_doc_family_style_name(self) -> str:
        try:
            return self._doc_family_style_name
        except AttributeError:
            self._doc_family_style_name = "PageStyles"
        return self._doc_family_style_name

    def _get_doc_standard_name(self) -> str:
        try:
            return self._doc_standard_name
        except AttributeError:
            self._doc_standard_name = "Standard"
        return self._doc_standard_name

    def get_write_doc(self) -> XTextDocument:
        """Gets current Writer active document"""
        return mWrite.Write.active_doc

    def get_page_text_size(self) -> Size:
        """
        Get page text size  in ``1/100 mm`` units.

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            Size: Page text Size in ``1/100 mm`` units.

        .. versionadded:: 0.9.0
        """
        props = mInfo.Info.get_style_props(
            doc=mWrite.Write.active_doc,
            family_style_name=self._get_doc_family_style_name(),
            prop_set_nm=self._get_doc_standard_name(),
        )
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style")
        try:
            width = int(props.getPropertyValue("Width"))
            height = int(props.getPropertyValue("Height"))
            left_margin = int(props.getPropertyValue("LeftMargin"))
            right_margin = int(props.getPropertyValue("RightMargin"))
            top_margin = int(props.getPropertyValue("TopMargin"))
            btm_margin = int(props.getPropertyValue("BottomMargin"))
            text_width = width - (left_margin + right_margin)
            text_height = height - (top_margin + btm_margin)

            return Size(text_width, text_height)
        except Exception as e:
            raise Exception("Could not access standard page style dimensions") from e

    def get_page_size(self) -> Size:
        """
        Get page size  in ``1/100 mm`` units.

        Raises:
            PropertiesError: If unable to access properties
            Exception: If unable to get page size

        Returns:
            Size: Page Size in ``1/100 mm`` units.
        """
        props = mInfo.Info.get_style_props(
            doc=mWrite.Write.active_doc,
            family_style_name=self._get_doc_family_style_name(),
            prop_set_nm=self._get_doc_standard_name(),
        )
        if props is None:
            raise mEx.PropertiesError("Could not access the standard page style")
        try:
            width = int(props.getPropertyValue("Width"))
            height = int(props.getPropertyValue("Height"))
            return Size(width, height)
        except Exception as e:
            raise Exception("Could not access standard page style dimensions") from e

    def get_text_frames(self) -> XNameAccess | None:
        """
        Gets document Text Frames.

        Args:
            doc (XComponent): Document

        Raises:
            MissingInterfaceError: if doc does not implement ``XTextFramesSupplier`` interface

        Returns:
            XNameAccess | None: Text Frames on success, Otherwise, None
        """
        # return mLo.Lo.this_component.TextFrames
        return mWrite.Write.get_text_frames(mWrite.Write.active_doc)

    # endregion methods
