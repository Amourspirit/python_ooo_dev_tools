from __future__ import annotations
from typing import Sequence, overload, TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from .write_doc import WriteDoc
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.type_var import PathOrStr, Table
    from ooodev.units import UnitT

from ooodev.utils.color import Color, CommonColor
from ooodev.office import write as mWrite
from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils import lo as mLo
from ooodev.adapter.drawing.graphic_object_shape_comp import GraphicObjectShapeComp

from . import write_text_content as mWriteTextContent
from . import write_text_table as mWriteTextTable

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange


class WriteTextCursor(TextCursorComp, PropPartial, QiPartial):
    """Represents a writer text cursor."""

    def __init__(self, owner: WriteDoc, component: XTextRange) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Sheet that owns this cell range.
            col_obj (Any): Range object.
        """
        self.__owner = owner
        TextCursorComp.__init__(self, component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def add_bookmark(self, name: str) -> bool:
        """
        Adds a bookmark with the specified name to the cursor.

        Args:
            name (str): Bookmark name

        Returns:
            bool: True if bookmark is added; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.BOOKMARK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.BOOKMARK_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``name`` and ``cursor``.
        """
        return mWrite.Write.add_bookmark(self.component, name)

    # region add_formula()
    @overload
    def add_formula(self, formula: str) -> mWriteTextContent.WriteTextContent:
        """
        Adds a formula

        Args:
            formula (str): formula

        Returns:
            WriteTextContent: Embedded Object.
        """
        ...

    @overload
    def add_formula(self, formula: str, styles: Sequence[StyleT]) -> mWriteTextContent.WriteTextContent:
        """
        Adds a formula

        Args:
            formula (str): formula
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Returns:
            WriteTextContent: Embedded Object.
        """
        ...

    def add_formula(self, formula: str, styles: Sequence[StyleT] | None = None) -> mWriteTextContent.WriteTextContent:
        """
        Adds a formula

        Args:
            formula (str): formula
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Raises:
            CreateInstanceMsfError: If unable to create text.TextEmbeddedObject
            CancelEventError: If event ``WriteNamedEvent.FORMULA_ADDING`` is cancelled
            Exception: If unable to add formula

        Returns:
            WriteTextContent: Embedded Object.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.FORMULA_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.FORMULA_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.obj </src/format/ooodev.format.writer.direct.obj>`

        Note:
            Event args ``event_data`` is a dictionary containing ``formula`` and ``cursor``.
        """
        if styles:
            return mWriteTextContent.WriteTextContent(self, mWrite.Write.add_formula(self.component, formula, styles))
        else:
            return mWriteTextContent.WriteTextContent(self, mWrite.Write.add_formula(self.component, formula))

    # endregion add_formula()

    def add_hyperlink(self, label: str, url_str: str) -> bool:
        """
        Add a hyperlink

        Args:
            label (str): Hyperlink label
            url_str (str): Hyperlink url

        Raises:
            CreateInstanceMsfError: If unable to create TextField.URL instance
            Exception: If unable to create hyperlink

        Returns:
            bool: True if hyperlink is added; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.HYPER_LINK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.HYPER_LINK_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``label``, ``url_str`` and ``cursor``.
        """
        return mWrite.Write.add_hyperlink(self.component, label, url_str)

    # region add_image_link
    @overload
    def add_image_link(self, fnm: PathOrStr) -> mWriteTextContent.WriteTextContent:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    # style: Sequence[StyleT] = None
    @overload
    def add_image_link(
        self, fnm: PathOrStr, *, width: int | UnitT, height: int | UnitT
    ) -> mWriteTextContent.WriteTextContent:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """

    @overload
    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        styles: Sequence[StyleT],
    ) -> mWriteTextContent.WriteTextContent:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    @overload
    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        width: int | UnitT,
        height: int | UnitT,
        styles: Sequence[StyleT],
    ) -> mWriteTextContent.WriteTextContent:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        width: int | UnitT = 0,
        height: int | UnitT = 0,
        styles: Sequence[StyleT] | None = None,
    ) -> mWriteTextContent.WriteTextContent:
        """
        Add Image Link

        Args:
            fnm (PathOrStr): Image path
            width (int, UnitT): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Raises:
            CreateInstanceMsfError: If Unable to create text.TextGraphicObject.
            MissingInterfaceError: If unable to obtain XPropertySet interface.
            Exception: If unable to add image.
            CancelEventError: If ``IMAGE_LINK_ADDING`` event is canceled.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_LINK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_LINK_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.image </src/format/ooodev.format.writer.direct.image>`

        Note:
            Event args ``event_data`` is a dictionary containing ``doc``, ``cursor``, ``fnm``, ``width`` and ``height``.
        """
        if styles is None:
            styles = ()
        result = mWrite.Write.add_image_link(
            doc=self.write_doc.component,
            cursor=self.component,
            fnm=fnm,
            width=width,
            height=height,
            styles=styles,
        )
        return mWriteTextContent.WriteTextContent(self, result)

    # endregion add_image_link

    # region add_image_shape()
    @overload
    def add_image_shape(self, fnm: PathOrStr) -> GraphicObjectShapeComp:
        """
        Add Image Shape.

        Args:
            fnm (PathOrStr): Image path.

        Returns:
            GraphicObjectShapeComp: Image Shape on success; Otherwise, ``None``.
        """
        ...

    @overload
    def add_image_shape(self, fnm: PathOrStr, width: int | UnitT, height: int | UnitT) -> GraphicObjectShapeComp:
        """
        Add Image Shape.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            GraphicObjectShapeComp: Image Shape on success; Otherwise, ``None``.
        """
        ...

    def add_image_shape(
        self, fnm: PathOrStr, width: int | UnitT = 0, height: int | UnitT = 0
    ) -> GraphicObjectShapeComp:
        """
        Add Image Shape.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.

        Raises:
            CreateInstanceMsfError: If unable to create drawing.GraphicObjectShape.
            ValueError: If unable to get image.
            MissingInterfaceError: If require interface cannot be obtained.
            Exception: If unable to add image shape.
            CancelEventError: if ``IMAGE_SHAPE_ADDING`` event is canceled.

        Returns:
            GraphicObjectShapeComp: Image Shape on success; Otherwise, ``None``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_SHAPE_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_SHAPE_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``doc``, ``cursor``, ``fnm``, ``width`` and ``height``.
        """
        result = mWrite.Write.add_image_shape(cursor=self.component, fnm=fnm, width=width, height=height)
        return GraphicObjectShapeComp(result)

    # endregion add_image_shape()

    def add_line_divider(self, line_width: int) -> None:
        """
        Adds a line divider

        Args:
            line_width (int): Line width

        Raises:
            CreateInstanceMsfError: If unable to create drawing.LineShape instance
            MissingInterfaceError: If unable to obtain XShape interface
            Exception: If unable to add Line divider
        """
        mWrite.Write.add_line_divider(self.component, line_width)

    def add_table(
        self,
        table_data: Table,
        header_bg_color: Color | None = CommonColor.DARK_BLUE,
        header_fg_color: Color | None = CommonColor.WHITE,
        tbl_bg_color: Color | None = CommonColor.LIGHT_BLUE,
        tbl_fg_color: Color | None = CommonColor.BLACK,
        first_row_header: bool = True,
        styles: Sequence[StyleT] | None = None,
    ) -> mWriteTextTable.WriteTextTable:
        """
        Adds a table.

        Each row becomes a row of the table. The first row is treated as a header.

        Args:
            cursor (XTextCursor): Text Cursor.
            table_data (Table): 2D Table with the the first row containing column names.
            header_bg_color (:py:data:`~.utils.color.Color`, optional): Table header background color.
                Set to None to ignore header color. Defaults to ``CommonColor.DARK_BLUE``.
            header_fg_color (:py:data:`~.utils.color.Color`, optional): Table header foreground color.
                Set to None to ignore header color. Defaults to ``CommonColor.WHITE``.
            tbl_bg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``CommonColor.LIGHT_BLUE``.
            tbl_fg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``CommonColor.BLACK``.
            first_row_header (bool, optional): If ``True`` First row is treated as header data. Default ``True``.
            styles (Sequence[StyleT], optional): One or more styles to apply to frame.
                Only styles that support ``com.sun.star.text.TextTable`` service are applied.

        Raises:
            ValueError: If table_data is empty
            CreateInstanceMsfError: If unable to create instance of text.TextTable
            CancelEventError:  If ``WriteNamedEvent.TABLE_ADDING`` event cancelled
            Exception: If unable to add table

        Returns:
            WriteTextTable: Table that is added to document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TABLE_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TABLE_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method args.

        See Also:
            - :ref:`help_writer_format_direct_table`
            - :py:class:`~.utils.color.CommonColor`
            - :py:meth:`~.utils.table_helper.TableHelper.table_2d_to_dict`
            - :py:meth:`~.utils.table_helper.TableHelper.table_dict_to_table`

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.writer.direct.table </src/format/ooodev.format.writer.direct.table>` subpackages.
        """
        result = mWrite.Write.add_table(
            cursor=self.component,
            table_data=table_data,
            header_bg_color=header_bg_color,
            header_fg_color=header_fg_color,
            tbl_bg_color=tbl_bg_color,
            tbl_fg_color=tbl_fg_color,
            first_row_header=first_row_header,
            styles=styles,
        )
        return mWriteTextTable.WriteTextTable(self, result)

    # region Properties
    @property
    def write_doc(self) -> WriteDoc:
        """Doc that owns this Cursor."""
        return self.__owner

    # endregion Properties
