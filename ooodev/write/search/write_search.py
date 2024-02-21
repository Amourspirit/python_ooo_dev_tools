from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.uno import XInterface
from com.sun.star.text import XTextRange
from ooodev.adapter.util.search_descriptor_comp import SearchDescriptorComp
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.write_text_range import WriteTextRange
from ooodev.write.write_text_ranges import WriteTextRanges
from ooodev.loader import lo as mLo
from ..write_doc import WriteDoc

if TYPE_CHECKING:
    from com.sun.star.util import XSearchable
    from ooodev.proto.component_proto import ComponentT
    from ooodev.adapter.container.index_access_comp import IndexAccessComp


class WriteSearch(WriteDocPropPartial, SearchDescriptorComp, LoInstPropsPartial):
    """
    Represents a writer search.

    .. versionadded:: 0.30.0
    """

    def __init__(self, doc: WriteDoc) -> None:
        """
        Constructor

        Args:
            doc (WriteDoc): doc that owns this component.
        """
        WriteDocPropPartial.__init__(self, obj=doc)
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        SearchDescriptorComp.__init__(self, component=self.write_doc.component.createSearchDescriptor())  # type: ignore

    def find_first(self) -> WriteTextRange[WriteSearch] | None:
        """
        Finds the first occurrence of the search string.

        Returns:
            XInterface | None: XInterface or None.
        """

        searchable = cast("XSearchable", self.write_doc.component)
        # result may be a text cursor but can be cast to XTextRange
        result = mLo.Lo.qi(XTextRange, searchable.findFirst(self.component))
        return None if result is None else WriteTextRange(owner=self, component=result, lo_inst=self.lo_inst)  # type: ignore

    def find_next(self, start: XInterface | ComponentT) -> WriteTextRange[WriteSearch] | None:
        """
        Finds the first occurrence of the search string.

        Args:
            start (XInterface | ComponentT): Start position.
                Any object that supports ``XInterface`` or an object that has a Component that supports ``XInterface``.

        Returns:
            XInterface | None: XInterface or None.
        """
        if mLo.Lo.qi(XInterface, start) is None:
            start_component = cast("XInterface", start.component)  # type: ignore
        else:
            start_component = cast("XInterface", start)

        searchable = cast("XSearchable", self.write_doc.component)
        # result may be a text cursor but can be cast to XTextRange
        result = mLo.Lo.qi(XTextRange, searchable.findNext(start_component, self.component))
        return None if result is None else WriteTextRange(owner=self, component=result, lo_inst=self.lo_inst)  # type: ignore

    def find_all(self) -> WriteTextRanges | None:
        """
        Searches the contained texts for all occurrences of whatever is specified.

        Returns:
            WriteTextRanges | None: The found occurrences.
        """
        searchable = cast("XSearchable", self.write_doc.component)
        result = searchable.findAll(self.component)
        return None if result is None else WriteTextRanges(owner=self, component=result)
