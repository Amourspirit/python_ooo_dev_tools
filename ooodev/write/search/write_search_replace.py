from __future__ import annotations
from typing import cast, Sequence, TYPE_CHECKING
import uno
from com.sun.star.uno import XInterface
from com.sun.star.text import XTextRange
from ooodev.adapter.util.property_replace_comp import PropertyReplaceComp
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.write_text_range import WriteTextRange
from ooodev.write.write_text_ranges import WriteTextRanges
from ooodev.loader import lo as mLo
from ooodev.write.write_doc import WriteDoc

if TYPE_CHECKING:
    from com.sun.star.util import XSearchable
    from com.sun.star.util import XReplaceable
    from com.sun.star.util import XReplaceDescriptor
    from com.sun.star.util import XPropertyReplace
    from ooodev.proto.component_proto import ComponentT

# https://wiki.documentfoundation.org/Documentation/DevGuide/Text_Documents#Search_and_Replace


class WriteSearchReplace(WriteDocPropPartial, PropertyReplaceComp, LoInstPropsPartial):
    """
    Represents a writer search.

    .. versionadded:: 0.30.0
    """

    def __init__(self, doc: WriteDoc, desc: XPropertyReplace) -> None:
        """
        Constructor

        Args:
            doc (WriteDoc): doc that owns this component.
        """
        WriteDocPropPartial.__init__(self, obj=doc)
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        PropertyReplaceComp.__init__(self, component=desc)  # type: ignore

    def find_first(self) -> WriteTextRange[WriteSearchReplace] | None:
        """
        Finds the first occurrence of the search string.

        Returns:
            XInterface | None: XInterface or None.
        """

        searchable = cast("XSearchable", self.write_doc.component)
        # result may be a text cursor but can be cast to XTextRange
        result = mLo.Lo.qi(XTextRange, searchable.findFirst(self.component))
        return None if result is None else WriteTextRange(owner=self, component=result, lo_inst=self.lo_inst)  # type: ignore

    def find_next(self, start: XInterface | ComponentT) -> WriteTextRange[WriteSearchReplace] | None:
        """
        Finds the first occurrence of the search string.

        Args:
            start (XInterface | ComponentT): Start position.
                Any object that supports ``XInterface`` or an object that has a Component that supports ``XInterface``.

        Returns:
            XInterface | None: XInterface or None.
        """
        if mLo.Lo.qi(XInterface, start) is None:
            start_component = cast(XInterface, start.component)  # type: ignore
        else:
            start_component = cast(XInterface, start)

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

    def replace_words(self, old_words: Sequence[str], new_words: Sequence[str]) -> int:
        """
        Replaces all occurrences of old_words with new_words.

        ``old_words`` and ``new_words`` must be the same length.

        Args:
            old_words (Sequence[str]): Sequence of words to be replaced.
                Can be a sequence of regular expressions if ``search_regular_expression`` is True.
            new_words (Sequence[str]): Sequence of words to replace with.

        Returns:
            int: The number of replacements.
        """
        replaceable = cast("XReplaceable", self.write_doc.component)
        replace_desc = cast("XReplaceDescriptor", self.component)
        count = 0
        for old, new in zip(old_words, new_words):
            replace_desc.setSearchString(old)
            replace_desc.setReplaceString(new)
            count += replaceable.replaceAll(replace_desc)
        return count

    def replace_all(self) -> int:
        """
        Replaces all occurrences of old_words with new_words.

        ``old_words`` and ``new_words`` must be the same length.

        Args:
            old_words (Sequence[str]): Sequence of words to be replaced.
                Can be a sequence of regular expressions if ``search_regular_expression`` is True.
            new_words (Sequence[str]): Sequence of words to replace with.

        Returns:
            int: The number of replacements.
        """
        replaceable = cast("XReplaceable", self.write_doc.component)
        replace_desc = cast("XReplaceDescriptor", self.component)
        return replaceable.replaceAll(replace_desc)

    @property
    def search_str(self) -> str:
        """
        Gets/Set search string.

        Returns:
            str: Search string.
        """
        return self.get_search_string()

    @search_str.setter
    def search_str(self, value: str) -> None:
        self.set_search_string(value)

    @property
    def replace_str(self) -> str:
        """
        Gets/Set replace string.

        Returns:
            str: Replace string.
        """
        return self.get_replace_string()

    @replace_str.setter
    def replace_str(self, value: str) -> None:
        self.set_replace_string(value)
