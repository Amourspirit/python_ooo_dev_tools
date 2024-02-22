"""
This class is a partial class for ``com.sun.star.util.Shape.SearchDescriptor`` service.
"""

from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from com.sun.star.util import SearchDescriptor  # service


class SearchDescriptorPartialProps:
    """
    Partial Class for Shape Service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SearchDescriptor) -> None:
        """
        Constructor

        Args:
            component (SearchDescriptor): UNO Component that implements ``com.sun.star.util.SearchDescriptor`` service.
        """
        self.__component = component

    # region Properties

    @property
    def search_backwards(self) -> bool:
        """
        If ``True``, the search is done backwards in the document.
        """
        return self.__component.SearchBackwards

    @search_backwards.setter
    def search_backwards(self, value: bool) -> None:
        self.__component.SearchBackwards = value

    @property
    def search_case_sensitive(self) -> bool:
        """
        If ``True``, the case of the letters is important for the match.
        """
        return self.__component.SearchCaseSensitive

    @search_case_sensitive.setter
    def search_case_sensitive(self, value: bool) -> None:
        self.__component.SearchCaseSensitive = value

    @property
    def search_regular_expression(self) -> bool:
        """
        If ``True``, the search string is evaluated as a regular expression.

        ``search_regular_expression``, ``search_wildcard`` and ``search_similarity`` are mutually exclusive, only one can be ``True`` at the same time.
        """
        return self.__component.SearchRegularExpression

    @search_regular_expression.setter
    def search_regular_expression(self, value: bool) -> None:
        self.__component.SearchRegularExpression = value

    @property
    def search_similarity(self) -> bool:
        """
        Gets/Sets - If ``True``, a **similarity search** is performed.

        In the case of a similarity search, the following properties specify the kind of similarity:

        ``search_regular_expression``, ``search_wildcard`` and ``search_similarity`` are mutually exclusive, only one can be ``True`` at the same time.
        """
        return self.__component.SearchSimilarity

    @search_similarity.setter
    def search_similarity(self, value: bool) -> None:
        self.__component.SearchSimilarity = value

    @property
    def search_similarity_add(self) -> int:
        """
        Gets/Sets the number of characters that must be added to match the search pattern.
        """
        return self.__component.SearchSimilarityAdd

    @search_similarity_add.setter
    def search_similarity_add(self, value: int) -> None:
        self.__component.SearchSimilarityAdd = value

    @property
    def search_similarity_exchange(self) -> int:
        """
        Gets/Sets - This property specifies the number of characters that must be replaced to match the search pattern.
        """
        return self.__component.SearchSimilarityExchange

    @search_similarity_exchange.setter
    def search_similarity_exchange(self, value: int) -> None:
        self.__component.SearchSimilarityExchange = value

    @property
    def search_similarity_relax(self) -> bool:
        """
        Gets/Sets If ``True``, all similarity rules are applied together.

        In the case of a relaxed similarity search, the following properties are applied together:
        """
        return self.__component.SearchSimilarityRelax

    @search_similarity_relax.setter
    def search_similarity_relax(self, value: bool) -> None:
        self.__component.SearchSimilarityRelax = value

    @property
    def search_similarity_remove(self) -> int:
        """
        Gets/Sets - This property specifies the number of characters that may be ignored to match the search pattern.
        """
        return self.__component.SearchSimilarityRemove

    @search_similarity_remove.setter
    def search_similarity_remove(self, value: int) -> None:
        self.__component.SearchSimilarityRemove = value

    @property
    def search_styles(self) -> bool:
        """
        Gets/Sets If ``True``, it is searched for positions where the paragraph style with the name of the search pattern is applied.
        """
        return self.__component.SearchStyles

    @search_styles.setter
    def search_styles(self, value: bool) -> None:
        self.__component.SearchStyles = value

    @property
    def search_wildcard(self) -> bool | None:
        """
        Gets/Sets - If ``True``, the search string is evaluated as a wildcard pattern.

        Wildcards are ``*`` (asterisk) for any sequence of characters, including an empty sequence, and ``?`` (question mark) for exactly one character.
        Escape character is ``\\`` (U+005C REVERSE SOLIDUS) aka backslash, it escapes the special meaning of a question mark, asterisk or escape
        character that follows immediately after the escape character.

        ``search_regular_expression``, ``search_wildcard`` and ``search_similarity`` are mutually exclusive, only one can be ``True`` at the same time.

        **optional**

        Returns:
            bool | None: ``True`` if the search string is evaluated as a wildcard pattern, otherwise ``False``. If the property is not supported, ``None`` is returned.
        """
        with contextlib.suppress(AttributeError):
            return self.__component.SearchWildcard
        return None

    @search_wildcard.setter
    def search_wildcard(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.SearchWildcard = value

    @property
    def search_words(self) -> bool:
        """
        If ``True``, only complete words will be found.
        """
        return self.__component.SearchWords

    @search_words.setter
    def search_words(self, value: bool) -> None:
        self.__component.SearchWords = value

    # endregion Properties
