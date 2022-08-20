.. _ch09:

**********************************
Chapter 9. Text Search and Replace
**********************************

.. topic:: Overview

    Finding the First Matching Phrase; Replacing all the Matching Words; Finding all Matching Phrases


The GenericTextDocument_ service supports the XSearchable_ and XReplaceable_ interfaces (see Chapter 5, :numref:`ch05fig_txt_doc_serv_interfaces`),
which are the entry points for doing regular expression based search and replace inside a document.

``XSearchable.createSearchDescriptor()`` builds a search description (an ordinary string or a regular expression).
The search is executed with ``XSearchable.findAll()`` or ``findFirst()`` and ``findNext()``.

XReplaceable_ works in a similar way but with a replace descriptor which combines a replacement string with the search string.
``XReplaceable.replaceAll()`` performs search and replacement, but the XSearchable_ searching methods are available as well.
This is shown in :numref:`ch09fig_search_replace_interfaces`.

.. cssclass:: diagram invert

    .. _ch09fig_search_replace_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/185747397-9ce45c81-8283-4ff3-8b9c-03fc0b0a36c4.png
        :alt: Diagram of The X Searchable and X Replaceable_Interfaces
        :figclass: align-center

        :The XSearchable_ and XReplaceable_ Interfaces.

The following code fragment utilizes the XSearchable_ and XSearchDescriptor_ interfaces:

.. tabs::

    .. code-tab:: python

        searchable = Lo.qi(XSearchable, doc)
        srch_desc = searchable.createSearchDescriptor()
        srch_desc.setSearchString("colou?r")

XReplaceable_ and XReplaceDescriptor_ objects are configured in a similar way, as shown in the examples.

XSearchDescriptor_ and XReplaceDescriptor_ contain get and set methods for their strings.
But a lot of the search functionality is expressed as properties in their SearchDescriptor_ and ReplaceDescriptor_ services.
:numref:`ch09fig_search_replace_services` summarizes these arrangements.

.. cssclass:: diagram invert

    .. _ch09fig_search_replace_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/185748858-71549f8d-9c54-47d3-acb0-c1a890966417.png
        :alt: Diagram of The Search Descriptor and Replace Descriptor Services.
        :figclass: align-center

        :The SearchDescriptor_ and ReplaceDescriptor_ Services.

The following code fragment utilizes the XSearchable_ and XSearchDescriptor_ interfaces:

The next code fragment accesses the SearchDescriptor_ properties, and switches on regular expression searching:

.. tabs::

    .. code-tab:: python

        srch_props = Lo.qi(XPropertySet, srch_desc, raise_err=True)
        srch_props.setPropertyValue("SearchRegularExpression", True)

Alternatively, :py:meth:`.Props.set_property` can be employed:

.. tabs::

    .. code-tab:: python

        Props.set_property(srch_desc, "SearchRegularExpression", True)

Once a search descriptor has been created (i.e. its string is set and any properties configured), then one of the ``findXXX()`` methods in XSearchable_ can be called.

For instance, ``XSearchable.findFirst()`` returns the text range of the first matching element (or ``None``), as in:

.. tabs::

    .. code-tab:: python

        srch = searchable.findFirst(srch_desc)

        if srch is not None:
            match_tr = Lo.qi(XTextRange, srch)

The example programs, |text_replace|_ and |italics_styler|_, demonstrate search and replacement.
|text_replace|_ uses XSearchable_ to find the first occurrence of a regular expression and XReplaceable_ to replace multiple occurrences of other words.

|italics_styler|_ calls XSearchable_'s ``findAll()`` to find every occurrence of a phrase.

9.1 Finding the First Matching Phrase
=====================================

|text_replace|_ repeatedly calls ``XSearchable.findFirst()`` with regular expressions taken from a tuple.
The first matching phrase for each expression is reported. For instance, the call:

.. tabs::

    .. code-tab:: python

        words = ("(G|g)rit", "colou?r",)
        find_words(doc, words)

prints the following when ``bigStory.doc`` is searched:

.. code-block:: text

    Searching for fist occurrence of '(G|g)rit'
    - found 'Grit'
        - on page 1
        - at char postion: 8
    Searching for fist occurrence of 'colou?r'
    - found 'colour'
        - on page 5
        - at char postion: 12

Three pieces of information are printed for each match: the text that matched, its page location, and its character position calculated from the start of the document.
The character position could be obtained from a text cursor or a text view cursor, but a page cursor is needed to access the page number.
Therefore the easiest thing to use a text view cursor, and a linked page cursor.

The code for ``find_words()``:

.. tabs::

    .. code-tab:: python

        def find_words(doc: XTextDocument, words: Sequence[str]) -> None:
            # get the view cursor and link the page cursor to it
            tvc = Write.get_view_cursor(doc)
            tvc.gotoStart(False)
            page_cursor = Write.get_page_cursor(tvc)
            searchable = Lo.qi(XSearchable, doc)
            srch_desc = searchable.createSearchDescriptor()

            for word in words:
                print(f"Searching for fist occurrence of '{word}'")
                srch_desc.setSearchString(word)

                srch_props = Lo.qi(XPropertySet, srch_desc, raise_err=True)
                srch_props.setPropertyValue("SearchRegularExpression", True)

                srch = searchable.findFirst(srch_desc)

                if srch is not None:
                    match_tr = Lo.qi(XTextRange, srch)

                    tvc.gotoRange(match_tr, False)
                    print(f"  - found '{match_tr.getString()}'")
                    print(f"    - on page {page_cursor.getPage()}")
                    # tvc.gotoStart(True)
                    tvc.goRight(len(match_tr.getString()), True)
                    print(f"    - at char postion: {len(tvc.getString())}")
                    Lo.delay(500)

``find_words()`` get the text view cursor (``tvc``) from :py:meth:`.Write.get_view_cursor`.

.. tabs::

    .. code-tab:: python

        page_cursor = Write.get_page_cursor(tvc)

Alternatively ``page_curser`` could be cast from view cursor:

.. tabs::

    .. code-tab:: python

        page_cursor = Lo.qi(XPageCursor, tvc)

``find_words()`` creates the text view cursor (``tvc``), moves it to the start of the document, and links the page cursor to it.

There is only one view cursor in an application, so when the text view cursor moves, so does the page cursor, and vice versa.

The XSearchable_ and XSearchDescriptor_ interfaces are instantiated, and a for-loop searches for each word in the supplied array.
If ``XSearchable.findFirst()`` returns a matching text range, it's used by ``XTextCursor.gotoRange()`` to update the position of the cursor.

After the page position has been printed, the cursor is moved to the right by the length of the current match string.

.. tabs::

    .. code-tab:: python

        tvc.goRight(len(match_tr.getString()), True)

9.2 Replacing all the Matching Words
====================================

|text_replace|_ also contains a method called ``replace_words()``, which takes two string sequences as arguments:

.. tabs::

    .. code-tab:: python

        uk_words = ("colour", "neighbour", "centre", "behaviour", "metre", "through")
        us_words = ("color", "neighbor", "center", "behavior", "meter", "thru")

``replace_words()`` cycles through the sequences, replacing all occurrences of the words in the first sequence (:abbreviation:`ex:` in ``uk_words``)
with the corresponding words in the second sequence (:abbreviation:`ex:` in ``us_words``). For instance, every occurrence of ``colour`` is replaced by ``color``.


.. code-block:: text

    Change all occurrences of ...

      colour -> color
        - no. of changes: 1
      neighbour -> neighbor
        - no. of changes: 2
      centre -> center
        - no. of changes: 2
      behaviour -> behavior
        - no. of changes: 0
      metre -> meter
        - no. of changes: 0
      through -> thru
        - no. of changes: 4

Since ``replace_words()`` doesn't report page and character positions, its code is somewhat shorter than ``find_words()``:

.. tabs::

    .. code-tab:: python

        def replace_words(doc: XTextDocument, old_words: Sequence[str], new_words: Sequence[str]) -> int:
            replaceable = Lo.qi(XReplaceable, doc, raise_err=True)
            replace_desc = Lo.qi(XReplaceDescriptor, replaceable.createSearchDescriptor())

            for old, new in zip(old_words, new_words):
                replace_desc.setSearchString(old)
                replace_desc.setReplaceString(new)
            return replaceable.replaceAll(replace_desc)

The XReplaceable_ and XReplaceDescriptor_ interfaces are created in a similar way to their search versions.
The replace descriptor has two set methods, one for the search string, the other for the replacement string.

9.3 Finding all Matching Phrases
================================

The |italics_styler|_ example also outputs matching details:

.. code-block:: shell

    python start.py --show --file "cicero_dummy.odt" --word pleasure green --word pain red

The program opens the file and uses the "search all' method in XSearchable_ to find all occurrences of the string in the document.
The matching strings are italicized and colored, and the changed document saved as "italicized.doc".
These changes are not performed using XReplaceable_ methods.

:numref:`ch09fig_italicize_doc_ss` shows a fragment of the resulting document, with the "pleasure" and "pain" changed in the text.
The search ignores case.

.. cssclass:: screen_shot invert

    .. _ch09fig_italicize_doc_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/185763894-adb25e29-270f-4085-834b-502cf48c86fe.png
        :alt: Screen shot of A Fragment of The Italicized Document
        :figclass: align-center

        :A Fragment of The Italicized Document.

The |italics_styler|_ example also outputs matching details (partial output):

.. code-block:: text

    No. of matches: 17
      - found: 'pleasure'
        - on page 1
        - starting at char position: 85
      - found: 'pleasure'
        - on page 1
        - starting at char position: 319
      - found: 'pleasure'
        - on page 1
        - starting at char position: 350
      - found: 'pleasure'
        - on page 1
        - starting at char position: 408
      :
    Found 17 results for "pleasure"
    Searching for all occurrences of 'pain'
    No. of matches: 15
      - found: 'pain'
        - on page 1
        - starting at char position: 107
      - found: 'pain'
        - on page 1
        - starting at char position: 548
      - found: 'pain'
        - on page 1
        - starting at char position: 578
      - found: 'pain'
        - on page 1
        - starting at char position: 647
        :
    Found 15 results for "pain"

As with |text_replace|_, the printed details include the page and character positions of the matches.

The searching in |italics_styler|_ is performed by ``italicize_all()``, which bears a close resemblance to ``find_words()``:

.. tabs::

    .. code-tab:: python

        def italicize_all(doc: XTextDocument, phrase: str, color: Color) -> int:
            # cursor = Write.get_view_cursor(doc) # can be used when visible
            cursor = Write.get_cursor(doc)
            cursor.gotoStart(False)
            page_cursor = Write.get_page_cursor(doc)
            result = 0
            try:
                xsearchable = Lo.qi(XSearchable, doc, True)
                srch_desc = xsearchable.createSearchDescriptor()
                print(f"Searching for all occurrences of '{phrase}'")
                pharse_len = len(phrase)
                srch_desc.setSearchString(phrase)
                Props.set_property(obj=srch_desc, name="SearchCaseSensitive", value=False)
                Props.set_property(
                    obj=srch_desc, name="SearchWords", value=True
                )  # If TRUE, only complete words will be found.

                matches = xsearchable.findAll(srch_desc)
                result = matches.getCount()

                print(f"No. of matches: {result}")

                for i in range(result):
                    match_tr = Lo.qi(XTextRange, matches.getByIndex(i))
                    if match_tr is not None:
                        cursor.gotoRange(match_tr, False)
                        print(f"  - found: '{match_tr.getString()}'")
                        print(f"    - on page {page_cursor.getPage()}")
                        cursor.gotoStart(True)
                        print(f"    - starting at char position: {len(cursor.getString()) - pharse_len}")

                        Props.set_properties(obj=match_tr, names=("CharColor", "CharPosture"), vals=(color, FontSlant.ITALIC))

            except Exception as e:
                raise
            return result

After the search descriptor string has been defined, the ``SearchCaseSensitive`` property in SearchDescriptor_ is set to ``False``:

.. tabs::

    .. code-tab:: python

        srch_desc.setSearchString(phrase)
        Props.set_property(obj=srch_desc, name="SearchCaseSensitive", value=False)

This allows the search to match text contains both upper and lower case letters, such as "Pleasure".
Many other search variants, such as restricting the search to complete words,
and the use of search similarity parameters are described in the SearchDescriptor_ documentation (``lodoc SearchDescriptor service``).

``XSearchable.findAll()`` returns an XIndexAccess_ collection, which is examined element-by-element inside a for-loop.
The text range for each element is obtained by applying :py:meth:`.Lo.qi`:

.. tabs::

    .. code-tab:: python

        match_tr = Lo.qi(XTextRange, matches.getByIndex(i))

The reporting of the matching page and character position use text view and page cursors in the same way as ``find_words()`` in |text_replace|_.

XTextRange_ is part of the TextRange_ service, which inherits ``ParagraphProperties`` and ``CharacterProperties``.
These properties are changed to adjust the character color and style of the selected range:

.. tabs::

    .. code-tab:: python

        Props.set_properties(
            obj=match_tr,
            names=("CharColor", "CharPosture"),
            vals=(color, FontSlant.ITALIC)
            )

This changes the ``CharColor`` and ``CharPosture`` properties are set to specified color and set to italic.

The color passed into command line can be a integer color such as ``16711680`` or any color name (case in-sensitive) in :py:class:`~.color.CommonColor`.

.. |text_replace| replace:: Text Replace
.. _text_replace: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_text_replace

.. |italics_styler| replace:: Italics Styler
.. _italics_styler: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_italics_styler

.. _GenericTextDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1GenericTextDocument.html
.. _ReplaceDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1ReplaceDescriptor.html
.. _SearchDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html
.. _TextRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextRange.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XReplaceable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XReplaceable.html
.. _XReplaceDescriptor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XReplaceDescriptor.html
.. _XSearchable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchable.html
.. _XSearchDescriptor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchDescriptor.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html