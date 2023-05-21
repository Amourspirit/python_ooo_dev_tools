.. _ch26:

******************************
Chapter 26. Search and Replace
******************************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

.. topic:: Overview

    XSearchable_; XReplaceable_; SearchDescriptor_; ReplaceDescriptor_; Searching Iteratively; Searching For All Matches; Replacing All Matches

    Examples: |g_secrets|_ and |replace_all|_

The ``_increase_garlic_cost()`` method in |g_secrets_py|_ (:ref:`ch23_increase_garlic_cost`) illustrates how to use loops and if - tests to search for and replace data.
Another approach is to employ the XSearchable_ and XReplaceable_ interfaces, as in the |replace_all_py|_ example shown below:

.. tabs::

    .. code-tab:: python

        # in replace_all.py

        from ooodev.format import Styler
        from ooodev.format.calc.direct.cell.background import Color as BgColor
        from ooodev.format.calc.direct.cell.borders import Borders, BorderLineKind, Side, LineSize
        from ooodev.format.calc.direct.cell.font import Font
        # ... other imports

        class ReplaceAll:
            # ReplaceAll Globals
            ANIMALS = (
                "ass", "cat", "cow", "cub", "doe", "dog", "elk", 
                "ewe", "fox", "gnu", "hog", "kid", "kit", "man",
                "orc", "pig", "pup", "ram", "rat", "roe", "sow", "yak"
            )
            TOTAL_ROWS = 15
            TOTAL_COLS = 6

            def main(self) -> None:
                loader = Lo.load_office(Lo.ConnectSocket())

                try:
                    doc = Calc.create_doc(loader)

                    GUI.set_visible(is_visible=True, odoc=doc)

                    sheet = Calc.get_sheet(doc=doc, index=0)

                    def cb(row: int, col: int, prev) -> str:
                        # call back function for make_2d_array, sets the value for the cell
                        # return animals repeating until all cells are filled
                        v = (row * ReplaceAll.TOTAL_COLS) + col

                        a_len = len(ReplaceAll.ANIMALS)
                        if v > a_len - 1:
                            i = v % a_len
                        else:
                            i = v
                        return ReplaceAll.ANIMALS[i]

                    tbl = TableHelper.make_2d_array(
                        num_rows=ReplaceAll.TOTAL_ROWS, num_cols=ReplaceAll.TOTAL_COLS, val=cb
                    )

                    # create styles that can be applied to the cells via Calc.set_array_range().
                    inner_side = Side()
                    outter_side = Side(width=LineSize.THICK)
                    bdr = Borders(
                        border_side=outter_side, vertical=inner_side, horizontal=inner_side
                    )
                    bg_color = BgColor(StandardColor.BLUE)
                    ft = Font(color=StandardColor.WHITE)

                    Calc.set_array_range(
                        sheet=sheet, range_name="A1:F15", values=tbl, styles=[bdr, bg_color, ft]
                    )

                    # A1:F15
                    cell_rng = Calc.get_cell_range(
                        sheet=sheet, col_start=0, row_start=0, col_end=5, row_end=15
                    )

                    for s in self._srch_strs:
                        if self._is_search_all:
                            self._search_all(sheet=sheet, cell_rng=cell_rng, srch_str=s)
                        else:
                            self._search_iter(sheet=sheet, cell_rng=cell_rng, srch_str=s)

                    if self._repl_str is not None:
                        for s in self._srch_strs:
                            self._replace_all(
                                cell_rng=cell_rng, srch_str=s, repl_str=self._repl_str
                            )

                    if self._out_fnm:
                        Lo.save_doc(doc=doc, fnm=self._out_fnm)

                    msg_result = MsgBox.msgbox(
                        "Do you wish to close document?",
                        "All done",
                        boxtype=MessageBoxType.QUERYBOX,
                        buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                    )
                    if msg_result == MessageBoxResultsEnum.YES:
                        Lo.close_doc(doc=doc, deliver_ownership=True)
                        Lo.close_office()
                    else:
                        print("Keeping document open")

                except Exception:
                    Lo.close_office()
                    raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A blank sheet is filled with a ``15 x 6`` grid of animal names, such as the one shown in :numref:`ch26fig_animials_sheet_grid`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch26fig_animials_sheet_grid:
    .. figure:: https://user-images.githubusercontent.com/4193389/205418740-58e4d6cd-8363-4264-aea9-578f66cad3a5.png
        :alt: A Grid Of Animals for Searching and Replacing.
        :figclass: align-center

        :A Grid Of Animals for Searching and Replacing.

The SheetCellRange_ supports the XReplaceable_ interface, which is a subclass of XSearchable_, as in :numref:`ch26fig_xreplaceable_xsearchable_interfaces`.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch26fig_xreplaceable_xsearchable_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/205418937-cb1d4473-3b4f-4dc8-991b-930be732541d.png
        :alt: The XReplaceable and XSearchable Interfaces.
        :figclass: align-center

        :The XReplaceable_ and XSearchable_ Interfaces.

A cell range's XSearchable_ interface is accessed through casting:

.. tabs::

    .. code-tab:: python

        # in replace_all.py
        srch = Lo.qi(XSearchable, cell_rng, True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The XReplaceable_ interface for the range is obtained in the same way:

.. tabs::

    .. code-tab:: python

        # in replace_all.py
        repl = Lo.qi(XReplaceable, cell_rng, True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

XSearchable_ offers iterative searching using its ``findFirst()`` and ``findNext()`` methods, which is demonstrated shortly in the ``_search_iter()`` method in |replace_all_py|_.
XSearchable_ can also search for all matches at once with :py:meth:`.Calc.find_all`, which is employ in the |replace_all_py|_ ``_search_all()``.
Only one of these methods is needed by the program, so the other is commented out in the ``main()`` function shown above.

XReplaceable_ only offers ``replaceAll()`` which searches for and replaces all of its matches in a single call.
It's utilized by the |replace_all_py|_ ``_replace_all()`` method.

Before a search can begin, it's usually necessary to tweak the search properties, :abbreviation:`i.e.` to employ regular expressions, be case sensitive, or use search similarity.
Similarity allows a text match to be a certain number of characters different from the search text.
These search properties are stored in the SearchDescriptor_ service, which is accessed by calling ``XSearchable.createSearchDescriptor()``. For example:

.. tabs::

    .. code-tab:: python

        # in ReplaceAll._search_iter() of replace_all.py
        # ...
        srch = Lo.qi(XSearchable, cell_rng, True)
        sd = srch.createSearchDescriptor()

        sd.setSearchString(srch_str)
        sd.setPropertyValue("SearchWords", True)
        # sd.setPropertyValue("SearchRegularExpression", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

XSearchDescriptor_ is the interface for the SearchDescriptor service, as shown in :numref:`ch26fig_search_and_replace_descriptors`.

..
    figure 3

.. cssclass:: diagram invert

    .. _ch26fig_search_and_replace_descriptors:
    .. figure:: https://user-images.githubusercontent.com/4193389/205419614-20a90c20-b240-456f-8b76-92880edef451.png
        :alt: The ReplaceDescriptor and SearchDescriptor Services.
        :figclass: align-center

        :The ReplaceDescriptor_ and SearchDescriptor_ Services.

Aside from being used to set search properties, XSearchDescriptor_ is also where the search string is stored:

.. tabs::

    .. code-tab:: python

        sd.setSearchString("dog")  # search for "dog"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

If regular expressions have been enabled, then the search string can utilize them:

.. tabs::

    .. code-tab:: python

        # search for a non-empty series of lower-case letters
        sd.setSearchString("[a-z]+")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The regular expression syntax is standard, and documented online at `List of Regular Expressions <https://help.libreoffice.org/latest/en-US/text/shared/01/02100001.html>`__.

..  _ch26_search_iteratively:

26.1 Searching Iteratively
==========================

The ``_search_iter()`` method in |replace_all_py|_ is passed the cell range for the ``15 x 6`` grid of animals, and creates a search based on finding complete words.
It uses ``XSearchable.findFirst()`` and ``XSearchable.findNext()`` to incrementally move through the grid:

.. tabs::

    .. code-tab:: python

        # in ReplaceAll._search_iter() of replace_all.py
        def _search_iter(self, sheet: XSpreadsheet, cell_rng: XCellRange, srch_str: str) -> None:
            print(f'Searching (iterating) for all occurrences of "{srch_str}"')
            try:
                srch = Lo.qi(XSearchable, cell_rng, True)
                sd = srch.createSearchDescriptor()

                sd.setSearchString(srch_str)
                # only complete words will be found
                sd.setPropertyValue("SearchWords", True)
                # sd.setPropertyValue("SearchRegularExpression", True)

                cr = Lo.qi(XCellRange, srch.findFirst(sd))
                if cr is None:
                    print(f'  No match found for "{srch_str}"')
                    return
                count = 0
                while cr is not None:
                    self._highlight(cr)
                    print(f"  Match {count + 1} : {Calc.get_range_str(cr)}")
                    cr = Lo.qi(XCellRange, srch.findNext(cr, sd))
                    count += 1

            except Exception as e:
                print(e)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``_highlight()`` method is as follows:

.. tabs::

    .. code-tab:: python

        from ooodev.format import Styler
        from ooodev.format.calc.direct.cell.background import Color as BgColor
        from ooodev.format.calc.direct.cell.borders import Borders, BorderLineKind, Side, LineSize
        from ooodev.format.calc.direct.cell.font import Font
        from ooodev.format.calc.direct.cell.standard_color import StandardColor
        # ... other imports

        # in ReplaceAll._highlight() of replace_all.py
        def _highlight(self, cr: XCellRange) -> None:
            # highlight by make cell bold, with text color of Light purple and
            # a background color of light blue.
            ft = Font(b=True, color=StandardColor.PURPLE_LIGHT1)
            bg_color = BgColor(StandardColor.DEFAULT_BLUE)
            bdrs = Borders(border_side=Side(line=BorderLineKind.SOLID, color=StandardColor.RED_DARK3))
            Styler.apply(cr, ft, bg_color, bdrs)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XSearchable.findNext()`` requires a reference to the previous match as its first input argument, so it can resume the search after that match.

.. tabs::

    .. code-tab:: python

        srch = Lo.qi(XSearchable, cell_rng, True)
        # ...
        o_first = srch.findFirst(sd)
        Info.show_services("Find First", o_first)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


When the services are listed for the references returned by ``XSearchable.findFirst()`` and ``XSearchable.findNext()`` by calling :py:meth:`.Info.show_services`
the following is show.

::

    Find First Supported Services (7)
    'com.sun.star.sheet.SheetCell'
    'com.sun.star.sheet.SheetCellRange'
    'com.sun.star.style.CharacterProperties'
    'com.sun.star.style.ParagraphProperties'
    'com.sun.star.table.Cell'
    'com.sun.star.table.CellProperties'
    'com.sun.star.table.CellRange'

The main service supported by the ``findFirst()`` result is SheetCell_.
This makes sense since the search is looking for a cell containing the search string.
As a consequence, the ``o_first`` reference can be converted to XCell_:

.. tabs::

    .. code-tab:: python

        cr = Lo.qi(XCell, srch.findFirst(sd))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

However, checking out ``XSearchable.findNext()`` in the same way showed an occasional problem:

.. tabs::

    .. code-tab:: python

        o_next  = srch.findNext(cr, sd)
        Info.show_services("Find Next", o_next)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The reference returned by ``findNext()`` usually supports the SheetCell_ service, but sometimes represents SheetCellRange_ instead!
When that occurs, code that attempts to convert ``o_next`` to XCell_ will return ``None``:

.. tabs::

    .. code-tab:: python

        cell = Lo.qi(XCell, srch.findNext(o_first, sd))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The solution is shown in the ``_search_iter()`` listing above - instead of converting the ``XSearchable.findFirst()`` and ``XSearchable.findNext()`` results to XCell_,
they're changed into XCellRange_ references, which always succeeds.

``_search_iter()`` calls ``_highlight()`` on each match so the user can see the results more clearly, as in :numref:`ch26fig_dog_search_result`.

..
    figure 4

.. cssclass:: screen_shot invert

    .. _ch26fig_dog_search_result:
    .. figure:: https://user-images.githubusercontent.com/4193389/205421272-cec25ea1-34d3-4b1d-90a0-39d1d866716e.png
        :alt: The Results of _search_iter when Looking for dog
        :figclass: align-center

        :The Results of ``_search_iter()`` when Looking for "dog".

.. _ch26_search_for_all_matches:

26.2 Searching For All Matches
==============================

The ``_search_all()`` method in |replace_all_py|_ utilizes ``XSearchable.findAll()`` to return all the search matches at once, in the form of an indexed container.
:py:meth:`.Calc.find_all` adds an extra conversion step, creating list of XCellRange_ objects from the values in the container:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def find_all(srch: XSearchable, sd: XSearchDescriptor) -> List[XCellRange] | None:
            con = srch.findAll(sd)
            if con is None:
                Lo.print("Match result is null")
                return None
            c_count = con.getCount()
            if c_count == 0:
                Lo.print("No matches found")
                return None

            crs = []
            for i in range(c_count):
                try:
                    cr = Lo.qi(XCellRange, con.getByIndex(i))
                    if cr is None:
                        continue
                    crs.append(cr)
                except Exception:
                    Lo.print(f"Could not access match index {i}")
            if len(crs) == 0:
                Lo.print(f"Found {c_count} matches but unable to access any match")
                return None
            return crs

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


``_search_all()`` iterates through the XCellRange_ list returned by :py:meth:`.Calc.find_all`, highlighting each match in the same way as the ``_search_iter()`` method:

.. tabs::

    .. code-tab:: python

        # in ReplaceAll._search_all() of replace_all.py
        def _search_all(self, sheet: XSpreadsheet, cell_rng: XCellRange, srch_str: str) -> None:
            print(f'Searching (find all) for all occurrences of "{srch_str}"')
            try:
                srch = Lo.qi(XSearchable, cell_rng, True)
                sd = srch.createSearchDescriptor()

                sd.setSearchString(srch_str)
                sd.setPropertyValue("SearchWords", True)

                match_crs = Calc.find_all(srch=srch, sd=sd)
                if not match_crs:
                    print(f'  No match found for "{srch_str}"')
                    return
                for i, cr in enumerate(match_crs):
                    self._highlight(cr)
                    print(f"  Index {i} : {Calc.get_range_str(cr)}")

            except Exception as e:
                print(e)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch26_replacing_all_matches:

26.3 Replacing All Matches
==========================

The XReplaceable_ interface only contains a ``replaceAll()`` method (see :numref:`ch26fig_search_and_replace_descriptors`), so there's no way to implement an iterative replace function.
In addition, ``XReplaceable.replaceAll()`` returns a count of the number of changes, not a container of the matched cells like ``XSearchable.findAll()``.
This means that its not possible to code a replace-like version of the ``_search_all()`` method which highlights all the changed cells.

The best that can be done is to execute two searches over the grid of animal names.
The first looks only for the search string so it can highlight the matching cells.
The second search calls ``XReplaceable.replaceAll()`` to make the changes.

The ``_replace_all()`` method is:

.. tabs::

    .. code-tab:: python

        # in ReplaceAll._replace_all() of replace_all.py
        def _replace_all(self, cell_rng: XCellRange, srch_str: str, repl_str: str) -> None:
            print(f'Replacing "{srch_str}" with "{repl_str}"')
            Lo.delay(2000)  # wait a bit before search & replace
            try:
                repl = Lo.qi(XReplaceable, cell_rng, True)
                rd = repl.createReplaceDescriptor()

                rd.setSearchString(srch_str)
                rd.setReplaceString(repl_str)
                rd.setPropertyValue("SearchWords", True)
                # rd.setPropertyValue("SearchRegularExpression", True)

                count = repl.replaceAll(rd)
                print(f"Search text replaced {count} times")
                print()

            except Exception as e:
                print(e)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The coding style is similar to my ``_search_all()`` method from above.
One difference is that XReplaceDescriptor_ is used to setup the search and replacement strings.

If ``rd.setPropertyValue("SearchRegularExpression", True)`` is uncommented then ``_replace_all()`` could be called using regular expressions in the function:

.. tabs::

    .. code-tab:: python

        self._replace_all(cell_rng=cell_rng, srch_str="[a-z]+", repl_str="ram")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The search string (``[a-z]+``) will match every cell's text, and change all the animal names to ``ram``.
Typical output is shown in :numref:`ch26fig_all_is_the_one`.

..
    figure 5

.. cssclass:: screen_shot invert

    .. _ch26fig_all_is_the_one:
    .. figure:: https://user-images.githubusercontent.com/4193389/205422346-019f174c-59fe-473f-b917-b82eaa8cc938.png
        :alt: All Animals Become One
        :figclass: align-center

        :All Animals Become One.

.. |g_secrets| replace::  Garlic Secrets
.. _g_secrets: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets

.. |g_secrets_py| replace:: garlic_secrets.py
.. _g_secrets_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_garlic_secrets/garlic_secrets.py

.. |replace_all| replace:: Replace All
.. _replace_all: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_replace_all

.. |replace_all_py| replace:: replace_all.py
.. _replace_all_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_replace_all/replace_all.py

.. _ReplaceDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1ReplaceDescriptor.html
.. _SearchDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html
.. _SheetCell: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCell.html
.. _SheetCell: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCell.html
.. _SheetCellRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1SheetCellRange.html
.. _XCell: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCell.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XInterface: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1uno_1_1XInterface.html
.. _XReplaceable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XReplaceable.html
.. _XReplaceDescriptor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XReplaceDescriptor.html
.. _XSearchable: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchable.html
.. _XSearchDescriptor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XSearchDescriptor.html
