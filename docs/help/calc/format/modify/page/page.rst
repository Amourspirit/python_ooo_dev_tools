.. _help_calc_format_modify_page_page:

Calc Modify Page Style
======================


.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

Allows you to define page layouts for single and multiple-page documents, as well as a numbering and paper formats.

The :py:class:`ooodev.format.calc.modify.page.page.PaperFormat`, :py:class:`ooodev.format.calc.modify.page.page.Margins`, and :py:class:`ooodev.format.calc.modify.page.page.LayoutSettings`
classes are used to modify the page style values seen in :numref:`236647843-47f49c90-f2d7-45a6-bed9-36e81728da61`.


Default Page Style Dialog

.. cssclass:: screen_shot

    .. _236647843-47f49c90-f2d7-45a6-bed9-36e81728da61:

    .. figure:: https://user-images.githubusercontent.com/4193389/236647843-47f49c90-f2d7-45a6-bed9-36e81728da61.png
        :alt: Calc dialog Page Style default
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.modify.page.page import PaperFormat, PaperFormatKind, CalcStylePageKind


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                style = PaperFormat.from_preset(
                    preset=PaperFormatKind.A3, landscape=False, style_name=CalcStylePageKind.DEFAULT
                )
                style.apply(doc)

                style_obj = PaperFormat.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Paper Format
------------

Select from a list of predefined paper sizes, or define a custom paper format.

Setting Paper Format
^^^^^^^^^^^^^^^^^^^^

A preset can be used to set the paper format via :py:class:`~ooodev.format.inner.preset.preset_paper_format.PaperFormatKind` class.

From a preset
"""""""""""""

.. tabs::

    .. code-tab:: python


        from ooodev.format.calc.modify.page.page import PaperFormat, PaperFormatKind, CalcStylePageKind
        # ... other code

        style = PaperFormat.from_preset(
            preset=PaperFormatKind.A3, landscape=False, style_name=CalcStylePageKind.DEFAULT
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results of preset can be seen in :numref:`236648019-b3d6b1ac-88b0-4f3f-97aa-2dcce1074698`.

.. cssclass:: screen_shot

    .. _236648019-b3d6b1ac-88b0-4f3f-97aa-2dcce1074698:

    .. figure:: https://user-images.githubusercontent.com/4193389/236648019-b3d6b1ac-88b0-4f3f-97aa-2dcce1074698.png
        :alt: Calc dialog Page Style Paper Format modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Paper Format modified

Using SizeMM
""""""""""""

Custom size can be set using :py:class:`~ooodev.utils.data_type.size_mm.SizeMM` class.
If the height is greater than the width, the page will be set to portrait mode; Otherwise, it will be set to landscape mode.

.. tabs::

    .. code-tab:: python


        from ooodev.format.calc.modify.page.page import PaperFormat, CalcStylePageKind
        from ooodev.format.calc.modify.page.page import SizeMM
        # ... other code

        style = PaperFormat(
            size=SizeMM(width=200.0, height=100.0),
            style_name=CalcStylePageKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results can be seen in :numref:`236648332-29e962db-76d7-45a9-89db-be622a8b44b8`.

.. cssclass:: screen_shot

    .. _236648332-29e962db-76d7-45a9-89db-be622a8b44b8:

    .. figure:: https://user-images.githubusercontent.com/4193389/236648332-29e962db-76d7-45a9-89db-be622a8b44b8.png
        :alt: Calc dialog Page Style Paper Format modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Paper Format modified


Set size in other units
""""""""""""""""""""""""

The :py:class:`~ooodev.utils.data_type.size_mm.SizeMM` class can also take other units as parameters.
Any unit that supports :ref:`proto_unit_obj` can used to set the size.

Setting the size using inches.

In this example the size is set to ``8.5`` inches by ``14`` inches.

.. tabs::

    .. code-tab:: python


        from ooodev.format.calc.modify.page.page import PaperFormat, CalcStylePageKind
        from ooodev.format.calc.modify.page.page import SizeMM
        from ooodev.units import UnitInch
        # ... other code

        style = PaperFormat(
            size=SizeMM(width=UnitInch(8.5), height=UnitInch(14)),
            style_name=CalcStylePageKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results can be seen in :numref:`236651021-87330d6a-cdd4-4405-9592-8bd146ec1089`.

.. cssclass:: screen_shot

    .. _236651021-87330d6a-cdd4-4405-9592-8bd146ec1089:

    .. figure:: https://user-images.githubusercontent.com/4193389/236651021-87330d6a-cdd4-4405-9592-8bd146ec1089.png
        :alt: Calc dialog Page Style Paper Format modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Paper Format modified

Getting Paper Format from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PaperFormat.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Margins
-------

Specify the amount of space to leave between the edges of the page and the document text.

Setting Margins
^^^^^^^^^^^^^^^

Set margins in millimeters
""""""""""""""""""""""""""

The default margin values are in millimeters.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.page import Margins, CalcStylePageKind
        # ... other code

        style = Margins(left=10, right=10, top=18, bottom=18, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236651212-5602facf-209c-436a-b91f-d19f82a97b04:

    .. figure:: https://user-images.githubusercontent.com/4193389/236651212-5602facf-209c-436a-b91f-d19f82a97b04.png
        :alt: Calc dialog Page Style Margins modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Margins modified

Set margins in other units
""""""""""""""""""""""""""

The margins can also take other units as parameters.
Any unit that supports :ref:`proto_unit_obj` can used to set the margin value.

In the following example the margins are set to ``1`` inch on the left and right, ``1.2`` inches on the top, and ``0.75`` inches on the bottom.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.page import Margins, CalcStylePageKind
        from ooodev.units import UnitInch
        # ... other code

        style = Margins(
            left=UnitInch(1.0),
            right=UnitInch(1.0),
            top=UnitInch(1.2),
            bottom=UnitInch(0.75),
            style_name=CalcStylePageKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236651329-79dc0d0b-d86b-4d63-a009-08f64c63940c:

    .. figure:: https://user-images.githubusercontent.com/4193389/236651329-79dc0d0b-d86b-4d63-a009-08f64c63940c.png
        :alt: Calc dialog Page Style Margins modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Margins modified

Getting Margins from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Margins.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Layout Settings
---------------

Setting Layout
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.page import LayoutSettings, PageStyleLayout
        from ooodev.format.calc.modify.page.page import NumberingTypeEnum, CalcStylePageKind
        # ... other code

        style = LayoutSettings(
            layout=PageStyleLayout.MIRRORED,
            numbers=NumberingTypeEnum.CHARS_UPPER_LETTER,
            align_hori=True,
            align_vert=True,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236651545-58dd23ba-a1d7-4b74-96fb-77c645577d61:

    .. figure:: https://user-images.githubusercontent.com/4193389/236651545-58dd23ba-a1d7-4b74-96fb-77c645577d61.png
        :alt: Calc dialog Page Style Borders style shadow modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Borders style shadow modified

Getting Layout Settings from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = LayoutSettings.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.page.page.PaperFormat`
        - :py:class:`ooodev.format.calc.modify.page.page.Margins`
        - :py:class:`ooodev.format.calc.modify.page.page.LayoutSettings`