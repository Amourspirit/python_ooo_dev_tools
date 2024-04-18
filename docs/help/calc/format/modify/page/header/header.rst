.. _help_calc_format_modify_page_header_header:

Calc Modify Page Header
=======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.modify.page.header.Header` class is used to modify the page header values seen in :numref:`236652959-e1ed9ada-6f58-4814-b2ed-ddc3a44a3615` of a Calc document.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.page.header import Header, CalcStylePageKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                header_style = Header(
                    on=True,
                    shared_first=True,
                    shared=True,
                    height=10.0,
                    spacing=3.0,
                    margin_left=1.5,
                    margin_right=2.0,
                    style_name=CalcStylePageKind.DEFAULT,
                )
                header_style.apply(doc)

                style_obj = Header.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Applying Header Style
---------------------

Before applying style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _236652959-e1ed9ada-6f58-4814-b2ed-ddc3a44a3615:

    .. figure:: https://user-images.githubusercontent.com/4193389/236652959-e1ed9ada-6f58-4814-b2ed-ddc3a44a3615.png
        :alt: Calc dialog Page Header default
        :figclass: align-center
        :width: 450px

        Calc dialog Page Header default

Apply Style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        header_style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            margin_left=1.5,
            margin_right=2.0,
            style_name=CalcStylePageKind.DEFAULT,
        )
        header_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _236653094-4fcc9ef8-628f-483d-856d-3af3deff767f:

    .. figure:: https://user-images.githubusercontent.com/4193389/236653094-4fcc9ef8-628f-483d-856d-3af3deff767f.png
        :alt: Calc dialog Page Header set with Header class
        :figclass: align-center
        :width: 450px

        Calc dialog Page Header set with Header class


Getting the Header from a style
-------------------------------

.. tabs::

    .. code-tab:: python

        style_obj = Header.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
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
        - :ref:`help_calc_format_modify_page_footer_footer`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.modify.page.header.Header`
        - :py:class:`Calc.set_head_foot() <ooodev.office.calc.Calc.set_head_foot>`
