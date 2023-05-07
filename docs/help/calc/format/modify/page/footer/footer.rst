.. _help_calc_format_modify_page_footer_footer:

Calc Modify Page Footer
=======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.modify.page.footer.Footer` class is used to modify the page footer values seen in :numref:`236653160-5c2eb66f-a672-4963-960b-6c8e165165b3` of a Calc document.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.modify.page.footer import Footer, CalcStylePageKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                footer_style = Footer(
                    on=True,
                    shared_first=True,
                    shared=True,
                    height=10.0,
                    spacing=3.0,
                    margin_left=1.5,
                    margin_right=2.0,
                    style_name=CalcStylePageKind.DEFAULT,
                )
                footer_style.apply(doc)

                style_obj = Footer.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Applying Footer Style
---------------------

Before applying style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _236653160-5c2eb66f-a672-4963-960b-6c8e165165b3:

    .. figure:: https://user-images.githubusercontent.com/4193389/236653160-5c2eb66f-a672-4963-960b-6c8e165165b3.png
        :alt: Calc dialog Page Footer default
        :figclass: align-center
        :width: 450px

        Calc dialog Page Footer default

Apply Style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        footer_style = Footer(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            margin_left=1.5,
            margin_right=2.0,
            style_name=CalcStylePageKind.DEFAULT,
        )
        footer_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _236653201-050a08eb-0ca3-48af-be6a-59a2a546836f:

    .. figure:: https://user-images.githubusercontent.com/4193389/236653201-050a08eb-0ca3-48af-be6a-59a2a546836f.png
        :alt: Calc dialog Page Footer set with Footer class
        :figclass: align-center
        :width: 450px

        Calc dialog Page Footer set with Footer class


Getting the Footer from a style
-------------------------------

.. tabs::

    .. code-tab:: python

        style_obj = Footer.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
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
        - :ref:`help_calc_format_modify_page_header_header`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.page.footer.Footer`
