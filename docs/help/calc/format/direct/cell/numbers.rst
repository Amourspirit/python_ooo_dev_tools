.. _help_calc_format_direct_cell_numbers:

Calc Direct Cell Numbers
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.direct.cell.numbers.Numbers` class gives you the similar options
as Calc's Font Dialog, but without the dialog. as seen in :numref:`236359913-91acf4fe-f6be-4e2b-be71-50a93b4991cb`.

.. cssclass:: screen_shot

    .. _236359913-91acf4fe-f6be-4e2b-be71-50a93b4991cb:

    .. figure:: https://user-images.githubusercontent.com/4193389/236359913-91acf4fe-f6be-4e2b-be71-50a93b4991cb.png
        :alt: Calc Format Cell dialog Numbers
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Numbers


Apply the Numbers to a cell
---------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 17,18

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.direct.cell.numbers import Numbers
        from ooodev.format.calc.direct.cell.numbers import NumberFormatEnum, NumberFormatIndexEnum

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                num_style = Numbers(num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED)
                Calc.set_val(value=-123.0, cell=cell, styles=[num_style])

                f_style = Numbers.from_obj(cell)
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the Numbers format
^^^^^^^^^^^^^^^^^^^^^^^^^^

The ``NumberFormatIndexEnum`` contains the values in |num_fmt_index|_ for easy lookup.

.. tabs::

    .. code-tab:: python

        num_style = Numbers(num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED)
        Calc.set_val(value=-123.0, cell=cell, styles=[num_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236360187-29a4270f-d133-4bd8-bd89-3a99436f9b91` and :numref:`236360255-51792c21-2b1c-4b30-9aae-4220aca8a79f`.

.. cssclass:: screen_shot

    .. _236360187-29a4270f-d133-4bd8-bd89-3a99436f9b91:

    .. figure:: https://user-images.githubusercontent.com/4193389/236360187-29a4270f-d133-4bd8-bd89-3a99436f9b91.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236360255-51792c21-2b1c-4b30-9aae-4220aca8a79f:

    .. figure:: https://user-images.githubusercontent.com/4193389/236360255-51792c21-2b1c-4b30-9aae-4220aca8a79f.png
        :alt: Calc Format Cell dialog Number Format set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Number Format set


Getting the number format from a cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Numbers.from_obj(cell)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply the Number to a range
---------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 19, 20

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.direct.cell.numbers import Numbers

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                Calc.set_val(value=0.000000034, sheet=sheet, cell_name="A1")
                Calc.set_val(value=0.000000013, sheet=sheet, cell_name="B1")
                rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

                num_style = Numbers().scientific
                num_style.apply(rng)

                f_style = Numbers.from_obj(rng)
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the Numbers format
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python
    

        # ... other code
        num_style = Numbers().scientific
        num_style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236360796-b4acf0fc-a2d5-4ce3-b303-c1ca5ecfd380` and :numref:`236360836-1554eea4-1386-400e-b5fb-e2879ba9913b`.

.. cssclass:: screen_shot

    .. _236360796-b4acf0fc-a2d5-4ce3-b303-c1ca5ecfd380:

    .. figure:: https://user-images.githubusercontent.com/4193389/236360796-b4acf0fc-a2d5-4ce3-b303-c1ca5ecfd380.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236360836-1554eea4-1386-400e-b5fb-e2879ba9913b:

    .. figure:: https://user-images.githubusercontent.com/4193389/236360836-1554eea4-1386-400e-b5fb-e2879ba9913b.png
        :alt: Calc Format Cell dialog Number Format set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Number Format set


Getting the number format from a range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Numbers.from_obj(rng)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_cell_numbers`
        - |num_fmt|_
        - |num_fmt_index|_
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`Calc.get_cell_range() <ooodev.office.calc.Calc.get_cell_range>`
        - :py:meth:`Calc.get_cell() <ooodev.office.calc.Calc.get_cell>`

.. |num_fmt| replace:: API NumberFormat
.. _num_fmt: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html

.. |num_fmt_index| replace:: API NumberFormatIndex
.. _num_fmt_index: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html