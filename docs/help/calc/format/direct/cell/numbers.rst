.. _help_calc_format_direct_cell_numbers:

Calc Direct Cell Numbers
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The ``style_numbers_numbers()`` method gives you the similar options
as Calc's Font Dialog, but without the dialog. as seen in :numref:`236359913-91acf4fe-f6be-4e2b-be71-50a93b4991cb_1`.

There are several other ``style_numbers_*`` methods that can be used to set the number format of a cell or range
such as, ``style_numbers_currency()`` and ``style_numbers_percent()``.

.. cssclass:: screen_shot

    .. _236359913-91acf4fe-f6be-4e2b-be71-50a93b4991cb_1:

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

        from __future__ import annotations
        import uno
        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = -123.0
                cell.style_numbers_numbers(
                    num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED,
                )

                Lo.delay(1_000)
                doc.close()
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

        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum

        # ... other code
        cell = sheet["A1"]
        cell.value = -123.0
        cell.style_numbers_numbers(
            num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236360187-29a4270f-d133-4bd8-bd89-3a99436f9b91_1` and :numref:`236360255-51792c21-2b1c-4b30-9aae-4220aca8a79f_1`.

.. cssclass:: screen_shot

    .. _236360187-29a4270f-d133-4bd8-bd89-3a99436f9b91_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236360187-29a4270f-d133-4bd8-bd89-3a99436f9b91.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236360255-51792c21-2b1c-4b30-9aae-4220aca8a79f_1:

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

        f_style = cell.style_numbers_numbers_get()
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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                rng = sheet.rng("A1:B1")
                sheet.set_array(values=[[0.000000034, 0.000000013]], range_obj=rng)

                cell_rng = sheet.get_range(range_obj=rng)
                cell_rng.style_numbers_scientific()

                Lo.delay(1_000)
                doc.close()
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
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_numbers_scientific()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236360796-b4acf0fc-a2d5-4ce3-b303-c1ca5ecfd380_1` and :numref:`236360836-1554eea4-1386-400e-b5fb-e2879ba9913b_1`.

.. cssclass:: screen_shot

    .. _236360796-b4acf0fc-a2d5-4ce3-b303-c1ca5ecfd380_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236360796-b4acf0fc-a2d5-4ce3-b303-c1ca5ecfd380.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236360836-1554eea4-1386-400e-b5fb-e2879ba9913b_1:

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

        f_style = cell_rng.style_numbers_numbers_get()
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
        - :py:class:`~ooodev.loader.Lo`

.. |num_fmt| replace:: API NumberFormat
.. _num_fmt: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html

.. |num_fmt_index| replace:: API NumberFormatIndex
.. _num_fmt_index: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html