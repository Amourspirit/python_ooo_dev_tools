.. _help_calc_format_direct_cell_font_effects:

Calc Direct Cell Font Effects
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.direct.cell.font.FontEffects` class gives you the same options
as Calc's Font Effects Dialog, but without the dialog as seen in :numref:`235955376-90701b09-42a9-472b-8b24-13e14bbd0a56`.

.. cssclass:: screen_shot

    .. _235955376-90701b09-42a9-472b-8b24-13e14bbd0a56:

    .. figure:: https://user-images.githubusercontent.com/4193389/235955376-90701b09-42a9-472b-8b24-13e14bbd0a56.png
        :alt: Calc Format Cell dialog Font Effects
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Font Effects


Apply the font effects to a cell
--------------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.color import CommonColor
        from ooodev.format.inner.direct.write.char.font.font_effects import (
            FontLine, FontUnderlineEnum
        )

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_font_effect(
                    color=CommonColor.RED,
                    underline=FontLine(
                        line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE
                    ),
                    shadowed=True,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the font effects
^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_font_effect(
            color=CommonColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`235963998-66f9c902-b97c-47ca-b8a2-048670e39511_1` and :numref:`235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686_1`.

.. cssclass:: screen_shot

    .. _235963998-66f9c902-b97c-47ca-b8a2-048670e39511_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/235963998-66f9c902-b97c-47ca-b8a2-048670e39511.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686.png
        :alt: Calc Format Cell dialog Font Effects set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Font Effects set


Getting the font effects from a cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = cell.style_font_effect_get()
        assert f_effects.prop_color == CommonColor.RED

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply the font effects to a range
---------------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.color import CommonColor
        from ooodev.format.inner.direct.write.char.font.font_effects import FontLine, FontUnderlineEnum

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                rng = sheet.rng("A1:B1")
                sheet.set_array(values=[["Hello", "World"]], range_obj=rng)

                cell_rng = sheet.get_range(range_obj=rng)
                cell_rng.style_font_effect(
                    color=CommonColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
                    shadowed=True,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the font effects
^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_font_effect(
            color=CommonColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`235968294-45fd9003-f462-4db1-bc92-982b88659b57` and :numref:`235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686_1`.

.. cssclass:: screen_shot

    .. _235968294-45fd9003-f462-4db1-bc92-982b88659b57_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/235968294-45fd9003-f462-4db1-bc92-982b88659b57.png
        :alt: Calc Range
        :figclass: align-center
        :width: 520px

        Calc Range

Getting the font effects from a range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = cell_rng.style_font_effect_get()
        assert f_effects.prop_color == CommonColor.RED

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font_effects`
        - :ref:`help_writer_format_modify_char_font_effects`
        - :ref:`help_writer_format_modify_para_font_effects`
        - :ref:`help_calc_format_direct_cell_font_only`
        - :ref:`help_calc_format_direct_cell_font`
        - :ref:`help_calc_format_modify_cell_font_effects`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.direct.cell.font.FontEffects`