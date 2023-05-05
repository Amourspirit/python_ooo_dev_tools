.. _help_calc_format_direct_cell_font_effects:

Calc Direct Cell FontEffects
============================

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
        :emphasize-lines: 17, 18, 19, 20, 21, 22

        from ooodev.format import CommonColor
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.direct.cell.font import FontEffects, FontLine, FontUnderlineEnum

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=True)):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")

                font_effects = FontEffects(
                    color=CommonColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
                    shadowed=True,
                )
                Calc.set_val(value="Hello", cell=cell, styles=[font_effects])

                f_effects = FontEffects.from_obj(cell)
                assert f_effects.prop_color == CommonColor.RED

                Lo.delay(1_000)
                Lo.close_doc(doc)
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

        font_effects = FontEffects(
            color=CommonColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
            shadowed=True,
        )
        Calc.set_val(value="Hello", cell=cell, styles=[font_effects])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`235963998-66f9c902-b97c-47ca-b8a2-048670e39511` and :numref:`235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686`.

.. cssclass:: screen_shot

    .. _235963998-66f9c902-b97c-47ca-b8a2-048670e39511:
    .. figure:: https://user-images.githubusercontent.com/4193389/235963998-66f9c902-b97c-47ca-b8a2-048670e39511.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686:
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

        f_effects = FontEffects.from_obj(cell)
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
        :emphasize-lines: 19, 20, 21, 22, 23, 24

        from ooodev.format import CommonColor
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.direct.cell.font import FontEffects, FontLine, FontUnderlineEnum

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=True)):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")
                Calc.set_val(value="World", sheet=sheet, cell_name="B1")
                rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

                font_effects = FontEffects(
                    color=CommonColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
                    shadowed=True,
                )
                font_effects.apply(rng)

                f_effects = FontEffects.from_obj(rng)
                assert f_effects.prop_color == CommonColor.RED

                Lo.delay(1_000)
                Lo.close_doc(doc)
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

        font_effects = FontEffects(
            color=CommonColor.RED,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
            shadowed=True,
        )
        font_effects.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`235968294-45fd9003-f462-4db1-bc92-982b88659b57` and :numref:`235963671-a3f8f543-26ec-4a91-b3cf-e1ef753de686`.

.. cssclass:: screen_shot

    .. _235968294-45fd9003-f462-4db1-bc92-982b88659b57:
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

        f_effects = FontEffects.from_obj(cell)
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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.direct.cell.font.FontEffects`
        - :py:meth:`Calc.get_cell_range() <ooodev.office.calc.Calc.get_cell_range>`
        - :py:meth:`Calc.get_cell() <ooodev.office.calc.Calc.get_cell>`