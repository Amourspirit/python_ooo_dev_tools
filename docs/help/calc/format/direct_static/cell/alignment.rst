.. _help_calc_format_direct_static_cell_alignment:

Calc Direct Cell Alignment (Static)
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

Calc has a dialog, as seen in :numref:`ss_calc_format_direct_cell_alignment_dialog`, that sets cell alignment.
In this section we will look the various classed that set the same options.

.. cssclass:: screen_shot

    .. _ss_calc_format_direct_cell_alignment_dialog:

    .. figure:: https://user-images.githubusercontent.com/4193389/236059462-8c9e8c63-de00-4505-9865-f1485d460c86.png
        :alt: Calc Format Cell dialog Direct Cell Alignment
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Direct Cell Alignment

.. seealso::

    - :ref:`help_calc_format_direct_cell_alignment`

Text Alignment
--------------

The :py:class:`ooodev.format.calc.direct.cell.alignment.TextAlign` class sets the text alignment of a cell or range.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.alignment import TextAlign
        from ooodev.format.calc.direct.cell.alignment import HoriAlignKind, VertAlignKind


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                style = TextAlign(hori_align=HoriAlignKind.CENTER, vert_align=VertAlignKind.MIDDLE)
                Calc.set_val(value="Hello", cell=cell, styles=[style])

                f_style = TextAlign.from_obj(cell)
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to a cell
^^^^^^^^^^^^^^^

Setting the text alignment
""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = Calc.get_cell(sheet=sheet, cell_name="A1")
        style = TextAlign(hori_align=HoriAlignKind.CENTER, vert_align=VertAlignKind.MIDDLE)
        Calc.set_val(value="Hello", cell=cell, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236063001-b8a31737-4f2d-4955-8a48-a6669d3e74eb` and :numref:`236063206-8094e9f5-b8de-49ea-aa25-375f1889e961`.

.. cssclass:: screen_shot

    .. _236063001-b8a31737-4f2d-4955-8a48-a6669d3e74eb:

    .. figure:: https://user-images.githubusercontent.com/4193389/236063001-b8a31737-4f2d-4955-8a48-a6669d3e74eb.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236063206-8094e9f5-b8de-49ea-aa25-375f1889e961:

    .. figure:: https://user-images.githubusercontent.com/4193389/236063206-8094e9f5-b8de-49ea-aa25-375f1889e961.png
        :alt: Calc Format Cell dialog Text Alignment set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Text Alignment set

Getting the text alignment from a cell
""""""""""""""""""""""""""""""""""""""


.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = TextAlign.from_obj(cell)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to a range
^^^^^^^^^^^^^^^^

Setting the text alignment
""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")
        Calc.set_val(value="World", sheet=sheet, cell_name="B1")
        rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

        style = TextAlign(hori_align=HoriAlignKind.LEFT, indent=3, vert_align=VertAlignKind.TOP)
        style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236066605-72802b3c-2a39-4f20-81c3-e6acebdf8328` and :numref:`236066708-228b4cf2-2763-4e08-b163-c35e76e9136e`.

.. cssclass:: screen_shot

    .. _236066605-72802b3c-2a39-4f20-81c3-e6acebdf8328:

    .. figure:: https://user-images.githubusercontent.com/4193389/236066605-72802b3c-2a39-4f20-81c3-e6acebdf8328.png
        :alt: Calc Cell Range
        :figclass: align-center
        :width: 520px

        Calc Cell Range

    .. _236066708-228b4cf2-2763-4e08-b163-c35e76e9136e:

    .. figure:: https://user-images.githubusercontent.com/4193389/236066708-228b4cf2-2763-4e08-b163-c35e76e9136e.png
        :alt: Calc Format Range dialog Text Alignment set
        :figclass: align-center
        :width: 450px

        Calc Format Range dialog Text Alignment set

Getting the text alignment from a cell range
""""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = TextAlign.from_obj(rng)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Orientation
----------------

The :py:class:`ooodev.format.calc.direct.cell.alignment.TextOrientation` class sets the text orientation of a cell or range.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.alignment import TextOrientation, EdgeKind


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                style = TextOrientation(vert_stack=False, rotation=-10, edge=EdgeKind.INSIDE)
                Calc.set_val(value="Hello", cell=cell, styles=[style])

                f_style = TextOrientation.from_obj(cell)
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to a cell
^^^^^^^^^^^^^^^

Setting the text orientation
""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = Calc.get_cell(sheet=sheet, cell_name="A1")
        style = TextOrientation(vert_stack=False, rotation=-10, edge=EdgeKind.INSIDE)
        Calc.set_val(value="Hello", cell=cell, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236069220-693024f3-dbd9-4c49-a16d-1d6c2b6e088b` and :numref:`236069303-908569cd-cc3c-4486-80f6-ba20c8c63c73`.

.. cssclass:: screen_shot

    .. _236069220-693024f3-dbd9-4c49-a16d-1d6c2b6e088b:

    .. figure:: https://user-images.githubusercontent.com/4193389/236069220-693024f3-dbd9-4c49-a16d-1d6c2b6e088b.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236069303-908569cd-cc3c-4486-80f6-ba20c8c63c73:

    .. figure:: https://user-images.githubusercontent.com/4193389/236069303-908569cd-cc3c-4486-80f6-ba20c8c63c73.png
        :alt: Calc Format Cell dialog Text Orientation set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Text Orientation set

Getting the text orientation from a cell
""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = TextAlign.from_obj(cell)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to a range
^^^^^^^^^^^^^^^^

Setting the text orientation
""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")
        Calc.set_val(value="World", sheet=sheet, cell_name="B1")
        rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

        style = TextOrientation(vert_stack=True)
        style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236071231-64e99eb6-6a59-4ab5-80de-5f5a165f7090` and :numref:`236071295-eaace095-5e8f-47e3-905f-01784d795486`.

.. cssclass:: screen_shot

    .. _236071231-64e99eb6-6a59-4ab5-80de-5f5a165f7090:

    .. figure:: https://user-images.githubusercontent.com/4193389/236071231-64e99eb6-6a59-4ab5-80de-5f5a165f7090.png
        :alt: Calc Cell Range
        :figclass: align-center
        :width: 520px

        Calc Cell Range

    .. _236071295-eaace095-5e8f-47e3-905f-01784d795486:

    .. figure:: https://user-images.githubusercontent.com/4193389/236071295-eaace095-5e8f-47e3-905f-01784d795486.png
        :alt: Calc Format Cell dialog Text Orientation set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Text Orientation set

Getting the text orientation from a range
"""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = TextOrientation.from_obj(rng)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Properties
---------------

The :py:class:`ooodev.format.calc.direct.cell.alignment.Properties` class sets the text properties of a cell or range.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.alignment import Properties, TextDirectionKind


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                style = Properties(wrap_auto=True, hyphen_active=True, direction=TextDirectionKind.PAGE)
                Calc.set_val(value="Hello World! Sunny Day!", cell=cell, styles=[style])

                f_style = Properties.from_obj(cell)
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to a cell
^^^^^^^^^^^^^^^

Setting the text properties
"""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = Calc.get_cell(sheet=sheet, cell_name="A1")
        style = Properties(wrap_auto=True, hyphen_active=True, direction=TextDirectionKind.PAGE)
        Calc.set_val(value="Hello World! Sunny Day!", cell=cell, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236075054-7ee77e37-7f93-4cef-8867-9d61b87eccef` and :numref:`236075133-1fe50a07-3e71-4090-aacf-b6da5d255ecc`.

.. cssclass:: screen_shot

    .. _236075054-7ee77e37-7f93-4cef-8867-9d61b87eccef:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075054-7ee77e37-7f93-4cef-8867-9d61b87eccef.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236075133-1fe50a07-3e71-4090-aacf-b6da5d255ecc:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075133-1fe50a07-3e71-4090-aacf-b6da5d255ecc.png
        :alt: Calc Format Cell dialog Text Orientation set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Text Orientation set

Getting the text properties from a cell
"""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Properties.from_obj(cell)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to a range
^^^^^^^^^^^^^^^^

Setting the text properties
"""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = Calc.get_cell(sheet=sheet, cell_name="A1")
        style = Properties(wrap_auto=True, hyphen_active=True, direction=TextDirectionKind.PAGE)
        Calc.set_val(value="Hello World! Sunny Day!", cell=cell, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236075781-396f1f66-2a89-413b-92af-3247c376ef09` and :numref:`236075827-4244bbab-9821-4c0d-842a-0ed03af3d921`.

.. cssclass:: screen_shot

    .. _236075781-396f1f66-2a89-413b-92af-3247c376ef09:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075781-396f1f66-2a89-413b-92af-3247c376ef09.png
        :alt: Calc Cell Range
        :figclass: align-center
        :width: 520px

        Calc Cell Range

    .. _236075827-4244bbab-9821-4c0d-842a-0ed03af3d921:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075827-4244bbab-9821-4c0d-842a-0ed03af3d921.png
        :alt: Calc Format Cell dialog Text Orientation set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Text Orientation set

Getting the text properties from a cell
"""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Properties.from_obj(rng)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_calc_format_direct_cell_alignment`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_alignment`
        - :ref:`help_calc_format_modify_cell_alignment`
        - :py:class:`ooodev.format.calc.direct.cell.alignment.TextAlign`
        - :py:class:`ooodev.format.calc.direct.cell.alignment.TextOrientation`
        - :py:class:`ooodev.format.calc.direct.cell.alignment.Properties`