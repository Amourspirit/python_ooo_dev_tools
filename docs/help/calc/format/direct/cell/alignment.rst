.. _help_calc_format_direct_cell_alignment:

Calc Direct Cell Alignment
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

Calc has a dialog, as seen in :numref:`ss_calc_format_direct_cell_alignment_dialog_1`, that sets cell alignment.
In this section we will look the various classed that set the same options.

.. cssclass:: screen_shot

    .. _ss_calc_format_direct_cell_alignment_dialog_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236059462-8c9e8c63-de00-4505-9865-f1485d460c86.png
        :alt: Calc Format Cell dialog Direct Cell Alignment
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Direct Cell Alignment

Text Alignment
--------------

The ``style_align_text()`` method is called to set the text alignment of a cell or range.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.format.calc.direct.cell.alignment import HoriAlignKind, VertAlignKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_align_text(
                    hori_align=HoriAlignKind.CENTER, vert_align=VertAlignKind.MIDDLE
                )

                Lo.delay(1_000)
                doc.close()
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
        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_align_text(
            hori_align=HoriAlignKind.CENTER,
            vert_align=VertAlignKind.MIDDLE,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236063001-b8a31737-4f2d-4955-8a48-a6669d3e74eb_1` and :numref:`236063206-8094e9f5-b8de-49ea-aa25-375f1889e961_1`.

.. cssclass:: screen_shot

    .. _236063001-b8a31737-4f2d-4955-8a48-a6669d3e74eb_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236063001-b8a31737-4f2d-4955-8a48-a6669d3e74eb.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236063206-8094e9f5-b8de-49ea-aa25-375f1889e961_1:

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

        f_style = cell.style_align_text_get()
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
        cell = sheet["A1"]
        cell.value = "Hello"
        cell = cell.get_cell_right()
        cell.value = "World"
        cell_rng = sheet.get_range(range_name="A1:B1")
        cell_rng.style_align_text(
            hori_align=HoriAlignKind.CENTER,
            indent=3,
            vert_align=VertAlignKind.MIDDLE,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236066605-72802b3c-2a39-4f20-81c3-e6acebdf8328_1` and :numref:`236066708-228b4cf2-2763-4e08-b163-c35e76e9136e_1`.

.. cssclass:: screen_shot

    .. _236066605-72802b3c-2a39-4f20-81c3-e6acebdf8328_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236066605-72802b3c-2a39-4f20-81c3-e6acebdf8328.png
        :alt: Calc Cell Range
        :figclass: align-center
        :width: 520px

        Calc Cell Range

    .. _236066708-228b4cf2-2763-4e08-b163-c35e76e9136e_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236066708-228b4cf2-2763-4e08-b163-c35e76e9136e.png
        :alt: Calc Format Range dialog Text Alignment set
        :figclass: align-center
        :width: 450px

        Calc Format Range dialog Text Alignment set

Getting the text alignment from a Cell Range
""""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = cell_rng.style_align_text_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Orientation
----------------

The ``style_align_orientation()`` method is called to set the text orientation of a cell or range.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.format.inner.direct.calc.alignment.text_orientation import EdgeKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_align_orientation(
                    vert_stack=False,
                    rotation=-10,
                    edge=EdgeKind.INSIDE,
                )

                f_style = cell.style_align_orientation_get()
                assert f_style is not None

                Lo.delay(1_000)
                doc.close()
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
        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_align_orientation(
            vert_stack=False,
            rotation=-10,
            edge=EdgeKind.INSIDE,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236069220-693024f3-dbd9-4c49-a16d-1d6c2b6e088b_1` and :numref:`236069303-908569cd-cc3c-4486-80f6-ba20c8c63c73_1`.

.. cssclass:: screen_shot

    .. _236069220-693024f3-dbd9-4c49-a16d-1d6c2b6e088b_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236069220-693024f3-dbd9-4c49-a16d-1d6c2b6e088b.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236069303-908569cd-cc3c-4486-80f6-ba20c8c63c73_1:

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
        f_style = cell.style_align_orientation_get()
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
        rng = sheet.rng("A1:B1")
        sheet.set_array(values=[["Hello", "World"]], range_obj=rng)
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_align_orientation(vert_stack=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236071231-64e99eb6-6a59-4ab5-80de-5f5a165f7090_1` and :numref:`236071295-eaace095-5e8f-47e3-905f-01784d795486_1`.

.. cssclass:: screen_shot

    .. _236071231-64e99eb6-6a59-4ab5-80de-5f5a165f7090_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236071231-64e99eb6-6a59-4ab5-80de-5f5a165f7090.png
        :alt: Calc Cell Range
        :figclass: align-center
        :width: 520px

        Calc Cell Range

    .. _236071295-eaace095-5e8f-47e3-905f-01784d795486_1:

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

        f_style = cell_rng.style_align_orientation_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Properties
---------------

The ``style_align_properties()`` method is called to set the text properties of a cell or range.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python

    from __future__ import annotations
    import uno
    from ooodev.calc import CalcDoc
    from ooodev.loader import Lo
    from ooodev.format.inner.direct.calc.alignment.properties import TextDirectionKind

    def main() -> int:
        with Lo.Loader(connector=Lo.ConnectSocket()):
            doc = CalcDoc.create_doc(visible=True)
            sheet = doc.sheets[0]
            Lo.delay(500)
            doc.zoom_value(400)

            cell = sheet["A1"]
            cell.value = "Hello World! Sunny Day!"
            cell.style_align_properties(
                wrap_auto=True,
                hyphen_active=True,
                direction=TextDirectionKind.PAGE,
            )

            Lo.delay(1_000)
            doc.close()
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
        cell = sheet["A1"]
        cell.value = "Hello World! Sunny Day!"
        cell.style_align_properties(
            wrap_auto=True,
            hyphen_active=True,
            direction=TextDirectionKind.PAGE,
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236075054-7ee77e37-7f93-4cef-8867-9d61b87eccef_1` and :numref:`236075133-1fe50a07-3e71-4090-aacf-b6da5d255ecc_1`.

.. cssclass:: screen_shot

    .. _236075054-7ee77e37-7f93-4cef-8867-9d61b87eccef_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075054-7ee77e37-7f93-4cef-8867-9d61b87eccef.png
        :alt: Calc Cell
        :figclass: align-center
        :width: 520px

        Calc Cell

    .. _236075133-1fe50a07-3e71-4090-aacf-b6da5d255ecc_1:

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

        f_style = cell.style_align_properties_get()
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
        rng = sheet.rng("A1:B1")
        sheet.set_array(
            values=[["Hello World! See the Shine", "Sunny Days are great!"]],
            range_obj=rng
        )
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_align_properties(shrink_to_fit=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236075781-396f1f66-2a89-413b-92af-3247c376ef09_1` and :numref:`236075827-4244bbab-9821-4c0d-842a-0ed03af3d921_1`.

.. cssclass:: screen_shot

    .. _236075781-396f1f66-2a89-413b-92af-3247c376ef09_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075781-396f1f66-2a89-413b-92af-3247c376ef09.png
        :alt: Calc Cell Range
        :figclass: align-center
        :width: 520px

        Calc Cell Range

    .. _236075827-4244bbab-9821-4c0d-842a-0ed03af3d921_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236075827-4244bbab-9821-4c0d-842a-0ed03af3d921.png
        :alt: Calc Format Cell dialog Text Orientation set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Text Orientation set

Getting the text properties from a cell range
"""""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = cell_rng.style_align_properties_get()
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
        - :ref:`help_writer_format_direct_para_alignment`
        - :ref:`help_calc_format_modify_cell_alignment`