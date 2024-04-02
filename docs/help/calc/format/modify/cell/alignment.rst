.. _help_calc_format_modify_cell_alignment:

Calc Modify Cell Alignment
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Calc has a dialog, as seen in :numref:`236448968-810ad7cc-04f4-4a73-9d1d-ca6598eac073`, that sets cell alignment. In this section we will look the various classes that set the same options.

.. cssclass:: screen_shot

    .. _236448968-810ad7cc-04f4-4a73-9d1d-ca6598eac073:

    .. figure:: https://user-images.githubusercontent.com/4193389/236448968-810ad7cc-04f4-4a73-9d1d-ca6598eac073.png
        :alt: Calc dialog style Alignment default
        :figclass: align-center
        :width: 450px

        Calc dialog style Alignment default


Text Alignment
--------------

The :py:class:`ooodev.format.calc.modify.cell.alignment.TextAlign` class sets the text alignment of a style.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14, 15, 16, 17, 18

        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.cell.alignment import TextAlign, StyleCellKind
        from ooodev.format.calc.modify.cell.alignment import HoriAlignKind, VertAlignKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                style = TextAlign(
                    hori_align=HoriAlignKind.CENTER,
                    vert_align=VertAlignKind.MIDDLE,
                    style_name=StyleCellKind.DEFAULT,
                )
                style.apply(doc)

                style_obj = TextAlign.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Setting the text alignment
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        style = TextAlign(
            hori_align=HoriAlignKind.CENTER,
            vert_align=VertAlignKind.MIDDLE,
            style_name=StyleCellKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236450660-af1cc6e9-feb5-47b5-a663-79781f8fcfda`.

.. cssclass:: screen_shot

    .. _236450660-af1cc6e9-feb5-47b5-a663-79781f8fcfda:

    .. figure:: https://user-images.githubusercontent.com/4193389/236450660-af1cc6e9-feb5-47b5-a663-79781f8fcfda.png
        :alt: Calc dialog style Text Alignment modified
        :figclass: align-center
        :width: 450px

        Calc dialog style Text Alignment modified

Getting the text alignment from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        style_obj = TextAlign.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
        assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Orientation
----------------

The :py:class:`ooodev.format.calc.modify.cell.alignment.TextOrientation` class sets the text alignment of a style.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 15, 16, 17, 18, 19, 20, 21

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.cell.alignment import TextOrientation
        from ooodev.format.calc.modify.cell.alignment import EdgeKind, StyleCellKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                style = TextOrientation(
                    vert_stack=False,
                    rotation=-10,
                    edge=EdgeKind.INSIDE,
                    style_name=StyleCellKind.DEFAULT,
                )
                style.apply(doc)

                style_obj = TextOrientation.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Setting the text alignment
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        style = TextOrientation(
            vert_stack=False,
            rotation=-10,
            edge=EdgeKind.INSIDE,
            style_name=StyleCellKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236453255-4822ad15-a1b8-4814-a5f8-695e28cde1a7`.

.. cssclass:: screen_shot

    .. _236453255-4822ad15-a1b8-4814-a5f8-695e28cde1a7:

    .. figure:: https://user-images.githubusercontent.com/4193389/236453255-4822ad15-a1b8-4814-a5f8-695e28cde1a7.png
        :alt: Calc dialog style Text Alignment modified
        :figclass: align-center
        :width: 450px

        Calc dialog style Text Alignment modified

Getting the text orientation from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        style_obj = TextOrientation.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
        assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Properties
---------------

The :py:class:`ooodev.format.calc.modify.cell.alignment.Properties` class sets the text properties of a style.

Setup
^^^^^

General setup for the examples in this section.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 15, 16, 17, 18, 19, 20, 21

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.cell.alignment import Properties
        from ooodev.format.calc.modify.cell.alignment import TextDirectionKind, StyleCellKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                style = Properties(
                    wrap_auto=True,
                    hyphen_active=True,
                    direction=TextDirectionKind.PAGE,
                    style_name=StyleCellKind.DEFAULT,
                )
                style.apply(doc)

                style_obj = Properties.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Setting the text properties
^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        style = Properties(
            wrap_auto=True,
            hyphen_active=True,
            direction=TextDirectionKind.PAGE,
            style_name=StyleCellKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236455445-a2c0afff-0c10-4fb5-8daf-930dd05cc953`.

.. cssclass:: screen_shot

    .. _236455445-a2c0afff-0c10-4fb5-8daf-930dd05cc953:

    .. figure:: https://user-images.githubusercontent.com/4193389/236455445-a2c0afff-0c10-4fb5-8daf-930dd05cc953.png
        :alt: Calc dialog style Text Properties modified
        :figclass: align-center
        :width: 450px

        Calc dialog style Text Properties modified

Getting the text properties from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        style_obj = Properties.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
        assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_alignment`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.modify.cell.alignment.TextAlign`
        - :py:class:`ooodev.format.calc.modify.cell.alignment.TextOrientation`
        - :py:class:`ooodev.format.calc.modify.cell.alignment.Properties`