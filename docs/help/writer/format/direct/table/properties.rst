.. _help_writer_format_direct_table_properties:

Write Direct Table TableProperties Class
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

The :py:class:`ooodev.format.writer.direct.table.properties.TableProperties` is used to set the properties of a table.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.table.properties import TableProperties, TableAlignKind
        from ooodev.office.write import Write
        from ooodev.units import UnitMM
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.intensity import Intensity
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.table_helper import TableHelper


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)
                cursor = Write.get_cursor(doc)

                tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)
                props_style = TableProperties(name="My_Table", relative=False, align=TableAlignKind.AUTO)

                table = Write.add_table(cursor=cursor, table_data=tbl_data, first_row_header=False, styles=[props_style])

                # getting the table properties
                tbl_props_style = TableProperties.from_obj(table)
                assert tbl_props_style.prop_name == "My_Table"

                Lo.delay(1_000)
                Lo.close_doc(doc)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Examples
--------

Absolute Position
+++++++++++++++++

Auto Position
^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(name="My_Table", relative=False, align=TableAlignKind.AUTO)
        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234003140-30186f30-d7d2-4f92-96e5-22f41e1af410:
    .. figure:: https://user-images.githubusercontent.com/4193389/234003140-30186f30-d7d2-4f92-96e5-22f41e1af410.png
        :alt: Auto Absolute Position
        :figclass: align-center
        :width: 520px

        Auto Absolute Position


.. cssclass:: screen_shot

    .. _234008850-27f1a75e-1b5b-414d-baf6-2d02ad83175f:
    .. figure:: https://user-images.githubusercontent.com/4193389/234008850-27f1a75e-1b5b-414d-baf6-2d02ad83175f.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Center Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.CENTER,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234004582-55ff4d13-ef74-41bb-9adf-e0ee3598ab55:
    .. figure:: https://user-images.githubusercontent.com/4193389/234004582-55ff4d13-ef74-41bb-9adf-e0ee3598ab55.png
        :alt: Align Center Position setting width
        :figclass: align-center
        :width: 520px

        Align Center Position setting width

.. cssclass:: screen_shot

    .. _234007765-6a6739f1-e6c1-4ae6-bb23-63f1ad4d89a3:
    .. figure:: https://user-images.githubusercontent.com/4193389/234007765-6a6739f1-e6c1-4ae6-bb23-63f1ad4d89a3.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Center Position setting left
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.CENTER,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            left=UnitMM(40.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234010297-d3cfaf1b-5037-47c0-820f-93ba0d6503ad:
    .. figure:: https://user-images.githubusercontent.com/4193389/234010297-d3cfaf1b-5037-47c0-820f-93ba0d6503ad.png
        :alt: Align Center Position setting left
        :figclass: align-center
        :width: 520px

        Align Center Position setting left

.. cssclass:: screen_shot

    .. _234010486-5e55d435-f382-4b31-87d0-7182d31752a9:
    .. figure:: https://user-images.githubusercontent.com/4193389/234010486-5e55d435-f382-4b31-87d0-7182d31752a9.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align From Left Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.FROM_LEFT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234015085-2bfcec71-e0a7-4e6c-9051-f67ab94e7948:
    .. figure:: https://user-images.githubusercontent.com/4193389/234015085-2bfcec71-e0a7-4e6c-9051-f67ab94e7948.png
        :alt: Align From Left Position setting width
        :figclass: align-center
        :width: 520px

        Align From Left Position setting width

.. cssclass:: screen_shot

    .. _234015381-e1e7bca8-be23-4a04-9ad7-7a380a4006ee:
    .. figure:: https://user-images.githubusercontent.com/4193389/234015381-e1e7bca8-be23-4a04-9ad7-7a380a4006ee.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Left Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.LEFT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234017855-2732b540-8d59-4b16-9c27-70fd13c0ae4b:
    .. figure:: https://user-images.githubusercontent.com/4193389/234017855-2732b540-8d59-4b16-9c27-70fd13c0ae4b.png
        :alt: Align Left Position setting width
        :figclass: align-center
        :width: 520px

        Align Left Position setting width

.. cssclass:: screen_shot

    .. _234018068-7b4b7329-f423-4ca2-af02-3c248bd1ff0f:
    .. figure:: https://user-images.githubusercontent.com/4193389/234018068-7b4b7329-f423-4ca2-af02-3c248bd1ff0f.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Left Position setting right
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.LEFT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            right=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234019567-037f2a71-cadf-4da9-8e4e-69d0c0a17ffb:
    .. figure:: https://user-images.githubusercontent.com/4193389/234019567-037f2a71-cadf-4da9-8e4e-69d0c0a17ffb.png
        :alt: Align Left Position setting right
        :figclass: align-center
        :width: 520px

        Align Left Position setting right

.. cssclass:: screen_shot

    .. _234019807-38ce580c-e57a-4b9c-8665-c473183fdabf:
    .. figure:: https://user-images.githubusercontent.com/4193389/234019807-38ce580c-e57a-4b9c-8665-c473183fdabf.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Right Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.RIGHT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234021353-e58bdb52-c5fb-4376-b928-390d59254022:
    .. figure:: https://user-images.githubusercontent.com/4193389/234021353-e58bdb52-c5fb-4376-b928-390d59254022.png
        :alt: Align Right Position setting width
        :figclass: align-center
        :width: 520px

        Align Right Position setting width

.. cssclass:: screen_shot

    .. _234021566-72dd687b-10e1-48d5-be3b-3826d4044313:
    .. figure:: https://user-images.githubusercontent.com/4193389/234021566-72dd687b-10e1-48d5-be3b-3826d4044313.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Right Position setting left
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.RIGHT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            left=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234022582-9e90ed0f-619a-40d9-b2ae-e373eb574051:
    .. figure:: https://user-images.githubusercontent.com/4193389/234022582-9e90ed0f-619a-40d9-b2ae-e373eb574051.png
        :alt: Align Right Position setting left
        :figclass: align-center
        :width: 520px

        Align Right Position setting left

.. cssclass:: screen_shot

    .. _234022939-ccd1e7e6-fb57-4881-af3f-5edcbb63d121:
    .. figure:: https://user-images.githubusercontent.com/4193389/234022939-ccd1e7e6-fb57-4881-af3f-5edcbb63d121.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Manual Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.MANUAL,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=UnitMM(60.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234023933-1c2041c5-5ee4-4312-bbab-94433373b16e:
    .. figure:: https://user-images.githubusercontent.com/4193389/234023933-1c2041c5-5ee4-4312-bbab-94433373b16e.png
        :alt: Align Manual Position setting width
        :figclass: align-center
        :width: 520px

        Align Manual Position setting width

.. cssclass:: screen_shot

    .. _234024282-797d5d09-2e86-485a-8a40-7cf92819229f:
    .. figure:: https://user-images.githubusercontent.com/4193389/234024282-797d5d09-2e86-485a-8a40-7cf92819229f.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Manual Position setting left & right
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=False,
            align=TableAlignKind.MANUAL,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            left=UnitMM(66.0),
            right=UnitMM(55.0),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234025419-16b043c9-972a-4a60-84a8-b7a4d2c431a2:
    .. figure:: https://user-images.githubusercontent.com/4193389/234025419-16b043c9-972a-4a60-84a8-b7a4d2c431a2.png
        :alt: Align Manual Position setting left & right
        :figclass: align-center
        :width: 520px

        Align Manual Position setting left & right

.. cssclass:: screen_shot

    .. _234025674-1985e1d3-381d-421b-b866-1c2320471a93:
    .. figure:: https://user-images.githubusercontent.com/4193389/234025674-1985e1d3-381d-421b-b866-1c2320471a93.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Relative Position
+++++++++++++++++

Align From Left Position setting left & width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=True,
            align=TableAlignKind.FROM_LEFT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            left=Intensity(20),
            width=Intensity(40),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234028263-95e62781-bc16-47e6-87ec-bbcf0b44bf89:
    .. figure:: https://user-images.githubusercontent.com/4193389/234028263-95e62781-bc16-47e6-87ec-bbcf0b44bf89.png
        :alt: Align Relative From Left Position setting left & width
        :figclass: align-center
        :width: 520px

        Align Relative From Left Position setting left & width

.. cssclass:: screen_shot

    .. _234028594-abe5737e-a4a6-4b1c-81fa-15a9522263b9:
    .. figure:: https://user-images.githubusercontent.com/4193389/234028594-abe5737e-a4a6-4b1c-81fa-15a9522263b9.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Left Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=True,
            align=TableAlignKind.LEFT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=Intensity(40),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234030209-03ad18c7-a193-43a2-b9f2-30b54c56bbdb:
    .. figure:: https://user-images.githubusercontent.com/4193389/234030209-03ad18c7-a193-43a2-b9f2-30b54c56bbdb.png
        :alt: Align Relative Left Position setting width
        :figclass: align-center
        :width: 520px

        Align Relative Left Position setting width

.. cssclass:: screen_shot

    .. _234030349-c2a9d533-9b08-4a23-a6e5-bc0ce85738dc:
    .. figure:: https://user-images.githubusercontent.com/4193389/234030349-c2a9d533-9b08-4a23-a6e5-bc0ce85738dc.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Left Position setting right
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=True,
            align=TableAlignKind.LEFT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            right=Intensity(40),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234031775-668229d6-2473-4fb5-885e-db1d5e4eee11:
    .. figure:: https://user-images.githubusercontent.com/4193389/234031775-668229d6-2473-4fb5-885e-db1d5e4eee11.png
        :alt: Align Relative Left Position setting right
        :figclass: align-center
        :width: 520px

        Align Relative Left Position setting right

.. cssclass:: screen_shot

    .. _234032176-c20e2da3-aa35-4f27-bd23-e2e3debe0fec:
    .. figure:: https://user-images.githubusercontent.com/4193389/234032176-c20e2da3-aa35-4f27-bd23-e2e3debe0fec.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Right Position setting width
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=True,
            align=TableAlignKind.RIGHT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            width=Intensity(40),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234033186-6f33f2fa-3e0e-4b50-ab62-942525c0724f:
    .. figure:: https://user-images.githubusercontent.com/4193389/234033186-6f33f2fa-3e0e-4b50-ab62-942525c0724f.png
        :alt: Align Relative Right Position setting width
        :figclass: align-center
        :width: 520px

        Align Relative Right Position setting width

.. cssclass:: screen_shot

    .. _234033630-ffc32292-baf3-4f06-a4ee-0ec403c85e34:
    .. figure:: https://user-images.githubusercontent.com/4193389/234033630-ffc32292-baf3-4f06-a4ee-0ec403c85e34.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Align Right Position setting left
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        props_style = TableProperties(
            name="My_Table",
            relative=True,
            align=TableAlignKind.RIGHT,
            above=UnitMM(2.0),
            below=UnitMM(1.8),
            left=Intensity(40),
        )

        table = Write.add_table(
            cursor=cursor,
            table_data=tbl_data,
            first_row_header=False,
            styles=[props_style],
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _234034670-07baf073-58ce-49bb-81c2-d56129158a93:
    .. figure:: https://user-images.githubusercontent.com/4193389/234034670-07baf073-58ce-49bb-81c2-d56129158a93.png
        :alt: Align Relative Right Position setting left
        :figclass: align-center
        :width: 520px

        Align Relative Right Position setting left

.. cssclass:: screen_shot

    .. _234034892-df0029a6-5935-4d06-9234-2fe113ca9806:
    .. figure:: https://user-images.githubusercontent.com/4193389/234034892-df0029a6-5935-4d06-9234-2fe113ca9806.png
        :alt: Table Properties Dialog
        :figclass: align-center
        :width: 450px

        Table Properties Dialog

Getting the Properties from the table
+++++++++++++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python

        # ... other code
        # getting the table properties
        tbl_props_style = TableProperties.from_obj(table)
        assert tbl_props_style.prop_name == "My_Table"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_table_borders`
        - :ref:`help_writer_format_direct_table_background`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_borders`
        - :py:meth:`Write.add_table() <ooodev.office.write.Write.add_table>`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.table.properties.TableProperties`