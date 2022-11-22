.. tabs::

    .. code-tab:: python

        # With Formatting
        >>> data = Calc.get_array(sheet=sheet, range_name="A1:E10")
        >>> fl = FormatterTable(format=(".2f", ">9"), idxs=(0, 9))
        >>> fl.custom_formats_row.append(FormatTableItem(format=">9", idxs_inc=(0, 9)))
        >>> fl.col_formats.append(FormatTableItem(format=(".0f", "^9"), idxs_inc=(0,), row_idxs_exc=(0, 9)))
        >>> fl.col_formats.append(FormatTableItem(format=(".0%", ">9"), idxs_inc=(4,), row_idxs_exc=(0, 9)))
        >>> Calc.print_array(data, fl)

        Row x Column size: 10 x 5
        Stud. No.   Proj/20    Mid/35    Fin/45    Total%
          22001       16.46     30.92     37.01       84%
          22028       11.88     23.04     25.46       60%
          22048       13.96     19.25     25.99       59%
          23715       12.08     18.67     20.48       51%
          23723       17.29     27.71     36.23       81%
          24277        0.00     16.04     19.69       36%
                      11.94     22.60     27.48       62%
                       0.60      0.65      0.61
                    Proj/20    Mid/35    Fin/45    Total%

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None
