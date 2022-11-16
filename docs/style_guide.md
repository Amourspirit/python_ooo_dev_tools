# ODEV Sytle Guide

## Book Naming conventions

### References

#### Chapter, ch01

`ch` followed by two digit chapter number

Example:

```rst
.. _ch01:
```

#### Sub Section

``ch`` followed by two digits folowed by `_` followed by section short name.

Example:

```rst
.. _ch03_examine_office:
```

#### Sub, sub section

sub section followed by `_` followed by sub section short name

Example:

```rst
.. _ch03_examine_office_cofig_prop:
```

### Chapter figures

chapter name followed by fig followed by `_` followed by short name

Example:

```rst
.. _ch03fig_prop_dialog:
```

### Chapter tables

chapter name followed by tbl followed by `_` followed by short name

Example:

```rst
.. _ch02tbl_some_doc_prop:
```

### Note

To reference a chapter or section use ref

Example:

```rst
:ref:`ch01`
:ref:`ch03_examine_office`
```

To reference figures and tables use numref

Example:

```rst
:numref:`ch03fig_detail_prop_lst`
```

## Headings

- `#` with overline, for parts
- `*` with overline, for chapters
- `=`, for sections
- `-`, for subsections
- `^`, for subsubsections
- `“`, for paragraphs

See Also:
<https://thomas-cokelaer.info/tutorials/sphinx/rest_syntax.html#headings>
<https://documentation-style-guide-sphinx.readthedocs.io/en/latest/index.html>
<https://www.sphinx-doc.org/en/master/usage/restructuredtext/basics.html#sections>


## List

By default rdt-theme displays un-ordered list without any list type formating.
To create an un-ordered list with bullets us the `ul-list` class from `readthedocs_custom.css`

```rst
.. cssclass:: ul-list

    * simple shapes: line, ellipse, rectangle, text;
    * shape fills: solid, gradients, hatching, bitmaps;
    * an OLE shape (a math formulae);
    * polygons, multiple lines, partial ellipses.
```

## Image Classes

- `screen_shot`: any still image that is a screen shot type image
- `invert`: any image that can can have color inverted such as jpeg, png for dark theme.
- `diagram`: diagram classes
- `a_gif`: animated gif images

**Notes:**

The cssclass directive is use to assign classes to images.
Screen shots typically have class `screen_shot`, `invert`

**Example:**

```rst
.. cssclass:: screen_shot invert

    .. _ch03fig_detail_prop_lst:
    .. figure:: https://user-images.githubusercontent.com/4193389/179298066-7acaa668-7b0b-4a59-bbb8-407ba354bf8a.png
        :alt: Details Properties List for algs.odp
        :figclass: align-center

        :Details Properties List for ``algs.odp``.
```

Diagrams typically have class `diagram`, `invert`

**Example:**

```rst
.. cssclass:: diagram invert

    .. _ch03fig_peek_services_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/179381798-efcb4f4a-a877-469f-9c6e-033e9cf7fe6b.png
        :alt: Methods to Investigate the Service and Interface Relationships and Hierarchies
        :figclass: align-center

        :Methods to Investigate the Service and Interface Relationships and Hierarchies.
```

Animated gif typically have class `a_gif`
Note that animated gif does not invert class.
Animated gif is not to be darkened using css styles

**Example:**

```rst
.. cssclass:: a_gif

    .. _ch02fig_lo_qi_auto_demo:
    .. figure:: https://user-images.githubusercontent.com/4193389/178285134-70b9aa56-5eaa-43c8-aa59-c19f2b495336.gif
        :alt: Lo.qi autocomplete demo image
        :figclass: align-center

        :Lo.qi autocomplete demo
```

## Spelling

To run spell check from `docs/` command line

```ps
docs\> sphinx-build -b spelling . _build
```

In some case accepted spelling of words need to added to a single rst file.
For one off words `:spelling:word` can be use for example: ``:spelling:word:`László` `` or ``:spelling:word:`Németh` ``

Adding a list or words to an rst file:
This can be done using `spelling:word-list::` directive. as shown below.

See src/conn/index.rst as an example

```rst
.. spelling:word-list::
    conn
```

## Abbreviations

Abbreviations are handled by the `:abbreviation:` role.

``:abbreviation:`i.e.` ``

## TODO

Docs that still need work such linking to chapter not yet created are to use the following directive.

```rst
.. todo::

    | Chapter 5, Add link to chapters 7
    | Chapter 5, Add link to chapters 8
```

To see the todo set `todo_include_todos = True` in `conf.py` and regenerate docs.
A master list of todo's will be on bottom of main page, also each document that contains a todo will disply it on its doc page.

See Also: [sphinx.ext.todo](https://www.sphinx-doc.org/en/master/usage/extensions/todo.html#module-sphinx.ext.todo)



## Comments

Adding comments to doc is straight forward. Use `..` followed by comment on a new intented line.
In this example adding a comment for a diagram.

```rst
..
    Figure 3

.. cssclass:: diagram invert

    .. _ch03fig_peek_services_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/179381798-efcb4f4a-a877-469f-9c6e-033e9cf7fe6b.png
        :alt: Methods to Investigate the Service and Interface Relationships and Hierarchies
        :figclass: align-center

        :Methods to Investigate the Service and Interface Relationships and Hierarchies.
```

## Linking to Source code

[sphinx.ext.extlinks](https://documentation.help/Sphinx/extlinks.html), see setting for extlinks in conf.py
Allows for custom roles to be set up the make for simple source code link injection.

Linking source code standart for directives such as `.. seealso`

`src-link` css class handles adding icon at end of link and formatting of `ul` lists.

```rst
.. seealso::

    .. cssclass:: src-link

        - :odev_src_draw_meth:`get_slide`
        - :odev_src_draw_meth:`get_slides`
        - :odev_src_draw_meth:`get_slides_count`
```

**Note:**

`sphinx.ext.extlinks` by defalult creates external links.

The `custom.js` file converts source code links back to internal links.


To see what role are available or to add new rows see extlinks in conf.py

## External Sources via sphinx.ext.intersphinx

Use `external+mapping` name.

In `conf.py`

```python
intersphinx_mapping = {"odevguiwin": (odevgui_win_url, None)}
```

**Example:**

```rst
:external+odevguiwin:ref:`class_robot_keys`
:external+odevguiwin:py:meth:`odevgui_win.draw_dispatcher.DrawDispatcher.create_dispatch_shape`
```

## Code Blocks

### Tabs

Code blocks are generally in a tab.
See Also: <https://sphinx-tabs.readthedocs.io/en/latest/>

```rst
.. tabs::

    .. code-tab:: python

        def _toggle_side_bar(self) -> None:
            if not RobotKeys:
                Lo.print("odevgui_win not found.")
                return
            RobotKeys.send_current(SendKeyInfo(WriterKeyCodes.KB_SIDE_BAR))
```

Hilighting can be added to lines by using `:emphasize-lines:`

```rst
.. tabs::

    .. code-tab:: python
        :emphasize-lines: 2, 3, 4, 5

        def _toggle_side_bar(self) -> None:
            if not RobotKeys:
                Lo.print("odevgui_win not found.")
                return
            RobotKeys.send_current(SendKeyInfo(WriterKeyCodes.KB_SIDE_BAR))
```

### Tab None

Adding a None Tab.
This option should only be used in conjunction with code tabs.
If one code tab on a page gets this option then all code tabs on the page should get this option.
If a None Tab is used this should be prefixed by `.. cssclass:: tab-none` for formatting purposes.

`.. only:: html` is used to indicate to Sphinx that this tab is to only be renedred for html.

```rst
.. tabs::

    .. code-tab:: python

        print("Hello World")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

```

## Collapse sections

Creates a generic collape section

```rst
.. collapse:: Some Text

    Congrats, you have revealed hidden text
```

Creates expanded collapse section

```rst
.. collapse:: Some Text
    :open:

    Congrats, you have revealed hidden text
```

Example has `.. code::` directive.

```rst
.. collapse:: Example

    .. code::

        special = get_my_code()
```

Creates collapse section with special css formatting applied, mostly for spacing

```rst
.. cssclass:: rst-collapse

    .. collapse:: Some Title

        Fancy stuff
```

## Other

[sphinx_design](https://sphinx-design.readthedocs.io/en/rtd-theme/get_started.html)
Allows setting icon ``:octicon:`code-square;1em;sd-text-info` ``
