.. _help_writer_format_direct_cursor_char_styler:

Styling with Cursor and CharacterStyler in Writer
=================================================

Introduction
------------

In version `0.30.1` the :py:class:`ooodev.write.style.WriteCharacterStyle` class was added.
It can be added via :py:attr:`ooodev.write.WriteTextCursor.style_direct_char`.

When styling using a cursor the style set the properties of the current select (text range) range of the cursor.
This means that cursor must be reset before adding new text. The :py:class:`~ooodev.write.style.direct.character_styler.CharacterStyler` has a
:py:meth:`~ooodev.write.style.direct.character_styler.CharacterStyler.clear` method that resets the cursor.

The simplest way is to set the select the text that is to be styled. Apply styles, Move curse back to end and then reset the cursor.
Resetting the cursor would only be necessary if end was selected in the styling; Otherwise, the last character properties would not change when style is applied.

The examples below all output the same result.

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f3fc1116-ec34-4a55-a7b1-2dd384a532c7
    :alt: Example of styled text

Example setting and clearing
----------------------------

.. code-block:: python

    from __future__ import annotations
    from ooodev.write import WriteDoc
    from ooodev.utils.color import StandardColor
    from ooodev.format.writer.direct.char.font import FontScriptKind
    from ooodev.format.writer.direct.char.borders import BorderLineKind

    # ... other code
    doc = WriteDoc.create_doc(visible=True)
    cursor = doc.get_cursor()

    # make an alias (shortcut) to make clearing easier
    clr = cursor.style_direct_char.clear

    cursor.append("Hello")
    # go left and select the the last word appended
    cursor.go_left(5, True)

    # Style the Hello word to be Font size 30, bold, italic, underline and color blue.
    cursor.style_direct_char.style_font_general(
        size=30.0,
        b=True,
        i=True,
        u=True,
        color=StandardColor.BLUE,
    )
    # put a border around Hello
    cursor.style_direct_char.style_borders_side(
        line=BorderLineKind.DOUBLE_THIN,
        color=StandardColor.RED,
    )

    # put the cursor back to the end so we can append more text.
    cursor.goto_end()
    # clear the cursor character styles.
    clr() # Alias of cursor.style_direct_char.clear() for brevity

    # append world without any character styling
    cursor.append(" World")
    # make the d in World subscript

    cursor.go_left(1, True) # select the d character
    style = cursor.style_direct_char.style_font_position(script_kind=FontScriptKind.SUBSCRIPT)
    # reset cursor
    cursor.goto_end()
    clr()

    cursor.append(". Nice Day!") # unformatted characters

Automatic resetting of the characters for the cursor is also possible by using events.
Each time the `CharacterStyler` set a character style it raises a ``WriteNamedEvent.CHARACTER_STYLE_APPLYING`` and a  ``WriteNamedEvent.CHARACTER_STYLE_APPLIED`` event.
By subscribing to the ``CHARACTER_STYLE_APPLIED`` event we can take action when the event occurs.

Following the above example we will add an event method that puts the cursor back to the end and clear character formatting when a character styled event occurs.

.. code-block:: python

    from __future__ import annotations
    from typing import cast, TYPE_CHECKING
    from ooodev.write import WriteDoc
    from ooodev.utils.color import StandardColor
    from ooodev.format.writer.direct.char.font import FontScriptKind
    from ooodev.format.writer.direct.char.borders import BorderLineKind
    from ooodev.write.style.direct.character_styler import CharacterStyler

    if TYPE_CHECKING:
        from ooodev.events.args.event_args import EventArgs

    if TYPE_CHECKING:
        from com.sun.star.text import XTextCursor

    def on_char_style_applied(src: Any, event_args: EventArgs) -> None:
        styler = cast(CharacterStyler, src)
        cursor = cast("XTextCursor", event_args.event_data.get("this_component", None))
        if cursor is None:
            return
        cursor.gotoEnd(False)
        styler.clear()
        

    # ... other code
    doc = WriteDoc.create_doc(visible=True)
    cursor = doc.get_cursor()

    # subscribe to the event that resets the cursor when a style is applied.
    cursor.subscribe_event(WriteNamedEvent.CHARACTER_STYLE_APPLIED, on_char_style_applied)

    cursor.append("Hello")
    # go left and select the the last word appended
    cursor.go_left(5, True)

    # Style the Hello word to be Font size 30, bold, italic, underline and color blue.
    _ = cursor.style_direct_char.style_font_general(
        size=30.0,
        b=True,
        i=True,
        u=True,
        color=StandardColor.BLUE,
    )
    # because the cursor position will be reset in the event need to select again.
    # A solution for this below.
    cursor.go_left(5, True)
    # put a border around Hello
    _ = cursor.style_direct_char.style_borders_side(
        line=BorderLineKind.DOUBLE_THIN,
        color=StandardColor.RED,
    )

    # append world without any character styling
    # note that there was no need to reset the cursor. It was done in the event.
    cursor.append(" World")
    # make the d in World subscript

    cursor.go_left(1, True) # select the d character
    _ = cursor.style_direct_char.style_font_position(script_kind=FontScriptKind.SUBSCRIPT)
    cursor.append(". Nice Day!") # unformatted characters

As see in the above example is was necessary to select ``Hello`` twice via ``cursor.go_left(5, True)``
because the event is resetting the cursor.

To solve this issue we can use the ``extra_data`` property of ``CharacterStyler``.
This is key, value data  (dict) that can be used to store and read extra information as needed.
We can access this in the event as well.

.. code-block:: python

    from __future__ import annotations
    from typing import cast, TYPE_CHECKING
    from ooodev.write import WriteDoc
    from ooodev.utils.color import StandardColor
    from ooodev.format.writer.direct.char.font import FontScriptKind
    from ooodev.format.writer.direct.char.borders import BorderLineKind
    from ooodev.write.style.direct.character_styler import CharacterStyler
    
    if TYPE_CHECKING:
        from com.sun.star.text import XTextCursor
        from ooodev.events.args.event_args import EventArgs

    def on_char_style_applied(src: Any, event_args: EventArgs) -> None:
        styler = cast(CharacterStyler, src)
        skip = styler.extra_data.get("skip", False)
        if skip:
            return
        cursor = cast("XTextCursor", event_args.event_data.get("this_component", None))
        if cursor is None:
            return
        cursor.gotoEnd(False)
        styler.clear()
        

    # ... other code
    doc = WriteDoc.create_doc(visible=True)
    cursor = doc.get_cursor()

    # subscribe to the event that resets the cursor when a style is applied.
    cursor.subscribe_event(WriteNamedEvent.CHARACTER_STYLE_APPLIED, on_char_style_applied)

    cursor.append("Hello")
    # go left and select the the last word appended
    cursor.go_left(5, True)

    # set extra data the we can use in event.
    # don't want to reset the cursor until after border is set around Hello
    cursor.style_direct_char.extra_data["skip"] = True
    # Style the Hello word to be Font size 30, bold, italic, underline and color blue.
    _ = cursor.style_direct_char.style_font_general(
        size=30.0,
        b=True,
        i=True,
        u=True,
        color=StandardColor.BLUE,
    )
    # set custom flag to let event reset cursor
    cursor.style_direct_char.extra_data["skip"] = False
    # put a border around Hello
    _ = cursor.style_direct_char.style_borders_side(
        line=BorderLineKind.DOUBLE_THIN,
        color=StandardColor.RED,
    )

    # append world without any character styling
    # note that there was no need to reset the cursor. It was done in the event.
    cursor.append(" World")
    # make the d in World subscript

    cursor.go_left(1, True) # select the d character
    _ = cursor.style_direct_char.style_font_position(script_kind=FontScriptKind.SUBSCRIPT)
    cursor.append(". Nice Day!") # unformatted characters