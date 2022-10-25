# Make Slides

This is not working.

In order to build this example it would be necessary to use 3rd party libraries
to build a module that could control Draw ( and other windows ).

For this example explicitly inserts Custom Shapes from Draw and places them on
the main document. The ONLY way to do this is to Use a `dispatch_command` to activate the
custom draw item on the tool bar followed by automating the mouse to draw the shape in the document.

What ever third party libraries are used would have to be able to find and focus Draw window and
Move the mouse to specific ( pre-determined ) coordinates, then preform a click and drag.

[pynput](https://pynput.readthedocs.io/en/latest/index.html) has some potential with controlling the mouse cross platform. But has Limitations using on Linux WayLand. ( `uintput`, must be set to `chmod 0666` or script runs as `root`).

It is not necessary to be precise putting the custom shape in the document. `set_size()` and
`set_position()` can be called after wards.
