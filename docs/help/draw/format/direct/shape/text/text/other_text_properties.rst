.. _help_draw_format_direct_shape_text_text_other_prop:

Draw Direct Shape Text - Other Properties
=========================================

Other properties of the shape text can be set.
The properties that can be set will depend on the type of the shape.

Here are some examples:

.. tabs::

    .. code-tab:: python

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)

        # Adjust to contour option
        rect.component.TextContourFrame = True  

        # Fit to Frame option
        # TextFitToSizeType.PROPORTIONAL (True)
        # TextFitToSizeType.AUTOFIT (True)
        # from ooo.dyn.drawing.text_fit_to_size_type import TextFitToSizeType
        rect.component.TextFitToSize = TextFitToSizeType.PROPORTIONAL

        # Word Wrap option
        # Word wrap text in shape.
        # TextWordWrap does not seem to be listed in the API
        rect.component.TextWordWrap = True

        # resize shape to fit text
        rect.component.TextAutoGrowHeight = True
        rect.component.TextAutoGrowWidth = True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

See the `API - TextProperties Member List <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1TextProperties-members.html>`__ for more options.