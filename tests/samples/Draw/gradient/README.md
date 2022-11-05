# Draw Gradient Examples

![gradients_transparent_sm](https://user-images.githubusercontent.com/4193389/199873235-517287a4-7514-4108-a6a3-2bb6d768e3ca.png)


Demonstrates several type of Draw gradients.

- Fill: Rectangle filled with solid color ( fill color).
- Fill Gradient: Rectangle filled with named gradient ( gradient color).
- Fill Gradient custom Props:  Rectangle filled with named gradient that has custom properties of angle, start color, and end color ( gradient color Custom props).
- Fill Gradient common color: Rectangle filled with custom start and stop colors (gradient CommonColor).
- Fill Hatching: Rectangle filled with specified hatch pattern ( hatching color).
- Fill Bitmap: Rectangle filled with specified Draw bitmap( bitmap color).
- Fill Custom Bitmap: Rectangle Filled with a custom image (bitmap file).

## Command line options

There are many options for this example due to all the options available for gradients.

Color names are flexible but must be an integer value representing a color or a name matching [CommonColor]

### Common

- `-x` `--x-pos`: Optional - Rectangle X position
- `-Y` `--y-pos`: Optional - Rectangle Y position
- `-w` `--width`: Optional - Rectangle Width
- `-t` `--height`: Optional - Rectangle Height

### `--Kind` fill

Creates *fill color*

- `-s` `--start-color`: Optional - Color of the fill.

### `--Kind` name

Creates *gradient color*

- `--gradient-kind`: Optional - Kind of Draw gradient. Must be one of the choices.

### `--Kind` name_props

Creates *gradient color Custom props*

- `--gradient-kind`: Optional - Kind of Draw gradient. Must be one of the choices.
- `-a` `--angle`: Optional - Angle of the gradient.
- `-s` `--start-color`: Optional - Start color of the fill.
- `-e` `--end-color`: Optional - End color of fill.

### `--Kind` gradient

Creates *gradient Common color*

- `-a` `--angle`: Optional - Angle of the gradient.
- `-s` `--start-color`: Optional - Start color of the fill.
- `-e` `--end-color`: Optional - End color of fill.


### `--Kind` hatch

Creates *hatching color*

- `--hatch-kind`: Optional - Kind of Draw hatch gradient. Must be one of the choices.

### `--Kind` bitmap

Creates *bitmap color*

- `--bitmap-kind`: Optional - Kind of Draw bitmap gradient. Must be one of the choices.

### `--Kind` file

Creates *bitmap file*

- NO extra choices

### Example

- `python -m start -k name --gradient-kind pastel_bouquet`
- `python -m start -k gradient -a 45`
- `python -m start -k gradient -s dark-red -e lime-green`
- `python -m start -k name_props --gradient-kind sunshine -a 88 -s "dark violet" -e light_goldenrod_yellow`

## Automate

A message box is display once the document has been created asking if you want to close the document.

### Cross Platform

From current example folder.

```sh
python -m start -h
```

### Linux/Mac

```sh
python ./tests/samples/Draw/gradient/start.py -h
```

### Windows

```ps
python .\tests\samples\Draw\gradient\start.py -h
```

[CommonColor]: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/color.html#ooodev.utils.color.CommonColor
