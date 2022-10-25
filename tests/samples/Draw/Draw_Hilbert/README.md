# Draw Hilbert curve

![img](https://upload.wikimedia.org/wikipedia/commons/0/06/Hilbert_curve_3.svg)

Generate a [Hilbert curve] of the specified level.

Created using a series of rounded blue lines.
Position/size the window, resize the page view

## Usage

run `python -m draw_hilbert 4`
Using `6` takes  2+ minutes to fully draw. It is fun to try once.
Using `7` causes the code to mis-calculate, so the line drawing goes off the left side of the canvas.
And it takes forever to do it.

## Automate

An extra parameter can be passed in:

An integer value the determines the levels to draw [Hilbert curve]. The default value is `4`.

**Example:**

```sh
python -m start 4
```

### Cross Platform

From current example folder.

```sh
python -m start 4
```

### Linux/Mac

```sh
python ./tests/samples/Draw/Draw_Hilbert/start.py 4
```

### Windows

```ps
python .\tests\samples\Draw\Draw_Hilbert\start.py 4
```

[Hilbert curve]: https://en.wikipedia.org/wiki/Hilbert_curve
