# Draw Bezier Curve

Demonstrates reading a text file contains Bezier curve data that is recreated in a Draw document.

A message box is display once the document has been created asking if you want to close the document.

## Automate

An extra parameter can be passed in:

A value between `0` and `3` The default value is `2`.
Each value represents a different Bezier curve file.

**Example:**

```sh
python -m start 1
```

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Draw/Bezier_Builder/start.py
```

### Windows

```ps
python .\tests\samples\Draw\Bezier_Builder\start.py
```
