# Custom Slide Show

Demonstrates opening a presentation file in Impress starting a slide show with only the slide indexes passes in.

A message box is display once slide show has completed asking if you want to close document.

## Automate

Slide index numbers can be passed in.

**Example:**

```sh
python tests/samples/Draw/Custom_Show/start.py 5 6 7 8
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux

```sh
python ./tests/samples/Draw/Custom_Show/start.py
```

### Windows

```ps
python .\tests\samples\Draw\Custom_Show\start.py
```
