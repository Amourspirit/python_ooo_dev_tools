# Impress Basic Slide Show

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198407936-7865b1c2-75b7-4530-8598-a1ce52821752.png" width="448" height="448">
</p>

Demonstrates displaying a slide show using default `algs.odp` file.

A message box is display once the slide show has ended asking if you want to close the document.

## Automate

A single parameters can be passed in which is the slide show document to display:

**Example:**

```sh
python ./tests/samples/Impress/Basic_Show/start.py "tests/fixtures/presentation/algs.odp"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Basic_Show/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Basic_Show\start.py
```
