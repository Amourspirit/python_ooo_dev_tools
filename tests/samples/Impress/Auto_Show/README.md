# Impress Auto Slide Show

Demonstrates displaying a slide show that automatically plays using default `algs.odp` file.

A message box is display once the slide show has ended asking if you want to close the document.

## Automate

A single parameters can be passed in which is the slide show document to modify:

**Example:**

```sh
python ./tests/samples/Impress/Auto_Show/start.py "tests/fixtures/presentation/algs.odp"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Auto_Show/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Auto_Show\start.py
```
