# Impress copy Slide

Demonstrates opening a presentation file in Impress and copying a slide from a given index to after another given index.

A message box is display once the document has been created asking if you want to close the document.

## Automate

Three parameters can be passed in:

1. File Name: Such as `"tests/fixtures/presentation/algs.odp"`
2. From Index: Such as `2`
3. To Index: Such as `4`

**Example:**

```sh
python ./tests/samples/Impress/Copy_Slide/start.py "tests/fixtures/presentation/algs.odp" 0 2
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Copy_Slide/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Copy_Slide\start.py
```
