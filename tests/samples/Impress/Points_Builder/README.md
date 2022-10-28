# Impress read text data and convert to slides with points

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198420085-ec183c1f-94a7-47ec-9b75-ec44804a94be.png">
</p>

Convert a text file of points into a series of slides. Uses a template from Office.

A message box is display once the document has been created asking if you want to close the document.

## Automate

A single parameters can be passed in which is the slide show document to data read from:

**Example:**

```sh
python ./tests/samples/Impress/Points_Builder/start.py "tests/fixtures/data/pointsInfo.txt"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Points_Builder/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Points_Builder\start.py
```
