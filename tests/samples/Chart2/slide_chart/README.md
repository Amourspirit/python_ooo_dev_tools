# Slide Chart

![slide_chart](https://user-images.githubusercontent.com/4193389/198894178-1c6b79bf-185f-44e0-b061-3c026da88384.png)

Generates a column chart using the "Sneakers Sold this Month" table from `chartsData.ods`, copies it to the clipboard, and closes the spreadsheet. Then an Impress document is created, and the chart image is pasted into it.

Also demonstrates saving the chart as an image.

A message box is display once the document has been created asking if you want to close the document.

## NOTE

There is currently an issue with LibreOffice `7.4` that does not allow the `Chart2` class to load.
The `Chart2` has been tested with LibreOffice `7.3`

## Automate

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Chart2/slide_chart/start.py
```

### Windows

```ps
python .\tests\samples\Chart2\slide_chart\start.py
```
