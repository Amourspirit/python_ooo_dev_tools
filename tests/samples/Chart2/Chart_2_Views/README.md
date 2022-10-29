# Demonstrates creating charts.

![charts_3d](https://user-images.githubusercontent.com/4193389/198852486-43ffdca8-9b26-4001-a734-1031cd4e42dd.png)


Demonstrates loading a spread sheet into Calc and dynamically inserting charts.
There are a total of 17 different charts that can be dynamically created by this demo.

A message box is display once the document has been created asking if you want to close the document.

## NOTE

There is currently an issue with LibreOffice `7.4` that does not allow the `Chart2` class to load.
The `Chart2` has been tested with LibreOffice `7.3`

## Options

The type of chart created is determined by the `-k` option.

Possible `-k` options are:

- area
- bar
- bubble_labeled
- col
- col_line
- col_multi
- donut
- happy_stock
- line
- lines
- net
- pie
- pie_3d
- scatter
- scatter_line_error
- scatter_line_log
- stock_prices

## Automate

### Cross Platform

From current example folder.

```sh
python -m start -k happy_stock
```

### Linux/Mac

```sh
python ./tests/samples/Chart2/Chart_2_Views/start.py -k happy_stock
```

### Windows

```ps
python .\tests\samples\Chart2\Chart_2_Views\start.py -k happy_stock
```
