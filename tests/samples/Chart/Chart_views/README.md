# Demonstrates creating charts

![chart_web_filled](https://user-images.githubusercontent.com/4193389/198901667-d7d1da10-4436-4cfc-adce-2e82d1f6692b.png)

Demonstrates loading a spread sheet into Calc and dynamically inserting charts using the `Chart` class.
There are a total of 10 different charts that can be dynamically created by this demo.

A message box is display once the document has been created asking if you want to close the document.

## Options

The type of chart created is determined by the `-k` option.

Possible `-k` options are:

- area
- bar
- bubble
- donut
- net
- net_filled
- line
- pie
- stock
- xy

## Automate

### Cross Platform

From current example folder.

```sh
python -m start -k bar
```

### Linux/Mac

```sh
python ./tests/samples/Chart/Chart_views/start.py -k bar
```

### Windows

```ps
python .\tests\samples\Chart\Chart_views\start.py -k bar
```
