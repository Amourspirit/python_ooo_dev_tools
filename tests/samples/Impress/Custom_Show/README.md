# Impress Custom Slide Show

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198407936-7865b1c2-75b7-4530-8598-a1ce52821752.png" width="448" height="448">
</p>


Demonstrates opening a presentation file in Impress starting a slide show with only the slide indexes passes in.

A message box is display once slide show has completed asking if you want to close document.

## Automate

Slide index numbers can be passed in.

**Example:**

```sh
python tests/samples/Impress/Custom_Show/start.py 5 6 7 8
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Custom_Show/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Custom_Show\start.py
```
