# Impress Modify Pages

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198418648-34ab1937-9d3d-4a10-bf38-54f4abc39775.png">
</p>

Add two new slides to the input document

- add a title-only slide with a graphic at the end
- add a title/subtitle slide at the start

A message box is display once the slide show has ended asking if you want to close the document.

## Automate

A single parameters can be passed in which is the slide show document to modify:

**Example:**

```sh
python ./tests/samples/Impress/Modify_Slides/start.py "tests/fixtures/presentation/algsSmall.ppt"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Modify_Slides/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Modify_Slides\start.py
```
