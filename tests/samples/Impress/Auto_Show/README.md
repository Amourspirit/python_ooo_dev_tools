# Impress Auto Slide Show

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198406431-6b28b28b-4949-4a41-bf67-ff485ab964a2.png" width="659" height="448">
</p>

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
