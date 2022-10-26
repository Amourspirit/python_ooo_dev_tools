# Impress Slide to Image

Saves a given page of a slide presentation (e.g. ppt, odp) as an image file (e.g. "gif", "png", "jpeg", "wmf", "bmp", "svg")


## Automate

Args are required to be passed.

**Get Help:**

```sh
python -m start --help
```


### Cross Platform

From current example folder.

```sh
python -m start --file "tests/fixtures/presentation/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Slide_to_Image/start.py --file "tests/fixtures/presentation/algs.ppt" --out_fmt "jpeg" --idx 0
```

### Windows

```ps
python .\tests\samples\Impress\Slide_to_Image\start.py --file "tests/fixtures/presentation/algs.ppt" --out_fmt "jpeg" --idx 0
```
