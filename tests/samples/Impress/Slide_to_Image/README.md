# Impress Slide to Image

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198423388-f8845bec-781a-42ef-b8cf-20bb13b9cb43.png">
</p>

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
