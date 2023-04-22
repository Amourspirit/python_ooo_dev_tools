# Gallery issue

## Description

See [Bug 151932](https://bugs.documentfoundation.org/show_bug.cgi?id=151932) - Gallery issue

See [LibreOffice wiping object properties (Issue with GalleryThemeProvider)](https://ask.libreoffice.org/t/libreoffice-wiping-object-properties-issue-with-gallerythemeprovider/83182) for screenshot and description.

## Steps to reproduce

In the `find_gallery_item()` method, the `item` has a `Drawing` property that is not `None` but the returned `result` has a `Drawing` property that is `None`.

For some reason as soon as the `item` is returned, the `Drawing` property is wiped.

Inside `find_gallery_item()` method reports as expected:

```txt
Gallery item information:
  URL: "private:gallery/svdraw/dd2000"
  Fnm: Unable to compute due to URL conversion error
  Path: Unable to compute due do URL conversion error
  Title: "Note-Gold"
  Type: drawing
```

Outside `find_gallery_item()` method reports Not existing:

```txt
Gallery item information:
  URL: Value is None
  Fnm: Unable to compute due to no URL is None
  Path: Unable to compute due do no URL is None
  Title: "TITLE NOT FOUND"
  Type: empty
```

Run the `start.py` script to reproduce the issue.

### Script output

```txt
Loading Office...

result.Drawing ref count: 0

Gallery item information:
  URL: "private:gallery/svdraw/dd2000"
  Fnm: Unable to compute due to URL conversion error
  Path: Unable to compute due do URL conversion error
  Title: "Note-Gold"
  Type: drawing

item.Drawing type reported on function result <class 'NoneType'>

Gallery item information:
  URL: Value is None
  Fnm: Unable to compute due to no URL is None
  Path: Unable to compute due do no URL is None
  Title: "TITLE NOT FOUND"
  Type: empty
Closing Office
Office terminated
Office bridge has gone!!
```

### Workaround

A workaround has been built into the [Gallery Class](https://github.com/Amourspirit/python_ooo_dev_tools/blob/c6704bb91d903a1a1ea964f4bfc2dcfba8ca5538/ooodev/utils/gallery.py) of `OooDev` to sort of avoid the issue.
