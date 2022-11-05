# Show Sheet

Example of opening a spreadsheet and inputting a password (foobar) to unlock sheet.

## Automate

A message box is display once the document has been processed asking if you want to close the document.

The following command will run automation that opens Calc document and ask for password.

### Cross Platform

From this folder.

```sh
python -m start --show --file "../../../../resources/ods/totals.ods"
```

### Linux/Mac

```sh
python ./tests/samples/Calc.Show_Sheet/start.py --show --file "tests/fixtures/calc/totals.ods"
```

### Windows

```ps
python .\tests\samples\Calc\Show_Sheet\start.py --show --file "tests/fixtures/calc/totals.ods"
```

![business-spreadsheet](https://user-images.githubusercontent.com/4193389/194169727-f5a61ab2-e336-42c3-8ef1-31299b81100d.jpg)
