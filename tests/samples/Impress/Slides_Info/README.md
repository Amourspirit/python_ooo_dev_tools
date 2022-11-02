# Impress Get Info

<p align="center">
    <img src="https://user-images.githubusercontent.com/4193389/198415603-a7ea1593-06a7-482f-b245-0933d0f5950d.png" width="396" height="314">
</p>


Demonstrates get info using default `algs.odp` file.

- open a draw/slides document
- report on layers
- report style information: families, styles, properties
- show no. of slides and slide size

## Automate

A message box is display once the document has been created asking if you want to close the document.

A single parameters can be passed in which is the slide show document to modify:

**Example:**

```sh
python ./tests/samples/Impress/Slides_Info/start.py "tests/fixtures/presentation/algs.odp"
```

If no parameters are passed then the script is run with the above parameters.

### Cross Platform

From current example folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./tests/samples/Impress/Slides_Info/start.py
```

### Windows

```ps
python .\tests\samples\Impress\Slides_Info\start.py
```

## Output

```txt
Loading Office...
Opening /home/user/LibreOffice/ooouno-dev-tools/tests/fixtures/presentation/algs.odp

No. of slides: 12

Size of slide (mm)(254, 190)

 Layer 0 Properties
  Description: 
  IsLocked: False
  IsPrintable: True
  IsVisible: True
  Name: layout
  Title: 

 Layer 1 Properties
  Description: 
  IsLocked: False
  IsPrintable: True
  IsVisible: True
  Name: background
  Title: 

 Layer 2 Properties
  Description: 
  IsLocked: False
  IsPrintable: True
  IsVisible: True
  Name: backgroundobjects
  Title: 

 Layer 3 Properties
  Description: 
  IsLocked: False
  IsPrintable: True
  IsVisible: True
  Name: controls
  Title: 

 Layer 4 Properties
  Description: 
  IsLocked: False
  IsPrintable: True
  IsVisible: True
  Name: measurelines
  Title: 

Background Object Props Properties
  Description: 
  IsLocked: False
  IsPrintable: True
  IsVisible: True
  Name: backgroundobjects
  Title: 

Style Families in this document:
No. of names: 4
  'cell'  'Default'  'graphics'  'table'



Styles in the "Default" style family:
No. of names: 14
  'background'  'backgroundobjects'  'notes'  'outline1'
  'outline2'  'outline3'  'outline4'  'outline5'
  'outline6'  'outline7'  'outline8'  'outline9'
  'subtitle'  'title'


Styles in the "cell" style family:
No. of names: 34
  'blue1'  'blue2'  'blue3'  'bw1'
  'bw2'  'bw3'  'default'  'earth1'
  'earth2'  'earth3'  'gray1'  'gray2'
  'gray3'  'green1'  'green2'  'green3'
  'lightblue1'  'lightblue2'  'lightblue3'  'orange1'
  'orange2'  'orange3'  'seetang1'  'seetang2'
  'seetang3'  'sun1'  'sun2'  'sun3'
  'turquoise1'  'turquoise2'  'turquoise3'  'yellow1'
  'yellow2'  'yellow3'


Styles in the "graphics" style family:
No. of names: 40
  'A4'  'A4'  'Arrow Dashed'  'Arrow Line'
  'Filled'  'Filled Blue'  'Filled Green'  'Filled Red'
  'Filled Yellow'  'Graphic'  'Heading A0'  'Heading A4'
  'headline'  'headline1'  'headline2'  'Lines'
  'measure'  'Object with no fill and no line'  'objectwitharrow'  'objectwithoutfill'
  'objectwithshadow'  'Outlined'  'Outlined Blue'  'Outlined Green'
  'Outlined Red'  'Outlined Yellow'  'Shapes'  'standard'
  'Text'  'text'  'Text A0'  'Text A4'
  'textbody'  'textbodyindent'  'textbodyjustfied'  'title'
  'Title A0'  'Title A4'  'title1'  'title2'


Styles in the "table" style family:
No. of names: 11
  'blue'  'bw'  'default'  'earth'
  'green'  'lightblue'  'orange'  'seetang'
  'sun'  'turquoise'  'yellow'


Closing the document
Closing Office
Office terminated
Office bridge has gone!!
```