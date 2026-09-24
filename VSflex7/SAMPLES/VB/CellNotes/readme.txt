CellNotes
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can implement Excel-style cell notes using
the FlexGrid control. This is done using the FlexGrid's ability to
attach data to each grid cell.

The CellNotes project demonstrates:

- How to use the Cell property with the flexcpData setting to attach
  arbitrary data to individual grid cells.
  This sample attaches strings containing notes about each cell, but 
  you could attach any Variant, even references to other objects.

- How to use the Cell property with the flexcpPicture and 
  flexcpPictureAlignment settings to attach pictures to individual cells.
  This sample uses small icons, which are transparent and thus allow
  the user to see the cell contents.

- How to use the MouseMove event to provide sophisticated user-interface 
  elements.

- How to build reusable forms than can be added to other projects that
  contain VSFlexGrid controls.