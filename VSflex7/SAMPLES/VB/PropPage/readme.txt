PropPage
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the FlexGrid control to build a
Property browser interface similar to the one in VB, alphabetic and
categorized views, custom editors, and owner-draw painting.

The PropPage project demonstrates:

- How to group items based on categories using the Subtotal method.
  This involves sorting the grid and using the Subtotal method to 
  clear and to insert subtotals.

- How to control which cells can be edited. This is done by trapping
  the BeforeEdit event and setting the Cancel parameter to True to
  prevent editing cells that hold fixed values (such as property names
  and categories).

- How to provide custom mouse and keyboard handling to facilitate editing.
  This is done by trapping the MouseUp and KeyDown events and changing the
  selection to the editable column.

- How to provide custom navigation. Typing CTRL plus a character takes the
  cursor to the next proeprty whose name starts with the character. This
  is similar to the VB-behavior, and is also done by trapping the KeyDown
  event.

- How to provide extended combo-list searching capabilities. The FontName
  properties have long combo lists that automatically search as the user
  types. This is the default behavior for combo boxes in VSFlexGrid7, but
  it can be disabled using the new ComboSearch property.

- How to provide automatic editing on double-clicks. This is done by 
  setting the Editable prperty to the new flexEDKbdMouse setting. This
  allows the user to start editing a cell by double-clicking on it, 
  without any code.

- How to provide custom editors, with buttons on the cells that bring up
  dialogs to edit the cell. This is used to provide color and font editors.
  This is done by trapping the BeforeEdit event and setting the ComboList
  property to an ellipsis ("..."), then trapping the CellButtonClick event
  to display the dialog.
  The sample also uses the new CellButtonPicture property to display a 
  custom picture in the cell button (instead of the default ellipsis).

- How to position and display a custom editor built from a VB Form. The
  sample shows a color picker built with two grids.

- How to use the OwnerDraw property to add custom elements to a cell's 
  appearance. The sample uses this technique to display the selected
  color in color properties, much like the VB properties window does.
  This requires setting the OwnerDraw property to one of the available
  settings, then trapping the DrawCell event to implement the custom 
  drawing.

- How to use the Cell property to provide custom cell formatting. The
  sample uses this technique to display the selected font in the font 
  property cells.

