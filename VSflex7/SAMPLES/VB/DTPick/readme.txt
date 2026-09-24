DTPick
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use a DateTimePicker control to edit
date entries on a VSFlexGrid. You may apply the same techniques to
other controls.

The project uses the VSFlexGrid control's CellLeft, CellTop, CellWidth, 
and CellHeight properties to position the DateTimePicker control over
the cell being edited, then monitors the control to retrieve the
user's selection and to finish the editing process.

The DTPick project demonstrates:

- How to display a drop-down arrow on cells to indicate they can be
  edited with a custom control.

- How to use the StartEdit event to suppress the default editing 
  process and display your own custom control.

- How to prevent the user from resizing or scrolling the control
  while the custom editor is active.
