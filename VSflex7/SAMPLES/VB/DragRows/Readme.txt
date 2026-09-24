DragRows
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the ExplorerBar property and the 
DragRow method to allow users to drag rows to new positions on the
grid.

The DragRows project demonstrates:

- How to set the ExplorerBar property so the grid will allow the end 
  user to sort columns by clicking on the column header cell, as well
  as drag rows and columns by dragging their header cells to a new
  position.

- How to use the RowDrag method to start a row dragging process when
  the user right-clicks on any cell. This technique is useful when
  grids have no fixed columns, in which case the rows have no "header"
  cells on the left and thus the user can't drag them in the standard
  maner.

- How to use the DragRow method to provide custom highlighting for the
  row being dragged. The trick here is to use DragRow's return value
  to update the row position so you can remove the custom highlighting
  when the dragging operation is finished. 

- How to use the AfterMoveRow event to provide custom behavior. The
  sample tests whether the SelectionMode property is set to 
  flexSelectionListBox. If it is, then some VB code is used to move 
  the selected rows to the new position (the default behavior is to
  move only the row that was being clicked).
