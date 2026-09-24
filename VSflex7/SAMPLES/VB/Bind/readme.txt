Bind
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can bind the FlexGrid control to Variant 
Arrays or to other FlexGrid controls.


The Bind project demonstrates:

- How to bind the FlexGrid to a Variant Array using the BindToArray 
  method. In this case, the grid and the array become connected. Changes
  made to the grid contents are saved back in the array, and changes
  made to the array are displayed on the grid.
  Note that after the array contents are changed, you need to call the
  grid's Refresh method before the changes become visible.

- How to load the FlexGrid with data from a Variant Array using the 
  LoadArray method. In this case, the array data is copied into the grid.
  After this, the grid and the array are not connected.

- How to bind the FlexGrid to another FlexGrid using the BindToArray 
  method. In this case, the grids become connected. Changes made to either
  grid are reflected on the other.
  Note that after the source grid contents are changed, you need to call the
  bound grid's Refresh method before the changes become visible.

- How to load the FlexGrid with data from another FlexGrid using the 
  LoadArray method. In this case, the sourc grid data is copied into the 
  bound grid. After this, the grids are not connected.
