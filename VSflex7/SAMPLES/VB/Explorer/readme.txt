Explorer
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the FlexGrid control to build an
Explorer-like interface, with a tree on the left showing the directories
and a list on the right showing files.


The Explorer project demonstrates:

- How to build an outline asynchronously using the VSFlexGrid control.
  Items are added only when the branch is expanded. This saves a lot
  of time, because we don't have to load the entire directory structure
  for every drive on the computer at startup.
  This is done using the new BeforeCollapse event.

- How to customize an outline to provide mouse and keyboard control.
  Double-clicking on outline nodes collapses or expands them. The left
  and right arrow keys can also be used to collapse/expand nodes.
  The outline also supports autosearching as the user types a directory
  name.
  The new NodeOpenPicture/NodeClosedPicture properties are used to customize
  the appearance of collapsed and expanded nodes.

- How to use the ExplorerBar property to provide headers that can be used
  to size and sort columns. The ExplorerBar property is set to the new 
  flexExSortShow setting to provide visual indicators of the last sorted
  column and sort direction (little arrows like those in the Explorer).

- Buffered painting. The Redraw property is set to the new flexRDBuffered
  setting to provide the outline with buffered painting. This makes 
  redrawing extremely smooth.

- Manual OLE drag and drop. Select one or more files by clicking or
  control-clicking them, then shift-drag them into NotePad or Word. This
  requires adding a list of selected files to the VSDataObject before
  the dragging actually starts.

- How the Ellipsis property can be used so the grid indicates long
  entries that won't fit in the space available by appending an ellipsis
  symbol (...) to the entry.

- How to use the Aggregate function to calculate aggregates over a range.
  (In this case, the total size of the selected files.)
  The Aggregate function allows you to specify a range explicitly or to 
  use the selected rows.

- How to use the new WallPaper property to improve the look of a grid
  by giving it a graphical background.

- A simple way to implement a splitter bar using just a label control and
  Drag/Drop code. This does not involve the FlexGrid, but still makes a
  pretty useful demo.

