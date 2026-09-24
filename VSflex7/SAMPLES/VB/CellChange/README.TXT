CellChange
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the CellChange event to provide
conditional formatting and dynamic data summaries (that get updated
automatically whenever the data changes).

The CellChange project demonstrates:

- How to use the FrozenRows property to provide a static (selectable
  but not scrollable) area on the grid where the dynamic data summary
  can be displayed.

- How to use the CellChanged event to format individual cells. The demo
  shows negative numbers in a red bold face, large numers in bold green,
  and removes the custom formatting for non-negative numbers.

- How to use the CellChanged event to update the dynamic data summary
  displayed across the top of the grid. The demo uses the new Aggregate
  property to get each column's statistics (count, average, min, max, 
  standard deviation, and variance).

The CellChanged event is useful when you want some processing done
whenever the contents of a cell have changed, regardless of how they 
were changed (e.g. user typed data into the cell, data got loaded from
a database, or data was assigned to the grid through code). To 
illustrate this, the demo includes a button that generates random data,
which is processed in the same way as the data you type in.

The Aggregate property calculates aggregate values for a range. It is
very fast.
