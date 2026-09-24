Database
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------

This project shows how you can use the FlexGrid control to browse
data from a database. The sample uses the NorthWind database that
ships with Microsoft Visual Studio.

The Database project demonstrates:

- How to freeze columns so that selected fields remain visible even
  when the user scrolls the control.
  This is done using the FrozenCols property.    

- How to map database fields to Image Lists. In the sample, this is
  used to display icons for each product category. It is also used
  to display a DB-cursor glyph like the one used in Access tables to
  indicate the current record.

- How to use the new ColIndex property to refer to columns by their
  keys instead of using absolute indices. The column key comes from
  the database (it is set to the field name by default), but you can
  set it in code using the ColKey property.
  This is nice because you can change the order of the fields ot add
  new fields and not worry about adjusting the column indices.

- How to build translated fields using the new BuildComboList method.
  Translated fields contain codes that get mapped into strings. The
  mapping is done via the ColComboList property.
  In this case, the grid will store only the code (e.g. ProductID), 
  but the user will see the product name instead.

- How to display a secondary recordset for a given field without
  using hierarchical recordsets. The secondary recordset is created
  on the fly with an SQL statement, and a secondary form is used to
  show it.

- How to provide "ScrollTips", a type of tooltip that appears while 
  the user drags the vertical scrollthumb and shows which record 
  would become visible at the top of the grid if the user released 
  the thumb. This helps users navigate grids with many rows.

- How to build ADO recordsets from SQL statements generated in Access.
  This is not specific to the FlexGrid control, but is a very common
  and useful technique.
