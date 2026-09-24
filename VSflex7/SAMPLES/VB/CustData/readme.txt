CustData
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the FlexGrid control's FlexDataSource
property to implement grid-based views of custom data structures without 
any conventional database support (no DAO, OLEDB, ADO, ODBC, etc).

This technique is useful if you want to display or edit large amounts
of data that is stored in custom data structures. The advantages of 
this technique are:

1) Speed:
   You don't need to copy all the data into the grids. The grid is used
   only to display the information, and does not store any data. The
   sample displays two grids with several thousand rows, which would 
   take a few seconds to load using the traditional AddItem or 
   TextMatrix methods.

2) Space:
   The data is not copied into the grid. It only exists in your custom
   structures.

3) Flexibility:
   You may display the data in any format and layout. You may validate
   the data as the user edits it, and even provide calculated fields
   that are obtained from combinations of other fields.

4) Simplicity:
   The FlexDatasource interface is extremely simple. It has only two 
   methods (GetData and SetData), so you don't need to implement 
   complicated methods to emulate recordset behavior.


The CustData project demonstrates:

- How to implement a VB class module that implements the IVSFlexDataSource interface.
  This requires a few easy steps:
  1) Insert a new call module into your project,
  2) Add the declaration "Implements VSFlex7LCtl.IVSFlexDataSource" to the module,
  3) Select the "IVSFlexDataSource" object in the code window (left drop-down),
  4) Select and implement the following functions:
     GetFieldCount returns the number of fields (columns) that are available
     GetFieldName returns the name of each field (columns)
     GetRecordCount returns the number of records (rows) that are available
     GetData returns the data for a given field/record (column/row)
     SetData validates and sets new data for a given field/record (column/row)

- How to implement GetData so it provides "real data", calculated fields, and labels
  for each grid cell.

- How to implement SetData so it provides validation on each data element.

- How to connect the custom data provider classes to the VSFlexGrid control. This is
  done by assigning an instance of the class to the FlexDataSource property.

- How to prevent editing of calculated fields, and how to invalidate controls when
  the user changes items that affect calculated fields.

- Other useful techniques that are not specific to custom data-binding, such as
  freezing columns (with the FrozenCols and AllowUserFreezing properties),
  formatting data (with the ColFormat property), editing with double-clicks
  (setting the Editable property to flexEDKbdMouse), and buffered repainting
  (setting the Redraw property to flexRDBuffered).

