CustDataADO
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This sample consists of two projects:


1) Simple
Shows how you can use the FlexDataSource property to bind the VSFlexGrid
control to a custom data source that is based on an ADO recordset.

This is similar to assigning the ADO recordset directly to the 
VSFlexGrid's DataSource property, with two differences:
- You can use this technique to bind an ADO recordset to FlexGrid light,
  which has no built-in ADO support.
- You have more flexibility over the binding process, since all data
  requests go through the custom FlexDataSource class.


2) Filtered
Builds on the previous example and takes advantage of the custom 
FlexDataSource class to implement a "filter line".

The top row on the grid is frozen, and can be used to enter filter 
criteria for each field. Whenever the user changes the filter, the
underlying recordset is requeries and the results are displayed on the
grid.

For example, if the user types "> 3" on the first grid row of the
EmployeedID column, the grid will display all entries where the
EmployeedID has values 4, 5, 6, etc.

