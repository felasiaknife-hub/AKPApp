FilterDB
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can filter data as it is retrieved or saved
to a bound recordset using the FlexGrid control. This is done with the
FilterData event.

The FilterData event is fired after each value is read from a bound
recordset and before each value is written back to the recordset. This
event is mostly useful when the data is stored in the database using 
a compressed format that is not ideal for displaying and editing.

For example, many databases store dates as a string of digits (usually 6 
or 8), without any separators. You could use the FilterDB event to
insert the separators at the proper places before displaying the data,
and to remove them again before writing them back into the database.

This is exactly what the FilterDB demo does.

Here is a description of the parameters used with the FilterData event:

fgADO_FilterData(ByVal Row As Long, ByVal Col As Long, Value As String, _
		         ByVal SavingToDB As Boolean, WantThisCol As Boolean)

Row, Col
	Address of the cell whose value is about to be read from or written 
	to the bound recordset.

Value
	Value just read from or about to be written to the bound recordset.

SavingToDB
	If True, the value was read from the grid and is about to be written
	to the bound recordset. If False, the value was read from the bound 
	recordset and is about to be assigned to the grid.

WantThisCol
	This value is set to False by default. If you set it to True, the 
	control will keep firing the FilterData event for this column. If
	you don't set it to True, the event will no longer be fired for 
	this column until the grid is bound to a new recordset or the 
	current recordset is refreshed.

	The WantThisCol parameter is important because the FilterData event
	is relatively slow. The WantThisCol parameter allows the 
	application to establish which columns need filtering, thus improving
	performance.
