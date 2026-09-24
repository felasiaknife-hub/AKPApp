===================================================================================
VSFlexGrid Pro Build  7.0.0.12		
===================================================================================
==================================================
DOCUMENTATION CHANGES (CUMULATIVE)     13 Dec 1999
==================================================

==================================================
DOCUMENTATION CHANGES (CUMULATIVE)     13 Dec 1999
==================================================
The following items are not reflected in the printed 
documentation. Please see the Help file for details on these 
items.

There is a new GetNode method that returns a VSFlexNode object. 
The VSFlexNode object corresponds to an outline row and has 
methods for adding, removing, locating, moving, sorting, and 
changing outline nodes.

The Cell property has a new flexcpSort setting that allows you 
to sort a range without changing the selection.

The new BookMark(Row) property returns a row's bookmark when 
the grid is bound to an ADO data source.

Translated combos now use strings instead of numbers for keys.

The ComboData property is now a string (used to be a long).

The Appearance property has a new setting: flex3DLight creates 
a thin 3-D border around the control.

CellBorder now supports double lines: use width values >= 100 
to get double lines.

Node.Move flexNMChildOf, nd makes a node a child of another node

MousePointer has a few new cursor options (OLEMove, OLECopy, 
Hand)

Tri-State check boxes (grayed) are now supported (CellChecked 
has new settings)

The flexcpCustomFormat setting of the Cell property has a new 
capability: set it to a string containing a cell reference 
(e.g. "200|5") to copy the given cell's format (everything: 
font, color, alignment, picture, borders, etc).

New ColIndent property allows you to specify an indentation for 
each column (in twips).


================================================== 
CORRECTED PROBLEMS                     13 Dec 1999 
================================================== 
ColFormat now recognizes leading zeros (e.g. "0##" formats "2" 
as "002")

Sort glyphs now change to reflect setting of the Sort property

We fire KeyPressEdit on <Return> and <Esc>

Reset LeftCol/TopRow when setting FrozenCols/FrozenRows

Finish edits before clicks on checkboxes (better event sequence)

Fixed coordinates in OLEDrag events (use client, twips)

No redraw while loading long text files

Improved EditText to return next setting of CheckBoxes

Improved ComboList/ColComboList to display Chr(11) as pipes

Improved response to system color change when control is a 
child of another control

Forced ScrollTips window to top so it's visible over topmost 
windows in NT

Expand rows before deleting them (good when clearing subtotals)

Refresh when setting ColComboLists (to show translated values)

PrintGrid was not working properly with FrozenCols

Now able to dag/drop columns into the first column of the grid

Added different ProgIDs for each flavor of the control: 
VSFlexGrid.VSFlexGridADO, VSFlexGrid.VSFlexGridDAO, 
SFlexGrid.VSFlexGridL, VSFlexGrid.VSFlexGridU. Previously all 
four controls had the ID VSFlexGrid.VSFlexGrid. This caused a 
problem when using the Add method of the Form.Controls 
collection.

==================================================
End of entries for build 7.0.0.12      13 Dec 1999
==================================================

===================================================================================
Initial build 7.0.0	
===================================================================================
Note: The initial release of VSFlexGrid 7.0 is not binarily 
compatible with some previous Beta versions. The GetOutlineNode 
property was deleted and reincarnated as the GetNodeRow method 
for better consistency with the new GetNode method. The data 
type of the ComboData property was also changed.

We refrain from making such changes to a released product, and 
apologize for any inconvenience caused to our testers by doing 
this during the beta cycle.

==================================================
UPGRADE FROM VERSION 6 TO VERSION 7    19 Nov 1999
==================================================
You can upgrade projects that use VSFlexGrid from version 6.0 
to version 7.0 with the Convert utility which is installed in 
C:\Vs\vsFlex7\Convert.  Simply run the utility and point it to 
your project. The utility will make backup copies of your 
project. This utility is provided in source-code form, and 
distributed "as-is." It will also upgrade projects that use 
MSFlexGrid.

==================================================
CORRECTED PROBLEMS                     19 Nov 1999
==================================================
Not applicable for the initial build.

==================================================
End of entries for the initial release 19 Nov 1999
==================================================

