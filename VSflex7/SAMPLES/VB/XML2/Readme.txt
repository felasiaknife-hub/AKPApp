XML2
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the VSFlexGrid control to display
XML data in an outline format. The project uses the Microsoft XML
parser (MSXML.DLL) to parse an XML document and populate the grid.

The XML2 project uses MSXML version 2.0, which includes a richer
object model than version 1.0. This sample is based on the
MSXML.DOMDocument and MSXML.IXMLDOMNode objects.

The XML project uses MSXML version 1.0, which includes a simpler
object model, based on the MSXML.XMLDocument and MSXML.IXMLElement2
objects.


The XML2 project demonstrates:

- How to use the IsSubtotal and RowOutlineLevel properties to build
  a custom outline.

- How to use the OutlineBar and OutlineCol properties to display
  a "tree-style" outline bar.

- How to use the MergeCells property and the new flexMergeOutline 
  setting to make subtotal (node) rows automatically spill across 
  into empty cells so their text never gets truncated even when
  the outline column is narrow.

- How to trap arrow keys and collapse or expand the tree when the
  user hits the left and right arrow keys.

