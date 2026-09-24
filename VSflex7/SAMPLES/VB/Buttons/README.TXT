Buttons
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


These projects show how you can implement custom cell buttons that
remain visible even when a cell does not have the focus. There are two
variations:


1) NOAPI
This project does not use any API calls. It uses the Cell(flexcpPicture)
property to display the buttons and the BeforeMouseDown event to handle 
mouse clicks.

This technique is efficient and easy to implement. The only minor 
drawbacks are the use of fixed colors for the buttons (instead of
system colors), the fact that clicks are triggered when the mouse 
button is pressed, rather than the usual (press, move, release, click)
sequence, and the fixed size of the buttons (if you resize a row, the
button aligns but does not stretch).


2) WITHAPI
This project uses the OwnerDraw property and several APIs to draw 
buttons that look exactly like the built-in ones.

This technique is a little less efficient (because the OwnerDraw property
slows down the repainting) and a requires a bit more code, but it 
overcomes the drawbacks listed earlier. It uses system colors to draw
the button faces, the mouse handling is standard, and the buttons
stretch to fit rows of different heights. 

