WallPaper
VSFlexGrid 7 VB Demo Project
------------------------------------------------------------------------


This project shows how you can use the new WallPaper and 
WallPaperAlignment properties to provide the FlexGrid control with a 
graphical background.

The project loads a NorthWind table to fill the grid with data, 
assuming there is an ODBC data source defined on your system and 
linked to the NorthWind database that ships with Microsoft Visual 
Studio.

Here's a few notes and tips on using the WallPaper property:

- How the WallPaper is displayed
  The WallPaper is a static backdrop. It does not scroll along with the
  grid contents. Thus, you cannot use it to add graphical elements that
  are related to grid contents such as boxes around specific cells or 
  range highlights.
  The wallpaper is applied only to the scrollable areas of the grid.
  Fixed cells are not affected by the WallPaper.
 
- Use "Faded" Pictures
  The WallPaper picture is displayed "behind" the grid contents. If you
  use a picture that has strong colors, the grid contents may be obscured.
  Note how the pictures used in the sample are "faded". This effect was
  created using filters provided by the Microsoft Image Composer
  application. If you want to use dark pictures, set the grid's ForeColor 
  to a light value so the grid contents will be clearly visible.
      
- BackColor property
  When using WallPaper, remember to set the BackColor property to a value
  that matches the predominant color on the WallPaper picture. The grid's
  BackColor is used to paint the small areas beyond the grid, next to the 
  fixed cells. It is also used to paint the background of edit controls
  while the grid is in edit mode. Finally, the transparent background is 
  used if you assign a transparent picture to the WallPaper property
  (such as an icon, metafile, or gif image).

- WallPaperAlignment
  The WallPaperAlignment property can be used to determine whether the 
  WallPaper picture should be stretchd or tiled to fit the grid's 
  scrollable area, or whether it should be simply aligned horizontally 
  and vertically.
  By default, the image is stretched. This is the most flexible option, 
  but is also the slowest. Stretching complex pictures is a time-
  consuming operation. You can notice the difference in the sample 
  application by setting the alignment to flexPicAlignStretch or
  flexPicAlignLeftTop, then scrolling the grid.
  
- Transparent pictures
  You can use transparent pictures such as icons, metafiles, or gifs as
  WallPaper. The transparent portions will be painted using the grid's
  BackColor property.
  However, rendering transparent pictures (especially gifs) is generally 
  much slower than rendering solid pictures.
  We recommend using solid, compressd JPG pictures as WallPaper. This
  yields good results in terms of rendering speed and disk space.
  