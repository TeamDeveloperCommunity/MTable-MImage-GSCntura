# Extensions

## M!Table
With M!Table your tables finally get the features you always wanted, e.g. Excel export, printing, tree structure, variable row heights, buttons, images, tooltips, hyperlinks, customizable fonts and colors, customizable lines, customizable appearance of selections,
merging of cells, grouping of column headers, sorting by any number of columns, automatic adjustment of column widths and row heights, and much more...

## M!Image
M!Image is a powerful and easy to use library for loading, drawing, saving and manipulating images incl. a custom control to easily display them.

## GSCntura
GSCntura is an interface to Graphics Server 5.x which contains all functions with the same export ordinals as the original gscntura.dll for CTD 1.x, so you can use the original APLs from Bits Per Second. Simply replace the old dll with the new one.

# Information regarding development

## Development environment
Microsoft Visual Studio Community (C++)<br>
Download: https://visualstudio.microsoft.com/de/vs/community/

## Documentation tool
HelpNDoc Personal Edition<br>
Download: https://www.helpndoc.com/de/download/

## Relevant folders

### M!Table
* **vc/mtbl**: Main project folder -> mtblXX.dll
* **doc/mtbl/helpndoc**: HelpNDoc Documentation Project
* **doc/mtbl/html**: Documentation in HTML format
* **doc/mtbl/chm**: Documentation as CHM file
* **lang/mtbl**: Language files
* **vc/mimg**: M!Image API header
* **vc/ctd**: Team Developer Interface
* **vc/vis**: Visual Toolchest Interface
* **vc/apihook**: API Hook
* **vc/alphawnd**: Window transparency
* **vc/winversion**: Windows version
* **vc/theme**: Windows Themes
* **vc/fileversion**: File version
* **vc/gradrect**: Gradient rectangle
* **vc/tstring**: STL TCHAR string
* **vc/export**: Excel Export constants

### CxImage
* **vc/cximglib**: Main project folder -> cximage.lib, jbig.lib, jpeg.lib png.lib, tiff.lib, zlib.lib

### M!Image
* **vc/mimg**: Main project folder -> mimgXX.dll
* **doc/mimg/helpndoc**: HelpNDoc Documentation Project
* **doc/mimg/html**: Documentation in HTML format
* **doc/mimg/chm**: Documentation as CHM file
* **vc/ctd**: Team Developer Interface
* **vc/vis**: Visual Toolchest Interface
* **vc/swcc**: TD Custom Control Interface
* **vc/tstring**: STL TCHAR string

### M!Table Excel Export
* **vc/mtblee**: Main project folder -> mtbleeXX.dll
* **vc/mimg**: M!Image API header
* **vc/mtbl**: M!Table API header
* **vc/ctd**: Team Developer Interface
* **vc/winversion**: Windows version
* **vc/export**: Excel Export constants

### GSCntura
* **vc/gscntura**: Main project folder -> gscntura.dll
* **vc/gsw**: Graphics Server Interface
* **vc/ctd**: Team Developer Interface
* **vc/tstring**: STL TCHAR string

## Dependencies among each other
* **M!Image** depends on CxImageLib
* **M!Table** depends on M!Image 
* **M!Table Excel Export** depends on M!Table and M!Image 