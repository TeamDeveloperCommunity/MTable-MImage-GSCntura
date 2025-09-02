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
* **MTbl**: Main project folder -> mtblXX.dll
* **MTblDoc/HelpNDoc**: HelpNDoc Documentation Project
* **MTblDoc/html**: Documentation in HTML format
* **MTblDoc/chm**: Documentation as CHM file
* **MTblLang**: Language files
* **MImg**: M!Image API header
* **ctd**: Team Developer Interface
* **vis**: Visual Toolchest Interface
* **apihook**: API Hook
* **alphawnd**: Window transparency
* **winversion**: Windows version
* **theme**: Windows Themes
* **fileversion**: File version
* **gradrect**: Gradient rectangle
* **tstring**: STL TCHAR string
* **export**: Excel Export constants

### CxImage
* **CxImageLib**: Main project folder -> cximage.lib, jbig.lib, jpeg.lib png.lib, tiff.lib, zlib.lib

### M!Image
* **MImg**: Main project folder -> mimgXX.dll
* **MImgDoc/HelpNDoc**: HelpNDoc Documentation Project
* **MImgDoc/html**: Documentation in HTML format
* **MImgDoc/chm**: Documentation as CHM file
* **ctd**: Team Developer Interface
* **vis**: Visual Toolchest Interface
* **swcc**: TD Custom Control Interface
* **tstring**: STL TCHAR string

### M!Table Excel Export
* **MTExcExp**: Main project folder -> mtbleeXX.dll
* **MImg**: M!Image API header
* **MTbl**: M!Table API header
* **ctd**: Team Developer Interface
* **winversion**: Windows version
* **export**: Excel Export constants

### GSCntura
* **GSCntura**: Main project folder -> gscntura.dll
* **gsw**: Graphics Server Interface
* **ctd**: Team Developer Interface
* **tstring**: STL TCHAR string

## Dependencies among each other
* **M!Image** depends on CxImageLib
* **M!Table** depends on M!Image 
* **M!Table Excel Export** depends on M!Table and M!Image 