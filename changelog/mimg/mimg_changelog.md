## V 3.5.0
- General<br>
  Development environment changed from VC 2013 to VC 2022.<br>
  Please note that from this version on the Visual C++ 2015-2022 Redistributable Package is required.

## V 3.4.0
- General<br>
  Ready for TD 7.5

## V 3.3.0
- General<br>
  Ready for TD 7.4

## V 3.2.0
- General<br>
  Ready for TD 7.3

## V 3.1.1
### Fixes
- MImgLoadFromStringBase64 always return 0  (since 3.0.2, worked in earlier versions)

## V 3.1.0
- General<br>
  Ready for TD 7.2

### Fixes
- MImgCreateHICON causes a GDI leak (since 3.0.1, no leak in earlier versions)

## V 3.0.2
### Fixes
- Fixed multithreading issue regarding strings (TD 7.1)

## V 3.0.1
- General<br>
   M!Image now uses the latest version of the CxImage library (7.0.2)

- MImgCreateHBITMAP<br>
   Bitmaps are now created with transparency ( if any )

- MImgCreateHICON<br>
   Icons are now created without using temporary files

### Fixes
- CMimgCtl doesn't show the vertical scrollbar when fit mode MIMG_FIT_WIDTH is set, and/or the horizontal scrollbar when MIMG_FIT_HEIGHT is set

## V 3.0.0
- General<br>
  Ready for TD 7.1<br>
  Ready for 64 Bit (TD 7.0 and TD 7.1)

## V 2.6.0
- General<br>
  Ready for TD 7.0 (Win32)

## V 2.5.5
- General<br>
  Ready for TD 6.3 (Win32)<br>
  Please note that from this version on the Visual C++ 2013 Redistributable Package (x86) is required

### Fixes
- MImgLoadFromHandle<br>
  Loading very large bitmaps may lead to a GPF

M!Image V 2.5.4

## V 2.5.4
- General<br>
  Ready for TD 6.2 (Win32)

- New functions
  - MImgPutTextAligned<br>
   Puts a text aligned into an image

- M!Image Custom Control
  - Added support for scrolling by mousewheel
  - New flag MIMG_FLAG_SET_FOCUS_ON_CLICK<br>
   If set, the control gets the keyboard focus when it's clicked with the left mouse button

### Fixes
- MImgLoadFromScreen<br>
   When capturing a window and Aero is activated, the title bar isn't captured correctly

## V 2.5.3
- General<br>
  Ready for TD 6.1 (Win32)

## V 2.5.2
- New functions
  - MImgGetStringNT<br>
   Gets an image as null terminated string.<br>
   In the Unicode versions of Team Developer, you have to use this function rather than MImgGetString if the string shall be used with SalPicSetString or SalPicSetImage.

### Fixes
- MImgCreateVisPic<br>
   Under certain conditions, a picture created from a 24BPP bitmap has no transparency.

- MImgPutImage<br>
   Under certain conditions, an alpha channel without any transparency is created for the destination image.<br>
   This bug was introduced in V2.5.0.

## V 2.5.1
- General<br>
  Ready for TD 6.0 (Win32)

## V 2.5.0
- Improved Help<br>
  M!Image now ships with a Windows help file ( mimg.chm ) which replaces the old Manual folder.<br>
  Furthermore, a help launcher tool is provided which can be intergrated in the IDE using the Tools interface.

- Icons<br>
  Added support for 256*256 Icons

  New frame constants to load 128*128 Icons:<br>
  MIMG_FRAME_ICON_XXXLARGE<br>
  MIMG_FRAME_ICON_XXXLARGE_4BPP<br>
  MIMG_FRAME_ICON_XXXLARGE_8BPP<br>
  MIMG_FRAME_ICON_XXXLARGE_24BPP

  New frame constants to load 256*256 Icons:<br>
  MIMG_FRAME_ICON_XXXXLARGE<br>
  MIMG_FRAME_ICON_XXXXLARGE_4BPP<br>
  MIMG_FRAME_ICON_XXXXLARGE_8BPP<br>
  MIMG_FRAME_ICON_XXXXLARGE_24BPP

- MImgPutImage<br>
  Modified behaviour when a source pixel is fully opaque or a destination pixel is fully transparent.

- MImgPutText<br>
  New font enhancement flag MIMG_FONT_ENH_CLEARTYPE. If set, the text is drawn using Cleartype (if possible).

### Fixes
- Saving an image as PNG fails when the image has pixel opacity and the color depth is < 24 BPP

## V 2.4.3
### Fixes
- Under certain conditions, a created icon ( by MImgCreateHICON or MImgSave ) looks frayed

## V 2.4.2
- General
  - Ready for TD 5.2

## V 2.4.1
- New functions:
  - MImgHasOpacity<br>
   Determines whether an image has any transparency.

  - MImgLoadFromModuleNamed<br>
   Loads an image from a named resource of a module, e.g. a DLL or an EXE.

## V 2.4.0
- General
  - Ready for TD 5.1

## V 2.3.9
- General
  - Ready for TD 4.2

- MImgLoadFromClipboard
  - Is able to load metafiles now

## V 2.3.8
### Fixes
- MImgLoadFromModule causes a memory leak ( the loaded module isn't freed )

## V 2.3.7
- General
  - From now on there's one DLL per TD version.<br>
   The TD version number is appended to the DLL name, e.g. mimg41.dll for TD 4.1 ( 2005.1 )<br>
   Futhermore, mimg.apl and mimg.app for TD11 have now the correct outline version number ( 4.0.26 )

  - All constants in mimg.apl were moved to the System section

- New functions:
  - MImgDither<br>
   Dithers an image with one of 8 methods

  - MImgGetFrameCount<br>
   Gets an image's total number of frames

  - MImgGetStringBase64<br>
   Gets an image as base64 encoded string

  - MImgLoadFromStringBase64<br>
   Loads an image from a base64 encoded string

- Image Control:
  - New fit modes
    - MIMG_FIT_WIDTH
    - MIMG_FIT_HEIGHT

### Fixes
- Under certain conditions, 32Bit Icons loaded with MImgLoadFromFileInfo, MImgLoadFromHandle or MImgLoadFromVisPic have no transparency

## V 2.3.6
- Ready for TD 4.1 ( TD 2005.1 )

### Fixes
- Under certain conditions and in particular with icons, the following functions may not work properly:<br>
   MImgCreateHICON<br>
   MImgGetString<br>
   MImgSave<br>
   This was introduced in 2.3.5 :o(.

## V 2.3.5
- New functions
  - MImgLoadEmpty<br>
   Loads an empty image

  - MImgPutImage<br>
   Puts an image into another image

- New MIMG_FRAME_ICON_MEDIUM* constants
  - Can be used to load 24*24 icons

### Fixes
- Drawing images in a device context with palettes ( e.g. when drawing on the screen and the color depth is 8 bit ) is very slow.<br>
   The reason was that M!Image created and realized a palette derived from the system palette every time an image was drawn.<br>
   Now M!Image uses the palette which is currently selected into the device context.

## V 2.3.4
- New functions:
  - MImgSetWMFLoadSize / MImgGetWMFLoadSize <br>
   Sets/Gets the fixed size of subsequent loaded WMF images

## V 2.3.3
- Ready for TD 4.0 ( TD 2005 )

- New functions:
  - MImgGetUsedColors<br>
   Gets the colors used in an image

- New draw state MIMG_DST_INVERTED and draw flag MIMG_DRAW_INVERTED

### Fixes
- When an icon - either saved with MImgSave or created with MImgCreateHICON - is drawn with the API function<br>
   DrawIcon and/or DrawIconEx, the transparency is missing.<br>
   This only occurs if the "transparent pixels" aren't black!

- An image with transparent color looses it's transparency after MImgSubstColor is called <br>
   when the color to substitute is MIMG_COLOR_UNDEF or equal to the transparent color

- An image with transparent color looses it's transparency after MImgDisable is called

- An image disabled with MImgDisable becomes type MIMG_TYPE_UNKNOWN

- An image loaded from a HBITMAP with MImgLoadFromHandle is type MIMG_TYPE_UNKNOWN

## V 2.3.2
- New M!Image Custom Control<br>
  Custom Control to easily display images created with M!Image

- New functions:
  - MImgSetLightness<br>
   Increases/decreases an image's lightness and/or contrast

## V 2.3.1
- New functions:
  - MImgInvert<br>
   Inverts an image

  - MImgLoadFromScreen<br>
   Loads an image from the current contents of the screen and/or a window.

  - MImgSetPaletteColor / MImgGetPaletteColor<br>
   Sets/Gets the color of an image's palette entry

### Fixes
- Under certain conditions, Icons loaded with one of the MIMG_FRAME_ICON constants look "damaged"

- An image loaded with MImgLoadFromClipboard is type MIMG_TYPE_UNKNOWN

- Under certain conditions, Icons loaded with MImgLoadFromFileInfo have the wrong size

## V 2.3.0
- Added some documentation in mimg.apl so that the basic function and/or parameter descriptions<br>
  are displayed along with the active coding assistant.<br>
  Thanks to Gerhard Achrainer for his effort!

- New functions:
  - MImgLoadFromVisPic<br>
   Loads an image from a visual toolchest picture

  - MImgPutText<br>
   Puts a text into the image

## V 2.2.0
- New functions:

  - MImgCopyToClipboard<br>
   Copies an image to the clipboard

  - MImgCreateHICON<br>
   Create a HICON from an image

  - MImgIsValidHandle<br>
   Checks whether an image handle is valid

  - MImgLoadFromClipboard<br>
   Loads an image from the clipboard

  - MImgSetSize<br>
   Resizes an image

## V 2.1.0
- New functions:

  - MImgAssign<br>
   Assigns one image to another

  - MImgCreateHBITMAP<br>
   Creates a HBITMAP from an image

  - MImgResample<br>
   Resamples an image to a new size

  - MImgSetJPGQuality / MImgGetJPGQuality<br>
   Let you set/get the JPEG quality
   
  - MImgSetGIFCompr / MImgGetGIFCompr<br>
   Let you set/get the GIF compression

  - MImgSetTIFCompr / MImgGetGIFCompr<br>
   Let you set/get the TIFF compression

  - MImgGetVersion<br>
   Retrieves the M!Image version

- Parameter "nFrame" in MImgLoadFromFile / MIMgLoadFromResource / MImgLoadFromModule<br>
  Now, if you load an icon, you can pass one of the new MIMG_FRAME_ICON constants<br>
  to load an icon with a particular size
