// MImgApi.h

#ifndef _INC_MIMGAPI
#define _INC_MIMGAPI

// Includes
#include "ctd.h"
#include "MImgConst.h"

// Exportierte Funktionen
BOOL WINAPI MImgAddShadow( HANDLE hImg, int iDispX, int iDispY, BYTE bySmoothRad, DWORD dwClr, DWORD dwClrBkgnd, BYTE byIntens );
BOOL WINAPI MImgAssign( HANDLE hImgSrc, HANDLE hImgDest );
BOOL WINAPI MImgCopyToClipboard( HANDLE hImg );
HANDLE WINAPI MImgCreateCopy( HANDLE hImg );
HANDLE WINAPI MImgCreateCopyEx( HANDLE hImg, LPRECT pRect, DWORD dwType, DWORD dwFlags );
HBITMAP WINAPI MImgCreateHBITMAP( HANDLE hImg, HDC hDC );
HICON WINAPI MImgCreateHICON( HANDLE hImg );
HANDLE WINAPI MImgCreateVisPic( HANDLE hImg, DWORD dwType );
BOOL WINAPI MImgCtlGetBackBrush( HWND hWndImg, LPINT piStyle, LPCOLORREF pclr );
BOOL WINAPI MImgCtlGetContentsExtent( HWND hWndImg, LPSIZE ps );
int WINAPI MImgCtlGetFitMode( HWND hWndImg );
HANDLE WINAPI MImgCtlGetImage( HWND hWndImg );
//BOOL WINAPI MImgCtlGetImagePixFromPoint( HWND hWndImg, POINT pt, LPPOINT pPt );
BOOL WINAPI MImgCtlGetImagePixFromPoint(HWND hWndImg, LONG lX, LONG lY, LPPOINT pPt);
BOOL WINAPI MImgCtlGetImagePos( HWND hWndImg, LPPOINT ppt );
double WINAPI MImgCtlGetScale( HWND hWndImg, int iWhat );
BOOL WINAPI MImgCtlGetTextPos( HWND hWndImg, LPPOINT ppt );
BOOL WINAPI MImgCtlQueryAlignFlags( HWND hWndImg, DWORD dwFlags );
BOOL WINAPI MImgCtlQueryFlags( HWND hWndImg, DWORD dwFlags );
BOOL WINAPI MImgCtlRefresh( HWND hWndImg );
BOOL WINAPI MImgCtlSetAlignFlags( HWND hWndImg, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MImgCtlSetBackBrush( HWND hWndImg, int iStyle, COLORREF clr );
BOOL WINAPI MImgCtlSetFitMode( HWND hWndImg, int iFitMode );
BOOL WINAPI MImgCtlSetFlags( HWND hWndImg, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MImgCtlSetImage( HWND hWndImg, HANDLE hImg, BOOL bOwn );
//BOOL WINAPI MImgCtlSetImagePos( HWND hWndImg, POINT pt );
BOOL WINAPI MImgCtlSetImagePos(HWND hWndImg, LONG lX, LONG lY);
BOOL WINAPI MImgCtlSetScale( HWND hWndImg, int iWhat, double dScale );
//BOOL WINAPI MImgCtlSetTextPos( HWND hWndImg, POINT pt );
BOOL WINAPI MImgCtlSetTextPos(HWND hWndImg, LONG lX, LONG lY);
BOOL WINAPI MImgCtlSizeToContents( HWND hWndImg );
BOOL WINAPI MImgDelete( HANDLE hImg );
BOOL WINAPI MImgDisable( HANDLE hImg );
BOOL WINAPI MImgDither( HANDLE hImg, long lMethod );
BOOL WINAPI MImgDraw( HANDLE hImg, HDC hDC, LPCRECT pRect, DWORD dwFlags = 0 );
BOOL WINAPI MImgDrawAligned( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid = -1, int iHt = -1, DWORD dwAlignFlags = 0, DWORD dwFlags = 0 );
BOOL WINAPI MImgDrawPattern( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid = -1, int iHt = -1, DWORD dwFlags = 0 );
int WINAPI MImgEnum( HARRAY ha );
WORD WINAPI MImgGetBpp( HANDLE hImg );
HANDLE WINAPI MImgGetDIBHandle( HANDLE hImg );
BOOL WINAPI MImgGetDIBString( HANDLE hImg, LPHSTRING phs );
BOOL WINAPI MImgGetDrawClipRect( HANDLE hImg, LPRECT pRect );
DWORD WINAPI MImgGetDrawState( HANDLE hImg );
int WINAPI MImgGetFrameCount( HANDLE hImg );
int WINAPI MImgGetGIFCompr( HANDLE hImg );
BYTE WINAPI MImgGetJPGQuality( HANDLE hImg );
BYTE WINAPI MImgGetMaxOpacity( HANDLE hImg );
COLORREF WINAPI MImgGetPaletteColor( HANDLE hImg, BYTE byIdx );
COLORREF WINAPI MImgGetPixelColor( HANDLE hImg, int iX, int iY );
BYTE WINAPI MImgGetPixelOpacity( HANDLE hImg, int iX, int iY );
BOOL WINAPI MImgGetSize( HANDLE hImg, LPSIZE pSize );
BOOL WINAPI MImgGetString( HANDLE hImg, LPHSTRING phs, DWORD dwType );
BOOL WINAPI MImgGetStringBase64( HANDLE hImg, LPHSTRING phs, DWORD dwType );
int WINAPI MImgGetTIFCompr( HANDLE hImg );
DWORD WINAPI MImgGetTranspColor( HANDLE hImg );
int WINAPI MImgGetType( HANDLE hImg );
int WINAPI MImgGetUsedColors( HANDLE hImg, HARRAY haColors );
HSTRING WINAPI MImgGetVersion( );
BOOL WINAPI MImgGetWMFLoadSize( LPLONG plWid, LPLONG plHt );
BOOL WINAPI MImgGrayScale( HANDLE hImg );
BOOL WINAPI MImgHasOpacity( HANDLE hImg );
BOOL WINAPI MImgInvert( HANDLE hImg );
BOOL WINAPI MImgIsValidHandle( HANDLE hImg );
HANDLE WINAPI MImgLoadEmpty( int iWid, int iHt, COLORREF clr, WORD wBpp, DWORD dwImgType );
HANDLE WINAPI MImgLoadFromClipboard( );
HANDLE WINAPI MImgLoadFromCResource( HMODULE hModule, HRSRC hRes, DWORD dwImgType, int iFrame );
HANDLE WINAPI MImgLoadFromFile( LPTSTR psFile, DWORD dwImgType, int iFrame );
HANDLE WINAPI MImgLoadFromFileInfo( LPCTSTR psFile, DWORD dwFlags );
HANDLE WINAPI MImgLoadFromHandle( HANDLE hImg, DWORD dwImgType );
HANDLE WINAPI MImgLoadFromModule( LPCTSTR lpsModule, DWORD dwImgType, int iIdx, int iFrame );
HANDLE WINAPI MImgLoadFromModuleNamed( LPCTSTR lpsModule, DWORD dwImgType, LPCTSTR lpsResName, int iFrame );
HANDLE WINAPI MImgLoadFromPicture( HWND hWndPic );
HANDLE WINAPI MImgLoadFromResource( HSTRING hsRes, DWORD dwImgType, int iFrame );
HANDLE WINAPI MImgLoadFromScreen( HWND hWnd, DWORD dwFlags );
HANDLE WINAPI MImgLoadFromString( HSTRING hsImg, DWORD dwImgType, int iFrame );
HANDLE WINAPI MImgLoadFromStringBase64( HSTRING hsImg, DWORD dwImgType, int iFrame );
HANDLE WINAPI MImgLoadFromVisPic( LPVOID lpVisPic );
BOOL WINAPI MImgMirror( HANDLE hImg, BOOL bHorz );
BOOL WINAPI MImgPrint( HANDLE hImg, HDC hDC, LPCRECT pRect, COLORREF clrPaper, DWORD dwFlags = 0 );
BOOL WINAPI MImgPutImage( HANDLE hImg, long lX, long lY, HANDLE hImgPut );
BOOL WINAPI MImgPutText( HANDLE hImg, long lX, long lY, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr );
BOOL WINAPI MImgPutTextAligned( HANDLE hImg, LPRECT pRect, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr, DWORD dwAlignFlags );
BOOL WINAPI MImgQueryFlags( HANDLE hImg, DWORD dwFlags );
BOOL WINAPI MImgResample( HANDLE hImg, int iNewX, int iNewY, int iMode );
BOOL WINAPI MImgRotate( HANDLE hImg, float fAngle );
BOOL WINAPI MImgSave( HANDLE hImg, LPTSTR lpsFile, DWORD dwType );
BOOL WINAPI MImgSetBpp( HANDLE hImg, WORD wBpp );
BOOL WINAPI MImgSetColorOpacity( HANDLE hImg, DWORD dwTranspColor, BYTE byVal );
BOOL WINAPI MImgSetDrawClipRect( HANDLE hImg, LPRECT pRect );
BOOL WINAPI MImgSetDrawState( HANDLE hImg, DWORD dwState );
BOOL WINAPI MImgSetFlags( HANDLE hImg, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MImgSetGIFCompr( HANDLE hImg, int iVal );
BOOL WINAPI MImgSetJPGQuality( HANDLE hImg, BYTE byVal );
BOOL WINAPI MImgSetLightness( HANDLE hImg, long lLevel, long Contrast );
BOOL WINAPI MImgSetMaxOpacity( HANDLE hImg, BYTE byVal );
BOOL WINAPI MImgSetOpacity( HANDLE hImg, BYTE byFrom, BYTE byTo, BYTE byVal );
BOOL WINAPI MImgSetPaletteColor( HANDLE hImg, BYTE byIdx, DWORD dwColor );
BOOL WINAPI MImgSetPixelColor( HANDLE hImg, int iX, int iY, DWORD dwColor );
BOOL WINAPI MImgSetPixelOpacity( HANDLE hImg, int iX, int iY, BYTE byVal );
BOOL WINAPI MImgSetSize( HANDLE hImg, int iNewWid, int iNewHt, DWORD dwOverflowColor, BYTE byOverflowOpacity );
BOOL WINAPI MImgSetTIFCompr( HANDLE hImg, int iVal );
BOOL WINAPI MImgSetTranspColor( HANDLE hImg, DWORD dwColor );
BOOL WINAPI MImgSetWMFLoadSize( long lWid, long lHt );
BOOL WINAPI MImgStripOpacity( HANDLE hImg, DWORD dwColor );
BOOL WINAPI MImgSubstColor( HANDLE hImg, DWORD dwOldColor, DWORD dwNewColor );

#endif // _INC_MIMGAPI