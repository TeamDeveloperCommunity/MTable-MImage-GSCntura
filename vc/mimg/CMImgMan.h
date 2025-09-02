// CMImgMan.h

#ifndef _INC_CMIMGMAN
#define _INC_CMIMGMAN

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <list>
#include "tstring.h"
#include "CMImg.h"

// Using
using namespace std;

// Typedefs
typedef list<CMImg*> MImgList;
typedef MImgList::iterator MImgListIt;

// Klasse CMImgMan
class CMImgMan
{
private:
	MImgList m_List;
public:
	// Konstruktor
	CMImgMan( void );
	// Destruktor
	~CMImgMan( void );

	// Funktionen
	BOOL AddShadow( HANDLE hImg, int iDispX, int iDispY, BYTE bySmoothRad, DWORD dwClr, DWORD dwClrBkgnd, BYTE byIntens );
	BOOL AssignImg( HANDLE hImgSrc, HANDLE hImgDest );
	BOOL CopyImgToClipboard( HANDLE hImg );
	HANDLE CreateImgCopy( HANDLE hImg );
	HANDLE CreateImgCopyEx( HANDLE hImg, LPRECT pRect, DWORD dwType, DWORD dwFlags );
	HBITMAP CreateImgHBITMAP( HANDLE hImg, HDC hDC );
	HICON CreateImgHICON( HANDLE hImg );
	LPVOID CreateImgVisPic( HANDLE hImg, DWORD dwType );
	BOOL DeleteImg( HANDLE hImg );
	BOOL DisableImg( HANDLE hImg );
	BOOL DitherImg( HANDLE hImg, long lMethod );
	BOOL DrawImg( HANDLE hImg, HDC hDC, LPCRECT pRect, DWORD dwFlags = 0 );
	BOOL DrawImgAligned( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid = -1, int iHt = -1, DWORD dwAlignFlags = 0, DWORD dwFlags = 0 );
	BOOL DrawImgPattern( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid = -1, int iHt = -1, DWORD dwFlags = 0 );
	int EnumImages( HARRAY ha );
	WORD GetImgBpp( HANDLE hImg );
	HANDLE GetImgDIBHandle( HANDLE hImg );
	BOOL GetImgDIBString( HANDLE hImg, LPHSTRING phs );
	BOOL GetImgDrawClipRect( HANDLE hImg, LPRECT pRect );
	DWORD GetImgDrawState( HANDLE hImg );
	int GetImgFrameCount( HANDLE hImg );
	int GetImgGIFCompr( HANDLE hImg );
	BYTE GetImgJPGQuality( HANDLE hImg );
	BYTE GetImgMaxOpacity( HANDLE hImg );
	COLORREF GetImgPixelColor( HANDLE hImg, int iX, int iY );
	COLORREF GetImgPaletteColor( HANDLE hImg, BYTE byIdx );
	BYTE GetImgPixelOpacity( HANDLE hImg, int iX, int iY );
	BOOL GetImgSize( HANDLE hImg, LPSIZE pSize );
	BOOL GetImgString( HANDLE hImg, LPHSTRING phs, DWORD dwType, BOOL bBase64 = FALSE, BOOL bNullTerminated = FALSE );
	int GetImgTIFCompr( HANDLE hImg );
	DWORD GetImgTranspColor( HANDLE hImg );
	int GetImgType( HANDLE hImg );
	int GetImgUsedColors( HANDLE hImg, HARRAY haColors );
	BOOL GrayScaleImg( HANDLE hImg );
	BOOL HasImgOpacity( HANDLE hImg );
	BOOL InvertImg( HANDLE hImg );
	BOOL IsValidImgHandle( HANDLE hImg );
	HANDLE LoadImgEmpty( int iWid, int iHt, COLORREF clr, WORD wBpp, DWORD dwImgType );
	HANDLE LoadImgFromClipboard( );
	HANDLE LoadImgFromCResource( HMODULE hModule, HRSRC hRes, DWORD dwImgType, int iFrame = 0 );
	HANDLE LoadImgFromFile( LPTSTR psFile, DWORD dwImgType, int iFrame = 0 );
	HANDLE LoadImgFromFileInfo( LPCTSTR psFile, DWORD dwFlags );
	HANDLE LoadImgFromHandle( HANDLE hImg, DWORD dwImgType );
	HANDLE LoadImgFromModule( LPCTSTR lpsModule, DWORD dwImgType, int iIdx, int iFrame = 0 );
	HANDLE LoadImgFromModuleNamed( LPCTSTR lpsModule, DWORD dwImgType, LPCTSTR lpsResName, int iFrame = 0 );
	HANDLE LoadImgFromPicture( HWND hWndPic );
	HANDLE LoadImgFromResource( HSTRING hsRes, DWORD dwImgType, int iFrame = 0 );
	HANDLE LoadImgFromScreen( HWND hWnd, DWORD dwFlags );
	HANDLE LoadImgFromString( HSTRING hs, DWORD dwImgType, int iFrame = 0, BOOL bBase64 = FALSE );
	HANDLE LoadImgFromVisPic( LPVOID lpVisPic );
	BOOL MirrorImg( HANDLE hImg, BOOL bHorz );
	BOOL PrintImg( HANDLE hImg, HDC hDC, LPCRECT pRect, COLORREF clrPaper, DWORD dwFlags = 0 );
	BOOL PutImgImage( HANDLE hImg, long lX, long lY, HANDLE hImgPut );
	BOOL PutImgText( HANDLE hImg, long lX, long lY, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr );
	BOOL PutImgTextAligned( HANDLE hImg, LPRECT pRect, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr, DWORD dwAlignFlags );
	BOOL QueryImgFlags( HANDLE hImg, DWORD dwFlags );
	BOOL ResampleImg( HANDLE hImg, int iNewX, int iNewY, int iMode );
	BOOL RotateImg( HANDLE hImg, float fAngle );
	BOOL SaveImg( HANDLE hImg, LPTSTR lpsFile, DWORD dwType );
	BOOL SetImgBpp( HANDLE hImg, WORD wBpp );
	BOOL SetImgColorOpacity( HANDLE hImg, DWORD dwTranspColor, BYTE byVal );
	BOOL SetImgDrawClipRect( HANDLE hImg, LPRECT pRect );
	BOOL SetImgDrawState( HANDLE hImg, DWORD dwState );
	BOOL SetImgFlags( HANDLE hImg, DWORD dwFlags, BOOL bSet );
	BOOL SetImgGIFCompr( HANDLE hImg, int iVal );
	BOOL SetImgJPGQuality( HANDLE hImg, BYTE byVal );
	BOOL SetImgLightness( HANDLE hImg, long lLevel, long lContrast );
	BOOL SetImgMaxOpacity( HANDLE hImg, BYTE byVal );
	BOOL SetImgOpacity( HANDLE hImg, BYTE byFrom, BYTE byTo, BYTE byVal );
	BOOL SetImgPaletteColor( HANDLE hImg, BYTE byIdx, DWORD dwColor );
	BOOL SetImgPixelColor( HANDLE hImg, int iX, int iY, DWORD dwColor );
	BOOL SetImgPixelOpacity( HANDLE hImg, int iX, int iY, BYTE byVal );
	BOOL SetImgSize( HANDLE hImg, int iNewWid, int iNewHt, DWORD dwOverflowColor, BYTE byOverflowOpacity );
	BOOL SetImgTIFCompr( HANDLE hImg, int iVal );
	BOOL SetImgTranspColor( HANDLE hImg, DWORD dwColor );
	BOOL StripImgOpacity( HANDLE hImg, DWORD dwColor );
	BOOL SubstImgColor( HANDLE hImg, DWORD dwOldColor, DWORD dwNewColor );
};

#endif // _INC_CMIMGMAN