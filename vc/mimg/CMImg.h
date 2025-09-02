// CMImg.h

#ifndef _INC_CMIMG
#define _INC_CMIMG

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <shlobj.h>
#include <comdef.h>
#include <limits.h>
#include <vector>
#include <algorithm>

#include "ctd.h"
#include "vis.h"

#include "MImgConst.h"
#include "CMImgDefPal.h"

#include "ximabmp.h"
#include "ximagif.h"
#include "ximaico.h"
//#include "ximaj2k.h"
#include "ximajbg.h"
#include "ximajpg.h"
#include "ximamng.h"
#include "ximapcx.h"
#include "ximapng.h"
#include "ximatga.h"
#include "ximatif.h"
#include "ximawbmp.h"
#include "ximawmf.h"

// Using
using namespace std;

// Makros
#ifndef IS_INTRESOURCE
#define IS_INTRESOURCE(s) (((ULONG)(s) >> 16) == 0)
#endif

// Externe Variablen
extern HANDLE ghInstance;

// Externe Objekte
extern CCtd Ctd;
extern CVis Vis;

// Klasse CMImg
class CMImg
{
private:
	// Typen
	#pragma pack(2)

	typedef struct tagRESICONDIRENTRY
	{
		BYTE bWidth;
		BYTE bHeight;
		BYTE bColorCount;
		BYTE bReserved;
		USHORT wPlanes;
		USHORT wBitCount;
		UINT dwBytesInRes;
		USHORT wID;
	} RESICONDIRENTRY;

	typedef struct tagENUMRESNAME
	{
		int iIdxDesired;
		int iIdxCurrent;
		LPTSTR lpName;
	} ENUMRESNAME;

	#pragma pack()

	// Member-Variablen
	CxImage * m_pImage;

	BYTE m_byJPGQuality;
	DWORD m_dwDrawState;
	DWORD m_dwFlags;
	int m_iGIFCompr;
	int m_iTIFCompr;
	RECT m_rDrawClipRect;

	// Statische Member
	static SIZE m_siWMF;

	// Funktionen
	BOOL AlphaPaletteToAlpha( );
	BOOL CopyMembersTo( CMImg *pImg );
	BOOL CreateAlpha( );
	BOOL CreateAlphaFromPalette( );
	CxImage* CreateDisabledImgCopy( CxImage *pImgSrc = NULL );
	BOOL CreateFromHBITMAP( HBITMAP hBmColor, HBITMAP hBmMask = NULL );
	BOOL CreateFromHICON( HICON hIcon );
	void DeleteDrawPalette( );
	int DetermineType( BYTE *pBuffer, DWORD dwSize );
	BOOL Encode( CxFile *pFile, DWORD dwType );
	void FreeImageData( );
	int GetColorIndex( DWORD dwColor );
	BOOL HasPalette( );
	BOOL IsAlphaFullyOpaque( );
	BOOL LoadFromBuffer( BYTE *pBuffer, DWORD dwSize, DWORD dwImgType, int iFrame = 0 );
	void TransformXY( int &iX, int &iY );
	BOOL TranspColorToAlpha( );
public:
	// Konstruktor
	CMImg( );
	// Destruktor
	~CMImg( );
	// Funktionen
	BOOL AddShadow( int iDispX, int iDispY, BYTE bySmoothRad, DWORD dwClr, DWORD dwClrBkgnd, BYTE byIntens );
	BOOL AssignTo( CMImg * pImg );
	BOOL CopyToClipboard( );
	CMImg * CreateCopy( );
	CMImg * CreateCopyEx( LPRECT pRect, DWORD dwType, DWORD dwFlags );
	HBITMAP CreateHBITMAP( HDC hDC );
	HICON CreateHICON( );
	LPVOID CreateVisPic( DWORD dwType );
	BOOL Disable( );
	BOOL Dither( long lMethod );
	BOOL Draw( HDC hDC, LPCRECT pRect, DWORD dwFlags = 0 );
	BOOL DrawAligned( HDC hDC, LPCRECT pRect, int iWid = -1, int iHt = -1, DWORD dwAlignFlags = 0, DWORD dwFlags = 0 );
	BOOL DrawPattern( HDC hDC, LPCRECT pRect, int iWid = -1, int iHt = -1, DWORD dwFlags = 0 );
	WORD GetBpp( );
	HANDLE GetDIBHandle( );
	BOOL GetDIBString( LPHSTRING phs );
	BOOL GetDrawClipRect( LPRECT pRect );
	DWORD GetDrawState( );
	int GetFrameCount( );
	int GetGIFCompr( );
	BYTE GetJPGQuality( );
	BYTE GetMaxOpacity( );
	COLORREF GetPaletteColor( BYTE byIdx );
	COLORREF GetPixelColor( int iX, int iY );
	BYTE GetPixelOpacity( int iX, int iY );
	BOOL GetSize( LPSIZE pSize );
	BOOL GetString( LPHSTRING phs, DWORD dwType, BOOL bBase64 = FALSE, BOOL bNullTerminated = FALSE );
	int GetTIFCompr( );
	DWORD GetTranspColor( );
	int GetType( );
	int GetUsedColors( vector<COLORREF> &vColors );
	int GetUsedColors( HARRAY haColors );
	BOOL GrayScale( );
	BOOL HasOpacity( );
	BOOL Invert( );
	BOOL LoadEmpty( int iWid, int iHt, COLORREF clr, WORD wBpp, DWORD dwImgType );
	BOOL LoadFromClipboard( );
	BOOL LoadFromCResource( HMODULE hModule, HRSRC hRes, DWORD dwImgType, int iFrame = 0 );
	BOOL LoadFromFile( LPTSTR psFile, DWORD dwImgType, int iFrame = 0 );
	BOOL LoadFromFileInfo( LPCTSTR psFile, DWORD dwFlags );
	BOOL LoadFromHandle( HANDLE hImg, DWORD dwImgType );
	BOOL LoadFromModule( LPCTSTR lpsModule, DWORD dwImgType, int iIdx, int iFrame = 0 );
	BOOL LoadFromModuleNamed( LPCTSTR lpsModule, DWORD dwImgType, LPCTSTR lpsResName, int iFrame = 0 );
	BOOL LoadFromPicture( HWND hWndPic );
	BOOL LoadFromResource( HSTRING hsRes, DWORD dwImgType, int iFrame = 0 );
	BOOL LoadFromScreen( HWND hWnd, DWORD dwFlags );
	BOOL LoadFromString( HSTRING hs, DWORD dwImgType, int iFrame = 0, BOOL bBase64 = FALSE );
	BOOL LoadFromVisPic( LPVOID lpVisPic );
	BOOL Mirror( BOOL bHorz );
	BOOL Print( HDC hDC, LPCRECT pRect, COLORREF clrPaper, DWORD dwFlags = 0 );
	BOOL PutImage( long lX, long lY, CMImg * pImgPut );
	BOOL PutText( long lX, long lY, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr );
	BOOL PutTextAligned( LPRECT pRect, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr, DWORD dwAlignFlags );
	BOOL QueryFlags( DWORD dwFlags );
	BOOL Resample( int iNewX, int iNewY, int iMode );
	BOOL Rotate( float fAngle );
	BOOL Save( LPTSTR lpsFile, DWORD dwType );
	BOOL SetBpp( WORD wBpp );
	BOOL SetColorOpacity( DWORD dwTranspColor, BYTE byVal );
	BOOL SetDrawClipRect( LPRECT pRect );
	BOOL SetDrawState( DWORD dwState );
	BOOL SetFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetGIFCompr( int iVal );
	BOOL SetJPGQuality( BYTE byVal );
	BOOL SetLightness( long lLevel, long lContrast );
	BOOL SetMaxOpacity( BYTE byVal );
	BOOL SetOpacity( BYTE byFrom, BYTE byTo, BYTE byVal );
	BOOL SetPaletteColor( BYTE byIdx, DWORD dwColor );
	BOOL SetPixelColor( int iX, int iY, DWORD dwColor );
	BOOL SetPixelOpacity( int iX, int iY, BYTE byVal );
	BOOL SetSize( int iNewWid, int iNewHt, DWORD dwOverflowColor, BYTE byOverflowOpacity );
	BOOL SetTIFCompr( int iVal );
	BOOL SetTranspColor( DWORD dwColor );
	BOOL StripOpacity( DWORD dwColor );
	BOOL SubstColor( DWORD dwOldColor, DWORD dwNewColor );
	// Statische Funktionen
	static BOOL CALLBACK EnumResNameProc( HANDLE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG lParam );
	static BYTE GetFrameBpp( int iFrame );
	static DWORD GetFrameIconSize( int iFrame );
	static BOOL GetWMFLoadSize( SIZE &si );
	static BOOL SetWMFLoadSize( SIZE &si );

};

#endif // _INC_CMIMG