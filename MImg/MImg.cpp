// MImg.cpp

#include "MImg.h"

BOOL APIENTRY DllMain( HANDLE hModule, DWORD  dwReasonForCall, LPVOID lpReserved )
{
	ghInstance = hModule;

    switch ( dwReasonForCall )
	{
		case DLL_PROCESS_ATTACH:
			MImgCtlRegisterClass( );
			break;
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

// Exportierte Funktionen

// MImgAddShadow
BOOL WINAPI MImgAddShadow( HANDLE hImg, int iDispX, int iDispY, BYTE bySmoothRad, DWORD dwClr, DWORD dwClrBkgnd, BYTE byIntens )
{
	return ImgMan.AddShadow( hImg, iDispX, iDispY, bySmoothRad, dwClr, dwClrBkgnd, byIntens );
}

// MImgAssign
BOOL WINAPI MImgAssign( HANDLE hImgSrc, HANDLE hImgDest )
{
	return ImgMan.AssignImg( hImgSrc, hImgDest );
}

// MImgCopyToClipboard
BOOL WINAPI MImgCopyToClipboard( HANDLE hImg )
{
	return ImgMan.CopyImgToClipboard( hImg );
}

// MImgCreateCopy
HANDLE WINAPI MImgCreateCopy( HANDLE hImg )
{
	return ImgMan.CreateImgCopy( hImg );
}

// MImgCreateCopyEx
HANDLE WINAPI MImgCreateCopyEx( HANDLE hImg, LPRECT pRect, DWORD dwType, DWORD dwFlags )
{
	return ImgMan.CreateImgCopyEx( hImg, pRect, dwType, dwFlags );
}

// MImgCreateHBITMAP
HBITMAP WINAPI MImgCreateHBITMAP( HANDLE hImg, HDC hDC )
{
	return ImgMan.CreateImgHBITMAP( hImg, hDC );
}

// MImgCreateHICON
HICON WINAPI MImgCreateHICON( HANDLE hImg )
{
	return ImgMan.CreateImgHICON( hImg );
}

// MImgCreateVisPic
HANDLE WINAPI MImgCreateVisPic( HANDLE hImg, DWORD dwType )
{
	return ImgMan.CreateImgVisPic( hImg, dwType );
}

// MImgCtlGet
CMImgCtl * MImgCtlGet( HWND hWndImg )
{
	return (CMImgCtl*)GetWindowLongPtr( hWndImg, 0 );
}

// MImgCtlGetFitMode
int WINAPI MImgCtlGetFitMode( HWND hWndImg )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->GetFitMode( );
}

// MImgCtlGetBackBrush
BOOL WINAPI MImgCtlGetBackBrush( HWND hWndImg, LPINT piStyle, LPCOLORREF pclr )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->GetBackBrush( piStyle, pclr );
}

// MImgCtlGetContentsExtent
BOOL WINAPI MImgCtlGetContentsExtent( HWND hWndImg, LPSIZE ps )
{
	if ( !ps ) return FALSE;

	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->GetContentsExtent( *ps );
}

// MImgCtlGetImage
HANDLE WINAPI MImgCtlGetImage( HWND hWndImg )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->GetImage( );
}

// MImgCtlGetImagePos
BOOL WINAPI MImgCtlGetImagePos( HWND hWndImg, LPPOINT ppt )
{
	if ( !ppt ) return FALSE;

	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->GetImagePos( *ppt );
}

// MImgCtlGetImagePixFromPoint
BOOL WINAPI MImgCtlGetImagePixFromPoint( HWND hWndImg, LONG lX, LONG lY, LPPOINT pPt )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	POINT pt;
	pt.x = lX;
	pt.y = lY;
	return pCtl->GetImagePixFromPoint( pt, pPt );
}

// MImgCtlGetScale
double WINAPI MImgCtlGetScale( HWND hWndImg, int iWhat )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return 0.0;

	return pCtl->GetScale( iWhat );
}

// MImgCtlGetTextPos
BOOL WINAPI MImgCtlGetTextPos( HWND hWndImg, LPPOINT ppt )
{
	if ( !ppt ) return FALSE;

	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->GetTextPos( *ppt );
}

// MImgCtlQueryAlignFlags
BOOL WINAPI MImgCtlQueryAlignFlags( HWND hWndImg, DWORD dwFlags )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->QueryAlignFlags( dwFlags );
}

// MImgCtlQueryFlags
BOOL WINAPI MImgCtlQueryFlags( HWND hWndImg, DWORD dwFlags )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->QueryFlags( dwFlags );
}

// MImgCtlRefresh
BOOL WINAPI MImgCtlRefresh( HWND hWndImg )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->Refresh( );
}

// MImgCtlRegisterClass
BOOL MImgCtlRegisterClass( )
{
	WNDCLASSEX wc;
	wc.cbSize = sizeof( WNDCLASSEX );

	BOOL bRegistered = GetClassInfoEx( (HINSTANCE)ghInstance, MIMG_CTLCLS, &wc );

	if ( !bRegistered )
	{
		// Fensterklasse registrieren
		wc.style = CS_OWNDC;
		wc.lpfnWndProc = MImgCtlWndProc;
		wc.cbClsExtra = 0;
		wc.cbWndExtra = 4;
		wc.hInstance = (HINSTANCE)ghInstance;
		wc.hIcon = NULL;
		wc.hCursor = LoadCursor( NULL, IDC_ARROW );
		wc.hbrBackground = NULL;
		wc.lpszMenuName = NULL;
		wc.lpszClassName = MIMG_CTLCLS;
		wc.hIconSm = NULL;

		bRegistered = ( RegisterClassEx( &wc ) != NULL );
	}

	return bRegistered;
}

// MImgCtlSetAlignFlags
BOOL WINAPI MImgCtlSetAlignFlags( HWND hWndImg, DWORD dwFlags, BOOL bSet )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SetAlignFlags( dwFlags, bSet );
}

// MImgCtlSetBackBrush
BOOL WINAPI MImgCtlSetBackBrush( HWND hWndImg, int iStyle, COLORREF clr )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SetBackBrush( iStyle, clr );
}

// MImgCtlSetFitMode
BOOL WINAPI MImgCtlSetFitMode( HWND hWndImg, int iFitMode )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SetFitMode( iFitMode );
}

// MImgCtlSetFlags
BOOL WINAPI MImgCtlSetFlags( HWND hWndImg, DWORD dwFlags, BOOL bSet )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SetFlags( dwFlags, bSet );
}

// MImgCtlSetImage
BOOL WINAPI MImgCtlSetImage( HWND hWndImg, HANDLE hImg, BOOL bOwn )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SetImage( hImg, bOwn );
}

// MImgCtlSetImagePos
BOOL WINAPI MImgCtlSetImagePos( HWND hWndImg, LONG lX, LONG lY )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	POINT pt;
	pt.x = lX;
	pt.y = lY;
	return pCtl->SetImagePos( pt );
}

// MImgCtlSetScale
BOOL WINAPI MImgCtlSetScale( HWND hWndImg, int iWhat, double dScale )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SetScale( iWhat, dScale );
}

// MImgCtlSetTextPos
BOOL WINAPI MImgCtlSetTextPos( HWND hWndImg, LONG lX, LONG lY )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	POINT pt;
	pt.x = lX;
	pt.y = lY;
	return pCtl->SetTextPos( pt );
}

// MImgCtlSizeToImage
BOOL WINAPI MImgCtlSizeToContents( HWND hWndImg )
{
	CMImgCtl *pCtl = MImgCtlGet( hWndImg );
	if ( !pCtl ) return FALSE;

	return pCtl->SizeToContents( );
}

// MImgCtlWndProc
LRESULT CALLBACK MImgCtlWndProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	CMImgCtl *pCtl = NULL;
	switch ( uMsg )
	{
	case WM_NCCREATE:
		pCtl = new CMImgCtl( hWnd );
		break;
	default:
		pCtl = MImgCtlGet( hWnd );
	}

	if ( pCtl )
		return pCtl->WndProc( hWnd, uMsg, wParam, lParam );
	else
		return DefWindowProc( hWnd, uMsg, wParam, lParam );
}

// MImgDelete
BOOL WINAPI MImgDelete( HANDLE hImg )
{
	return ImgMan.DeleteImg( hImg );
}

// MImgDisable
BOOL WINAPI MImgDisable( HANDLE hImg )
{
	return ImgMan.DisableImg( hImg );
}

// MImgDither
BOOL WINAPI MImgDither( HANDLE hImg, long lMethod )
{
	return ImgMan.DitherImg( hImg, lMethod );
}

// MImgDraw
BOOL WINAPI MImgDraw( HANDLE hImg, HDC hDC, LPCRECT pRect, DWORD dwFlags /*= 0*/ )
{
	return ImgMan.DrawImg( hImg, hDC, pRect, dwFlags );
}

// MImgDrawAligned
BOOL WINAPI MImgDrawAligned( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid /*= -1*/, int iHt /*= -1*/, DWORD dwAlignFlags /*= 0*/, DWORD dwFlags /*= 0*/ )
{
	return ImgMan.DrawImgAligned( hImg, hDC, pRect, iWid, iHt, dwAlignFlags, dwFlags );
}

// MImgDrawPattern
BOOL WINAPI MImgDrawPattern( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid /*= -1*/, int iHt /*= -1*/, DWORD dwFlags /*= 0*/ )
{
	return ImgMan.DrawImgPattern( hImg, hDC, pRect, iWid, iHt, dwFlags );
}

// MImgEnum
int WINAPI MImgEnum( HARRAY ha )
{
	return ImgMan.EnumImages( ha );	
}

// MImgGetBpp
WORD WINAPI MImgGetBpp( HANDLE hImg )
{
	return ImgMan.GetImgBpp( hImg );
}

// MImgGetDIBHandle
HANDLE WINAPI MImgGetDIBHandle( HANDLE hImg )
{
	return ImgMan.GetImgDIBHandle( hImg );
}

// MImgGetDIBString
BOOL WINAPI MImgGetDIBString( HANDLE hImg, LPHSTRING phs )
{
	return ImgMan.GetImgDIBString( hImg, phs );
}

// MImgGetDrawClipRect
BOOL WINAPI MImgGetDrawClipRect( HANDLE hImg, LPRECT pRect )
{
	return ImgMan.GetImgDrawClipRect( hImg, pRect );
}

// MImgGetDrawState
DWORD WINAPI MImgGetDrawState( HANDLE hImg )
{
	return ImgMan.GetImgDrawState( hImg );
}

// MImgGetFrameCount
int WINAPI MImgGetFrameCount( HANDLE hImg )
{
	return ImgMan.GetImgFrameCount( hImg );
}

// MImgGetGIFCompr
int WINAPI MImgGetGIFCompr( HANDLE hImg )
{
	return ImgMan.GetImgGIFCompr( hImg );
}

// MImgGetJPGQuality
BYTE WINAPI MImgGetJPGQuality( HANDLE hImg )
{
	return ImgMan.GetImgJPGQuality( hImg );
}

// MImgGetMaxOpacity
BYTE WINAPI MImgGetMaxOpacity( HANDLE hImg )
{
	return ImgMan.GetImgMaxOpacity( hImg );
}

// MImgGetPaletteColor
COLORREF WINAPI MImgGetPaletteColor( HANDLE hImg, BYTE byIdx )
{
	return ImgMan.GetImgPaletteColor( hImg, byIdx );
}

// MImgGetPixelColor
COLORREF WINAPI MImgGetPixelColor( HANDLE hImg, int iX, int iY )
{
	return ImgMan.GetImgPixelColor( hImg, iX, iY );
}

// MImgGetPixelOpacity
BYTE WINAPI MImgGetPixelOpacity( HANDLE hImg, int iX, int iY )
{
	return ImgMan.GetImgPixelOpacity( hImg, iX, iY );
}

// MImgGetSize
BOOL WINAPI MImgGetSize( HANDLE hImg, LPSIZE pSize )
{
	return ImgMan.GetImgSize( hImg, pSize );
}

// MImgGetString
BOOL WINAPI MImgGetString( HANDLE hImg, LPHSTRING phs, DWORD dwType )
{
	return ImgMan.GetImgString( hImg, phs, dwType );
}

// MImgGetStringNT
BOOL WINAPI MImgGetStringNT( HANDLE hImg, LPHSTRING phs, DWORD dwType )
{
	return ImgMan.GetImgString( hImg, phs, dwType, FALSE, TRUE );
}

// MImgGetStringBase64
BOOL WINAPI MImgGetStringBase64( HANDLE hImg, LPHSTRING phs, DWORD dwType )
{
	return ImgMan.GetImgString( hImg, phs, dwType, TRUE );
}

// MImgGetTIFCompr
int WINAPI MImgGetTIFCompr( HANDLE hImg )
{
	return ImgMan.GetImgTIFCompr( hImg );
}

// MImgGetTranspColor
DWORD WINAPI MImgGetTranspColor( HANDLE hImg )
{
	return ImgMan.GetImgTranspColor( hImg );
}

// MImgGetType
int WINAPI MImgGetType( HANDLE hImg )
{
	return ImgMan.GetImgType( hImg );
}

// MImgGetUsedColors
int WINAPI MImgGetUsedColors( HANDLE hImg, HARRAY haColors )
{
	return ImgMan.GetImgUsedColors( hImg, haColors );
}

/*
// MImgGetVersion
HSTRING WINAPI MImgGetVersion( )
{
	HSTRING hsVer = 0;
	LPTSTR pBuf = NULL;
	long lLen;

	Ctd.InitLPHSTRINGParam(&hsVer, Ctd.BufLenFromStrLen(lstrlen(MIMG_VERSION)));
	pBuf = Ctd.StringGetBuffer( hsVer, &lLen );
	lstrcpy( pBuf, MIMG_VERSION );
	return hsVer;
}
*/

HSTRING WINAPI MImgGetVersion()
{
	TCHAR szFilename[MAX_PATH + 1] = { 0 };
	TCHAR szVersion[100] = { 0 };

	if (GetModuleFileName((HMODULE)ghInstance, szFilename, MAX_PATH))
	{
		DWORD dwTemp;
		DWORD dwSize = GetFileVersionInfoSize(szFilename, &dwTemp);
		if (dwSize > 0)
		{
			std::vector<BYTE> data(dwSize);
			if (GetFileVersionInfo(szFilename, NULL, dwSize, &data[0]))
			{
				UINT uiVerLen = 0;
				VS_FIXEDFILEINFO* pFixedInfo = 0;
				if (VerQueryValue(&data[0], TEXT("\\"), (void**)&pFixedInfo, (UINT *)&uiVerLen))
				{
					//_stprintf(szVersion, _T("%u.%u.%u.%u"), HIWORD(pFixedInfo->dwProductVersionMS), LOWORD(pFixedInfo->dwProductVersionMS), HIWORD(pFixedInfo->dwProductVersionLS), LOWORD(pFixedInfo->dwProductVersionLS));
					_stprintf(szVersion, _T("%u.%u.%u"), HIWORD(pFixedInfo->dwProductVersionMS), LOWORD(pFixedInfo->dwProductVersionMS), HIWORD(pFixedInfo->dwProductVersionLS));
				}
			}
		}
	}

	HSTRING hsVer = 0;
	LPTSTR pBuf;
	long lLen;

	Ctd.InitLPHSTRINGParam(&hsVer, Ctd.BufLenFromStrLen(lstrlen(szVersion)));
	pBuf = Ctd.StringGetBuffer(hsVer, &lLen);
	if (pBuf)
	{
		lstrcpy(pBuf, szVersion);
		Ctd.HStrUnlock(hsVer);
	}	

	return hsVer;
}

// MImgGetWMFLoadSize
BOOL WINAPI MImgGetWMFLoadSize( LPLONG plWid, LPLONG plHt )
{
	if ( !plWid && plHt )
		return FALSE;

	SIZE si;
	if ( !CMImg::GetWMFLoadSize( si ) )
		return FALSE;

	*plWid = si.cx;
	*plHt = si.cy;

	return TRUE;
}

// MImgGrayScale
BOOL WINAPI MImgGrayScale( HANDLE hImg )
{
	return ImgMan.GrayScaleImg( hImg );
}

// MImgHasOpacity
BOOL WINAPI MImgHasOpacity( HANDLE hImg )
{
	return ImgMan.HasImgOpacity( hImg );
}

// MImgInvert
BOOL WINAPI MImgInvert( HANDLE hImg )
{
	return ImgMan.InvertImg( hImg );
}

// MImgIsValidHandle
BOOL WINAPI MImgIsValidHandle( HANDLE hImg )
{
	return ImgMan.IsValidImgHandle( hImg );
}

// MImgSetLightness
BOOL WINAPI MImgSetLightness( HANDLE hImg, long lLevel, long lContrast )
{
	return ImgMan.SetImgLightness( hImg, lLevel, lContrast );
}

// MImgLoadEmpty
HANDLE WINAPI MImgLoadEmpty( int iWid, int iHt, COLORREF clr, WORD wBpp, DWORD dwImgType )
{
	return ImgMan.LoadImgEmpty( iWid, iHt, clr, wBpp, dwImgType );
}

// MImgLoadFromClipboard
HANDLE WINAPI MImgLoadFromClipboard( )
{
	return ImgMan.LoadImgFromClipboard( );
}

// MImgLoadFromCResource
HANDLE WINAPI MImgLoadFromCResource( HMODULE hModule, HRSRC hRes, DWORD dwImgType, int iFrame )
{
	return ImgMan.LoadImgFromCResource( hModule, hRes, dwImgType, iFrame );
}

// MImgLoadFromFile
HANDLE WINAPI MImgLoadFromFile( LPTSTR psFile, DWORD dwImgType, int iFrame )
{
	return ImgMan.LoadImgFromFile( psFile, dwImgType, iFrame );
}

// MImgLoadFromFileInfo
HANDLE WINAPI MImgLoadFromFileInfo( LPCTSTR psFile, DWORD dwFlags )
{
	return ImgMan.LoadImgFromFileInfo( psFile, dwFlags );
}

// MImgLoadFromHandle
HANDLE WINAPI MImgLoadFromHandle( HANDLE hImg, DWORD dwImgType )
{
	return ImgMan.LoadImgFromHandle( hImg, dwImgType );
}

// MImgLoadFromModule
HANDLE WINAPI MImgLoadFromModule( LPCTSTR lpsModule, DWORD dwImgType, int iIdx, int iFrame )
{
	return ImgMan.LoadImgFromModule( lpsModule, dwImgType, iIdx, iFrame );
}

// MImgLoadFromModuleNamed
HANDLE WINAPI MImgLoadFromModuleNamed( LPCTSTR lpsModule, DWORD dwImgType, LPCTSTR lpsResName, int iFrame )
{
	return ImgMan.LoadImgFromModuleNamed( lpsModule, dwImgType, lpsResName, iFrame );
}

// MImgLoadFromPicture
HANDLE WINAPI MImgLoadFromPicture( HWND hWndPic )
{
	return ImgMan.LoadImgFromPicture( hWndPic );
}

// MImgLoadFromResource
HANDLE WINAPI MImgLoadFromResource( HSTRING hsRes, DWORD dwImgType, int iFrame )
{
	return ImgMan.LoadImgFromResource( hsRes, dwImgType, iFrame );
}

// MImgLoadFromScreen
HANDLE WINAPI MImgLoadFromScreen( HWND hWnd, DWORD dwFlags )
{
	return ImgMan.LoadImgFromScreen( hWnd, dwFlags );
}

// MImgLoadFromString
HANDLE WINAPI MImgLoadFromString( HSTRING hsImg, DWORD dwImgType, int iFrame )
{
	return ImgMan.LoadImgFromString( hsImg, dwImgType, iFrame );
}

// MImgLoadFromStringBase64
HANDLE WINAPI MImgLoadFromStringBase64( HSTRING hsImg, DWORD dwImgType, int iFrame )
{
	return ImgMan.LoadImgFromString( hsImg, dwImgType, iFrame, TRUE );
}

// MImgLoadFromVisPic
HANDLE WINAPI MImgLoadFromVisPic( LPVOID lpVisPic )
{
	return ImgMan.LoadImgFromVisPic( lpVisPic );
}

// MImgMirror
BOOL WINAPI MImgMirror( HANDLE hImg, BOOL bHorz )
{
	return ImgMan.MirrorImg( hImg, bHorz );
}

// MImgPrint
BOOL WINAPI MImgPrint( HANDLE hImg, HDC hDC, LPCRECT pRect, COLORREF clrPaper, DWORD dwFlags )
{
	return ImgMan.PrintImg( hImg, hDC, pRect, clrPaper, dwFlags );
}

// MImgPutImage
BOOL WINAPI MImgPutImage( HANDLE hImg, long lX, long lY, HANDLE hImgPut )
{
	return ImgMan.PutImgImage( hImg, lX, lY, hImgPut );
}

// MImgPutText
BOOL WINAPI MImgPutText( HANDLE hImg, long lX, long lY, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr )
{
	return ImgMan.PutImgText( hImg, lX, lY, lpsText, lpsFontName, lFontSize, dwFontEnh, clr );
}

// MImgPutTextAligned
BOOL WINAPI MImgPutTextAligned( HANDLE hImg, LPRECT pRect, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr, DWORD dwAlignFlags )
{
	return ImgMan.PutImgTextAligned( hImg, pRect, lpsText, lpsFontName, lFontSize, dwFontEnh, clr, dwAlignFlags );
}

// MImgQueryFlags
BOOL WINAPI MImgQueryFlags( HANDLE hImg, DWORD dwFlags )
{
	return ImgMan.QueryImgFlags( hImg, dwFlags );
}

// MImgResample
BOOL WINAPI MImgResample( HANDLE hImg, int iNewX, int iNewY, int iMode )
{
	return ImgMan.ResampleImg( hImg, iNewX, iNewY, iMode );
}

// MImgRotate
BOOL WINAPI MImgRotate( HANDLE hImg, float fAngle )
{
	return ImgMan.RotateImg( hImg, fAngle );
}

// MImgSave
BOOL WINAPI MImgSave( HANDLE hImg, LPTSTR lpsFile, DWORD dwType )
{
	return ImgMan.SaveImg( hImg, lpsFile, dwType );
}

// MImgSetBpp
BOOL WINAPI MImgSetBpp( HANDLE hImg, WORD wBpp )
{
	return ImgMan.SetImgBpp( hImg, wBpp );
}

// MImgSetColorTransp
BOOL WINAPI MImgSetColorOpacity( HANDLE hImg, DWORD dwTranspColor, BYTE byVal )
{
	return ImgMan.SetImgColorOpacity( hImg, dwTranspColor, byVal );
}

// MImgSetDrawClipRect
BOOL WINAPI MImgSetDrawClipRect( HANDLE hImg, LPRECT pRect )
{
	return ImgMan.SetImgDrawClipRect( hImg, pRect );
}

// MImgSetDrawState
BOOL WINAPI MImgSetDrawState( HANDLE hImg, DWORD dwState )
{
	return ImgMan.SetImgDrawState( hImg, dwState );
}

// MImgSetFlags
BOOL WINAPI MImgSetFlags( HANDLE hImg, DWORD dwFlags, BOOL bSet )
{
	return ImgMan.SetImgFlags( hImg, dwFlags, bSet );
}

// MImgSetGIFCompr
BOOL WINAPI MImgSetGIFCompr( HANDLE hImg, int iVal )
{
	return ImgMan.SetImgGIFCompr( hImg, iVal );
}

// MImgSetJPGQuality
BOOL WINAPI MImgSetJPGQuality( HANDLE hImg, BYTE byVal )
{
	return ImgMan.SetImgJPGQuality( hImg, byVal );
}

// MImgSetMaxOpacity
BOOL WINAPI MImgSetMaxOpacity( HANDLE hImg, BYTE byVal )
{
	return ImgMan.SetImgMaxOpacity( hImg, byVal );
}

// MImgSetOpacity
BOOL WINAPI MImgSetOpacity( HANDLE hImg, BYTE byFrom, BYTE byTo, BYTE byVal )
{
	return ImgMan.SetImgOpacity( hImg, byFrom, byTo, byVal );
}

// MImgSetPaletteColor
BOOL WINAPI MImgSetPaletteColor( HANDLE hImg, BYTE byIdx, DWORD dwColor )
{
	return ImgMan.SetImgPaletteColor( hImg, byIdx, dwColor );
}

// MImgSetPixelColor
BOOL WINAPI MImgSetPixelColor( HANDLE hImg, int iX, int iY, DWORD dwColor )
{
	return ImgMan.SetImgPixelColor( hImg, iX, iY, dwColor );
}

// MImgSetPixelOpacity
BOOL WINAPI MImgSetPixelOpacity( HANDLE hImg, int iX, int iY, BYTE byVal )
{
	return ImgMan.SetImgPixelOpacity( hImg, iX, iY, byVal );
}

// MImgSetSize
BOOL WINAPI MImgSetSize( HANDLE hImg, int iNewWid, int iNewHt, DWORD dwOverflowColor, BYTE byOverflowOpacity )
{
	return ImgMan.SetImgSize( hImg, iNewWid, iNewHt, dwOverflowColor, byOverflowOpacity );
}

// MImgSetTIFCompr
BOOL WINAPI MImgSetTIFCompr( HANDLE hImg, int iVal )
{
	return ImgMan.SetImgTIFCompr( hImg, iVal );
}

// MImgSetTranspColor
BOOL WINAPI MImgSetTranspColor( HANDLE hImg, DWORD dwColor )
{
	return ImgMan.SetImgTranspColor( hImg, dwColor );
}

// MImgSetWMFLoadSize
BOOL WINAPI MImgSetWMFLoadSize( long lWid, long lHt )
{
	SIZE si = {lWid,lHt};
	return CMImg::SetWMFLoadSize( si );
}

// MImgStripOpacity
BOOL WINAPI MImgStripOpacity( HANDLE hImg, DWORD dwColor )
{
	return ImgMan.StripImgOpacity( hImg, dwColor );
}

// MImgSubstColor
BOOL WINAPI MImgSubstColor( HANDLE hImg, DWORD dwOldColor, DWORD dwNewColor )
{
	return ImgMan.SubstImgColor( hImg, dwOldColor, dwNewColor );
}
