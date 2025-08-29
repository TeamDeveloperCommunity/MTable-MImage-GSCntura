// CMImgMan.cpp

// Includes
#include "CMImgMan.h"

// Konstruktor
CMImgMan::CMImgMan( )
{
}


// Destruktor
CMImgMan::~CMImgMan( void )
{
	// Bilder löschen
	if ( m_List.size( ) > 0 ) 
	{
		CMImg *pImg;
		MImgListIt it;
		for ( it = m_List.begin( ); it != m_List.end( ); ++it )
		{
			pImg = *it;
			delete pImg;
		}
	}
}

// AddImgShadow
// Fügt Schatten zu einem Bild hinzu
BOOL CMImgMan::AddShadow( HANDLE hImg, int iDispX, int iDispY, BYTE bySmoothRad, DWORD dwClr, DWORD dwClrBkgnd, BYTE byIntens )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return FALSE;

	return pImg->AddShadow( iDispX, iDispY, bySmoothRad, dwClr, dwClrBkgnd, byIntens );
}

// AssignImg
// Weist ein Bild einem andern zu
BOOL CMImgMan::AssignImg( HANDLE hImgSrc, HANDLE hImgDest )
{
	CMImg * pImgSrc = (CMImg*)hImgSrc;
	if ( !pImgSrc )
		return FALSE;

	CMImg * pImgDest = (CMImg*)hImgDest;
	if ( !pImgDest )
		return FALSE;

	return pImgSrc->AssignTo( pImgDest );
}

// CopyImgToClipboard
// Kopiert ein Bild mit dem Format CF_DIB in die Zwischenablage
BOOL CMImgMan::CopyImgToClipboard( HANDLE hImg )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return FALSE;

	return pImg->CopyToClipboard( );
}

// CreateImgCopy
// Erstellt Kopie eines Bildes
HANDLE CMImgMan::CreateImgCopy( HANDLE hImg )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return NULL;

	CMImg * pImgCopy = pImg->CreateCopy( );
	if ( !pImgCopy )
		return NULL;

	m_List.push_back( pImgCopy );

	return HANDLE( pImgCopy );
}

// CreateImgCopyEx
// Erstellt Kopie eines Bildes mit erw. Möglichkeiten
HANDLE CMImgMan::CreateImgCopyEx( HANDLE hImg, LPRECT pRect, DWORD dwType, DWORD dwFlags )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return NULL;

	CMImg * pImgCopy = pImg->CreateCopyEx( pRect, dwType, dwFlags );
	if ( !pImgCopy )
		return NULL;

	m_List.push_back( pImgCopy );

	return HANDLE( pImgCopy );
}

// CreateImgHBITMAP
// Erzeugt ein HBITMAP aus einem Bild
HBITMAP CMImgMan::CreateImgHBITMAP( HANDLE hImg, HDC hDC )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return NULL;

	return pImg->CreateHBITMAP( hDC );
}

// CreateImgHICON
// Erzeugt ein HICON aus einem Bild
HICON CMImgMan::CreateImgHICON(HANDLE hImg)
{
	CMImg * pImg = (CMImg*)hImg;
	if (!pImg) return NULL;

	return pImg->CreateHICON();
}

// CreateImgVisPic
// Erstellt ein Visual Toolchest - Picture aus einem Bild
LPVOID CMImgMan::CreateImgVisPic( HANDLE hImg, DWORD dwType )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return NULL;

	return pImg->CreateVisPic( dwType );
}

// DeleteImg
// Löscht ein Bild
BOOL CMImgMan::DeleteImg( HANDLE hImg )
{
	CMImg * pImg = (CMImg*)hImg;
	if ( !pImg ) return FALSE;

	MImgListIt itEnd = m_List.end( );
	MImgListIt it = find( m_List.begin( ), itEnd, pImg );
	if ( it != m_List.end( ) )
	{
		delete *it;
		m_List.erase( it );
		return TRUE;
	}
	else
		return FALSE;
}

// DisableImg
// Disabled ein Bild
BOOL CMImgMan::DisableImg( HANDLE hImg )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Disable( );
}

// DitherImg
// Dithert ein Bild
BOOL CMImgMan::DitherImg( HANDLE hImg, long lMethod )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Dither( lMethod );
}
// DrawImg
// Zeichnet ein Bild
BOOL CMImgMan::DrawImg( HANDLE hImg, HDC hDC, LPCRECT pRect, DWORD dwFlags /*= 0*/ )
{
	if ( !( hImg && hDC && pRect ) )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Draw( hDC, pRect, dwFlags );
}

// DrawImgAligned
// Zeichnet ein Bild ausgerichtet in einem Bereich
BOOL CMImgMan::DrawImgAligned( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid /*= -1*/, int iHt /*= -1*/, DWORD dwAlignFlags /*= 0*/, DWORD dwFlags /*= 0*/ )
{
	if ( !( hImg && hDC && pRect ) )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->DrawAligned( hDC, pRect, iWid, iHt, dwAlignFlags, dwFlags );
}

// DrawImgPattern
// Füllt einen Bereich mit einem Bild
BOOL CMImgMan::DrawImgPattern( HANDLE hImg, HDC hDC, LPCRECT pRect, int iWid /*= -1*/, int iHt /*= -1*/, DWORD dwFlags /*= 0*/ )
{
	if ( !( hImg && hDC && pRect ) )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->DrawPattern( hDC, pRect, iWid, iHt, dwFlags );
}

// EnumImages
// Liefert die Handles aller Bilder
int CMImgMan::EnumImages( HARRAY ha )
{
	long lDimCount;
	if ( !Ctd.ArrayDimCount( ha, &lDimCount ) ) return -1;
	if ( lDimCount != 1 ) return -1;

	long lLowerBound;
	if ( !Ctd.ArrayGetLowerBound( ha, 1, &lLowerBound ) ) return -1;

	int iCount = 0;
	MImgListIt it;
	MImgListIt itEnd = m_List.end( );
	NUMBER n;
	for ( it = m_List.begin( ); it != itEnd; ++it )
	{
		if ( !(*it)->QueryFlags( MIMG_FLAG_PRIVATE ) )
		{
			if ( !Ctd.CvtULongToNumber( DWORD( *it ), &n ) ) return -1;
			if ( !Ctd.MDArrayPutNumber( ha, &n, lLowerBound + iCount ) ) return -1;
			++iCount;
		}
	}

	return iCount;
}

// GetImgBpp
// Liefert die Bits pro Pixel eines Bildes
WORD CMImgMan::GetImgBpp( HANDLE hImg )
{
	if ( !hImg )
		return 0;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetBpp( );
}

// GetImgGIFCompr
// Liefert die GIF-Kompression
int CMImgMan::GetImgGIFCompr( HANDLE hImg )
{
	if ( !hImg )
		return -1;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetGIFCompr( );
}

// GetImgTIFCompr
// Liefert die TIF-Kompression
int CMImgMan::GetImgTIFCompr( HANDLE hImg )
{
	if ( !hImg )
		return -1;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetTIFCompr( );
}

// GetImgDIBHandle
// Liefert DIB-Handle
HANDLE CMImgMan::GetImgDIBHandle( HANDLE hImg )
{
	if ( !hImg )
		return NULL;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetDIBHandle( );
}

// GetImgDIBString
// Liefert ein Pixel eines Bildes
BOOL CMImgMan::GetImgDIBString( HANDLE hImg, LPHSTRING phs )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetDIBString( phs );
}

// GetImgDrawClipRect
// Liefert das Zeichnungs-Clip-Rechteck
BOOL CMImgMan::GetImgDrawClipRect( HANDLE hImg, LPRECT pRect )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetDrawClipRect( pRect );
}

// GetImgDrawState
// Liefert den Zeichnungs-Status eines Bildes
DWORD CMImgMan::GetImgDrawState( HANDLE hImg )
{
	if ( !hImg )
		return MIMG_DST_NONE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetDrawState( );
}

// GetImgFrameCount
// Liefert die Gesamtanzahl Frames eines Bildes
int CMImgMan::GetImgFrameCount( HANDLE hImg )
{
	if ( !hImg )
		return MIMG_DST_NONE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetFrameCount( );
}

// GetImgJPGQuality
// Setzt die Bildqualität
BYTE CMImgMan::GetImgJPGQuality( HANDLE hImg )
{
	if ( !hImg )
		return 0;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetJPGQuality( );
}

// GetImgMaxOpacity
// Liefert die max. Durchsichtigkeit eines Bildes
BYTE CMImgMan::GetImgMaxOpacity( HANDLE hImg )
{
	if ( !hImg )
		return 0;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetMaxOpacity( );
}

// GetImgPaletteColor
// Liefert die Farbe eines Paletteneintrags
COLORREF CMImgMan::GetImgPaletteColor( HANDLE hImg, BYTE byIdx )
{
	if ( !hImg )
		return MIMG_COLOR_UNDEF;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetPaletteColor( byIdx );
}

// GetImgPixelColor
// Liefert die Farbe eines Bild-Pixels
COLORREF CMImgMan::GetImgPixelColor( HANDLE hImg, int iX, int iY )
{
	if ( !hImg )
		return MIMG_COLOR_UNDEF;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetPixelColor( iX, iY );
}

// GetImgPixelOpacity
// Liefert die Durchlässigkeit eines Bild-Pixels
BYTE CMImgMan::GetImgPixelOpacity( HANDLE hImg, int iX, int iY )
{
	if ( !hImg )
		return 255;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetPixelOpacity( iX, iY );
}

// GetImgSize
// Liefert die Größe eines Bildes
BOOL CMImgMan::GetImgSize( HANDLE hImg, LPSIZE pSize )
{
	if ( !( hImg && pSize ) )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetSize( pSize );
}

// GetImgString
// Liefert ein Bild als String
BOOL CMImgMan::GetImgString( HANDLE hImg, LPHSTRING phs, DWORD dwType, BOOL bBase64 /*= FALSE*/, BOOL bNullTerminated /*= FALSE*/ )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetString( phs, dwType, bBase64, bNullTerminated );
}

// GetImgTranspColor
// Liefert die transparente Farbe eines Bildes
DWORD CMImgMan::GetImgTranspColor( HANDLE hImg )
{
	if ( !hImg )
		return MIMG_COLOR_UNDEF;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetTranspColor( );
}

// GetImgType
// Liefert den Typ eines Bildes
int CMImgMan::GetImgType( HANDLE hImg )
{
	if ( !hImg )
		return MIMG_TYPE_UNKNOWN;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetType( );
}

// GetImgUsedColors
// Liefert die verwendeten Farben eines Bildes
int CMImgMan::GetImgUsedColors( HANDLE hImg, HARRAY haColors )
{
	if ( !hImg )
		return -1;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GetUsedColors( haColors );
}

// GrayScaleImg
// Konvertiert ein Bild in Graustufen
BOOL CMImgMan::GrayScaleImg( HANDLE hImg )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->GrayScale( );
}

// HasImgOpacity
// Ermittelt, ob ein Bild Transparenz enthält
BOOL CMImgMan::HasImgOpacity( HANDLE hImg )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->HasOpacity( );
}

// InvertImg
// Invertiert ein Bild
BOOL CMImgMan::InvertImg( HANDLE hImg )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Invert( );
}

// IsValidImgHandle
// Ermittelt, ob ein Bild-Handle gültig ist
BOOL CMImgMan::IsValidImgHandle( HANDLE hImg )
{
	if ( !hImg )
		return FALSE;

	return find( m_List.begin( ), m_List.end( ), (CMImg*)hImg ) != m_List.end( );
}

// LoadImgEmpty
// Lädt ein "leeres" Bild
HANDLE CMImgMan::LoadImgEmpty( int iWid, int iHt, COLORREF clr, WORD wBpp, DWORD dwImgType )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadEmpty( iWid, iHt, clr, wBpp, dwImgType ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromClipboard
// Lädt ein Bild aus der Zwischenablage
HANDLE CMImgMan::LoadImgFromClipboard( )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromClipboard( ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromCResource
// Lädt ein Bild aus einer C-Ressource
HANDLE CMImgMan::LoadImgFromCResource( HMODULE hModule, HRSRC hRes, DWORD dwImgType, int iFrame /* = 0 */ )
{
	if ( !hRes ) return NULL;

	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromCResource( hModule, hRes, dwImgType, iFrame ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromFile
// Lädt ein Bild aus einer Datei und fügt es der internen Liste hinzu
HANDLE CMImgMan::LoadImgFromFile( LPTSTR psFile, DWORD dwImgType, int iFrame /* = 0 */ )
{
	if ( !psFile ) return NULL;

	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromFile( psFile, dwImgType, iFrame ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromFileInfo
// Lädt das einer Datei zugepordnete Bild
HANDLE CMImgMan::LoadImgFromFileInfo( LPCTSTR psFile, DWORD dwFlags )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromFileInfo( psFile, dwFlags ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromHandle
// Lädt ein Bild aus anhand eines Handles
HANDLE CMImgMan::LoadImgFromHandle( HANDLE hImg, DWORD dwImgType )
{
	if ( !hImg ) return NULL;

	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromHandle( hImg, dwImgType ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromModule
// Lädt ein Bild aus einem Modul
HANDLE CMImgMan::LoadImgFromModule( LPCTSTR lpsModule, DWORD dwImgType, int iIdx, int iFrame /*= 0*/ )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromModule( lpsModule, dwImgType, iIdx, iFrame ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromModuleNamed
// Lädt ein Bild anhand des Ressourcen-Namens aus einem Modul
HANDLE CMImgMan::LoadImgFromModuleNamed( LPCTSTR lpsModule, DWORD dwImgType, LPCTSTR lpsResName, int iFrame /*= 0*/ )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromModuleNamed( lpsModule, dwImgType, lpsResName, iFrame ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromPicture
// Lädt ein Bild aus einem Picture
HANDLE CMImgMan::LoadImgFromPicture( HWND hWndPic )
{
	if ( !hWndPic ) return NULL;

	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromPicture( hWndPic ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromResource
// Lädt ein Bild aus einer Sal-Resource
HANDLE CMImgMan::LoadImgFromResource( HSTRING hsRes, DWORD dwImgType, int iFrame /* = 0 */ )
{
	if ( !( hsRes ) ) return NULL;

	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromResource( hsRes, dwImgType, iFrame ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromScreen
// Lädt ein Bild anhand des akt. Inhalts des Bildschirms bzw. eines Fensters  
HANDLE CMImgMan::LoadImgFromScreen( HWND hWnd, DWORD dwFlags )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromScreen( hWnd, dwFlags ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromString
// Lädt ein Bild aus einem String
HANDLE CMImgMan::LoadImgFromString( HSTRING hs, DWORD dwImgType, int iFrame /* = 0 */, BOOL bBase64 /*=FALSE*/ )
{
	if ( !( hs ) ) return NULL;

	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromString( hs, dwImgType, iFrame, bBase64 ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// LoadImgFromVisPic
// Lädt ein Bild aus einem Visual Toolchest Picture
HANDLE CMImgMan::LoadImgFromVisPic( LPVOID lpVisPic )
{
	CMImg *pImg = new CMImg;
	if ( !pImg->LoadFromVisPic( lpVisPic ) )
	{
		delete pImg;
		return NULL;
	}

	m_List.push_back( pImg );

	return HANDLE( pImg );
}

// MirrorImg
// Spiegelt das Bild vertikal oder horizontal
BOOL CMImgMan::MirrorImg( HANDLE hImg, BOOL bHorz )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Mirror( bHorz );
}

// PrintImg
// Druckt ein Bild
BOOL CMImgMan::PrintImg( HANDLE hImg, HDC hDC, LPCRECT pRect, COLORREF clrPaper, DWORD dwFlags )
{
	if ( !( hImg && hDC && pRect ) )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Print( hDC, pRect, clrPaper, dwFlags );
}

// PutImgImage
// Fügt ein Bild in ein Bild ein
BOOL CMImgMan::PutImgImage( HANDLE hImg, long lX, long lY, HANDLE hImgPut )
{
	if ( !( hImg && hImgPut ) )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	CMImg *pImgPut = (CMImg*)hImgPut;
	return pImg->PutImage( lX, lY, pImgPut );
}

// PutImgText
// Gibt Text in einem Bild aus
BOOL CMImgMan::PutImgText( HANDLE hImg, long lX, long lY, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->PutText( lX, lY, lpsText, lpsFontName, lFontSize, dwFontEnh, clr );
}

// PutImgTextAligned
// Gibt Text ausgerichtet in einem Bild aus
BOOL CMImgMan::PutImgTextAligned( HANDLE hImg, LPRECT pRect, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr, DWORD dwAlignFlags )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->PutTextAligned( pRect, lpsText, lpsFontName, lFontSize, dwFontEnh, clr, dwAlignFlags );
}

// QueryImgFlags
// Ermittelt, on Flags eines Bildes gesetzt sind
BOOL CMImgMan::QueryImgFlags( HANDLE hImg, DWORD dwFlags )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->QueryFlags( dwFlags );	
}

// ResampleImg
// Setzt eine neue Größe
BOOL CMImgMan::ResampleImg( HANDLE hImg, int iNewX, int iNewY, int iMode )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Resample( iNewX, iNewY, iMode );	
}

// RotateImg
// Dreht ein Bild
BOOL CMImgMan::RotateImg( HANDLE hImg, float fAngle )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Rotate( fAngle );
}

// SaveImg
// Speichert ein Bild in einer Datei
BOOL CMImgMan::SaveImg( HANDLE hImg, LPTSTR lpsFile, DWORD dwType )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->Save( lpsFile, dwType );	
}

// SetImgBpp
// Setzt die Farbtiefe eines Bildes
BOOL CMImgMan::SetImgBpp( HANDLE hImg, WORD wBpp )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetBpp( wBpp );
}

// SetImgColorOpacity
// Setzt die Durchlässigkeit einer Farbe im Bild
BOOL CMImgMan::SetImgColorOpacity( HANDLE hImg, DWORD dwTranspColor, BYTE byVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetColorOpacity( dwTranspColor, byVal );
}

// SetImgDrawClipRect
// Setzt das Zeichnungs-Clip-Rechteck
BOOL CMImgMan::SetImgDrawClipRect( HANDLE hImg, LPRECT pRect )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetDrawClipRect( pRect );
}

// SetImgDrawState
// Liefert den Zeichnungs-Status eines Bildes
BOOL CMImgMan::SetImgDrawState( HANDLE hImg, DWORD dwState )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetDrawState( dwState );
}

// SetImgFlags
// Setzt/Entfernt Bild-Flags
BOOL CMImgMan::SetImgFlags( HANDLE hImg, DWORD dwFlags, BOOL bSet )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetFlags( dwFlags, bSet );
}

// SetImgGIFCompr
// Setzt die GIF-Kompression
BOOL CMImgMan::SetImgGIFCompr( HANDLE hImg, int iVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetGIFCompr( iVal );	
}

// SetImgJPGQuality
// Setzt die Bildqualität
BOOL CMImgMan::SetImgJPGQuality( HANDLE hImg, BYTE byVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetJPGQuality( byVal );
}

// SetImgLightness
// Erhöht/verringert die Bildhelligkeit bzw. Kontrast
BOOL CMImgMan::SetImgLightness( HANDLE hImg, long lLevel, long lContrast )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetLightness( lLevel, lContrast );
}

// SetImgMaxOpacity
// Setzt die max. Durchlässigkeit des Bildes
BOOL CMImgMan::SetImgMaxOpacity( HANDLE hImg, BYTE byVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetMaxOpacity( byVal );
}

// SetImgOpacity
// Setzt die Durchlässigkeit des Bildes
BOOL CMImgMan::SetImgOpacity( HANDLE hImg, BYTE byFrom, BYTE byTo, BYTE byVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetOpacity( byFrom, byTo, byVal );
}

// SetImgPaletteColor
// Setzt die Farbe eines Paletteneintrags
BOOL CMImgMan::SetImgPaletteColor( HANDLE hImg, BYTE byIdx, DWORD dwColor )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetPaletteColor( byIdx, dwColor );

}

// SetImgPixelColor
// Setzt die Farbe eines Bild-Pixels
BOOL CMImgMan::SetImgPixelColor( HANDLE hImg, int iX, int iY, DWORD dwColor )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetPixelColor( iX, iY, dwColor );
}

// SetImgPixelOpacity
// Setzt die Durchlässigkeit eines Bild-Pixels
BOOL CMImgMan::SetImgPixelOpacity( HANDLE hImg, int iX, int iY, BYTE byVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetPixelOpacity( iX, iY, byVal );
}

// SetImgSize
// Setzt die Größe eines Bildes
BOOL CMImgMan::SetImgSize( HANDLE hImg, int iNewWid, int iNewHt, DWORD dwOverflowColor, BYTE byOverflowOpacity )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetSize( iNewWid, iNewHt, dwOverflowColor, byOverflowOpacity );
}

// SetImgTIFCompr
// Setzt die TIF-Kompression
BOOL CMImgMan::SetImgTIFCompr( HANDLE hImg, int iVal )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetTIFCompr( iVal );	
}

// SetImgTranspColor
// Setzt die transparente Farbe eines Bildes
BOOL CMImgMan::SetImgTranspColor( HANDLE hImg, DWORD dwColor )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SetTranspColor( dwColor );
}

// StripImgOpacity
// Überführt die Transparenz ( Transp. Farbe und Opazität ) in die Daten eines Bildes
BOOL CMImgMan::StripImgOpacity( HANDLE hImg, DWORD dwColor )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->StripOpacity( dwColor );
}

// SubstImgColor
// Ersetzt eine oder alle Farben durch eine andere Farbe
BOOL CMImgMan::SubstImgColor( HANDLE hImg, DWORD dwOldColor, DWORD dwNewColor )
{
	if ( !hImg )
		return FALSE;

	CMImg *pImg = (CMImg*)hImg;
	return pImg->SubstColor( dwOldColor, dwNewColor );
}