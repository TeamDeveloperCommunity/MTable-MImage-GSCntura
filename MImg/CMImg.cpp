// CMImg.cpp

// Includes
#include "CMImg.h"

// Klasse CMImg

// Statische Member
SIZE CMImg::m_siWMF = {0,0};

// Konstruktor
CMImg::CMImg( )
{
	m_pImage = NULL;

	m_byJPGQuality = 75;
	m_dwDrawState = 0;
	m_dwFlags = 0;
	m_iGIFCompr = 0;
	m_iTIFCompr = 0;
	SetRectEmpty( &m_rDrawClipRect );
}

// Destruktor
CMImg::~CMImg( )
{
	FreeImageData( );
}

// AddShadow
BOOL CMImg::AddShadow( int iDispX, int iDispY, BYTE bySmoothRad, DWORD dwClr, DWORD dwClrBkgnd, BYTE byIntens )
{
	if ( !m_pImage )
		return FALSE;

	iDispX = -iDispX;

	RGBQUAD c0 = CxImage::RGBtoRGBQUAD(dwClr);
	RGBQUAD c1 = CxImage::RGBtoRGBQUAD(dwClrBkgnd);

	CxImage *mix = new CxImage(*m_pImage);
	mix->IncreaseBpp(24);
	mix->SelectionClear();
	mix->SelectionAddColor(c1);
	CxImage iShadow;
	mix->SelectionSplit(&iShadow);
	mix->SelectionDelete();

	iShadow.GaussianBlur( bySmoothRad );

	for (int n = 0; n<256; n++){
		BYTE r = (BYTE)(c1.rgbRed   + ((byIntens*n*((long)c0.rgbRed   - (long)c1.rgbRed))>>16));
		BYTE g = (BYTE)(c1.rgbGreen + ((byIntens*n*((long)c0.rgbGreen - (long)c1.rgbGreen))>>16));
		BYTE b = (BYTE)(c1.rgbBlue  + ((byIntens*n*((long)c0.rgbBlue  - (long)c1.rgbBlue))>>16));
		iShadow.SetPaletteColor((BYTE)(255-n),r,g,b);
	}

	mix->SetTransColor(c1);
	mix->SetTransIndex(0);
	mix->Mix(iShadow,CxImage::OpSrcCopy,iDispX,iDispY);
	mix->SetTransIndex(-1);

	DWORD dwTransClr = GetTranspColor( );
	FreeImageData( );
	m_pImage = mix;
	SetTranspColor( dwTransClr );

	return TRUE;
}

// AlphaPaletteToAlpha
BOOL CMImg::AlphaPaletteToAlpha( )
{
	if ( !m_pImage )
		return FALSE;

	if ( !m_pImage->AlphaPaletteIsValid( ) )
		return TRUE;

	if ( !CreateAlphaFromPalette( ) )
		return FALSE;

	return TRUE;
}

// Assign
BOOL CMImg::AssignTo( CMImg *pImg )
{
	if ( !pImg )
		return FALSE;
	if ( !m_pImage )
		return FALSE;

	pImg->FreeImageData( );

	pImg->m_pImage = new CxImage( *m_pImage, TRUE, TRUE, TRUE );

	BOOL bError = FALSE;
	if ( !CopyMembersTo( pImg ) )
		bError = TRUE;

	return !bError;
}

// CopyMembersTo
BOOL CMImg::CopyMembersTo( CMImg *pImg )
{
	if ( !pImg )
		return FALSE;

	pImg->m_byJPGQuality = m_byJPGQuality;
	pImg->m_dwDrawState = m_dwDrawState;
	pImg->m_dwFlags = m_dwFlags;
	pImg->m_iGIFCompr = m_iGIFCompr;
	pImg->m_iTIFCompr = m_iTIFCompr;
	CopyRect( &pImg->m_rDrawClipRect, &m_rDrawClipRect );

	return TRUE;
}

// CopyToClipboard
BOOL CMImg::CopyToClipboard( )
{
	if ( !m_pImage )
		return FALSE;

	BOOL bOk = FALSE;
	if ( OpenClipboard( NULL ) )
	{
		HANDLE hGlobal = m_pImage->CopyToHandle( );
		if ( hGlobal )
		{
			if ( SetClipboardData( CF_DIB, hGlobal ) )
				bOk = TRUE;
			else
				GlobalFree( hGlobal );
		}

		CloseClipboard( );
	}

	return bOk;
}

// CreateAlpha
BOOL CMImg::CreateAlpha( )
{
	if ( !m_pImage )
		return FALSE;

	if ( !m_pImage->AlphaIsValid( ) )
	{
		m_pImage->AlphaCreate( );
		m_pImage->AlphaSet( 255 );
	}

	return TRUE;
}

// CreateAlphaFromPalette
BOOL CMImg::CreateAlphaFromPalette( )
{
	if ( !m_pImage )
		return FALSE;

	// Palette da?
	if ( !HasPalette( ) )
		return FALSE;

	// Alpha erzeugen
	if ( !CreateAlpha( ) )
		return FALSE;

	// Pixel durchlaufen
	BYTE byIdx;
	DWORD dwHt = m_pImage->GetHeight( );
	DWORD dwWid = m_pImage->GetWidth( );
	RGBQUAD r;
	for ( DWORD dwY = 0; dwY < dwHt; dwY++ )
	{
		for ( DWORD dwX = 0; dwX < dwWid; dwX++ )
		{
			byIdx = m_pImage->GetPixelIndex( dwX, dwY );
			r = m_pImage->GetPaletteColor( byIdx );
			if ( r.rgbReserved != 0 )
				m_pImage->AlphaSet( dwX, dwY, r.rgbReserved );
		}
	}

	return TRUE;
}

// CreateCopy
CMImg * CMImg::CreateCopy( )
{
	if ( !m_pImage )
		return NULL;

	CMImg *pImg = new CMImg;
	AssignTo( pImg );

	return pImg;
}

// CreateCopyEx
CMImg * CMImg::CreateCopyEx( LPRECT pRect, DWORD dwType, DWORD dwFlags )
{
	if ( !( m_pImage && pRect ) )
		return NULL;

	if ( dwType == MIMG_TYPE_UNKNOWN )
		dwType = GetType( );

	int iX = max( 0, pRect->left );
	int iY = max( 0, pRect->top );
	int iWid = int( m_pImage->GetWidth( ) );
	int iHt = int( m_pImage->GetHeight( ) );

	int jWid = min( iWid - iX, pRect->right - pRect->left );
	if ( jWid <= 0 )
		return NULL;
	int jHt = min( iHt - iY, pRect->bottom - pRect->top );
	if ( jHt <= 0 )
		return NULL;

	CMImg *pImg = new CMImg;
	pImg->m_pImage = new CxImage( *m_pImage, FALSE, FALSE, FALSE );
	if ( !pImg->m_pImage->Create( jWid, jHt, GetBpp( ), dwType ) )
	{
		delete pImg;
		return NULL;
	};

	RGBQUAD *pPal = m_pImage->GetPalette( );
	if ( pPal )
		pImg->m_pImage->SetPalette( pPal, m_pImage->GetNumColors( ) );

	BOOL bCopyAlpha = m_pImage->AlphaIsValid( ) && ( dwFlags & MIMG_COPY_OPACITY );
	if ( bCopyAlpha )
		pImg->m_pImage->AlphaCreate( );
	else
		pImg->m_pImage->AlphaSetMax( 255 );

	int jX, jY, kX, kY;
	for ( jY = 0, kY = iHt - jHt - iY; jY < jHt; jY++, kY++ )
	{
		for ( jX = 0, kX = iX; jX < jWid; jX++, kX++ )
		{
			pImg->m_pImage->SetPixelColor( jX, jY, m_pImage->GetPixelColor( kX, kY ) );
			if ( bCopyAlpha )
				pImg->m_pImage->AlphaSet( jX, jY, m_pImage->AlphaGet( kX, kY ) );
		}
	}

	CopyMembersTo( pImg );

	return pImg;
}

// CreateDisabledImgCopy
CxImage* CMImg::CreateDisabledImgCopy( CxImage *pImgSrc )
{
	CxImage *pImg = ( pImgSrc ? pImgSrc : m_pImage );
	if ( !pImg && pImg->IsValid( ) )
		return NULL;

	CxImage *pImgCopy = NULL;

	int iWid = pImg->GetWidth( );
	int iHt = pImg->GetHeight( );

	struct MONOBMI
	{
		BITMAPINFOHEADER bmiHeader;
		RGBQUAD bmiColors[2];
	};

	MONOBMI mbmi = 
	{
		{
			sizeof( BITMAPINFOHEADER ),	// biSize
			iWid,						// biWidth
			iHt,						// biHeight
			1,							// biPlanes
			1,							// biBitCount
			BI_RGB,						// biCompression
			0,							// biSizeImage
			0,							// biXPelsPerMeter
			0,							// biYPelsPerMeter
			0,							// biClrUsed
			0							// biClrImportant
		},
		{
			{ 0x00, 0x00, 0x00, 0x00 },	// Black
			{ 0xFF, 0xFF, 0xFF, 0x00 }	// White
		}
	};

	BOOL bOk = FALSE;
	HDC hDC = GetDC( NULL );
	if ( hDC )
	{
		HDC hClrDC = CreateCompatibleDC( hDC );
		if ( hClrDC )
		{
			void *pBits = NULL;
			HBITMAP hBitm = CreateDIBSection( hClrDC, (LPBITMAPINFO)pImg->GetDIB( ), DIB_RGB_COLORS, &pBits, NULL, 0);
			if ( hBitm )
			{
				HGDIOBJ hOldBitm = SelectObject( hClrDC, hBitm );
				pImg->Draw2( hClrDC );

				HDC hMonoDC = CreateCompatibleDC( hDC );
				if ( hMonoDC )
				{
					void *pMonoBits = NULL;
					HBITMAP hMonoBitm = CreateDIBSection( hMonoDC, (LPBITMAPINFO)&mbmi, DIB_RGB_COLORS, &pMonoBits, NULL, 0);
					if ( hMonoBitm )
					{
						HGDIOBJ hOldMonoBitm = SelectObject( hMonoDC, hMonoBitm );

						BitBlt( hMonoDC, 0, 0, iWid, iHt, hClrDC, 0, 0, SRCCOPY );

						RECT r = { 0, 0, iWid + 1, iHt + 1 };
						FillRect( hClrDC, &r, GetSysColorBrush( COLOR_3DFACE ) );

						HBRUSH hBrush = CreateSolidBrush( GetSysColor( COLOR_3DHILIGHT ) );
						HGDIOBJ hOldBrush = (HBRUSH)SelectObject( hClrDC, hBrush );
						BitBlt( hClrDC, 1, 1, iWid, iHt, hMonoDC, 0, 0, 0xB8074A );
						SelectObject( hClrDC, hOldBrush );
						DeleteObject( hBrush );

						hBrush = CreateSolidBrush( GetSysColor( COLOR_3DSHADOW ) );
						hOldBrush = SelectObject( hClrDC, hBrush );
						BitBlt( hClrDC, 0, 0, iWid, iHt, hMonoDC, 0, 0, 0xB8074A );
						SelectObject( hClrDC, hOldBrush );
						DeleteObject( hBrush );

						SelectObject( hMonoDC, hOldMonoBitm );
						DeleteObject( hMonoBitm );

						pImgCopy = new CxImage( pImg->GetType( ) );
						pImgCopy->CreateFromHBITMAP( hBitm );
						pImgCopy->AlphaCopy( *pImg );
						pImgCopy->SelectionCopy( *pImg );

						if ( pImg->GetTransIndex( ) >= 0 )
						{
							// Farbtransparenz in Alpha-Transparenz konvertieren
							if ( !pImgCopy->AlphaIsValid( ) )
							{
								pImgCopy->AlphaCreate( );
								pImgCopy->AlphaSet( 255 );
							}

							DWORD dwTransColor = pImg->RGBQUADtoRGB( pImg->GetTransColor( ) );

							DWORD dwX,dwY;
							DWORD dwHt = pImg->GetHeight( );
							DWORD dwWid = pImg->GetWidth( );

							for ( dwY = 0; dwY < dwHt; dwY++ )
							{
								for ( dwX = 0; dwX < dwWid; dwX++ )
								{
									if ( pImg->RGBQUADtoRGB( pImg->GetPixelColor( dwX, dwY ) ) == dwTransColor )
										pImgCopy->AlphaSet( dwX, dwY, 0 );
								}
							}
						}

						bOk = TRUE;
					}

					DeleteDC( hMonoDC );
				}
				SelectObject( hClrDC, hOldBitm );
				DeleteObject( hBitm );
			}

			DeleteDC( hClrDC );
		}
		
		ReleaseDC( NULL, hDC );
	}

	return pImgCopy;
}

// CreateFromHBITMAP
BOOL CMImg::CreateFromHBITMAP( HBITMAP hBmColor, HBITMAP hBmMask /*=NULL*/ )
{
	if ( !( m_pImage && hBmColor ) )
		return FALSE;
	
	// Aus der Farb-Bitmap erzeugen
	BITMAP bmp;
    if ( !GetObject( hBmColor, sizeof(BITMAP), (LPVOID)&bmp ) )
		return FALSE;

    if ( !m_pImage->Create( bmp.bmWidth, bmp.bmHeight, bmp.bmBitsPixel, m_pImage->GetType( ) ) )
		return FALSE;

    HDC hDC = GetDC( NULL );
	if ( !hDC )
		return FALSE;
    int iRet = GetDIBits( hDC, hBmColor, 0, m_pImage->GetHeight( ), m_pImage->GetBits( ),(LPBITMAPINFO)m_pImage->GetDIB( ), DIB_RGB_COLORS );

	// Alpha initialisieren
	BOOL bNeedAlpha = FALSE;

	if ( !m_pImage->AlphaIsValid( ) )
		m_pImage->AlphaCreate( );
	m_pImage->AlphaSet( 255 );

	// Alpha-Transparenz anhand der letzten 8 Bit setzen, wenn es eine 32Bit-Bitmap ist
	if (bmp.bmBitsPixel == 32)
	{
		BOOL bAlphaOk = FALSE;

		BITMAPINFOHEADER bih =
		{
			sizeof(BITMAPINFOHEADER),	// biSize
			bmp.bmWidth,				// biWidth
			bmp.bmHeight,				// biHeight
			1,							// biPlanes
			32,							// biBitCount
			BI_RGB,						// biCompression
			0,							// biSizeImage
			0,							// biXPelsPerMeter
			0,							// biYPelsPerMeter
			0,							// biClrUsed
			0							// biClrImportant
		};

		RGBQUAD *pBits = new RGBQUAD[bmp.bmWidth * bmp.bmHeight];
		if (pBits)
		{
			iRet = GetDIBits(hDC, hBmColor, 0, bmp.bmHeight, pBits, (LPBITMAPINFO)&bih, DIB_RGB_COLORS);
			if (iRet)
			{
				RGBQUAD r;
				for (int iY = 0; iY < bmp.bmHeight; ++iY)
				{
					for (int iX = 0; iX < bmp.bmWidth; ++iX)
					{
						r = pBits[(iY * bmp.bmWidth) + iX];
						/*if ( r.rgbReserved != 0 )
						{
						m_pImage->AlphaSet( iX, iY, r.rgbReserved );
						bNeedAlpha = TRUE;
						}*/
						m_pImage->AlphaSet(iX, iY, r.rgbReserved);

						if (r.rgbReserved != 0)
							bAlphaOk = TRUE;
					}
				}
			}
			delete[]pBits;
		}

		if (bAlphaOk)
			bNeedAlpha = TRUE;
		else
		{
			// Alle Alpha-Werte sind 0, d.h. die Bitmap hat doch keine Alpha-Daten, also alles auf 255 setzen
			m_pImage->AlphaInvert();
		}
	}

	// Alpha-Transparenz anhand der Maske setzen
	if ( hBmMask )
	{
		CxImage ImgMask( MIMG_TYPE_BMP );
		ImgMask.CreateFromHBITMAP( hBmMask );

		COLORREF clr;
		DWORD dwHt = m_pImage->GetHeight( );
		DWORD dwWid = m_pImage->GetWidth( );
		RGBQUAD r;
		for ( DWORD dwY = 0; dwY < dwHt; dwY++ )
		{
			for ( DWORD dwX = 0; dwX < dwWid; dwX++ )
			{
				r = ImgMask.GetPixelColor( dwX, dwY );
				clr = ImgMask.RGBQUADtoRGB( r );
				if ( clr )
				{
					m_pImage->AlphaSet( dwX, dwY, 0 );
					bNeedAlpha = TRUE;
				}
			}
		}
	}
	
	ReleaseDC( NULL, hDC );            

	if ( !bNeedAlpha )
		m_pImage->AlphaDelete( );

	return TRUE;
}

// CreateFromHICON
BOOL CMImg::CreateFromHICON( HICON hIcon )
{
	if ( !( m_pImage && hIcon ) )
		return FALSE;
	
	ICONINFO ii;
	if ( !GetIconInfo( hIcon, &ii ) )
		return FALSE;
	if ( !( ii.hbmColor && ii.hbmMask ) )
		return FALSE;

	BOOL bOk = CreateFromHBITMAP( ii.hbmColor, ii.hbmMask );

	DeleteObject( ii.hbmColor );
	DeleteObject( ii.hbmMask );

	return bOk;
}

// CreateHBITMAP
HBITMAP CMImg::CreateHBITMAP( HDC hDC )
{
	if ( !m_pImage )
		return NULL;

	return m_pImage->MakeBitmap( hDC, TRUE );
}

// CreateHICON
HICON CMImg::CreateHICON( )
{
	if ( !m_pImage )
		return NULL;

	/*
	// Temporäre Datei erzeugen
	TCHAR szPath[1024];
	DWORD dwLen = GetTempPath( 1024, szPath );
	if ( !dwLen )
		return NULL;

	TCHAR szFile[1024];
	if ( !GetTempFileName( szPath, _T("MIM"), 0, szFile ) )
		return NULL;

	if ( !Save( szFile, MIMG_TYPE_ICO ) )
		return NULL;

	HICON hIcon = (HICON)LoadImage( 0, szFile, IMAGE_ICON, m_pImage->GetWidth( ), m_pImage->GetHeight( ), LR_LOADFROMFILE );

	DeleteFile( szFile );

	return hIcon;
	*/

	return m_pImage->MakeIcon(NULL, TRUE);
}

// CreateVisPic
LPVOID CMImg::CreateVisPic( DWORD dwType )
{
	if ( !( m_pImage && Vis.IsPresent( ) ) )
		return NULL;
	
	CMImg *pImg = NULL;
	HSTRING hs = 0;
	LPTSTR pBuf = NULL;
	LPVOID pVisPic = NULL;
	WORD wFlags = PIC_LoadSWinStr;

	// Typ festelegen
	if ( dwType == MIMG_TYPE_UNKNOWN )
	{
		DWORD dwMyType = GetType( );
		if ( dwMyType == MIMG_TYPE_ICO || dwMyType == MIMG_TYPE_BMP )
			dwType = dwMyType;
		else
			dwType = MIMG_TYPE_BMP;
	}

	// Ggf. jegliche Nicht-Alpha-Transparenz in Aplha umwandeln
	if ( dwType == MIMG_TYPE_BMP && GetBpp( ) == 24 && ( GetTranspColor( ) != MIMG_COLOR_UNDEF || m_pImage->AlphaPaletteIsValid( ) ) )
	{
		pImg = CreateCopy( );
		if ( !pImg)
			goto CLEANUP;

		if ( !pImg->CreateAlpha( ) )
			goto CLEANUP;

		if ( !pImg->AlphaPaletteToAlpha( ) )
			goto CLEANUP;

		if ( !pImg->TranspColorToAlpha( ) )
			goto CLEANUP;		
	}
	else
		pImg = this;

	// Wenn Alpha-Channel ohne jegliche Transparenz, Alpha-Channel entfernen
	if ( pImg->IsAlphaFullyOpaque( ) )
		pImg->m_pImage->AlphaDelete( );

	// String-Buffer ermitteln
	if ( !pImg->GetString( &hs, dwType ) )
		goto CLEANUP;

	long lLen;
	pBuf = Ctd.StringGetBuffer( hs, &lLen );
	if (pBuf)
	{
		// Flags setzen
		if (dwType == MIMG_TYPE_ICO)
		{
			wFlags |= PIC_FormatIcon;
			if (m_pImage->GetWidth() <= 16)
				wFlags |= PIC_LoadSmallIcon;
			else
				wFlags |= PIC_LoadLargeIcon;
		}
		else
			wFlags |= PIC_FormatBitmap;

		// Erzeugen
		pVisPic = Vis.PicLoad(wFlags, pBuf, NULL);
		Ctd.HStrUnlock(hs);
	}	

CLEANUP:
	if ( hs )
		Ctd.HStringUnRef( hs );

	if ( pImg && pImg != this )
		delete pImg;

	return pVisPic;
}


// DetermineType
int CMImg::DetermineType( BYTE *pBuffer, DWORD dwSize )
{
	if ( !pBuffer )
		return MIMG_TYPE_UNKNOWN;

	// BMP?
	if ( pBuffer[0] == 0x42 && pBuffer[1] == 0x4D ) // 'BM'
		return MIMG_TYPE_BMP;

	// ICO?
	if ( pBuffer[0] == 0x00 && pBuffer[1] == 0x00 && pBuffer[2] == 0x01 && pBuffer[3] == 0x00 )
		return MIMG_TYPE_ICO;

	// GIF?
	if ( pBuffer[0] == 0x47 && pBuffer[1] == 0x49 && pBuffer[2] == 0x46 && pBuffer[3] == 0x38 ) // 'GIF8'
		return MIMG_TYPE_GIF;

	// JPEG?
	if ( pBuffer[0] == 0xFF && pBuffer[1] == 0xD8 && pBuffer[2] == 0xFF && pBuffer[3] == 0xE0 )
		return MIMG_TYPE_JPG;

	// PNG?
	if ( pBuffer[0] == 0x89 && pBuffer[1] == 0x50 && pBuffer[2] == 0x4E && pBuffer[3] == 0x47 &&
		 pBuffer[4] == 0x0D && pBuffer[5] == 0x0A && pBuffer[6] == 0x1A && pBuffer[7] == 0x0A
	   )
		 return MIMG_TYPE_PNG;

	// TIF?
	if ( ( pBuffer[0] == 0x49 && pBuffer[1] == 0x49 ) || // 'II'
		 ( pBuffer[0] == 0x4D && pBuffer[1] == 0x4D )    // 'MM'
	   )
		 return MIMG_TYPE_TIF;

	// WMF?
	if ( pBuffer[0] == 0xD7 && pBuffer[1] == 0xCD && pBuffer[2] == 0xC6 && pBuffer[3] == 0x9A )
		return MIMG_TYPE_WMF;

	// PCX?
	if ( pBuffer[0] == 0x0A )
		return MIMG_TYPE_PCX;

	// WBMP?
	if ( pBuffer[0] == 0x00 && pBuffer[1] == 0x00 && pBuffer[2] != 0x00 && pBuffer[3] != 0x00 )
		return MIMG_TYPE_WBMP;

	// Hmmm... dann laden wir mal
	CxImage Img( pBuffer, dwSize, MIMG_TYPE_UNKNOWN );
	return Img.GetType( );
}

// Disable
BOOL CMImg::Disable( )
{
	if ( !m_pImage )
		return FALSE;

	CxImage *pImg = CreateDisabledImgCopy( );
	if ( !pImg )
		return FALSE;

	FreeImageData( );
	m_pImage = pImg;

	return TRUE;
}

// Dither
BOOL CMImg::Dither( long lMethod )
{
	if ( lMethod < 0 || lMethod > 7 )
		lMethod = 0;

	DWORD dwTranspColor = GetTranspColor( );
	BOOL bOk = m_pImage->Dither( lMethod );
	SetTranspColor( dwTranspColor );

	return bOk;
}

// Draw
BOOL CMImg::Draw( HDC hDC, LPCRECT pRect, DWORD dwFlags )
{
    if ( !( m_pImage && hDC && pRect ) )
		return FALSE;

	BOOL bOk = FALSE;


	CMImg *pMImg = this;
	BOOL bDisabled = ( m_dwDrawState & MIMG_DST_DISABLED ) || ( dwFlags & MIMG_DRAW_DISABLED );
	BOOL bInverted = ( m_dwDrawState & MIMG_DST_INVERTED ) || ( dwFlags & MIMG_DRAW_INVERTED );
	if ( bDisabled || bInverted )
	{
		pMImg = CreateCopy( );
		if ( !pMImg )
			return FALSE;

		if ( bDisabled )
			pMImg->Disable( );

		if ( bInverted )
			pMImg->Invert( );
	}

	LPRECT pClipRect = NULL;
	if ( dwFlags & MIMG_DRAW_USECLIPRECT )
		pClipRect = &m_rDrawClipRect;

	HPALETTE hPal = NULL;
	if ( GetDeviceCaps( hDC, RASTERCAPS ) & RC_PALETTE )
	{
		HPALETTE hPalStock = HPALETTE( GetStockObject( DEFAULT_PALETTE ) );
		hPal = SelectPalette( hDC, hPalStock, TRUE );
		SelectPalette( hDC, hPal, TRUE );

		if ( hPal == hPalStock )
			hPal = theMImgDefPal( );
	}

	bOk = pMImg->m_pImage->Draw( hDC, *pRect, pClipRect, hPal );

	if ( pMImg != this )
		delete pMImg;

	return bOk;
}

// DrawAligned
BOOL CMImg::DrawAligned( HDC hDC, LPCRECT pRect, int iWid /*= -1*/, int iHt /*= -1*/, DWORD dwAlignFlags /*= 0*/, DWORD dwFlags /*= 0*/ )
{
    if ( !( m_pImage && hDC && pRect ) )
		return FALSE;

	if ( iWid < 0 )
		iWid = m_pImage->GetWidth( );
	if ( iHt < 0 )
		iHt = m_pImage->GetHeight( );

	int iAreaWid = ( pRect->right - pRect->left );
	int iAreaHt = ( pRect->bottom - pRect->top );

	BOOL bWider = ( iWid > iAreaWid );
	BOOL bHigher = ( iHt > iAreaHt );

	RECT rImg;
	if ( ( dwAlignFlags & MIMG_ALIGN_HCENTER ) && !bWider )
		rImg.left = pRect->left + ( ( iAreaWid - iWid ) / 2 );
	else if ( ( dwAlignFlags & MIMG_ALIGN_RIGHT ) && !bWider )
		rImg.left = pRect->right - iWid;
	else
		rImg.left = pRect->left;

	if ( ( dwAlignFlags & MIMG_ALIGN_VCENTER ) && !bHigher )
		rImg.top = pRect->top + ( ( iAreaHt - iHt ) / 2 );
	else if ( ( dwAlignFlags & MIMG_ALIGN_BOTTOM ) && !bHigher )
		rImg.top = pRect->bottom - iHt;
	else
		rImg.top = pRect->top;

	rImg.right = rImg.left + iWid;
	rImg.bottom = rImg.top + iHt;

	BOOL bRestoreClipRect = FALSE;
	RECT rOldClipRect;
	if ( !( dwFlags & MIMG_DRAW_USECLIPRECT ) )
	{
		bRestoreClipRect = TRUE;
		GetDrawClipRect( &rOldClipRect );
		SetDrawClipRect( (LPRECT)pRect );
	}

	BOOL bOk = Draw( hDC, (LPCRECT)&rImg, dwFlags );

	if ( bRestoreClipRect )
		SetDrawClipRect( &rOldClipRect );

	return bOk;
}

// DrawPattern
BOOL CMImg::DrawPattern( HDC hDC, LPCRECT pRect, int iWid /*= -1*/, int iHt /*= -1*/, DWORD dwFlags /*= 0*/ )
{
    if ( !( m_pImage && hDC && pRect ) )
		return FALSE;

	if ( iWid < 0 )
		iWid = m_pImage->GetWidth( );
	if ( iHt < 0 )
		iHt = m_pImage->GetHeight( );

	BOOL bRestoreClipRect = FALSE;
	LPRECT pClipRect = (LPRECT)pRect;
	RECT rImg = { pRect->left, pRect->top, pRect->left + iWid, pRect->top + iHt }, rIS, rOldClipRect;

	if ( !( dwFlags & MIMG_DRAW_USECLIPRECT ) )
	{
		bRestoreClipRect = TRUE;
		GetDrawClipRect( &rOldClipRect );
		SetDrawClipRect( pClipRect );
	}

	while ( rImg.top <= pRect->bottom )
	{
		while ( rImg.left <= pRect->right )
		{
			if ( IntersectRect( &rIS, &rImg, pClipRect ) )
			{
				Draw( hDC, &rImg, dwFlags );
			}
			rImg.left = rImg.right;
			rImg.right = rImg.left + iWid;
		}
		rImg.left = pRect->left;
		rImg.top = rImg.top + iHt;
		rImg.right = pRect->left + iWid;
		rImg.bottom = rImg.top + iHt;
	}

	if ( bRestoreClipRect )
		SetDrawClipRect( &rOldClipRect );

	return TRUE;
}

// Save
BOOL CMImg::Encode( CxFile *pFile, DWORD dwType )
{
	if ( !( m_pImage && pFile ) )
		return FALSE;

	if ( dwType == MIMG_TYPE_UNKNOWN )
		dwType = m_pImage->GetType( );

	CMImg *pMImg = this;

	BOOL bGIF = ( dwType == MIMG_TYPE_GIF );
	BOOL bICO = ( dwType == MIMG_TYPE_ICO );
	BOOL bJPG = ( dwType == MIMG_TYPE_JPG );
	BOOL bTGA = ( dwType == MIMG_TYPE_TGA );
	BOOL bTIF = ( dwType == MIMG_TYPE_TIF );
	BOOL bWBMP = ( dwType == MIMG_TYPE_WBMP );

	if ( bJPG && GetBpp( ) < 24 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 24 ) )
			return FALSE;
	}
	else if ( bTGA && GetBpp( ) < 8 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 8 ) )
			return FALSE;
	}
	else if ( bWBMP && GetBpp( ) > 1 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 1 ) )
			return FALSE;
	}
	else if ( bICO )
	{
		// Farbe der "transparenten Pixel" auf Schwarz setzen 
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		CxImage *pImage = pMImg->m_pImage;

		BOOL bAlpha = pImage->AlphaIsValid( );
		BOOL bAlphaPalette = pImage->AlphaPaletteIsValid( );
		DWORD dwTranspColor = pMImg->GetTranspColor( );

		if ( bAlpha || bAlphaPalette || dwTranspColor != MIMG_COLOR_UNDEF )
		{
			if ( !bAlpha )
				pImage->AlphaCreate( );

			pMImg->SetTranspColor( MIMG_COLOR_UNDEF );

			if ( pMImg->HasPalette( ) && pMImg->GetColorIndex( 0 ) < 0 )
				pMImg->SetBpp( 24 );

			DWORD dwColor;
			DWORD dwHt = pImage->GetHeight( );
			DWORD dwWid = pImage->GetWidth( );
			DWORD dwX, dwY;

			for ( dwY = 0; dwY < dwHt; ++dwY )
			{
				for ( dwX = 0; dwX < dwWid; ++dwX )
				{
					dwColor = pImage->RGBQUADtoRGB( pImage->GetPixelColor( dwX, dwY, FALSE ) );
					if (
						( dwColor == dwTranspColor ) ||
						( bAlpha && pImage->AlphaGet( dwX, dwY ) == 0 ) ||
						( bAlphaPalette && pImage->GetPixelColor( dwX, dwY, TRUE ).rgbReserved == 0 )
					   )
					{
						pImage->SetPixelColor( dwX, dwY, 0 );
						pImage->AlphaSet( dwX, dwY, 0 );
					}
				}
			}
		}
	}

	BYTE byJPGQuality;
	int iEncodeOption;
	if ( bJPG )
	{
		byJPGQuality = pMImg->m_pImage->GetJpegQuality( );
		pMImg->m_pImage->SetJpegQuality( pMImg->GetJPGQuality( ) );
	}
	else if ( bGIF || bTIF )
	{
		iEncodeOption = pMImg->m_pImage->GetCodecOption( );
		if ( bGIF )
			pMImg->m_pImage->SetCodecOption( GetGIFCompr( ) );
		else
		{
			WORD wBpp = pMImg->GetBpp( );
			int iCompr = pMImg->GetTIFCompr( );
			if ( ( iCompr == MIMG_COMPR_TIF_CCITTRLE || iCompr == MIMG_COMPR_TIF_CCITTFAX3 || iCompr == MIMG_COMPR_TIF_CCITTFAX4 ) && wBpp != 1 )
				return FALSE;
			else if ( ( iCompr == MIMG_COMPR_TIF_NONE || iCompr == MIMG_COMPR_TIF_LZW || iCompr == MIMG_COMPR_TIF_PACKBITS ) && wBpp == 24 )
				return FALSE;
			else if ( iCompr == MIMG_COMPR_TIF_JPEG && wBpp != 24 )
				return FALSE;

			pMImg->m_pImage->SetCodecOption( iCompr );
		}
	}

	BOOL bOk = pMImg->m_pImage->Encode( pFile, dwType );

	if ( pMImg != this )
	{
		if ( bJPG )
		{
			pMImg->m_pImage->SetJpegQuality( byJPGQuality );
		}
		else if ( bGIF || bTIF )
		{
			pMImg->m_pImage->SetCodecOption( iEncodeOption );
		}

		delete pMImg;
	}

	return bOk;
}

// EnumResNameProc
BOOL CALLBACK CMImg::EnumResNameProc( HANDLE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG lParam )
{
	ENUMRESNAME *pern = (ENUMRESNAME*)lParam;
	if ( !pern )
		return FALSE;

	if ( pern->iIdxCurrent == pern->iIdxDesired && IS_INTRESOURCE( lpszName ) )
	{
		pern->lpName = lpszName;
		return FALSE;
	}

	++pern->iIdxCurrent;
	return ( pern->iIdxCurrent <= pern->iIdxDesired );
}

// FreeImageData
void CMImg::FreeImageData( )
{
	if ( m_pImage )
	{
		delete m_pImage;
		m_pImage = NULL;
	}
}

// GetBpp
WORD CMImg::GetBpp( )
{
	WORD wBpp = 0;
	if ( m_pImage )
		wBpp = m_pImage->GetBpp( );
	return wBpp;
}

// GetColorIndex
int CMImg::GetColorIndex( DWORD dwColor )
{
	if ( !m_pImage )
		return -1;

	if ( dwColor == MIMG_COLOR_UNDEF )
		return -1;

	if ( !HasPalette( ) )
		return -1;

	int iColors = int( m_pImage->GetNumColors( ) );
	RGBQUAD r;
	for ( int i = 0; i < iColors; ++i )
	{
		r = m_pImage->GetPaletteColor( i );
		if ( m_pImage->RGBQUADtoRGB( r ) == dwColor )
			return i;
	}

	return -1;
}

// GetDIBHandle
HANDLE CMImg::GetDIBHandle( )
{
	if ( !m_pImage )
		return NULL;

	return m_pImage->GetDIB( );
}

// GetDIBString
BOOL CMImg::GetDIBString( LPHSTRING phs )
{
	if ( !( m_pImage && phs ) )
		return FALSE;

	void *pDib = m_pImage->GetDIB( );
	if ( !pDib )
		return FALSE;

	long lSize = m_pImage->GetSize( );
	if ( lSize <= 0 )
		return FALSE;

	if ( !Ctd.InitLPHSTRINGParam( phs, lSize ) )
		return FALSE;

	long lLen;
	void *pBuf = Ctd.StringGetBuffer( *phs, &lLen );
	if (pBuf)
	{
		memcpy(pBuf, pDib, lSize);
		Ctd.HStrUnlock(*phs);
	}
	else
		return FALSE;	

	return TRUE;
}

// GetDrawClipRect
BOOL CMImg::GetDrawClipRect( LPRECT pRect )
{
	if ( !pRect )
		return FALSE;

	return CopyRect( pRect, &m_rDrawClipRect );
}

// GetDrawState
DWORD CMImg::GetDrawState( )
{
	return m_dwDrawState;
}

// GetFrameBpp
BYTE CMImg::GetFrameBpp( int iFrame )
{
	BYTE byBpp = 0;
	HDC hDC;

	switch ( iFrame )
	{
	case MIMG_FRAME_ICON_XXXXLARGE_4BPP:
	case MIMG_FRAME_ICON_XXXLARGE_4BPP:
	case MIMG_FRAME_ICON_XXLARGE_4BPP:
	case MIMG_FRAME_ICON_XLARGE_4BPP:
	case MIMG_FRAME_ICON_LARGE_4BPP:
	case MIMG_FRAME_ICON_MEDIUM_4BPP:
	case MIMG_FRAME_ICON_SMALL_4BPP:
		byBpp = 4;
		break;
	
	case MIMG_FRAME_ICON_XXXXLARGE_8BPP:
	case MIMG_FRAME_ICON_XXXLARGE_8BPP:
	case MIMG_FRAME_ICON_XXLARGE_8BPP:
	case MIMG_FRAME_ICON_XLARGE_8BPP:
	case MIMG_FRAME_ICON_LARGE_8BPP:
	case MIMG_FRAME_ICON_MEDIUM_8BPP:
	case MIMG_FRAME_ICON_SMALL_8BPP:
		byBpp = 8;
		break;

	case MIMG_FRAME_ICON_XXXXLARGE_24BPP:
	case MIMG_FRAME_ICON_XXXLARGE_24BPP:
	case MIMG_FRAME_ICON_XXLARGE_24BPP:
	case MIMG_FRAME_ICON_XLARGE_24BPP:
	case MIMG_FRAME_ICON_LARGE_24BPP:
	case MIMG_FRAME_ICON_MEDIUM_24BPP:
	case MIMG_FRAME_ICON_SMALL_24BPP:
		byBpp = 24;
		break;

	default:
		hDC = GetDC( NULL );
		if ( hDC )
		{
			byBpp = (BYTE)GetDeviceCaps( hDC, BITSPIXEL );
			ReleaseDC( NULL, hDC );
		}
		break;
	}

	return byBpp;
}

// GetFrameCount
int CMImg::GetFrameCount( )
{
	int iFrameCount = -1;

	if ( m_pImage )
		iFrameCount = max( 1, m_pImage->GetNumFrames( ) );

	return iFrameCount;
}

// GetFrameIconSize
DWORD CMImg::GetFrameIconSize( int iFrame )
{
	switch ( iFrame )
	{
	case MIMG_FRAME_ICON_XXXXLARGE:
	case MIMG_FRAME_ICON_XXXXLARGE_4BPP:
	case MIMG_FRAME_ICON_XXXXLARGE_8BPP:
	case MIMG_FRAME_ICON_XXXXLARGE_24BPP:
		return 256;
	case MIMG_FRAME_ICON_XXXLARGE:
	case MIMG_FRAME_ICON_XXXLARGE_4BPP:
	case MIMG_FRAME_ICON_XXXLARGE_8BPP:
	case MIMG_FRAME_ICON_XXXLARGE_24BPP:
		return 128;
	case MIMG_FRAME_ICON_XXLARGE:
	case MIMG_FRAME_ICON_XXLARGE_4BPP:
	case MIMG_FRAME_ICON_XXLARGE_8BPP:
	case MIMG_FRAME_ICON_XXLARGE_24BPP:
		return 64;
	case MIMG_FRAME_ICON_XLARGE:
	case MIMG_FRAME_ICON_XLARGE_4BPP:
	case MIMG_FRAME_ICON_XLARGE_8BPP:
	case MIMG_FRAME_ICON_XLARGE_24BPP:
		return 48;
	case MIMG_FRAME_ICON_LARGE:
	case MIMG_FRAME_ICON_LARGE_4BPP:
	case MIMG_FRAME_ICON_LARGE_8BPP:
	case MIMG_FRAME_ICON_LARGE_24BPP:
		return 32;
	case MIMG_FRAME_ICON_MEDIUM:
	case MIMG_FRAME_ICON_MEDIUM_4BPP:
	case MIMG_FRAME_ICON_MEDIUM_8BPP:
	case MIMG_FRAME_ICON_MEDIUM_24BPP:
		return 24;
	default:
		return 16; // MIMG_FRAME_ICON_SMALL
	}
}

// GetGIFCompr
int CMImg::GetGIFCompr( )
{
	return m_iGIFCompr;
}

// GetJPGQuality
BYTE CMImg::GetJPGQuality( )
{
	return m_byJPGQuality;
}

// GetMaxOpacity
BYTE CMImg::GetMaxOpacity( )
{
	if ( !m_pImage )
		return 0;

	return m_pImage->AlphaGetMax( );
}

// GetPaletteColor
COLORREF CMImg::GetPaletteColor( BYTE byIdx )
{
	if ( !( m_pImage && HasPalette( ) && byIdx < m_pImage->GetNumColors( ) ) )
		return MIMG_COLOR_UNDEF;

	RGBQUAD r = m_pImage->GetPaletteColor( byIdx );
	return RGB( r.rgbRed, r.rgbGreen, r.rgbBlue );
}

// GetPixelColor
COLORREF CMImg::GetPixelColor( int iX, int iY )
{
	if ( !m_pImage )
		return MIMG_COLOR_UNDEF;

	TransformXY( iX, iY );

	RGBQUAD r = m_pImage->GetPixelColor( iX, iY );
	return RGB( r.rgbRed, r.rgbGreen, r.rgbBlue );
}

// GetPixelOpacity
BYTE CMImg::GetPixelOpacity( int iX, int iY )
{
	if ( !( m_pImage && m_pImage->AlphaIsValid( ) ) )
		return 255;

	TransformXY( iX, iY );
	return m_pImage->AlphaGet( iX, iY );
}

// GetSize
BOOL CMImg::GetSize( LPSIZE pSize )
{
	if ( !( m_pImage && pSize ) )
		return FALSE;

	pSize->cx = m_pImage->GetWidth( );
	pSize->cy = m_pImage->GetHeight( );

	return TRUE;
}

// GetString
BOOL CMImg::GetString( LPHSTRING phs, DWORD dwType, BOOL bBase64 /* = FALSE */, BOOL bNullTerminated /* = FALSE */ )
{
	if ( !( m_pImage && phs ) )
		return FALSE;

	// Temporäre Datei erzeugen
	TCHAR szPath[1024];
	DWORD dwLen = GetTempPath( 1024, szPath );
	if ( !dwLen )
		return FALSE;

	TCHAR szFile[1024];
	if ( !GetTempFileName( szPath, _T("MIM"), 0, szFile ) )
		return FALSE;

	if ( !Save( szFile, dwType ) )
		return FALSE;

	// Memory-Datei für evtl. Base64-Encoding
	CxMemFile file64( NULL, 0 );

	// Temporäre Datei einlesen
	BOOL bOk = FALSE;
	FILE *hFile = NULL;

	for ( ; ; )
	{		
		hFile = _tfopen( szFile, _T("rb") );
		if ( !hFile )
			break;

		CxIOFile file( hFile );
		long lSize = file.Size( );
		if ( lSize <= 0 )
			break;

		if ( bBase64 )
		{
			static const char cb64[]="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

			unsigned char in[3], out[4];
			int	i, blocksout = 0;
			WORD len, linesize = 72;

			if ( !file64.Open( ) )
				break;
			while( !file.Eof( ) )
			{
				len = 0;
				for( i = 0; i < 3; i++ )
				{
					in[i] = (unsigned char) file.GetC( );
					if( !file.Eof( ) )
						len++;
					else
						in[i] = 0;
				}
				if ( len )
				{
					out[0] = cb64[ in[0] >> 2 ];
					out[1] = cb64[ ((in[0] & 0x03) << 4) | ((in[1] & 0xf0) >> 4) ];
					out[2] = (unsigned char) (len > 1 ? cb64[ ((in[1] & 0x0f) << 2) | ((in[2] & 0xc0) >> 6) ] : '=');
					out[3] = (unsigned char) (len > 2 ? cb64[ in[2] & 0x3f ] : '=');

					for( i = 0; i < 4; i++ )
						file64.PutC( out[i] );
					blocksout++;
				}
				if( blocksout >= ( linesize/4 ) || file.Eof( ) )
				{
					if( blocksout )
					{
						file64.PutC( '\r' );
						file64.PutC( '\n' );
					}
					blocksout = 0;
				}
			}
			long lSize = file64.Size( );			
			if ( !Ctd.InitLPHSTRINGParam( phs, lSize ) )
				break;

			long lLen;
			void *pBuf = Ctd.StringGetBuffer( *phs, &lLen );
			if (pBuf)
			{
				BOOL bBreak = FALSE;

				if (file64.Seek(0, SEEK_SET))
				{
					if (file64.Read(pBuf, 1, lSize) != size_t(lSize))
						bBreak = TRUE;
				}
				else
					bBreak = TRUE;

				Ctd.HStrUnlock(*phs);

				if (bBreak)
					break;
			}
			else
				break;
		}
		else
		{
			if ( !Ctd.InitLPHSTRINGParam( phs, lSize + (bNullTerminated ? sizeof(TCHAR) : 0) ) )
				break;

			long lLen;
			void *pBuf = Ctd.StringGetBuffer( *phs, &lLen );
			if (pBuf)
			{
				BOOL bBreak = FALSE;

				if (file.Read(pBuf, 1, lSize) == size_t(lSize))
				{
					if (bNullTerminated)
					{
						LPBYTE lpb = (LPBYTE)pBuf;
						lpb += lSize;
						for (size_t s = 1; s <= sizeof(TCHAR); s++, lpb++)
						{
							lpb = 0;
						}
					}
				}
				else
					bBreak = TRUE;

				Ctd.HStrUnlock(*phs);

				if (bBreak)
					break;
			}
			else
				break;				
		}

		bOk = TRUE;
		break;
	}

	// Dateien schließen
	if ( hFile )
		fclose( hFile );

	file64.Close( );

	// Temporäre Datei löschen
	DeleteFile( szFile ); 

	return bOk;
}

// GetTIFCompr
int CMImg::GetTIFCompr( )
{
	return m_iTIFCompr;
}

// GetTranspColor
DWORD CMImg::GetTranspColor( )
{
	if ( !m_pImage )
		return MIMG_COLOR_UNDEF;

	int iIdx = m_pImage->GetTransIndex( );
	if ( iIdx < 0 )
		return MIMG_COLOR_UNDEF;
	else
		return m_pImage->RGBQUADtoRGB( m_pImage->GetTransColor( ) );
}

// GetType
int CMImg::GetType( )
{
	if ( !m_pImage )
		return MIMG_TYPE_UNKNOWN;

	return m_pImage->GetType( );
}

// GetUsedColors
int CMImg::GetUsedColors( vector<COLORREF> &vColors )
{
	if ( !m_pImage )
		return -1;

	vColors.clear( );
	DWORD dwColors = m_pImage->GetNumColors( );
	if ( !dwColors )
		dwColors = 256;
	vColors.reserve( dwColors );

	COLORREF clr, clrLast = CLR_INVALID;
	DWORD dwX, dwY;
	DWORD dwWid = m_pImage->GetWidth( );
	DWORD dwHt = m_pImage->GetHeight( );
	for ( dwY = 0; dwY < dwHt; dwY++ )
	{
		for ( dwX = 0; dwX < dwWid; dwX++ )
		{
			clr = m_pImage->RGBQUADtoRGB( m_pImage->GetPixelColor( dwX, dwHt - dwY - 1 ) );
			if ( clr != clrLast && find( vColors.begin( ), vColors.end( ), clr ) == vColors.end( ) )
			{
					vColors.push_back( clr );
			}
		}
	}

	return vColors.size( );
}

// GetUsedColors
int CMImg::GetUsedColors( HARRAY haColors )
{
	vector<COLORREF> vColors;
	int iColors = GetUsedColors( vColors );
	if ( iColors < 0 )
		return -1;

	// Array checken
	long lDims;
	if ( !Ctd.ArrayDimCount( haColors, &lDims ) )
		return -1;
	if ( lDims > 1 )
		return -1;

	// Untergrenze des Arrays ermitteln
	long lLower;
	if ( !Ctd.ArrayGetLowerBound( haColors, 1, &lLower ) )
		return -1;

	// Übertragen
	NUMBER n;
	for ( int i = 0; i < iColors; ++i )
	{
		if ( !Ctd.CvtULongToNumber( vColors[i], &n ) )
			return -1;
		if ( !Ctd.MDArrayPutNumber( haColors, &n, i ) )
			return -1;
	}

	return iColors;
}

// GetWMFLoadSize
BOOL CMImg::GetWMFLoadSize( SIZE &si )
{
	si = m_siWMF;
	return TRUE;
}

// GrayScale
BOOL CMImg::GrayScale( )
{
	if ( !m_pImage )
		return FALSE;

	return m_pImage->GrayScale( );
}

// HasOpacity
BOOL CMImg::HasOpacity( )
{
	if ( !m_pImage )
		return FALSE;

	if ( GetTranspColor( ) != MIMG_COLOR_UNDEF )
		return TRUE;

	if ( m_pImage->AlphaIsValid( ) )
		return TRUE;

	if ( m_pImage->AlphaPaletteIsEnabled( ) )
		return TRUE;

	return FALSE;
}

// Invert
BOOL CMImg::Invert( )
{
	if ( !m_pImage )
		return FALSE;

	return m_pImage->Negative( );
}

// IsAlphaFullyOpaque
BOOL CMImg::IsAlphaFullyOpaque( )
{
	if ( !m_pImage || !m_pImage->AlphaIsValid( ) )
		return TRUE;

	DWORD dwX, dwY;
	DWORD dwWid = m_pImage->GetWidth( );
	DWORD dwHt = m_pImage->GetHeight( );
	for ( dwY = 0; dwY < dwHt; dwY++ )
	{
		for ( dwX = 0; dwX < dwWid; dwX++ )
		{
			if ( GetPixelOpacity( (int)dwX, (int)dwY ) != 255 )
				return FALSE;
		}
	}

	return TRUE;
}

// SetLightness
BOOL CMImg::SetLightness( long lLevel, long lContrast )
{
	if ( !m_pImage )
		return FALSE;

	return m_pImage->Light( lLevel, lContrast );
}

// HasPalette
BOOL CMImg::HasPalette( )
{
	if ( !m_pImage )
		return FALSE;

	return ( m_pImage->GetPalette( ) != NULL );
}

// LoadEmpty
BOOL CMImg::LoadEmpty( int iWid, int iHt, COLORREF clr, WORD wBpp, DWORD dwImgType )
{
	if ( iWid <= 0 || iHt <= 0 )
		return FALSE;

	if ( dwImgType == MIMG_TYPE_UNKNOWN )
		dwImgType = MIMG_TYPE_BMP;

	FreeImageData( );

	CxImage *pImg = new CxImage( iWid, iHt, wBpp, dwImgType );
	if ( !pImg )
		return FALSE;

	m_pImage = pImg;

	int iIdx = -1;
	if ( wBpp < 24 )
	{
		if ( wBpp == 1 )
		{
			m_pImage->SetPaletteColor( 0, 0, 0, 0 );
			m_pImage->SetPaletteColor( 1, 255, 255, 255);
		}
		else
			m_pImage->SetStdPalette( );

		iIdx = m_pImage->GetNearestIndex( m_pImage->RGBtoRGBQUAD( clr ) );
	}

	int x, y;
	for ( y = 0; y < iHt; ++y )
	{
		for ( x = 0; x < iWid; ++x )
		{
			if ( iIdx >= 0 )
				m_pImage->SetPixelIndex( x, y, (BYTE)iIdx );
			else
				m_pImage->SetPixelColor( x, y, clr );
		}
	}

	return TRUE;
}

// LoadFromBuffer
BOOL CMImg::LoadFromBuffer( BYTE *pBuffer, DWORD dwSize, DWORD dwImgType, int iFrame /* = 0 */ )
{
	if ( !pBuffer )
		return FALSE;

	if ( dwImgType == MIMG_TYPE_UNKNOWN )
	{
		dwImgType = DetermineType( pBuffer, dwSize );
		if ( dwImgType == MIMG_TYPE_UNKNOWN )
			return FALSE;
	}

	FreeImageData( );

	BYTE byDesiredBpp = 0;
	DWORD dwDesiredSize = 0;
	if ( dwImgType == MIMG_TYPE_ICO && iFrame < 0 )
	{
		BOOL bBppOk, bSizeOk;
		BYTE byBpp, byBppDiff;
		DWORD dwWid, dwWidDiff;
		vector<BYTE> vBppDiff;
		vector<DWORD> vSizeDiff;

		dwDesiredSize = GetFrameIconSize( iFrame );
		byDesiredBpp = GetFrameBpp( iFrame );

		CxImage *pImg;
		for ( int i = 0; ; ++i )
		{
			pImg = NULL;
			pImg = new CxImage( dwImgType );
			if ( !pImg )
				return FALSE;

			pImg->SetFrame( i );
			if ( !pImg->Decode( pBuffer, dwSize, dwImgType ) )
			{
				delete pImg;
				break;
			}

			if ( !pImg->IsValid( ) )
			{
				delete pImg;
				break;
			}

			dwWid = pImg->GetWidth( );
			dwWidDiff = max( dwWid, dwDesiredSize ) - min( dwWid, dwDesiredSize );
			vSizeDiff.push_back( dwWidDiff );
			bSizeOk = ( dwWidDiff == 0 );

			byBpp = (BYTE)pImg->GetBpp( );
			byBppDiff = max( byBpp, byDesiredBpp ) - min( byBpp, byDesiredBpp );
			vBppDiff.push_back( byBppDiff );
			bBppOk = ( byBppDiff == 0 );

			delete pImg;

			if ( bSizeOk && bBppOk )
				iFrame = i;

			if ( iFrame >= 0 )
				break;
		}

		if ( iFrame < 0 )
		{
			BYTE byMaxBppDiff = UCHAR_MAX;
			DWORD dwMaxSizeDiff = ULONG_MAX;
			int i, iSize = vSizeDiff.size( );
			for ( i = 0; i < iSize; ++i )
			{

				if ( vSizeDiff[i] <= dwMaxSizeDiff )
				{
					if ( vSizeDiff[i] < dwMaxSizeDiff )
						byMaxBppDiff = UCHAR_MAX;

					dwMaxSizeDiff = vSizeDiff[i];
					if ( vBppDiff[i] <= byMaxBppDiff )
					{
						byMaxBppDiff = vBppDiff[i];
						iFrame = i;
					}
				}
			}
		}
	}

	m_pImage = new CxImage( dwImgType );
	if ( !m_pImage )
		return FALSE;

	if ( dwImgType == MIMG_TYPE_WMF && ( m_siWMF.cx || m_siWMF.cy ) )
	{
		CxImageWMF WMFImg;
		CxMemFile MemFile( pBuffer, dwSize );
		if ( !WMFImg.Decode( &MemFile, m_siWMF.cx, m_siWMF.cy ) )
			return FALSE;
		m_pImage->Copy( WMFImg, TRUE, TRUE, TRUE );
	}
	else
	{
		m_pImage->SetFrame( max( 0, iFrame ) );
		m_pImage->Decode( pBuffer, dwSize, dwImgType );
	}
	
	if ( !m_pImage->IsValid( ) )
		return FALSE;

	if ( dwDesiredSize > 0 && ( m_pImage->GetWidth( ) != dwDesiredSize || m_pImage->GetHeight( ) != dwDesiredSize ) )
		Resample( int( dwDesiredSize ), int( dwDesiredSize ), MIMG_RESAMPLE_NC );

	if ( byDesiredBpp > 0 && ( m_pImage->GetBpp( ) != byDesiredBpp ) )
		SetBpp( byDesiredBpp );

	return TRUE;
}

// LoadFromClipboard
BOOL CMImg::LoadFromClipboard( )
{
	/*if ( !IsClipboardFormatAvailable( CF_DIB ) )
		return FALSE;*/

	BOOL bOk = FALSE;
	if ( OpenClipboard( NULL ) )
	{
		HANDLE hGlobal = GetClipboardData( CF_DIB );
		if ( hGlobal )
		{
			FreeImageData( );
			m_pImage = new CxImage( MIMG_TYPE_BMP );
			bOk = m_pImage->CreateFromHANDLE( hGlobal );
		}
		else
		{
			HENHMETAFILE hmf = (HENHMETAFILE)GetClipboardData( CF_ENHMETAFILE );
			if ( hmf )
			{				
				UINT uiSize = GetEnhMetaFileBits( hmf, 0, NULL );
				if ( uiSize > 0 )
				{
					BYTE *pBuf = new BYTE[uiSize];
					if ( uiSize = GetEnhMetaFileBits( hmf, uiSize, pBuf ) )
					{
						bOk = LoadFromBuffer( pBuf, uiSize, MIMG_TYPE_WMF );
					}
					delete [] pBuf;
				}
			}
		}

		CloseClipboard( );
	}

	return bOk;
}

// LoadFromCResource
BOOL CMImg::LoadFromCResource( HMODULE hModule, HRSRC hRes, DWORD dwImgType, int iFrame /* = 0 */ )
{
	if ( !hRes )
		return FALSE;

	DWORD dwSize = SizeofResource( hModule, hRes );
	if ( !dwSize )
		return FALSE;
	HGLOBAL hMem = LoadResource( hModule, hRes );
	if ( !hMem)
		return FALSE;
	BYTE *pBuffer = (BYTE*)LockResource( hMem );
	if ( !pBuffer )
		return FALSE;

	if ( dwImgType == MIMG_TYPE_ICO )
	{
		// Gewünschte Größe ermitteln
		BYTE byDesiredSize = 0;
		if ( iFrame < 0 )
		{			
			byDesiredSize = (BYTE)GetFrameIconSize( iFrame );
		}

		// Directory-Eintrag des gewünschten Frames ermitteln
		BOOL bFound = FALSE;
		CxImageICO::tagIconDir *pHeader = (CxImageICO::tagIconDir*)pBuffer;
		if ( pHeader->idCount > 0 )
		{
			RESICONDIRENTRY **pEntries = new RESICONDIRENTRY*[pHeader->idCount];
			RESICONDIRENTRY *pEntry = NULL;
			for ( int i = 0; i < pHeader->idCount; ++i )
			{
				pEntries[i] = (RESICONDIRENTRY*)(pBuffer + sizeof(CxImageICO::tagIconDir) + (i * sizeof(RESICONDIRENTRY)));
				pEntry = pEntries[i];
				if ( iFrame >= 0 && i == iFrame )
				{
					bFound = TRUE;
					break;
				}
				else if ( pEntry->bWidth == byDesiredSize )
				{
					bFound = TRUE;
					break;
				}
			}

			if ( iFrame < 0 && !bFound )
			{
				// Das der gewünschten Größe am nächsten kommende Icon verwenden
				int iDiff = abs( byDesiredSize - pEntries[0]->bWidth ), iEntry = 0;
				for ( int i = 1; i < pHeader->idCount; ++i )
				{
					if ( abs( byDesiredSize - pEntries[i]->bWidth ) < iDiff )
					{
						iDiff = abs( byDesiredSize - pEntries[i]->bWidth );
						iEntry = i;
					}
				}

				pEntry = pEntries[iEntry];
				bFound = TRUE;
			}

			delete []pEntries;

			if ( bFound )
			{
				// Icon laden
				HRSRC hResIco = FindResource( hModule, MAKEINTRESOURCE( pEntry->wID ), RT_ICON );
				if ( !hResIco )
					return FALSE;
				DWORD dwSizeIco = SizeofResource( hModule, hResIco );
				if ( !dwSizeIco )
					return FALSE;
				HGLOBAL hMemIco = LoadResource( hModule, hResIco );
				if ( !hMemIco)
					return FALSE;
				BYTE *pBufferIco = (BYTE*)LockResource( hMemIco );
				if ( !pBufferIco )
					return FALSE;

				// "Datei-Icon" bauen
				int iSize = sizeof(CxImageICO::tagIconDir) + sizeof(CxImageICO::tagIconDirectoryEntry) + dwSizeIco;
				BYTE *pIco = new BYTE[iSize];
				CxImageICO::tagIconDir hdr = {0,1,1};
				CxImageICO::tagIconDirectoryEntry de;
				memcpy( &de, pEntry, sizeof(RESICONDIRENTRY) );
				de.dwImageOffset = sizeof(CxImageICO::tagIconDir) + sizeof(CxImageICO::tagIconDirectoryEntry);
				memcpy( pIco, &hdr, sizeof(CxImageICO::tagIconDir) );
				memcpy( pIco + sizeof(CxImageICO::tagIconDir), &de, sizeof(CxImageICO::tagIconDirectoryEntry) );
				memcpy( pIco + sizeof(CxImageICO::tagIconDir) + sizeof(CxImageICO::tagIconDirectoryEntry), pBufferIco, dwSizeIco );

				BOOL bOk = LoadFromBuffer( pIco, iSize, dwImgType, iFrame < 0 ? iFrame : 0 );
				delete []pIco;
				return bOk;
			}
		}
	}
	else
		return LoadFromBuffer( pBuffer, dwSize, dwImgType, iFrame );

	return FALSE;
}

// LoadFromFile
BOOL CMImg::LoadFromFile( LPTSTR psFile, DWORD dwImgType, int iFrame /* = 0 */ )
{
	if ( !psFile )
		return FALSE;

	FILE *hFile;
	hFile = _tfopen( psFile, _T("rb") );
	if ( !hFile )
		return FALSE;
	CxIOFile cf( hFile );
	DWORD dwSize = cf.Size( );
	BYTE *pBuffer = new BYTE[dwSize];
	cf.Read( pBuffer, dwSize, 1 );
	fclose( hFile );
	
	BOOL bOk = LoadFromBuffer( pBuffer, dwSize, dwImgType, iFrame );
	
	delete []pBuffer;

	return bOk;
}

// LoadFromFileInfo
BOOL CMImg::LoadFromFileInfo( LPCTSTR psFile, DWORD dwFlags )
{
	CoInitialize( NULL );

	SHFILEINFO sfi;
	ZeroMemory( &sfi, sizeof( SHFILEINFO ) );

	DWORD dwFileAttributes = 0;
	UINT uiFlags = SHGFI_ICON | SHGFI_SHELLICONSIZE;
	if ( dwFlags & MIMG_FI_ICON_LARGE )
		uiFlags |= SHGFI_LARGEICON;
	if ( dwFlags & MIMG_FI_ICON_SMALL )
		uiFlags |= SHGFI_SMALLICON;
	if ( dwFlags & MIMG_FI_ICON_OPEN )
		uiFlags |= SHGFI_OPENICON;

	// Spezial-Ordner?
	int iSpecialFolder = 0;
	LPITEMIDLIST lpList = NULL;

	if ( lstrcmp( psFile, MIMG_FI_SPECIAL_CONTROLPANEL ) == 0 )
		iSpecialFolder = CSIDL_CONTROLS;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_FAVORITES ) == 0 )
		iSpecialFolder = CSIDL_FAVORITES;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_INTERNET ) == 0 )
		iSpecialFolder = CSIDL_INTERNET;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_NETWORK ) == 0 )
		iSpecialFolder = CSIDL_NETWORK;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_MYCOMPUTER ) == 0 )
		iSpecialFolder = CSIDL_DRIVES;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_MYDOCUMENTS ) == 0 )
	{
		LPSHELLFOLDER lpDesktop;
		if ( SHGetDesktopFolder( &lpDesktop ) != NOERROR )
			return FALSE;
		lpDesktop->ParseDisplayName( NULL, NULL, L"::{450d8fba-ad25-11d0-98a8-0800361b1103}", NULL, &lpList, NULL );
		lpDesktop->Release( );
	}
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_PRINTERS ) == 0 )
		iSpecialFolder = CSIDL_PRINTERS;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_PROGRAMS ) == 0 )
		iSpecialFolder = CSIDL_PROGRAMS;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_RECENT ) == 0 )
		iSpecialFolder = CSIDL_RECENT;
	else if ( lstrcmp( psFile, MIMG_FI_SPECIAL_RECYCLEBIN ) == 0 )
		iSpecialFolder = CSIDL_BITBUCKET;

	if ( iSpecialFolder )
	{
		if ( SHGetSpecialFolderLocation( NULL, iSpecialFolder, &lpList ) != NOERROR )
			return FALSE;
	}

	if ( lpList )
	{
		uiFlags |= SHGFI_PIDL;
		SHGetFileInfo( (LPCTSTR)lpList, dwFileAttributes, &sfi, sizeof( sfi ), uiFlags );

		LPMALLOC lpMalloc;
		if ( SHGetMalloc( &lpMalloc ) == NOERROR )
			lpMalloc->Free( lpList );
	}
	else
	{
		// Existiert die Datei?
		if ( psFile && lstrlen( psFile ) > 0 )
		{
			HANDLE hFile = CreateFile( psFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL );
			if ( hFile != INVALID_HANDLE_VALUE )
			{
				CloseHandle( hFile );
			}
			else
			{
				// Aha, wir wollen also das Icon für eine Dateigruppe
				uiFlags |= SHGFI_USEFILEATTRIBUTES;
				dwFileAttributes = FILE_ATTRIBUTE_NORMAL;
			}
		}

		SHGetFileInfo( psFile, dwFileAttributes, &sfi, sizeof( sfi ), uiFlags );
	}

	if ( sfi.hIcon )
	{
		BOOL bOk = LoadFromHandle( sfi.hIcon, MIMG_TYPE_ICO );
		DestroyIcon( sfi.hIcon );
		return bOk;
	}

	return FALSE;
}

// LoadFromHandle
BOOL CMImg::LoadFromHandle( HANDLE hImg, DWORD dwImgType )
{
	if ( !hImg )
		return FALSE;
	if ( !( dwImgType == MIMG_TYPE_BMP || dwImgType == MIMG_TYPE_ICO ) )
		return FALSE;

	FreeImageData();
	m_pImage = new CxImage( dwImgType );

	if ( dwImgType == MIMG_TYPE_BMP )
		CreateFromHBITMAP( HBITMAP( hImg ) );
	else
		CreateFromHICON( HICON( hImg ) );

	if ( !m_pImage->IsValid( ) )
		return FALSE;

	return TRUE;
}

// LoadFromModule
BOOL CMImg::LoadFromModule( LPCTSTR lpsModule, DWORD dwImgType, int iIdx, int iFrame /*= 0*/ )
{
	if ( !lpsModule )
		return FALSE;

	BOOL bOk = FALSE;

	HMODULE hModule = HMODULE( LoadLibraryEx( lpsModule, NULL, LOAD_LIBRARY_AS_DATAFILE ) );
	if ( hModule )
	{
		LPTSTR lpType = NULL;
		if ( dwImgType == MIMG_TYPE_ICO )
			lpType = RT_GROUP_ICON;
		else if ( dwImgType == MIMG_TYPE_BMP )
			lpType = RT_BITMAP;

		if ( lpType )
		{
			LPTSTR lpName = NULL;
			if ( iIdx < 0 )
			{
				lpName = MAKEINTRESOURCE( abs( iIdx ) );
			}
			else if ( iIdx > 0 )
			{
				ENUMRESNAME ern = {iIdx,1,NULL};
				EnumResourceNames( hModule, lpType, (ENUMRESNAMEPROC)EnumResNameProc, (LPARAM)&ern );
				lpName = ern.lpName;
			}

			if ( lpName )
			{
				HRSRC hRes = FindResource( hModule, lpName, lpType );
				if ( hRes )
				{
					bOk = LoadFromCResource( hModule, hRes, dwImgType, iFrame );
				}
			}
		}
	
		FreeLibrary( hModule );
	}

	return bOk;
}

// LoadFromModuleNamed
BOOL CMImg::LoadFromModuleNamed( LPCTSTR lpsModule, DWORD dwImgType, LPCTSTR lpsResName, int iFrame /*= 0*/ )
{
	if ( !lpsModule )
		return FALSE;

	if ( !lpsResName )
		return FALSE;

	BOOL bOk = FALSE;

	HMODULE hModule = HMODULE( LoadLibraryEx( lpsModule, NULL, LOAD_LIBRARY_AS_DATAFILE ) );
	if ( hModule )
	{
		LPTSTR lpType = NULL;
		if ( dwImgType == MIMG_TYPE_ICO )
			lpType = RT_GROUP_ICON;
		else if ( dwImgType == MIMG_TYPE_BMP )
			lpType = RT_BITMAP;

		if ( lpType )
		{
			HRSRC hRes = FindResource( hModule, lpsResName, lpType );
			if ( hRes )
			{
				bOk = LoadFromCResource( hModule, hRes, dwImgType, iFrame );
			}
		}
	
		FreeLibrary( hModule );
	}

	return bOk;
}

// LoadFromPicture
BOOL CMImg::LoadFromPicture( HWND hWndPic )
{
	BOOL bOk = FALSE;

	HSTRING hsData = 0;
	int iType = PIC_ImageTypeNone;
	long lLen = Ctd.PicGetImage( hWndPic, &hsData, &iType );
	if ( lLen > 0 )
	{
		bOk = LoadFromString( hsData, MIMG_TYPE_UNKNOWN );
	}

	if ( hsData )
		Ctd.HStringUnRef( hsData );

	return bOk;
}

// LoadFromResource
BOOL CMImg::LoadFromResource( HSTRING hsRes, DWORD dwImgType, int iFrame /* = 0 */ )
{
	BOOL bOk = FALSE;
	long lLen;
	LPTSTR lpsRes = Ctd.StringGetBuffer( hsRes, &lLen );

	if (lpsRes)
	{
		CTDRESID dwResId = Ctd.ResId(lpsRes);
		Ctd.HStrUnlock(hsRes);

		if (dwResId)
		{
			HSTRING hsData = 0;
			if (Ctd.ResLoad(dwResId, &hsData))
				bOk = LoadFromString(hsData, dwImgType, iFrame);

			if (hsData)
				Ctd.HStringUnRef(hsData);
		}		
	}	

	return bOk;
}

// LoadFromScreen
BOOL CMImg::LoadFromScreen( HWND hWnd, DWORD dwFlags )
{
	BOOL bOk = FALSE;

	BOOL bWndClAreaOnly = ( dwFlags & MIMG_SCREEN_WNDCLAREAONLY );
	BOOL bWndForeground = ( dwFlags & MIMG_SCREEN_WNDFOREGROUND );

	if ( bWndForeground )
	{
		hWnd = GetForegroundWindow( );
		if ( !hWnd )
			return FALSE;
	}

	HDC hDC = GetDC( NULL );
	if ( hDC )
	{
		int iHt, iLeft = 0, iTop = 0, iWid;

		if ( hWnd )
		{
			RECT r;
			SetRectEmpty( &r );
			if ( bWndClAreaOnly )
			{
				GetClientRect( hWnd, &r );
				MapWindowPoints( hWnd, NULL, (LPPOINT)&r, 2 );
			}
			else
				GetWindowRect( hWnd, &r );

			iLeft = r.left;
			iTop = r.top;
			iWid = r.right - r.left;
			iHt = r.bottom - r.top;
		}
		else
		{
			iWid = GetDeviceCaps( hDC, HORZRES );
			iHt = GetDeviceCaps( hDC, VERTRES );
		}

		HDC hDCComp = CreateCompatibleDC( hDC );
		if ( hDCComp )
		{
			HBITMAP hBmp = CreateCompatibleBitmap( hDC, iWid, iHt );
			if ( hBmp )
			{
				HGDIOBJ hBmpOld = SelectObject( hDCComp, hBmp );

				BitBlt( hDCComp, 0, 0, iWid, iHt, hDC, iLeft, iTop, SRCCOPY );

				SelectObject( hDCComp, hBmpOld );

				FreeImageData( );
				m_pImage = new CxImage( MIMG_TYPE_BMP );
				bOk = CreateFromHBITMAP( hBmp, NULL );

				DeleteObject( hBmp );
			}

			DeleteDC( hDCComp );
		}

		ReleaseDC( hWnd, hDC );
	}

	return bOk;
}

// LoadFromString
BOOL CMImg::LoadFromString( HSTRING hs, DWORD dwImgType, int iFrame /* = 0 */, BOOL bBase64 /*=FALSE*/ )
{
	BOOL bOk = FALSE;
	long lLen;
	LPBYTE pBuf = LPBYTE( Ctd.StringGetBuffer( hs, &lLen ) );

	if (pBuf)
	{
		if (bBase64)
		{
			static const char cd64[] = "|$$$}rstuvwxyz{$$$$$$$>?@ABCDEFGHIJKLMNOPQRSTUVW$$$$$$XYZ[\\]^_`abcdefghijklmnopq";

			CxMemFile infile(pBuf, lLen), outfile(NULL, 0);
			unsigned char in[4], out[3], v;
			int i, len;

			if (outfile.Open())
			{
				while (!infile.Eof())
				{
					for (len = 0, i = 0; i < 4 && !infile.Eof(); i++)
					{
						v = 0;
						while (!infile.Eof() && v == 0)
						{
							v = (unsigned char)infile.GetC();
							v = (unsigned char)((v < 43 || v > 122) ? 0 : cd64[v - 43]);
							if (v)
								v = (unsigned char)((v == '$') ? 0 : v - 61);
						}
						if (!infile.Eof())
						{
							len++;
							if (v)
								in[i] = (unsigned char)(v - 1);
						}
						else
							in[i] = 0;
					}
					if (len)
					{
						out[0] = (unsigned char)(in[0] << 2 | in[1] >> 4);
						out[1] = (unsigned char)(in[1] << 4 | in[2] >> 2);
						out[2] = (unsigned char)(((in[2] << 6) & 0xc0) | in[3]);
						for (i = 0; i < len - 1; i++)
							outfile.PutC(out[i]);
					}
				}

				LPBYTE pBuf64 = (LPBYTE)outfile.GetBuffer(false);
				bOk = LoadFromBuffer(pBuf64, outfile.Size(), dwImgType, iFrame);
				outfile.Close();
			}
		}
		else
			bOk = LoadFromBuffer(pBuf, lLen, dwImgType, iFrame);

		Ctd.HStrUnlock(hs);
	}
	
	return bOk;
}

// LoadFromVisPic
BOOL CMImg::LoadFromVisPic( LPVOID lpVisPic )
{
	if ( !lpVisPic )
		return FALSE;

	FreeImageData();

	int iOffset;
	DWORD dwType;
	HBITMAP hBmColor = NULL, hBmMask = NULL;
	HICON hIcon = NULL;

	BYTE byFormat = *(LPBYTE( lpVisPic ) + 4);
	if ( byFormat == 1 )
	{
		dwType = MIMG_TYPE_BMP;
		iOffset = 13;
#ifdef CTD7x
		iOffset += 3;
#endif
		hBmColor = *(HBITMAP*)(LPBYTE( lpVisPic ) + iOffset);
		
		iOffset = 17;
#ifdef CTD7x
		iOffset += 3;
#endif
		hBmMask = *(HBITMAP*)(LPBYTE( lpVisPic ) + iOffset);

		m_pImage = new CxImage( dwType );
		CreateFromHBITMAP( hBmColor, hBmMask );
	}
	else
	{
		dwType = MIMG_TYPE_ICO;
		iOffset = 21;
#ifdef CTD7x
		iOffset += 3;
#endif
		hIcon = *(HICON*)(LPBYTE( lpVisPic ) + iOffset);

		m_pImage = new CxImage( dwType );
		CreateFromHICON( hIcon );
	}

	if ( !m_pImage->IsValid( ) )
		return FALSE;

	return TRUE;
}


// Mirror
BOOL CMImg::Mirror( BOOL bHorz )
{
	if ( !m_pImage )
		return FALSE;

	return bHorz ? m_pImage->Mirror( ) : m_pImage->Flip( );
}

// Print
BOOL CMImg::Print( HDC hDC, LPCRECT pRect, COLORREF clrPaper, DWORD dwFlags )
{
    if ( !( m_pImage && hDC && pRect ) )
		return FALSE;

	CMImg *pMImg = CreateCopy( );
	if ( !pMImg )
		return FALSE;

	pMImg->StripOpacity( clrPaper );

	BOOL bOk = pMImg->Draw( hDC, pRect, dwFlags );

	delete pMImg;

	return bOk;
}

// PutImage
BOOL CMImg::PutImage( long lX, long lY, CMImg * pImgPut )
{
	if ( !( m_pImage && pImgPut && pImgPut->m_pImage ) )
		return FALSE;

	long lWid = m_pImage->GetWidth( );
	long lHt = m_pImage->GetHeight( );

	long lWidPut = pImgPut->m_pImage->GetWidth( );
	long lHtPut = pImgPut->m_pImage->GetHeight( );

	BYTE op, op1, opd;
	BOOL bSetOp;
	COLORREF clr, clrPut, clrPutTransp = pImgPut->GetTranspColor( ), clrTransp = GetTranspColor( );
	int xpt, ypt;
	long x, y, xp, yp;
	RGBQUAD rgb;
	for ( y = lY, yp = 0; y < lHt && yp < lHtPut; ++y, ++yp )
	{
		if ( y >= 0 )
		{
			for ( x = lX, xp = 0; x < lWid && xp < lWidPut; ++x, ++xp )
			{
				if ( x >= 0 )
				{
					clrPut = pImgPut->GetPixelColor( xp, yp );
					if ( clrPut == clrPutTransp )
						op = 0;
					else
					{
						op = pImgPut->GetPixelOpacity( xp, yp );
						op =(BYTE)( ( op * ( 1 + pImgPut->GetMaxOpacity( ) ) ) >> 8 );
						if ( pImgPut->HasPalette( ) && pImgPut->m_pImage->AlphaPaletteIsEnabled( ) )
						{
							xpt = xp;
							ypt = yp;
							pImgPut->TransformXY( xpt, ypt );
							rgb = pImgPut->m_pImage->GetPaletteColor( pImgPut->m_pImage->GetPixelIndex( xpt, ypt ) );
							op = (BYTE)( ( op * ( 1 + rgb.rgbReserved ) ) >> 8 );
						}
					}

					if ( op )
					{
						bSetOp = FALSE;

						if ( op == 255 )
						{
							clr = clrPut;
							bSetOp = ( op != GetPixelOpacity( x, y ) );
						}
						else
						{
							clr = GetPixelColor( x, y );

							if ( clr == clrTransp )
								opd = 0;
							else
							{
								opd = GetPixelOpacity( x, y );
								//opd =(BYTE)( ( opd * ( 1 + GetMaxOpacity( ) ) ) >> 8 );
								if ( HasPalette( ) && m_pImage->AlphaPaletteIsEnabled( ) )
								{
									xpt = x;
									ypt = y;
									TransformXY( xpt, ypt );
									rgb = m_pImage->GetPaletteColor( m_pImage->GetPixelIndex( xpt, ypt ) );
									opd = (BYTE)( ( opd * ( 1 + rgb.rgbReserved ) ) >> 8 );
								}
							}

							if ( opd )
							{
								op1 = (BYTE)~op;
								rgb.rgbRed = (BYTE)( ( GetRValue( clr ) * op1 + op * GetRValue( clrPut ) ) >> 8 );
								rgb.rgbGreen = (BYTE)( ( GetGValue( clr ) * op1 + op * GetGValue( clrPut ) ) >> 8 );
								rgb.rgbBlue = (BYTE)( ( GetBValue( clr ) * op1 + op * GetBValue( clrPut ) ) >> 8 );
								clr = RGB( rgb.rgbRed, rgb.rgbGreen, rgb.rgbBlue );
							}
							else
							{
								clr = clrPut;
								bSetOp = bSetOp = ( op != GetPixelOpacity( x, y ) );
							}
						}

						SetPixelColor( x, y, clr );
						if ( bSetOp )
							SetPixelOpacity( x, y, op );
					}
				}
			}
		}
	}

	return TRUE;
}

// PutText
BOOL CMImg::PutText( long lX, long lY, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr )
{
	/*if ( !( m_pImage && lpsText && lpsFontName ) )
		return FALSE;

	HDC hDC = GetDC( NULL );
	if ( !hDC )
		return FALSE;

	CxImage::CXTEXTINFO cxti;
	m_pImage->InitTextInfo( &cxti );

	cxti.lfont.lfHeight = -MulDiv( lFontSize, GetDeviceCaps( hDC, LOGPIXELSY ), 72 );
	cxti.lfont.lfWeight = ( dwFontEnh & FONT_EnhBold ? FW_BOLD : FW_NORMAL );
	cxti.lfont.lfItalic = ( dwFontEnh & FONT_EnhItalic ? TRUE : FALSE );
	cxti.lfont.lfUnderline = ( dwFontEnh & FONT_EnhUnderline ? TRUE : FALSE );
	cxti.lfont.lfStrikeOut = ( dwFontEnh & FONT_EnhStrikeOut ? TRUE : FALSE );
	cxti.lfont.lfCharSet = DEFAULT_CHARSET;
	cxti.lfont.lfOutPrecision = OUT_DEFAULT_PRECIS;
	cxti.lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	cxti.lfont.lfQuality = ( dwFontEnh & MIMG_FONT_ENH_CLEARTYPE ? 5 : PROOF_QUALITY );
	cxti.lfont.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
	lstrcpyn( cxti.lfont.lfFaceName, lpsFontName, LF_FACESIZE );

	cxti.fcolor = clr;
	cxti.bcolor = 0;
	cxti.opaque = FALSE;
	cxti.b_opacity = 0.0;
	cxti.b_outline = 0;
	cxti.b_round = 0;

	cxti.align = DT_LEFT;

	long lCopy = sizeof( cxti.text ) / sizeof( TCHAR );
	lstrcpyn( cxti.text, lpsText, lCopy );	

	m_pImage->DrawStringEx( hDC, lX, lY, &cxti, FALSE );

	ReleaseDC( NULL, hDC );

	return TRUE;*/

	RECT r = {lX,lY,0,0};
	return PutTextAligned( &r, lpsText, lpsFontName, lFontSize, dwFontEnh, clr, 0 );
}

// PutTextAligned
BOOL CMImg::PutTextAligned( LPRECT pRect, LPCTSTR lpsText, LPCTSTR lpsFontName, long lFontSize, DWORD dwFontEnh, COLORREF clr, DWORD dwAlignFlags )
{
	if ( !( m_pImage && pRect && lpsText && lpsFontName ) )
		return FALSE;

	HDC hDC = GetDC( NULL );
	if ( !hDC )
		return FALSE;

	CxImage::CXTEXTINFO cxti;
	m_pImage->InitTextInfo( &cxti );

	cxti.lfont.lfHeight = -MulDiv( lFontSize, GetDeviceCaps( hDC, LOGPIXELSY ), 72 );
	cxti.lfont.lfWeight = ( dwFontEnh & FONT_EnhBold ? FW_BOLD : FW_NORMAL );
	cxti.lfont.lfItalic = ( dwFontEnh & FONT_EnhItalic ? TRUE : FALSE );
	cxti.lfont.lfUnderline = ( dwFontEnh & FONT_EnhUnderline ? TRUE : FALSE );
	cxti.lfont.lfStrikeOut = ( dwFontEnh & FONT_EnhStrikeOut ? TRUE : FALSE );
	cxti.lfont.lfCharSet = DEFAULT_CHARSET;
	cxti.lfont.lfOutPrecision = OUT_DEFAULT_PRECIS;
	cxti.lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	cxti.lfont.lfQuality = ( dwFontEnh & MIMG_FONT_ENH_CLEARTYPE ? 5 : PROOF_QUALITY );
	cxti.lfont.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
	lstrcpyn( cxti.lfont.lfFaceName, lpsFontName, LF_FACESIZE );

	cxti.fcolor = clr;
	cxti.bcolor = 0;
	cxti.opaque = FALSE;
	cxti.b_opacity = 0.0;
	cxti.b_outline = 0;
	cxti.b_round = 0;

	long lMaxTextLen = sizeof( cxti.text ) / sizeof( TCHAR );
	lstrcpyn( cxti.text, lpsText, lMaxTextLen );

	// Position oben anpassen
	if ( dwAlignFlags & MIMG_ALIGN_VCENTER || dwAlignFlags & MIMG_ALIGN_BOTTOM )
	{
		HFONT hFont = CreateFontIndirect( &cxti.lfont );
		if ( hFont )
		{
			HFONT hOldFont = (HFONT)SelectObject( hDC, hFont );

			long lLen = min( lstrlen( lpsText ), lMaxTextLen );
			RECT r = {0,0,0,0};
			int iRet = DrawText( hDC, lpsText, lLen, &r, DT_NOPREFIX | DT_CALCRECT );
			if ( iRet )
			{
				if ( dwAlignFlags & MIMG_ALIGN_VCENTER || dwAlignFlags & MIMG_ALIGN_BOTTOM )
				{
					long lTxtHeight = r.bottom - r.top;
					long lRectHeight = pRect->bottom - pRect->top;
					long lTop;
					if ( dwAlignFlags & MIMG_ALIGN_VCENTER )
						lTop = pRect->top + ( ( lRectHeight - lTxtHeight ) / 2 );
					else if ( dwAlignFlags & MIMG_ALIGN_BOTTOM )
						lTop = pRect->bottom - lTxtHeight;

					pRect->top = max( lTop, pRect->top );
				}
			}
			
			if ( hOldFont )
				SelectObject( hDC, hOldFont );
			DeleteObject( hFont );
		}
	}
	
	// Position links anpassen
	if ( dwAlignFlags & MIMG_ALIGN_HCENTER || dwAlignFlags & MIMG_ALIGN_RIGHT )
	{
		long lRectWidth = pRect->right - pRect->left;
		long lLeft;
		if ( dwAlignFlags & MIMG_ALIGN_HCENTER )
			lLeft = pRect->left + (lRectWidth / 2); // CxImage erwartet bei zentrierter Ausgabe, dass links = Mitte ist
		else if ( dwAlignFlags & MIMG_ALIGN_RIGHT )
			lLeft = pRect->right; // CxImage erwartet bei rechtsbündiger Ausgabe, dass links = rechts ist

		pRect->left = max( lLeft, pRect->left );
	}

	if ( dwAlignFlags & MIMG_ALIGN_HCENTER )
		cxti.align = DT_CENTER;
	else if ( dwAlignFlags & MIMG_ALIGN_RIGHT )
		cxti.align = DT_RIGHT;
	else	
		cxti.align = DT_LEFT;
	

	m_pImage->DrawStringEx( hDC, pRect->left, pRect->top, &cxti, FALSE );

	ReleaseDC( NULL, hDC );

	return TRUE;
}

// QueryFlags
BOOL CMImg::QueryFlags( DWORD dwFlags )
{
	return m_dwFlags & dwFlags;
}

// Resample
BOOL CMImg::Resample( int iNewX, int iNewY, int iMode )
{
	if ( !m_pImage )
		return FALSE;

	return m_pImage->Resample( iNewX, iNewY, iMode );
}

// Rotate
BOOL CMImg::Rotate( float fAngle )
{
	if ( !m_pImage )
		return FALSE;
	if ( fAngle == 0.0 )
		return TRUE;

	if ( fAngle == 90.0 )
		return m_pImage->RotateRight( );
	else if ( fAngle == -90.0 )
		return m_pImage->RotateLeft( );
	else if ( fAngle == 180.0 || fAngle == -180.0 )
		return m_pImage->Rotate180( );
	else
		return m_pImage->Rotate( fAngle );
}

// Save
BOOL CMImg::Save( LPTSTR lpsFile, DWORD dwType )
{
	if ( !( m_pImage && lpsFile ) )
		return FALSE;

	if ( dwType == MIMG_TYPE_UNKNOWN )
		dwType = m_pImage->GetType( );

	CMImg *pMImg = this;

	BOOL bGIF = ( dwType == MIMG_TYPE_GIF );
	BOOL bICO = ( dwType == MIMG_TYPE_ICO );
	BOOL bJPG = ( dwType == MIMG_TYPE_JPG );
	BOOL bPNG = ( dwType == MIMG_TYPE_PNG );
	BOOL bTGA = ( dwType == MIMG_TYPE_TGA );
	BOOL bTIF = ( dwType == MIMG_TYPE_TIF );
	BOOL bWBMP = ( dwType == MIMG_TYPE_WBMP );

	if ( bJPG && GetBpp( ) < 24 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 24 ) )
			return FALSE;
	}
	else if ( bPNG && m_pImage->AlphaIsValid( ) && GetBpp( ) < 24 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 24 ) )
			return FALSE;
	}
	else if ( bTGA && GetBpp( ) < 8 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 8 ) )
			return FALSE;
	}
	else if ( bWBMP && GetBpp( ) > 1 )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		if ( !pMImg->SetBpp( 1 ) )
			return FALSE;
	}
	else if ( bICO )
	{
		if ( !( pMImg = CreateCopy( ) ) ) return FALSE;
		CxImage *pImage = pMImg->m_pImage;

		BOOL bAlpha = pImage->AlphaIsValid( );
		BOOL bAlphaPalette = pImage->AlphaPaletteIsValid( );
		DWORD dwTranspColor = pMImg->GetTranspColor( );

		if ( bAlpha || bAlphaPalette || dwTranspColor != MIMG_COLOR_UNDEF )
		{
			if ( !bAlpha )
				pImage->AlphaCreate( );

			pMImg->SetTranspColor( MIMG_COLOR_UNDEF );

			BOOL bBlackNotInPalette = pMImg->HasPalette( ) && pMImg->GetColorIndex( 0 ) < 0;
			if ( bBlackNotInPalette || bAlpha || bAlphaPalette )
				pMImg->SetBpp( 24 );

			// Farbe der "transparenten Pixel" auf Schwarz und Alpha = 0 setzen
			DWORD dwColor;
			DWORD dwHt = pImage->GetHeight( );
			DWORD dwWid = pImage->GetWidth( );
			DWORD dwX, dwY;

			for ( dwY = 0; dwY < dwHt; ++dwY )
			{
				for ( dwX = 0; dwX < dwWid; ++dwX )
				{
					dwColor = pImage->RGBQUADtoRGB( pImage->GetPixelColor( dwX, dwY, FALSE ) );
					if (
						( dwColor == dwTranspColor ) ||
						( bAlpha && pImage->AlphaGet( dwX, dwY ) == 0 ) ||
						( bAlphaPalette && pImage->GetPixelColor( dwX, dwY, TRUE ).rgbReserved == 0 )
					   )
					{
						pImage->SetPixelColor( dwX, dwY, 0 );
						pImage->AlphaSet( dwX, dwY, 0 );
					}
				}
			}
		}
	}

	BYTE byJPGQuality;
	int iEncodeOption;
	if ( bJPG )
	{
		byJPGQuality = pMImg->m_pImage->GetJpegQuality( );
		pMImg->m_pImage->SetJpegQuality( pMImg->GetJPGQuality( ) );
	}
	else if ( bGIF || bTIF )
	{
		iEncodeOption = pMImg->m_pImage->GetCodecOption( );
		if ( bGIF )
			pMImg->m_pImage->SetCodecOption( GetGIFCompr( ) );
		else
		{
			WORD wBpp = pMImg->GetBpp( );
			int iCompr = pMImg->GetTIFCompr( );
			if ( ( iCompr == MIMG_COMPR_TIF_CCITTRLE || iCompr == MIMG_COMPR_TIF_CCITTFAX3 || iCompr == MIMG_COMPR_TIF_CCITTFAX4 ) && wBpp != 1 )
				return FALSE;
			else if ( ( iCompr == MIMG_COMPR_TIF_NONE || iCompr == MIMG_COMPR_TIF_LZW || iCompr == MIMG_COMPR_TIF_PACKBITS ) && wBpp == 24 )
				return FALSE;
			else if ( iCompr == MIMG_COMPR_TIF_JPEG && wBpp != 24 )
				return FALSE;

			pMImg->m_pImage->SetCodecOption( iCompr );
		}
	}

	BOOL bOk = pMImg->m_pImage->Save( lpsFile, dwType );

	if ( pMImg != this )
	{
		if ( bJPG )
		{
			pMImg->m_pImage->SetJpegQuality( byJPGQuality );
		}
		else if ( bGIF || bTIF )
		{
			pMImg->m_pImage->SetCodecOption( iEncodeOption );
		}

		delete pMImg;
	}

	return bOk;
}

// SetBpp
BOOL CMImg::SetBpp( WORD wBpp )
{
	if ( !m_pImage )
		return FALSE;

	BOOL bOk = FALSE;
	WORD wCurBpp = GetBpp( );

	if ( wCurBpp == wBpp )
		return TRUE;
	else if ( wCurBpp < wBpp )
		bOk = m_pImage->IncreaseBpp( wBpp );
	else
	{
		int iTransIndex = m_pImage->GetTransIndex( );
		DWORD dwTransColor = m_pImage->RGBQUADtoRGB( m_pImage->GetTransColor( ) );

		RGBQUAD *pPal = NULL;
		if ( wBpp < 24 )
		{
			int iColors = 0;
			if ( wBpp == 8 )
				iColors = 256;
			else if ( wBpp == 4 )
				iColors = 16;
			else if ( wBpp == 1 )
				iColors = 2;

			vector<COLORREF> vColors;
			int iUsedColors = GetUsedColors( vColors );
			if ( iUsedColors > 0 && iUsedColors <= iColors )
			{
				pPal = (RGBQUAD*) malloc( iColors * sizeof( RGBQUAD ) );
				memset( pPal, 0, iColors * sizeof( RGBQUAD ) );

				for ( int i = 0; i < iUsedColors; ++i )
				{
					pPal[i] = m_pImage->RGBtoRGBQUAD( vColors[i] );
				}
			}
		}

		if ( ( bOk = m_pImage->DecreaseBpp( wBpp, FALSE, pPal ) ) && iTransIndex >= 0 )
		{
			int iTransIndexNew = GetColorIndex( dwTransColor );
			if ( iTransIndexNew != iTransIndex  )
				m_pImage->SetTransIndex( iTransIndexNew );
		}

		if ( pPal )
			free( pPal );
	}

	return bOk;
}

// SetColorOpacity
BOOL CMImg::SetColorOpacity( DWORD dwTranspColor, BYTE byVal )
{
	if ( !m_pImage )
		return FALSE;

	if ( !CreateAlpha( ) )
		return FALSE;

	DWORD dwWid = m_pImage->GetWidth( );
	DWORD dwHt = m_pImage->GetHeight( );
	RGBQUAD r;
	for ( DWORD dwY = 0; dwY < dwHt; dwY++ )
	{
		for ( DWORD dwX = 0; dwX < dwWid; dwX++ )
		{
			r = m_pImage->GetPixelColor( dwX, dwY );
			if ( RGB( r.rgbRed, r.rgbGreen, r.rgbBlue ) == dwTranspColor )
				m_pImage->AlphaSet( dwX, dwY, byVal );
		}
	}

	return TRUE;
}

// SetDrawClipRect
BOOL CMImg::SetDrawClipRect( LPRECT pRect )
{
	if ( !pRect )
		return FALSE;

	return CopyRect( &m_rDrawClipRect, pRect );
}


// SetDrawState
BOOL CMImg::SetDrawState( DWORD dwState )
{
	m_dwDrawState = dwState;
	return TRUE;
}

// SetFlags
BOOL CMImg::SetFlags( DWORD dwFlags, BOOL bSet )
{
	m_dwFlags = ( bSet ? ( m_dwFlags | dwFlags ) : ( ( m_dwFlags & dwFlags ) ^ m_dwFlags ) );
	return TRUE;
}

// SetGIFCompr
BOOL CMImg::SetGIFCompr( int iVal )
{
	m_iGIFCompr = iVal;
	return TRUE;
}

// SetJPGQuality
BOOL CMImg::SetJPGQuality( BYTE byVal )
{
	m_byJPGQuality = max( min( byVal, 100 ), 1 );
	return TRUE;
}

// SetMaxOpacity
BOOL CMImg::SetMaxOpacity( BYTE byVal )
{
	if ( !m_pImage )
		return FALSE;

	// Alpha-Daten initialisieren
	m_pImage->AlphaPaletteEnable( FALSE );
	if ( m_pImage->AlphaPaletteIsValid( ) )
	{
		if ( !CreateAlphaFromPalette( ) )
			return FALSE;
		m_pImage->AlphaPaletteClear( );
	}
	else
	{
		if ( !m_pImage->AlphaIsValid( ) )
		{
			m_pImage->AlphaCreate( );
			m_pImage->AlphaSet( 255 );
		}
	}

	// Setzen
	m_pImage->AlphaSetMax( byVal );

	return TRUE;
}

// SetOpacity
BOOL CMImg::SetOpacity( BYTE byFrom, BYTE byTo, BYTE byVal )
{
	if ( !m_pImage )
		return FALSE;

	if ( byFrom > byTo )
		return FALSE;

	if ( !CreateAlpha( ) )
		return FALSE;

	if ( byFrom == 0 && byTo == 255 )
	{
		m_pImage->AlphaSet( byVal );
	}
	else
	{
		BYTE by;
		int iWid = int( m_pImage->GetWidth( ) );
		int iHt = int( m_pImage->GetHeight( ) );
		int iX, iY;
		for ( iY = 0; iY < iHt; iY++ )
		{
			for ( iX = 0; iX < iWid; iX++ )
			{
				by = m_pImage->AlphaGet( iX, iY );
				if ( by >= byFrom && by <= byTo )
					m_pImage->AlphaSet( iX, iY, byVal );
			}
		}
	}

	return TRUE;
}

// SetPaletteColor
BOOL CMImg::SetPaletteColor( BYTE byIdx, DWORD dwColor )
{
	if ( !( m_pImage && HasPalette( ) && byIdx < m_pImage->GetNumColors( ) ) )
		return FALSE;

	m_pImage->SetPaletteColor( byIdx, dwColor );

	return TRUE;
}

// SetPixelColor
BOOL CMImg::SetPixelColor( int iX, int iY, DWORD dwColor )
{
	if ( !( m_pImage && m_pImage->IsInside( iX, iY ) ) )
		return FALSE;

	TransformXY( iX, iY );
	m_pImage->SetPixelColor( iX, iY, dwColor );

	return TRUE;
}

// SetPixelOpacity
BOOL CMImg::SetPixelOpacity( int iX, int iY, BYTE byVal )
{
	if ( !m_pImage )
		return FALSE;

	if ( !CreateAlpha( ) )
		return FALSE;

	TransformXY( iX, iY );
	if ( !m_pImage->IsInside( iX, iY ) )
		return FALSE;
	m_pImage->AlphaSet( iX, iY, byVal );

	return TRUE;
}

// SetSize
BOOL CMImg::SetSize( int iNewWid, int iNewHt, DWORD dwOverflowColor, BYTE byOverflowOpacity )
{
	if ( !m_pImage )
		return FALSE;

	if ( iNewWid <= 0 || iNewHt <= 0 )
		return FALSE;

	BYTE byAlpha;
	int iWid = m_pImage->GetWidth( );
	int iHt = m_pImage->GetHeight( );
	int iX, iNewX, iY, iNewY;
	RGBQUAD rgb, rgbOverflow = m_pImage->RGBtoRGBQUAD( dwOverflowColor );

	CxImage *pImg = new CxImage( *m_pImage, FALSE, FALSE, FALSE );
	pImg->Create( iNewWid, iNewHt, m_pImage->GetBpp( ), m_pImage->GetType( ) );

	pImg->SetPalette( m_pImage->GetPalette( ) );

	BOOL bAlpha = m_pImage->AlphaIsValid( ) || ( ( iNewHt > iHt || iNewWid > iWid ) && byOverflowOpacity != 255 );
	if ( bAlpha )
	{
		pImg->AlphaCreate( );
		pImg->AlphaSet( 255 );
	}

	for ( iNewY = 0, iY = iHt - iNewHt; iNewY < iNewHt; ++iNewY, ++iY )
	{
		for ( iNewX = iX = 0; iNewX < iNewWid; ++iNewX, ++iX )
		{
			if ( iY < 0 || iX >= iWid )
			{
				rgb = rgbOverflow;
				if ( bAlpha )
					byAlpha = byOverflowOpacity;
			}
			else
			{
				rgb = m_pImage->GetPixelColor( iX, iY );
				if ( bAlpha && m_pImage->AlphaIsValid( ) )
					byAlpha = m_pImage->AlphaGet( iX, iY );
				else
					byAlpha = 255;
			}

			pImg->SetPixelColor( iNewX, iNewY, rgb );
			if ( bAlpha )
				pImg->AlphaSet( iNewX, iNewY, byAlpha );
		}
	}

	FreeImageData( );
	m_pImage = pImg;

	return TRUE;
}

// SetTIFCompr
BOOL CMImg::SetTIFCompr( int iVal )
{
	m_iTIFCompr = iVal;
	return TRUE;
}

// SetTranspColor
BOOL CMImg::SetTranspColor( DWORD dwColor )
{
	if ( !m_pImage )
		return FALSE;

	if ( dwColor == MIMG_COLOR_UNDEF )
	{
		m_pImage->SetTransIndex( -1 );
	}
	else
	{
		if ( HasPalette( ) )
		{
			int iIdx = GetColorIndex( dwColor );
			if ( iIdx < 0 )
				return FALSE;
			m_pImage->SetTransIndex( iIdx );
			m_pImage->SetTransColor( m_pImage->GetPaletteColor( iIdx ) );
		}
		else
		{
			m_pImage->SetTransIndex( 0 );
			m_pImage->SetTransColor( m_pImage->RGBtoRGBQUAD( dwColor ) );
		}
	}

	return TRUE;
}

// SetWMFLoadSize
BOOL CMImg::SetWMFLoadSize( SIZE &si )
{
	m_siWMF.cx = max( 0, si.cx );
	m_siWMF.cy = max( 0, si.cy );
	return TRUE;
}

// StripOpacity
BOOL CMImg::StripOpacity( DWORD dwColor )
{
	if ( !m_pImage )
		return FALSE;

	if ( dwColor == MIMG_COLOR_UNDEF )
		return FALSE;

	if ( GetBpp( ) < 24 )
	{
		int iIdx = GetColorIndex( dwColor );
		if ( iIdx < 0 )
		{
			if ( !SetBpp( 24 ) )
				return FALSE;
		}
	}

	if ( !CreateAlpha( ) )
		return FALSE;

	if ( !AlphaPaletteToAlpha( ) )
		return FALSE;

	if ( !TranspColorToAlpha( ) )
		return FALSE;

	if ( !SetTranspColor( dwColor ) )
		return FALSE;

	m_pImage->AlphaStrip( );

	SetTranspColor( MIMG_COLOR_UNDEF );	

	return TRUE;
}

// SubstColor
BOOL CMImg::SubstColor( DWORD dwOldColor, DWORD dwNewColor )
{
	if ( !m_pImage )
		return FALSE;

	BOOL bAll = ( dwOldColor == MIMG_COLOR_UNDEF );

	// Evtl. Farbtransparenz in Alpha-Transparenz konvertieren bzw. transparente Farbe anpassen
	DWORD dwTranspColor = GetTranspColor( );
	if ( dwTranspColor != MIMG_COLOR_UNDEF )
	{
		if ( bAll || dwOldColor == dwTranspColor )
		{
			TranspColorToAlpha( );
			SetTranspColor( MIMG_COLOR_UNDEF );
		}
	}
	
	BOOL bSubst;
	DWORD dwHt = m_pImage->GetHeight( );
	DWORD dwWid = m_pImage->GetWidth( );
	RGBQUAD r;
	for ( DWORD dwY = 0; dwY < dwHt; dwY++ )
	{
		for ( DWORD dwX = 0; dwX < dwWid; dwX++ )
		{
			if ( bAll )
				bSubst = TRUE;
			else
			{
				r = m_pImage->GetPixelColor( dwX, dwY );
				bSubst = ( dwOldColor == m_pImage->RGBQUADtoRGB( r ) );
			}
			if ( bSubst )
				m_pImage->SetPixelColor( dwX, dwY, dwNewColor );
		}
	}	

	return TRUE;
}

// TransformXY
void CMImg::TransformXY( int &iX, int &iY )
{
	if ( m_pImage )
		iY = m_pImage->GetHeight( ) - iY - 1;
}

// TranspColorToAlpha
BOOL CMImg::TranspColorToAlpha( )
{
	if ( !m_pImage )
		return FALSE;

	if ( !CreateAlpha( ) )
		return FALSE;

	DWORD dwTranspColor = GetTranspColor( );
	if ( dwTranspColor != MIMG_COLOR_UNDEF )
	{
		DWORD dwWid = m_pImage->GetWidth( );
		DWORD dwHt = m_pImage->GetHeight( );
		RGBQUAD r;
		for ( DWORD dwY = 0; dwY < dwHt; dwY++ )
		{
			for ( DWORD dwX = 0; dwX < dwWid; dwX++ )
			{
				r = m_pImage->GetPixelColor( dwX, dwY );
				if ( RGB( r.rgbRed, r.rgbGreen, r.rgbBlue ) == dwTranspColor )
					m_pImage->AlphaSet( dwX, dwY, 0 );
			}
		}
	}

	return TRUE;
}
