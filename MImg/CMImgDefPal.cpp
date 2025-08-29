// CMImgDefPal.cpp

// Includes
#include "CMImg.h"

// Klasse CMImgDefPal

// Konstruktoren
CMImgDefPal::CMImgDefPal( )
{
	m_hPal = NULL;
	CreatePalette( );
}

// Destruktor
CMImgDefPal::~CMImgDefPal( )
{
	DeletePalette( );
}

// Methoden
void CMImgDefPal::CreatePalette( )
{
	DeletePalette( );

	HPALETTE hPalStock = (HPALETTE)GetStockObject( DEFAULT_PALETTE );
	UINT uiEntries = GetPaletteEntries( hPalStock, 0, 0, NULL );
	if ( uiEntries )
	{
		size_t sizeEntries = uiEntries * sizeof( PALETTEENTRY );
		LPLOGPALETTE pPal = (LPLOGPALETTE)malloc( sizeof( LOGPALETTE ) - sizeof( PALETTEENTRY ) + sizeEntries );
		if ( pPal )
		{
			pPal->palVersion = 0x300;
			pPal->palNumEntries = uiEntries;
			GetPaletteEntries( hPalStock, 0, uiEntries, pPal->palPalEntry );

			m_hPal = ::CreatePalette( pPal );

			free( pPal );
		}
	}
}

void CMImgDefPal::DeletePalette( )
{
	if ( m_hPal )
	{
		DeleteObject( m_hPal );
		m_hPal = NULL;
	}
}

HPALETTE CMImgDefPal::GetHandle( )
{
	return m_hPal;
}

// Funktionen
HPALETTE theMImgDefPal( )
{
	static CMImgDefPal thePal;
	return thePal.GetHandle( );
}
