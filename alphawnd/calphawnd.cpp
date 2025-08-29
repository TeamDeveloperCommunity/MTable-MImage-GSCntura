// calphawnd.cpp

#include "calphawnd.h"

// Klasse CAlphaWnd

// Konstruktor
CAlphaWnd::CAlphaWnd( )
{
	// Initialisierung
	m_bCanUse = FALSE;
	m_hUserDLL = NULL;

	SetLayeredWindowAttributes = NULL;

	if ( WinVerIs2000_OG( ) )
	{
		// Funktionsaddressen ermitteln
		if ( m_hUserDLL = LoadLibrary( _T("user32.dll") ) )
		{
			if ( !( SetLayeredWindowAttributes = (PSLWA)GetProcAddress( m_hUserDLL, "SetLayeredWindowAttributes" ) ) )
				return;
			if ( !( AnimateWindow = (PAW)GetProcAddress( m_hUserDLL, "AnimateWindow" ) ) )
				return;
		}

		// OK
		m_bCanUse = TRUE;		
	}
}

// Destruktor
CAlphaWnd::~CAlphaWnd( )
{
	if ( m_hUserDLL )
		FreeLibrary( m_hUserDLL );
}

// Animate
BOOL CAlphaWnd::Animate( HWND hWnd, DWORD dwTime, DWORD dwFlags )
{
	if ( !CanUse( ) )
		return FALSE;

	return AnimateWindow( hWnd, dwTime, dwFlags );
}

// SetAlpha
BOOL CAlphaWnd::SetAlpha( HWND hWnd, BYTE byAlpha )
{
	if ( !CanUse( ) )
		return FALSE;

	if ( byAlpha == 255 )
		return SetLayeredStyle( hWnd, FALSE );
	else
	{
		SetLayeredStyle( hWnd, TRUE );
		return SetLayeredWindowAttributes( hWnd, 0, byAlpha, LWA_ALPHA );
	}
}

// SetLayeredStyle
BOOL CAlphaWnd::SetLayeredStyle( HWND hWnd, BOOL bSet )
{
	if ( !CanUse( ) )
		return FALSE;

	if ( !IsWindow( hWnd ) )
		return FALSE;

	long lLong = GetWindowLongPtr(hWnd, GWL_EXSTYLE);
	if ( bSet )
	{
		lLong |= WS_EX_LAYERED;
	}
	else
	{
		lLong = lLong & ~WS_EX_LAYERED;
	}

	SetLastError( 0 );
	SetWindowLongPtr( hWnd, GWL_EXSTYLE, lLong );

	return ( !GetLastError( ) );
}


// Funktionen

// theAlphaWnd
CAlphaWnd & theAlphaWnd( )
{
	static CAlphaWnd aw;
	return aw;
}