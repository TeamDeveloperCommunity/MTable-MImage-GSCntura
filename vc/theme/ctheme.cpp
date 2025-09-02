// ctheme.cpp

#include "ctheme.h"

// Klasse CTheme

// Konstruktor
CTheme::CTheme( )
{
	// Initialisierung
	m_bCanUse = FALSE;
	m_hUxThemeDLL = NULL;
	
	// Funktionsvariablen initialisieren
	CloseThemeData = NULL;
	DrawThemeBackground = NULL;
	GetThemePartSize = NULL;
	GetThemeSysColor = NULL;
	IsAppThemed = NULL;
	OpenThemeData = NULL;

	if ( WinVerIsXP_OG( ) )
	{
		// Funktionsaddressen ermitteln
		if ( m_hUxThemeDLL = LoadLibrary( _T("uxtheme.dll") ) )
		{
			if ( !( CloseThemeData = (CLOSE_THEME_DATA)GetProcAddress( m_hUxThemeDLL, "CloseThemeData" ) ) )
				return;
			if ( !( DrawThemeBackground = (DRAW_THEME_BACKGROUND)GetProcAddress( m_hUxThemeDLL, "DrawThemeBackground" ) ) )
				return;
			if ( !( IsAppThemed = (IS_APP_THEMED)GetProcAddress( m_hUxThemeDLL, "IsAppThemed" ) ) )
				return;
			if ( !( GetThemeColor = (GET_THEME_COLOR)GetProcAddress( m_hUxThemeDLL, "GetThemeColor" ) ) )
				return;
			if ( !( GetThemePartSize = (GET_THEME_PART_SIZE)GetProcAddress( m_hUxThemeDLL, "GetThemePartSize" ) ) )
				return;
			if ( !( GetThemeSysColor = (GET_THEME_SYS_COLOR)GetProcAddress( m_hUxThemeDLL, "GetThemeSysColor" ) ) )
				return;
			if ( !( OpenThemeData = (OPEN_THEME_DATA)GetProcAddress( m_hUxThemeDLL, "OpenThemeData" ) ) )
				return;
		}

		// OK
		m_bCanUse = TRUE;
	}
}

// Destruktor
CTheme::~CTheme( )
{
	if ( m_hUxThemeDLL )
		FreeLibrary( m_hUxThemeDLL );
}

// Funktionen

// theTheme
CTheme & theTheme( )
{
	static CTheme th;
	return th;
}