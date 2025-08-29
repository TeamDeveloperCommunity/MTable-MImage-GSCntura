// winversion.cpp

// Includes
#include "winversion.h"

// WinVerGet
// Ermittelt die Windows-Version
int WinVerGet( )
{
	OSVERSIONINFO ovi;

	ovi.dwOSVersionInfoSize = sizeof( OSVERSIONINFO );
	if ( GetVersionEx( &ovi ) )
	{
		switch( ovi.dwPlatformId )
		{
		case VER_PLATFORM_WIN32_NT:
			if ( ovi.dwMajorVersion == 3 )
				return WINVER_NT_3;
			else if ( ovi.dwMajorVersion == 4 )
				return WINVER_NT_4;
			else if ( ovi.dwMajorVersion == 5 && ovi.dwMinorVersion == 0 )
				return WINVER_2000;
			else if ( ovi.dwMajorVersion == 5 && ovi.dwMinorVersion == 1 )
				return WINVER_XP;
			break;
		case VER_PLATFORM_WIN32_WINDOWS:
			if ( ovi.dwMajorVersion == 4 && ovi.dwMinorVersion == 0 )
				return WINVER_95;
			else if ( ovi.dwMajorVersion == 4 && ovi.dwMinorVersion == 10 )
				return WINVER_98;
			else if ( ovi.dwMajorVersion == 4 && ovi.dwMinorVersion == 90 )
				return WINVER_ME;
			break;
		case VER_PLATFORM_WIN32s:
			return WINVER_31;
		}
	}

	return WINVER_UNKNOWN;
}

// WinVerIs2000_OG
// Ermittelt, ob die Windows-Version >= 2000 ist
BOOL WinVerIs2000_OG( )
{
	OSVERSIONINFO ovi;

	ovi.dwOSVersionInfoSize = sizeof( OSVERSIONINFO );
	if ( GetVersionEx( &ovi ) )
	{
		return ( ovi.dwPlatformId == VER_PLATFORM_WIN32_NT && ovi.dwMajorVersion >= 5 );
	}

	return FALSE;
}

// WinVerIs98_OG
// Ermittelt, ob die Windows-Version >= 98 ist
BOOL WinVerIs98_OG( )
{
	OSVERSIONINFO ovi;

	ovi.dwOSVersionInfoSize = sizeof( OSVERSIONINFO );
	if ( GetVersionEx( &ovi ) )
	{
		return ( ovi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS &&
			   ( ovi.dwMajorVersion > 4 ||
			   ( ovi.dwMajorVersion == 4 && ovi.dwMinorVersion >= 10 ) ) );
	}

	return FALSE;
}

// WinVerIsNT4_OG
// Ermittelt, ob die Windows-Version >= NT4 ist
BOOL WinVerIsNT4_OG( )
{
	OSVERSIONINFO ovi;

	ovi.dwOSVersionInfoSize = sizeof( OSVERSIONINFO );
	if ( GetVersionEx( &ovi ) )
	{
		return ( ovi.dwPlatformId == VER_PLATFORM_WIN32_NT && ovi.dwMajorVersion >= 4 );
	}

	return FALSE;
}

// WinVerIsXP_OG
// Ermittelt, ob die Windows-Version >= XP ist
BOOL WinVerIsXP_OG( )
{
	OSVERSIONINFO ovi;

	ovi.dwOSVersionInfoSize = sizeof( OSVERSIONINFO );
	if ( GetVersionEx( &ovi ) )
	{
		return ( ovi.dwPlatformId == VER_PLATFORM_WIN32_NT &&
			   ( ovi.dwMajorVersion > 5 ||
			   ( ovi.dwMajorVersion == 5 && ovi.dwMinorVersion >= 1 ) ) );
	}

	return FALSE;
}