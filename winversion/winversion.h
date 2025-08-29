// winversion.h

#ifndef _INC_WINVERSION
#define _INC_WINVERSION

// Includes
#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif

// Konstanten
#define WINVER_UNKNOWN 0
#define WINVER_NT_3 1
#define WINVER_NT_4 2
#define WINVER_2000 3
#define WINVER_XP 4
#define WINVER_31 5
#define WINVER_95 6
#define WINVER_98 7
#define WINVER_ME 8

// Funktionsdeklarationen
int WinVerGet( );
BOOL WinVerIs2000_OG( );
BOOL WinVerIs98_OG( );
BOOL WinVerIsNT4_OG( );
BOOL WinVerIsXP_OG( );

#endif // _INC_WINVERSION