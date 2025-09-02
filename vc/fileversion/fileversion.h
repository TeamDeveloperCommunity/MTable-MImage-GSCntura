// fileversion.h

#ifndef _INC_FILEVERSION
#define _INC_FILEVERSION

// Pragmas
#pragma warning(disable:4786)

// Includes
#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif
#include <winver.h>
#include "tstring.h"

// Using
using namespace std;

// Funktionsdeklarationen
BOOL FileVerGetVersionStr( LPTSTR lpsFile, tstring & sVer );

#endif // _INC_FILEVERSION