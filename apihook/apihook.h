// apihook.h

#ifndef _APIHOOK_H
#define _APIHOOK_H

/*//////////////////////////////////////////////////////////////////////
                                Includes
//////////////////////////////////////////////////////////////////////*/
#include <windows.h>
#include <tchar.h>

/*//////////////////////////////////////////////////////////////////////
                            Structures
//////////////////////////////////////////////////////////////////////*/
typedef struct tag_HOOKFUNCDESCA
{
    char szFunc[MAX_PATH];	// The name of the function to hook
    PROC pProc;		// The procedure to blast in
} HOOKFUNCDESCA , * LPHOOKFUNCDESCA ;

typedef struct tag_HOOKFUNCDESCW
{
    WCHAR szFunc[MAX_PATH];	// The name of the function to hook
    PROC pProc;		// The procedure to blast in
} HOOKFUNCDESCW , * LPHOOKFUNCDESCW ;

#ifdef UNICODE
#define HOOKFUNCDESC   HOOKFUNCDESCW
#define LPHOOKFUNCDESC LPHOOKFUNCDESCW
#else
#define HOOKFUNCDESC   HOOKFUNCDESCA
#define LPHOOKFUNCDESC LPHOOKFUNCDESCA
#endif  // UNICODE

/*//////////////////////////////////////////////////////////////////////
                            Functions
//////////////////////////////////////////////////////////////////////*/
BOOL HookImportedFunctionsByName( HMODULE hModule, LPCSTR szImportMod, UINT uiCount, LPHOOKFUNCDESCA paHookArray, PROC *paOrigFuncs,LPDWORD pdwHooked );
BOOL GetLoadedModules( DWORD dwPID, UINT uiCount, HMODULE * paModArray, LPDWORD pdwRealCount );
DWORD GetModuleBaseNamePortable( HANDLE hProcess, HMODULE hModule, LPTSTR lpBaseName, DWORD nSize );

#endif  // _APIHOOK_H
