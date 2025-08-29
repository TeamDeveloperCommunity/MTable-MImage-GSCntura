// apihook.cpp

/*//////////////////////////////////////////////////////////////////////
                                Includes
//////////////////////////////////////////////////////////////////////*/
#include "apihook.h"
#include "TLHELP32.h"

/*//////////////////////////////////////////////////////////////////////
                                Typedefs
//////////////////////////////////////////////////////////////////////*/
// The typedefs for the PSAPI.DLL functions used by this module.
typedef BOOL (WINAPI *ENUMPROCESSMODULES) ( HANDLE    hProcess   ,
                                            HMODULE * lphModule  ,
                                            DWORD     cb         ,
                                            LPDWORD   lpcbNeeded  ) ;

typedef DWORD (WINAPI *GETMODULEBASENAME) ( HANDLE  hProcess   ,
                                            HMODULE hModule    ,
                                            LPTSTR  lpBaseName ,
                                            DWORD   nSize       ) ;

// The typedefs for the TOOLHELP32.DLL functions used by this module.
// Type definitions for pointers to call tool help functions.
typedef BOOL (WINAPI *MODULEWALK) ( HANDLE          hSnapshot ,
                                    LPMODULEENTRY32 lpme       ) ;
typedef BOOL (WINAPI *THREADWALK) ( HANDLE          hSnapshot ,
                                    LPTHREADENTRY32 lpte       ) ;
typedef BOOL (WINAPI *PROCESSWALK) ( HANDLE           hSnapshot ,
                                     LPPROCESSENTRY32 lppe       ) ;
typedef HANDLE (WINAPI *CREATESNAPSHOT) ( DWORD dwFlags       ,
                                          DWORD th32ProcessID  ) ;

/*//////////////////////////////////////////////////////////////////////
                            File Static Data
//////////////////////////////////////////////////////////////////////*/
// Has the function stuff here been initialized? This is only to be
// used by the InitPSAPI function and nothing else.
static BOOL g_bInitialized_PSAPI = FALSE ;
// The pointer to EnumProcessModules.
static ENUMPROCESSMODULES g_pEnumProcessModules = NULL ;
// The pointer to GetModuleBaseName.
static GETMODULEBASENAME g_pGetModuleBaseName = NULL ;

// Has the function stuff here been initialized? This is only to be
// used by the InitTOOLHELP32 function and nothing else.
static BOOL           g_bInitialized_TLHELP       = FALSE ;
static CREATESNAPSHOT g_pCreateToolhelp32Snapshot = NULL  ;
static MODULEWALK     g_pModule32First            = NULL  ;
static MODULEWALK     g_pModule32Next             = NULL  ;
static PROCESSWALK    g_pProcess32First           = NULL  ;
static PROCESSWALK    g_pProcess32Next            = NULL  ;
static THREADWALK     g_pThread32First            = NULL  ;
static THREADWALK     g_pThread32Next             = NULL  ;

// This is only to be used by the IsNT function and nothing else.
// Indicates that the version information is valid.
static BOOL g_bHasVersion = FALSE ;
// Indicates NT or 9x.
static BOOL g_bIsNT = TRUE ;

/*//////////////////////////////////////////////////////////////////////
                            Macros
//////////////////////////////////////////////////////////////////////*/
#ifdef _WIN64
#define MakePtr( cast , ptr , AddValue ) (cast)( (ULONG64)(ptr)+(ULONG64)(AddValue))
#else
#define MakePtr( cast , ptr , AddValue ) (cast)( (DWORD)(ptr)+(DWORD)(AddValue))
#endif

/*//////////////////////////////////////////////////////////////////////
                            Forward Declarations
//////////////////////////////////////////////////////////////////////*/
BOOL GetLoadedModulesNT4( DWORD dwPID, UINT uiCount, HMODULE * paModArray, LPDWORD pdwRealCount );
BOOL GetLoadedModulesTLHELP( DWORD dwPID, UINT uiCount, HMODULE * paModArray, LPDWORD pdwRealCount );
DWORD GetModuleBaseNameNT( HANDLE hProcess, HMODULE hModule, LPTSTR lpBaseName, DWORD nSize );
DWORD GetModuleBaseName9x( HANDLE hProcess, HMODULE hModule, LPSTR lpBaseName, DWORD nSize );
PIMAGE_IMPORT_DESCRIPTOR GetNamedImportDescriptor( HMODULE hModule, LPCSTR szImportMod );
static BOOL InitPSAPI( void );
static BOOL InitTOOLHELP32( void );
BOOL IsNT ( void );

/*//////////////////////////////////////////////////////////////////////
                            Functions
//////////////////////////////////////////////////////////////////////*/
/*----------------------------------------------------------------------
FUNCTION        :   InitPSAPI
DISCUSSION      :
    Loads PSAPI.DLL and initializes all the pointers needed by this
file.  If BugslayerUtil.DLL statically linked to PSAPI.DLL, it would not
work on Windows95.
    Note that I conciously chose to allow the resource leak on loading
PSAPI.DLL.
PARAMETERS      :
    None.
RETURNS         :
    TRUE  - Everything initialized properly.
    FALSE - There was a problem.
----------------------------------------------------------------------*/
static BOOL InitPSAPI( void )
{
    if ( TRUE == g_bInitialized_PSAPI )
    {
        return ( TRUE ) ;
    }

    // Load up PSAPI.DLL.
    HINSTANCE hInst = LoadLibrary( _T("PSAPI.DLL") ) ;
    if ( NULL == hInst )
    {
        return ( FALSE ) ;
    }

    // Now do the GetProcAddress stuff.
    g_pEnumProcessModules = (ENUMPROCESSMODULES)GetProcAddress( hInst, "EnumProcessModules" );
    if ( NULL == g_pEnumProcessModules )
    {
        return ( FALSE ) ;
    }
#ifdef UNICODE
    g_pGetModuleBaseName = (GETMODULEBASENAME)GetProcAddress( hInst, "GetModuleBaseNameW" );
#else
	g_pGetModuleBaseName = (GETMODULEBASENAME)GetProcAddress( hInst, "GetModuleBaseNameA" );
#endif
    if ( NULL == g_pGetModuleBaseName )
    {
        return ( FALSE ) ;
    }

    // All OK, Jumpmaster!
    g_bInitialized_PSAPI = TRUE ;
    return ( TRUE ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   InitTOOLHELP32
DISCUSSION      :
    Retrieves the function pointers needed to access the tool help
routines.  Since this code is supposed to work on NT4, it cannot link
to the non-existant addresses in KERNEL32.DLL.
    This is pretty much lifter right from the MSDN.
PARAMETERS      :
    None.
RETURNS         :
    TRUE  - Everything initialized properly.
    FALSE - There was a problem.
----------------------------------------------------------------------*/
static BOOL InitTOOLHELP32( void )
{
    if ( TRUE == g_bInitialized_TLHELP )
    {
        return ( TRUE ) ;
    }

    BOOL      bRet    = FALSE ;
    HINSTANCE hKernel = NULL  ;

    // Obtain the module handle of the kernel to retrieve addresses of
    //  the tool helper functions.
    hKernel = GetModuleHandleA ( "KERNEL32.DLL" ) ;

    if ( NULL != hKernel )
    {
        g_pCreateToolhelp32Snapshot =
           (CREATESNAPSHOT)GetProcAddress ( hKernel ,
                                            "CreateToolhelp32Snapshot");
        g_pModule32First = (MODULEWALK)GetProcAddress (hKernel ,
                                                       "Module32First");
        g_pModule32Next = (MODULEWALK)GetProcAddress (hKernel        ,
                                                      "Module32Next"  );
        g_pProcess32First =
                (PROCESSWALK)GetProcAddress ( hKernel          ,
                                              "Process32First"  ) ;
        g_pProcess32Next =
                (PROCESSWALK)GetProcAddress ( hKernel         ,
                                              "Process32Next" ) ;
        g_pThread32First =
                (THREADWALK)GetProcAddress ( hKernel         ,
                                             "Thread32First"  ) ;
        g_pThread32Next =
                (THREADWALK)GetProcAddress ( hKernel        ,
                                             "Thread32Next"  ) ;
        // All addresses must be non-NULL to be successful.  If one of
        //  these addresses is NULL, one of the needed lists cannot be
        //  walked.

        bRet =  g_pModule32First            &&
                g_pModule32Next             &&
                g_pProcess32First           &&
                g_pProcess32Next            &&
                g_pThread32First            &&
                g_pThread32Next             &&
                g_pCreateToolhelp32Snapshot    ;
    }
    else
    {
        // Could not get the module handle of kernel.
        SetLastErrorEx ( ERROR_DLL_INIT_FAILED , SLE_ERROR ) ;
        bRet = FALSE ;
    }

    if ( TRUE == bRet )
    {
        // All OK, Jumpmaster!
        g_bInitialized_TLHELP = TRUE ;
    }
    return ( bRet ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   HookImportedFunctionsByName
DISCUSSION      :
    Hooks the specified functions imported into hModule by the module
indicated by szImportMod.  This function can be used to hook from one
to 'n' of the functions imported.
    The techniques used in the function are slightly different than
that shown by Matt Pietrek in his book, "Windows 95 System Programming
Secrets."  He uses the address of the function to hook as returned by
GetProcAddress.  Unfortunately, while this works in almost all cases, it
does not work when the program being hooked is running under a debugger
on Windows95 (and presumably, Windows98).  The problem is that
GetProcAddress under a debugger returns a "debug thunk," not the address
that is stored in the Import Address Table (IAT).
    This function gets around that by using the real thunk list in the
PE file, the one not bashed by the loader when the module is loaded and
fixed up, to find where the named import is located.  Once the named
import is found, then the original table is blasted to make the hook.
As the name implies, this function will only hook functions imported by
name.
PARAMETERS      :
    hModule     - The module where the imports will be hooked.
    szImportMod - The name of the module whose functions will be
                  imported.
    uiCount     - The number of functions to hook.  This is the size of
                  the paHookArray and paOrigFuncs arrays.
    paHookArray - The array of function descriptors that list which
                  functions to hook.  At this point, the array does not
                  have to be in szFunc name order.  Also, if a
                  particular pProc is NULL, then that item will just be
                  skipped.  This makes it much easier for debugging.
    paOrigFuncs - The array of original addresses that were hooked.  If
                  a function was not hooked, then that item will be
                  NULL.  This parameter can be NULL if the returned
                  information is not needed.
    pdwHooked   - Returns the number of functions hooked out of
                  paHookArray.  This parameter can be NULL if the
                  returned information is not needed.
RETURNS         :
    FALSE - There was a problem, check GetLastError.
    TRUE  - The function succeeded.  See the parameter discussion for
            the output parameters.
----------------------------------------------------------------------*/
BOOL HookImportedFunctionsByName( HMODULE hModule, LPCSTR szImportMod, UINT uiCount, LPHOOKFUNCDESCA paHookArray, PROC *paOrigFuncs, LPDWORD pdwHooked )
{
	
#ifdef _DEBUG
	// Assert the parameters.
	/*ASSERT(FALSE == IsBadReadPtr(hModule,
		sizeof(IMAGE_DOS_HEADER)));
	ASSERT(FALSE == IsBadStringPtr(szImportMod, MAX_PATH));
	ASSERT(0 != uiCount);
	ASSERT(NULL != paHookArray);
	ASSERT(FALSE == IsBadReadPtr(paHookArray,
		sizeof(HOOKFUNCDESC) * uiCount));

	if (NULL != paOrigFuncs)
	{
		ASSERT(FALSE == IsBadWritePtr(paOrigFuncs,
			sizeof(PROC) * uiCount));

	}
	if (NULL != pdwHooked)
	{
		ASSERT(FALSE == IsBadWritePtr(pdwHooked, sizeof(UINT)));
	}

	// Check each item in the hook array.
	{
		for (UINT i = 0; i < uiCount; i++)
		{
			ASSERT(NULL != paHookArray[i].szFunc);
			ASSERT('\0' != *paHookArray[i].szFunc);
			// If the function address isn't NULL, it is validated.
			if (NULL != paHookArray[i].pProc)
			{
				ASSERT(FALSE == IsBadCodePtr(paHookArray[i].pProc));
			}
		}
	}*/
#endif

	// Perform the error checking for the parameters.
	if ((0 == uiCount) ||
		(NULL == szImportMod) ||
		(TRUE == IsBadReadPtr(paHookArray,
		sizeof(HOOKFUNCDESC) * uiCount)))
	{
		SetLastErrorEx(ERROR_INVALID_PARAMETER, SLE_ERROR);
		return (FALSE);
	}
	if ((NULL != paOrigFuncs) &&
		(TRUE == IsBadWritePtr(paOrigFuncs,
		sizeof(PROC) * uiCount)))
	{
		SetLastErrorEx(ERROR_INVALID_PARAMETER, SLE_ERROR);
		return (FALSE);
	}
	if ((NULL != pdwHooked) &&
		(TRUE == IsBadWritePtr(pdwHooked, sizeof(UINT))))
	{
		SetLastErrorEx(ERROR_INVALID_PARAMETER, SLE_ERROR);
		return (FALSE);
	}

	// Is this a system DLL above the 2-GB line, which Windows 98 won't
	// let you patch?
	if ((FALSE == IsNT()) && ((DWORD)hModule >= 0x80000000))
	{
		SetLastErrorEx(ERROR_INVALID_HANDLE, SLE_ERROR);
		return (FALSE);
	}


	// TODO TODO
	// Should each item in the hook array be checked in release builds?

	if (NULL != paOrigFuncs)
	{
		// Set all the values in paOrigFuncs to NULL.
		memset(paOrigFuncs, NULL, sizeof(PROC) * uiCount);
	}
	if (NULL != pdwHooked)
	{
		// Set the number of functions hooked to 0.
		*pdwHooked = 0;
	}

	// Get the specific import descriptor.
	PIMAGE_IMPORT_DESCRIPTOR pImportDesc = GetNamedImportDescriptor(hModule, szImportMod);
	if (NULL == pImportDesc)
	{
		// The requested module wasn't imported. Don't return an error.
		return (TRUE);
	}

	// Get the original thunk information for this DLL. I can't use
	// the thunk information stored in pImportDesc->FirstThunk
	// because the loader has already changed that array to fix up
	// all the imports. The original thunk gives me access to the
	// function names.
	PIMAGE_THUNK_DATA pOrigThunk =
		MakePtr(PIMAGE_THUNK_DATA,
		hModule,
		pImportDesc->OriginalFirstThunk);
	// Get the array the pImportDesc->FirstThunk points to because
	// I'll do the actual bashing and hooking there.
	PIMAGE_THUNK_DATA pRealThunk = MakePtr(PIMAGE_THUNK_DATA,
		hModule,
		pImportDesc->FirstThunk);

	// Loop through and find the functions to hook.
	while (NULL != pOrigThunk->u1.Function)
	{
		// Look only at functions that are imported by name, not those
		// that are imported by ordinal value.
		if (IMAGE_ORDINAL_FLAG !=
			(pOrigThunk->u1.Ordinal & IMAGE_ORDINAL_FLAG))
		{
			// Look at the name of this imported function.
			PIMAGE_IMPORT_BY_NAME pByName;

			pByName = MakePtr(PIMAGE_IMPORT_BY_NAME,
				hModule,
				pOrigThunk->u1.AddressOfData);

			// If the name starts with NULL, skip it.
			if ('\0' == pByName->Name[0])
			{
				continue;
			}

			// Determines whether I hook the function
			BOOL bDoHook = FALSE;

			// TODO TODO
			// Might want to consider bsearch here.

			// See whether the imported function name is in the hook
			// array. Consider requiring paHookArray to be sorted by
			// function name so that bsearch can be used, which
			// will make the lookup faster. However, the size of
			// uiCount coming into this function should be rather
			// small, so it's OK to search the entire paHookArray for
			// each function imported by szImportMod.

			UINT i = 0;
			for (i = 0; i < uiCount; i++)
			{
				if ((paHookArray[i].szFunc[0] ==
					pByName->Name[0]) &&
					(0 == _strcmpi(paHookArray[i].szFunc,
					(char*)pByName->Name)))
				{
					// If the function address is NULL, exit now;
					// otherwise, go ahead and hook the function.
					if (NULL != paHookArray[i].pProc)
					{
						bDoHook = TRUE;
					}
					break;
				}
			}

			if (TRUE == bDoHook)
			{
				// I found a function to hook. Now I need to change
				// the memory protection to writable before I overwrite
				// the function pointer. Note that I'm now writing into
				// the real thunk area!

				MEMORY_BASIC_INFORMATION mbi_thunk;

				VirtualQuery(pRealThunk,
					&mbi_thunk,
					sizeof(MEMORY_BASIC_INFORMATION));

				if (FALSE == VirtualProtect(mbi_thunk.BaseAddress,
					mbi_thunk.RegionSize,
					PAGE_READWRITE,
					&mbi_thunk.Protect))
				{
#ifdef _DEBUG
					/*ASSERT(!"VirtualProtect failed!");*/
#endif
					SetLastErrorEx(ERROR_INVALID_HANDLE, SLE_ERROR);
					return (FALSE);
				}

				// Save the original address if requested.
				if (NULL != paOrigFuncs)
				{
					paOrigFuncs[i] = (PROC)pRealThunk->u1.Function;
				}

				// Microsoft has two different definitions of the
				// PIMAGE_THUNK_DATA fields as they are moving to
				// support Win64. The W2K RC2 Platform SDK is the
				// latest header, so I'll use that one and force the
				// Visual C++ 6 Service Pack 3 headers to deal with it.

				// Hook the function.
				/*DWORD * pTemp = (DWORD*)&pRealThunk->u1.Function;
				*pTemp = (DWORD)(paHookArray[i].pProc);*/
				pRealThunk->u1.Function = (uintptr_t)(paHookArray[i].pProc);

				DWORD dwOldProtect;

				// Change the protection back to what it was before I
				// overwrote the function pointer.
#ifdef _DEBUG
				/*VERIFY(VirtualProtect(mbi_thunk.BaseAddress,
					mbi_thunk.RegionSize,
					mbi_thunk.Protect,
					&dwOldProtect));*/
#else
				VirtualProtect(mbi_thunk.BaseAddress,
					mbi_thunk.RegionSize,
					mbi_thunk.Protect,
					&dwOldProtect);
#endif

				if (NULL != pdwHooked)
				{
					// Increment the total number of functions hooked.
					*pdwHooked += 1;
				}
			}
		}
		// Increment both tables.
		pOrigThunk++;
		pRealThunk++;
	}

	// All OK, JumpMaster!
	SetLastError(ERROR_SUCCESS);
	return (TRUE);
}

/*----------------------------------------------------------------------
FUNCTION        :   GetLoadedModules
DISCUSSION      :
    For the specified process id, this function returns the HMODULES for
all modules loaded into that process address space.  This function works
for both NT and 95.
PARAMETERS      :
    dwPID        - The process ID to look into.
    uiCount      - The number of slots in the paModArray buffer.  If
                   this value is 0, then the return value will be TRUE
                   and pdwRealCount will hold the number of items
                   needed.
    paModArray   - The array to place the HMODULES into.  If this buffer
                   is too small to hold the result and uiCount is not
                   zero, then FALSE is returned, but pdwRealCount will
                   be the real number of items needed.
    pdwRealCount - The count of items needed in paModArray, if uiCount
                   is zero, or the real number of items in paModArray.
RETURNS         :
    FALSE - There was a problem, check GetLastError.
    TRUE  - The function succeeded.  See the parameter discussion for
            the output parameters.
----------------------------------------------------------------------*/
BOOL GetLoadedModules( DWORD dwPID, UINT uiCount, HMODULE * paModArray, LPDWORD pdwRealCount )
{
    // Do the parameter checking for real.  Note that I only check the
    //  memory in paModArray if uiCount is > 0.  The user can pass zero
    //  in uiCount if they are just interested in the total to be
    //  returned so they could dynamically allocate a buffer.
    if ( ( TRUE == IsBadWritePtr ( pdwRealCount , sizeof(UINT) ) )    ||
         ( ( uiCount > 0 ) &&
           ( TRUE == IsBadWritePtr ( paModArray ,
                                     uiCount * sizeof(HMODULE) ) ) )   )
    {
        SetLastErrorEx ( ERROR_INVALID_PARAMETER , SLE_ERROR ) ;
        return ( 0 ) ;
    }

    // Figure out which OS we are on.
    OSVERSIONINFO stOSVI ;

    memset ( &stOSVI , NULL , sizeof ( OSVERSIONINFO ) ) ;
    stOSVI.dwOSVersionInfoSize = sizeof ( OSVERSIONINFO ) ;

    BOOL bRet = GetVersionEx ( &stOSVI ) ;
    if ( FALSE == bRet )
    {
        return ( FALSE ) ;
    }

    // Check the version and call the appropriate thing.
    if ( ( VER_PLATFORM_WIN32_NT == stOSVI.dwPlatformId ) &&
         ( 4 == stOSVI.dwMajorVersion                   )    )
    {
        // This is NT 4 so call it's specific version.
        return ( GetLoadedModulesNT4 ( dwPID        ,
                                       uiCount      ,
                                       paModArray   ,
                                       pdwRealCount  ) );
    }
    else
    {
        // Everyone other OS goes through TOOLHELP32.
        return ( GetLoadedModulesTLHELP ( dwPID         ,
                                          uiCount       ,
                                          paModArray    ,
                                          pdwRealCount   ) ) ;
    }
}

/*----------------------------------------------------------------------
FUNCTION        :   GetLoadedModulesNT4
DISCUSSION      :
    The NT4 specific version of GetLoadedModules.  This function assumes
that GetLoadedModules does the work to validate the parameters.
PARAMETERS      :
    dwPID        - The process ID to look into.
    uiCount      - The number of slots in the paModArray buffer.  If
                   this value is 0, then the return value will be TRUE
                   and pdwRealCount will hold the number of items
                   needed.
    paModArray   - The array to place the HMODULES into.  If this buffer
                   is too small to hold the result and uiCount is not
                   zero, then FALSE is returned, but pdwRealCount will
                   be the real number of items needed.
    pdwRealCount - The count of items needed in paModArray, if uiCount
                   is zero, or the real number of items in paModArray.
RETURNS         :
    FALSE - There was a problem, check GetLastError.
    TRUE  - The function succeeded.  See the parameter discussion for
            the output parameters.
----------------------------------------------------------------------*/
BOOL GetLoadedModulesNT4( DWORD dwPID, UINT uiCount, HMODULE * paModArray, LPDWORD pdwRealCount )
{
    // Initialize PSAPI.DLL, if needed.
    if ( FALSE == InitPSAPI ( ) )
    {
        SetLastErrorEx ( ERROR_DLL_INIT_FAILED , SLE_ERROR ) ;
        return ( FALSE ) ;
    }

    // Convert the process ID into a process handle.
    HANDLE hProc = OpenProcess ( PROCESS_QUERY_INFORMATION |
                                    PROCESS_VM_READ         ,
                                 FALSE                      ,
                                 dwPID                       ) ;
    if ( NULL == hProc )
    {
        return ( FALSE ) ;
    }

    // Now get the modules for the specified process.
    // Because of possible DLL unload order differences, make sure that
    //  PSAPI.DLL is still loaded in case this function is called during
    //  shutdown.
    if ( TRUE == IsBadCodePtr ( (FARPROC)g_pEnumProcessModules ) )
    {
        // Close the process handle used.
        CloseHandle ( hProc );

        SetLastErrorEx ( ERROR_INVALID_DLL , SLE_ERROR ) ;

        return ( FALSE ) ;
    }

    DWORD dwTotal = 0 ;
    BOOL bRet = g_pEnumProcessModules ( hProc                        ,
                                        paModArray                   ,
                                        uiCount * sizeof ( HMODULE ) ,
                                        &dwTotal                      );

    // Close the process handle used.
    CloseHandle ( hProc );

    // Convert the count from bytes to HMODULE values.
    *pdwRealCount = dwTotal / sizeof ( HMODULE ) ;

    // If bRet was FALSE, and the user was not just asking for the
    //  total, there was a problem.
    if ( ( ( FALSE == bRet ) && ( uiCount > 0 ) ) || ( 0 == dwTotal ) )
    {
        return ( FALSE ) ;
    }

    // If the total returned in pdwRealCount is larger than the value in
    //  uiCount, then return an error.  If uiCount is zero, then it is
    //  not an error.
    if ( ( *pdwRealCount > uiCount ) && ( uiCount > 0 ) )
    {
        SetLastErrorEx ( ERROR_INSUFFICIENT_BUFFER , SLE_ERROR ) ;
        return ( FALSE ) ;
    }

    // All OK, Jumpmaster!
    SetLastError ( ERROR_SUCCESS ) ;
    return ( TRUE ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   GetLoadedModulesTLHELP
DISCUSSION      :
    The TOOLHELP32 specific version of GetLoadedModules.  This function
assumes that GetLoadedModules does the work to validate the parameters.
PARAMETERS      :
    dwPID        - The process ID to look into.
    uiCount      - The number of slots in the paModArray buffer.  If
                   this value is 0, then the return value will be TRUE
                   and pdwRealCount will hold the number of items
                   needed.
    paModArray   - The array to place the HMODULES into.  If this buffer
                   is too small to hold the result and uiCount is not
                   zero, then FALSE is returned, but pdwRealCount will
                   be the real number of items needed.
    pdwRealCount - The count of items needed in paModArray, if uiCount
                   is zero, or the real number of items in paModArray.
RETURNS         :
    FALSE - There was a problem, check GetLastError.
    TRUE  - The function succeeded.  See the parameter discussion for
            the output parameters.
----------------------------------------------------------------------*/
BOOL GetLoadedModulesTLHELP( DWORD dwPID, UINT uiCount, HMODULE * paModArray, LPDWORD pdwRealCount )
{
    // Always set pdwRealCount to a know value before anything else.
    *pdwRealCount = 0 ;

    if ( FALSE == InitTOOLHELP32 ( ) )
    {
        SetLastErrorEx ( ERROR_DLL_INIT_FAILED , SLE_ERROR ) ;
        return ( FALSE ) ;
    }

    // The snapshot handle.
    HANDLE        hModSnap     = NULL ;
    // The module structure.
    MODULEENTRY32 stME32              ;
    // A flag kept to report if the buffer was too small.
    BOOL          bBuffToSmall = FALSE ;


    // Get the snapshot for the specified process.
    hModSnap = g_pCreateToolhelp32Snapshot ( TH32CS_SNAPMODULE ,
                                             dwPID              ) ;
    if ( INVALID_HANDLE_VALUE == hModSnap )
    {
        return ( FALSE ) ;
    }

    memset ( &stME32 , NULL , sizeof ( MODULEENTRY32 ) ) ;
    stME32.dwSize = sizeof ( MODULEENTRY32 ) ;

    // Start getting the module values.
    if ( TRUE == g_pModule32First ( hModSnap , &stME32 ) )
    {
        do
        {
            // If uiCount is not zero, copy values.
            if ( 0 != uiCount )
            {
                // If the passed in buffer is to small, set the flag.
                //  This is so we match the functionality of the NT4
                //  version of this function which will return the
                //  correct total needed.
                if ( ( TRUE == bBuffToSmall     ) ||
                     ( *pdwRealCount == uiCount )   )
                {
                    bBuffToSmall = TRUE ;
                    break ;
                }
                else
                {
                    // Copy this value in.
                    paModArray[ *pdwRealCount ] =
                                         (HINSTANCE)stME32.modBaseAddr ;
                }
            }
            // Bump up the real total count.
            *pdwRealCount += 1 ;
        }
        while ( ( TRUE == g_pModule32Next ( hModSnap , &stME32 ) ) ) ;
    }
    else
    {
        DWORD dwLastErr = GetLastError ( ) ;
        SetLastErrorEx ( dwLastErr , SLE_ERROR ) ;
        return ( FALSE ) ;
    }

    // Close the snapshot handle.
    CloseHandle ( hModSnap );

    // Check if the buffer was too small.
    if ( TRUE == bBuffToSmall )
    {
        SetLastErrorEx ( ERROR_INSUFFICIENT_BUFFER , SLE_ERROR ) ;
        return ( FALSE ) ;
    }

    // All OK, Jumpmaster!
    SetLastError ( ERROR_SUCCESS ) ;
    return ( TRUE ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   GetModuleBaseNamePortable
DISCUSSION      :
    Returns the base name of the specified module in a manner that is
portable between NT and Win95/98.
PARAMETERS      :
    hProcess   - The handle to the process.  In Win95/98 this is
                 ignored.
    hModule    - The module to look up.  If this is NULL, then the
                 module returned is the executable.
    lpBaseName - The buffer that recieves the base name.
    nSize      - The size of the buffer.
RETURNS         :
    !0 - The length of the string copied to the buffer.
    0  - The function failed.  To get extended error information,
         call GetLastError
----------------------------------------------------------------------*/
DWORD GetModuleBaseNamePortable( HANDLE hProcess, HMODULE hModule, LPTSTR lpBaseName, DWORD nSize )
{
    if ( TRUE == IsNT ( ) )
    {
        // Call the NT version.  It is in NT4ProcessInfo because that is
        //  where all the PSAPI wrappers are kept.
        return ( GetModuleBaseNameNT ( hProcess     ,
                                       hModule      ,
                                       lpBaseName   ,
                                       nSize         ) ) ;
    }
    return ( GetModuleBaseName9x ( hProcess     ,
                                      hModule      ,
                                      LPSTR(lpBaseName)   ,
                                      nSize         ) ) ;

}

/*----------------------------------------------------------------------
FUNCTION        :   GetModuleBaseNameNT
----------------------------------------------------------------------*/
DWORD GetModuleBaseNameNT ( HANDLE  hProcess   ,
                                      HMODULE hModule    ,
                                      LPTSTR  lpBaseName ,
                                      DWORD   nSize       )
{
    // Initialize PSAPI.DLL, if needed.
    if ( FALSE == InitPSAPI ( ) )
    {
        SetLastErrorEx ( ERROR_DLL_INIT_FAILED , SLE_ERROR ) ;
        return ( FALSE ) ;
    }
    return ( g_pGetModuleBaseName ( hProcess    ,
                                    hModule     ,
                                    lpBaseName  ,
                                    nSize        ) ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   GetModuleBaseName9x
----------------------------------------------------------------------*/
static DWORD GetModuleBaseName9x( HANDLE /*hProcess*/, HMODULE hModule, LPSTR lpBaseName, DWORD nSize )
{
    if ( TRUE == IsBadWritePtr ( lpBaseName , nSize ) )
    {
        SetLastError ( ERROR_INVALID_PARAMETER ) ;
        return ( 0 ) ;
    }

    // This could blow the stack...
    char szBuff[ MAX_PATH + 1 ] ;
    DWORD dwRet = GetModuleFileNameA( hModule , szBuff , MAX_PATH ) ;
    if ( 0 == dwRet )
    {
        return ( 0 ) ;
    }

    // Find the last '\' mark.
    char * pStart = strrchr ( szBuff , '\\' ) ;
    int iMin ;
    if ( NULL != pStart )
    {
        // Move up one character.
        pStart++ ;
        //lint -e666
        iMin = min ( (int)nSize , (lstrlenA( pStart ) + 1) ) ;
        //lint +e666
        lstrcpynA( lpBaseName , pStart , iMin ) ;
    }
    else
    {
        // Copy the szBuff buffer in.
        //lint -e666
        iMin = min ( (int)nSize , (lstrlenA( szBuff ) + 1) ) ;
        //lint +e666
        lstrcpynA( lpBaseName , szBuff , iMin ) ;
    }
    // Always NULL terminate.
    lpBaseName[ iMin ] = '\0' ;
    return ( (DWORD)(iMin - 1) ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   IsNT
DISCUSSION      :
    Returns TRUE if the operating system is NT.  I simply got tired of
always having to call GetVersionEx each time I needed to check.
Additionally, I also need to check the OS inside loops so this function
caches the results so it is faster.
PARAMETERS      :
    None.
RETURNS         :

----------------------------------------------------------------------*/
BOOL IsNT ( void )
{
    if ( TRUE == g_bHasVersion )
    {
        return ( TRUE == g_bIsNT ) ;
    }

    OSVERSIONINFO stOSVI ;

    memset ( &stOSVI , NULL , sizeof ( OSVERSIONINFO ) ) ;
    stOSVI.dwOSVersionInfoSize = sizeof ( OSVERSIONINFO ) ;

    BOOL bRet = GetVersionEx ( &stOSVI ) ;
    if ( FALSE == bRet )
    {
        return ( FALSE ) ;
    }

    // Check the version and call the appropriate thing.
    if ( VER_PLATFORM_WIN32_NT == stOSVI.dwPlatformId )
    {
        g_bIsNT = TRUE ;
    }
    else
    {
        g_bIsNT = FALSE ;
    }
    g_bHasVersion = TRUE ;
    return ( TRUE == g_bIsNT ) ;
}

/*----------------------------------------------------------------------
FUNCTION        :   GetNamedImportDescriptor
DISCUSSION      :
Gets the import descriptor for the requested module.  If the module
is not imported in hModule, NULL is returned.
This is a potential useful function in the future.
PARAMETERS      :
hModule      - The module to hook in.
szImportMod  - The module name to get the import descriptor for.
RETURNS         :
NULL  - The module was not imported or hModule is invalid.
!NULL - The import descriptor.
----------------------------------------------------------------------*/
PIMAGE_IMPORT_DESCRIPTOR GetNamedImportDescriptor( HMODULE hModule, LPCSTR szImportMod)
{
#ifdef _DEBUG
	// Always check parameters.
	/*ASSERT(NULL != szImportMod);
	ASSERT(NULL != hModule);*/
#endif
	if ((NULL == szImportMod) || (NULL == hModule))
	{
		SetLastErrorEx(ERROR_INVALID_PARAMETER, SLE_ERROR);
		return (NULL);
	}

	PIMAGE_DOS_HEADER pDOSHeader = (PIMAGE_DOS_HEADER)hModule;

	// Is this the MZ header?
	if ((TRUE == IsBadReadPtr(pDOSHeader,
		sizeof(IMAGE_DOS_HEADER))) ||
		(IMAGE_DOS_SIGNATURE != pDOSHeader->e_magic))
	{
#ifdef _DEBUG
		/*ASSERT(!"Unable to read MZ header!");*/
#endif
		SetLastErrorEx(ERROR_INVALID_PARAMETER, SLE_ERROR);
		return (NULL);
	}

	// Get the PE header.
	PIMAGE_NT_HEADERS pNTHeader = MakePtr(PIMAGE_NT_HEADERS,
		pDOSHeader,
		pDOSHeader->e_lfanew);

	// Is this a real PE image?
	if ((TRUE == IsBadReadPtr(pNTHeader, sizeof(IMAGE_NT_HEADERS))) || (IMAGE_NT_SIGNATURE != pNTHeader->Signature))
	{
#ifdef _DEBUG
		/*ASSERT(!"Not able to read PE header");*/
#endif
		SetLastErrorEx(ERROR_INVALID_PARAMETER, SLE_ERROR);
		return (NULL);
	}

	// If there is no imports section, leave now.
	if (0 == pNTHeader->OptionalHeader.
		DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].
		VirtualAddress)
	{
		return (NULL);
	}

	// Get the pointer to the imports section.
	PIMAGE_IMPORT_DESCRIPTOR pImportDesc
		= MakePtr(PIMAGE_IMPORT_DESCRIPTOR,
		pDOSHeader,
		pNTHeader->OptionalHeader.
		DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].
		VirtualAddress);

	// Loop through the import module descriptors looking for the
	// module whose name matches szImportMod.
	while (NULL != pImportDesc->Name)
	{
		PSTR szCurrMod = MakePtr(PSTR,
			pDOSHeader,
			pImportDesc->Name);
		if (0 == _stricmp(szCurrMod, szImportMod))
		{
			// Found it.
			break;
		}
		// Look at the next one.
		pImportDesc++;
	}

	// If the name is NULL, then the module is not imported.
	if (NULL == pImportDesc->Name)
	{
		return (NULL);
	}

	// All OK, Jumpmaster!
	return (pImportDesc);

}
