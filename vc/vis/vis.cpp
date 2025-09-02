// vis.cpp

// Includes
#include "vis.h"

// Konstruktor
CVis::CVis( void )
{
	memset( this, 0, sizeof(CVis) );

	// DLL laden
	TCHAR szDLL[MAX_PATH] = _T("");

	#ifdef TD11
	lstrcpy( szDLL, _T("VTI11.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD15
	lstrcpy( szDLL, _T("VTI15.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD20
	lstrcpy( szDLL, _T("VTI20.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD21
	lstrcpy( szDLL, _T("VTI21.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD30
	lstrcpy( szDLL, _T("VTI30.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD31
	lstrcpy( szDLL, _T("VTI31.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD40
	lstrcpy( szDLL, _T("VTI40.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD41
	lstrcpy( szDLL, _T("VTI41.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD42
	lstrcpy( szDLL, _T("VTI42.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD51
	lstrcpy( szDLL, _T("VTI51.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD52
	lstrcpy( szDLL, _T("VTI52.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD60
	lstrcpy( szDLL, _T("VTI60.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD61
	lstrcpy( szDLL, _T("VTI61.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD62
	lstrcpy( szDLL, _T("VTI62.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD63
	lstrcpy( szDLL, _T("VTI63.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD70
	lstrcpy( szDLL, _T("VTI70.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD71
	lstrcpy( szDLL, _T("VTI71.DLL") );
	m_hInst = LoadLibrary(szDLL);
	#elif TD72
	lstrcpy(szDLL, _T("VTI72.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#elif TD73
	lstrcpy(szDLL, _T("VTI73.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#elif TD74
	lstrcpy(szDLL, _T("VTI74.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#elif TD75
	lstrcpy(szDLL, _T("VTI75.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#endif


	if ( m_hInst )
	{
		ListGetItemHandle = (VISLISTGETITEMHANDLE)GetProcAddress(m_hInst, "VisListGetItemHandle");
		PicLoad = (VISPICLOAD)GetProcAddress( m_hInst, "VisPicLoad" );
		TblFindString = (VISTBLFINDSTRING)GetProcAddress(m_hInst, "VisTblFindString");
	}
	else
	{
		TCHAR szMsg[1024] = _T("");
		wsprintf( szMsg, _T("Visual Toolchest runtime environment could not be loaded ( %s )."), szDLL );
		MessageBox( NULL, szMsg, _T("MICSTO"), MB_ICONSTOP | MB_OK );
		return;
	}

	m_Initialized = ListGetItemHandle && PicLoad && TblFindString;

	if ( !m_Initialized )
		MessageBox( NULL, _T("Not all required Visual Toolchest functions could be loaded."), _T("MICSTO"), MB_ICONSTOP | MB_OK );
}


// Destruktor
CVis::~CVis( void )
{
	if ( m_hInst )
	{
		FreeLibrary( m_hInst );
		m_hInst = NULL;
	}
}