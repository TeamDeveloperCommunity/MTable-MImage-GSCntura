// lang.cpp

// Includes
#include "lang.h"
#include "mtblapi.h"

// Globale Variablen
//extern CCtd Ctd;

// Funktionen

BOOL GetUIStr( int iLanguage, const CString &sKey, CString &sString )
{
	int iLen = MTblGetLanguageStringC( iLanguage, _T("Export"), sKey, NULL );
	if ( iLen <= 0 )
		return FALSE;
	
	LPTSTR lpsBuf = sString.GetBuffer( iLen );
	MTblGetLanguageStringC( iLanguage, _T("Export"), sKey, lpsBuf );
	sString.ReleaseBuffer( -1 );

	return TRUE;
}