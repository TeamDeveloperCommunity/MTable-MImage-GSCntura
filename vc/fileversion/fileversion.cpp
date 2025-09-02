// fileversion.cpp

// Includes
#include "fileversion.h"

BOOL FileVerGetVersionStr( LPTSTR lpsFile, tstring & sVer )
{
	BOOL bOk = FALSE;
	DWORD dwHandle, dwLen;
	int i, iLen;
	LPVOID pData = NULL, pValue;
	UINT uiLen;
	TCHAR szKey[255] = _T(""), szLang[255] = _T("");

	dwLen = GetFileVersionInfoSize( lpsFile, &dwHandle );
	if ( dwLen == 0 )
		return FALSE;

	pData = malloc( dwLen );
	if ( !pData )
		return FALSE;

	if ( !GetFileVersionInfo( lpsFile, dwHandle, dwLen, pData ) )
		goto cleanup;

	if ( !VerQueryValue( pData, _T("\\VarFileInfo\\Translation"), &pValue, &uiLen ) )
		goto cleanup;
	wsprintf( szLang, _T("%4x%4x"), LOWORD(*LPDWORD(pValue)), HIWORD(*LPDWORD(pValue)) );
	iLen = lstrlen( szLang );
	for ( i = 0; i < iLen; i++ )
	{
		if ( szLang[i] == _T(' ') )
			szLang[i] = _T('0');
	}

	lstrcpy( szKey, _T("\\StringFileInfo\\") );
	lstrcat( szKey, szLang );
	lstrcat( szKey, _T("\\ProductVersion") );
	if ( !VerQueryValue( pData, szKey, &pValue, &uiLen ) )
		goto cleanup;

	sVer = LPTSTR(pValue);

	bOk = TRUE;

cleanup:
		if ( pData )
			free( pData );

	return bOk;
}