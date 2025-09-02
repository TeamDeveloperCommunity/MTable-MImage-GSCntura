// cmtblsortval.h

#ifndef _INC_CMTBLSORTVAL
#define _INC_CMTBLSORTVAL

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include "ctd.h"

extern CCtd Ctd;

class CMTblSortVal
{
private:
	// Variablen
	int m_iType;				// Datentyp
	LPVOID m_pData;				// Daten-Zeiger
	// Funktionen
	void FreeData( );
	void InitMembers( );
	int CmpDate( DATETIME &dt1, DATETIME &dt2 );
	//int CmpNumber( NUMBER &n1, NUMBER &n2 );
	int CmpNumber( NUMBER &n1, NUMBER &n2 );
	int CmpString( LPCTSTR lpcts1, LPCTSTR lpcts2, BOOL bCaseInsensitive = FALSE, BOOL bStringSort = FALSE );
public:
	// Konstruktor(en)
	CMTblSortVal( );
	// Destruktor
	~CMTblSortVal( );
	// Funktionen
	int Compare( CMTblSortVal &Val, BOOL bCaseInsensitive = FALSE, BOOL bStringSort = FALSE );
	BOOL SetDate( DATETIME &dtVal );
	BOOL SetNumber( NUMBER &nVal );
	BOOL SetString( LPCTSTR lpctsVal );
};

#endif // _INC_CMTBLSORTVAL