// cmtblcolhdrgrps.h

#ifndef _INC_CMTBLCOLHDRGRPS
#define _INC_CMTBLCOLHDRGRPS

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <list>
#include <map>
#include <stack>
#include "tstring.h"
#include "cmtblcolhdrgrp.h"

// Using
using namespace std;

// Typedefs
typedef list<CMTblColHdrGrp*> MTblColHdrGrpList;
typedef MTblColHdrGrpList *PMTblColHdrGrpList;

// Klasse CMTblColHdrGrps
class CMTblColHdrGrps
{
public:
	MTblColHdrGrpList m_List;
	// Konstruktore
	CMTblColHdrGrps( );
	// Destruktor
	~CMTblColHdrGrps( );
	// Funktionen
	CMTblColHdrGrp * CreateGrp( tstring &sText, list<HWND> &liCols, DWORD dwFlags = 0 );
	void DeleteAll( );
	BOOL DeleteGrp( CMTblColHdrGrp *pGrp );
	long EnumGrps( HARRAY haGrps );
	CMTblColHdrGrp * GetGrpOfCol( HWND hWndCol );
	BOOL IsValidGrp( CMTblColHdrGrp *pGrp );
	BOOL SetCol( CMTblColHdrGrp *pGrp, HWND hWndCol, BOOL bSet );
};

#endif // _INC_CMTBLCOLHDRGRPS