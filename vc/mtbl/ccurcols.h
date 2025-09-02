// ccurcols.h

#ifndef _CCURCOLS_H
#define _CCURCOLS_H

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <map>
#include <vector>
#include "ctd.h"

// Globale Variablen
extern CCtd Ctd;

// Using
using namespace std;

// Typedefs
typedef map<HWND, int> CURCOLS_POS_MAP;
typedef vector<POINT> CURCOLS_COORD_VECTOR;
typedef vector<HWND> CURCOLS_HANDLE_VECTOR;
typedef vector<long> CURCOLS_WIDTH_VECTOR;

// Klasse CCurCols
class CCurCols
{
public:
	CCurCols( HWND hWndTbl );
	
	int GetPos( HWND hWndCol );
	CURCOLS_HANDLE_VECTOR & GetVectorHandle( );
	CURCOLS_WIDTH_VECTOR & GetVectorWidth( );
	void Refresh( );
	void SetInvalid( ) { m_bValid = FALSE; };
	void UpdateWidth( HWND hWndCol );

private:
	BOOL m_bValid;
	HWND m_hWndTbl;
	CURCOLS_POS_MAP m_pm;
	CURCOLS_HANDLE_VECTOR m_vh;
	CURCOLS_WIDTH_VECTOR m_vw;

	CCurCols( );

};

#endif // _CCURCOLS_H