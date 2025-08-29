// mtoocalcexp.h

#ifndef _INC_MTOOCALCEXP
#define _INC_MTOOCALCEXP

// Includes
#include "stdafx.h"
#include "resource.h"		// Hauptsymbole
#include "winversion.h"
#include "ctd.h"
#include "openoffice.h"
#include "ExpTblDlgThread.h"
#include "mtexp.h"
#include "mtexcexpapi.h"
#include "mtblapi.h"

// Strukturen
struct OO_RANGEFORMAT
{	
	DWORD dwUse;
	OOCellRange R;
	int iColFrom;
	int iRowFrom;
	int iColTo;
	int iRowTo;
	DWORD dwTextColor;
	DWORD dwBackColor;
	long lNumberFormat;
	long lIndentLevel;
	FONTINFO fi;
	long lJustify;
	long lVAlign;
	//int iUnionCount;
};

struct OO_RANGELINES
{
	DWORD dwUse;
	OOCellRange R;
	LINEDEF ldHorzInner;
	LINEDEF ldVertInner;
	LINEDEF ldEdgeLeft;
	LINEDEF ldEdgeTop;
	LINEDEF ldEdgeRight;
	LINEDEF ldEdgeBottom;
	LINEDEF ldDiagUp;
	LINEDEF ldDiagDown;
};

// Funktionen
void SetBorder( HWND hWndTbl, OOCellRange &R, int iType, LINEDEF &ld );
void SetBorder( HWND hWndTbl, OOCellRange &R, int iType, short shOuterWidth, long lColor = 0, short shInnerWidth = 0, short shLineDistance = 0 );

#endif //_INC_MTOOCALCEXP