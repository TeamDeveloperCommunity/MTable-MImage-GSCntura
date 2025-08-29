// mtexcexp.h : Haupt-Header-Datei für die DLL MEXCEL8
//

#if !defined(AFX_MTEXCEXP_H__7F26ADC4_E250_11D5_A9AB_00104B88A5A3__INCLUDED_)
#define AFX_MTEXCEXP_H__7F26ADC4_E250_11D5_A9AB_00104B88A5A3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// Hauptsymbole
#include "winversion.h"
#include "ctd.h"
#include "excel8.h"
#include "openoffice.h"
#include "excconst.h"
#include "ExpTblDlgThread.h"
#include "mtexp.h"
#include "mtexcexpapi.h"
#include "mtblapi.h"

// Konstanten

// Strukturen
struct EXCEL_EXP_PARAMS
{
	HWND hWndTbl;
	int iLanguage;
	CString sExcelStartCell;
	CString sSaveAsFile;
	CString sTemplateFile;
	CString sSheetName;
	DWORD dwExcelFlags;
	DWORD dwExpFlags;
	WORD wRowFlagsOn;
	WORD wRowFlagsOff;
	DWORD dwColFlagsOn;
	DWORD dwColFlagsOff;
};

#pragma pack(UDV_PACK)
struct EXCEL_EXP_PARAMS_UDV
{
#ifdef UDV_OFFSET_SIZE
	BYTE Offset[UDV_OFFSET_SIZE];
#endif
	NUMBER Language;
	HSTRING ExcelStartCell;
	HSTRING ExcelSaveAsFile;
	HSTRING ExcelTemplateFile;
	HSTRING ExcelSheetName;
	NUMBER ExcelFlags;
	NUMBER ExpFlags;
	NUMBER RowFlagsOn;
	NUMBER RowFlagsOff;
	NUMBER ColFlagsOn;
	NUMBER ColFlagsOff;
};
#pragma pack()

struct RANGEFORMAT
{	
	DWORD dwUse;
	Range R;
	DWORD dwTextColor;
	DWORD dwBackColor;
	CString sNumberFormat;
	long lIndentLevel;
	FONTINFO fi;
	long lJustify;
	long lVAlign;
	int iUnionCount;

	void Init();

	BOOL operator==(RANGEFORMAT& rf);
};

struct RANGELINES
{
	DWORD dwUse;
	Range R;
	MTBLLINEDEF ldHorzInner;
	MTBLLINEDEF ldVertInner;
	MTBLLINEDEF ldEdgeLeft;
	MTBLLINEDEF ldEdgeTop;
	MTBLLINEDEF ldEdgeRight;
	MTBLLINEDEF ldEdgeBottom;
	MTBLLINEDEF ldDiagUp;
	MTBLLINEDEF ldDiagDown;

	int iUnionCount;

	void Init();
	void ApplyLineDef(DWORD dwDefType, MTBLLINEDEF& ld);

	BOOL operator==(RANGELINES& rf);
};

// Typedefs
typedef CArray<COLORREF, COLORREF> COLORARRAY;
typedef CArray<BOOL, BOOL> BOOLARRAY;
typedef CList<RANGEFORMAT, RANGEFORMAT&> RANGEFORMAT_LIST;
typedef CList<RANGELINES, RANGELINES&> RANGELINE_LIST;

// Funktionen
void AttachRange( _Worksheet &Sheet, Range &R, int iColFrom, int iRowFrom, int iColTo, int iRowTo );
void AttachRange( _Worksheet &Sheet, Range &R, CString &sFrom, CString &sTo );
int ExportToExcel( EXCEL_EXP_PARAMS &params );
void GetRangeStr( int iCol, int iRow, CString &s );
void GetStart( Range &R, int &iStartCol, int &iStartRow );
void SetBorder( HWND hWndTbl, Border &B, MTBLLINEDEF &ld );
void SetBorder(HWND hWndTbl, Range &R, int iBorder, MTBLLINEDEF &ld);
void SetFont( Font &F, FONTINFO &fi );
void SetRangelines(RANGELINE_LIST& RLs);

/////////////////////////////////////////////////////////////////////////////
// CMtexcexpApp
// Siehe mexcel8.cpp für die Implementierung dieser Klasse
//

class CMtexcexpApp : public CWinApp	
{
public:
	CMtexcexpApp();

// Überladungen
	// Vom Klassenassistenten generierte Überladungen virtueller Funktionen
	//{{AFX_VIRTUAL(CMtexcexpApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CMtexcexpApp)
		// HINWEIS - An dieser Stelle werden Member-Funktionen vom Klassen-Assistenten eingefügt und entfernt.
		//    Innerhalb dieser generierten Quelltextabschnitte NICHTS VERÄNDERN!
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ fügt unmittelbar vor der vorhergehenden Zeile zusätzliche Deklarationen ein.

#endif // !defined(AFX_MTEXCEXP_H__7F26ADC4_E250_11D5_A9AB_00104B88A5A3__INCLUDED_)
