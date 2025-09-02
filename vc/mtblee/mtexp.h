// mtexp.h

#ifndef _INC_MTEXP
#define _INC_MTEXP

// Includes
#include "stdafx.h"
#include "export.h"
#include "mtblapi.h"
#include "excconst.h"
#include "ooconst.h"
#include "ExpTblDlgThread.h"
#include "lang.h"

// Konstanten
#define APP_EXCEL 0
#define APP_OOCALC 1

#define ERR_COPY_TO_CLIPBOARD _T("Error.CopyToClipboard")
#define ERR_COPY_TO_CLIPBOARD_OPEN _T("Error.CopyToClipboard.Open")
#define ERR_COPY_TO_CLIPBOARD_ALLOC _T("Error.CopyToClipboard.Alloc")
#define ERR_COPY_TO_CLIPBOARD_SETDATA _T("Error.CopyToClipboard.SetData")
#define ERR_GET_FONT_INFO _T("Error.GetFontInfo")
#define ERR_GET_OBJECT _T("Error.GetObject")
#define ERR_GET_START_CELL _T("Error.GetStartCell")
#define ERR_INSERT_SHEET _T("Error.InsertSheet")
#define ERR_SAVE_CLIPBOARD _T("Error.SaveClipboard")
#define ERR_SAVE_CLIPBOARD_OPEN _T("Error.SaveClipboard.Open")

#define GC_ERR_CANCELLED -1

#define MAX_INDENT_LEVEL 15L
#define MAX_TEXT_LEN 255L

#define USE_TEXT_COLOR 1
#define USE_BACK_COLOR 2
#define USE_NUMBERFORMAT 4
#define USE_FONT 8
#define USE_INDENT_LEVEL 16
#define USE_JUSTIFY 32
#define USE_VALIGN 64

#define USE_HORZ_INNER 1
#define USE_VERT_INNER 2
#define USE_EDGE_LEFT 4
#define USE_EDGE_TOP 8
#define USE_EDGE_RIGHT 16
#define USE_EDGE_BOTTOM 32
#define USE_DIAG_UP 64
#define USE_DIAG_DOWN 128

// Strukturen

struct LINEDEF
{
	int iStyle;
	DWORD dwColor;
	int iThickness;

	BOOL operator==(LINEDEF &ld);
	BOOL operator!=(LINEDEF &ld);
};

struct CELLMERGE
{
	int iFromCol;
	int iFromLine;
	int iToCol;
	int iToLine;
};

struct COLHDRGRP
{
	int iFromCol;
	int iFromLine;
	int iToCol;
	int iToLine;
	BOOL bNoInnerVLines;
	BOOL bNoInnerHLine;
	BOOL bNoColHeaders;
	BOOL bTxtAlignLeft;
	BOOL bTxtAlignRight;
	MTBLLINEDEF HorzLineDef;
	MTBLLINEDEF VertLineDef;
};

struct COLINFO
{
	HWND hWnd;
	int iPos;
	long lJustify;
	long lHdrJustify;
	long lVAlign;
	BOOL bInvisible;
	BOOL bTreeCol;
	long lDataType;
	int iCellType;
	BOOL bLocked;
	LPVOID pHdrGrp;
	BOOL bNoRowLines;
	BOOL bNoColLine;
	MTBLLINEDEF HorzLineDef;
	MTBLLINEDEF VertLineDef;
	BOOL bExportAsText;
	CString sFormat;
	CArray<CString,CString&> CellTexts;
};

struct FONTINFO
{
	CString sName;
	int iPointSize;
	BOOL bBold;
	BOOL bItalic;
	BOOL bUnderline;
	BOOL bStrikeOut;

	void Init();

	BOOL operator==( FONTINFO &fi );
	BOOL operator!=( FONTINFO &fi );
};

//Typedefs
typedef CArray<COLINFO,COLINFO> COLINFO_ARRAY;

// Makros
#define IS_CANCELED ( pDlgThread != NULL ? pDlgThread->IsCanceled( ) : FALSE )
#define MAKE_STATUS_DLG_TOPMOST(b) ( pDlgThread != NULL ? pDlgThread->MakeDlgTopmost( b ) : NULL )
#define SET_STATUS(s) ( pDlgThread != NULL ? PostMessage( pDlgThread->GetDlgHandle( ), WM_USER, 0, LPARAM( LPCTSTR( s ) ) ) : NULL )
#define SET_STATUS_LNG(i) ( pDlgThread != NULL ? pDlgThread->SetLanguage( i ) : NULL )

// Funktionen
int CharToUtf8( LPCTSTR lpChar, LPSTR lpCharUtf8, int iLenUtf8 );
void CopyToClipboard( CString & s, HWND hWnd = NULL, BOOL bHTML = FALSE );
int GetCols( BOOL bOpenOffice, HWND hWndTbl, DWORD dwColFlagsOn, DWORD dwColFlagsOff, int iRealLockedCols, HWND hWndTreeCol, COLINFO_ARRAY &Cols, int iStartCol, int iMaxCols, BOOL bStringColsAsText, BOOL &bAnyHdrGrp, BOOL &bAnyHdrGrpWCH, CExpTblDlgThread *pDlgThread );
void GetErrorMsg( LPCTSTR lpctsID, int iLanguage, CString & sMsg );
void GetFontInfo( HWND hWnd, HFONT hFont, FONTINFO & fi );
void HandleException( COleDispatchException * pE, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage );
void HandleException( COleException * pE, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage );
void HandleException( CException * pE, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage );
void HandleException( CString & sError, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage );
void MakeHtmlStr( CString & sText );
int MsgBoxException( CString & sText, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage );
void SaveClipboard( CString & s, HWND hWnd = NULL );
void ThrowError( LPCTSTR lpctsID );
void VMakeOptional( VARIANT & v );

#endif // _INC_MTEXP
