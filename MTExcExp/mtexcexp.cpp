// mtexcexp.cpp : Legt die Initialisierungsroutinen für die DLL fest.
//

#include "stdafx.h"
#include "mtexcexp.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMtexcexpApp

BEGIN_MESSAGE_MAP(CMtexcexpApp, CWinApp)
	//{{AFX_MSG_MAP(CMtexcexpApp)
		// HINWEIS - Hier werden Mapping-Makros vom Klassen-Assistenten eingefügt und entfernt.
		//    Innerhalb dieser generierten Quelltextabschnitte NICHTS VERÄNDERN!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMtexcexpApp Konstruktion

CMtexcexpApp::CMtexcexpApp()
{
	// ZU ERLEDIGEN: Hier Code zur Konstruktion einfügen
	// Alle wichtigen Initialisierungen in InitInstance platzieren
}

/////////////////////////////////////////////////////////////////////////////
// Das einzige CMtexcexpApp-Objekt

CMtexcexpApp theApp;

BOOL CMtexcexpApp::InitInstance() 
{
	AfxOleInit( );	
	return CWinApp::InitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// Globale Variablen
CCtd Ctd;
CString sError;

/////////////////////////////////////////////////////////////////////////////
// LINEDEF

BOOL LINEDEF::operator==( LINEDEF &ld )
{
	return iStyle == ld.iStyle &&
		   dwColor == ld.dwColor &&
		   iThickness == ld.iThickness;
}

BOOL LINEDEF::operator!=( LINEDEF &ld )
{
	return !operator==( ld );
}

/////////////////////////////////////////////////////////////////////////////
// RANGEFORMAT
void RANGEFORMAT::Init()
{
	dwUse = 0;
	dwTextColor = xlAutomatic;
	dwBackColor = xlAutomatic;
	sNumberFormat.Empty();
	lIndentLevel = 0;
	fi.Init();
	lJustify = xlNone;
	lVAlign = xlNone;
	iUnionCount = 0;
}

BOOL RANGEFORMAT::operator == (RANGEFORMAT& rf)
{
	return
	dwUse == rf.dwUse &&
	dwTextColor == rf.dwTextColor &&
	dwBackColor == rf.dwBackColor &&
	sNumberFormat == rf.sNumberFormat &&
	lIndentLevel == rf.lIndentLevel &&
	fi == rf.fi &&
	lJustify == rf.lJustify &&
	lVAlign == rf.lVAlign;
}

/////////////////////////////////////////////////////////////////////////////
// RANGELINES
void RANGELINES::Init()
{
	dwUse = 0;
	ldHorzInner.Init();
	ldVertInner.Init();
	ldEdgeLeft.Init();
	ldEdgeTop.Init();
	ldEdgeRight.Init();
	ldEdgeBottom.Init();
	ldDiagUp.Init();
	ldDiagDown.Init();

	iUnionCount = 0;
}

void RANGELINES::ApplyLineDef(DWORD dwDefType, MTBLLINEDEF& ld)
{
	LPMTBLLINEDEF pld = NULL;

	switch (dwDefType)
	{
	case USE_HORZ_INNER:
		pld = &ldHorzInner;
		break;
	case USE_VERT_INNER:
		pld = &ldVertInner;
		break;
	case USE_EDGE_LEFT:
		pld = &ldEdgeLeft;
		break;
	case USE_EDGE_TOP:
		pld = &ldEdgeTop;
		break;
	case USE_EDGE_RIGHT:
		pld = &ldEdgeRight;
		break;
	case USE_EDGE_BOTTOM:
		pld = &ldEdgeBottom;
		break;
	case USE_DIAG_UP:
		pld = &ldDiagUp;
		break;
	case USE_DIAG_DOWN:
		pld = &ldDiagDown;
		break;
	}

	if (pld)
	{
		*pld += ld;

		if (pld->AnyUndef())
			dwUse |= dwDefType;
	}
}

BOOL RANGELINES::operator == (RANGELINES& rl)
{
	return
	dwUse == rl.dwUse &&
	dwUse & USE_HORZ_INNER ? ldHorzInner == rl.ldHorzInner : TRUE &&
	dwUse & USE_VERT_INNER ? ldVertInner == rl.ldVertInner : TRUE &&
	dwUse & USE_EDGE_LEFT ? ldEdgeLeft == rl.ldEdgeLeft : TRUE &&
	dwUse & USE_EDGE_TOP ? ldEdgeTop == rl.ldEdgeTop : TRUE &&
	dwUse & USE_EDGE_RIGHT ? ldEdgeRight == rl.ldEdgeRight : TRUE &&
	dwUse & USE_EDGE_BOTTOM ? ldEdgeBottom == rl.ldEdgeBottom : TRUE &&
	dwUse & USE_DIAG_UP ? ldDiagUp == rl.ldDiagUp : TRUE &&
	dwUse & USE_DIAG_DOWN ? ldDiagDown == rl.ldDiagDown : TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// Interne Functionen

void AttachRange( _Worksheet & Sheet, Range & R, int iColFrom, int iRowFrom, int iColTo = 0, int iRowTo = 0 )
{
	CString s, s2;
	COleVariant v, v2;
	LPDISPATCH pDisp;

	GetRangeStr( iColFrom, iRowFrom, s );
	v = s;

	if ( iColTo > 0 )
	{
		GetRangeStr( iColTo, iRowTo, s2 );
		v2 = s2;
	}
	else
		v2 = v;

	pDisp = Sheet.GetRange( v, v2 );	
	R.AttachDispatch( pDisp );
}

void AttachRange( _Worksheet & Sheet, Range & R, CString & sFrom, CString & sTo = CString( _T("") ) )
{
	COleVariant v, v2;
	LPDISPATCH pDisp;

	v = sFrom;

	if ( !sTo.IsEmpty( ) )
	{
		v2 = sTo;
	}
	else
		v2 = sFrom;

	pDisp = Sheet.GetRange( v, v2 );	
	R.AttachDispatch( pDisp );
}

/*void GetRangeStr( int iCol, int iRow, CString & s )
{
	CString sCol, sRow;
	int iMod;

	s.Empty( );

	iCol = min( max( iCol, 1), 256 );
	if ( iCol > 26 )
	{
		iMod = iCol % 26;
		iCol = iCol / 26;

		if ( iMod )
			sCol.Insert( 0, iMod + 64 );
		else
		{
			sCol.Insert( 0, _T('Z') );
			--iCol;
		}

		sCol.Insert( 0, iCol + 64 );
	}
	else
		sCol = TCHAR(iCol + 64);

	if ( iRow > 0 )
	{
		iRow = min( max( iRow, 1), 65536 );
		sRow.Format( _T("%d"), iRow );
		s = sCol + sRow;
	}
	else
		s = sCol;	
}*/


void GetRangeStr( int iCol, int iRow, CString & s )
{
	CString sCol, sRow;
	int iChar;

	s.Empty( );

	while ( iCol > 0 )
	{
		if ( iCol > 26 )
		{
			iChar = iCol % 26;
			iCol = iCol / 26;

			if ( !iChar )
			{
				iChar = 26;
				iCol--;
			}
		}
		else
		{
			iChar = iCol;
			iCol = 0;
		}

		sCol.Insert( 0, iChar + 64 );
	}

	if ( iRow > 0 )
	{
		sRow.Format( _T("%u"), iRow );
		s = sCol + sRow;
	}
	else
		s = sCol;	
}


void GetStart( _Application & App, int & iStartCol, int & iStartRow )
{
	COleVariant v;
	CString sAddr, sCol, sRow;
	int iPos;
	LPDISPATCH pDisp;
	Range R;

	pDisp = App.GetActiveCell( );
	R.AttachDispatch( pDisp );

	VMakeOptional( v );
	sAddr = R.GetAddress( v, v, xlR1C1, v, v );

	iPos = sAddr.ReverseFind( _T('C') );
	if ( iPos >= 2 )
	{
		sCol = sAddr.Mid( iPos + 1 );
		sRow = sAddr.Mid( 1, iPos - 1 );

		iStartCol = _ttoi( sCol );
		iStartRow = _ttoi( sRow );
	}
	else
		ThrowError( ERR_GET_START_CELL );
}

void OffsetRange( Range & R, int iRowOffs, int iColOffs )
{
	COleVariant v((long)iRowOffs), v2((long)iColOffs);
	
	LPDISPATCH pDisp = R.GetOffset( v, v2 );	
	R.AttachDispatch( pDisp );
}

void SetBorder(HWND hWndTbl, Border &B, MTBLLINEDEF &ld)
{
	COleVariant v;
	int iStyle;

	if (ld.Style != MTLS_UNDEF)
	{
		if (ld.Style == MTLS_INVISIBLE)
		{
			/*v = xlAutomatic;
			B.SetColorIndex(v);

			v = xlThin;
			B.SetWeight(v);*/

			v = xlNone;
			B.SetLineStyle(v);
		}
		else
		{
			v = xlContinuous;
			B.SetLineStyle(v);

			if (ld.Style == MTLS_DOT)
			{
				v = xlHairline;
				B.SetWeight(v);
			}				
		}
	}
	if (ld.Style != MTLS_INVISIBLE)
	{
		if (ld.Color != MTBL_COLOR_UNDEF && ld.Style != MTLS_INVISIBLE)
		{
			DWORD dwColor = ld.Color;

			DWORD dwSubstColor = MTblGetExportSubstColor(hWndTbl, ld.Color);
			if (dwSubstColor != MTBL_COLOR_UNDEF)
				dwColor = dwSubstColor;

			v = long(dwColor);
			B.SetColor(v);
		}

		if (ld.Thickness != MTBL_LINE_UNDEF_THICKNESS && ld.Style != MTLS_DOT && ld.Style != MTLS_INVISIBLE)
		{
			if (ld.Thickness >= 3)
				v = xlThick;
			else if (ld.Thickness >= 2)
				v = xlMedium;
			else
				v = xlThin;

			B.SetWeight(v);
		}
	}	
}

void SetBorder(HWND hWndTbl, Range &R, int iBorder, MTBLLINEDEF &ld)
{
	Borders Bs;
	Border B;
	LPDISPATCH pDisp;

	pDisp = R.GetBorders();
	Bs.AttachDispatch(pDisp);

	pDisp = Bs.GetItem(iBorder);
	B.AttachDispatch(pDisp);

	SetBorder(hWndTbl, B, ld);
}

/*
void SetBorder( HWND hWndTbl, Border &B, MTBLLINEDEF &ld )
{
	COleVariant v;
	DWORD dwSubstColor;

	if ( ld.Style == MTLS_INVISIBLE )
	{
		v = xlAutomatic;
		B.SetColorIndex(v);

		v = xlNone;
		B.SetLineStyle( v );
	}
	else
	{		
		v = xlContinuous;
		B.SetLineStyle(v);

		if (ld.Style == MTLS_SOLID)
		{
			if (ld.Thickness >= 3)
				v = xlThick;
			else if (ld.Thickness >= 2)
				v = xlMedium;
			else
				v = xlThin;
		}			
		else
			v = xlHairline;
		B.SetWeight(v);
		
		if ( ld.Color != MTBL_COLOR_UNDEF )
		{
			dwSubstColor = MTblGetExportSubstColor( hWndTbl, ld.Color );
			if ( dwSubstColor != MTBL_COLOR_UNDEF )
				ld.Color = dwSubstColor;
		}

		if ( ld.Color == MTBL_COLOR_UNDEF )
		{
			v = xlAutomatic;
			B.SetColorIndex( v );
		}
		else
		{
			v = long( ld.Color );
			B.SetColor( v );
		}						
	}	
}
*/

void SetFont( Font & F, FONTINFO & fi )
{
	COleVariant v;

	v = fi.sName;
	F.SetName( v );

	v = long( fi.iPointSize );
	F.SetSize( v );

	v = VARIANT_BOOL( fi.bBold );
	F.SetBold( v );

	v = VARIANT_BOOL( fi.bItalic );
	F.SetItalic( v );

	v = long( fi.bUnderline ? xlUnderlineStyleSingle : xlNone );
	F.SetUnderline( v );

	v = VARIANT_BOOL( fi.bStrikeOut );
	F.SetStrikethrough( v );
}

void SetRangelines(HWND hWndTbl, RANGELINE_LIST& RLs)
{
	Border B;
	Borders Bs;
	LPDISPATCH pDisp;
	POSITION pos;
	RANGELINES* prl;

	pos = RLs.GetHeadPosition();
	while (pos)
	{
		prl = &RLs.GetNext(pos);

		pDisp = prl->R.GetBorders();
		Bs.AttachDispatch(pDisp);

		if (prl->dwUse & USE_HORZ_INNER)
		{
			pDisp = Bs.GetItem(xlInsideHorizontal);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldHorzInner);
		}
		if (prl->dwUse & USE_VERT_INNER)
		{
			pDisp = Bs.GetItem(xlInsideVertical);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldVertInner);
		}
		if (prl->dwUse & USE_EDGE_LEFT)
		{
			pDisp = Bs.GetItem(xlEdgeLeft);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldEdgeLeft);
		}
		if (prl->dwUse & USE_EDGE_TOP)
		{
			pDisp = Bs.GetItem(xlEdgeTop);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldEdgeTop);
		}
		if (prl->dwUse & USE_EDGE_RIGHT)
		{
			pDisp = Bs.GetItem(xlEdgeRight);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldEdgeRight);
		}
		if (prl->dwUse & USE_EDGE_BOTTOM)
		{
			pDisp = Bs.GetItem(xlEdgeBottom);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldEdgeBottom);
		}
		if (prl->dwUse & USE_DIAG_UP)
		{
			pDisp = Bs.GetItem(xlDiagonalUp);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldDiagUp);
		}
		if (prl->dwUse & USE_DIAG_DOWN)
		{
			pDisp = Bs.GetItem(xlDiagonalDown);
			B.AttachDispatch(pDisp);
			SetBorder(hWndTbl, B, prl->ldDiagDown);
		}
	}
}


// ExportToExcel
int ExportToExcel( EXCEL_EXP_PARAMS &params )
{
	AFX_MANAGE_STATE( AfxGetStaticModuleState( ) );

	_Application App;
	_Workbook WBook;
	_Worksheet WSheet;
	BOOL bAnyHdrGrp, bAnyHdrGrpWCH, bAttached, bMergeCell, bCellNoColLine, bCellNoRowLine, bCloseInstance, bCloseWorkbook, bColHeaders, bColors, bCurrentPos, bFindSplitRows, bFonts, bFound, bGrid, bHideInstance, bHorz, bMustDeleteFont, bNewInstance, bNewWorkbook, bNewWorksheet,  bNoAutoFitCol, bNoAutoFitRow, bNoMoreRows, bNullsAsSpace, bRowNoColLines, bRowNoRowLine, bRunningInstance, bShowStatus, bSplitRows, bStdFormat, bStringColsAsText, bTemplate, bTemplateLoaded = FALSE, bUseClipboard, bVert;
	Border B;
	Borders Bs;
	COLINFO_ARRAY Cols;
	CArray<COLORREF,COLORREF&> Colors;
	CArray<CString,CString&> PasteStrings;
	CList<CELLMERGE,CELLMERGE&>CMs;
	CList<COLHDRGRP,COLHDRGRP&> CHGs;
	RANGEFORMAT_LIST RFs;
	RANGEFORMAT_LIST RFVs;
	CList<long,long&>TblRows;
	CLSID clsid;
	CMTblItem item;
	COleVariant v, v2, v3;
	CString sClipboardData, sCR( TCHAR(13) ), sCRLF, sEncl( _T("\"") ), sStartCell( params.sExcelStartCell ), sLine, sPaste, sRange, sRange2, sStatus, sStatusFmt, sTab, sText, sTmp, /*sVersion,*/ sWBookFullName;
	CTime tBegin, tEnd;
	CTimeSpan ts;
	CExpTblDlgThread * pDlgThread = NULL;
	CELLMERGE cm;
	CELLMERGE * pcm;
	COLHDRGRP chg;
	COLHDRGRP * pchg;
	DWORD dwBackColorCell = 0, dwBackColorTbl = 0, dwHString, dwSubstColor = 0, dwTextColorCell = 0, dwTextColorTbl = 0;
	Font F;
	FONTINFO fi, fiTbl;
	HFONT hFont;
	HSTRING hsText = 0, hsFormat = 0, hsLanguageFile = 0, hsValue = 0;
	HWND hWndCol, hWndTreeCol = NULL;
	int iCHGs = 0, iCMs = 0, iCol, iCol2, iColFrom, iColTo, iCols, iDataRows, iHdrRows, iID, iIndent, iLastCol, iLastHdrGrpCol, iLastRow, iLines, iMaxTries, iNodeSize = 0, iPos, iPos2, iPasteRow, iPasteLastRow, iRealLockedCols, iRet = 1, iRFs = 0, iRFVs = 0, iRow, iRow2, iRowFrom, iRowTo, iRows, iRowDataBegin, iRowLevel, iStartCol, iStartRow, iTries, /*iVersion = 0,*/ j;
	Interior I;
	MTBLLINEDEF ColLinesTbl, RowLinesTbl, HdrLinesTbl;
	MTBLLINEDEF HorzLineDef, HorzLineDef2, VertLineDef, VertLineDef2;
	long lCalculationOrig, lCellJustify, lCellVAlign, lColWidth, lIndentLevel, lLen, lRow, lRow2, lRowHeight, lSheets, lTextLen;
	LPDISPATCH pDisp = NULL, pFont, pInterior, pRange, pWBook, pWBooks, pWSheet, pWSheets;
	LPMTBLMERGECELL pmc = NULL;
	LPTSTR lps;
	LPUNKNOWN pUnk;
	LPVOID pHdrGrp, pMergeCells = NULL;
	POSITION pos, pos2;
	Range R;
	RANGEFORMAT rf, rfv;
	RANGEFORMAT *prf, *prfv;
	TCHAR cPwd;
	VARIANT vOpt, vIndentLevel;
	Workbooks WBooks;
	Worksheets WSheets;

	const int MAX_UNION = 64;
	const int PASTE_AT = 1000;

	try
	{
		// Subclassed?
		if ( !MTblIsSubClassed( params.hWndTbl ) )
			return MTE_ERR_NOT_SUBCLASSED;

		VMakeOptional( vOpt );

		tBegin = CTime::GetCurrentTime( );

		// Tree-Definition ermitteln
		MTblQueryTree( params.hWndTbl, &hWndTreeCol, &iNodeSize, &iIndent );

		// Passwort-Zeichen ermitteln
		cPwd = MTblGetPasswordCharVal( params.hWndTbl );

		// Sprache
		if ( !MTblGetLanguageFile( params.iLanguage, &hsLanguageFile ) )
			return MTE_ERR_INVALID_LANG;

		if ( hsLanguageFile != 0 )
			Ctd.HStringUnRef( hsLanguageFile );

		// Template
		bTemplate = params.sTemplateFile.GetLength() > 0;

		// Flags ermitteln
		bNewInstance = params.dwExcelFlags & MTE_EXCEL_NEW_INSTANCE;
		bRunningInstance = params.dwExcelFlags & MTE_EXCEL_RUNNING_INSTANCE;
		bCloseInstance = params.dwExcelFlags & MTE_EXCEL_CLOSE_INSTANCE;
		bHideInstance = params.dwExcelFlags & MTE_EXCEL_HIDE_INSTANCE;
		bNewWorkbook = ( params.dwExcelFlags & MTE_EXCEL_NEW_WORKBOOK );
		bCloseWorkbook = ( params.dwExcelFlags & MTE_EXCEL_CLOSE_WORKBOOK );
		bNewWorksheet = params.dwExcelFlags & MTE_EXCEL_NEW_WORKSHEET;
		bCurrentPos = params.dwExcelFlags & MTE_EXCEL_CURRENT_POS;
		bStringColsAsText = params.dwExcelFlags & MTE_EXCEL_STRING_COLS_AS_TEXT;
		bNullsAsSpace = params.dwExcelFlags & MTE_EXCEL_NULLS_AS_SPACE;
		bUseClipboard = !(params.dwExcelFlags & MTE_EXCEL_NO_CLIPBOARD);
		bNoAutoFitCol = params.dwExcelFlags & MTE_EXCEL_NO_AUTO_FIT_COL;
		bNoAutoFitRow = params.dwExcelFlags & MTE_EXCEL_NO_AUTO_FIT_ROW;

		if ( !( bNewInstance || bRunningInstance ) )
		{
			// Wenn keine Instanz läuft, lfd. Instanz verwenden oder neue Instanz erzeugen
			bRunningInstance = TRUE;
			bNewInstance = TRUE;
		}

		bColHeaders = params.dwExpFlags & MTE_COL_HEADERS;
		bColors = params.dwExpFlags & MTE_COLORS;
		bFonts = params.dwExpFlags & MTE_FONTS;
		bGrid = params.dwExpFlags & MTE_GRID;
		bShowStatus = params.dwExpFlags & MTE_SHOW_STATUS;
		bSplitRows = params.dwExpFlags & MTE_SPLIT_ROWS;
		if ( bSplitRows )
		{
			BOOL bDragAdjust;
			int iSplitRows = 0;
			if ( !( Ctd.TblQuerySplitWindow( params.hWndTbl, &iSplitRows, &bDragAdjust ) && iSplitRows > 0 ) )
				bSplitRows = FALSE;
		}

		// Aktive Excel-Instanz ermitteln
		if ( bRunningInstance )
		{
			// Klassen-ID ermitteln
			if ( CLSIDFromProgID( L"Excel.Application", &clsid ) == S_OK )
			{
				// Laufende Instanz ermitteln
				if ( GetActiveObject( clsid, NULL, &pUnk ) == S_OK )
				{
					// Dispatch-Interface ermitteln
					if ( pUnk->QueryInterface( IID_IDispatch, (void**)&pDisp ) != S_OK )
					{
						pUnk->Release( );
						return MTE_ERR_RUNNING_INSTANCE;
					}
					pUnk->Release( );
				}
			}

			// Mit Dispatch verbinden
			if ( pDisp )
			{
				App.AttachDispatch( pDisp );
			}
		}

		// Dispatch erzeugen
		if ( !pDisp )
		{
			if ( bNewInstance )
			{
				if ( !App.CreateDispatch( _T("Excel.Application") ) )
					return MTE_ERR_NEW_INSTANCE;

				// Wir brauchen auf jeden Fall ein neues Workbook
				//bNewWorkbook = TRUE;
			}
			else
				return MTE_ERR_RUNNING_INSTANCE;
		}

		// Sichtbarkeit Instanz
		App.SetVisible( !bHideInstance );

		// Dialog-Thread starten
		if ( bShowStatus )
		{
			pDlgThread = (CExpTblDlgThread*)AfxBeginThread( RUNTIME_CLASS( CExpTblDlgThread ) );
			if ( !pDlgThread )
				return MTE_ERR_STATUS;

			// Sprache setzen
			SET_STATUS_LNG( params.iLanguage );
		}

		// Status
		sStatus = _T("");
		GetUIStr( params.iLanguage, _T("StatDlg.Text.Init"), sStatus );
		SET_STATUS( sStatus );

		// Sanduhr
		App.SetCursor( xlWait );

		// Workbooks-Aufzählung ermitteln
		pWBooks = App.GetWorkbooks( );
		WBooks.AttachDispatch( pWBooks );

		// Ggfs. Template laden
		if ( bTemplate )
		{
			MAKE_STATUS_DLG_TOPMOST( FALSE );

			// Status
			//SET_STATUS( params.iLanguage == MTE_LNG_GERMAN ? _T("Vorlagendatei wird geöffnet...") : _T("Opening template file...") );
			sStatus = _T("");
			GetUIStr( params.iLanguage, _T("StatDlg.Text.OpenTempl"), sStatus );
			SET_STATUS( sStatus );

			try
			{
				if ( bNewWorkbook )
				{
					v = params.sTemplateFile;
					pWBook = WBooks.Add( v );
				}
				else
				{
					pWBook = WBooks.Open( params.sTemplateFile, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt );
				}
				
				bTemplateLoaded = TRUE;
			}
			catch( CException * pE )
			{
				pE->Delete( );
			}

			MAKE_STATUS_DLG_TOPMOST( TRUE );

			if ( bTemplateLoaded )
			{
				bNewWorkbook = FALSE;
			}
		}

		// Neues Workbook erstellen, falls keins da ist
		if ( WBooks.GetCount( ) < 1 )
		{			
			bNewWorkbook = TRUE;
		}

		// Workbook erstellen bzw. ermitteln
		if ( bNewWorkbook )
		{
			lSheets = App.GetSheetsInNewWorkbook( );
			App.SetSheetsInNewWorkbook( 1 );
			pWBook = WBooks.Add( );
			App.SetSheetsInNewWorkbook( lSheets );

			// Kein neues Worksheet, da ja jetzt eins da sein sollte
			bNewWorksheet = FALSE;
		}
		else
		{
			pWBook = App.GetActiveWorkbook( );
		}
		WBook.AttachDispatch( pWBook );

		// Worksheets ermitteln
		pWSheets = App.GetWorksheets( );
		WSheets.AttachDispatch( pWSheets );
		if ( WSheets.GetCount( ) < 1 )
		{
			// Neues Worksheet erstellen, da keins da ist
			bNewWorksheet = TRUE;
		}

		// Worksheet erstellen bzw. ermitteln
		if ( bNewWorksheet )
		{
			pWSheet = WSheets.Add( vOpt, vOpt, vOpt, vOpt );
		}
		else
		{
			pWSheet = App.GetActiveSheet( );
		}
		WSheet.AttachDispatch( pWSheet );

		// Max. Spalten/Zeilen
		pRange = WSheet.GetColumns( );
		R.AttachDispatch( pRange );
		const long MAX_COLS = R.GetCount( );

		pRange = WSheet.GetRows( );
		R.AttachDispatch( pRange );
		const long MAX_ROWS = R.GetCount( );

		// Worksheet aktivieren
		WSheet.Activate( );

		// Ggfs. Name setzen
		if ( params.sSheetName.GetLength() > 0 )
		{
			try
			{
				WSheet.SetName( params.sSheetName );
			}
			catch( CException * pE )
			{
				pE->Delete( );
			}
		}

		// Ggf. automatische Berechnung abschalten
		lCalculationOrig =  App.GetCalculation( );
		if ( lCalculationOrig != xlManual )
			App.SetCalculation( xlManual );

		// Startposition selektieren
		if ( !bCurrentPos )
		{
			if ( sStartCell.IsEmpty( ) )
				AttachRange( WSheet, R, 1, 1 );
			else
				AttachRange( WSheet, R, sStartCell );
			R.Select( );
		}

		// Startspalte/zeile ermitteln
		GetStart( App, iStartCol, iStartRow );

		// Tatsächliche Anzahl gelockter Spalten ermitteln
		iRealLockedCols = MTblQueryLockedColumns( params.hWndTbl );

		// Spaltenlinien der Tabelle ermitteln
		ColLinesTbl.Style = MTLS_DOT;
		ColLinesTbl.Color = MTBL_COLOR_UNDEF;
		ColLinesTbl.Thickness = 1;
		MTblQueryColLines(params.hWndTbl, &ColLinesTbl.Style, &ColLinesTbl.Color);

		// Zeilenlinien der Tabelle ermitteln
		RowLinesTbl.Style = MTLS_DOT;
		RowLinesTbl.Color = MTBL_COLOR_UNDEF;
		RowLinesTbl.Thickness = 1;
		MTblQueryRowLines(params.hWndTbl, &RowLinesTbl.Style, &RowLinesTbl.Color);

		// Headerlinien der Tabelle ermitteln
		HdrLinesTbl.Style = MTLS_SOLID;
		HdrLinesTbl.Color = MTBL_COLOR_UNDEF;
		HdrLinesTbl.Thickness = 1;

		// Spalten ermitteln
		//SET_STATUS( params.iLanguage == MTE_LNG_GERMAN ? _T("Zu exportierende Spalten werden ermittelt...") : _T("Columns to export are queried...") );
		sStatus = _T("");
		GetUIStr( params.iLanguage, _T("StatDlg.Text.QueryCols"), sStatus );
		SET_STATUS( sStatus );

		Cols.SetSize( MAX_COLS );
		bAnyHdrGrp = bAnyHdrGrpWCH = FALSE;
		iCols = 0;
		iPos = 1;
		while ( hWndCol = Ctd.TblGetColumnWindow( params.hWndTbl, iPos, COL_GetPos ) )
		{
			if ( ( params.dwColFlagsOn ? Ctd.TblQueryColumnFlags( hWndCol, params.dwColFlagsOn ) : TRUE ) && ( params.dwColFlagsOff ? !Ctd.TblQueryColumnFlags( hWndCol, params.dwColFlagsOff ) : TRUE ) )
			{
				if ( ( iCols + iStartCol - 1 ) >= MAX_COLS )
					break;

				// Handle
				Cols[iCols].hWnd = hWndCol;

				// Position
				Cols[iCols].iPos = Ctd.TblQueryColumnPos( hWndCol );

				// Gelockt?
				Cols[iCols].bLocked = Cols[iCols].iPos <= iRealLockedCols;

				// Tree-Spalte?
				Cols[iCols].bTreeCol = ( hWndCol == hWndTreeCol );

				// Horizontale Ausrichtung ermitteln
				if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_LEFT ) )
					Cols[iCols].lJustify = xlLeft;
				else if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_CENTER ) )
					Cols[iCols].lJustify = xlCenter;
				else if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_RIGHT ) )
					Cols[iCols].lJustify = xlRight;
				else if ( Ctd.TblQueryColumnFlags( hWndCol, COL_RightJustify ) )
					Cols[iCols].lJustify = xlRight;
				else if ( Ctd.TblQueryColumnFlags( hWndCol, COL_CenterJustify ) )
					Cols[iCols].lJustify = xlCenter;
				else
					Cols[iCols].lJustify = xlLeft;

				// Horizontale Ausrichtung Header ermitteln
				if ( MTblQueryColumnHdrFlags( hWndCol, MTBL_COLHDR_FLAG_TXTALIGN_LEFT ) )
					Cols[iCols].lHdrJustify = xlLeft;
				else if ( MTblQueryColumnHdrFlags( hWndCol, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT ) )
					Cols[iCols].lHdrJustify = xlRight;
				else
					Cols[iCols].lHdrJustify = xlCenter;

				// Vertikale Ausrichtung ermitteln
				if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_VCENTER ) )
					Cols[iCols].lVAlign = xlCenter;
				else if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_BOTTOM ) )
					Cols[iCols].lVAlign = xlBottom;
				else
					Cols[iCols].lVAlign = xlTop;

				// Invisible?
				Cols[iCols].bInvisible = Ctd.TblQueryColumnFlags( hWndCol, COL_Invisible );

				// Datentyp
				Cols[iCols].lDataType = Ctd.GetDataType( hWndCol );

				// Zelltyp
				Ctd.TblQueryColumnCellType( hWndCol, &Cols[iCols].iCellType );

				// Header-Gruppe
				Cols[iCols].pHdrGrp = MTblGetColHdrGrp( hWndCol );
				if ( Cols[iCols].pHdrGrp )
				{
					bAnyHdrGrp = TRUE;
					if ( !MTblQueryColHdrGrpFlags( params.hWndTbl, Cols[iCols].pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS ) )
						bAnyHdrGrpWCH = TRUE;
				}

				// Keine Zeilenlinien?
				if ( hWndCol == hWndTreeCol )
					Cols[iCols].bNoRowLines = MTblQueryTreeFlags( params.hWndTbl, MTBL_TREE_FLAG_NO_ROWLINES );
				else
					Cols[iCols].bNoRowLines = MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_NO_ROWLINES );

				// Keine Spaltenlinie?
				Cols[iCols].bNoColLine = MTblQueryColumnFlags(hWndCol, MTBL_COL_FLAG_NO_COLLINE) || Cols[iCols].VertLineDef.Style == MTLS_INVISIBLE;

				// Liniendefinitionen
				item.DefineAsCol(hWndCol);
				MTblQueryItemLineC(&item, &Cols[iCols].HorzLineDef, MTLT_HORZ);
				MTblQueryItemLineC(&item, &Cols[iCols].VertLineDef, MTLT_VERT);

				// Zahlenformat ermitteln
				dwHString = (DWORD)SendMessage(params.hWndTbl, MTM_QueryExcelNumberFormat, (WPARAM)hWndCol, TBL_Error);
				if (dwHString)
				{
					hsFormat = Ctd.NumberToHString(dwHString);
					if (hsFormat)
					{
						lps = Ctd.StringGetBuffer(hsFormat, &lLen);
						if (lps)
						{
							Cols[iCols].sFormat = lps;
							Ctd.HStrUnlock(hsFormat);
						}
							
					}
				}

				if (Cols[iCols].sFormat.IsEmpty())
				{
					if (MTblQueryColumnFlags(hWndCol, MTBL_COL_FLAG_EXPORT_AS_TEXT) || (bStringColsAsText && (Cols[iCols].lDataType == DT_String || Cols[iCols].lDataType == DT_LongString)))
						Cols[iCols].sFormat = _T("@");
				}

				// Array für Zellentexte vorbereiten
				if ( !bUseClipboard )
					Cols[iPos].CellTexts.SetSize( PASTE_AT );

				++iCols;
			}
			++iPos;

			// Abbruch?
			if ( IS_CANCELED )
			{
				iRet = MTE_ERR_CANCELLED;
				goto Terminate;
			}
		}

		if ( iCols == 0 )
		{
			goto Terminate;
		}

		// Letzte Spalte setzen
		iLastCol = iStartCol + iCols - 1;

		// Merge-Zellen ermitteln
		pMergeCells = MTblGetMergeCells( params.hWndTbl, params.wRowFlagsOn, params.wRowFlagsOff, params.dwColFlagsOn, params.dwColFlagsOff );

		// Farben der Tabelle ermitteln
		if ( bColors )
		{
			dwTextColorTbl = Ctd.ColorGet( params.hWndTbl, COLOR_IndexWindowText );
			dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwTextColorTbl );
			if ( dwSubstColor != MTBL_COLOR_UNDEF )
				dwTextColorTbl = dwSubstColor;

			dwBackColorTbl = Ctd.ColorGet( params.hWndTbl, COLOR_IndexWindow );
			dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwBackColorTbl );
			if ( dwSubstColor != MTBL_COLOR_UNDEF )
				dwBackColorTbl = dwSubstColor;
		}

		// Font der Tabelle ermitteln
		if ( bFonts )
		{
			GetFontInfo( params.hWndTbl, HFONT( SendMessage( params.hWndTbl, WM_GETFONT, 0, 0 ) ), fiTbl );
		}

		// Spaltenköpfe exportieren
		sStatus = _T("");
		GetUIStr( params.iLanguage, _T("StatDlg.Text.ExpColHdrs"), sStatus );
		SET_STATUS( sStatus );

		iRow = iStartRow - 1;
		iHdrRows = 0;

		if ( bColHeaders )
		{
			// Zeilenzähler erhöhen
			++iRow;
			if ( bAnyHdrGrpWCH )
				++iRow;

			// Formatieren
			AttachRange( WSheet, R, iStartCol, iStartRow, iLastCol, iRow );

			// Alles löschen
			R.ClearContents( );

			// Datentyp = Text
			v = _T("@");
			R.SetNumberFormatLocal( v );

			// Ausrichtung: vertikal = oben
			v = xlTop;
			R.SetVerticalAlignment(v);

			pHdrGrp = NULL;
			for( iPos = 0; iPos < iCols; ++iPos )
			{
				hWndCol = Cols[iPos].hWnd;
				iCol = iStartCol + iPos;

				// Header-Gruppen
				if ( bAnyHdrGrp )
				{
					if ( !Cols[iPos].pHdrGrp )
					{
						// Akt. Gruppe setzen
						pHdrGrp = NULL;

						// Keine Gruppe, Zellen verbinden
						AttachRange( WSheet, R, iCol, iStartRow, iCol, iRow );
						//R.ClearContents( );
						v = 1L;
						R.SetMergeCells( v );
					}
					else if ( Cols[iPos].pHdrGrp && Cols[iPos].pHdrGrp != pHdrGrp )
					{
						// Akt. Gruppe setzen
						pHdrGrp = Cols[iPos].pHdrGrp;

						// Letzte Gruppenspalte ermitteln
						iLastHdrGrpCol = iCol;
						for ( j = iPos + 1; j < iCols; ++j )
						{
							if ( Cols[j].pHdrGrp == pHdrGrp )
								++iLastHdrGrpCol;
							else
								break;
						}

						// Range der Gruppe ermitteln
						if ( MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS ) )
							AttachRange( WSheet, R, iCol, iStartRow, iLastHdrGrpCol, iRow );
						else
							AttachRange( WSheet, R, iCol, iStartRow, iLastHdrGrpCol, iStartRow );

						// Obere Zellen der Gruppe verbinden
						if ( iLastHdrGrpCol > iCol )
						{
							//R.ClearContents( );
							v = 1L;
							R.SetMergeCells( v );
						}

						// Gruppe merken
						chg.iFromCol = iCol;
						chg.iFromLine = iStartRow;
						chg.iToCol = iLastHdrGrpCol;
						chg.iToLine = iRow;
						chg.bNoInnerVLines = MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOINNERVERTLINES );
						chg.bNoInnerHLine = MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOINNERHORZLINE );
						chg.bNoColHeaders = MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS );
						chg.bTxtAlignLeft = MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_TXTALIGN_LEFT );
						chg.bTxtAlignRight = MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_TXTALIGN_RIGHT );

						item.DefineAsColHdrGrp(params.hWndTbl, pHdrGrp);
						MTblQueryEffItemLineC(item, chg.HorzLineDef, MTLT_HORZ);
						MTblQueryEffItemLineC(item, chg.VertLineDef, MTLT_VERT);

						CHGs.AddTail( chg );
						++iCHGs;

						// Text ermitteln
						MTblGetColHdrGrpText( params.hWndTbl, pHdrGrp, &hsText );
						sText = Ctd.StringGetBuffer( hsText, &lLen );
						sText.Replace( sCR, _T("") );
						Ctd.HStrUnlock(hsText);

						// Horizontale Textausrichtung setzen
						if (chg.bTxtAlignLeft)
							v = xlLeft;
						else if (chg.bTxtAlignRight)
							v = xlRight;
						else
							v = xlCenter;
						R.SetHorizontalAlignment(v);

						// Evtl. Farben setzen
						if ( bColors )
						{
							// Hintergrundfarbe = Hintergrundfarbe der Gruppe
							pInterior = R.GetInterior( );
							I.AttachDispatch( pInterior );
							if ( MTblGetEffColHdrGrpBackColor( params.hWndTbl, pHdrGrp, &dwBackColorCell ) )
							{
								dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwBackColorCell );
								if ( dwSubstColor != MTBL_COLOR_UNDEF )
									dwBackColorCell = dwSubstColor;

								v = long( dwBackColorCell );
								I.SetColor( v );
							}
							v = xlSolid;
							I.SetPattern( v );

							// Textfarbe = Textfarbe der Gruppe
							pFont = R.GetFont( );
							F.AttachDispatch( pFont );
							if ( MTblGetEffColHdrGrpTextColor( params.hWndTbl, pHdrGrp, &dwTextColorCell ) )
							{
								dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwTextColorCell );
								if ( dwSubstColor != MTBL_COLOR_UNDEF )
									dwTextColorCell = dwSubstColor;

								v = long( dwTextColorCell );
								F.SetColor( v );
							}
						}

						// Evtl. Font setzen
						if ( bFonts )
						{
							hFont = NULL;
							
							if ( hFont = MTblGetEffColHdrGrpFontHandle( params.hWndTbl, pHdrGrp, &bMustDeleteFont ) )
							{
								GetFontInfo( params.hWndTbl, hFont, fi );
								if ( bMustDeleteFont )
									DeleteObject( hFont );
							}
							else
								fi = fiTbl;

							pFont = R.GetFont( );
							F.AttachDispatch( pFont );

							SetFont( F, fi );
						}

						// Text setzen
						v = sText;
						R.SetValue( v );

						// Range für Spaltenkopf ermitteln
						AttachRange( WSheet, R, iCol, iRow );
					}
					else
						// Range für Spaltenkopf ermitteln
						AttachRange( WSheet, R, iCol, iRow );
				}
				else
					// Range für Spaltenkopf ermitteln
					AttachRange( WSheet, R, iCol, iRow );

				if ( pHdrGrp && MTblQueryColHdrGrpFlags( params.hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS ) )
					continue;

				// Horizontale Textausrichtung setzen
				v = Cols[iPos].lHdrJustify;
				R.SetHorizontalAlignment(v);

				// Evtl. Farben setzen
				if ( bColors )
				{
					// Hintergrundfarbe = Hintergrundfarbe des Spaltenkopfes
					pInterior = R.GetInterior( );
					I.AttachDispatch( pInterior );
					if ( MTblGetEffColumnHdrBackColor( Cols[iPos].hWnd, &dwBackColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwBackColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwBackColorCell = dwSubstColor;

						v = long( dwBackColorCell );
						I.SetColor( v );
					}
					v = xlSolid;
					I.SetPattern( v );

					// Textfarbe = Textfarbe des Spaltenkopfes
					pFont = R.GetFont( );
					F.AttachDispatch( pFont );
					if ( MTblGetEffColumnHdrTextColor( Cols[iPos].hWnd, &dwTextColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwTextColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwTextColorCell = dwSubstColor;

						v = long( dwTextColorCell );
						F.SetColor( v );
					}
				}

				// Evtl. Font setzen
				if ( bFonts )
				{
					hFont = NULL;
					
					if ( hFont = MTblGetEffColumnHdrFontHandle( Cols[iPos].hWnd, &bMustDeleteFont ) )
					{
						GetFontInfo( params.hWndTbl, hFont, fi );
						if ( bMustDeleteFont )
							DeleteObject( hFont );
					}
					else
						fi = fiTbl;

					pFont = R.GetFont( );
					F.AttachDispatch( pFont );

					SetFont( F, fi );
				}

				// Text ermitteln
				MTblGetColumnTitle( hWndCol, &hsText );
				sText = Ctd.StringGetBuffer( hsText, &lLen );
				sText.Replace( sCR, _T("") );
				Ctd.HStrUnlock(hsText);

				// Text setzen
				v = sText;
				R.SetValue( v );

				// Abbruch?
				if ( IS_CANCELED )
				{
					iRet = MTE_ERR_CANCELLED;
					goto Terminate;
				}
			}

			// Header-Zeilenzähler erhöhen
			++iHdrRows;
			if ( bAnyHdrGrpWCH )
				++iHdrRows;
		}

		// Zellen exportieren
		sStatus = _T("");
		GetUIStr( params.iLanguage, _T("StatDlg.Text.GetData"), sStatus );
		SET_STATUS( sStatus );

		// Inhalt der Zwischenablage sichern
		if ( bUseClipboard )
		{
			sClipboardData = _T("");
			SaveClipboard( sClipboardData, params.hWndTbl );
		}

		sCRLF = TCHAR(13);
		sCRLF += TCHAR(10);
		sTab = _T("	");
		++iRow;
		iRowDataBegin = iRow;
		iPasteRow = iRow;
		iDataRows = 0;
		PasteStrings.SetSize( iCols );
		iLines = 0;
		lRow = -1;
		bNoMoreRows = FALSE;
		bFindSplitRows = FALSE;
		for( ; ; )
		{
			if ( !Ctd.TblFindNextRow( params.hWndTbl, &lRow, 0, 0 ) || ( bFindSplitRows ? lRow >= 0 : FALSE ) )
			{
				if ( bSplitRows )
				{
					lRow = TBL_MinRow;
					bFindSplitRows = TRUE;
					bSplitRows = FALSE;
					continue;
				}
				else
				{
					bNoMoreRows = TRUE;
					goto Paste;
				}
			}

			if ( iRow > MAX_ROWS )
			{
				bNoMoreRows = TRUE;
				goto Paste;
			}

			if ( params.wRowFlagsOn && !Ctd.TblQueryRowFlags( params.hWndTbl, lRow, params.wRowFlagsOn ) )
				continue;
			if ( params.wRowFlagsOff && Ctd.TblQueryRowFlags( params.hWndTbl, lRow, params.wRowFlagsOff ) )
				continue;

			TblRows.AddTail( lRow );

			Ctd.TblSetContextEx( params.hWndTbl, lRow );

			for ( iPos = 0; iPos < iCols; ++iPos )
			{
				hWndCol = Cols[iPos].hWnd;
				iCol = iStartCol + iPos;

				// Merge-Zelle?
				bMergeCell = FALSE;
				pmc = (LPMTBLMERGECELL)MTblFindMergeCell( hWndCol, lRow, FMC_MASTER, pMergeCells );
				if ( pmc )
				{
					// Merge-Bereich setzen
					cm.iFromCol = iCol;
					cm.iToCol = iCol + pmc->iMergedCols;
					cm.iFromLine = iRow;
					cm.iToLine = iRow + pmc->iMergedRows;

					bMergeCell = TRUE;
					CMs.AddTail( cm );
					++iCMs;
				}
				
				// Text ermitteln
				hsValue = 0;
				
				// Nachricht zur Ermittlung des Wertes an Tabelle senden
				dwHString = (DWORD)SendMessage( params.hWndTbl, MTM_QueryExcelValue, (WPARAM)hWndCol, lRow );
				if ( dwHString )
				{
					hsValue = Ctd.NumberToHString( dwHString );
				}
				
				// Text initialisieren
				sText.Empty( );
				lTextLen = 0;
				
				// Endgültigen Text ermitteln
				if ( hsValue )
				{
					lps = Ctd.StringGetBuffer( hsValue, &lLen );
					if ( lps )
					{
						sText = lps;
						lTextLen = sText.GetLength( );
						Ctd.HStrUnlock(hsValue);
					}
				}
				else if ( !MTblQueryCellFlags( hWndCol, lRow, MTBL_CELL_FLAG_HIDDEN_TEXT ) )
				{
					if ( Cols[iPos].lDataType == DT_Number && !Cols[iPos].sFormat.IsEmpty( ) )
					{
						Ctd.FmtFieldToStr( hWndCol, &hsText, FALSE );
					}
					else
					{
						iID = Ctd.TblQueryColumnID( hWndCol );
						Ctd.TblGetColumnText( params.hWndTbl, iID, &hsText );
					}
					sText = Ctd.StringGetBuffer( hsText, &lLen );
					lTextLen = sText.GetLength( );
					Ctd.HStrUnlock(hsText);

					// Evtl. Zeichen durch Passwort-Zeichen ersetzen
					if ( lTextLen > 0 && Cols[iPos].bInvisible && Cols[iPos].iCellType != COL_CellType_CheckBox )
					{
						for ( int i = 0; i < lTextLen; ++i )
						{
							sText.SetAt( i, cPwd );
						}
					}
				}

				if ( !sText.IsEmpty( ) )
				{
					// Carriage Return entfernen
					sText.Replace( sCR, _T("") );
					lTextLen = sText.GetLength( );

					// Evtl. Anführungszeichen verdoppeln und mit Anführungszeichen umschließen
					if ( bUseClipboard )
					{						
						sText.Replace( _T("\""), _T("\"\"") );
						lTextLen = sText.GetLength( );

						sText = sEncl + sText + sEncl;
					}
					
				}
				else if ( bNullsAsSpace )
				{
					// NULL durch Leerzeichen ersetzen
					sText = _T(" ");
					lTextLen = 1;
				}

				// Text setzen
				if ( bUseClipboard )
				{
					// Pufferlänge des Paste-Strings setzen
					if ( iLines == 0 )
					{
						lps = PasteStrings[iPos].GetBufferSetLength( PASTE_AT * 254 );
						PasteStrings[iPos].ReleaseBuffer( 0 );
					}

					// Paste-String der Spalte erweitern
					if ( iLines > 0 )
						sText = sCRLF + sText;
					PasteStrings[iPos] += sText;
				}
				else
				{
					if ( Cols[iPos].CellTexts.GetUpperBound( ) < iLines )
						Cols[iPos].CellTexts.Add( sText );
					else
						Cols[iPos].CellTexts[iLines] = sText;
				}

				// Formate ermitteln, welche VOR Einfügen der Zellentexte gesetzt werden
				rfv.dwUse = 0;
				rfv.dwTextColor = dwTextColorTbl;
				rfv.dwBackColor = dwBackColorTbl;
				rfv.sNumberFormat.Empty( );
				rfv.lIndentLevel = 0;
				rfv.fi = fiTbl;
				rfv.lJustify = xlNone;
				rfv.lVAlign = xlNone;
				rfv.iUnionCount = 0;

				// Zahlen-Format
				hsFormat = SendMessage(params.hWndTbl, MTM_QueryExcelNumberFormat, (WPARAM)hWndCol, (LPARAM)lRow);
				if (hsFormat)
				{
					lps = Ctd.StringGetBuffer(hsFormat, &lLen);
					if (lps)
					{
						rfv.dwUse |= USE_NUMBERFORMAT;
						rfv.sNumberFormat = lps;
						Ctd.HStrUnlock(hsFormat);
					}
				}

				// Horizontale Textausrichtung
				lCellJustify = MTblGetEffCellTextJustify(hWndCol, lRow);
				if (lCellJustify == DT_LEFT)
					lCellJustify = xlLeft;
				else if (lCellJustify == DT_CENTER)
					lCellJustify = xlCenter;
				else if (lCellJustify == DT_RIGHT)
					lCellJustify = xlRight;
				else
					lCellJustify = xlNone;

				if (lCellJustify != xlNone && lCellJustify != Cols[iPos].lJustify)
				{
					rfv.dwUse |= USE_JUSTIFY;
					rfv.lJustify = lCellJustify;
				}				

				// Vertikale Textausrichtung
				lCellVAlign = MTblGetEffCellTextVAlign(hWndCol, lRow);
				if (lCellVAlign == DT_TOP)
					lCellVAlign = xlTop;
				else if (lCellVAlign == DT_VCENTER)
					lCellVAlign = xlCenter;
				else if (lCellVAlign == DT_BOTTOM)
					lCellVAlign = xlBottom;
				else
					lCellVAlign = xlNone;

				if (lCellVAlign != xlNone && lCellVAlign != Cols[iPos].lVAlign)
				{
					rfv.dwUse |= USE_VALIGN;
					rfv.lVAlign = lCellVAlign;
				}

				// Formatierungsbereich erstellen bzw. zu bestehendem Bereich hinzufügen
				if ( rfv.dwUse )
				{
					bFound = FALSE;
					if ( !RFVs.IsEmpty( ) )
					{
						pos = RFVs.GetHeadPosition( );
						while ( pos )
						{
							prfv = &RFVs.GetNext( pos );

							if ( rfv.dwUse == prfv->dwUse &&
								 rfv.dwTextColor == prfv->dwTextColor &&
								 rfv.dwBackColor == prfv->dwBackColor &&
								 rfv.sNumberFormat == prfv->sNumberFormat &&
								 rfv.lIndentLevel == prfv->lIndentLevel &&
								 rfv.fi == prfv->fi &&
								 rfv.lJustify == prfv->lJustify && 
								 rfv.lVAlign == prfv->lVAlign &&
								 prfv->iUnionCount < MAX_UNION )
							{
								AttachRange( WSheet, rfv.R, iCol, iRow );
								pRange = App.Union( prfv->R.m_lpDispatch, rfv.R.m_lpDispatch );
								prfv->R.AttachDispatch( pRange );
								++prfv->iUnionCount;
								bFound = TRUE;
								break;
							}
						}
					}

					if ( !bFound )
					{
						AttachRange( WSheet, rfv.R, iCol, iRow );
						RFVs.AddTail( rfv );
						++iRFVs;
					}
				}

				
				// Formate ermitteln, welche NACH Einfügen der Zellentexte gesetzt werden
				rf.Init();

				// Farben
				if ( bColors )
				{
					// Text
					if ( MTblGetEffCellTextColor( hWndCol, lRow, &dwTextColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwTextColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwTextColorCell = dwSubstColor;

						if ( dwTextColorCell != dwTextColorTbl )
						{
							rf.dwUse |= USE_TEXT_COLOR;
							rf.dwTextColor = dwTextColorCell;
						}
					}

					// Hintergrund
					if ( MTblGetEffCellBackColor( hWndCol, lRow, &dwBackColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( params.hWndTbl, dwBackColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwBackColorCell = dwSubstColor;

						if ( dwBackColorCell != dwBackColorTbl )
						{
							rf.dwUse |= USE_BACK_COLOR;
							rf.dwBackColor = dwBackColorCell;
						}
					}
				}

				// Font
				if ( bFonts )
				{
					if ( hFont = MTblGetEffCellFontHandle( hWndCol, lRow, &bMustDeleteFont ) )
					{
						GetFontInfo( params.hWndTbl, hFont, fi );
						if ( fi != fiTbl )
						{
							rf.dwUse |= USE_FONT;
							rf.fi = fi;
						}
						if ( bMustDeleteFont )
							DeleteObject( hFont );
					}
				}

				// Einrückung
				if ( Cols[iPos].bTreeCol )
				{
					if ( MTblQueryTreeFlags( params.hWndTbl, MTBL_TREE_FLAG_FLAT_STRUCT ) )
						iRowLevel = 0;
					else
						iRowLevel = MTblGetRowLevel( params.hWndTbl, lRow );

					if ( iRowLevel > 0 )
					{
						rf.dwUse |= USE_INDENT_LEVEL;
						rf.lIndentLevel = min( iRowLevel, MAX_INDENT_LEVEL );
					}
				}
				else
				{
					if ( MTblGetCellIndentLevel( hWndCol, lRow, &lIndentLevel ) )
					{
						if ( lIndentLevel > 0 )
						{
							rf.dwUse |= USE_INDENT_LEVEL;
							rf.lIndentLevel = min( lIndentLevel, MAX_INDENT_LEVEL );
						}
					}
				}

				// Formatierungsbereich erstellen bzw. zu bestehendem Bereich hinzufügen
				if ( rf.dwUse )
				{
					bFound = FALSE;
					if ( !RFs.IsEmpty( ) )
					{
						pos = RFs.GetHeadPosition( );
						while ( pos )
						{
							prf = &RFs.GetNext( pos );

							if ( rf == *prf && prf->iUnionCount < MAX_UNION )
							{
								AttachRange( WSheet, rf.R, iCol, iRow );
								pRange = App.Union( prf->R.m_lpDispatch, rf.R.m_lpDispatch );
								prf->R.AttachDispatch( pRange );
								++prf->iUnionCount;
								bFound = TRUE;
								break;
							}
						}
					}

					if ( !bFound )
					{
						AttachRange( WSheet, rf.R, iCol, iRow );
						RFs.AddTail( rf );
						++iRFs;
					}
				}				
			}

			++iLines;
			++iDataRows;

			// Status
			sStatusFmt = _T("");
			GetUIStr( params.iLanguage, _T("StatDlg.Text.GetDataProgr"), sStatusFmt );
			sStatus.Format( sStatusFmt, iDataRows );
			SET_STATUS( sStatus );

			// Einfügen
Paste:
			if ( ( iLines == PASTE_AT || bNoMoreRows ) && iLines > 0 )
			{
				// Status
				sStatus = _T("");
				GetUIStr( params.iLanguage, _T("StatDlg.Text.Format"), sStatus );
				SET_STATUS( sStatus );

				// Letzte Zeile ermitteln
				iPasteLastRow = iPasteRow + iLines - 1;

				// Gesamten Bereich formatieren
				AttachRange( WSheet, R, iStartCol, iPasteRow, iLastCol, iPasteLastRow );

				// Alles löschen
				R.ClearContents( );

				// Datentyp = Standard
				v = _T("");
				R.SetNumberFormatLocal( v );

				// Spaltenweise formatieren
				for( iPos = 0; iPos < iCols; ++iPos )
				{
					iCol = iStartCol + iPos;
					AttachRange(WSheet, R, iCol, iPasteRow, iCol, iPasteLastRow);

					// Datentyp setzen, wenn Spalte einen Format-String hat
					if ( !Cols[iPos].sFormat.IsEmpty( ) )
					{						
						v = Cols[iPos].sFormat;				
						R.SetNumberFormatLocal( v );
					}

					// Horizontale Ausrichtung = wie Tabellenspalte
					v = Cols[iPos].lJustify;
					R.SetHorizontalAlignment(v);

					// Vertikale Ausrichtung = wie Tabellenspalte
					v = Cols[iPos].lVAlign;
					R.SetVerticalAlignment(v);
				}
				
				// Zellenspezifische Formatierungen VOR Einfügen
				if ( iRFVs > 0 )
				{
					pos = RFVs.GetHeadPosition( );
					while ( pos )
					{
						prfv = &RFVs.GetNext( pos );

						if ( prfv->dwUse & USE_NUMBERFORMAT )
						{
							v = prfv->sNumberFormat;
							prfv->R.SetNumberFormatLocal( v );
						}

						if (prfv->dwUse & USE_JUSTIFY)
						{
							v = prfv->lJustify;
							prfv->R.SetHorizontalAlignment(v);
						}

						if (prfv->dwUse & USE_VALIGN)
						{
							v = prfv->lVAlign;
							prfv->R.SetVerticalAlignment(v);
						}
					}

					RFVs.RemoveAll( );
					iRFVs = 0;
				}

				// Spaltenweise einfügen
				sStatus = _T("");
				GetUIStr( params.iLanguage, _T("StatDlg.Text.InsData"), sStatus );
				SET_STATUS( sStatus );
				for( iPos = 0; iPos < iCols; ++iPos )
				{
					iCol = iStartCol + iPos;					

					// Einfügen
					if ( bUseClipboard )
					{
						if ( !PasteStrings[iPos].IsEmpty( ) )
						{
							iMaxTries = 5;
							iTries = 0;
							while (iTries < iMaxTries)
							{
								try
								{
									CopyToClipboard(PasteStrings[iPos], params.hWndTbl);

									AttachRange(WSheet, R, iCol, iPasteRow);
									R.Select();
									
									WSheet.PasteSpecial( );
									iTries = iMaxTries;
								}
								catch (COleDispatchException * pE)
								{
									iTries++;
									if (iTries == iMaxTries)
										throw;
									else
									{
										pE->Delete();
										Sleep(100 * iTries);
									}
								}
								catch (COleException * pE)
								{
									iTries++;
									if (iTries == iMaxTries)
										throw;
									else
									{
										pE->Delete();
										Sleep(100 * iTries);
									}										
								}
								catch (CException * pE)
								{
									iTries++;
									if (iTries == iMaxTries)
										throw;
									else
									{
										pE->Delete();
										Sleep(100 * iTries);
									}
								}
								catch (CString & E)
								{
									iTries++;
									if (iTries == iMaxTries)
										throw;
									else
										Sleep(100 * iTries);
								}
							}

							PasteStrings[iPos].Empty( );
						}
					}
					else
					{
						for ( int i = 0; i < iLines; i++ )
						{
							if ( i == 0 )
								AttachRange( WSheet, R, iCol, iPasteRow );
							else
								OffsetRange( R, 1, 0 );
							v = Cols[iPos].CellTexts[i];
							R.SetFormulaLocal( v );
						}
					}
				}				 

				iPasteRow = iPasteLastRow + 1;
				iLines = 0;
			}
			
			if ( bNoMoreRows )
				break;

			++iRow;

			// Abbruch?
			if ( IS_CANCELED )
			{
				iRet = MTE_ERR_CANCELLED;
				goto Terminate;
			}
		}

		// Gesamtanzahl Zeilen setzen
		iRows = iHdrRows + iDataRows;

		// Letzte Zeile setzen
		iLastRow = iStartRow + iRows - 1;

		if ( iRows > 0 )
		{
			// Formatierungen
			//SET_STATUS( params.iLanguage == MTE_LNG_GERMAN ? _T("Formatierungen werden durchgeführt...") : _T("Formattings...") );
			sStatus = _T("");
			GetUIStr( params.iLanguage, _T("StatDlg.Text.Format"), sStatus );
			SET_STATUS( sStatus );

			// Datenbereich
			if ( iDataRows > 0 )
			{
				// Gesamter Datenbereich
				AttachRange( WSheet, R, iStartCol, iRowDataBegin, iLastCol, iLastRow );

				// Farben
				if ( bColors )
				{
					// Hintergrundfarbe der Tabelle setzen
					pInterior = R.GetInterior( );
					I.AttachDispatch( pInterior );
					if ( dwBackColorTbl == RGB( 255, 255, 255 ) )
					{
						v = xlNone;
						I.SetColorIndex( v );
					}
					else
					{
						v = long( dwBackColorTbl );
						I.SetColor( v );
						v = xlSolid;
						I.SetPattern( v );
					}

					// Textfarbe der Tabelle setzen
					pFont = R.GetFont( );
					F.AttachDispatch( pFont );
					v = long( dwTextColorTbl );
					F.SetColor( v );			
				}

				// Fonts
				if ( bFonts )
				{
					// Font der Tabelle setzen
					pFont = R.GetFont( );
					F.AttachDispatch( pFont );
					SetFont( F, fiTbl );
				}
								
				// Bereichsformatierungen
				if ( iRFs > 0 )
				{
					pos = RFs.GetHeadPosition( );
					while ( pos )
					{
						prf = &RFs.GetNext( pos );

						if ( prf->dwUse & USE_TEXT_COLOR || prf->dwUse & USE_FONT )
						{
							pFont = prf->R.GetFont( );
							F.AttachDispatch( pFont );

							if ( prf->dwUse & USE_TEXT_COLOR )
							{
								v = long( prf->dwTextColor );
								F.SetColor( v );
							}

							if ( prf->dwUse & USE_FONT )
							{
								SetFont( F, prf->fi );
							}
						}

						if ( prf->dwUse & USE_BACK_COLOR )
						{
							pInterior = prf->R.GetInterior( );
							I.AttachDispatch( pInterior );
							v = long( prf->dwBackColor );
							I.SetColor( v );
							v = xlSolid;
							I.SetPattern( v );
						}

						if ( prf->dwUse & USE_NUMBERFORMAT )
						{
							v = prf->sNumberFormat;
							prf->R.SetNumberFormatLocal( v );
						}

						if ( prf->dwUse & USE_INDENT_LEVEL )
						{
							v = prf->lIndentLevel;
							prf->R.SetIndentLevel( v );
						}

						if ( prf->dwUse & USE_JUSTIFY )
						{
							v = prf->lJustify;
							prf->R.SetHorizontalAlignment( v );
						}

						if ( prf->dwUse & USE_VALIGN )
						{
							v = prf->lVAlign;
							prf->R.SetVerticalAlignment( v );
						}

						// Abbruch?
						if ( IS_CANCELED )
						{
							iRet = MTE_ERR_CANCELLED;
							goto Terminate;
						}
					}
				}

				// Merge-Bereiche
				if ( iCMs > 0 )
				{
					pos = CMs.GetHeadPosition( );
					while ( pos )
					{
						pcm = &CMs.GetNext( pos );

						// Alle Zellen außer der Ersten löschen, damit keine Warnung kommt
						if ( pcm->iToCol > pcm->iFromCol )
						{
							AttachRange( WSheet, R, pcm->iFromCol + 1, pcm->iFromLine, pcm->iToCol, pcm->iToLine );
							R.ClearContents( );
						}
						if ( pcm->iToLine > pcm->iFromLine )
						{
							AttachRange( WSheet, R, pcm->iFromCol, pcm->iFromLine + 1, pcm->iFromCol, pcm->iToLine );
							R.ClearContents( );
						}

						// Einrückung der Master-Zelle ermitteln ( wird durch Merge zurückgesetzt :( )
						AttachRange( WSheet, R, pcm->iFromCol, pcm->iFromLine, pcm->iFromCol, pcm->iFromLine );
						vIndentLevel = R.GetIndentLevel( );

						// Zellen verbinden
						AttachRange( WSheet, R, pcm->iFromCol, pcm->iFromLine, pcm->iToCol, pcm->iToLine );
						v = 1L;
						R.SetMergeCells( v );

						// Einrückung wiederherstellen
						R.SetIndentLevel( vIndentLevel );

						// Abbruch?
						if ( IS_CANCELED )
						{
							iRet = MTE_ERR_CANCELLED;
							goto Terminate;
						}
					}
				}
			}

			// Gitternetz
			if ( bGrid )
			{
				// Gesamter Bereich
				AttachRange( WSheet, R, iStartCol, iStartRow, iLastCol, iLastRow );
				pDisp = R.GetBorders( );
				Bs.AttachDispatch( pDisp );

				// Diagonale Linien ausschalten
				pDisp = Bs.GetItem( xlDiagonalDown );
				B.AttachDispatch( pDisp );
				v = xlNone;
				B.SetLineStyle( v );

				pDisp = Bs.GetItem( xlDiagonalUp );
				B.AttachDispatch( pDisp );
				v = xlNone;
				B.SetLineStyle( v );
				
				// Vertikal
				// Außen
				pDisp = Bs.GetItem(xlEdgeRight);
				B.AttachDispatch(pDisp);

				SetBorder(params.hWndTbl, B, ColLinesTbl);

				// Innen
				if ( iCols > 1 )
				{
					pDisp = Bs.GetItem(xlInsideVertical);
					B.AttachDispatch(pDisp);

					SetBorder( params.hWndTbl, B, ColLinesTbl );
				}

				// Horizontal
				// Unten
				pDisp = Bs.GetItem(xlEdgeBottom);
				B.AttachDispatch(pDisp);

				SetBorder(params.hWndTbl, B, RowLinesTbl);

				// Innen
				if ( iRows > 1 )
				{
					pDisp = Bs.GetItem(xlInsideHorizontal);
					B.AttachDispatch(pDisp);

					SetBorder ( params.hWndTbl, B, RowLinesTbl );
				}
				
				if ( bColHeaders )
				{
					// Gesamter Header-Bereich
					AttachRange( WSheet, R, iStartCol, iStartRow, iLastCol, iStartRow + iHdrRows - 1 );

					pDisp = R.GetBorders( );
					Bs.AttachDispatch( pDisp );

					// Innere vertikale Linien					
					if ( iCols > 1 )
					{
						pDisp = Bs.GetItem(xlInsideVertical);
						B.AttachDispatch(pDisp);
						SetBorder(params.hWndTbl, B, HdrLinesTbl);
					}

					// Innere horizontale Linien
					if ( iHdrRows > 1 )
					{
						pDisp = Bs.GetItem(xlInsideHorizontal);
						B.AttachDispatch(pDisp);
						SetBorder(params.hWndTbl, B, HdrLinesTbl);
					}

					// Untere Linien
					pDisp = Bs.GetItem( xlEdgeBottom );
					B.AttachDispatch( pDisp );
					SetBorder(params.hWndTbl, B, HdrLinesTbl);

					// Column-Header
					for (iPos = 0; iPos < iCols; ++iPos)
					{
						item.DefineAsColHdr(Cols[iPos].hWnd);
						MTblQueryEffItemLineC(item, HorzLineDef, MTLT_HORZ);
						MTblQueryEffItemLineC(item, VertLineDef, MTLT_VERT);

						iCol = iStartCol + iPos;

						AttachRange(WSheet, R, iCol, iStartRow, iCol, iStartRow + iHdrRows - 1);

						pDisp = R.GetBorders();
						Bs.AttachDispatch(pDisp);

						pDisp = Bs.GetItem(xlEdgeBottom);
						B.AttachDispatch(pDisp);
						SetBorder(params.hWndTbl, B, HorzLineDef);

						pDisp = Bs.GetItem(xlEdgeRight);
						B.AttachDispatch(pDisp);
						SetBorder(params.hWndTbl, B, VertLineDef);
					}

					if (IS_CANCELED)
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}

					// Column-Header Gruppen
					if ( iCHGs > 0 )
					{
						pos = CHGs.GetHeadPosition( );
						while ( pos )
						{
							pchg = &CHGs.GetNext( pos );

							// Untere und rechte Linie
							AttachRange(WSheet, R, pchg->iFromCol, pchg->iFromLine, pchg->iToCol, pchg->iFromLine);

							pDisp = R.GetBorders();
							Bs.AttachDispatch(pDisp);

							pDisp = Bs.GetItem(xlEdgeBottom);
							B.AttachDispatch(pDisp);
							SetBorder(params.hWndTbl, B, pchg->HorzLineDef);

							pDisp = Bs.GetItem(xlEdgeRight);
							B.AttachDispatch(pDisp);
							SetBorder(params.hWndTbl, B, pchg->VertLineDef);

							// Evtl. innere Linien (gesamter Bereich inkl. Spaltenköpfe) ausschalten
							if ( pchg->bNoInnerVLines || pchg->bNoInnerHLine )
							{
								AttachRange( WSheet, R, pchg->iFromCol, pchg->iFromLine, pchg->iToCol, pchg->iToLine );

								pDisp = R.GetBorders( );
								Bs.AttachDispatch( pDisp );

								// Vertikal
								if ( pchg->bNoInnerVLines )
								{
									pDisp = Bs.GetItem( xlInsideVertical );
									B.AttachDispatch( pDisp );
									v = xlNone;
									B.SetLineStyle( v );
								}
								
								// Horizontal
								if ( pchg->bNoInnerHLine )
								{
									pDisp = Bs.GetItem( xlInsideHorizontal );
									B.AttachDispatch( pDisp );
									v = xlNone;
									B.SetLineStyle( v );
								}
							}
						}
					}

					if (IS_CANCELED)
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}				

				// Vertikale Linienformate
				for (iPos = 0; iPos < iCols; ++iPos)
				{
					iCol = iStartCol + iPos;

					VertLineDef2 = ColLinesTbl;

					iRow = iRowDataBegin - 1;
					pos = TblRows.GetHeadPosition();
					while (pos)
					{
						lRow = TblRows.GetNext(pos);
						iRow++;

						item.DefineAsCell(Cols[iPos].hWnd, lRow);
						MTblQueryEffItemLineC(item, VertLineDef, MTLT_VERT);

						// Teil eines Zellenverbundes?
						pmc = (LPMTBLMERGECELL)MTblFindMergeCell(Cols[iPos].hWnd, lRow, FMC_ANY, pMergeCells);
						if (pmc)
						{
							// Master-Zelle?
							if (pmc->hWndColFrom == Cols[iPos].hWnd && pmc->lRowFrom == lRow)
							{
								// Linienformat setzen
								AttachRange(WSheet, R, iCol, iRow, iCol + pmc->iMergedCols, iRow + pmc->iMergedRows);
								SetBorder(params.hWndTbl, R, xlEdgeRight, VertLineDef);
							}
						}
						else if (VertLineDef != VertLineDef2)
						{
							// Folgende Zellen mit gleichem Linienformat?
							pos2 = pos;
							iRowFrom = iRowTo = iRow2 = iRow;
							while (pos2)
							{
								lRow2 = TblRows.GetNext(pos2);
								iRow2++;

								item.DefineAsCell(Cols[iPos].hWnd, lRow2);
								MTblQueryEffItemLineC(item, VertLineDef2, MTLT_VERT);

								// Teil eines Zellenverbundes?
								pmc = (LPMTBLMERGECELL)MTblFindMergeCell(Cols[iPos].hWnd, lRow2, FMC_ANY, pMergeCells);

								if (!pmc && VertLineDef == VertLineDef2)
								{
									iRowTo = iRow2;
									iRow = iRow2;
									pos = pos2;
								}
								else
								{
									VertLineDef2 = ColLinesTbl;
									break;
								}
									
							}

							// Linienformat setzen								
							AttachRange(WSheet, R, iCol, iRowFrom, iCol, iRowTo);
							SetBorder(params.hWndTbl, R, xlEdgeRight, VertLineDef);
						}
					}

					// Abbruch?
					if (IS_CANCELED)
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}

				// Horizontale Linienformate
				iRow = iRowDataBegin - 1;
				pos = TblRows.GetHeadPosition();
				while (pos)
				{
					lRow = TblRows.GetNext(pos);
					iRow++;

					HorzLineDef2 = RowLinesTbl;

					for (iPos = 0; iPos < iCols; ++iPos)
					{
						iCol = iStartCol + iPos;

						if (MTblQueryCellFlags(hWndCol, lRow, MTBL_CELL_FLAG_NO_ROWLINE))
						{
							HorzLineDef.Init();
							HorzLineDef.Style = MTLS_INVISIBLE;
						}
						else
						{
							item.DefineAsCell(Cols[iPos].hWnd, lRow);
							MTblQueryEffItemLineC(item, HorzLineDef, MTLT_HORZ);
						}

						// Teil eines Zellenverbundes?
						pmc = (LPMTBLMERGECELL)MTblFindMergeCell(Cols[iPos].hWnd, lRow, FMC_ANY, pMergeCells);
						if (pmc)
						{
							// Master-Zelle?
							if (pmc->hWndColFrom == Cols[iPos].hWnd && pmc->lRowFrom == lRow)
							{
								// Linienformat setzen
								AttachRange(WSheet, R, iCol, iRow, iCol + pmc->iMergedCols, iRow + pmc->iMergedRows);

								pDisp = R.GetBorders();
								Bs.AttachDispatch(pDisp);

								pDisp = Bs.GetItem(xlEdgeBottom);
								B.AttachDispatch(pDisp);

								SetBorder(params.hWndTbl, B, HorzLineDef);
							}
						}
						else if (HorzLineDef != HorzLineDef2)
						{
							// Folgende Zellen mit gleichem Linienformat?
							iColFrom = iColTo = iCol;
							for (iPos2 = iPos + 1; iPos2 < iCols; ++iPos2)
							{
								iCol2 = iStartCol + iPos2;

								item.DefineAsCell(Cols[iPos2].hWnd, lRow);
								MTblQueryEffItemLineC(item, HorzLineDef2, MTLT_HORZ);

								// Teil eines Zellenverbundes?
								pmc = (LPMTBLMERGECELL)MTblFindMergeCell(Cols[iPos2].hWnd, lRow, FMC_ANY, pMergeCells);

								if (!pmc && HorzLineDef == HorzLineDef2)
								{
									iColTo = iCol2;
									iCol = iCol2;
								}
								else
								{
									HorzLineDef2 = RowLinesTbl;
									break;
								}
							}

							// Linienformat setzen
							AttachRange(WSheet, R, iColFrom, iRow, iColTo, iRow);

							pDisp = R.GetBorders();
							Bs.AttachDispatch(pDisp);

							pDisp = Bs.GetItem(xlEdgeBottom);
							B.AttachDispatch(pDisp);

							SetBorder(params.hWndTbl, B, HorzLineDef);
						}
					}

					// Abbruch?
					if (IS_CANCELED)
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}

				// Linien-Sichtbarkeits-Flags der Spalten berücksichtigen
				for (iPos = 0; iPos < iCols; ++iPos)
				{
					if (Cols[iPos].bNoColLine || Cols[iPos].bNoRowLines)
					{
						iCol = iStartCol + iPos;
						AttachRange(WSheet, R, iCol, iRowDataBegin, iCol, iLastRow);

						if (Cols[iPos].bNoColLine)
						{
							VertLineDef.Init();
							VertLineDef.Style = MTLS_INVISIBLE;
							
							SetBorder(params.hWndTbl, R, xlEdgeRight, VertLineDef);
						}

						if (Cols[iPos].bNoRowLines)
						{
							HorzLineDef.Init();
							HorzLineDef.Style = MTLS_INVISIBLE;

							SetBorder(params.hWndTbl, R, xlInsideHorizontal, HorzLineDef);
							SetBorder(params.hWndTbl, R, xlEdgeBottom, HorzLineDef);
						}
					}

					// Abbruch?
					if (IS_CANCELED)
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}

				// Linien-Sichtbarkeits-Flags der Zeilen/Zellen berücksichtigen
				iRow = iRowDataBegin - 1;
				pos = TblRows.GetHeadPosition();
				while (pos)
				{
					lRow = TblRows.GetNext(pos);
					iRow++;

					bRowNoColLines = MTblQueryRowFlags(params.hWndTbl, lRow, MTBL_ROW_NO_COLLINES);
					bRowNoRowLine = MTblQueryRowFlags(params.hWndTbl, lRow, MTBL_ROW_NO_ROWLINE);

					if (bRowNoColLines || bRowNoRowLine)
					{
						AttachRange(WSheet, R, iStartCol, iRow, iStartCol + iCols - 1, iRow);

						// Horizontal
						if (bRowNoRowLine)
						{
							HorzLineDef.Init();
							HorzLineDef.Style = MTLS_INVISIBLE;
							
							SetBorder(params.hWndTbl, R, xlEdgeBottom, HorzLineDef);
						}

						// Vertikal
						if (bRowNoColLines)
						{
							HorzLineDef.Init();
							HorzLineDef.Style = MTLS_INVISIBLE;

							SetBorder(params.hWndTbl, R, xlInsideVertical, HorzLineDef);
							SetBorder(params.hWndTbl, R, xlEdgeRight, HorzLineDef);
						}
					}

					// Zellen
					if (!(bRowNoColLines && bRowNoRowLine))
					{
						for (iPos = 0; iPos < iCols; ++iPos)
						{
							bCellNoColLine = bRowNoColLines || Cols[iPos].bNoColLine ? FALSE : MTblQueryCellFlags(Cols[iPos].hWnd, lRow, MTBL_CELL_FLAG_NO_COLLINE);
							bCellNoRowLine = bRowNoRowLine || Cols[iPos].bNoRowLines ? FALSE : MTblQueryCellFlags(Cols[iPos].hWnd, lRow, MTBL_CELL_FLAG_NO_ROWLINE);

							if (bCellNoColLine || bCellNoRowLine)
							{
								iCol = iStartCol + iPos;
								AttachRange(WSheet, R, iCol, iRow);

								// Horizontal
								if (bCellNoRowLine)
								{
									HorzLineDef.Init();
									HorzLineDef.Style = MTLS_INVISIBLE;

									SetBorder(params.hWndTbl, R, xlEdgeBottom, HorzLineDef);
								}

								// Vertikal
								if (bCellNoColLine)
								{
									VertLineDef.Init();
									VertLineDef.Style = MTLS_INVISIBLE;

									SetBorder(params.hWndTbl, R, xlEdgeRight, VertLineDef);
								}
							}
						}
					}

					// Abbruch?
					if (IS_CANCELED)
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}
			}

			// Spaltenbreiten anpassen
			if ( !bNoAutoFitCol )
			{
				AttachRange( WSheet, R, iStartCol, iStartRow, iLastCol, iLastRow );
				pRange = R.GetColumns( );
				R.AttachDispatch( pRange );
				R.AutoFit( );
			}

			// Abbruch?
			if (IS_CANCELED)
			{
				iRet = MTE_ERR_CANCELLED;
				goto Terminate;
			}

			for ( iPos = 0; iPos < iCols; ++iPos )
			{
				hWndCol = Cols[iPos].hWnd;
				iCol = iStartCol + iPos;

				lColWidth = SendMessage( params.hWndTbl, MTM_QueryExcelColWidth, (WPARAM)hWndCol, 0 );
				if ( lColWidth )
				{
					AttachRange( WSheet, R, iCol, iStartRow, iCol, iLastRow );
					pRange = R.GetColumns( );
					R.AttachDispatch( pRange );

					v = double(lColWidth) / 100.0;
					R.SetColumnWidth( v );
				}

				// Abbruch?
				if (IS_CANCELED)
				{
					iRet = MTE_ERR_CANCELLED;
					goto Terminate;
				}
			}

			// Zeilenhöhen anpassen
			if ( !bNoAutoFitRow )
			{
				AttachRange( WSheet, R, iStartCol, iStartRow, iLastCol, iLastRow );
				pRange = R.GetRows( );
				R.AttachDispatch( pRange );
				R.AutoFit( );
			}

			// Abbruch?
			if (IS_CANCELED)
			{
				iRet = MTE_ERR_CANCELLED;
				goto Terminate;
			}

			iRow = iRowDataBegin - 1;
			pos = TblRows.GetHeadPosition( );
			while ( pos )
			{
				lRow = TblRows.GetNext( pos );
				iRow++;

				lRowHeight = SendMessage( params.hWndTbl, MTM_QueryExcelRowHeight, 0, lRow );
				if ( lRowHeight > 0 )
				{					
					AttachRange( WSheet, R, iStartCol, iRow, iLastCol, iRow );
					pRange = R.GetRows( );
					R.AttachDispatch( pRange );

					v = double(lRowHeight) / 100.0;
					R.SetRowHeight( v );
				}

				// Abbruch?
				if (IS_CANCELED)
				{
					iRet = MTE_ERR_CANCELLED;
					goto Terminate;
				}
			}
		}

		// Startzelle selektieren
		AttachRange( WSheet, R, iStartCol, iStartRow );
		R.Select( );

		// Autom. Berechnung auf Originalwert stellen
		if ( lCalculationOrig != xlManual )
			App.SetCalculation( lCalculationOrig );

		// Benötigte Zeit ermitteln
		tEnd = CTime::GetCurrentTime( );
		ts = tEnd - tBegin;

		// Evtl. speichern
		if ( params.sSaveAsFile.GetLength() > 0 )
		{
			MAKE_STATUS_DLG_TOPMOST( FALSE );

			// Status
			//SET_STATUS( params.iLanguage == MTE_LNG_GERMAN ? _T("Datei wird gespeichert...") : _T("Saving file...") );
			sStatus = _T("");
			GetUIStr( params.iLanguage, _T("StatDlg.Text.SaveFile"), sStatus );
			SET_STATUS( sStatus );

			sWBookFullName = WBook.GetFullName( );
			v = params.sSaveAsFile;

			if ( sWBookFullName == params.sSaveAsFile )
				WBook.Save( );
			else
				WBook.SaveAs( v, vOpt, vOpt, vOpt, vOpt, vOpt, 0, vOpt, vOpt, vOpt, vOpt );

			MAKE_STATUS_DLG_TOPMOST( TRUE );
		}

		// Evtl. Workbook schließen
		if ( bCloseWorkbook || bCloseInstance )
		{
			// Status
			//SET_STATUS( params.iLanguage == MTE_LNG_GERMAN ? _T("Workbook wird geschlossen...") : _T("Closing workbook...") );
			sStatus = _T("");
			GetUIStr( params.iLanguage, _T("StatDlg.Text.CloseWBook"), sStatus );
			SET_STATUS( sStatus );

			v.Clear( );
			v.ChangeType( VT_BOOL );
			v = (short)0;
			WBook.Close( v, vOpt, vOpt );
			WBook.ReleaseDispatch( );
		}

		// Evtl. Instanz schließen
		if ( bCloseInstance )
		{
			// Status
			//SET_STATUS( params.iLanguage == MTE_LNG_GERMAN ? _T("Instanz wird geschlossen...") : _T("Closing instance...") );
			sStatus = _T("");
			GetUIStr( params.iLanguage, _T("StatDlg.Text.CloseInst"), sStatus );
			SET_STATUS( sStatus );

			App.Quit( );
			App.ReleaseDispatch( );
		}
	}
	catch( COleDispatchException * pE )
	{
		iRet = MTE_ERR_OLE_EXCEPTION;
		HandleException( pE, pDlgThread, APP_EXCEL, params.iLanguage );
		pE->Delete( );
	}
	catch( COleException * pE )
	{
		iRet = MTE_ERR_OLE_EXCEPTION;
		HandleException( pE, pDlgThread, APP_EXCEL, params.iLanguage );
		pE->Delete( );
	}
	catch( CException * pE )
	{
		iRet = MTE_ERR_EXCEPTION;
		HandleException( pE, pDlgThread, APP_EXCEL, params.iLanguage );
		pE->Delete( );
	}
	catch( CString & E )
	{
		iRet = MTE_ERR_EXCEPTION;

		CString sError;
		GetErrorMsg( E, params.iLanguage, sError );

		HandleException( sError, pDlgThread, APP_EXCEL, params.iLanguage );
	}

Terminate:
	// Sanduhr aus
	if ( App.m_lpDispatch )
	{
		try
		{
			App.SetCursor( xlNormal );
			App.SetVisible( TRUE );
		}
		catch( CException * pE )
		{
			pE->Delete( );
		}
	}

	// Ressourcen freigeben
	if ( hsText )
		Ctd.HStringUnRef( hsText );

	if ( pMergeCells )
		MTblFreeMergeCells( params.hWndTbl, pMergeCells );

	// Zwischenablage wiederherstellen
	if ( bUseClipboard )
		CopyToClipboard( sClipboardData, params.hWndTbl );

	// Dialog-Thread beenden
	if ( pDlgThread )
	{
		DestroyWindow( pDlgThread->GetDlgHandle( ) );
		PostThreadMessage( pDlgThread->m_nThreadID, WM_QUIT, 0, 0 );
	}

	return iRet;
}

/////////////////////////////////////////////////////////////////////////////
// Exportierte Funktionen

// MTblExportToExcel
int WINAPI MTblExportToExcel( HWND hWndTbl, int iLanguage, LPCTSTR lpsExcelStartCell, DWORD dwExcelFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff )
{
	EXCEL_EXP_PARAMS params;

	params.hWndTbl = hWndTbl;
	params.iLanguage = iLanguage;
	params.sExcelStartCell = lpsExcelStartCell;
	params.sSaveAsFile = _T("");
	params.sTemplateFile = _T("");
	params.sSheetName = _T("");
	params.dwExcelFlags = dwExcelFlags;
	params.dwExpFlags = dwExpFlags;
	params.wRowFlagsOn = wRowFlagsOn;
	params.wRowFlagsOff = wRowFlagsOff;
	params.dwColFlagsOn = dwColFlagsOn;
	params.dwColFlagsOff = dwColFlagsOff;

	return ExportToExcel( params );
}

// MTblExportToExcelEx
int WINAPI MTblExportToExcelEx( HWND hWndTbl, HUDV hUdv )
{
	EXCEL_EXP_PARAMS_UDV *pUdvParams = (EXCEL_EXP_PARAMS_UDV*) Ctd.UdvDeref( hUdv );
	EXCEL_EXP_PARAMS params;
	long lLen;

	params.hWndTbl = hWndTbl;
	params.iLanguage = Ctd.NumToInt( pUdvParams->Language );
	params.sExcelStartCell = Ctd.StringGetBuffer( pUdvParams->ExcelStartCell, &lLen );
	Ctd.HStrUnlock(pUdvParams->ExcelStartCell);
	params.sSaveAsFile = Ctd.StringGetBuffer( pUdvParams->ExcelSaveAsFile, &lLen );
	Ctd.HStrUnlock(pUdvParams->ExcelSaveAsFile);
	params.sTemplateFile = Ctd.StringGetBuffer( pUdvParams->ExcelTemplateFile, &lLen );
	Ctd.HStrUnlock(pUdvParams->ExcelTemplateFile);
	params.sSheetName = Ctd.StringGetBuffer( pUdvParams->ExcelSheetName, &lLen );
	Ctd.HStrUnlock(pUdvParams->ExcelSheetName);
	params.dwExcelFlags = Ctd.NumToULong( pUdvParams->ExcelFlags );
	params.dwExpFlags = Ctd.NumToULong( pUdvParams->ExpFlags );
	params.wRowFlagsOn = (WORD)Ctd.NumToULong( pUdvParams->RowFlagsOn );
	params.wRowFlagsOff = (WORD)Ctd.NumToULong( pUdvParams->RowFlagsOff );
	params.dwColFlagsOn = Ctd.NumToULong( pUdvParams->ColFlagsOn );
	params.dwColFlagsOff = Ctd.NumToULong( pUdvParams->ColFlagsOff );

	return ExportToExcel( params );
}