#pragma warning(disable:4786)

// Includes
#include "mtoocalcexp.h"

// Globale Variablen
extern CCtd Ctd;
extern CString sError;

static const int MAX_COLS = 1024;
static const int MAX_COL_IDX = MAX_COLS - 1;
static const int MAX_ROWS = 65536;
static const int MAX_ROW_IDX = MAX_ROWS - 1;
static const int MAX_SHEETS = 256;

// Funktionen
void SetBorder( HWND hWndTbl, OOCellRange &R, int iType, LINEDEF &ld )
{
	short shOuterWidth = ld.iStyle == MTLS_INVISIBLE ? 0 : 35;
	long lColor = (long)ld.dwColor;
	SetBorder( hWndTbl, R, iType, shOuterWidth, lColor );
}

void SetBorder( HWND hWndTbl, OOCellRange &R, int iType, short shOuterWidth, long lColor /*=0*/, short shInnerWidth /*=0*/, short shLineDistance /*=0*/ )
{
	BOOL bLeftRight_TopBottom = shOuterWidth < 1 && (iType == BL_Top || iType == BL_Left || iType == BL_Right || iType == BL_Bottom);
	BOOL bContinue = FALSE;
	BOOL bTableBorder = iType == BL_Top || iType == BL_Left || iType == BL_Right || iType == BL_Bottom || iType == BL_Vertical || iType == BL_Horizontal;
	int iColFrom, iColTo, iRowFrom, iRowTo, iTypeNext = -1, iTypePrev = -1;
	LPDISPATCH pDisp;
	OOBorderLine BL;
	OOCellRange ooCellRange;
	OOTableBorder TB;

	if ( lColor == MTBL_COLOR_UNDEF )
	{
		lColor = 0;
	}
	else
	{
		DWORD dwSubstColor = MTblGetExportSubstColor( hWndTbl, lColor );
		if ( dwSubstColor != MTBL_COLOR_UNDEF )
			lColor = (long)dwSubstColor;
	}

	do
	{
		pDisp = R.GetRangeAddress( );
		if ( !pDisp )
			break;
		OOCellRangeAddress ooCellRangeAddr( pDisp );

		iColFrom = ooCellRangeAddr.StartColumn( );
		iRowFrom = ooCellRangeAddr.StartRow( );
		iColTo = ooCellRangeAddr.EndColumn( );
		iRowTo = ooCellRangeAddr.EndRow( );

		pDisp = R.GetSpreadsheet( );
		if ( !pDisp )
			break;			
		OOSpreadsheet ooSheet( pDisp );		

		if ( bLeftRight_TopBottom )
		{
			switch ( iType )
			{
			case BL_Left:						
				if ( iTypePrev == BL_Right )
				{
					iColFrom = iColTo = iColTo + 1;
					iTypeNext = -1;
				}
				else if ( iColFrom - 1 >= 0 )
				{
					iTypeNext = BL_Right;
				}
				
				break;

			case BL_Right:
				if ( iTypePrev == BL_Left )
				{
					iColFrom = iColTo = iColFrom -1;
					iTypeNext = -1;
				}
				else if ( iColTo + 1 <= MAX_COL_IDX )
				{
					iTypeNext = BL_Left;
				}
				break;

			case BL_Top:
				if ( iTypePrev == BL_Bottom )
				{
					iRowFrom = iRowTo = iRowTo + 1;
					iTypeNext = -1;
				}
				else if ( iRowFrom - 1 >= 0 )
				{
					iTypeNext = BL_Bottom;
				}
				break;

			case BL_Bottom:
				if ( iTypePrev == BL_Top )
				{
					iRowFrom = iRowTo = iRowFrom - 1;
					iTypeNext = -1;
				}
				else if ( iRowTo + 1 <= MAX_ROW_IDX )
				{
					iTypeNext = BL_Top;
				}
				break;
			}

		}

		AttachRange( ooSheet, ooCellRange, iColFrom, iRowFrom, iColTo, iRowTo );

		if ( bTableBorder )
		{
			pDisp = ooCellRange.GetTableBorder( );
			if ( !pDisp )
				return;
			TB.AttachDispatch( pDisp );

			pDisp = TB.GetBorderLine( iType );
			if ( !pDisp )
				return;
			BL.AttachDispatch( pDisp );
		}
		else
		{
			pDisp = ooCellRange.GetBorderLine( iType );
			if ( !pDisp )
				return;
			BL.AttachDispatch( pDisp );
		}

		BL.SetColor( lColor );
		BL.SetInnerLineWidth( shInnerWidth );
		BL.SetOuterLineWidth( shOuterWidth );
		BL.SetLineDistance( shLineDistance );

		if ( bTableBorder )
		{
			TB.SetLineValid( iType, TRUE );
			TB.SetBorderLine( iType, BL.m_lpDispatch );
			ooCellRange.SetTableBorder( TB.m_lpDispatch );
			TB.SetLineValid( iType, FALSE );
		}
		else
		{
			ooCellRange.SetBorderLine( iType, BL.m_lpDispatch );
		}

		if ( iTypeNext >= 0 )
		{
			iTypePrev = iType;
			iType = iTypeNext;
		}
		else
		{
			iType = -1;
		}
	}
	while ( iType >= 0 );
}

// MTblExportToOOCalc
int WINAPI MTblExportToOOCalc( HWND hWndTbl, int iLanguage, LPCTSTR lpsCalcStartCell, DWORD dwCalcFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff )
{
	AFX_MANAGE_STATE( AfxGetStaticModuleState( ) );

	BOOL bAnyHdrGrp, bAnyHdrGrpWCH, bMergeCell, bCellNoColLine, bCellNoRowLine, bColHeaders, bColors, bCurrentPos, bFindSplitRows, bFonts, bFound, bGrid, bMustDeleteFont, bNewWorkbook, bNewWorksheet, bNoMoreRows, bNullsAsSpace, bRowNoColLines, bRowNoRowLine, bShowStatus, bSplitRows, bStringColsAsText;
	CArray<CString,CString&> PasteStrings;
	//CArray< CArray<CString,CString&>, CArray<CString,CString&>& > Values;
	CArray< CList<OO_RANGEFORMAT,OO_RANGEFORMAT&>, CList<OO_RANGEFORMAT,OO_RANGEFORMAT&>& >CRFs;
	CELLMERGE cm;
	CELLMERGE * pcm;
	CExpTblDlgThread * pDlgThread = NULL;
	CList<CELLMERGE,CELLMERGE&>CMs;
	CList<COLHDRGRP,COLHDRGRP&> CHGs;
	CList<OO_RANGELINES,OO_RANGELINES&>RLs;
	COleSafeArray Props;
	COLHDRGRP chg;
	COLHDRGRP * pchg;
	COLINFO_ARRAY Cols;
	CString sClipboardData, sCR( TCHAR(13) ), sEncl( _T("\"") ), sEnclLeft( _T("<tr><td>") ), sEnclRight( _T("</td></tr>") ), sName, sLF( TCHAR(10) ), sStartCell( lpsCalcStartCell ), sStatus, sTab, sText;
	CTime tBegin, tEnd;
	CTimeSpan ts;
	CWinThread *pConfirmerThread = NULL;
	DWORD dwBackColorCell, dwBackColorTbl, dwSubstColor, dwTextColorCell, dwTextColorTbl;
	FONTINFO fi, fiTbl;
	HFONT hFont;
	HWND hWndCol, hWndTreeCol;
	HSTRING hsText = 0, hsLanguageFile = 0;
	int iCHGs = 0, iCMs = 0, iCol, iCols, iDataRows, iHdrRows, iID, iIndent, iLastCol, iLastHdrGrpCol, iLastRow, iLines, iNodeSize, iPasteRow, iPos, iRealLockedCols, iRet = 1, iRFs = 0, iRLs = 0, iRow, iRowDataBegin, iRowLevel, iRows, iStartCol, iStartRow, j;
	long lCellJustify, lCellVAlign, lIndentLevel, lJustify, lLen, lRow, lSheet;
	LINEDEF ColLinesTbl, RowLinesTbl;
	LPDISPATCH pDispDocument = NULL, pDispSheet = NULL, pDispTmp = NULL;
	LPMTBLMERGECELL pmc;
	LPTSTR lps;
	LPVOID pHdrGrp, pMergeCells = NULL;
	OOBorderLine ooBorderLine;
	OOCell ooCell;
	OOCellRange ooCellRange;
	OOCellRangeAddress ooCellRangeAddr;
	OOComponents ooComponents;
	OODesktop ooDesktop;
	OODispatchHelper ooDispatchHelper;
	OOEnumeration ooEnumeration;
	OOFrame ooFrame;
	OOIdlClass ooIdlPropertyValue;
	OOIndexAccess ooIndexAccess;
	OOModel ooModel;
	OONumberFormatter ooNumberFormatter;
	OOConfirmer ooConfirmer;
	OOPropertyValue ooPropertyValue;
	OOReflection ooReflection;
	OOServiceManager ooSvcMan;
	OOSpreadsheetDocument ooDocument;
	OOSpreadsheetView ooView;
	OOSpreadsheet ooSheet;
	OOSpreadsheets ooSheets;
	OOSystemClipboard ooClipboard;
	OOTableBorder ooTableBorder;
	OOTransferable ooTransferable;
	OO_RANGEFORMAT rf;
	OO_RANGEFORMAT * prf;
	OO_RANGELINES rl;
	OO_RANGELINES *prl;
	POSITION pos;
	SAFEARRAY *pArr = NULL;
	VARIANT v;
	TCHAR cPwd;
	
	const int PASTE_AT = 1000;

	try
	{
		// Subclassed?
		if ( !MTblIsSubClassed( hWndTbl ) )
			return MTE_ERR_NOT_SUBCLASSED;

		tBegin = CTime::GetCurrentTime( );

		// Tree-Definition ermitteln
		MTblQueryTree( hWndTbl, &hWndTreeCol, &iNodeSize, &iIndent );

		// Passwort-Zeichen ermitteln
		cPwd = MTblGetPasswordCharVal( hWndTbl );

		// Sprache
		/*if ( !( iLanguage == MTE_LNG_GERMAN || iLanguage == MTE_LNG_ENGLISH ) )
			iLanguage = MTE_LNG_ENGLISH;*/
		if ( !MTblGetLanguageFile( iLanguage, &hsLanguageFile ) )
		{
			return MTE_ERR_INVALID_LANG;
		}

		if ( hsLanguageFile != 0 )
			Ctd.HStringUnRef( hsLanguageFile );

		// Flags ermitteln
		bNewWorkbook = dwCalcFlags & MTE_OOCALC_NEW_WORKBOOK;
		bNewWorksheet = dwCalcFlags & MTE_OOCALC_NEW_WORKSHEET;
		bCurrentPos = dwCalcFlags & MTE_OOCALC_CURRENT_POS;
		bStringColsAsText = dwCalcFlags & MTE_OOCALC_STRING_COLS_AS_TEXT;
		bNullsAsSpace = dwCalcFlags & MTE_OOCALC_NULLS_AS_SPACE;

		bColHeaders = dwExpFlags & MTE_COL_HEADERS;
		bColors = dwExpFlags & MTE_COLORS;
		bFonts = dwExpFlags & MTE_FONTS;
		bGrid = dwExpFlags & MTE_GRID;
		bShowStatus = dwExpFlags & MTE_SHOW_STATUS;
		bSplitRows = dwExpFlags & MTE_SPLIT_ROWS;
		if ( bSplitRows )
		{
			BOOL bDragAdjust;
			int iSplitRows = 0;
			if ( !( Ctd.TblQuerySplitWindow( hWndTbl, &iSplitRows, &bDragAdjust ) && iSplitRows > 0 ) )
				bSplitRows = FALSE;
		}

		// Service-Manger-Instanz erzeugen
		if ( !ooSvcMan.CreateDispatch( _T("com.sun.star.ServiceManager") ) )
			ThrowError( ERR_GET_OBJECT );

		// Desktop-Instanz erzeugen
		pDispTmp = ooSvcMan.CreateInstance( _T("com.sun.star.frame.Desktop") );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooDesktop.AttachDispatch( pDispTmp );

		// Dokument ermitteln
		if ( !bNewWorkbook )
		{
			pDispTmp = ooDesktop.GetComponents( );
			if ( !pDispTmp )
				return MTE_ERR_RUNNING_INSTANCE;
			ooComponents.AttachDispatch( pDispTmp );

			pDispTmp = ooComponents.CreateEnumeration( );
			if ( !pDispTmp )
				return MTE_ERR_RUNNING_INSTANCE;
			ooEnumeration.AttachDispatch( pDispTmp );

			VariantInit( &v );
			while ( ooEnumeration.HasMoreElements( ) )
			{
				v = ooEnumeration.NextElement( );
				if ( v.vt == VT_DISPATCH )
				{
					if ( OOSpreadsheetDocument::IsSpreadsheetDocument( v.pdispVal ) )
					{
						pDispDocument = v.pdispVal;
						pDispDocument->AddRef( );
					}
				}
				VariantClear( &v );

				if ( pDispDocument )
					break;
			}
		}

		if ( !pDispDocument )
			pDispDocument = ooDesktop.LoadComponentFromUrl( _T( "private:factory/scalc" ), _T( "_blank" ), 0 );

		if ( !pDispDocument )
			ThrowError( ERR_GET_OBJECT );
		ooDocument.AttachDispatch( pDispDocument );

		// View ermitteln
		ooModel.AttachDispatch( ooDocument.m_lpDispatch, FALSE );
		pDispTmp = ooModel.GetCurrentController( );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooView.AttachDispatch( pDispTmp );

		// Frame ermitteln
		pDispTmp = ooView.GetFrame( );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooFrame.AttachDispatch( pDispTmp );

		// Dispatch-Helper erzeugen
		pDispTmp = ooSvcMan.CreateInstance( _T("com.sun.star.frame.DispatchHelper") );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooDispatchHelper.AttachDispatch( pDispTmp );

		// Reflection erzeugen
		pDispTmp = ooSvcMan.CreateInstance( _T("com.sun.star.reflection.CoreReflection") );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooReflection.AttachDispatch( pDispTmp );

		// Idl-Klasse für PropertySet ermittlen
		pDispTmp = ooReflection.ForName( _T("com.sun.star.beans.PropertyValue") );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooIdlPropertyValue.AttachDispatch( pDispTmp );

		// Number-Formatter erzeugen
		pDispTmp = ooSvcMan.CreateInstance( _T("com.sun.star.util.NumberFormatter") );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooNumberFormatter.AttachDispatch( pDispTmp );
		ooNumberFormatter.AttachNumberFormatsSupplier( pDispDocument );

		// System-Clipboard erzeugen
		pDispTmp = ooSvcMan.CreateInstance( _T("com.sun.star.datatransfer.clipboard.SystemClipboard") );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooClipboard.AttachDispatch( pDispTmp );

		// Dialog-Thread starten
		if ( bShowStatus )
		{
			pDlgThread = (CExpTblDlgThread*)AfxBeginThread( RUNTIME_CLASS( CExpTblDlgThread ) );
			if ( !pDlgThread )
				return MTE_ERR_STATUS;

			// Sprache setzen
			SET_STATUS_LNG( iLanguage );
		}

		// Spreadsheets-Objekt ermitteln
		pDispTmp = ooDocument.GetSheets( );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooSheets.AttachDispatch( pDispTmp );

		// Worksheet hinzufügen oder ermitteln
		if ( bNewWorksheet )
		{
			lSheet = ooSheets.Add( iLanguage );
			if ( lSheet < 0 )
				ThrowError( ERR_INSERT_SHEET );

			pDispSheet = ooSheets.GetByIndex( lSheet - 1);
			if ( pDispSheet )
				ooView.SetActiveSheet( pDispSheet );
		}
		else
			pDispSheet = ooView.GetActiveSheet( );

		if ( !pDispSheet )
			ThrowError( ERR_GET_OBJECT );
		ooSheet.AttachDispatch( pDispSheet );

		
		// Startposition selektieren
		if ( !bCurrentPos )
		{
			if ( sStartCell.IsEmpty( ) )
				pDispTmp = ooSheet.GetCellRangeByPosition( 0, 0, 0, 0 );
			else
				pDispTmp = ooSheet.GetCellRangeByName( sStartCell );
			
			if ( !pDispTmp )
				ThrowError( ERR_GET_OBJECT );
			ooCellRange.AttachDispatch( pDispTmp );
			ooView.Select( ooCellRange.m_lpDispatch );
		}

		// Startposition ermitteln
		pDispTmp = ooView.GetSelection( );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooCellRange.AttachDispatch( pDispTmp );

		pDispTmp = ooCellRange.GetRangeAddress( );
		if ( !pDispTmp )
			ThrowError( ERR_GET_OBJECT );
		ooCellRangeAddr.AttachDispatch( pDispTmp );

		iStartCol = ooCellRangeAddr.StartColumn( );
		if ( iStartCol < 0 )
			ThrowError( ERR_GET_START_CELL );

		iStartRow = ooCellRangeAddr.StartRow( );
		if ( iStartRow < 0 )
			ThrowError( ERR_GET_START_CELL );

		// Tatsächliche Anzahl gelockter Spalten ermitteln
		iRealLockedCols = MTblQueryLockedColumns( hWndTbl );

		// Zu exportierende Spalten ermitteln
		iCols = GetCols( TRUE, hWndTbl, dwColFlagsOn, dwColFlagsOff, iRealLockedCols, hWndTreeCol, Cols, iStartCol, MAX_COLS, bStringColsAsText, bAnyHdrGrp, bAnyHdrGrpWCH, pDlgThread );
		if ( iCols <= 0 )
		{
			if ( iCols == GC_ERR_CANCELLED )
				iRet = MTE_ERR_CANCELLED;
			goto Terminate;
		}

		// Letzte Spalte setzen
		iLastCol = iStartCol + iCols - 1;

		// Merge-Zellen ermitteln
		pMergeCells = MTblGetMergeCells( hWndTbl, wRowFlagsOn, wRowFlagsOff, dwColFlagsOn, dwColFlagsOff );

		// Farben der Tabelle ermitteln
		if ( bColors )
		{
			dwTextColorTbl = Ctd.ColorGet( hWndTbl, COLOR_IndexWindowText );
			dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwTextColorTbl );
			if ( dwSubstColor != MTBL_COLOR_UNDEF )
				dwTextColorTbl = dwSubstColor;

			dwBackColorTbl = Ctd.ColorGet( hWndTbl, COLOR_IndexWindow );
			dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwBackColorTbl );
			if ( dwSubstColor != MTBL_COLOR_UNDEF )
				dwBackColorTbl = dwSubstColor;
		}

		// Font der Tabelle ermitteln
		if ( bFonts )
		{
			GetFontInfo( hWndTbl, HFONT( SendMessage( hWndTbl, WM_GETFONT, 0, 0 ) ), fiTbl );
		}

		// Spaltenlinien der Tabelle ermitteln
		ColLinesTbl.iStyle = MTLS_SOLID;
		ColLinesTbl.dwColor = MTBL_COLOR_UNDEF;
		MTblQueryColLines( hWndTbl, &ColLinesTbl.iStyle, &ColLinesTbl.dwColor );		

		// Zeilenlinien der Tabelle ermitteln
		RowLinesTbl.iStyle = MTLS_SOLID;
		RowLinesTbl.dwColor = MTBL_COLOR_UNDEF;
		MTblQueryRowLines( hWndTbl, &RowLinesTbl.iStyle, &RowLinesTbl.dwColor );

		// Spaltenköpfe exportieren
		//SET_STATUS( iLanguage == MTE_LNG_GERMAN ? _T("Spaltenköpfe werden exportiert...") : _T("Exporting column headers...") );
		sStatus = _T("");
		GetUIStr( iLanguage, _T("StatDlg.Text.ExpColHdrs"), sStatus );
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
			AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow, iLastCol, iRow );

			// Alles löschen
			ooCellRange.ClearContents( CF_All );

			// Datentyp = Text
			ooCellRange.SetNumberFormat( 100 );
			
			// Ausrichtung: vertikal = oben
			ooCellRange.SetVertJustify( CVJ_Top );
			
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
						AttachRange( ooSheet, ooCellRange, iCol, iStartRow, iCol, iRow );
						ooCellRange.Merge( TRUE );
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
						if ( MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS ) )
							AttachRange( ooSheet, ooCellRange, iCol, iStartRow, iLastHdrGrpCol, iRow );
						else
							AttachRange( ooSheet, ooCellRange, iCol, iStartRow, iLastHdrGrpCol, iStartRow );

						// Obere Zellen der Gruppe verbinden
						if ( iLastHdrGrpCol > iCol )
						{
							//ooCellRange.ClearContents( CF_All );
							ooCellRange.Merge( TRUE );
						}

						// Gruppe merken
						chg.iFromCol = iCol;
						chg.iFromLine = iStartRow;
						chg.iToCol = iLastHdrGrpCol;
						chg.iToLine = iRow;
						chg.bNoInnerVLines = MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOINNERVERTLINES );
						chg.bNoInnerHLine = MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOINNERHORZLINE );
						chg.bNoColHeaders = MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS );
						chg.bTxtAlignLeft = MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_TXTALIGN_LEFT );
						chg.bTxtAlignRight = MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_TXTALIGN_RIGHT );
						CHGs.AddTail( chg );
						++iCHGs;

						// Text ermitteln
						MTblGetColHdrGrpText( hWndTbl, pHdrGrp, &hsText );
						sText = Ctd.StringGetBuffer( hsText, &lLen );
						sText.Replace( sCR, _T("") );
						Ctd.HStrUnlock(hsText);

						// Horizontale Textausrichtung setzen
						if ( chg.bTxtAlignLeft )
							lJustify = CHJ_Left;
						else if ( chg.bTxtAlignRight )
							lJustify = CHJ_Right;
						else
							lJustify = CHJ_Center;
						ooCellRange.SetHoriJustify( lJustify );

						// Evtl. Farben setzen
						if ( bColors )
						{
							// Hintergrundfarbe = Hintergrundfarbe der Gruppe
							if ( MTblGetEffColHdrGrpBackColor( hWndTbl, pHdrGrp, &dwBackColorCell ) )
							{
								dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwBackColorCell );
								if ( dwSubstColor != MTBL_COLOR_UNDEF )
									dwBackColorCell = dwSubstColor;

								ooCellRange.SetBackColor( long( dwBackColorCell ) );
							}

							// Textfarbe = Textfarbe der Gruppe
							if ( MTblGetEffColHdrGrpTextColor( hWndTbl, pHdrGrp, &dwTextColorCell ) )
							{
								dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwTextColorCell );
								if ( dwSubstColor != MTBL_COLOR_UNDEF )
									dwTextColorCell = dwSubstColor;

								ooCellRange.SetTextColor( long( dwTextColorCell ) );
							}
						}

						// Evtl. Font setzen
						if ( bFonts )
						{
							hFont = NULL;
							
							if ( hFont = MTblGetEffColHdrGrpFontHandle( hWndTbl, pHdrGrp, &bMustDeleteFont ) )
							{
								GetFontInfo( hWndTbl, hFont, fi );
								if ( bMustDeleteFont )
									DeleteObject( hFont );
							}
							else
								fi = fiTbl;

							SetFont( ooCellRange, fi );
						}

						// Text setzen
						ooCellRange.SetValue( sText );

						// Range für Spaltenkopf ermitteln
						AttachRange( ooSheet, ooCellRange, iCol, iRow );
					}
					else
						// Range für Spaltenkopf ermitteln
						AttachRange( ooSheet, ooCellRange, iCol, iRow );
				}
				else
					// Range für Spaltenkopf ermitteln
					AttachRange( ooSheet, ooCellRange, iCol, iRow );

				if ( pHdrGrp && MTblQueryColHdrGrpFlags( hWndTbl, pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS ) )
					continue;

				// Horizontale Textausrichtung setzen
				ooCellRange.SetHoriJustify( Cols[iPos].lHdrJustify );

				// Evtl. Farben setzen
				if ( bColors )
				{
					// Hintergrundfarbe = Hintergrundfarbe des Spaltenkopfes
					if ( MTblGetEffColumnHdrBackColor( Cols[iPos].hWnd, &dwBackColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwBackColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwBackColorCell = dwSubstColor;

						ooCellRange.SetBackColor( long( dwBackColorCell ) );
					}

					// Textfarbe = Textfarbe des Spaltenkopfes
					if ( MTblGetEffColumnHdrTextColor( Cols[iPos].hWnd, &dwTextColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwTextColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwTextColorCell = dwSubstColor;

						ooCellRange.SetTextColor( long( dwTextColorCell ) );
					}
				}

				// Evtl. Font setzen
				if ( bFonts )
				{
					hFont = NULL;
					
					if ( hFont = MTblGetEffColumnHdrFontHandle( Cols[iPos].hWnd, &bMustDeleteFont ) )
					{
						GetFontInfo( hWndTbl, hFont, fi );
						if ( bMustDeleteFont )
							DeleteObject( hFont );
					}
					else
						fi = fiTbl;

					SetFont( ooCellRange, fi );
				}

				// Text ermitteln
				MTblGetColumnTitle( hWndCol, &hsText );
				sText = Ctd.StringGetBuffer( hsText, &lLen );
				sText.Replace( sCR, _T("") );
				Ctd.HStrUnlock(hsText);

				// Text setzen
				ooCellRange.SetValue( sText );

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
		//SET_STATUS( iLanguage == MTE_LNG_GERMAN ? _T("Daten werden ermittelt...") : _T("Getting data...") );
		sStatus = _T("");
		GetUIStr( iLanguage, _T("StatDlg.Text.GetData"), sStatus );
		SET_STATUS( sStatus );

		// Inhalt der Zwischenablage sichern
		sClipboardData = _T("");
		SaveClipboard( sClipboardData, hWndTbl );

		sTab = _T("	");
		++iRow;
		iRowDataBegin = iRow;
		iPasteRow = iRow;
		iDataRows = 0;
		PasteStrings.SetSize( iCols );
		CRFs.SetSize( iCols );
		/*Values.SetSize( iCols );
		for ( int i = 0; i < iCols; i++ )
		{
			Values[i].SetSize( PASTE_AT );
		}*/
		iLines = 0;
		lRow = -1;
		bNoMoreRows = FALSE;
		bFindSplitRows = FALSE;
		for( ; ; )
		{
			if ( !Ctd.TblFindNextRow( hWndTbl, &lRow, 0, 0 ) || ( bFindSplitRows ? lRow >= 0 : FALSE ) )
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

			if ( iRow > MAX_ROW_IDX )
			{
				bNoMoreRows = TRUE;
				goto Paste;
			}

			if ( wRowFlagsOn && !Ctd.TblQueryRowFlags( hWndTbl, lRow, wRowFlagsOn ) )
				continue;
			if ( wRowFlagsOff && Ctd.TblQueryRowFlags( hWndTbl, lRow, wRowFlagsOff ) )
				continue;

			Ctd.TblSetContextEx( hWndTbl, lRow );

			// Linienformat Zeile
			if ( bGrid )
			{
				bRowNoColLines = MTblQueryRowFlags( hWndTbl, lRow, MTBL_ROW_NO_COLLINES );
				bRowNoRowLine = MTblQueryRowFlags( hWndTbl, lRow, MTBL_ROW_NO_ROWLINE );

				rl.dwUse = 0;
				if ( bRowNoRowLine )
				{
					rl.dwUse |= USE_EDGE_BOTTOM;
					rl.ldEdgeBottom.iStyle = MTLS_INVISIBLE;
				}

				if ( bRowNoColLines )
				{
					rl.dwUse |= USE_VERT_INNER;
					rl.ldVertInner.iStyle = MTLS_INVISIBLE;
				}

				if ( rl.dwUse )
				{
					AttachRange( ooSheet, rl.R, iStartCol, iRow, iStartCol + iCols - 1, iRow );
					RLs.AddTail( rl );
					++iRLs;
				}
			}

			for ( iPos = 0; iPos < iCols; ++iPos )
			{
				hWndCol = Cols[iPos].hWnd;
				iCol = iStartCol + iPos;

				// Merge-Zelle?
				bMergeCell = FALSE;
				pmc = (LPMTBLMERGECELL)MTblFindMergeCell( Cols[iPos].hWnd, lRow, FMC_MASTER, pMergeCells );
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
				if ( MTblQueryCellFlags( hWndCol, lRow, MTBL_CELL_FLAG_HIDDEN_TEXT ) )
				{
					sText.Empty( );
				}
				else
				{
					iID = Ctd.TblQueryColumnID( hWndCol );
					Ctd.TblGetColumnText( hWndTbl, iID, &hsText );
					sText = Ctd.StringGetBuffer( hsText, &lLen );
					Ctd.HStrUnlock(hsText);
					/*if ( Cols[iPos].bInvisible && Cols[iPos].iCellType != COL_CellType_CheckBox )
					{
						lLen = sText.GetLength( );
						for ( int i = 0; i < lLen; ++i )
						{
							sText.SetAt( i, cPwd );
						}
					}
					else if ( !sText.IsEmpty( ) )
					{
						BOOL bEnclose = FALSE;
						if ( sText.Replace( sCR, _T("") ) )
						{
							sText.Replace( _T("\""), _T("\"\"") );
							bEnclose = TRUE;
						}
						lLen = sText.GetLength( );
						if ( bEnclose )
							sText = sEncl + sText + sEncl;
					}*/
					if ( Cols[iPos].bInvisible && Cols[iPos].iCellType != COL_CellType_CheckBox )
					{
						lLen = sText.GetLength( );
						for ( int i = 0; i < lLen; ++i )
						{
							sText.SetAt( i, cPwd );
						}
					}
				}

				// Evtl. NULL durch Leerzeichen ersetzen
				if ( bNullsAsSpace && sText.IsEmpty( ) )
				{
					sText = _T(" ");
					lLen = 1;
				}

				// In HTML-String umwandeln
				MakeHtmlStr( sText );
				lLen = sText.GetLength( );

				// "Einpacken"
				sText = sEnclLeft + sText + sEnclRight;

				// Pufferlänge des Paste-Strings setzen
				if ( iLines == 0 )
				{
					lps = PasteStrings[iPos].GetBufferSetLength( PASTE_AT * 254 );
					PasteStrings[iPos].ReleaseBuffer( 0 );
				}

				// Paste-String der Spalte erweitern
				if ( iLines > 0 )
					sText = sCR + sText;
				PasteStrings[iPos] += sText;
				//Values[iPos][iLines] = sText;

				// Formate ermitteln
				rf.dwUse = 0;
				rf.iColFrom = rf.iColTo = iCol;
				rf.iRowFrom = rf.iRowTo = iRow;
				rf.dwTextColor = dwTextColorTbl;
				rf.dwBackColor = dwBackColorTbl;
				rf.lNumberFormat = 0;
				rf.lIndentLevel = 0;
				rf.fi = fiTbl;
				rf.lJustify = xlNone;
				rf.lVAlign = xlNone;
				//rf.iUnionCount = 0;

				// Farben
				if ( bColors )
				{
					// Text
					if ( MTblGetEffCellTextColor( hWndCol, lRow, &dwTextColorCell ) )
					{
						dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwTextColorCell );
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
						dwSubstColor = MTblGetExportSubstColor( hWndTbl, dwBackColorCell );
						if ( dwSubstColor != MTBL_COLOR_UNDEF )
							dwBackColorCell = dwSubstColor;

						if ( dwBackColorCell != dwBackColorTbl )
						{
							rf.dwUse |= USE_BACK_COLOR;
							rf.dwBackColor = dwBackColorCell;
						}
					}
				}

				// Format
				/*if ( Cols[iPos].bExportAsText && lLen > MAX_TEXT_LEN )
				{
					rf.dwUse |= USE_NUMBERFORMAT;
					rf.lNumberFormat = 0;
				}*/

				// Font
				if ( bFonts )
				{
					if ( hFont = MTblGetEffCellFontHandle( hWndCol, lRow, &bMustDeleteFont ) )
					{
						GetFontInfo( hWndTbl, hFont, fi );
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
					if ( MTblQueryTreeFlags( hWndTbl, MTBL_TREE_FLAG_FLAT_STRUCT ) )
						iRowLevel = 0;
					else
						iRowLevel = MTblGetRowLevel( hWndTbl, lRow );

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

				// Horizontale Textausrichtung
				lCellJustify = MTblGetEffCellTextJustify( hWndCol, lRow );
				if ( lCellJustify == DT_LEFT )
					lCellJustify = CHJ_Left;
				else if ( lCellJustify == DT_CENTER )
					lCellJustify = CHJ_Center;
				else if ( lCellJustify == DT_RIGHT )
					lCellJustify = CHJ_Right;
				else
					lCellJustify = CHJ_Standard;

				if ( lCellJustify != CHJ_Standard && lCellJustify != Cols[iPos].lJustify )
				{
					rf.dwUse |= USE_JUSTIFY;
					rf.lJustify = lCellJustify;
				}

				// Vertikale Textausrichtung
				lCellVAlign = MTblGetEffCellTextVAlign( hWndCol, lRow );
				if ( lCellVAlign == DT_TOP )
					lCellVAlign = CVJ_Top;
				else if ( lCellVAlign == DT_VCENTER )
					lCellVAlign = CVJ_Center;
				else if ( lCellVAlign == DT_BOTTOM )
					lCellVAlign = CVJ_Center;
				else
					lCellVAlign = CVJ_Standard;

				if ( lCellVAlign != CVJ_Standard && lCellVAlign != Cols[iPos].lVAlign )
				{
					rf.dwUse |= USE_VALIGN;
					rf.lVAlign = lCellVAlign;
				}

				// Formatierungsbereich erstellen bzw. zu bestehendem Bereich hinzufügen
				if ( rf.dwUse )
				{
					bFound = FALSE;
					if ( !CRFs[iPos].IsEmpty( ) )
					{
						pos = CRFs[iPos].GetTailPosition( );
						while ( pos )
						{
							prf = &CRFs[iPos].GetPrev( pos );

							if ( rf.dwUse == prf->dwUse &&
								 rf.dwTextColor == prf->dwTextColor &&
								 rf.dwBackColor == prf->dwBackColor &&
								 rf.lNumberFormat == prf->lNumberFormat &&
								 rf.lIndentLevel == prf->lIndentLevel &&
								 rf.lJustify == prf->lJustify && 
								 rf.lVAlign == prf->lVAlign &&
								 rf.fi == prf->fi /*&&
								 prf->iUnionCount < MAX_UNION*/ )
							{
								if ( prf->iColFrom == iCol && prf->iRowTo == iRow - 1 )
								{
									AttachRange( ooSheet, prf->R, iCol, prf->iRowFrom, iCol, iRow );
									prf->iRowTo = iRow;
									//++prf->iUnionCount;
									bFound = TRUE;
								}
							}

							break;
						}
					}

					if ( !bFound )
					{
						AttachRange( ooSheet, rf.R, iCol, iRow );
						CRFs[iPos].AddTail( rf );
						++iRFs;
					}
				}

				// Linienformat Zelle
				if ( bGrid )
				{
					bCellNoColLine = MTblQueryCellFlags( hWndCol, lRow, MTBL_CELL_FLAG_NO_COLLINE );
					bCellNoRowLine = MTblQueryCellFlags( hWndCol, lRow, MTBL_CELL_FLAG_NO_ROWLINE );

					rl.dwUse = 0;
					if ( bCellNoColLine && ( ( !bRowNoColLines && !Cols[iPos].bNoColLine ) || bMergeCell ) )
					{
						rl.dwUse |= USE_EDGE_RIGHT;
						rl.ldEdgeRight.iStyle = MTLS_INVISIBLE;
					}

					if ( bCellNoRowLine && ( ( !bRowNoRowLine && !Cols[iPos].bNoRowLines ) || bMergeCell ) )
					{
						rl.dwUse |= USE_EDGE_BOTTOM;
						rl.ldEdgeBottom.iStyle = MTLS_INVISIBLE;
					}

					if ( rl.dwUse )
					{
						if ( bMergeCell )
							AttachRange( ooSheet, rl.R, cm.iFromCol, cm.iFromLine, cm.iToCol, cm.iToLine );
						else
							AttachRange( ooSheet, rl.R, iCol, iRow );
						RLs.AddTail( rl );
						++iRLs;
					}
				}
			}

			++iLines;
			++iDataRows;

			// Status
			if ( iLanguage == MTE_LNG_GERMAN )
				sStatus.Format( _T("Daten werden ermittelt... ( %d Zeilen ermittelt )"), iDataRows );
			else
				sStatus.Format( _T("Getting data... ( %d rows )"), iDataRows );
			SET_STATUS( sStatus );

			// Einfügen
Paste:
			if ( ( iLines == PASTE_AT || bNoMoreRows ) && iLines > 0 )
			{
				// Status
				//SET_STATUS( iLanguage == MTE_LNG_GERMAN ? _T("Daten werden eingefügt...") : _T("Inserting data...") );
				sStatus = _T("");
				GetUIStr( iLanguage, _T("StatDlg.Text.InsData"), sStatus );
				SET_STATUS( sStatus );

				// Letzte Zeile des Einfügebereichs setzen
				iLastRow = iRow;
				if ( bNoMoreRows )
					iLastRow--;

				// Gesamten Bereich formatieren
				AttachRange( ooSheet, ooCellRange, iStartCol, iPasteRow, iLastCol, iLastRow );

				// Alles löschen
				ooCellRange.ClearContents( CF_All );

				// Datentyp = Standard
				ooCellRange.SetNumberFormat( 0 );

				// Spaltenweise formatieren + einfügen
				for( iPos = 0; iPos < iCols; ++iPos )
				{
					iCol = iStartCol + iPos;
					
					// Range
					AttachRange( ooSheet, ooCellRange, iCol, iPasteRow, iCol, iLastRow );

					// Datentyp = Text, wenn Spalte als Text exportiert werden soll
					if ( Cols[iPos].bExportAsText )
					{						
						ooCellRange.SetNumberFormat( 100 );
					}

					// Einfügen
					if ( !PasteStrings[iPos].IsEmpty( ) )
					{
						/*for ( int i = 0; i < iLines; i++ )
						{
							CopyToClipboard( Values[iPos][i] );
							pDispTmp = ooClipboard.GetContents( );
							if ( !pDispTmp )
								ThrowError( ERR_GET_OBJECT );
							ooTransferable.AttachDispatch( pDispTmp );

							pDispTmp = ooCellRange.GetCellByPosition( 0, i );
							if ( !pDispTmp )
								ThrowError( ERR_GET_OBJECT );
							ooCell.AttachDispatch( pDispTmp );

							ooView.Select( ooCell.m_lpDispatch );
							ooView.InsertTransferable( ooTransferable.m_lpDispatch );
						}*/

						CopyToClipboard( PasteStrings[iPos], NULL, TRUE );
						pDispTmp = ooClipboard.GetContents( );
						if ( !pDispTmp )
							ThrowError( ERR_GET_OBJECT );
						ooTransferable.AttachDispatch( pDispTmp );

						pDispTmp = ooCellRange.GetCellByPosition( 0, 0 );
						if ( !pDispTmp )
							ThrowError( ERR_GET_OBJECT );
						ooCell.AttachDispatch( pDispTmp );

						ooView.Select( ooCell.m_lpDispatch );

						pConfirmerThread = AfxBeginThread( OOConfirmer::Start, &ooConfirmer );
						ooView.InsertTransferable( ooTransferable.m_lpDispatch );
						if ( !ooConfirmer.IsStopped( ) )
							ooConfirmer.Stop( );

						PasteStrings[iPos].Empty( );
					}

					// Abbruch?
					if ( IS_CANCELED )
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}				 

				iPasteRow = iRow + 1;
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
			//SET_STATUS( iLanguage == MTE_LNG_GERMAN ? _T("Formatierungen werden durchgeführt...") : _T("Formattings...") );
			sStatus = _T("");
			GetUIStr( iLanguage, _T("StatDlg.Text.Format"), sStatus );
			SET_STATUS( sStatus );


			// Datenbereich
			if ( iDataRows > 0 )
			{
				// Spaltenspezifisch
				for ( iPos = 0; iPos < iCols; ++iPos )
				{
					iCol = iStartCol + iPos;

					AttachRange( ooSheet, ooCellRange, iCol, iRowDataBegin, iCol, iLastRow );

					// Horizontale Ausrichtung = wie Tabellenspalte
					ooCellRange.SetHoriJustify( Cols[iPos].lJustify );

					// Vertikale Ausrichtung = wie Tabellenspalte
					ooCellRange.SetVertJustify( Cols[iPos].lVAlign );

					// Abbruch?
					if ( IS_CANCELED )
					{
						iRet = MTE_ERR_CANCELLED;
						goto Terminate;
					}
				}

				// Gesamter Datenbereich
				AttachRange( ooSheet, ooCellRange, iStartCol, iRowDataBegin, iLastCol, iLastRow );

				// Autom. Zeilenumbruch entfernen
				ooCellRange.SetTextWrap( FALSE );

				// Farben
				if ( bColors )
				{
					// Hintergrundfarbe der Tabelle setzen
					if ( dwBackColorTbl == RGB( 255, 255, 255 ) )
					{
						ooCellRange.SetBackColor( -1 );
					}
					else
					{
						ooCellRange.SetBackColor( long( dwBackColorTbl ) );
						//v = xlSolid;
						//I.SetPattern( v );
					}

					// Textfarbe der Tabelle setzen
					ooCellRange.SetTextColor( long( dwTextColorTbl ) );
				}

				// Fonts
				if ( bFonts )
				{
					// Font der Tabelle setzen
					SetFont( ooCellRange, fiTbl );
				}
				else
				{
					// Standard ( Arial 10 ) setzen
					ooCellRange.SetFontName( _T("Arial" ) );
					ooCellRange.SetFontSize( 10 );
				}

				// Bereichsformatierungen
				if ( iRFs > 0 )
				{
					/*pos = RFs.GetHeadPosition( );
					while ( pos )
					{
						prf = &RFs.GetNext( pos );

						if ( prf->dwUse & USE_TEXT_COLOR || prf->dwUse & USE_FONT )
						{
							if ( prf->dwUse & USE_TEXT_COLOR )
							{
								prf->R.SetTextColor( long( prf->dwTextColor ) );
							}

							if ( prf->dwUse & USE_FONT )
							{
								SetFont( prf->R, prf->fi );
							}
						}

						if ( prf->dwUse & USE_BACK_COLOR )
						{
							prf->R.SetBackColor( long( prf->dwBackColor ) );
							//v = xlSolid;
							//I.SetPattern( v );
						}

						if ( prf->dwUse & USE_NUMBERFORMAT )
						{
							prf->R.SetNumberFormat( prf->lNumberFormat );
						}

						if ( prf->dwUse & USE_INDENT_LEVEL )
						{
							prf->R.SetIndentLevel( prf->lIndentLevel );
						}

						if ( prf->dwUse & USE_JUSTIFY )
						{
							prf->R.SetHoriJustify( prf->lJustify );
						}

						if ( prf->dwUse & USE_VALIGN )
						{
							prf->R.SetVertJustify( prf->lVAlign );
						}

						// Abbruch?
						if ( IS_CANCELED )
						{
							iRet = MTE_ERR_CANCELLED;
							goto Terminate;
						}
					}*/

					for ( int i = 0; i < iCols; ++i )
					{
						pos = CRFs[i].GetHeadPosition( );
						while ( pos )
						{
							prf = &CRFs[i].GetNext( pos );

							if ( prf->dwUse & USE_TEXT_COLOR || prf->dwUse & USE_FONT )
							{
								if ( prf->dwUse & USE_TEXT_COLOR )
								{
									prf->R.SetTextColor( long( prf->dwTextColor ) );
								}

								if ( prf->dwUse & USE_FONT )
								{
									SetFont( prf->R, prf->fi );
								}
							}

							if ( prf->dwUse & USE_BACK_COLOR )
							{
								prf->R.SetBackColor( long( prf->dwBackColor ) );
								//v = xlSolid;
								//I.SetPattern( v );
							}

							if ( prf->dwUse & USE_NUMBERFORMAT )
							{
								prf->R.SetNumberFormat( prf->lNumberFormat );
							}

							if ( prf->dwUse & USE_INDENT_LEVEL )
							{
								prf->R.SetIndentLevel( prf->lIndentLevel );
							}

							if ( prf->dwUse & USE_JUSTIFY )
							{
								prf->R.SetHoriJustify( prf->lJustify );
							}

							if ( prf->dwUse & USE_VALIGN )
							{
								prf->R.SetVertJustify( prf->lVAlign );
							}

							// Abbruch?
							if ( IS_CANCELED )
							{
								iRet = MTE_ERR_CANCELLED;
								goto Terminate;
							}
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
							AttachRange( ooSheet, ooCellRange, pcm->iFromCol + 1, pcm->iFromLine, pcm->iToCol, pcm->iToLine );
							ooCellRange.ClearContents( CF_All );
						}
						if ( pcm->iToLine > pcm->iFromLine )
						{
							AttachRange( ooSheet, ooCellRange, pcm->iFromCol, pcm->iFromLine + 1, pcm->iFromCol, pcm->iToLine );
							ooCellRange.ClearContents( CF_All );
						}

						// Zellen verbinden
						AttachRange( ooSheet, ooCellRange, pcm->iFromCol, pcm->iFromLine, pcm->iToCol, pcm->iToLine );
						ooCellRange.Merge( TRUE );

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
				AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow, iLastCol, iLastRow );				

				// Diagonale Linien ausschalten
				SetBorder( hWndTbl, ooCellRange, BL_TopLeft_BottomRight, 0 );
				SetBorder( hWndTbl, ooCellRange, BL_BottomLeft_TopRight, 0 );

				// Gesamter Datenbereich
				AttachRange( ooSheet, ooCellRange, iStartCol, iRowDataBegin, iLastCol, iLastRow );
				
				// Vertikale Linien innen
				if ( iCols > 1 )
				{
					SetBorder( hWndTbl, ooCellRange, BL_Vertical, ColLinesTbl );
				}

				// Horizontale Linien innen
				if ( iRows > 1 )
				{
					SetBorder( hWndTbl, ooCellRange, BL_Horizontal, RowLinesTbl );
				}		
				
				if ( bColHeaders )
				{
					AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow, iLastCol, iStartRow + iHdrRows - 1 );

					// Innere Linien
					SetBorder( hWndTbl, ooCellRange, BL_Vertical, 35, 0 );
					SetBorder( hWndTbl, ooCellRange, BL_Horizontal, 35, 0 );

					// Untere Linien
					SetBorder( hWndTbl, ooCellRange, BL_Bottom, 35, 0 );

					// Column-Header Gruppen
					if ( iCHGs > 0 )
					{
						pos = CHGs.GetHeadPosition( );
						while ( pos )
						{
							pchg = &CHGs.GetNext( pos );

							// Evtl. innere Linien ausschalten
							if ( pchg->bNoInnerVLines || pchg->bNoInnerHLine )
							{
								AttachRange( ooSheet, ooCellRange, pchg->iFromCol, pchg->iFromLine, pchg->iToCol, pchg->iToLine );

								// Vertikal
								if ( pchg->bNoInnerVLines )
								{
									SetBorder( hWndTbl, ooCellRange, BL_Vertical, 0 );
								}
								
								// Horizontal
								if ( pchg->bNoInnerHLine )
								{
									SetBorder( hWndTbl, ooCellRange, BL_Horizontal, 0 );
								}
							}
						}
					}
				}

				// Spalten ohne Zeilenlinien oder ohne Spaltenlinie?
				if ( ColLinesTbl.iStyle != MTLS_INVISIBLE && RowLinesTbl.iStyle != MTLS_INVISIBLE )
				{
					for ( iPos = 0; iPos < iCols; ++iPos )
					{
						if ( Cols[iPos].bNoRowLines || Cols[iPos].bNoColLine )
						{
							iCol = iStartCol + iPos;

							AttachRange( ooSheet, ooCellRange, iCol, iRowDataBegin, iCol, iLastRow );

							if ( Cols[iPos].bNoRowLines && iDataRows > 1 )
							{
								SetBorder( hWndTbl, ooCellRange, BL_Horizontal, 0 );
							}

							if ( Cols[iPos].bNoColLine )
							{								
								SetBorder( hWndTbl, ooCellRange, BL_Right, 0 );
							}
						}
					}
				}

				// Linienformate
				if ( iRLs > 0 )
				{
					pos = RLs.GetHeadPosition( );
					while ( pos )
					{
						prl = &RLs.GetNext( pos );

						pDispTmp = prl->R.GetRangeAddress( );
						if ( !pDispTmp )
							break;
						ooCellRangeAddr.AttachDispatch( pDispTmp );

						int iColFrom = ooCellRangeAddr.StartColumn( );
						int iRowFrom = ooCellRangeAddr.StartRow( );
						int iColTo = ooCellRangeAddr.EndColumn( );
						int iRowTo = ooCellRangeAddr.EndRow( );
						if ( iColFrom < 0 || iRowFrom < 0 || iColTo < 0 || iRowTo < 0 )
							break;

						if ( prl->dwUse & USE_HORZ_INNER )
						{
							SetBorder( hWndTbl, prl->R, BL_Horizontal, prl->ldHorzInner );
						}

						if ( prl->dwUse & USE_VERT_INNER )
						{
							SetBorder( hWndTbl, prl->R, BL_Vertical, prl->ldVertInner );
						}

						if ( prl->dwUse & USE_EDGE_LEFT )
						{
							SetBorder( hWndTbl, prl->R, BL_Left, prl->ldEdgeLeft );
						}

						if ( prl->dwUse & USE_EDGE_TOP )
						{
							SetBorder( hWndTbl, prl->R, BL_Top, prl->ldEdgeTop );
						}

						if ( prl->dwUse & USE_EDGE_RIGHT )
						{
							SetBorder( hWndTbl, prl->R, BL_Right, prl->ldEdgeRight );
						}

						if ( prl->dwUse & USE_EDGE_BOTTOM )
						{
							SetBorder( hWndTbl, prl->R, BL_Bottom, prl->ldEdgeBottom );
						}

						if ( prl->dwUse & USE_DIAG_UP )
						{
							SetBorder( hWndTbl, prl->R, BL_BottomLeft_TopRight, prl->ldDiagUp );
						}

						if ( prl->dwUse & USE_DIAG_DOWN )
						{
							SetBorder( hWndTbl, prl->R, BL_TopLeft_BottomRight, prl->ldDiagDown );
						}
					}
				}

				// Äußere Linien gesamter Bereich einschalten
				AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow, iLastCol, iLastRow );

				SetBorder( hWndTbl, ooCellRange, BL_Left, 35, 0 );
				SetBorder( hWndTbl, ooCellRange, BL_Top, 35, 0 );
				SetBorder( hWndTbl, ooCellRange, BL_Right, 35, 0 );
				SetBorder( hWndTbl, ooCellRange, BL_Bottom, 35, 0 );
			}

			// Äußere Linien gesamter Bereich setzen
			AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow, iLastCol, iLastRow );

			short shWidth = bGrid ? 35 : 0;
			SetBorder( hWndTbl, ooCellRange, BL_Left, shWidth, 0 );
			SetBorder( hWndTbl, ooCellRange, BL_Top, shWidth, 0 );
			SetBorder( hWndTbl, ooCellRange, BL_Right, shWidth, 0 );
			SetBorder( hWndTbl, ooCellRange, BL_Bottom, shWidth, 0 );

				
			// Spaltenbreiten und Zeilenhöhen anpassen			
			long lIndices[] = { 0 };

			AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow, iLastCol, iLastRow );
			ooView.Select( ooCellRange.m_lpDispatch );

			// Spaltenbreiten
			Props.CreateOneDim( VT_DISPATCH, 1 );

			pDispTmp = ooIdlPropertyValue.CreateObject( );
			if ( !pDispTmp )
				ThrowError( ERR_GET_OBJECT );
			ooPropertyValue.AttachDispatch( pDispTmp, FALSE );

			ooPropertyValue.SetName( _T("aExtraWidth") );
			ooPropertyValue.SetValue( 100L );

			Props.PutElement( lIndices, ooPropertyValue.m_lpDispatch );			
			
			VariantInit( &v );
			v = ooDispatchHelper.ExecuteDispatch( ooFrame.m_lpDispatch, _T(".uno:SetOptimalColumnWidth"), _T(""), 0, Props );
			VariantClear( &v );

			Props.Detach( );

			// Zeilenhöhen
			Props.CreateOneDim( VT_DISPATCH, 1 );

			pDispTmp = ooIdlPropertyValue.CreateObject( );
			if ( !pDispTmp )
				ThrowError( ERR_GET_OBJECT );
			ooPropertyValue.AttachDispatch( pDispTmp, FALSE );

			ooPropertyValue.SetName( _T("aExtraHeight") );
			ooPropertyValue.SetValue( 0L );

			Props.PutElement( lIndices, ooPropertyValue.m_lpDispatch );			
			
			VariantInit( &v );
			v = ooDispatchHelper.ExecuteDispatch( ooFrame.m_lpDispatch, _T(".uno:SetOptimalRowHeight"), _T(""), 0, Props );
			VariantClear( &v );

			Props.Detach( );
		}

		// Startzelle selektieren
		AttachRange( ooSheet, ooCellRange, iStartCol, iStartRow );
		ooView.Select( ooCellRange.m_lpDispatch );

		// Benötigte Zeit ermitteln
		tEnd = CTime::GetCurrentTime( );
		ts = tEnd - tBegin;

	}
	catch( COleDispatchException * pE )
	{
		iRet = MTE_ERR_OLE_EXCEPTION;
		HandleException( pE, pDlgThread, APP_OOCALC, iLanguage );
		pE->Delete( );
	}
	catch( COleException * pE )
	{
		iRet = MTE_ERR_OLE_EXCEPTION;
		HandleException( pE, pDlgThread, APP_OOCALC, iLanguage );
		pE->Delete( );
	}
	catch( CException * pE )
	{
		iRet = MTE_ERR_EXCEPTION;
		HandleException( pE, pDlgThread, APP_OOCALC, iLanguage );
		pE->Delete( );
	}
	catch( CString & E )
	{
		iRet = MTE_ERR_EXCEPTION;

		CString sError;
		GetErrorMsg( E, iLanguage, sError );

		HandleException( sError, pDlgThread, APP_OOCALC, iLanguage );
	}

Terminate:
	// Sanduhr aus
	/*if ( iRet >= 0 && App.m_lpDispatch )
		App.SetCursor( xlNormal );*/

	// Ressourcen freigeben
	if ( hsText )
		Ctd.HStringUnRef( hsText );

	if ( pMergeCells )
		MTblFreeMergeCells( hWndTbl, pMergeCells );

	// Zwischenablage wiederherstellen
	CopyToClipboard( sClipboardData, hWndTbl );

	// Paste-Confirmer beenden
	if ( !ooConfirmer.IsStopped( ) )
		ooConfirmer.Stop( );

	// Dialog-Thread beenden
	if ( pDlgThread )
	{
		DestroyWindow( pDlgThread->GetDlgHandle( ) );
		PostThreadMessage( pDlgThread->m_nThreadID, WM_QUIT, 0, 0 );
	}

	return iRet;
}
