// Includes
#include "mtexp.h"

// Globale Variablen
extern CString sError;
extern CCtd Ctd;

/////////////////////////////////////////////////////////////////////////////
// FONTINFO

void FONTINFO::Init()
{
	sName.Empty();
	iPointSize = 0;
	bBold = bItalic = bUnderline = bStrikeOut = FALSE;
}

BOOL FONTINFO::operator==( FONTINFO &fi )
{
	return sName == fi.sName &&
		iPointSize == fi.iPointSize &&
		bBold == fi.bBold &&
		bItalic == fi.bItalic &&
		bUnderline == fi.bUnderline &&
		bStrikeOut == fi.bStrikeOut;
}

BOOL FONTINFO::operator!=(FONTINFO &fi)
{
	return !operator==(fi);
}

// Funktionen
int CharToUtf8( LPCTSTR lpChar, LPSTR lpCharUtf8, int iLenUtf8 )
{
#ifdef UNICODE
	return WideCharToMultiByte( CP_UTF8, 0, lpChar, -1, lpCharUtf8, iLenUtf8, NULL, NULL );
#else
	int iRet = 0;

	int iChars = MultiByteToWideChar( CP_THREAD_ACP, 0, lpChar, -1, NULL, 0 );
	if ( iChars )
	{
		WCHAR *tmp = new WCHAR[iChars];
		if ( MultiByteToWideChar( CP_THREAD_ACP, 0, lpChar, -1, tmp, iChars ) )
		{
			iRet = WideCharToMultiByte( CP_UTF8, 0, tmp, iChars, lpCharUtf8, iLenUtf8, NULL, NULL );
		}
		delete [] tmp;
	}
	return iRet;
#endif;
}

/*void CopyToClipboard( CString & s, HWND hWnd )
{
	if ( s.GetLength( ) < 1 )
		return;

	BOOL bClipboardOpen = FALSE;
	int iMaxTries = 5;
	for ( int i = 0; i < iMaxTries; ++i )
	{
		if ( bClipboardOpen = OpenClipboard( hWnd ) )
			break;
		else
		{
			CloseClipboard( );
			Sleep( 100 );
		}
	}

	if ( bClipboardOpen )
	{
		EmptyClipboard( );
		HGLOBAL hGlobal = GlobalAlloc( GMEM_MOVEABLE | GMEM_DDESHARE, ( s.GetLength( ) * sizeof(TCHAR) ) + sizeof(TCHAR) );
		if ( hGlobal )
		{
			LPTSTR pBuf = (LPTSTR)GlobalLock( hGlobal );
			lstrcpy( pBuf, s );
			GlobalUnlock( hGlobal );
#ifdef UNICODE
			UINT uiFormat = CF_UNICODETEXT;
#else
			UINT uiFormat = CF_TEXT;
#endif
			if ( SetClipboardData( uiFormat, hGlobal ) )
			{
				CloseClipboard( );				
			}
			else
			{				
				CloseClipboard( );
				GlobalFree( hGlobal );
				ThrowError( ERR_COPY_TO_CLIPBOARD_SETDATA );
			}
		}
		else
		{
			CloseClipboard( );
			ThrowError( ERR_COPY_TO_CLIPBOARD_ALLOC );
		}
	}
	else
	{
		ThrowError( ERR_COPY_TO_CLIPBOARD_OPEN );
	}
}*/

void CopyToClipboard( CString & s, HWND hWnd /*= NULL*/, BOOL bHTML /*= FALSE*/ )
{
	if ( s.GetLength( ) < 1 )
		return;

	static UINT uiFormatHtml = RegisterClipboardFormat( _T("HTML Format") );

	BOOL bCvtToUTF8 = FALSE;
	BYTE *pData = NULL;
	char *pDataUtf8 = NULL;
	CString sData;
	int iDataLen = 0;
	UINT uiFormat;

	if ( bHTML )
	{
		uiFormat = uiFormatHtml;

		static CString sCR( TCHAR(13) );
		static CString sLF( TCHAR(10) );
		static CString sCRLF = sCR + sLF;

		static CString sHdr1 = _T("Version:0.9") + sCRLF;
		static CString sHdr2 = _T("StartHTML:XStartHtml") + sCRLF;
		static CString sHdr3 = _T("EndHTML:XXXEndHtml") + sCRLF;
		static CString sHdr4 = _T("StartFragment:XXXStartFr") + sCRLF;
		static CString sHdr5 = _T("EndFragment:XXXXXEndFr") + sCRLF;
		static CString sHdr = sHdr1 + sHdr2 + sHdr3 + sHdr4 + sHdr5;

		static CString sStartTag = _T("<HTML><BODY>");
		static CString sEndTag = _T("</BODY></HTML>");
		//static CString sStartFr = _T("<!--StartFragment -->");
		//static CString sEndFr = _T("<!--EndFragment -->");

		CString sBegin = sHdr + sStartTag /* sStartFr*/;
		CString sEnd = /*sEndFr +*/ sEndTag;

		CString sFmt = _T("%010u");
		CString sPos;
		int iPos;
		
		int iHtmlLen = CharToUtf8( s, NULL, 0 );
		char *pHtmlUtf8 = new char[iHtmlLen];
		CharToUtf8( s, pHtmlUtf8, iHtmlLen );

		iPos = sHdr.GetLength( );
		sPos.Format( sFmt, iPos );
		sBegin.Replace( _T("XStartHtml"), sPos );

		iPos += sStartTag.GetLength( );
		sPos.Format( sFmt, iPos );
		sBegin.Replace( _T("XXXStartFr"), sPos );

		iPos += (iHtmlLen -1 );
		sPos.Format( sFmt, iPos );
		sBegin.Replace( _T("XXXXXEndFr"), sPos );

		iPos += sEnd.GetLength( );
		sPos.Format( sFmt, iPos );
		sBegin.Replace( _T("XXXEndHtml"), sPos );

		int iBeginLen = CharToUtf8( sBegin, NULL, 0 );
		char *pBeginUtf8 = new char[iBeginLen];
		CharToUtf8( sBegin, pBeginUtf8, iBeginLen );

		int iEndLen = CharToUtf8( sEnd, NULL, 0 );
		char *pEndUtf8 = new char[iEndLen];
		CharToUtf8( sEnd, pEndUtf8, iEndLen );

		iDataLen = (iBeginLen-1) + (iHtmlLen-1) + (iEndLen - 1) + 1;
		pDataUtf8 = new char[iDataLen];
		char *p = pDataUtf8;
		memcpy( p, pBeginUtf8, iBeginLen );
		p += (iBeginLen - 1);
		memcpy( p, pHtmlUtf8, iHtmlLen );
		p += (iHtmlLen - 1);
		memcpy( p, pEndUtf8, iEndLen );

		delete [] pBeginUtf8;
		delete [] pHtmlUtf8;
		delete [] pEndUtf8;
	}
	else
	{
#ifdef UNICODE
		uiFormat = CF_UNICODETEXT;
#else
		uiFormat = CF_TEXT;
#endif
		sData = s;
		iDataLen = ( sData.GetLength( ) * sizeof(TCHAR) ) + sizeof(TCHAR);
	}

	BOOL bClipboardOpen = FALSE;
	int iMaxTries = 5;
	for ( int i = 0; i < iMaxTries; ++i )
	{
		if ( bClipboardOpen = OpenClipboard( hWnd ) )
			break;
		else
		{
			CloseClipboard( );
			Sleep( 100 );
		}
	}

	if ( bClipboardOpen )
	{
		EmptyClipboard( );
		if ( uiFormat == CF_TEXT )
			SetClipboardData( CF_LOCALE, NULL );

		HGLOBAL hGlobal = NULL;
		if ( iDataLen > 0 )
			hGlobal = GlobalAlloc( GMEM_MOVEABLE | GMEM_DDESHARE, iDataLen );

		if ( hGlobal )
		{
			LPTSTR pBuf = (LPTSTR)GlobalLock( hGlobal );
			if ( uiFormat == uiFormatHtml )
			{
				memcpy( pBuf, pDataUtf8, iDataLen );
				//MessageBox( NULL, pDataUtf8, _T("Test"), 0 );
				delete [] pDataUtf8;
			}
			else
			{
				lstrcpy( pBuf, sData );
			}
			GlobalUnlock( hGlobal );

			if ( SetClipboardData( uiFormat, hGlobal ) )
			{
				CloseClipboard( );				
			}
			else
			{				
				CloseClipboard( );
				GlobalFree( hGlobal );
				ThrowError( ERR_COPY_TO_CLIPBOARD_SETDATA );
			}
		}
		else
		{
			CloseClipboard( );
			ThrowError( ERR_COPY_TO_CLIPBOARD_ALLOC );
		}
	}
	else
	{
		ThrowError( ERR_COPY_TO_CLIPBOARD_OPEN );
	}
}

int GetCols( BOOL bOpenOffice, HWND hWndTbl, DWORD dwColFlagsOn, DWORD dwColFlagsOff, int iRealLockedCols, HWND hWndTreeCol, COLINFO_ARRAY &Cols, int iStartCol, int iMaxCols, BOOL bStringColsAsText, BOOL &bAnyHdrGrp, BOOL &bAnyHdrGrpWCH, CExpTblDlgThread *pDlgThread )
{
	Cols.SetSize( iMaxCols );
	bAnyHdrGrp = bAnyHdrGrpWCH = FALSE;
	int iCols = 0;
	int iPos = 1;
	while ( HWND hWndCol = Ctd.TblGetColumnWindow( hWndTbl, iPos, COL_GetPos ) )
	{
		if ( ( dwColFlagsOn ? Ctd.TblQueryColumnFlags( hWndCol, dwColFlagsOn ) : TRUE ) && ( dwColFlagsOff ? !Ctd.TblQueryColumnFlags( hWndCol, dwColFlagsOff ) : TRUE ) )
		{
			if ( ( iCols + iStartCol - ( bOpenOffice ? 0 : 1 ) ) >= iMaxCols )
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
				Cols[iCols].lJustify = bOpenOffice ? CHJ_Left : xlLeft;
			else if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_CENTER ) )
				Cols[iCols].lJustify = bOpenOffice ? CHJ_Center : xlCenter;
			else if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_RIGHT ) )
				Cols[iCols].lJustify = bOpenOffice ? CHJ_Right : xlRight;
			else if ( Ctd.TblQueryColumnFlags( hWndCol, COL_RightJustify ) )
				Cols[iCols].lJustify = bOpenOffice ? CHJ_Right : xlRight;
			else if ( Ctd.TblQueryColumnFlags( hWndCol, COL_CenterJustify ) )
				Cols[iCols].lJustify = bOpenOffice ? CHJ_Center : xlCenter;
			else
				Cols[iCols].lJustify = bOpenOffice ? CHJ_Left : xlLeft;

			// Horizontale Ausrichtung Header ermitteln
			if ( MTblQueryColumnHdrFlags( hWndCol, MTBL_COLHDR_FLAG_TXTALIGN_LEFT ) )
				Cols[iCols].lHdrJustify = bOpenOffice ? CHJ_Left : xlLeft;
			else if ( MTblQueryColumnHdrFlags( hWndCol, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT ) )
				Cols[iCols].lHdrJustify = bOpenOffice ? CHJ_Right : xlRight;
			else
				Cols[iCols].lHdrJustify = bOpenOffice ? CHJ_Center : xlCenter;

			// Vertikale Ausrichtung ermitteln
			if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_VCENTER ) )
				Cols[iCols].lVAlign = bOpenOffice ? CVJ_Center : xlCenter;
			else if ( MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_TXTALIGN_BOTTOM ) )
				Cols[iCols].lVAlign = bOpenOffice ? CVJ_Bottom : xlBottom;
			else
				Cols[iCols].lVAlign = bOpenOffice ? CVJ_Top : xlTop;

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
				if ( !MTblQueryColHdrGrpFlags( hWndTbl, Cols[iCols].pHdrGrp, MTBL_CHG_FLAG_NOCOLHEADERS ) )
					bAnyHdrGrpWCH = TRUE;
			}

			// Keine Zeilenlinien?
			if ( hWndCol == hWndTreeCol )
				Cols[iCols].bNoRowLines = MTblQueryTreeFlags( hWndTbl, MTBL_TREE_FLAG_NO_ROWLINES );
			else
				Cols[iCols].bNoRowLines = MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_NO_ROWLINES );

			// Keine Spaltenlinie?
			Cols[iCols].bNoColLine = MTblQueryColumnFlags(hWndCol, MTBL_COL_FLAG_NO_COLLINE) || Cols[iCols].VertLineDef.Style == MTLS_INVISIBLE;

			// Als Text exportieren?
			Cols[iCols].bExportAsText = MTblQueryColumnFlags( hWndCol, MTBL_COL_FLAG_EXPORT_AS_TEXT ) || ( bStringColsAsText && ( Cols[iCols].lDataType == DT_String || Cols[iCols].lDataType == DT_LongString ) );

			++iCols;
		}
		++iPos;

		// Abbruch?
		if ( IS_CANCELED )
		{
			iCols = GC_ERR_CANCELLED;
			break;
		}
	}

	return iCols;
}

void GetErrorMsg( LPCTSTR lpctsID, int iLanguage, CString & sMsg )
{
	CString sID = lpctsID;

	/*'if ( sID == ERR_COPY_TO_CLIPBOARD )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Kopieren in die Zwischenablage.") : _T("Error while copying to clipboard.");
	else if ( sID == ERR_COPY_TO_CLIPBOARD_OPEN )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Kopieren in die Zwischenablage: Zwischenablage konnte nicht geöffnet werden.") : _T("Error while copying to clipboard: Could not open clipboard.");
	else if ( sID == ERR_COPY_TO_CLIPBOARD_ALLOC )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Kopieren in die Zwischenablage: Globales Speicherobjekt konnte nicht alloziert werden.") : _T("Error while copying to clipboard: Could not allocate global memory object.");
	else if ( sID == ERR_COPY_TO_CLIPBOARD_SETDATA )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Kopieren in die Zwischenablage: Daten konnten nicht gesetzt werden.") : _T("Error while copying to clipboard: Could not set data.");
	else if ( sID == ERR_GET_FONT_INFO )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Ermitteln von Schriftartinformationen.") : _T("Error while getting font informations.");
	else if ( sID == ERR_GET_OBJECT )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Ermitteln eines Objektes.") : _T("Error while getting an object.");
	else if ( sID == ERR_GET_START_CELL )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Ermitteln der Ausgangszelle.") : _T("Error while getting the start cell.");
	else if ( sID == ERR_INSERT_SHEET )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Einfügen einer Tabelle.") : _T("Error while inserting a sheet.");
	else if ( sID == ERR_SAVE_CLIPBOARD )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Sichern der Zwischenablage.") : _T("Error while saving clipboard data.");
	else if ( sID == ERR_SAVE_CLIPBOARD_OPEN )
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Fehler beim Sichern der Zwischenablage: Zwischenablage konnte nicht geöffnet werden.") : _T("Error while saving clipboard data: Could not open clipboard.");
	else
		sMsg = iLanguage == MTE_LNG_GERMAN ? _T("Es ist ein unbekannter Fehler aufgetreten") : _T("An unknown error occured.");*/

	sMsg = _T("");
	GetUIStr( iLanguage, sID, sMsg );
}

void GetFontInfo( HWND hWnd, HFONT hFont, FONTINFO & fi )
{
	if ( hWnd && hFont )
	{
		TCHAR szBuf[256] = _T("");
		HDC hDC;
		HGDIOBJ hOldFont;
		int iPixPerInch;
		TEXTMETRIC tm;

		hDC = GetDC( hWnd );
		if ( hDC )
		{
			iPixPerInch = GetDeviceCaps( hDC, LOGPIXELSX );

			hOldFont = SelectObject( hDC, hFont );

			if ( !GetTextFace( hDC, sizeof( szBuf ), szBuf ) )
				ThrowError( ERR_GET_FONT_INFO );
			fi.sName = szBuf;

			if ( !GetTextMetrics( hDC, &tm ) )
				ThrowError( ERR_GET_FONT_INFO );

			fi.iPointSize = MulDiv(tm.tmHeight - tm.tmInternalLeading, 72, iPixPerInch);
			fi.bBold = ( tm.tmWeight >= 700 );
			fi.bItalic = BOOL( tm.tmItalic );
			fi.bUnderline = BOOL( tm.tmUnderlined );
			fi.bStrikeOut = BOOL( tm.tmStruckOut );

			SelectObject( hDC, hOldFont );

			ReleaseDC( hWnd, hDC );
		}
		else
			ThrowError( ERR_GET_FONT_INFO );
	}
	else
		ThrowError( ERR_GET_FONT_INFO );
}

void HandleException( COleDispatchException * pE, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage )
{
	CString sError( pE->m_strDescription );
	if ( sError.IsEmpty( ) )
	{
		//sError = ( iLanguage == MTE_LNG_GERMAN ? _T("Es ist ein unbekannter OLE-Dispatch-Fehler aufgetreten.") : _T("An unknown OLE dispatch exception occured.") );
		GetUIStr( iLanguage, _T("Exception.OleDispatch.Unknown"), sError );
		if ( pE->m_wCode )
		{
			TCHAR szCode[10] = _T("");
			_ultot( pE->m_wCode, szCode, 10 );			
			//sError += ( iLanguage == MTE_LNG_GERMAN ? _T(" (Fehler-Nr.:") : _T(" (Error Code:") );
			CString sTxt( _T("") );
			GetUIStr( iLanguage, _T("Exception.ErrorCode"), sTxt );
			sError += sTxt;
			sError += szCode;
			sError += _T(")");
		}
	}

	MsgBoxException( sError, pDlgThread, iApp, iLanguage );
}

void HandleException( COleException * pE, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage )
{
	TCHAR szBuf[2048] = _T("");
	pE->GetErrorMessage( szBuf, sizeof( szBuf ) );
	CString sError( szBuf );
	if ( sError.IsEmpty( ) )
	{
		//sError = ( iLanguage == MTE_LNG_GERMAN ? _T("Es ist ein unbekannter OLE-Fehler aufgetreten.") : _T("An unknown OLE exception occured.") );
		GetUIStr( iLanguage, _T("Exception.Ole.Unknown"), sError );
		WORD wCode = LOWORD( pE->m_sc );
		if ( wCode )
		{
			TCHAR szCode[10] = _T("");
			_ultot( wCode, szCode, 10 );
			//sError += ( iLanguage == MTE_LNG_GERMAN ? _T(" (Fehler-Nr.:") : _T(" (Error Code:") );
			CString sTxt( _T("") );
			GetUIStr( iLanguage, _T("Exception.ErrorCode"), sTxt );
			sError += sTxt;
			sError += szCode;
			sError += _T(")");
		}
	}

	MsgBoxException( sError, pDlgThread, iApp, iLanguage );
}

void HandleException( CException * pE, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage )
{
	TCHAR szBuf[2048] = _T("");
	pE->GetErrorMessage( szBuf, sizeof( szBuf ) );
	CString sError( szBuf );
	if ( sError.IsEmpty( ) )
		//sError = ( iLanguage == MTE_LNG_GERMAN ? _T("Es ist ein unbekannter Fehler aufgetreten.") : _T("An unknown exception occured.") );
		GetUIStr( iLanguage, _T("Exception.Unknown"), sError );

	MsgBoxException( sError, pDlgThread, iApp, iLanguage );
}

void HandleException( CString & sError, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage )
{
	MsgBoxException( sError, pDlgThread, iApp, iLanguage );
}

void MakeHtmlStr( CString & sText )
{
	if ( sText.IsEmpty( ) )
	{
		//sText = _T("<p>&#160;</p>");
		return;
	}
	else
	{
		BOOL bDBCSSecByte = FALSE;
		int i, j;
		CString sModText;
		for ( i = 0; i < sText.GetLength( ); i++ )
		{
			if ( IsDBCSLeadByte( sText[i] ) )
			{
				bDBCSSecByte = TRUE;
				sModText += sText[i];
			}
			else if ( bDBCSSecByte )
			{
				bDBCSSecByte = FALSE;
				sModText += sText[i];
			}
			else
			{
				switch ( sText[i] )
				{
				case _T(' '):					
					sModText += _T("&#160;");
					break;

				case _T('ä'):
					sModText += _T("&#228;");
					break;
				case _T('Ä'):
					sModText += _T("&#196;");
					break;
				case _T('ö'):
					sModText += _T("&#246;");
					break;
				case _T('Ö'):
					sModText += _T("&#214;");
					break;
				case _T('ü'):
					sModText += _T("&#252;");
					break;
				case _T('Ü'):
					sModText += _T("&#220;");
					break;
				case _T('ß'):
					sModText += _T("&#223;");
					break;
				case _T('<'):
					sModText += _T("&#60;");
					break;
				case _T('>'):
					sModText += _T("&#62;");
					break;
				case _T('&'):
					sModText += _T("&#38;");
					break;
				case _T('"'):
					sModText += _T("&#34");
					break;

				case _T('\t'):
					for ( j = 0; j < 8; j++ )
						sModText += _T("&#160;");
					break;
				case _T('\r'):
					sModText += _T("<br>");
					break;
				case _T('\n'):
					break;

				default:
					sModText += sText[i];
				}
			}
		}

		sText = sModText;
	}
}

int MsgBoxException( CString & sText, CExpTblDlgThread * pDlgThread, int iApp, int iLanguage )
{
	HWND hWndParent = ( pDlgThread ? pDlgThread->GetDlgHandle( ) : GetForegroundWindow( ) );
	CString sTitle = _T("");

	switch( iApp )
	{
	case APP_EXCEL:
		//sTitle = _T("Excel Export");
		GetUIStr( iLanguage, _T("Exception.Msg.Title.Excel"), sTitle );
		break;
	case APP_OOCALC:
		//sTitle = _T("Open Office Calc Export");
		GetUIStr( iLanguage, _T("Exception.Msg.Title.OOCalc"), sTitle );
		break;
	default:
		//sTitle = _T("Export");
		GetUIStr( iLanguage, _T("Exception.Msg.Title.Default"), sTitle );
	}
	
	//sTitle += ( iLanguage == MTE_LNG_GERMAN ? _T(" - Fehler") : _T(" - Error") );

	return MessageBox( hWndParent, sText, sTitle, MB_ICONERROR | MB_OK );
}

void SaveClipboard( CString & s, HWND hWnd /*= NULL*/ )
{
	if ( OpenClipboard( hWnd ) )
	{
#ifdef UNICODE
		UINT uiFormat = CF_UNICODETEXT;
#else
		UINT uiFormat = CF_TEXT;
#endif
		HGLOBAL hGlobal = GetClipboardData( uiFormat );
		if ( hGlobal )
		{
			DWORD dwLen = GlobalSize( hGlobal );
			LPTSTR pBuf = (LPTSTR)GlobalLock( hGlobal );
			LPTSTR pStrBuf = s.GetBuffer( dwLen );
			lstrcpy( pStrBuf, pBuf );
			s.ReleaseBuffer( -1 );
			GlobalUnlock( hGlobal );
			CloseClipboard( );
		}
		else
		{
			s = _T("");
			CloseClipboard( );
		}
	}
	else
	{
		ThrowError( ERR_SAVE_CLIPBOARD_OPEN );
	}
}

void ThrowError( LPCTSTR lpctsID )
{
	sError = lpctsID;
	throw sError;
}

void VMakeOptional( VARIANT & v )
{
	V_VT(&v) = VT_ERROR;
    V_ERROR(&v) = DISP_E_PARAMNOTFOUND;
}