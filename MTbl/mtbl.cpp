// mtbl.cpp

#include "mtbl.h"
#include "mtblsubclass.h"

BOOL gbCtdHooked;
HANDLE ghInstance;
LanguageMap glm;
SubClassMap gscm;
UINT guiMsgOffs = 100;
tstring gsEmpty(_T(""));

// Funktionscontainer
CCtd Ctd;
CVis Vis;

// Funktionsvariablen
FLATSB_SETSCROLLPOS FlatSB_SetScrollPos_Hooked = FlatSB_SetScrollPos;
FLATSB_SETSCROLLRANGE FlatSB_SetScrollRange_Hooked = FlatSB_SetScrollRange;
INVERTRECT InvertRect_Hooked = InvertRect;
PATBLT PatBlt_Hooked = PatBlt;
SCROLLWINDOW ScrollWindow_Hooked = ScrollWindow;
SETSCROLLPOS SetScrollPos_Hooked = SetScrollPos;
SETSCROLLRANGE SetScrollRange_Hooked = SetScrollRange;

// DllMain
BOOL APIENTRY DllMain( HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved )
{
	ghInstance = hModule;

	LanguageMap::iterator it;
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
			gbCtdHooked = FALSE;
#ifdef USE_APIHOOK
			gbCtdHooked = Hook( TRUE );
			if ( !gbCtdHooked )
			{
				MessageBox( NULL, _T("Could not implement all needed function hooks."), _T("MICSTO"), MB_OK | MB_ICONSTOP );
				return FALSE;
			}
#endif // USE_APIHOOK

			// Standardsprachen registrieren
			it = glm.find( MTBL_LNG_GERMAN );
			if ( it == glm.end( ) )
				MTblRegisterLanguage( MTBL_LNG_GERMAN, _T( "mtbllang.de" ), FALSE );

			it = glm.find( MTBL_LNG_ENGLISH );
			if ( it == glm.end( ) )
				MTblRegisterLanguage( MTBL_LNG_ENGLISH, _T( "mtbllang.en" ), FALSE );

			break;
		case DLL_THREAD_ATTACH:
			break;
		case DLL_THREAD_DETACH:
			break;
		case DLL_PROCESS_DETACH:
#ifdef USE_APIHOOK
			Hook( FALSE );
#endif // USE_APIHOOK
			break;
    }
    return TRUE;
}

// Interne Funktionen

#ifdef USE_APIHOOK

// Hook_alt
BOOL Hook( BOOL bHook )
{
	BOOL bOk = HookTabliXX( bHook );
	return bOk;
}

// HookTabliXX
BOOL HookTabliXX( BOOL bHook )
{
	//static const int HOOKFUNCS = 10;
	static HOOKDESC s_Hook[] =
	{
		{"COMCTL32.DLL", "FlatSB_SetScrollPos", (PROC)MyFlatSB_SetScrollPos },
		{"COMCTL32.DLL", "FlatSB_SetScrollRange", (PROC)MyFlatSB_SetScrollRange },
		{"USER32.DLL","InvertRect",(PROC)MyInvertRect},
		{"GDI32.DLL","PatBlt",(PROC)MyPatBlt},
		{"USER32.DLL", "ScrollWindow", (PROC)MyScrollWindow},
		{"USER32.DLL","SetScrollPos",(PROC)MySetScrollPos},
		{"USER32.DLL", "SetScrollRange", (PROC)MySetScrollRange }
	};
	static HOOKDESC s_UnHook[] =
	{
		{"COMCTL32.DLL", "FlatSB_SetScrollPos", NULL },
		{"COMCTL32.DLL", "FlatSB_SetScrollRange", NULL },
		{"USER32.DLL","InvertRect",NULL},
		{"GDI32.DLL","PatBlt",NULL},
		{"USER32.DLL", "ScrollWindow", NULL},
		{"USER32.DLL","SetScrollPos",NULL},
		{"USER32.DLL", "SetScrollRange", NULL }
	};

	static HMODULE s_hModuleTabli = NULL;

	// Get number of loaded modules
	DWORD dwLoadedModules = 0;
	DWORD dwProcessId = GetCurrentProcessId( );
	if ( !GetLoadedModules( dwProcessId, 0, NULL, &dwLoadedModules ) )
		return FALSE;

	// Allocate HMODULE array
	HMODULE *pModules = new HMODULE[dwLoadedModules];

	// Get loaded modules
	UINT uiCount = dwLoadedModules;
	if ( !GetLoadedModules( dwProcessId, uiCount, pModules, &dwLoadedModules ) )
	{
		delete []pModules;
		return FALSE;
	}

	// Find module handle of tablixx.dll if not already found in a previous call
	if ( !s_hModuleTabli )
	{
		TCHAR szDLL[MAX_PATH] = _T("DUMMY");
		#ifdef TD11
		lstrcpy( szDLL, _T("TABLI11.DLL") );
		#elif TD15
		lstrcpy( szDLL, _T("TABLI15.DLL") );
		#elif TD20
		lstrcpy( szDLL, _T("TABLI20.DLL") );
		#elif TD21
		lstrcpy( szDLL, _T("TABLI21.DLL") );
		#elif TD30
		lstrcpy( szDLL, _T("TABLI30.DLL") );
		#elif TD31
		lstrcpy( szDLL, _T("TABLI31.DLL") );
		#elif TD40
		lstrcpy( szDLL, _T("TABLI40.DLL") );
		#elif TD41
		lstrcpy( szDLL, _T("TABLI41.DLL") );
		#elif TD42
		lstrcpy( szDLL, _T("TABLI42.DLL") );
		#elif TD51
		lstrcpy( szDLL, _T("TABLI51.DLL") );
		#elif TD52
		lstrcpy( szDLL, _T("TABLI52.DLL") );
		#elif TD60
		lstrcpy( szDLL, _T("TABLI60.DLL") );
		#elif TD61
		lstrcpy( szDLL, _T("TABLI61.DLL") );
		#elif TD62
		lstrcpy( szDLL, _T("TABLI62.DLL") );
		#elif TD63
		lstrcpy( szDLL, _T("TABLI63.DLL") );
		#elif TD70
		lstrcpy( szDLL, _T("TABLI70.DLL") );
		#elif TD71
		lstrcpy(szDLL, _T("TABLI71.DLL"));
		#elif TD72
		lstrcpy(szDLL, _T("TABLI72.DLL"));
		#elif TD73
		lstrcpy(szDLL, _T("TABLI73.DLL"));
		#elif TD74
		lstrcpy(szDLL, _T("TABLI74.DLL"));
		#elif TD75
		lstrcpy(szDLL, _T("TABLI75.DLL"));
		#endif

		HANDLE hProcess = GetCurrentProcess();
		TCHAR szModuleName[MAX_PATH];
		for ( DWORD i = 0; i < dwLoadedModules; i++ )
		{
			if ( !GetModuleBaseNamePortable( hProcess, pModules[i], szModuleName, MAX_PATH ) )
			{
				delete []pModules;
				return FALSE;
			}

			if ( lstrcmpi( szDLL, szModuleName ) == 0 )
			{
				s_hModuleTabli = pModules[i];
				break;
			}
		}
	}

	// If found, hook or unhook...
	BOOL bHookError = FALSE;
	if ( s_hModuleTabli )
	{
		DWORD dwMy;
		int iHooks = sizeof(s_Hook) / sizeof(HOOKDESC);
		PROC  pProc;
		if ( bHook )
		{		
			// Hook
			for ( int i = 0; i < iHooks; i++ )
			{
				// Already My?
				if ( s_UnHook[i].HookFunc.pProc )
					continue;

				// Hook
				if (HookImportedFunctionsByName(s_hModuleTabli, s_Hook[i].szImpDll, 1, &(s_Hook[i].HookFunc), &pProc, &dwMy))
				{
					s_UnHook[i].HookFunc.pProc = pProc;

					if (pProc)
					{
						if (strcmp(s_Hook[i].HookFunc.szFunc, "FlatSB_SetScrollPos") == 0)
							FlatSB_SetScrollPos_Hooked = (FLATSB_SETSCROLLPOS)pProc;
						else if (strcmp(s_Hook[i].HookFunc.szFunc, "FlatSB_SetScrollRange") == 0)
							FlatSB_SetScrollRange_Hooked = (FLATSB_SETSCROLLRANGE)pProc;
						else if (strcmp(s_Hook[i].HookFunc.szFunc, "InvertRect") == 0)
							InvertRect_Hooked = (INVERTRECT)pProc;
						else if (strcmp(s_Hook[i].HookFunc.szFunc, "PatBlt") == 0)
							PatBlt_Hooked = (PATBLT)pProc;
						else if (strcmp(s_Hook[i].HookFunc.szFunc, "ScrollWindow") == 0)
							ScrollWindow_Hooked = (SCROLLWINDOW)pProc;
						else if (strcmp(s_Hook[i].HookFunc.szFunc, "SetScrollPos") == 0)
							SetScrollPos_Hooked = (SETSCROLLPOS)pProc;
						else if (strcmp(s_Hook[i].HookFunc.szFunc, "SetScrollRange") == 0)
							SetScrollRange_Hooked = (SETSCROLLRANGE)pProc;
					}
				}
					
				else
					bHookError = TRUE;
			}
			
		}
		else
		{
			// Unhook
			for (int i = 0; i < iHooks; i++)
			{
				// Not My?
				if ( !s_UnHook[i].HookFunc.pProc )
					continue;

				// Unhook
				if ( HookImportedFunctionsByName( s_hModuleTabli, s_Hook[i].szImpDll, 1, &(s_UnHook[i].HookFunc), &pProc, &dwMy ) )
					s_UnHook[i].HookFunc.pProc = NULL;
				else
					bHookError = TRUE;
			}
		}
	}
	else
	{
		bHookError = TRUE;
		SetLastError( ERROR_MOD_NOT_FOUND );
	}
	
	delete []pModules;
	return !bHookError;
}

// MyFlatSB_SetScrollPos
int WINAPI MyFlatSB_SetScrollPos(HWND hWnd, int fnBar, int nPos, BOOL fRedraw)
{
	HWND hWndTbl = NULL;
	LPMTBLSUBCLASS psc = NULL;

	if (fnBar == SB_CTL)
		hWndTbl = GetParent(hWnd);
	else
		hWndTbl = hWnd;
	psc = MTblGetSubClass(hWndTbl);

	if (psc && psc->m_bNoVScrolling && psc->IsScrollBarV(hWnd, fnBar))
	{
		return FlatSB_GetScrollPos(hWnd, fnBar);
	}

	int iRet = FlatSB_SetScrollPos_Hooked(hWnd, fnBar, nPos, fRedraw);

	return iRet;
}

// MyFlatSB_SetScrollRange
int WINAPI MyFlatSB_SetScrollRange(HWND hWnd, int nBar, int nMinPos, int nMaxPos, BOOL bRedraw)
{
	HWND hWndTbl = NULL;
	LPMTBLSUBCLASS psc = NULL;

	if (nBar == SB_CTL)
		hWndTbl = GetParent(hWnd);
	else
		hWndTbl = hWnd;
	psc = MTblGetSubClass(hWndTbl);

	if (psc)
	{
		int iScrBarV = psc->IsScrollBarV(hWnd, nBar);
		if (iScrBarV == 1)
			psc->m_iVScrollMaxCtd = nMaxPos;
		else if (iScrBarV == 2)
			psc->m_iVScrollMaxSplitCtd = nMaxPos;

		if (psc->m_bNoVScrolling && iScrBarV)
			return TRUE;
		else
			return psc->OnSetScrollRange(hWnd, nBar, nMinPos, nMaxPos, bRedraw);
	}
	else
		return FlatSB_SetScrollRange_Hooked(hWnd, nBar, nMinPos, nMaxPos, bRedraw);
}

// MyInvertRect
BOOL WINAPI MyInvertRect( HDC hDC, CONST RECT *lprc )
{
	LPMTBLSUBCLASS psc = MTblGetSubClass( WindowFromDC( hDC ) );
	if ( psc )
		return TRUE;

	BOOL bRet = InvertRect_Hooked( hDC, lprc );
	return bRet;
}

// MyPatBlt
BOOL WINAPI MyPatBlt(HDC hDC, int iXLeft, int iYLeft, int iWidth, int iHeight, DWORD dwRop )
{
	LPMTBLSUBCLASS psc = MTblGetSubClass( WindowFromDC( hDC ) );
	if ( psc && iWidth > 1 && iHeight > 1 && dwRop == DSTINVERT )
		return TRUE;

	BOOL bRet = PatBlt_Hooked( hDC, iXLeft, iYLeft, iWidth, iHeight, dwRop );
	return bRet;
}

// MyScrollWindow
BOOL WINAPI MyScrollWindow( HWND hWnd, int XAmount, int YAmount, CONST RECT *lpRect, CONST RECT *lpClipRect )
{
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWnd );

	if ( psc && psc->m_bNoVScrolling )
		YAmount = 0;

	if ( psc && psc->m_bCellMode && ( psc->m_bColMoving || psc->m_bColSizing ) )
	{
		psc->m_bNoFocusPaint = TRUE;
		psc->UpdateFocus( );
	}

	BOOL bRet = TRUE;
	if ( XAmount || YAmount )
		bRet = ScrollWindow_Hooked( hWnd, XAmount, YAmount, lpRect, lpClipRect );

	return bRet;
}

// MySetScrollPos
int WINAPI MySetScrollPos( HWND hWnd, int fnBar, int nPos, BOOL fRedraw )
{
	HWND hWndTbl = NULL;
	LPMTBLSUBCLASS psc = NULL;

	if ( fnBar == SB_CTL )
		hWndTbl = GetParent( hWnd );
	else
		hWndTbl = hWnd;
	psc = MTblGetSubClass( hWndTbl );

	if ( psc && psc->m_bNoVScrolling && psc->IsScrollBarV( hWnd, fnBar ) )
	{
		return GetScrollPos(hWnd, fnBar);
	}

	int iRet = SetScrollPos_Hooked(hWnd, fnBar, nPos, fRedraw);

	return iRet;
}

// MySetScrollRange
BOOL WINAPI MySetScrollRange( HWND hWnd, int nBar, int nMinPos, int nMaxPos, BOOL bRedraw )
{
	HWND hWndTbl = NULL;
	LPMTBLSUBCLASS psc = NULL;

	if ( nBar == SB_CTL )
		hWndTbl = GetParent( hWnd );
	else
		hWndTbl = hWnd;
	psc = MTblGetSubClass( hWndTbl );

	if ( psc )
	{
		int iScrBarV = psc->IsScrollBarV( hWnd, nBar );
		if ( iScrBarV == 1 )
			psc->m_iVScrollMaxCtd = nMaxPos;
		else if ( iScrBarV == 2 )
			psc->m_iVScrollMaxSplitCtd = nMaxPos;

		if ( psc->m_bNoVScrolling && iScrBarV )
			return TRUE;
		else
			return psc->OnSetScrollRange( hWnd, nBar, nMinPos, nMaxPos, bRedraw );
	}
	else
		return SetScrollRange_Hooked(hWnd, nBar, nMinPos, nMaxPos, bRedraw);
}

#endif // USE_APIHOOK

// ********************
// *****MTBLBTNDEF*****
// ********************
BOOL MTBLBTNDEF::FromUdv( HUDV hUdv )
{
	// UDV dereferenzieren
	LPMTBLBTNDEFUDV pUdv = (LPMTBLBTNDEFUDV)Ctd.UdvDeref( hUdv );
	if ( !pUdv )
		return FALSE;

	// Werte setzen
	this->bVisible = Ctd.NumToBool( pUdv->Visible );
	this->iWidth = Ctd.NumToInt( pUdv->Width );
#ifdef _WIN64
	this->hImage = (HANDLE)Ctd.NumToULongLong( pUdv->Image );
#else
	this->hImage = (HANDLE)Ctd.NumToULong( pUdv->Image );
#endif
	this->dwFlags = Ctd.NumToULong( pUdv->Flags );

	BOOL bOk = TRUE;
	long lLen;
	LPTSTR lps = Ctd.StringGetBuffer(pUdv->Text, &lLen);
	if (lps)
		this->sText = lps;
	else
	{
		this->sText = _T("");
		bOk = FALSE;
	}
	Ctd.HStrUnlock(pUdv->Text);
			

	// Fertig
	return bOk;
}

BOOL MTBLBTNDEF::ToUdv( HUDV hUdv )
{
	// UDV dereferenzieren
	LPMTBLBTNDEFUDV pUdv = (LPMTBLBTNDEFUDV)Ctd.UdvDeref( hUdv );
	if ( !pUdv )
		return FALSE;

	// Werte setzen
	Ctd.CvtLongToNumber( this->bVisible ? 1 : 0, &pUdv->Visible );
	Ctd.CvtIntToNumber( this->iWidth, &pUdv->Width );
	Ctd.CvtULongToNumber( (ULONG)this->hImage, &pUdv->Image );
	Ctd.CvtULongToNumber( this->dwFlags, &pUdv->Flags );

	Ctd.InitLPHSTRINGParam( &pUdv->Text, Ctd.BufLenFromStrLen( lstrlen( this->sText.c_str( ) ) ) );

	BOOL bOk = TRUE;
	long lLen;
	LPTSTR pBuf = Ctd.StringGetBuffer( pUdv->Text, &lLen );
	if (pBuf)
		lstrcpy(pBuf, this->sText.c_str());
	else
		bOk = FALSE;
	Ctd.HStrUnlock(pUdv->Text);

	// Fertig
	return bOk;
}

// ********************
// ****MTBLCELLRANGE***
// ********************

// CopyTo
void MTBLCELLRANGE::CopyTo( MTBLCELLRANGE &cr )
{
	cr.hWndColFrom = hWndColFrom;
	cr.lRowFrom = lRowFrom;

	cr.hWndColTo = hWndColTo;
	cr.lRowTo = lRowTo;
}

// GetDiff
int MTBLCELLRANGE::GetDiff( HWND hWndTbl,  MTBLCELLRANGE &cr1,  MTBLCELLRANGE &cr2, MTBLCELLRANGE craDiff[4] )
{
	int i = 0;
	if ( !cr1.IsEqualTo( cr2 ) )
	{
		BOOL bCellIn1, bCellIn2, bColIn1, bColIn2, bExpandRange, bFirstNotEqualFound, bRowIn1, bRowIn2;
		int iColPos, iColPosFrom, iColPosFrom1, iColPosFrom2, iColPosFromR, iColPosTo, iColPosTo1, iColPosTo2, iColPosToR;
		long lRow, lRowFrom, lRowFromR, lRowTo, lRowToR;

		iColPosFrom1 = cr1.GetColFromPos( );
		iColPosTo1 = cr1.GetColToPos( );

		iColPosFrom2 = cr2.GetColFromPos( );
		iColPosTo2 = cr2.GetColToPos( );

		iColPosFrom = min( iColPosFrom1, iColPosFrom2 );
		iColPosTo = max( iColPosTo1, iColPosTo2 );

		lRowFrom = min( cr1.lRowFrom, cr2.lRowFrom );
		lRowTo = max( cr1.lRowTo, cr2.lRowTo );

		bFirstNotEqualFound = FALSE;
		lRow = lRowFrom;
		while ( lRow <= lRowTo )
		{
			bRowIn1 = ( lRow >= cr1.lRowFrom && lRow <= cr1.lRowTo );
			bRowIn2 = ( lRow >= cr2.lRowFrom && lRow <= cr2.lRowTo );

			iColPos = iColPosFrom;
			while ( iColPos <= iColPosTo )
			{
				bColIn1 = ( iColPos >= iColPosFrom1 && iColPos <= iColPosTo1 );
				bColIn2 = ( iColPos >= iColPosFrom2 && iColPos <= iColPosTo2 );

				bCellIn1 = ( bRowIn1 && bColIn1 );
				bCellIn2 = ( bRowIn2 && bColIn2 );

				if ( bCellIn1 != bCellIn2 )
				{
					if ( !bFirstNotEqualFound )
					{
						iColPosFromR = iColPosToR = iColPos;
						lRowFromR = lRowToR = lRow;
						bFirstNotEqualFound = TRUE;
					}
					else
					{
						bExpandRange = FALSE;
						if ( lRow == lRowToR )
						{
							bExpandRange = iColPos == iColPosToR + 1;
						}
						else if ( lRow == lRowToR + 1 )
						{
							if ( iColPos >= iColPosFromR && iColPos <= iColPosToR )
							{
								bExpandRange = TRUE;
							}
						}

						if ( bExpandRange )
						{
							iColPosToR = iColPos;
							lRowToR = lRow;
						}
						else
						{
							iColPosFromR = iColPosToR = iColPos;
							lRowFromR = lRowToR = lRow;
						}
					}
				}
				else
					break;
			}
		}
		
	}

	return i;
}

// IsEqualTo
BOOL MTBLCELLRANGE::IsEqualTo( MTBLCELLRANGE &cr )
{
	if ( hWndColFrom != cr.hWndColFrom )
		return FALSE;

	if ( lRowFrom != cr.lRowFrom )
		return FALSE;

	if ( hWndColTo != cr.hWndColTo )
		return FALSE;

	if ( lRowTo != cr.lRowTo )
		return FALSE;

	return TRUE;
}

// IsValid
BOOL MTBLCELLRANGE::IsValid( )
{
	if ( !hWndColFrom || !hWndColTo || lRowFrom == TBL_Error || lRowTo == TBL_Error )
		return FALSE;

	if ( ( lRowFrom >= 0 && lRowTo < 0 ) || ( lRowFrom < 0 && lRowTo >= 0 ) )
		return FALSE;

	return TRUE;
}

// Normalize
void MTBLCELLRANGE::Normalize( )
{
	if ( Ctd.TblQueryColumnPos( hWndColFrom ) > Ctd.TblQueryColumnPos( hWndColTo ) )
	{
		HWND hWndTmp = hWndColFrom;
		hWndColFrom = hWndColTo;
		hWndColTo = hWndTmp;
	}

	if ( lRowFrom > lRowTo )
	{
		long lTmp = lRowFrom;
		lRowFrom = lRowTo;
		lRowTo = lTmp;
	}
}

// ********************
// ****MTBLPAINTATA***
// ********************

LPMTBLPAINTCOL MTBLPAINTDATA::EnsureValidPaintCol( int iIndex )
{
	if ( iIndex < 0 || iIndex >= MTBL_MAX_COLS )
		return NULL;

	if ( !PaintCols[iIndex] )
		PaintCols[iIndex] = new MTBLPAINTCOL;

	return PaintCols[iIndex];
}

MTBLPAINTDATA::MTBLPAINTDATA( )
{
	this->hDC = NULL;
	this->hDCTbl = NULL;
	this->hDCDoubleBuffer = NULL;
	this->hBitmDoubleBuffer = NULL;
	this->hOldBitmDoubleBuffer = NULL;
	this->hBrTblBack = NULL;
	SetRectEmpty( &this->rUpd );
	this->clrTblBackColor = MTBL_COLOR_UNDEF;
	this->clrTblTextColor = 0;
	this->clrHdrBack = 0;
	this->clrHdrLine = 0;
	this->clr3DHiLight = 0;
	this->lRow = 0;
	this->lEffPaintRow = 0;
	this->pRow = NULL;
	this->pEffPaintRow = NULL;
	this->pFirstMergeRow = NULL;
	this->pLastMergeRow = NULL;
	this->iRowLevel = 0;
	this->bRowSelected = FALSE;
	ZeroMemory( this->PaintCols, sizeof( LPMTBLPAINTCOL ) * MTBL_MAX_COLS );
	this->pPaintCol = NULL;
	this->pEffPaintCol = NULL;
	this->pFirstMergeCol = NULL;
	this->pLastMergeCol = NULL;
	this->pFirstGroupColHdr = NULL;
	this->pLastGroupColHdr = NULL;
	this->iCols = 0;
	this->iNormalCols = 0;
	ZeroMemory( &this->pc, sizeof( MTBLPAINTCELL ) );
	this->ppmc = NULL;
	this->ppmrh = NULL;
	this->pmr = NULL;
	this->lMinSplitRow = 0;
	this->lMaxSplitRow = 0;
	this->lFocusRow = 0;
	this->hWndFocusCol = NULL;
	this->bColHdr = FALSE;
	this->bEmptyRow = FALSE;
	this->bGrayedHeaders = FALSE;
	this->bSplit = FALSE;
	this->bNoRowLines = FALSE;
	this->bSuppressLastColLine = FALSE;
	this->bSplitRow = FALSE;
	this->bAnyRow = FALSE;
	this->bAnySplitRow = FALSE;
	this->wRowHdrFlags = 0;
	this->bRowHdrMirror = FALSE;
	this->hWndRowHdrCol = NULL;
	this->bRowHdrColWordWrap = FALSE;
	this->lRowHdrColJustify = 0;
	this->iScrPosV = 0;
	this->iScrPosVSplit = 0;
	this->dwRowTop = 0;
	this->dwRowBottom = 0;
	this->dwRowTopLog = 0;
	this->dwRowBottomLog = 0;
	this->dwRowTopVis = 0;
	this->dwRowBottomVis = 0;
	this->hTheme = NULL;
	this->bAnySelColors = FALSE;
	this->bClipFocusArea = FALSE;
	this->bPaintFocus = FALSE;
	this->iPaintColFocusCell = 0;
	SetRectEmpty( &this->rFocusClip );
}


// ********************
// ****MTBLPAINTCELL***
// ********************

BOOL MTBLPAINTCELL::IsCellTypeCheckedVal( LPCTSTR lps )
{
	if ( !pCellType || !lps )
		return FALSE;

	if ( pCellType->m_dwFlags & COL_CheckBox_IgnoreCase )
		return lstrcmpi( lps, pCellType->m_sCheckedVal.c_str( ) ) == 0;
	else
		return lstrcmp( lps, pCellType->m_sCheckedVal.c_str( ) ) == 0;
}


// ********************
// ****MTBLROWRANGE***
// ********************

// IsValid
BOOL MTBLROWRANGE::IsValid( )
{
	if ( lRowFrom == TBL_Error || lRowTo == TBL_Error )
		return FALSE;

	if ( ( lRowFrom >= 0 && lRowTo < 0 ) || ( lRowFrom < 0 && lRowTo >= 0 ) )
		return FALSE;

	return TRUE;
}

// Normalize
void MTBLROWRANGE::Normalize( )
{
	if ( lRowFrom > lRowTo )
	{
		long lTmp = lRowFrom;
		lRowFrom = lRowTo;
		lRowTo = lTmp;
	}
}

// ****************
// ***** MTbl *****
// ****************

// MTblApplyProperties
BOOL WINAPI MTblApplyProperties( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	const int TMP = 24;

	BOOL baProp[25];
	BOOL baVal[25];
	DWORD dwaVal[25];
	HANDLE hImg, hImg1, hImg2;
	HSTRING hsaVal[25];
	HWND hWndaVal[25];
	int iaVal[25];
	NUMBER naVal[25];
	tstring sBoolNo = MTPROP_VAL_BOOL_NO;
	tstring sBoolYes = MTPROP_VAL_BOOL_YES;
	vector<tstring> saVal(25);

	memset( hsaVal, 0, sizeof( hsaVal ) );

	// (General)
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_EXT_MSGS"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableExtMsgs( hWndTbl, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_EXT_MSGS_CORNERSEP_EXCL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetExtMsgsOptions( hWndTbl, MTEM_CORNERSEP_EXCLUSIVE, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_EXT_MSGS_COLHDRSEP_EXCL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetExtMsgsOptions( hWndTbl, MTEM_COLHDRSEP_EXCLUSIVE, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_EXT_MSGS_ROWHDRSEP_EXCL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetExtMsgsOptions( hWndTbl, MTEM_ROWHDRSEP_EXCLUSIVE, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_PWD_CHAR"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		MTblSetPasswordChar( hWndTbl, saVal[0].c_str( ) );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ADAPT_LIST_WIDTH"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_ADAPT_LIST_WIDTH, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_EXPAND_CRLF"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_EXPAND_CRLF, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_EXPAND_TAB"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_EXPAND_TABS, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_IGNORE_TD_THEME_COLORS"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_IGNORE_TD_THEME_COLORS, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_IGNORE_WIN_THEME"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_IGNORE_WIN_THEME, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_NO_DITHER"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_NO_DITHER, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_USE_TBL_FONT_DDL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_USE_TBL_FONT_DDL, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_USE_TBL_FONT_PE"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_USE_TBL_FONT_PE, baVal[0] );
	}

	// Buttons
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_BUTTONS_FLAT"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_BUTTONS_FLAT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_BUTTONS_OUTSIDE_CELL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_BUTTONS_OUTSIDE_CELL, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_BUTTONS_PERMANENT"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_BUTTONS_PERMANENT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_HIDE_PERM_BTN_NE"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_HIDE_PERM_BTN_NE, baVal[0] );
	}

	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_BTN_HOTKEY"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_BTN_HOTKEY_CTRL"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_BTN_HOTKEY_SHIFT"), &hsaVal[2] );
	baProp[3] = Ctd.WindowGetProperty( hWndTbl, _T("MT_BTN_HOTKEY_ALT"), &hsaVal[3] );
	if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] )
	{
		iaVal[0] = baProp[0] ? Ctd.HStrToInt( hsaVal[0] ) : 0;
		baVal[1] = baProp[1] ? MTblPropStrToBool( hsaVal[1] ) : 0;
		baVal[2] = baProp[2] ? MTblPropStrToBool( hsaVal[2] ) : 0;
		baVal[3] = baProp[3] ? MTblPropStrToBool( hsaVal[3] ) : 0;

		MTblDefineButtonHotkey( hWndTbl, iaVal[0], baVal[1], baVal[2], baVal[3] );
	}

	// Cell Mode
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_CM_ACTIVE"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_CM_SINGLE_SELECTION"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_CM_AUTO_EDIT_MODE"), &hsaVal[2] );
	baProp[3] = Ctd.WindowGetProperty( hWndTbl, _T("MT_CM_NO_SELECTION"), &hsaVal[3] );
	if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] )
	{
		baVal[0] = baProp[0] ? MTblPropStrToBool( hsaVal[0] ) : FALSE;

		dwaVal[0] = baProp[1] && MTblPropStrToBool( hsaVal[1] ) ? MTBL_CM_FLAG_SINGLE_SELECTION : 0;
		dwaVal[0] |= baProp[2] && MTblPropStrToBool( hsaVal[2] ) ? MTBL_CM_FLAG_AUTO_EDIT_MODE : 0;
		dwaVal[0] |= baProp[3] && MTblPropStrToBool( hsaVal[3] ) ? MTBL_CM_FLAG_NO_SELECTION : 0;

		MTblDefineCellMode( hWndTbl, baVal[0], dwaVal[0] );

	}

	// Effective Cell Property Determination
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_ECPD_ORDER"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_ECPD_COMB_FONT_ENH"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		dwaVal[0] = baProp[0] ? Ctd.HStrToULong( hsaVal[0] ) : MTECPD_ORDER_CELL_COL_ROW;
		dwaVal[1] = baProp[1] && MTblPropStrToBool( hsaVal[1] ) ? MTECPD_FLAG_COMB_FONT_ENH : 0;

		MTblDefineEffCellPropDet( hWndTbl, dwaVal[0], dwaVal[1] );

	}

	// Headers
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_GRADIENT_HEADERS"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_GRADIENT_HEADERS, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_HEADERS_BACK_COLOR"), &hsaVal[0] ) )
	{
		dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
		MTblSetHeadersBackColor( hWndTbl, dwaVal[0], MTSC_REDRAW );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_COLHDR_FREE_BACK_COLOR"), &hsaVal[0] ) )
	{
		dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
		MTblSetColumnHdrFreeBackColor( hWndTbl, dwaVal[0], MTSC_REDRAW );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_SOLID_ROWHDR"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_SOLID_ROWHDR, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_CORNER_BACK_COLOR"), &hsaVal[0] ) )
	{
		dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
		MTblSetCornerBackColor( hWndTbl, dwaVal[0], MTSC_REDRAW );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_CORNER_OWNERDRAWN"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetOwnerDrawnItem( hWndTbl, TBL_Error, MTBL_ODI_CORNER, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_NO_COLUMN_HEADER"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_NO_COLUMN_HEADER, baVal[0] );
	}

	// Highlighting
	typedef list< pair<LONG_PAIR,tstring> > HILI_LIST;
	HILI_LIST hl;
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_CELL, MTBL_PART_UNDEF ), _T("MT_HILI_CELL_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_COLUMN, MTBL_PART_UNDEF ), _T("MT_HILI_COL_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_COLHDR, MTBL_PART_UNDEF ), _T("MT_HILI_COLHDR_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_COLHDRGRP, MTBL_PART_UNDEF ), _T("MT_HILI_COLHDRPRP_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_ROW, MTBL_PART_UNDEF ), _T("MT_HILI_ROW_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_ROWHDR, MTBL_PART_UNDEF ), _T("MT_HILI_ROWHDR_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_CORNER, MTBL_PART_UNDEF ), _T("MT_HILI_CORNER_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_CELL, MTBL_PART_NODE ), _T("MT_HILI_NODE_") ) );
	hl.push_back( HILI_LIST::value_type( LONG_PAIR( MTBL_ITEM_ROWHDR, MTBL_PART_NODE ), _T("MT_HILI_NODE_") ) );

	int iItemPart, iItemType;
	MTBLHILI hili;
	tstring sPrefix, sProp;
	for ( HILI_LIST::iterator it = hl.begin( ); it != hl.end( ); it++ )
	{
		hili.Init( );

		iItemType = (*it).first.first;
		iItemPart = (*it).first.second;
		sPrefix = (*it).second;

		// Background color
		sProp = sPrefix;
		sProp += _T("BACK_COLOR");
		if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[0] ) )
		{
			hili.dwBackColor = Ctd.HStrToULong( hsaVal[0] );
		}

		// Text color
		sProp = sPrefix;
		sProp += _T("TEXT_COLOR");
		if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[0] ) )
		{
			hili.dwTextColor = Ctd.HStrToULong( hsaVal[0] );
		}

		// Font enhancements
		sProp = sPrefix;
		sProp += _T("FONT_ENH_BOLD");
		baProp[0] = Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[0] );
		
		sProp = sPrefix;
		sProp += _T("FONT_ENH_ITALIC");
		baProp[1] = Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[1] );

		sProp = sPrefix;
		sProp += _T("FONT_ENH_UNDERLINE");
		baProp[2] = Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[2] );

		sProp = sPrefix;
		sProp += _T("FONT_ENH_STRIKEOUT");
		baProp[3] = Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[3] );
		
		if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] )
		{
			
			iaVal[0] = 0;
			if ( baProp[0] && MTblPropStrToBool( hsaVal[0] ) )
				iaVal[0] |= FONT_EnhBold;
			if ( baProp[1] && MTblPropStrToBool( hsaVal[1] ) )
				iaVal[0] |= FONT_EnhItalic;
			if ( baProp[2] && MTblPropStrToBool( hsaVal[2] ) )
				iaVal[0] |= FONT_EnhUnderline;
			if ( baProp[3] && MTblPropStrToBool( hsaVal[3] ) )
				iaVal[0] |= FONT_EnhStrikeOut;			

			if ( iaVal[0] )
				hili.lFontEnh = iaVal[0];
		}

		// Image
		sProp = sPrefix;
		sProp += _T("IMG_SOURCE");
		if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[0] ) )
		{
			saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );

			sProp = sPrefix;
			sProp += _T("IMG_NAME");
			if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[1] ) )
			{
				saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
				hImg = NULL;
				if ( saVal[0] == _T("FILE") )
				{
					hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
				}
				else if ( saVal[0] == _T("TDRES") )
				{
					hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
				}

				if ( hImg )
				{
					sProp = sPrefix;
					sProp += _T("IMG_TRANSP_COLOR");
					if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[2] ) )
					{
						dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
						MImgSetTranspColor( hImg, dwaVal[2] );
					}

					hili.Img.SetHandle( hImg, NULL );
				}
			}
		}

		// Image expanded
		sProp = sPrefix;
		sProp += _T("IMGEXP_SOURCE");
		if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[0] ) )
		{
			saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );

			sProp = sPrefix;
			sProp += _T("IMGEXP_NAME");
			if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[1] ) )
			{
				saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
				hImg = NULL;
				if ( saVal[0] == _T("FILE") )
				{
					hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
				}
				else if ( saVal[0] == _T("TDRES") )
				{
					hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
				}

				if ( hImg )
				{
					sProp = sPrefix;
					sProp += _T("IMGEXP_TRANSP_COLOR");
					if ( Ctd.WindowGetProperty( hWndTbl, (LPTSTR)sProp.c_str( ), &hsaVal[2] ) )
					{
						dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
						MImgSetTranspColor( hImg, dwaVal[2] );
					}

					hili.ImgExp.SetHandle( hImg, NULL );
				}
			}
		}


		// Setzen
		if ( !hili.IsUndef( ) )
		{
			psc->m_pHLDM->insert( HiLiDefMap::value_type( (*it).first, hili ) );
		}
	}

	// Row Flag Images
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_EDITED_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_EDITED_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_EDITED_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_EDITED_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_Edited, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_MARKDELETED_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_MARKDELETED_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_MARKDELETED_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_MARKDELETED_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_MarkDeleted, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_NEW_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_NEW_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_NEW_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_NEW_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_New, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG1_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG1_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG1_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG1_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_UserFlag1, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG2_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG2_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG2_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG2_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_UserFlag2, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG3_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG3_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG3_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG3_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_UserFlag3, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG4_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG4_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG4_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG4_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_UserFlag4, hImg, dwaVal[3] );
			}
		}
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG5_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG5_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			hImg = NULL;
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG5_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}

				dwaVal[3] = 0;
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_USERFLAG5_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
					dwaVal[3] |= MTSI_NOSELINV;
				dwaVal[3] |= MTSI_REDRAW;

				MTblSetRowFlagImage( hWndTbl, ROW_UserFlag5, hImg, dwaVal[3] );
			}
		}
	}

	// Hyperlink HOVER
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_HYPHO_ENABLED"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_HYPHO_FONT_ENH_BOLD"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_HYPHO_FONT_ENH_ITALIC"), &hsaVal[2] );
	baProp[3] = Ctd.WindowGetProperty( hWndTbl, _T("MT_HYPHO_FONT_ENH_STRIKEOUT"), &hsaVal[3] );
	baProp[4] = Ctd.WindowGetProperty( hWndTbl, _T("MT_HYPHO_FONT_ENH_UNDERLINE"), &hsaVal[4] );
	baProp[5] = Ctd.WindowGetProperty( hWndTbl, _T("MT_HYPHO_TEXT_COLOR"), &hsaVal[5] );
	if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] || baProp[4] || baProp[5] )
	{
		baVal[0] = baProp[0] ? MTblPropStrToBool( hsaVal[0] ) : FALSE;

		dwaVal[1] = 0;
		if ( baProp[1] && MTblPropStrToBool( hsaVal[1] ) )
			dwaVal[1] |= FONT_EnhBold;
		if ( baProp[2] && MTblPropStrToBool( hsaVal[2] ) )
			dwaVal[1] |= FONT_EnhItalic;
		if ( baProp[3] && MTblPropStrToBool( hsaVal[3] ) )
			dwaVal[1] |= FONT_EnhStrikeOut;
		if ( baProp[4] && MTblPropStrToBool( hsaVal[4] ) )
			dwaVal[1] |= FONT_EnhUnderline;

		dwaVal[2] = baProp[5] ? Ctd.HStrToULong( hsaVal[5] ) : MTBL_COLOR_UNDEF;

		MTblDefineHyperlinkHover( hWndTbl, baVal[0], dwaVal[1], dwaVal[2] );
	}

	// Columns
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_NO_LINES_FREE_COL_AREA"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_NO_FREE_COL_AREA_LINES, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_COLOR_ENTIRE_COLUMN"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_COLOR_ENTIRE_COLUMN, baVal[0] );
	}
	
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_COL_LINES_STYLE"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_COL_LINES_COLOR"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		if ( baProp[0] )
			iaVal[0] = Ctd.HStrToInt( hsaVal[0] );
		else
		{
#ifdef TD_SOLID_TBL_LINES
			iaVal[0] = MTLS_SOLID;
#else
			iaVal[0] = MTLS_DOT;
#endif
		}

		dwaVal[1] = baProp[1] ? Ctd.HStrToULong( hsaVal[1] ) : MTBL_COLOR_UNDEF;

		MTblDefineColLines( hWndTbl, iaVal[0], dwaVal[1] );
	}

	
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_LLCOL_LINE_STYLE"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_LLCOL_LINE_COLOR"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		iaVal[0] = baProp[0] ? Ctd.HStrToInt( hsaVal[0] ) : MTLS_SOLID;
		dwaVal[1] = baProp[1] ? Ctd.HStrToULong( hsaVal[1] ) : MTBL_COLOR_UNDEF;

		MTblDefineLastLockedColLine( hWndTbl, iaVal[0], dwaVal[1] );
	}

	// Edit Mode
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_MOVE_INP_FOCUS_UD_EX"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_MOVE_INP_FOCUS_UD_EX, baVal[0] );
	}

	// Rows
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_VARIABLE_ROW_HEIGHT"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_VARIABLE_ROW_HEIGHT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_NO_LINES_FREE_ROW_AREA"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_NO_FREE_ROW_AREA_LINES, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_COLOR_ENTIRE_ROW"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_COLOR_ENTIRE_ROW, baVal[0] );
	}

	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_LINES_STYLE"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_ROW_LINES_COLOR"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		if ( baProp[0] )
			iaVal[0] = Ctd.HStrToInt( hsaVal[0] );
		else
		{
#ifdef TD_SOLID_TBL_LINES
			iaVal[0] = MTLS_SOLID;
#else
			iaVal[0] = MTLS_DOT;
#endif
		}

		dwaVal[1] = baProp[1] ? Ctd.HStrToULong( hsaVal[1] ) : MTBL_COLOR_UNDEF;

		MTblDefineRowLines( hWndTbl, iaVal[0], dwaVal[1] );
	}

	// Suppress Focus
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_SUPPRESS_FOCUS"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_SUPPRESS_FOCUS, baVal[0] );
	}

	// Clip Row Focus Area
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_CLIP_ROW_FOCUS_AREA"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_CLIP_FOCUS_AREA, baVal[0] );
	}

	// Focus Frame
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_FOCUS_FRAME_THICKNESS"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_FOCUS_FRAME_OFFSET_X"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_FOCUS_FRAME_OFFSET_Y"), &hsaVal[2] );
	baProp[3] = Ctd.WindowGetProperty( hWndTbl, _T("MT_FOCUS_FRAME_COLOR"), &hsaVal[3] );
	if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] )
	{
		iaVal[0] = baProp[0] ? Ctd.HStrToInt( hsaVal[0] ) : 3;
		iaVal[1] = baProp[1] ? Ctd.HStrToInt( hsaVal[1] ) : -1;
		iaVal[2] = baProp[2] ? Ctd.HStrToInt( hsaVal[2] ) : -1;
		dwaVal[0] = baProp[3] ? Ctd.HStrToULong( hsaVal[3] ) : MTBL_COLOR_UNDEF;
		MTblDefineFocusFrame( hWndTbl, iaVal[0], iaVal[1], iaVal[2], dwaVal[0] );
	}

	// Scrolling
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_THUMBTRACK_HSCROLL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_THUMBTRACK_HSCROLL, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_THUMBTRACK_VSCROLL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_THUMBTRACK_VSCROLL, baVal[0] );
	}

	// Mousewheel Scrolling
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_MWSCR_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableMWheelScroll( hWndTbl, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_MWSCR_ROWS_PER_NOTCH"), &hsaVal[0] ) )
	{
		iaVal[0] = Ctd.HStrToInt( hsaVal[0] );
		MTblSetMWheelOption( hWndTbl, MTMW_ROWS_PER_NOTCH, iaVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_MWSCR_SPLITROWS_PER_NOTCH"), &hsaVal[0] ) )
	{
		iaVal[0] = Ctd.HStrToInt( hsaVal[0] );
		MTblSetMWheelOption( hWndTbl, MTMW_SPLITROWS_PER_NOTCH, iaVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_MWSCR_PAGE_KEY"), &hsaVal[0] ) )
	{
		iaVal[0] = Ctd.HStrToInt( hsaVal[0] );
		MTblSetMWheelOption( hWndTbl, MTMW_PAGE_KEY, iaVal[0] );
	}

	// Selections
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_SEL_TEXT_COLOR"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_SEL_BACK_COLOR"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		dwaVal[0] = baProp[0] ? Ctd.HStrToULong( hsaVal[0] ) : MTBL_COLOR_UNDEF;
		dwaVal[1] = baProp[1] ? Ctd.HStrToULong( hsaVal[1] ) : MTBL_COLOR_UNDEF;
		MTblSetSelectionColors( hWndTbl, dwaVal[0], dwaVal[1] );
	}

	if (Ctd.WindowGetProperty(hWndTbl, _T("MT_SEL_DARKENING"), &hsaVal[0]))
	{
		dwaVal[0] = Ctd.HStrToULong(hsaVal[0]);
		MTblSetSelectionDarkening(hWndTbl, (BYTE)dwaVal[0]);
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_FULL_OVERLAP_SELECTION"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_FULL_OVERLAP_SELECTION, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_SUPPRESS_ROW_SEL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_SUPPRESS_ROW_SELECTION, baVal[0] );
	}

	// Sort
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_SORT_KEEP_ROWMERGE"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_SORT_KEEP_ROWMERGE, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_SORT_RESTORE_TREE"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_SORT_RESTORE_TREE, baVal[0] );
	}

	// Tooltips
	// Enabled Tooltip Types
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_CELL_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_CELL, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_CELL_TEXT_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_CELL_TEXT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_CELL_COMPLETETEXT_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_CELL_COMPLETETEXT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_CELL_IMAGE_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_CELL_IMAGE, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_CELL_BTN_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_CELL_BTN, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_COLHDR_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_COLHDR, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_COLHDR_BTN_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_COLHDR_BTN, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_COLHDRGRP_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_COLHDRGRP, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_CORNER_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_CORNER, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_ROWHDR_ENABLED"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblEnableTipType( hWndTbl, MTBL_TIP_ROWHDR, baVal[0] );
	}

	// Tooltip Default Settings
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FONT_NAME"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FONT_SIZE"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FONT_ENH_BOLD"), &hsaVal[2] );
	baProp[3] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FONT_ENH_ITALIC"), &hsaVal[3] );
	baProp[4] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FONT_ENH_STRIKEOUT"), &hsaVal[4] );
	baProp[5] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FONT_ENH_UNDERLINE"), &hsaVal[5] );
	if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] || baProp[4] || baProp[5] )
	{
		
		saVal[0] = baProp[0] ? Ctd.HStrToLPTSTR( hsaVal[0] ) : _T("Arial");
		iaVal[1] = baProp[1] ? Ctd.HStrToInt( hsaVal[1] ) : 8;
		dwaVal[2] = 0;
		if ( baProp[2] && MTblPropStrToBool( hsaVal[2] ) )
			dwaVal[2] |= FONT_EnhBold;
		if ( baProp[3] && MTblPropStrToBool( hsaVal[3] ) )
			dwaVal[2] |= FONT_EnhItalic;
		if ( baProp[4] && MTblPropStrToBool( hsaVal[4] ) )
			dwaVal[2] |= FONT_EnhStrikeOut;
		if ( baProp[5] && MTblPropStrToBool( hsaVal[5] ) )
			dwaVal[2] |= FONT_EnhUnderline;

		MTblSetTipFont( hWndTbl, 0, MTBL_TIP_DEFAULT, (LPTSTR)saVal[0].c_str( ), iaVal[1], dwaVal[2] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_TEXT_COLOR"), &hsaVal[0] ) )
	{
		dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
		MTblSetTipTextColor( hWndTbl, 0, MTBL_TIP_DEFAULT, dwaVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FRAME_COLOR"), &hsaVal[0] ) )
	{
		dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
		MTblSetTipFrameColor( hWndTbl, 0, MTBL_TIP_DEFAULT, dwaVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_BACK_COLOR"), &hsaVal[0] ) )
	{
		dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
		MTblSetTipBackColor( hWndTbl, 0, MTBL_TIP_DEFAULT, dwaVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_DELAY"), &hsaVal[0] ) )
	{
		naVal[0] = Ctd.HStrToNum( hsaVal[0] );
		MTblSetTipDelay( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_FADE_IN_TIME"), &hsaVal[0] ) )
	{
		naVal[0] = Ctd.HStrToNum( hsaVal[0] );
		MTblSetTipFadeInTime( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0] );
	}

	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_POS_X"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_POS_Y"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		naVal[0] = baProp[0] ? Ctd.HStrToNum( hsaVal[0] ) : Ctd.NumGetNull( );
		naVal[1] = baProp[1] ? Ctd.HStrToNum( hsaVal[1] ) : Ctd.NumGetNull( );
		MTblSetTipPos( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0], naVal[1] );
	}

	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_OFFSET_X"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_OFFSET_Y"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		naVal[0] = baProp[0] ? Ctd.HStrToNum( hsaVal[0] ) : Ctd.NumGetNull( );
		naVal[1] = baProp[1] ? Ctd.HStrToNum( hsaVal[1] ) : Ctd.NumGetNull( );
		MTblSetTipOffset( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0], naVal[1] );
	}

	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_WIDTH"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_HEIGHT"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		naVal[0] = baProp[0] ? Ctd.HStrToNum( hsaVal[0] ) : Ctd.NumGetNull( );
		naVal[1] = baProp[1] ? Ctd.HStrToNum( hsaVal[1] ) : Ctd.NumGetNull( );
		MTblSetTipSize( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0], naVal[1] );
	}

	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_RCE_WIDTH"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_RCE_HEIGHT"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		naVal[0] = baProp[0] ? Ctd.HStrToNum( hsaVal[0] ) : Ctd.NumGetNull( );
		naVal[1] = baProp[1] ? Ctd.HStrToNum( hsaVal[1] ) : Ctd.NumGetNull( );
		MTblSetTipRCESize( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0], naVal[1] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_GRADIENT"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTipFlags( hWndTbl, 0, MTBL_TIP_DEFAULT, MTBL_TIP_FLAG_GRADIENT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_OPACITY"), &hsaVal[0] ) )
	{
		naVal[0] = Ctd.HStrToNum( hsaVal[0] );
		MTblSetTipOpacity( hWndTbl, 0, MTBL_TIP_DEFAULT, naVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_NO_FRAME"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTipFlags( hWndTbl, 0, MTBL_TIP_DEFAULT, MTBL_TIP_FLAG_NOFRAME, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TT_DEF_PERMEABLE"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTipFlags( hWndTbl, 0, MTBL_TIP_DEFAULT, MTBL_TIP_FLAG_PERMEABLE, baVal[0] );
	}

	// Tree-Definition
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_COL"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_SIZE"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_INDENT"), &hsaVal[2] );
	if ( baProp[0] || baProp[1] || baProp[2] )
	{
		hWndaVal[0] = NULL;
		if ( baProp[0] )
		{
			HWND hWndCol = NULL;
			while ( MTblFindNextCol( hWndTbl, &hWndCol, 0, 0 ) )
			{
				if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_IS_TREE_COL"), &hsaVal[0] ) )
				{
					if ( MTblPropStrToBool( hsaVal[0] ) )
					{
						hWndaVal[0] = hWndCol;
						break;
					}
				}
			}
		}

		iaVal[0] = baProp[1] ? Ctd.HStrToInt( hsaVal[1] ) : 0;
		iaVal[1] = baProp[2] ? Ctd.HStrToInt( hsaVal[2] ) : 0;

		MTblDefineTree( hWndTbl, hWndaVal[0], iaVal[0], iaVal[1] );
	}

	// Tree Node Colors
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_OUTER_COLOR"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_INNER_COLOR"), &hsaVal[1] );
	baProp[2] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_BACK_COLOR"), &hsaVal[2] );
	if ( baProp[0] || baProp[1] || baProp[2] )
	{
		dwaVal[0] = baProp[0] ? Ctd.HStrToULong( hsaVal[0] ) : MTBL_COLOR_UNDEF;
		dwaVal[1] = baProp[1] ? Ctd.HStrToULong( hsaVal[1] ) : MTBL_COLOR_UNDEF;
		dwaVal[2] = baProp[2] ? Ctd.HStrToULong( hsaVal[2] ) : MTBL_COLOR_UNDEF;
		MTblSetTreeNodeColors( hWndTbl, dwaVal[0], dwaVal[1], dwaVal[2] );
	}

	// Tree Node Images
	hImg1 = hImg2 = NULL;

	hImg = NULL;
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_COLL_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_COLL_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_COLL_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}
			}
		}
	}
	hImg1 = hImg;

	hImg = NULL;
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_EXP_IMG_SOURCE"), &hsaVal[0] ) )
	{
		saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
		if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_EXP_IMG_NAME"), &hsaVal[1] ) )
		{
			saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
			if ( saVal[0] == _T("FILE") )
			{
				hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
			}
			else if ( saVal[0] == _T("TDRES") )
			{
				hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
			}

			if ( hImg )
			{
				if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NODE_EXP_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
				{
					dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
					MImgSetTranspColor( hImg, dwaVal[2] );
				}
			}
		}
	}
	hImg2 = hImg;

	if ( hImg1 || hImg2 )
		MTblSetNodeImages( hWndTbl, hImg1, hImg2, 0 );


	// Tree Lines Appearance
	baProp[0] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_LINES_STYLE"), &hsaVal[0] );
	baProp[1] = Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_LINES_COLOR"), &hsaVal[1] );
	if ( baProp[0] || baProp[1] )
	{
		
		iaVal[0] = baProp[0] ?Ctd.HStrToInt( hsaVal[0] ) : MTLS_INVISIBLE;
		dwaVal[1] = baProp[1] ? Ctd.HStrToULong( hsaVal[1] ) : MTBL_COLOR_UNDEF;
		MTblDefineTreeLines( hWndTbl, iaVal[0], dwaVal[1] );
	}

	// Tree-Flags
	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_AUTO_NORM_HIER"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_AUTO_NORM_HIER, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_INDENT_ALL"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_INDENT_ALL, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_INDENT_NONE"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_INDENT_NONE, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_BOTTOM_UP"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_BOTTOM_UP, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_SPLIT_BOTTOM_UP"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_SPLIT_BOTTOM_UP, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_FLAT_STRUCT"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_FLAT_STRUCT, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NO_ROWLINES"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_NO_ROWLINES, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_EXP_ROW_ON_DBLCLK"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetFlags( hWndTbl, MTBL_FLAG_EXPAND_ROW_ON_DBLCLK, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NOSELINV_NODES"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_NOSELINV_NODES, baVal[0] );
	}

	if ( Ctd.WindowGetProperty( hWndTbl, _T("MT_TREE_NOSELINV_LINES"), &hsaVal[0] ) )
	{
		baVal[0] = MTblPropStrToBool( hsaVal[0] );
		MTblSetTreeFlags( hWndTbl, MTBL_TREE_FLAG_NOSELINV_LINES, baVal[0] );
	}

	// ***Columns***
	HWND hWndCol = NULL;
	while ( MTblFindNextCol( hWndTbl, &hWndCol, 0, 0 ) )
	{
		//
		// Kategorie "(General)"
		//
		// Ownerdrawn
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_OWNERDRAWN"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetOwnerDrawnItem( hWndCol, TBL_Error, MTBL_ODI_COLUMN, baVal[0] );
		}

		// End Ellipsis
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_END_ELLIPSIS"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_END_ELLIPSIS, TRUE );
		}

		// Export As Text
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_EXPORT_AS_TEXT"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_EXPORT_AS_TEXT, TRUE );
		}

		//
		// Kategorie "(Header)"
		//
		// Ownerdrawn
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_OWNERDRAWN"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetOwnerDrawnItem( hWndCol, TBL_Error, MTBL_ODI_COLHDR, baVal[0] );
		}

		// Text Alignment Horizontal
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_TXTALIGN_H"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnHdrFlags( hWndCol, dwaVal[0], TRUE );
		}

		// Gruppe "Font"
		baProp[0] = Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_FONT_NAME"), &hsaVal[0] );
		baProp[1] = Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_FONT_SIZE"), &hsaVal[1] );
		baProp[2] = Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_FONT_ENH_BOLD"), &hsaVal[2] );
		baProp[3] = Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_FONT_ENH_ITALIC"), &hsaVal[3] );
		baProp[4] = Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_FONT_ENH_STRIKEOUT"), &hsaVal[4] );
		baProp[5] = Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_FONT_ENH_UNDERLINE"), &hsaVal[5] );
		if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] || baProp[4] || baProp[5] )
		{
			
			saVal[0] = baProp[0] ? Ctd.HStrToLPTSTR( hsaVal[0] ) : MTBL_FONT_UNDEF_NAME;
			iaVal[1] = baProp[1] ? Ctd.HStrToInt( hsaVal[1] ) : MTBL_FONT_UNDEF_SIZE;

			dwaVal[2] = 0;
			if ( baProp[2] && MTblPropStrToBool( hsaVal[2] ) )
				dwaVal[2] |= FONT_EnhBold;
			if ( baProp[3] && MTblPropStrToBool( hsaVal[3] ) )
				dwaVal[2] |= FONT_EnhItalic;
			if ( baProp[4] && MTblPropStrToBool( hsaVal[4] ) )
				dwaVal[2] |= FONT_EnhStrikeOut;
			if ( baProp[5] && MTblPropStrToBool( hsaVal[5] ) )
				dwaVal[2] |= FONT_EnhUnderline;

			if ( !dwaVal[2] )
				dwaVal[2] = MTBL_FONT_UNDEF_ENH;

			MTblSetColumnHdrFont( hWndCol, (LPTSTR)saVal[0].c_str( ), iaVal[1], dwaVal[2], MTSF_REDRAW );
		}

		// Gruppe "Colors"
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_TEXT_COLOR"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnHdrTextColor( hWndCol, dwaVal[0], MTSC_REDRAW );
		}

		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_BACK_COLOR"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnHdrBackColor( hWndCol, dwaVal[0], MTSC_REDRAW );
		}

		// Gruppe "Image"
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_SOURCE"), &hsaVal[0] ) )
		{
			saVal[0] = Ctd.HStrToLPTSTR( hsaVal[0] );
			if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_NAME"), &hsaVal[1] ) )
			{
				saVal[1] = Ctd.HStrToLPTSTR( hsaVal[1] );
				hImg = NULL;
				if ( saVal[0] == _T("FILE") )
				{
					hImg = MImgLoadFromFile( (LPTSTR)saVal[1].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
				}
				else if ( saVal[0] == _T("TDRES") )
				{
					hImg = MImgLoadFromResource( hsaVal[1], MIMG_TYPE_UNKNOWN, 0 );
				}

				if ( hImg )
				{
					if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_TRANSP_COLOR"), &hsaVal[2] ) )
					{
						dwaVal[2] = Ctd.HStrToULong( hsaVal[2] );
						MImgSetTranspColor( hImg, dwaVal[2] );
					}

					dwaVal[3] = 0;
					if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_ALIGN_TILE"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
						dwaVal[3] |= MTSI_ALIGN_TILE;
					else
					{
						if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_ALIGN_H"), &hsaVal[3] ) )
							dwaVal[3] |= Ctd.HStrToULong( hsaVal[3] );
						if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_ALIGN_V"), &hsaVal[3] ) )
							dwaVal[3] |= Ctd.HStrToULong( hsaVal[3] );
					}
					if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_NOLEAD_LEFT"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
						dwaVal[3] |= MTSI_NOLEAD_LEFT;
					if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_NOLEAD_TOP"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
						dwaVal[3] |= MTSI_NOLEAD_TOP;
					if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_HDR_IMG_NOSELINV"), &hsaVal[3] ) && MTblPropStrToBool( hsaVal[3] ) )
						dwaVal[3] |= MTSI_NOSELINV;
					dwaVal[3] |= MTSI_REDRAW;

					MTblSetColumnHdrImage( hWndCol, hImg, dwaVal[3] );
				}
			}
		}

		//
		// Kategorie "Button"
		//
		baProp[0] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_ACTIVE"), &hsaVal[0] );
		baProp[1] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_WIDTH"), &hsaVal[1] );
		baProp[2] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_TEXT"), &hsaVal[2] );
		baProp[3] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_IMG_SOURCE"), &hsaVal[3] );
		baProp[4] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_DISABLED"), &hsaVal[4] );
		baProp[5] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_SHARE_FONT"), &hsaVal[5] );
		baProp[6] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_LEFT"), &hsaVal[6] );
		baProp[7] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_FLAT"), &hsaVal[7] );
		baProp[8] = Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_NO_BKGND"), &hsaVal[8] );
		if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] || baProp[4] || baProp[5] || baProp[6] || baProp[7] || baProp[8] )
		{
			baVal[0] = baProp[0] ? MTblPropStrToBool( hsaVal[0] ) : FALSE;
			iaVal[1] = baProp[1] ? Ctd.HStrToInt( hsaVal[1] ) : 0;
			saVal[2] = baProp[2] ? Ctd.HStrToLPTSTR( hsaVal[2] ) : _T("");

			hImg = NULL;
			if ( baProp[3] )
			{
				saVal[3] = Ctd.HStrToLPTSTR( hsaVal[3] );
				if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_IMG_NAME"), &hsaVal[TMP] ) )
				{
					saVal[TMP] = Ctd.HStrToLPTSTR( hsaVal[TMP] );
					
					if ( saVal[3] == _T("FILE") )
					{
						hImg = MImgLoadFromFile( (LPTSTR)saVal[TMP].c_str( ), MIMG_TYPE_UNKNOWN, 0 );
					}
					else if ( saVal[3] == _T("TDRES") )
					{
						hImg = MImgLoadFromResource( hsaVal[TMP], MIMG_TYPE_UNKNOWN, 0 );
					}
				}
			}
			if ( hImg && Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_IMG_TRANSP_COLOR"), &hsaVal[TMP] ) )
			{
				dwaVal[TMP] = Ctd.HStrToULong( hsaVal[TMP] );
				MImgSetTranspColor( hImg, dwaVal[TMP] );
			}

			dwaVal[4] = 0;
			if ( baProp[4] && MTblPropStrToBool( hsaVal[4] ) )
				dwaVal[4] |= MTBL_BTN_DISABLED;
			if ( baProp[5] && MTblPropStrToBool( hsaVal[5] ) )
				dwaVal[4] |= MTBL_BTN_SHARE_FONT;
			if ( baProp[6] && MTblPropStrToBool( hsaVal[6] ) )
				dwaVal[4] |= MTBL_BTN_LEFT;
			if ( baProp[7] && MTblPropStrToBool( hsaVal[7] ) )
				dwaVal[4] |= MTBL_BTN_FLAT;
			if ( baProp[8] && MTblPropStrToBool( hsaVal[8] ) )
				dwaVal[4] |= MTBL_BTN_NO_BKGND;

			MTblDefineButtonColumn( hWndCol, baVal[0], iaVal[1], (LPTSTR)saVal[2].c_str( ), hImg, dwaVal[4] );
		}

		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_BUTTONS_PERMANENT"), &hsaVal[0] ) && MTblPropStrToBool( hsaVal[0] ) )
		{
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_BUTTONS_PERMANENT, TRUE );
		}

		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_SHOW_PERM_BTN_NE"), &hsaVal[0] ) && MTblPropStrToBool( hsaVal[0] ) )
		{
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_SHOW_PERM_BTN_NE, TRUE );
		}

		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_BTN_HIDE_PERM"), &hsaVal[0] ) && MTblPropStrToBool( hsaVal[0] ) )
		{
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_HIDE_PERM_BTN, TRUE );
		}

		//
		// Kategorie "Colors"
		//
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_TEXT_COLOR"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnTextColor( hWndCol, dwaVal[0], MTSC_REDRAW );
		}

		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_BACK_COLOR"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnBackColor( hWndCol, dwaVal[0], MTSC_REDRAW );
		}

		//
		// Kategorie "Edit Mode"
		//
		// Read Only
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_READ_ONLY"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_READ_ONLY, baVal[0] );
		}

		// Clip Non Editable Area
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_CLIP_NONEDIT_AREA"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_CLIP_NONEDIT, baVal[0] );
		}

		// Autosize Edit
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_CLIP_AUTOSIZE_EDIT"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_AUTOSIZE_EDIT, baVal[0] );
		}

		// Horizontal Text Alignment
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_EDITALIGN"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, dwaVal[0], TRUE );
		}

		//
		// Kategorie "Font"
		//
		baProp[0] = Ctd.WindowGetProperty( hWndCol, _T("MTC_FONT_NAME"), &hsaVal[0] );
		baProp[1] = Ctd.WindowGetProperty( hWndCol, _T("MTC_FONT_SIZE"), &hsaVal[1] );
		baProp[2] = Ctd.WindowGetProperty( hWndCol, _T("MTC_FONT_ENH_BOLD"), &hsaVal[2] );
		baProp[3] = Ctd.WindowGetProperty( hWndCol, _T("MTC_FONT_ENH_ITALIC"), &hsaVal[3] );
		baProp[4] = Ctd.WindowGetProperty( hWndCol, _T("MTC_FONT_ENH_STRIKEOUT"), &hsaVal[4] );
		baProp[5] = Ctd.WindowGetProperty( hWndCol, _T("MTC_FONT_ENH_UNDERLINE"), &hsaVal[5] );
		if ( baProp[0] || baProp[1] || baProp[2] || baProp[3] || baProp[4] || baProp[5] )
		{
			
			saVal[0] = baProp[0] ? Ctd.HStrToLPTSTR( hsaVal[0] ) : MTBL_FONT_UNDEF_NAME;
			iaVal[1] = baProp[1] ? Ctd.HStrToInt( hsaVal[1] ) : MTBL_FONT_UNDEF_SIZE;

			dwaVal[2] = 0;
			if ( baProp[2] && MTblPropStrToBool( hsaVal[2] ) )
				dwaVal[2] |= FONT_EnhBold;
			if ( baProp[3] && MTblPropStrToBool( hsaVal[3] ) )
				dwaVal[2] |= FONT_EnhItalic;
			if ( baProp[4] && MTblPropStrToBool( hsaVal[4] ) )
				dwaVal[2] |= FONT_EnhStrikeOut;
			if ( baProp[5] && MTblPropStrToBool( hsaVal[5] ) )
				dwaVal[2] |= FONT_EnhUnderline;

			if ( !dwaVal[2] )
				dwaVal[2] = MTBL_FONT_UNDEF_ENH;

			MTblSetColumnFont( hWndCol, (LPTSTR)saVal[0].c_str( ), iaVal[1], dwaVal[2], MTSF_REDRAW );
		}

		//
		// Kategorie "Lines"
		//
		// No Row Lines
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_NO_ROWLINES"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_NO_ROWLINES, baVal[0] );
		}

		// No Column Line
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_NO_COLLINE"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_NO_COLLINE, baVal[0] );
		}

		//
		// Kategorie "Selection"
		//		
		// Don't Invert Text
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_NOSELINV_TEXT"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_NOSELINV_TEXT, baVal[0] );
		}

		// Don't Invert Image
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_NOSELINV_IMAGE"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_NOSELINV_IMAGE, baVal[0] );
		}

		// Don't Invert Background
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_NOSELINV_BKGND"), &hsaVal[0] ) )
		{
			baVal[0] = MTblPropStrToBool( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, MTBL_COL_FLAG_NOSELINV_BKGND, baVal[0] );
		}

		//
		// Kategorie "Text Alignment"
		//		
		// Horizontal
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_TXTALIGN_H"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, dwaVal[0], TRUE );
		}

		// Vertical
		if ( Ctd.WindowGetProperty( hWndCol, _T("MTC_TXTALIGN_V"), &hsaVal[0] ) )
		{
			dwaVal[0] = Ctd.HStrToULong( hsaVal[0] );
			MTblSetColumnFlags( hWndCol, dwaVal[0], TRUE );
		}
	}



	// H-Strings freigeben
	for ( int i = 0; i < 25; i++ )
	{
	if ( hsaVal[i] )
		Ctd.HStringUnRef( hsaVal[i] );
	}

	return TRUE;
}


// MTblAutoSizeColumn
// Paßt die Breite von Spalten an
BOOL WINAPI MTblAutoSizeColumn( HWND hWndTbl, HWND hWndCol, DWORD dwFlags )
{	
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Spalteninfo-Struktur
	typedef struct
	{
		HWND hWnda[MTBL_MAX_COLS];
		int iaType[MTBL_MAX_COLS];
		int iaWidth[MTBL_MAX_COLS];
		int iaWidthOrg[MTBL_MAX_COLS];
	}
	COLINFO;

	COLINFO ci;
	DWORD dwColFlagsOn = 0, dwColFlagsOff = 0;
	HWND hWndColFound = NULL;
	int iCols = 0, iCol = 0, iMaxWidth, iMinWidth, iWidth;
	long lPos = 0, lFirstRow = 0, lLastRow = 0, lRow = 0, lLen = 0, lColWidth = 0;
	WORD wRowFlagsOn = 0, wRowFlagsOff = 0;

	// Tabellen-Handle Ok?
	if ( ! IsWindow( hWndTbl ) )
		return FALSE;

	// Spalteninfos setzen
	if ( !hWndCol )
	{
		if ( !( dwFlags & MTASC_HIDDENCOLS ) )
			dwColFlagsOn = COL_Visible;

		//while ( MTblFindNextCol( hWndTbl, &hWndColFound, dwColFlagsOn, dwColFlagsOff ) )
		while ( psc->FindNextCol( &hWndColFound, dwColFlagsOn, dwColFlagsOff ) )
		{
			ci.hWnda[iCols] = hWndColFound;
			//ci.iaType[iCols] = Ctd.TblGetCellType( hWndColFound );
			ci.iaType[iCols] = psc->GetEffCellType( TBL_Error, hWndColFound );
			ci.iaWidth[iCols] = 0;
			ci.iaWidthOrg[iCols] = SendMessage( hWndColFound, TIM_QueryColumnWidth, 0, 0 );
			iCols++;
			if ( iCols == MTBL_MAX_COLS )
				break;
		}
	}
	else
	{
		if ( ! IsWindow( hWndCol ) )
			return FALSE;
		ci.hWnda[0] = hWndCol;
		//ci.iaType[0] = Ctd.TblGetCellType( hWndCol );
		ci.iaType[0] = psc->GetEffCellType( TBL_Error, hWndCol );
		ci.iaWidth[0] = 0;
		ci.iaWidthOrg[0] = SendMessage( hWndCol, TIM_QueryColumnWidth, 0, 0 );
		iCols = 1;
	}

	// Tabellenmaße ermitteln
	CMTblMetrics tm( hWndTbl );

	// Suchflags setzen
	if ( !( dwFlags & MTASC_HIDDENROWS ) )
		wRowFlagsOff = ROW_Hidden;

	// Buttons außerhalb von Zellen?
	BOOL bButtonsOutsideCell = psc->QueryFlags( MTBL_FLAG_BUTTONS_OUTSIDE_CELL );

	// Los geht's
	BOOL bBtnEditSpace, bBtnPermSpace, bMargin;

	// *** Normale Zeilen ***
	// Bereich ermitteln
	if ( dwFlags & MTASC_ALLROWS )
	{
		Ctd.TblQueryScroll( hWndTbl, &lPos, &lFirstRow, &lLastRow );
	}
	else
	{
		Ctd.TblQueryVisibleRange( hWndTbl, &lFirstRow, &lLastRow );
		lLastRow++;
	}

	// Optimale Breite ermitteln 
	lRow = lFirstRow - 1;
	while( psc->FindNextRow( &lRow, wRowFlagsOn, wRowFlagsOff ) && lRow <= lLastRow )
	{
		for ( iCol = 0; iCol < iCols; iCol++ )
		{
			iWidth = psc->GetBestCellWidth( ci.hWnda[iCol], lRow, tm.m_CellLeadingX );

			// Button-Platz einberechnen?
			if ( dwFlags & MTASC_BTNSPACE )
			{
				bMargin = FALSE;
				bBtnEditSpace = !bButtonsOutsideCell && psc->MustShowBtn( ci.hWnda[iCol], lRow, TRUE );
				bBtnPermSpace = psc->MustShowBtn( ci.hWnda[iCol], lRow, TRUE, TRUE );
				if ( bBtnEditSpace || bBtnPermSpace )
				{
					CMTblBtn *pBtn = psc->GetEffCellBtn( ci.hWnda[iCol], lRow );
					iWidth += psc->CalcBtnWidth( pBtn, ci.hWnda[iCol], lRow );
					
					if ( pBtn && !pBtn->IsLeft( ) )
						bMargin = TRUE;
				}

				if ( !bButtonsOutsideCell && psc->GetEffCellType( lRow, ci.hWnda[iCol] ) == COL_CellType_DropDownList )
				{
					iWidth += MTBL_DDBTN_WID;
					bMargin = TRUE;
				}

				if ( bMargin )
					++iWidth;
			}

			ci.iaWidth[iCol] = max(iWidth, ci.iaWidth[iCol]);
		}
	}

	// *** Split-Zeilen ***
	if ( dwFlags & MTASC_SPLITROWS && psc->IsSplitted( ) )
	{
		// Bereich ermitteln
		if ( dwFlags & MTASC_ALLROWS )
		{
			lFirstRow = TBL_MinSplitRow;
			lLastRow = TBL_MaxSplitRow;
		}
		else
		{
			//if ( MTblQueryVisibleSplitRange( hWndTbl, &lFirstRow, &lLastRow ) )
			if ( psc->QueryVisRange( lFirstRow, lLastRow, TRUE ) )
				lLastRow++;
			else
			{
				lFirstRow = TBL_MinSplitRow;
				lLastRow = TBL_MaxSplitRow;
			}
		}

		// Optimale Breite ermitteln
		lRow = lFirstRow - 1;
		if ( lRow < TBL_MinSplitRow )
			lRow = TBL_MinRow;
		while( psc->FindNextRow( &lRow, wRowFlagsOn, wRowFlagsOff ) && lRow <= lLastRow )
		{
			for ( iCol = 0; iCol < iCols; iCol++ )
			{
				iWidth = psc->GetBestCellWidth( ci.hWnda[iCol], lRow, tm.m_CellLeadingX );

				// Button-Platz einberechnen?
				if ( dwFlags & MTASC_BTNSPACE )
				{
					bMargin = FALSE;
					bBtnEditSpace = !bButtonsOutsideCell && psc->MustShowBtn( ci.hWnda[iCol], lRow, TRUE );
					bBtnPermSpace = psc->MustShowBtn( ci.hWnda[iCol], lRow, TRUE, TRUE );
					if ( bBtnEditSpace || bBtnPermSpace )
					{
						CMTblBtn *pBtn = psc->GetEffCellBtn( ci.hWnda[iCol], lRow );
						iWidth += psc->CalcBtnWidth( pBtn, ci.hWnda[iCol], lRow );

						if ( pBtn && !pBtn->IsLeft( ) )
							bMargin = TRUE;
					}

					if ( !bButtonsOutsideCell && psc->GetEffCellType( lRow, ci.hWnda[iCol] ) == COL_CellType_DropDownList )
					{
						iWidth += MTBL_DDBTN_WID;
						bMargin = TRUE;
					}

					if ( bMargin )
						++iWidth;
				}

				ci.iaWidth[iCol] = max(iWidth, ci.iaWidth[iCol]);
			}
		}
	}

	// *** Spalten-Header + min./max. Breite ermitteln + Breite setzen ***
	BOOL bSize;
	for ( iCol = 0; iCol < iCols; iCol++ )
	{
		if (!psc->QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER) && !Ctd.TblQueryTableFlags(hWndTbl, TBL_Flag_HideColumnHeaders))
		{
			iWidth = psc->GetBestColHdrWidth( ci.hWnda[iCol], tm.m_CellLeadingX );
			if ( dwFlags & MTASC_BTNSPACE )
			{
				CMTblBtn *pBtn = psc->m_Cols->GetHdrBtn( ci.hWnda[iCol] );
				if ( pBtn && pBtn->m_bShow )
					iWidth += psc->CalcBtnWidth( pBtn, ci.hWnda[iCol], TBL_Error );
			}
			

			ci.iaWidth[iCol] = max(iWidth, ci.iaWidth[iCol]);
		}

		iMinWidth = SendMessage( hWndTbl, MTM_QueryMinColWidth, WPARAM(ci.hWnda[iCol]), ci.iaWidth[iCol] );
		iMaxWidth = SendMessage( hWndTbl, MTM_QueryMaxColWidth, WPARAM(ci.hWnda[iCol]), ci.iaWidth[iCol] );

		if ( iMinWidth > 0 )
			ci.iaWidth[iCol] = max(iMinWidth, ci.iaWidth[iCol]);			
		if ( iMaxWidth > 0 )
			ci.iaWidth[iCol] = min(iMaxWidth, ci.iaWidth[iCol]);
		
		if ( dwFlags & MTASC_INCREASE_ONLY )
			bSize = ci.iaWidth[iCol] > ci.iaWidthOrg[iCol];
		else
			bSize = TRUE;

		if ( bSize )
			Ctd.TblSetColumnWidth( ci.hWnda[iCol], Ctd.PixelsToFormUnits( ci.hWnda[iCol], ci.iaWidth[iCol], FALSE ) );
	}

	return TRUE;
}


// MTblAutoSizeRows
// Paßt die Höhe von Zeilen an
BOOL WINAPI MTblAutoSizeRows( HWND hWndTbl, DWORD dwFlags )
{	
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Spalteninfo-Struktur
	typedef struct
	{
		HWND hWnda[MTBL_MAX_COLS];
		int iaType[MTBL_MAX_COLS];
	}
	COLINFO;

	BOOL bSizeRows = psc->QueryFlags( MTBL_FLAG_VARIABLE_ROW_HEIGHT );
	BOOL bIncreaseOnly = (dwFlags & MTASR_INCREASE_ONLY) > 0;
	BOOL bContextRowOnly = (dwFlags & MTASR_CONTEXTROW_ONLY) > 0;
	BOOL bRowsSized = FALSE;
	COLINFO ci;
	DWORD dwColFlagsOn = 0, dwColFlagsOff = 0;
	HWND hWndColFound = NULL;
	int iCols = 0, iCol = 0, iBestBtnHt, iBestCellHt, iBestRowHt, iMaxHeight = 0, iMaxRowHt, iMinRowHt;
	long lPos = 0, lFirstRow = 0, lLastRow = 0, lRow = 0, lLen = 0, lColWidth = 0;
	WORD wRowFlagsOn = 0, wRowFlagsOff = 0;

	// Tabellen-Handle Ok?
	if ( ! IsWindow( hWndTbl ) )
		return FALSE;

	// Spalteninfos setzen
	if ( !( dwFlags & MTASR_HIDDENCOLS ) )
		dwColFlagsOn = COL_Visible;

	while ( psc->FindNextCol( &hWndColFound, dwColFlagsOn, dwColFlagsOff ) )
	{
		ci.hWnda[iCols] = hWndColFound;
		ci.iaType[iCols] = psc->GetEffCellType( TBL_Error, hWndColFound );
		iCols++;
		if ( iCols == MTBL_MAX_COLS )
			break;
	}

	// Tabellenmaße ermitteln
	CMTblMetrics tm( hWndTbl );

	// Suchflags Zeilen setzen
	if ( !( dwFlags & MTASR_HIDDENROWS ) && !bContextRowOnly )
		wRowFlagsOff = ROW_Hidden;

	// Los geht's
	int irr = 0;
	MTBLROWRANGE rra[2];

	// Nur Zeile mit Kontext?
	if ( bContextRowOnly )
	{
		lFirstRow = lLastRow = Ctd.TblQueryContext( hWndTbl );

		LPMTBLMERGEROW pmr = psc->FindMergeRow( lFirstRow, FMR_ANY );
		if ( pmr )
		{
			lFirstRow = pmr->lRowFrom;
			lLastRow = pmr->lRowTo;
		}

		rra[irr].lRowFrom = lFirstRow;
		rra[irr].lRowTo = lLastRow;
		++irr;
	}
	else
	{	
		// *** Normale Zeilen ***
		// Bereich ermitteln
		if ( dwFlags & MTASR_ALLROWS )
		{
			Ctd.TblQueryScroll( hWndTbl, &lPos, &lFirstRow, &lLastRow );
		}
		else
		{
			Ctd.TblQueryVisibleRange( hWndTbl, &lFirstRow, &lLastRow );
			lLastRow++;
		}

		rra[irr].lRowFrom = lFirstRow;
		rra[irr].lRowTo = lLastRow;
		++irr;		

		// *** Split-Zeilen ***
		if ( dwFlags & MTASR_SPLITROWS && psc->IsSplitted( ) )
		{
			// Bereich ermitteln
			if ( dwFlags & MTASR_ALLROWS )
			{
				lFirstRow = TBL_MinSplitRow;
				lLastRow = TBL_MaxSplitRow;
			}
			else
			{
				//if ( MTblQueryVisibleSplitRange( hWndTbl, &lFirstRow, &lLastRow ) )
				if ( psc->QueryVisRange( lFirstRow, lLastRow, TRUE ) )
					lLastRow++;
				else
				{
					lFirstRow = TBL_MinSplitRow;
					lLastRow = TBL_MaxSplitRow;
				}
			}

			rra[irr].lRowFrom = lFirstRow;
			rra[irr].lRowTo = lLastRow;
			++irr;			
		}
	}

	// Optimale Höhe ermitteln bzw. optimale Höhen setzen
	for ( int i = 0; i < irr; ++i )
	{
		lFirstRow = rra[i].lRowFrom;
		lLastRow = rra[i].lRowTo;

		lRow = lFirstRow - 1;
		if ( lRow < TBL_MinSplitRow )
			lRow = TBL_MinRow;
		while( psc->FindNextRow( &lRow, wRowFlagsOn, wRowFlagsOff ) && lRow <= lLastRow )
		{
			iBestRowHt = 0;
			for ( iCol = 0; iCol < iCols; iCol++ )
			{
				iBestCellHt = psc->GetBestCellHeight( ci.hWnda[iCol], lRow, tm.m_CellLeadingY, tm.m_CharBoxY, NULL, NULL, FALSE, FALSE );

				// Button-Platz einberechnen?
				if ( ( dwFlags & MTASR_BTNSPACE ) && ( psc->MustShowBtn( ci.hWnda[iCol], lRow, TRUE ) || psc->MustShowBtn( ci.hWnda[iCol], lRow, TRUE, TRUE ) ) )
				{
					iBestBtnHt = psc->GetBtnBestHeight( psc->GetEffCellBtn( ci.hWnda[iCol], lRow ), lRow, ci.hWnda[iCol] );
					iBestCellHt = max( iBestCellHt, iBestBtnHt + tm.m_CellLeadingY );
				}

				iBestRowHt = max( iBestRowHt, iBestCellHt );
			}

			iMinRowHt = SendMessage( hWndTbl, MTM_QueryMinRowHeight, lRow, iBestRowHt );
			iMaxRowHt = SendMessage( hWndTbl, MTM_QueryMaxRowHeight, lRow, iBestRowHt );

			if ( iMinRowHt > 0 )
				iBestRowHt = max( iMinRowHt, iBestRowHt );
			if ( iMaxRowHt > 0 )
				iBestRowHt = min( iMaxRowHt, iBestRowHt );

			iMaxHeight = max( iMaxHeight, iBestRowHt );

			if ( bSizeRows && iBestRowHt != psc->GetEffRowHeight( lRow ) )
			{
				psc->m_Rows->SetRowHeight( lRow, (iBestRowHt == tm.m_RowHeight ? 0 : iBestRowHt) );
				bRowsSized = TRUE;
			}
		}
	}

	if ( bSizeRows )
	{
		if ( bRowsSized )
		{
			//Ctd.TblKillEdit( hWndTbl );

			psc->AdaptScrollRangeV( FALSE );
			if ( dwFlags & MTASR_SPLITROWS )
				psc->AdaptScrollRangeV( TRUE );

			psc->AdaptListClip( );
			psc->UpdateFocus( );

			InvalidateRect( hWndTbl, NULL, FALSE );
		}
	}
	else
	{
		// Anzahl "LinesPerRow" anhand max. Höhe setzen
		int iLinesPerRow = 1;
		while ( tm.CalcRowHeight( iLinesPerRow ) < iMaxHeight )
		{
			++iLinesPerRow;
		}

		BOOL bSet = TRUE;
		if ( bIncreaseOnly )
		{
			int iLinesPerRowCurrent = 0;
			if ( Ctd.TblQueryLinesPerRow( hWndTbl, &iLinesPerRowCurrent ) )
				bSet = ( iLinesPerRow > iLinesPerRowCurrent );
		}
		
		if ( bSet )
		{
			if ( !Ctd.TblSetLinesPerRow( hWndTbl, iLinesPerRow ) )
				return FALSE;
		}
	}

	return TRUE;
}

// MTblCalcRowHeight
// Berechnet die Höhe einer Zeile für eine best. Anzahl Lines per Row
int WINAPI MTblCalcRowHeight( HWND hWndTbl, int iLinesPerRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	CMTblMetrics tm( hWndTbl );
	return tm.CalcRowHeight( iLinesPerRow );
}

// MTblClearCellSelection
BOOL WINAPI MTblClearCellSelection( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	BOOL bOk = psc->ClearCellSelection( );
	UpdateWindow( hWndTbl );

	return bOk;
}


// MTblClearExportSubstColors
// Löscht alle Export-Ersatzfarben
BOOL WINAPI MTblClearExportSubstColors( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Löschen
	psc->m_escm->clear( );

	return TRUE;
}


// MTblColFromPoint
// Liefert das Handle der Spalte an einem Punkt
HWND WINAPI MTblColFromPoint( HWND hWndTbl, int iX, int iY )
{
	if ( !Ctd.IsPresent( ) ) return NULL;

	HWND hWndCol = NULL;
	long lRow;
	DWORD dwFlags;
	Ctd.TblObjectsFromPoint( hWndTbl, iX, iY, &lRow, &hWndCol, &dwFlags );

	return hWndCol;
}


// MTblCopyRows
// Kopiert Zeilen ohne 64KB-Beschränkung in die Zwischenablage
BOOL WINAPI MTblCopyRows( HWND hWndTbl, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->CopyRows( wRowFlagsOn, wRowFlagsOff, dwColFlagsOn, dwColFlagsOff, FALSE );

}


// MTblCopyRowsAndHeaders
// Kopiert Zeilen + Column Header ohne 64KB-Beschränkung in die Zwischenablage
BOOL WINAPI MTblCopyRowsAndHeaders( HWND hWndTbl, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->CopyRows( wRowFlagsOn, wRowFlagsOff, dwColFlagsOn, dwColFlagsOff, TRUE );

}


// MTblCollapseRow
// Schließt eine Zeile
long WINAPI MTblCollapseRow( HWND hWndTbl, long lRow, WPARAM wFlags )
{
	return SendMessage( hWndTbl, MTM_CollapseRow, wFlags, lRow );
}


// MTblCreateColHdrGrp
// Erzeugt eine Spalten-Header-Gruppe
LPVOID WINAPI MTblCreateColHdrGrp( HWND hWndTbl, LPTSTR lpsText, HARRAY haCols, int iCols, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Array in Liste übertragen
	long lDimCount;
	if ( !Ctd.ArrayDimCount( haCols, &lDimCount ) ) return NULL;
	long lLowerBound;
	if ( !Ctd.ArrayGetLowerBound( haCols, 1, &lLowerBound ) ) return NULL;

	HANDLE hCol;
	int i;
	list<HWND> liCols;
	for ( i = 0; i < iCols; ++i )
	{
		if ( !Ctd.MDArrayGetHandle( haCols, &hCol, lLowerBound + i ) ) return NULL;
		if ( !hCol ) return NULL;
		if ( Ctd.ParentWindow( HWND(hCol) ) != hWndTbl ) return NULL;
		liCols.push_back( HWND(hCol) );
	}
	if ( i < 1 ) return NULL;

	// Text
	tstring sText( _T("") );
	if ( lpsText )
		sText = lpsText;

	// Gruppe erstellen
	CMTblColHdrGrp * pGrp = psc->m_ColHdrGrps->CreateGrp( sText, liCols, dwFlags );
	if ( pGrp )
	{
		// Texte der zugehörigen Spalten-Header anpassen
		list<HWND>::iterator it, itEnd = pGrp->m_Cols.end( );
		for ( it = pGrp->m_Cols.begin( ); it != itEnd; ++it )
		{
			psc->AdaptColHdrText( *it );
		}
	}

	// Neu zeichnen
	psc->InvalidateColHdr( FALSE );

	return pGrp;
}



// MTblDefineButtonCell
// Definiert Button-Zelle
BOOL WINAPI MTblDefineButtonCell( HWND hWndCol, long lRow, BOOL bShow, int iWidth, LPCTSTR lpsText, HANDLE hImage, DWORD dwFlags )
{
	// Ctd da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	// Subclass-Struktur ermitteln
	HWND hWndTbl = Ctd.ParentWindow( hWndCol );
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Aktuelle Definition ermitteln
	CMTblBtn bOld;
	CMTblBtn *pbOld = pRow->m_Cols->GetBtn( hWndCol );
	if ( pbOld )
		bOld = *pbOld;

	// Definieren
	CMTblBtn b;
	b.m_bShow = bShow ? true : false;
	b.m_iWidth = iWidth;
	b.m_sText = lpsText ? lpsText : _T("");
	b.m_hImage = hImage;
	b.m_dwFlags = dwFlags;

	if ( !pRow->m_Cols->SetBtn( hWndCol, b ) )
		return FALSE;

	HWND hWndFocusCol = NULL;
	long lFocusRow = NULL;
	Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol );

	if ( hWndCol == hWndFocusCol && lRow == lFocusRow )
	{
		psc->AdaptListClip( );
	}
	else if ( psc->QueryFlags( MTBL_FLAG_BUTTONS_PERMANENT ) || psc->QueryColumnFlags( hWndCol, MTBL_COL_FLAG_BUTTONS_PERMANENT ) || psc->QueryCellFlags( lRow, hWndCol, MTBL_CELL_FLAG_BUTTONS_PERMANENT ) )
	{
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );
	}

	return TRUE;
}


// MTblDefineButtonColumn
// Definiert Button-Spalte
BOOL WINAPI MTblDefineButtonColumn( HWND hWndCol, BOOL bShow, int iWidth, LPCTSTR lpsText, HANDLE hImage, DWORD dwFlags )
{
	// Ctd da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	// Subclass-Struktur ermitteln
	HWND hWndTbl = Ctd.ParentWindow( hWndCol );
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Aktuelle Definition ermitteln
	CMTblBtn bOld;
	CMTblBtn *pbOld = psc->m_Cols->GetBtn( hWndCol );
	if ( pbOld )
		bOld = *pbOld;

	// Definieren
	CMTblBtn b;
	b.m_bShow = bShow ? true : false;
	b.m_iWidth = iWidth;
	b.m_sText = lpsText ? lpsText : _T("");
	b.m_hImage = hImage;
	b.m_dwFlags = dwFlags;

	if ( !psc->m_Cols->SetBtn( hWndCol, b ) )
		return FALSE;

	// Gesetzte Spalte == Spalte mit Focus?
	if ( hWndCol == MTblGetFocusCol( hWndTbl ) )
	{
		psc->AdaptListClip( );
	}

	// Ggfs. Spalte neu zeichnen
	if ( psc->QueryFlags( MTBL_FLAG_BUTTONS_PERMANENT ) || psc->QueryColumnFlags( hWndCol, MTBL_COL_FLAG_BUTTONS_PERMANENT ) )
	{
		psc->InvalidateCol( hWndCol, FALSE );
	}

	return TRUE;
}


// MTblDefineButtonHotkey
// Definiert Hotkey für Buttons
BOOL WINAPI MTblDefineButtonHotkey( HWND hWndTbl, int iKey, BOOL bCtrl, BOOL bShift, BOOL bAlt )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	psc->m_iBtnKey = max( iKey, 0 );
	psc->m_bBtnKeyShift = bShift;
	psc->m_bBtnKeyCtrl = bCtrl;
	psc->m_bBtnKeyAlt = bAlt;

	return TRUE;
}


// MTblDefineCellMode
// Definiert Zellen-Modus
BOOL WINAPI MTblDefineCellMode( HWND hWndTbl, BOOL bActive, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Flags prüfen
	if ( ( dwFlags & MTBL_CM_FLAG_NO_SELECTION ) && ( dwFlags & MTBL_CM_FLAG_SINGLE_SELECTION ) )
	{
		dwFlags ^= MTBL_CM_FLAG_SINGLE_SELECTION;
	}

	// Daten setzen
	BOOL bActiveChanged = FALSE, bAnyChanges = FALSE;

	if ( dwFlags != psc->m_dwCellModeFlags )
	{
		psc->m_dwCellModeFlags = dwFlags;
		bAnyChanges = TRUE;
	}

	if ( bActive != psc->m_bCellMode )
	{
		if ( bActive )
			//Ctd.TblKillFocus( hWndTbl );
			psc->KillFocusFrameI( );

		psc->m_bCellMode = bActive;

		bActiveChanged = TRUE;
		bAnyChanges = TRUE;
	}

	if ( bAnyChanges )
	{
		int iRedrawMode = UF_REDRAW_INVALIDATE;
		int iSelChangeMode = UF_SELCHANGE_NONE;

		if ( bActive && ( dwFlags & MTBL_CM_FLAG_SINGLE_SELECTION ) )
			iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
		else if ( !bActive || ( dwFlags & MTBL_CM_FLAG_NO_SELECTION  ) )
			iSelChangeMode = UF_SELCHANGE_CLEAR_CELLS;

		if (bActiveChanged)
			psc->m_bNoFocusPaint = FALSE;

		psc->UpdateFocus( TBL_Error, NULL, iRedrawMode, iSelChangeMode  );
	}

	return TRUE;
}


// MTblDefineCellType
// Definiert den Typ einer Zelle
BOOL WINAPI MTblDefineCellType( HWND hWndCol, long lRow, int iType, DWORD dwFlags, int iLines, LPCTSTR lpsCheckedVal, LPCTSTR lpsUnCheckedVal )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Zelle ermitteln
	CMTblCol * pCell = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCell ) return FALSE;

	// Evtl. Zelltyp-Objekt erzeugen
	if ( !pCell->m_pCellType )
		pCell->m_pCellType = new CMTblCellType;
	else
	{
		// Wenn aktueller Zellentyp der Spalte = Zellentyp der Zelle, Zellentyp der Zeile auf NULL setzen
		// Wichtig, damit in BeforeSetFocusCell ein Unterschied erkannt wird
		CMTblCol *pCol = psc->m_Cols->Find( hWndCol );
		if ( pCol && pCol->m_pCellTypeCurr == pCell->m_pCellType )
			pCol->m_pCellTypeCurr = NULL;
	}

	// Typ definieren
	switch ( iType )
	{
	case COL_CellType_Standard:
		pCell->m_pCellType->DefineStandard( );
		break;
	case COL_CellType_PopupEdit:
		pCell->m_pCellType->DefinePopupEdit( dwFlags, iLines );
		break;
	case COL_CellType_DropDownList:		
		pCell->m_pCellType->DefineDropDownList( dwFlags, iLines, NULL );
		psc->SetDropDownListContext( hWndCol, lRow );
		break;
	case COL_CellType_CheckBox:
		pCell->m_pCellType->DefineCheckBox( dwFlags, tstring(lpsCheckedVal), tstring(lpsUnCheckedVal) );
		break;
	default:
		delete pCell->m_pCellType;
		pCell->m_pCellType = NULL;
		break;
	}

	// Focus killen?
	if ( psc->IsFocusCell( hWndCol, lRow ) )
		Ctd.TblKillEdit( psc->m_hWndTbl );

	// Neu zeichnen
	psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );

	//
	return TRUE;
}


// MTblDefineColLines
// Definiert Spaltenlinien
BOOL WINAPI MTblDefineColLines( HWND hWndTbl, int iStyle, COLORREF clr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Stil prüfen
	if ( !psc->IsValidLineStyle( iStyle ) || iStyle == MTLS_UNDEF )
		return FALSE;

	// Setzen
	psc->m_pColsVertLineDef->Style = iStyle;
	psc->m_pColsVertLineDef->Color = clr;

	// Invalidate
	InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblDefineEffCellPropDet
// Definiert die Ermittlung von effektiven Zellen-Eigenschaften
BOOL WINAPI MTblDefineEffCellPropDet( HWND hWndTbl, DWORD dwOrder, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;
	
	// Setzen und ggfs. neu zeichnen
	if ( psc->SetEffCellPropDetOrder( dwOrder ) )
	{
		psc->m_dwEffCellPropDetFlags = dwFlags;
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );
		return TRUE;
	}
	else
		return FALSE;
}


// MTblDefineFocusFrame
// Definiert Fokus-Rahmen
BOOL WINAPI MTblDefineFocusFrame( HWND hWndTbl, int iThickness, int iOffsX, int iOffsY, DWORD dwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Parameter anpassen/prüfen
	iThickness = max( iThickness, 0 );
	/*if ( iThickness + iOffsX < 0 || iThickness + iOffsY < 0 )
		return FALSE;*/

	// Daten setzen
	BOOL bAnyChanges = FALSE;

	if ( iThickness != psc->m_lFocusFrameThick )
	{
		psc->m_lFocusFrameThick = iThickness;
		bAnyChanges = TRUE;
	}

	if ( iOffsX != psc->m_lFocusFrameOffsX )
	{
		psc->m_lFocusFrameOffsX = iOffsX;
		bAnyChanges = TRUE;
	}

	if ( iOffsY != psc->m_lFocusFrameOffsY )
	{
		psc->m_lFocusFrameOffsY = iOffsY;
		bAnyChanges = TRUE;
	}


	if ( dwColor != psc->m_dwFocusFrameColor )
	{
		psc->m_dwFocusFrameColor = dwColor;
		bAnyChanges = TRUE;
	}

	if ( bAnyChanges )
		psc->UpdateFocus( TBL_Error, NULL, UF_REDRAW_INVALIDATE );

	return TRUE;
}


// MTblDefineHighlighting
// Definiert die Hervorhebung für die verschiedenen Items
BOOL WINAPI MTblDefineHighlighting( HWND hWndTbl, long lItem, long lPart, HUDV hUdv, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Entfernen?
	if ( dwFlags & MTDHL_REMOVE )
	{
		psc->m_pHLDM->erase( LONG_PAIR( lItem, lPart ) );
		return TRUE;
	}

	// UDV dereferenzieren
	LPMTBLHILIUDV pUDV = (LPMTBLHILIUDV)Ctd.UdvDeref( hUdv );
	if ( !pUDV )
		return FALSE;

	MTBLHILI hld;

	// Schon in Map?
	HiLiDefMap::iterator it;
	it = psc->m_pHLDM->find( LONG_PAIR( lItem, lPart ) );
	if ( it != psc->m_pHLDM->end( ) )
	{
		hld = (*it).second;
		psc->m_pHLDM->erase( it );
	}

	// Daten in Struktur setzen
	HANDLE hImg;
	Ctd.CvtNumberToULong( &pUDV->BackColor, &hld.dwBackColor );
	Ctd.CvtNumberToULong( &pUDV->TextColor, &hld.dwTextColor );
	Ctd.CvtNumberToLong( &pUDV->FontEnh, &hld.lFontEnh );
	Ctd.CvtNumberToULong( &pUDV->Image, (LPDWORD)&hImg );
	hld.Img.SetHandle( hImg, NULL );
	Ctd.CvtNumberToULong( &pUDV->ImageExp, (LPDWORD)&hImg );
	hld.ImgExp.SetHandle( hImg, NULL );

	// In Map einfügen
	psc->m_pHLDM->insert( HiLiDefMap::value_type( LONG_PAIR( lItem, lPart ), hld ) );

	// Ggf. neu zeichnen
	if ( dwFlags & MTDHL_REDRAW )
		InvalidateRect( hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblDefineHyperlinkCell
// Definiert eine Hyperlink-Zelle
BOOL WINAPI MTblDefineHyperlinkCell( HWND hWndCol, long lRow, BOOL bEnable, DWORD dwFlags )
{
	// CTD da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	// Tabelle ermitteln
	HWND hWndTbl = Ctd.ParentWindow( hWndCol );

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Hyperlink hinzufügen/entfernen
	CMTblHyLnk hl;
	hl.SetEnabled( bEnable );
	hl.m_wFlags = dwFlags;

	return psc->m_Rows->SetCellHyperlink( lRow, hWndCol, hl );
}


// MTblDefineHyperlinkHover
// Definition Hyperlink-HOVER
BOOL WINAPI MTblDefineHyperlinkHover( HWND hWndTbl, BOOL bEnabled, long lAddFontEnh, DWORD dwTextColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	psc->m_bHLHover = bEnabled;
	psc->m_lHLHoverFontEnh = lAddFontEnh;
	psc->m_dwHLHoverTextColor = dwTextColor;

	return TRUE;
}



// MTblDefineItemButton
// Definiert Button für ein Item
BOOL WINAPI MTblDefineItemButton( HUDV hUdvItem, HUDV hUdvBtnDef )
{
	MTBLITEM item;
	if ( !item.FromUdv( hUdvItem ) ) return FALSE;

	MTBLBTNDEF btndef;
	if ( !btndef.FromUdv( hUdvBtnDef ) ) return FALSE;

	return MTblDefineItemButtonC( &item, &btndef );
}


// MTblDefineItemButtonC
// Definiert Button für ein Item, Variante mit C-Strukturen
BOOL WINAPI MTblDefineItemButtonC( LPVOID pItemV, LPVOID pBtnDefV )
{
	// Check Parameter
	if ( !(pItemV && pBtnDefV) ) return FALSE;

	// Void-Pointer in Strukturpointer umwandeln
	LPMTBLITEM pItem = (LPMTBLITEM)pItemV;
	LPMTBLBTNDEF pBtnDef = (LPMTBLBTNDEF)pBtnDefV;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem( pItem );
	if ( !psc ) return FALSE;

	// Zeile prüfen
	CMTblRow * pRow = NULL;
	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
		pRow = pItem->GetRowPtr( );

		// Gültige Zeile?
		if ( !pRow ) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if ( pRow->IsDelAdj( ) ) return FALSE;		

		break;
	}
	
	// Button definieren
	CMTblBtn btn;
	btn.m_bShow = pBtnDef->bVisible ? true : false;
	btn.m_iWidth = pBtnDef->iWidth;
	btn.m_sText = pBtnDef->sText;
	btn.m_hImage = pBtnDef->hImage;
	btn.m_dwFlags = pBtnDef->dwFlags;

	// Button setzen
	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
		if ( !pRow->m_Cols->SetBtn( pItem->WndHandle, btn ) )
			return FALSE;
		break;
	case MTBL_ITEM_COLUMN:
		if ( !psc->m_Cols->SetBtn( pItem->WndHandle, btn ) )
			return FALSE;
		break;
	case MTBL_ITEM_COLHDR:
		if ( !psc->m_Cols->SetHdrBtn( pItem->WndHandle, btn ) )
			return FALSE;
		break;
	}

	// Listclip anpassen
	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
	case MTBL_ITEM_COLUMN:
		if ( pItem->WndHandle == MTblGetFocusCol( psc->m_hWndTbl ) )
			psc->AdaptListClip( );
		break;
	}

	// Neu zeichnen
	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
		if ( psc->QueryFlags( MTBL_FLAG_BUTTONS_PERMANENT ) || psc->QueryColumnFlags( pItem->WndHandle, MTBL_COL_FLAG_BUTTONS_PERMANENT ) || psc->QueryCellFlags( pRow->GetNr( ), pItem->WndHandle, MTBL_CELL_FLAG_BUTTONS_PERMANENT ) )
			psc->InvalidateCell( pItem->WndHandle, pRow->GetNr( ), NULL, FALSE );
		break;
	case MTBL_ITEM_COLUMN:
		if ( psc->QueryFlags( MTBL_FLAG_BUTTONS_PERMANENT ) || psc->QueryColumnFlags( pItem->WndHandle, MTBL_COL_FLAG_BUTTONS_PERMANENT ) )
			psc->InvalidateCol( pItem->WndHandle, FALSE );
		break;
	case MTBL_ITEM_COLHDR:
		psc->InvalidateColHdr( pItem->WndHandle, FALSE );
		break;
	}

	// Fertig
	return TRUE;
}


// MTblDefineItemLine
// Definiert eine Linie für einItem
BOOL WINAPI MTblDefineItemLine(HUDV hUdvItem, HUDV hUdvLineDef, int iLineType)
{
	MTBLITEM item;
	if (!item.FromUdv(hUdvItem)) return FALSE;

	CMTblLineDef ld;
	if (!ld.FromUdv(hUdvLineDef)) return FALSE;

	return MTblDefineItemLineC(&item, &ld, iLineType);
}


// MTblDefineItemLineC
// Definiert eine Linie für ein Item, Variante mit C-Strukturen
BOOL WINAPI MTblDefineItemLineC(LPVOID pItemV, LPVOID pLineDefV, int iLineType)
{
	// Check Parameter
	if (!(pItemV && pLineDefV)) return FALSE;

	// Void-Pointer in Strukturpointer umwandeln
	LPMTBLITEM pItem = (LPMTBLITEM)pItemV;
	LPMTBLLINEDEF pLineDef = (LPMTBLLINEDEF)pLineDefV;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem(pItem);
	if (!psc) return FALSE;

	// Liniendefinition prüfen
	if (!psc->IsValidLineStyle(pLineDef->Style))
		return FALSE;
	if (!(pLineDef->Thickness >= 1 || pLineDef->Thickness == MTBL_LINE_UNDEF_THICKNESS))
		return FALSE;
	if (pLineDef->Style == MTLS_DOT && pLineDef->Thickness > 1)
		return FALSE;

	if (pItem->Type == MTBL_ITEM_ROWS && iLineType == MTLT_HORZ)
	{
		if (pLineDef->Style == MTLS_UNDEF)
			return FALSE;
		if (pLineDef->Thickness == MTBL_LINE_UNDEF_THICKNESS)
			return FALSE;
	}

	if ((pItem->Type == MTBL_ITEM_COLUMNS || pItem->Type == MTBL_ITEM_LAST_LOCKED_COLUMN) && iLineType == MTLT_VERT)
	{
		if (pLineDef->Style == MTLS_UNDEF)
			return FALSE;
		if (pLineDef->Thickness == MTBL_LINE_UNDEF_THICKNESS)
			return FALSE;
	}

	// Item prüfen
	CMTblRow * pRow = NULL;
	CMTblColHdrGrp* pCHG = NULL;
	switch (pItem->Type)
	{
	case MTBL_ITEM_CELL:
	case MTBL_ITEM_ROW:
	case MTBL_ITEM_ROWHDR:
		pRow = pItem->GetRowPtr();

		// Gültige Zeile?
		if (!pRow) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if (pRow->IsDelAdj()) return FALSE;

		break;
	case MTBL_ITEM_COLHDRGRP:
		pCHG = (CMTblColHdrGrp*)pItem->Id;
		if (!pCHG)
			return FALSE;

		if (!psc->m_ColHdrGrps->IsValidGrp(pCHG))
			return FALSE;
		break;
	}

	// Liniendefinition setzen
	switch (pItem->Type)
	{
	case MTBL_ITEM_CELL:
		if (!pRow->m_Cols->SetLineDef(pItem->WndHandle, *pLineDef, iLineType))
			return FALSE;
		break;
	case MTBL_ITEM_ROW:
		if (!pRow->SetLineDef(*pLineDef, iLineType))
			return FALSE;
		break;
	case MTBL_ITEM_ROWS:
		if (iLineType == MTLT_HORZ)
			*psc->m_pRowsHorzLineDef = *pLineDef;
		else
			return FALSE;
		break;
	case MTBL_ITEM_ROWHDR:
		if (!pRow->SetHdrLineDef(*pLineDef, iLineType))
			return FALSE;
		break;
	case MTBL_ITEM_COLUMN:		
		if (!psc->m_Cols->SetLineDef(pItem->WndHandle, *pLineDef, iLineType))
			return FALSE;
		break;
	case MTBL_ITEM_LAST_LOCKED_COLUMN:
		if (iLineType == MTLT_VERT)
			*psc->m_pLastLockedColVertLineDef = *pLineDef;
		else
			return FALSE;
		break;
	case MTBL_ITEM_COLUMNS:
		if (iLineType == MTLT_VERT)
			*psc->m_pColsVertLineDef = *pLineDef;
		else
			return FALSE;
		break;
	case MTBL_ITEM_COLHDR:
		if (!psc->m_Cols->SetHdrLineDef(pItem->WndHandle, *pLineDef, iLineType))
			return FALSE;
		break;
	case MTBL_ITEM_COLHDRGRP:
		if (!psc->m_ColHdrGrps->IsValidGrp(pCHG))
			return false;
		if (!pCHG->SetLineDef(*pLineDef, iLineType))
			return FALSE;
		break;
	case MTBL_ITEM_CORNER:
		if (!psc->m_Corner->SetLineDef(*pLineDef, iLineType))
			return FALSE;
		break;
	}

	// Neu zeichnen
	HWND hWndLastLockedCol = NULL;

	switch (pItem->Type)
	{
	case MTBL_ITEM_CELL:
		psc->InvalidateCell(pItem->WndHandle, pRow->GetNr(), NULL, FALSE);
		break;
	case MTBL_ITEM_ROW:
		psc->InvalidateRow(pRow->GetNr(), FALSE, TRUE, FALSE);
		break;
	case MTBL_ITEM_ROWS:
		InvalidateRect(psc->m_hWndTbl, NULL, FALSE);
		break;
	case MTBL_ITEM_ROWHDR:
		psc->InvalidateRowHdr(pRow->GetNr(), FALSE, TRUE);
		break;
	case MTBL_ITEM_COLUMN:
		psc->InvalidateCol(pItem->WndHandle, FALSE);
		break;
	case MTBL_ITEM_LAST_LOCKED_COLUMN:
		hWndLastLockedCol = psc->GetLastLockedColumn();
		if (hWndLastLockedCol)
			psc->InvalidateCol(hWndLastLockedCol, FALSE);
		break;
	case MTBL_ITEM_COLUMNS:
		InvalidateRect(psc->m_hWndTbl, NULL, FALSE);
		break;
	case MTBL_ITEM_COLHDR:
		psc->InvalidateColHdr(pItem->WndHandle, TRUE, FALSE);
		break;
	case MTBL_ITEM_COLHDRGRP:
		psc->InvalidateColHdr(FALSE);
		break;
	case MTBL_ITEM_CORNER:
		psc->InvalidateCorner(FALSE);
		break;
	}

	// Fertig
	return TRUE;
}


// MTblDefineLastLockedColLine
// Definiert Spaltenlinie der letzten verankerten Spalte
BOOL WINAPI MTblDefineLastLockedColLine( HWND hWndTbl, int iStyle, COLORREF clr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Stil prüfen
	if (!psc->IsValidLineStyle(iStyle) || iStyle == MTLS_UNDEF)
		return FALSE;

	// Setzen
	psc->m_pLastLockedColVertLineDef->Style = iStyle;
	psc->m_pLastLockedColVertLineDef->Color = clr;

	// Invalidate
	InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblDefineRowLines
// Definiert Zeilenlinien
BOOL WINAPI MTblDefineRowLines( HWND hWndTbl, int iStyle, COLORREF clr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Stil prüfen
	if (!psc->IsValidLineStyle(iStyle) || iStyle == MTLS_UNDEF)
		return FALSE;

	// Setzen
	psc->m_pRowsHorzLineDef->Style = iStyle;
	psc->m_pRowsHorzLineDef->Color = clr;

	// Invalidate
	InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblDefineTree
// Definiert Baumstruktur
BOOL WINAPI MTblDefineTree( HWND hWndTbl, HWND hWndTreeCol, int iNodeSize, int iIndent )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( hWndTreeCol )
	{
		// Spalte ok?
		if ( Ctd.GetType( hWndTreeCol ) != TYPE_TableColumn ) return FALSE;
		if ( Ctd.ParentWindow( hWndTreeCol ) != hWndTbl ) return FALSE;
	}

	BOOL bRedraw = ( hWndTreeCol != psc->m_hWndTreeCol || iNodeSize != psc->m_iNodeSize || iIndent != psc->m_iIndent );

	psc->m_hWndTreeCol = hWndTreeCol;
	psc->m_iNodeSize = max( 0, iNodeSize );
	psc->m_iIndent = max( psc->m_iNodeSize, iIndent );

	if ( bRedraw )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblDefineTreeLines
// Definiert Tree-Linien
BOOL WINAPI MTblDefineTreeLines( HWND hWndTbl, int iStyle, COLORREF clr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Stil prüfen
	if (!psc->IsValidLineStyle(iStyle) || iStyle == MTLS_UNDEF)
		return FALSE;

	// Setzen
	psc->m_TreeLineDef.Style = iStyle;
	psc->m_TreeLineDef.Color = clr;

	// Invalidate
	InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblDeleteColHdrGrp
// Löscht eine Spaltenheader-Gruppe
BOOL WINAPI MTblDeleteColHdrGrp( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Spalten merken
	list<HWND> liCols( p->m_Cols );

	// Löschen
	if ( !psc->m_ColHdrGrps->DeleteGrp( p ) ) return FALSE;

	// Ggf. Enter/Leave-Gruppe zurücksetzen
	if ( p == psc->m_pELColHdrGrp )
		psc->m_pELColHdrGrp = NULL;

	// Texte der zugehörigen Spalten-Header anpassen
	list<HWND>::iterator it, itEnd = liCols.end( );
	for ( it = liCols.begin( ); it != itEnd; ++it )
	{
		psc->AdaptColHdrText( *it );
	}

	// Neu zeichnen?
	psc->InvalidateColHdr( FALSE );

	return TRUE;
}


// MTblDeleteDescRows
// Löscht die Nachkommen einer Eltern-Zeile
int WINAPI MTblDeleteDescRows( HWND hWndTbl, long lRow, WORD wMethod )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	CMTblRow *pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( !pRow ) return 0;

	return psc->DeleteDescRows( pRow, wMethod );
}


// MTblDeleteRowUserData
// Löscht Benutzerdaten einer Zeile
BOOL WINAPI MTblDeleteRowUserData( HWND hWndTbl, long lRow, LPTSTR lpsKey )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Löschen
	tstring sKey( lpsKey );
	return psc->m_Rows->DelRowUserData( lRow, sKey );
}


// MTblDisableCell
// Sperrt eine Zelle
BOOL WINAPI MTblDisableCell( HWND hWndCol, long lRow, BOOL bDisable )
{
	// CTD da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	// Tabelle ermitteln
	HWND hWndTbl = Ctd.ParentWindow( hWndCol );

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Setzen
	if ( !psc->m_Rows->SetCellDisabled( lRow, hWndCol, bDisable ) ) return FALSE;

	// Ggfs. Editiermodus beenden
	if ( bDisable )
	{
		// Gesperrte Zelle = Zelle mit Focus?
		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( lRow == lFocusRow && hWndCol == hWndFocusCol )
			{
				Ctd.TblKillEdit( hWndTbl );
			}
		}
	}

	// Ggfs. neu zeichnen
	if ( psc->QueryFlags( MTBL_FLAG_BUTTONS_PERMANENT ) || psc->QueryColumnFlags( hWndCol, MTBL_COL_FLAG_BUTTONS_PERMANENT ) || psc->QueryCellFlags( lRow, hWndCol, MTBL_CELL_FLAG_BUTTONS_PERMANENT ) )
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );
	
	return TRUE;
}


// MTblDisableRow
// Sperrt eine Zeile
BOOL WINAPI MTblDisableRow( HWND hWndTbl, long lRow, BOOL bDisable )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Disable-Flag der Zeile setzen
	if ( !psc->m_Rows->SetRowDisabled( lRow, bDisable ) ) return FALSE;

	// Ggfs. Editiermodus beenden
	if ( bDisable )
	{
		// Gesperrte Zeile = Zeile mit Focus?
		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( lRow == lFocusRow && hWndFocusCol )
			{
				Ctd.TblKillEdit( hWndTbl );
			}
		}
	}

	// Neu zeichnen wegen ggf. vorhandenen permanenten Buttons
	psc->InvalidateRow( lRow, FALSE, TRUE );
	
	return TRUE;
}


// MTblEnableExtMsgs
// Schaltet die Generierung der erweiterten Nachrichten an/ab

BOOL WINAPI MTblEnableExtMsgs( HWND hWndTbl, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Setzen
	psc->m_bExtMsgs = bSet;

	return TRUE;
}


// MTblEnableMWheelScroll
// Schaltet das Scrollen bei WM_MOUSEWHEEL-Nachrichten an/ab
BOOL WINAPI MTblEnableMWheelScroll( HWND hWndTbl, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Setzen
	psc->m_bMWheelScroll = bSet;

	return TRUE;
}


// MTblEnableTipType
// Aktiviert / deaktiviert einen Tip-Typ
BOOL WINAPI MTblEnableTipType( HWND hWndTbl, DWORD dwTipType, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Setzen
	return psc->EnableTipType( dwTipType, bSet );
}


// MTblEnumColHdrGrpCols
// Liefert alle Spalten einer Spaltenheader-Gruppe
long WINAPI MTblEnumColHdrGrpCols( HWND hWndTbl, LPVOID pGrp, HARRAY haCols )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return -1;

	return p->EnumCols( haCols );
}


// MTblEnumColHdrGrps
// Liefert alle Spaltenheader-Gruppen
long WINAPI MTblEnumColHdrGrps( HWND hWndTbl, HARRAY haGrps )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->m_ColHdrGrps->EnumGrps( haGrps );
}


// MTblEnumExportSubstColors
// Liefert alle Export-Ersatzfarben
long WINAPI MTblEnumExportSubstColors( HWND hWndTbl, HARRAY haTblClr, HARRAY haExcClr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	// Check Arrays
	long lDims;
	if ( !Ctd.ArrayDimCount( haTblClr, &lDims ) ) return -1;
	if ( lDims > 1 ) return -1;
	if ( !Ctd.ArrayDimCount( haExcClr, &lDims ) ) return -1;
	if ( lDims > 1 ) return -1;


	// Untergrenzen Array ermitteln
	long lLowExc, lLowTbl;
	if ( !Ctd.ArrayGetLowerBound( haTblClr, 1, &lLowTbl ) ) return -1;
	if ( !Ctd.ArrayGetLowerBound( haExcClr, 1, &lLowExc ) ) return -1;

	// Daten ermitteln
	long lCount;
	SubstColorMap::iterator it;
	NUMBER n;
	for ( it = psc->m_escm->begin( ), lCount = 0; it != psc->m_escm->end( ); ++it )
	{
		if ( !Ctd.CvtULongToNumber( ULONG( (*it).first ), &n ) ) return -1;
		if ( !Ctd.MDArrayPutNumber( haTblClr, &n, lCount + lLowTbl ) ) return -1;

		if ( !Ctd.CvtULongToNumber( ULONG( (*it).second ), &n ) ) return -1;
		if ( !Ctd.MDArrayPutNumber( haExcClr, &n, lCount + lLowExc ) ) return -1;

		++lCount;
	}

	return lCount;
}


// MTblEnumFonts
long WINAPI MTblEnumFonts( HARRAY haFontNames )
{
	// Check Array
	long lDims;
	if ( !Ctd.ArrayDimCount( haFontNames, &lDims ) )
		return -1;
	if ( lDims > 1 ) return -1;

	// Untergrenze Array ermitteln
	long lLower;
	if ( !Ctd.ArrayGetLowerBound( haFontNames, 1, &lLower ) )
		return -1;

	// Fonts ermitteln
	HDC hDC = GetDC( NULL );
	if ( hDC )
	{
		LOGFONT lf;
		MTBLENUM en;

		ZeroMemory( &lf, sizeof( lf ) );
		lf.lfCharSet = DEFAULT_CHARSET;

		en.ha = haFontNames;
		en.lLower = lLower;
		en.lItems = 0;

		EnumFontFamiliesEx( hDC, &lf, (FONTENUMPROC)MTblEnumFontsProc, (LPARAM)&en, 0 );
		ReleaseDC( NULL, hDC );

		return en.lItems;
	}
	else
		return -1;
}


// MTblEnumFontSizes
long WINAPI MTblEnumFontSizes( HSTRING hsFontName, HARRAY haFontSizes )
{
	// Check Array
	long lDims;
	if (!Ctd.ArrayDimCount(haFontSizes, &lDims))
		return -1;
	if (lDims > 1) return -1;

	// Untergrenze Array ermitteln
	long lLower;
	if (!Ctd.ArrayGetLowerBound(haFontSizes, 1, &lLower))
		return -1;

	// Font-Größen ermitteln
	BOOL bError = FALSE;
	HDC hDC = NULL;
	LOGFONT lf;
	long lBufLen;
	LPTSTR lpsBuf = NULL;
	MTBLENUM en;

	ZeroMemory(&lf, sizeof(lf));

	lpsBuf = Ctd.StringGetBuffer(hsFontName, &lBufLen);
	if (lpsBuf)
		lstrcpy(lf.lfFaceName, lpsBuf);
	else
		bError = TRUE;
	Ctd.HStrUnlock(hsFontName);

	if (bError)
		return -1;

	lf.lfCharSet = DEFAULT_CHARSET;

	en.ha = haFontSizes;
	en.lLower = lLower;
	en.lItems = 0;

	hDC = GetDC(NULL);
	if (hDC)
	{
		EnumFontFamiliesEx(hDC, &lf, (FONTENUMPROC)MTblEnumFontSizesProc, (LPARAM)&en, 0);
		ReleaseDC(NULL, hDC);
	}
	else
		return -1;

	return en.lItems;
}


// MTblEnumFontsProc
int CALLBACK MTblEnumFontsProc( ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, int iFontType, LPARAM lParam )
{
	BOOL bAdd = TRUE;
	int iCmp;
	LPMTBLENUM pen = (LPMTBLENUM)lParam;
	long lBufLen = 0;
	long lIndex = pen->lLower + pen->lItems;
	LPTSTR lpsBuf = NULL, lpsBufArr = NULL, lpsFontName = lpelfe->elfLogFont.lfFaceName;
	
	if ( pen->lItems > 0 )
	{
		HSTRING hsArr = 0;
		for( long lc = pen->lLower + pen->lItems - 1; lc >= pen->lLower; lc-- )
		{
			if ( Ctd.MDArrayGetHString( pen->ha, &hsArr, lc ) )
			{
				lpsBufArr = Ctd.StringGetBuffer( hsArr, &lBufLen );
				iCmp = lstrcmp( lpsBufArr, lpsFontName );
				Ctd.HStrUnlock(hsArr);

				if ( iCmp == 0 )
				{
					bAdd = FALSE;
					break;
				}
				else if ( iCmp > 0 )
				{
					lIndex--;
				}
			}
		}
	}
	
	if ( bAdd )
	{
		HSTRING hs = 0;
		lBufLen = Ctd.BufLenFromStrLen( lstrlen( lpsFontName ) );
		Ctd.InitLPHSTRINGParam( &hs, lBufLen );
		lpsBuf = Ctd.StringGetBuffer( hs, &lBufLen );
		lstrcpy( lpsBuf, lpsFontName );
		Ctd.HStrUnlock(hs);

		if ( pen->lItems > 0 )
		{
			HSTRING hsArr = 0;
			for( long lc = pen->lLower + pen->lItems - 1; lc >= lIndex; lc-- )
			{
				if ( Ctd.MDArrayGetHString( pen->ha, &hsArr, lc ) )
				{
					lpsBufArr = Ctd.StringGetBuffer( hsArr, &lBufLen );
					HSTRING hsNew = 0;
					lBufLen = Ctd.BufLenFromStrLen( lstrlen( lpsBufArr ) );
					Ctd.InitLPHSTRINGParam( &hsNew, lBufLen );
					lpsBuf = Ctd.StringGetBuffer( hsNew, &lBufLen );
					lstrcpy( lpsBuf, lpsBufArr );

					Ctd.HStrUnlock(hsArr);
					Ctd.HStrUnlock(hsNew);

					Ctd.MDArrayPutHString( pen->ha, hsNew, lc + 1 );
				}
			}
		}

		Ctd.MDArrayPutHString( pen->ha, hs, lIndex );
		pen->lItems++;
	}

	return 1;
}


// MTblEnumFontSizesProc
int CALLBACK MTblEnumFontSizesProc( ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, int iFontType, LPARAM lParam )
{
	LPMTBLENUM pen = (LPMTBLENUM)lParam;
	NUMBER n;

	if ( iFontType == TRUETYPE_FONTTYPE || iFontType == 0 )
	{
		long laFontSizes[] = {8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72};
		for ( int i = 0; i < sizeof( laFontSizes ) / sizeof( long ); i++ )
		{
			Ctd.CvtLongToNumber( laFontSizes[i], &n );
			Ctd.MDArrayPutNumber( pen->ha, &n, pen->lLower + pen->lItems );
			pen->lItems++;
		}
		return 0;
	}
	else
	{
		BOOL bAdd = TRUE;
		long lFontHt = lpntme->ntmTm.tmHeight - lpntme->ntmTm.tmInternalLeading;
		long lIndex = pen->lLower + pen->lItems;
		long lLogPixelsY = 96;
		
		HDC hDC = GetDC( NULL );
		if ( hDC )
		{
			lLogPixelsY = GetDeviceCaps( hDC, LOGPIXELSY );
			ReleaseDC( NULL, hDC );
		}

		long lFontSize = MulDiv( 72, lFontHt, lLogPixelsY );
		
		if ( pen->lItems > 0 )
		{
			for( long lc = pen->lLower + pen->lItems - 1; lc >= pen->lLower; lc-- )
			{
				if ( Ctd.MDArrayGetNumber( pen->ha, &n, lc ) )
				{
					long lFontSizeArr = Ctd.NumToLong( n );
					if ( lFontSizeArr == lFontSize )
					{
						bAdd = FALSE;
						break;
					}
					else if ( lFontSizeArr > lFontSize )
						lIndex--;
				}
			}
		}
		
		if ( bAdd )
		{
			Ctd.CvtLongToNumber( lFontSize, &n );

			if ( pen->lItems > 0 )
			{
				NUMBER nArr;
				for( long lc = pen->lLower + pen->lItems - 1; lc >= lIndex; lc-- )
				{
					if ( Ctd.MDArrayGetNumber( pen->ha, &nArr, lc ) )
						Ctd.MDArrayPutNumber( pen->ha, &nArr, lc + 1 );
				}
			}

			Ctd.MDArrayPutNumber( pen->ha, &n, lIndex );
			pen->lItems++;
		}

		return 1;
	}
}


// MTblEnumRowUserData
// Liefert alle Benutzerdaten einer Zeile
long WINAPI MTblEnumRowUserData( HWND hWndTbl, long lRow, HARRAY haKeys, HARRAY haValues )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return -1;
	
	return psc->m_Rows->EnumRowUserData( lRow, haKeys, haValues );
}


// MTblExpandRow
// Öffnet eine Zeile
long WINAPI MTblExpandRow( HWND hWndTbl, long lRow, WPARAM wFlags )
{
	return SendMessage( hWndTbl, MTM_ExpandRow, wFlags, lRow );
}


// MTblExportToHTML
// Exportiert in eine HTML-Datei
int WINAPI MTblExportToHTML( HWND hWndTbl, LPCTSTR lpsFile, DWORD dwHTMLFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return MTE_ERR_NOT_SUBCLASSED;

	return psc->ExportToHTML( lpsFile, dwHTMLFlags, dwExpFlags, wRowFlagsOn, wRowFlagsOff, dwColFlagsOn, dwColFlagsOff );
}


// MTblFetchMissingRows
// Fetcht alle fehlenden Zeilen
int WINAPI MTblFetchMissingRows( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->FetchMissingRows( );
}



// MTblFindMergeCell
// Sucht Merge-Zelle
LPVOID WINAPI MTblFindMergeCell( HWND hWndCol, long lRow, int iMode, LPVOID pMergeCells )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return NULL;

	// Parameter prüfen
	if ( !pMergeCells )
		return NULL;

	return (LPVOID) psc->FindMergeCell( hWndCol, lRow, iMode, (LPMTBLMERGECELLS)pMergeCells );
}


// MTblFindNextCol
// Sucht nächste Spalte
int WINAPI MTblFindNextCol( HWND hWndTbl, LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return 0;

	return psc->FindNextCol( phWndCol, dwFlagsOn, dwFlagsOff );
}


// MTblFindNextRow
// Sucht nächste Zeile
BOOL WINAPI MTblFindNextRow( HWND hWndTbl, LPLONG plRow, long lMaxRow, WORD wFlagsOn, WORD wFlagsOff )
{
	// Checks
	if ( !plRow ) return FALSE;
	if ( !Ctd.IsPresent( ) ) return FALSE;

	long lRow = *plRow;
	if ( Ctd.TblFindNextRow( hWndTbl, &lRow, wFlagsOn, wFlagsOff ) )
	{
		if ( lRow > lMaxRow ) return FALSE;
		*plRow = lRow;
		return TRUE;
	}
	else
		return FALSE;
}


// MTblFindPrevCol
// Sucht vorherige Spalte
int WINAPI MTblFindPrevCol( HWND hWndTbl, LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff )
{
	BOOL bFlagsOnOk, bFlagsOffOk, bFound = FALSE;
	int iPos;
	HWND hWndCol;

	// CTD da?
	if ( !Ctd.IsPresent( ) ) return 0;

	if ( *phWndCol )
	{
		iPos = Ctd.TblQueryColumnPos( *phWndCol );
		if ( iPos < 0 ) return 0;
		iPos--;
	}
	else
	{
		HWND hWndNextCol = NULL;
		while( MTblFindNextCol( hWndTbl, &hWndNextCol, 0, 0 ) );
		iPos = Ctd.TblQueryColumnPos( hWndNextCol );
		if ( iPos < 0 ) return 0;
	}

	while ( TRUE )
	{
		hWndCol = Ctd.TblGetColumnWindow( hWndTbl, iPos, COL_GetPos );
		if ( ! hWndCol )
			break;

		bFlagsOnOk = dwFlagsOn ? SendMessage( hWndCol, TIM_QueryColumnFlags, 0, dwFlagsOn ) : TRUE;
		bFlagsOffOk = dwFlagsOff ? !SendMessage( hWndCol, TIM_QueryColumnFlags, 0, dwFlagsOff ) : TRUE;

		if ( bFlagsOnOk && bFlagsOffOk )
		{
			bFound = TRUE;
			break;
		}

		iPos--;
	}

	if ( bFound )
	{
		*phWndCol = hWndCol;
		return iPos;
	}
	else
		return 0;
}


// MTblFindPrevRow
// Sucht vorherige Zeile
BOOL WINAPI MTblFindPrevRow( HWND hWndTbl, LPLONG plRow, long lMinRow, WORD wFlagsOn, WORD wFlagsOff )
{
	// Checks
	if ( !plRow ) return FALSE;
	if ( !Ctd.IsPresent( ) ) return FALSE;

	long lRow = *plRow;
	if ( Ctd.TblFindPrevRow( hWndTbl, &lRow, wFlagsOn, wFlagsOff ) )
	{
		if ( lRow < lMinRow ) return FALSE;
		*plRow = lRow;
		return TRUE;
	}
	else
		return FALSE;
}


// MTblFreeMergeCells
// Gibt Merge-Zellen-Liste frei
BOOL WINAPI MTblFreeMergeCells( HWND hWndTbl, LPVOID pMergeCells )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Parameter prüfen
	if ( !pMergeCells )
		return FALSE;

	// Freigeben
	psc->FreeMergeCells( (LPMTBLMERGECELLS)pMergeCells );
	
	return TRUE;
}


// MTblGetAltRowBackColors
BOOL WINAPI MTblGetAltRowBackColors( HWND hWndTbl, BOOL bSplitRows, LPDWORD pdwColor1, LPDWORD pdwColor2 )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( bSplitRows )
	{
		*pdwColor1 = psc->m_dwSplitRowAlternateBackColor1;
		*pdwColor2 = psc->m_dwSplitRowAlternateBackColor2;

	}
	else
	{
		*pdwColor1 = psc->m_dwRowAlternateBackColor1;
		*pdwColor2 = psc->m_dwRowAlternateBackColor2;
	}

	return TRUE;
}

// MTblGetBestCellHeight
// Liefert die optimale Höhe einer Zelle
int WINAPI MTblGetBestCellHeight( HWND hWndTbl, HWND hWndCol, long lRow, long lCellLeadingY, long lCharBoxY, HFONT hUseFont /*=NULL*/, LPVOID pMergeCells /*=NULL*/, BOOL bCheckboxText /*=FALSE*/, BOOL bIgnoreButtons /*=FALSE*/ )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->GetBestCellHeight( hWndCol, lRow, lCellLeadingY, lCharBoxY, hUseFont, (LPMTBLMERGECELLS)pMergeCells, bCheckboxText, bIgnoreButtons );
}


// MTblGetBestCellWidth
// Liefert die optimale Breite einer Zelle
int WINAPI MTblGetBestCellWidth( HWND hWndTbl, HWND hWndCol, long lRow, long lCellLeadingX, HFONT hUseFont /*=NULL*/, LPVOID pMergeCells /*=NULL*/, BOOL bCheckboxText /*=FALSE*/ )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->GetBestCellWidth( hWndCol, lRow, lCellLeadingX, hUseFont, (LPMTBLMERGECELLS)pMergeCells, bCheckboxText );
}


// MTblGetBestColHdrWidth
// Liefert die optimale Breite eines Spaltenkopfes
int WINAPI MTblGetBestColHdrWidth( HWND hWndTbl, HWND hWndCol, long lCellLeadingX, HFONT hUseFont /*=NULL*/ )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->GetBestColHdrWidth( hWndCol, lCellLeadingX, hUseFont );
}


// MTblGetCellBackColor
// Liefert die Hintergrundfarbe einer Zelle
BOOL WINAPI MTblGetCellBackColor( HWND hWndCol, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
			dwColor = DWORD( pCol->m_BackColor.Get( ) );
	}

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetCellFont
// Liefert den Font einer Zelle
BOOL WINAPI MTblGetCellFont( HWND hWndCol, long lRow, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Font ermitteln
	CMTblFont f;
	CMTblFont *pf = &f;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
			pf = &pCol->m_Font;
	}

	// Font zurückgeben
	return pf->Get( phsName, plSize, plEnh );
}


// MTblGetCellIndentLevel
// Liefert die Einrückungs-Ebene einer Zelle
BOOL WINAPI MTblGetCellIndentLevel( HWND hWndCol, long lRow, LPLONG plIndentLevel )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) )	return FALSE;
	
	// Einrückung setzen
	*plIndentLevel = psc->GetCellIndentLevel( lRow, hWndCol );

	return TRUE;
}


// MTblGetCellImage
// Liefert das normale Bild einer Zelle
BOOL WINAPI MTblGetCellImage( HWND hWndCol, long lRow, LPHANDLE phImage )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Bild ermitteln
	HANDLE hImage = NULL;
	
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
		{
			hImage = pCol->m_Image.GetHandle( );
		}
	}

	// In Parameter übertragen
	*phImage = hImage;

	return TRUE;
}


// MTblGetCellImageExp
// Liefert das Bild "expandierte Zeile" einer Zelle
BOOL WINAPI MTblGetCellImageExp( HWND hWndCol, long lRow, LPHANDLE phImage )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Bild ermitteln
	HANDLE hImage = NULL;
	
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
		{
			hImage = pCol->m_Image2.GetHandle( );
		}
	}

	// In Parameter übertragen
	*phImage = hImage;

	return TRUE;
}


// MTblGetCellMerge
// Liefert Zellenverknüpfung
BOOL WINAPI MTblGetCellMerge( HWND hWndCol, long lRow, LPINT piCount )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Ermitteln
	int iCount = 0;

	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
		iCount = pRow->m_Cols->GetMerged( hWndCol );

	// In Parameter übertragen
	*piCount = iCount;

	return TRUE;
}


// MTblGetCellMergeEx
// Liefert erweiterte Zellenverknüpfung ( Spalten + Zeilen )
BOOL WINAPI MTblGetCellMergeEx( HWND hWndCol, long lRow, LPINT piColCount, LPINT piRowCount )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Ermitteln
	int iColCount = 0, iRowCount = 0;

	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		iColCount = pRow->m_Cols->GetMerged( hWndCol );
		iRowCount = pRow->m_Cols->GetMergedRows( hWndCol );
	}

	// In Parameter übertragen
	*piColCount = iColCount;
	*piRowCount = iRowCount;

	return TRUE;
}


// MTblGetCellRect
// Liefert das Rechteck einer Zelle
int WINAPI MTblGetCellRect( HWND hWndCol, long lRow, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom )
{
	return MTblGetCellRectEx( hWndCol, lRow, piLeft, piTop, piRight, piBottom, MTGR_VISIBLE_PART );
}


// MTblGetCellRectEx
// Liefert das Rechteck einer Zelle ( erweiterte Version )
int WINAPI MTblGetCellRectEx( HWND hWndCol, long lRow, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return MTGR_RET_ERROR;

	// Rechteck ermitteln
	RECT r;
	int iRet = psc->GetCellRect( hWndCol, lRow, NULL, &r, dwFlags );

	*piLeft = r.left;
	*piTop = r.top;
	*piRight = r.right;
	*piBottom = r.bottom;

	return iRet;
}


// MTblGetCellTextColor
// Liefert die Textfarbe einer Zelle
BOOL WINAPI MTblGetCellTextColor( HWND hWndCol, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
			dwColor = DWORD( pCol->m_TextColor.Get( ) );
	}

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetCellType
// Liefert Zelltyp
int WINAPI MTblGetCellType( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return -1;

	// Definition ermitteln
	CMTblCellType *pct = NULL;
	if ( lRow != TBL_Error )
	{
		// Gültige Zeile?
		if ( !psc->IsRow( lRow ) ) return -1;

		// Zeile mit TBL_Adjust gelöscht?
		if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return -1;

		// Zelltyp ermitteln		
		CMTblRow *pRow = psc->m_Rows->GetRowPtr( lRow );
		if ( pRow )
		{
			CMTblCol *pCell = pRow->m_Cols->Find( hWndCol );
			if ( pCell && pCell->m_pCellType )
			{
				pct = pCell->m_pCellType;
			}
		}
	}
	else
	{
		CMTblCol *pCol = psc->m_Cols->Find( hWndCol );
		if ( pCol && pCol->m_pCellType )
		{
			pct = pCol->m_pCellType;
		}
	}

	// Typ zurückliefern
	if ( pct )
		return pct->m_iType;
	else
		return COL_CellType_Undefined;
}


// MTblGetChildRowCount
// Liefert die Anzahl der Kindzeilen einer Zeile
int WINAPI MTblGetChildRowCount( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	// Zeiger auf Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( !pRow ) return 0;

	// Anzahl Kindzeilen ermitteln
	return pRow->GetChildCount( );
}


// MTblGetColHdrRect
// Liefert das sichtbare Rechteck eines Spaltenkopfes
int WINAPI MTblGetColHdrRect( HWND hWndCol, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom )
{
	return MTblGetColHdrRectEx( hWndCol, piLeft, piTop, piRight, piBottom, MTGR_VISIBLE_PART );
}

// MTblGetColHdrRectEx
// Liefert das Rechteck eines Spaltenkopfes ( erweiterte Version )
int WINAPI MTblGetColHdrRectEx( HWND hWndCol, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return MTGR_RET_ERROR;

	RECT r;
	int iRet = psc->GetColHdrRect( hWndCol, r, dwFlags );

	*piLeft = r.left;
	*piTop = r.top;
	*piRight = r.right;
	*piBottom = r.bottom;

	return iRet;	
}


// MTblGetColSelChanges
// Liefert die Änderungen der Spaltenselektion
long WINAPI MTblGetColSelChanges( HWND hWndTbl, HARRAY haCols, HARRAY haChanges )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->GetColSelChanges( haCols, haChanges );
}


// MTblGetColumnBackColor
// Liefert die Hintergrundfarbe einer Spalte
BOOL WINAPI MTblGetColumnBackColor( HWND hWndCol, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblCol * pCol = psc->m_Cols->Find( hWndCol );
	if ( pCol )
		dwColor = DWORD( pCol->m_BackColor.Get( ) );

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetColumnFont
// Liefert den Font einer Spalte
BOOL WINAPI MTblGetColumnFont( HWND hWndCol, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Font ermitteln
	CMTblFont f;
	CMTblFont *pf = &f;

	CMTblCol * pCol = psc->m_Cols->Find( hWndCol );
	if ( pCol )
		pf = &pCol->m_Font;

	// Font zurückgeben
	return pf->Get( phsName, plSize, plEnh );
}


// MTblGetColumnTextColor
// Liefert die Textfarbe einer Spalte
BOOL WINAPI MTblGetColumnTextColor( HWND hWndCol, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblCol * pCol = psc->m_Cols->Find( hWndCol );
	if ( pCol )
		dwColor = DWORD( pCol->m_TextColor.Get( ) );

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetColumnTitle
// Liefert den Titel einer Spalte
BOOL WINAPI MTblGetColumnTitle( HWND hWndCol, LPHSTRING phsTitle )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	return psc->GetColumnTitle( hWndCol, phsTitle );
}



// MTblGetCornerBackColor
// Liefert Corner-Hintergrundfarbe
BOOL WINAPI MTblGetCornerBackColor( HWND hWndTbl, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->m_Corner->m_BackColor.Get( );

	return TRUE;
}


// MTblGetCornerImage
// Liefert das Bild vom Corner
BOOL WINAPI MTblGetCornerImage( HWND hWndTbl, LPHANDLE phImage )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Bild ermitteln
	*phImage = psc->m_Corner->m_Image.GetHandle( );

	return TRUE;
}


// MTblGetColumnHdrBackColor
// Liefert die Hintergrundfarbe eines Spaltenheaders
BOOL WINAPI MTblGetColumnHdrBackColor( HWND hWndCol, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblCol * pCol = psc->m_Cols->FindHdr( hWndCol );
	if ( pCol )
		dwColor = DWORD( pCol->m_BackColor.Get( ) );

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetColumnHdrFont
// Liefert den Font eines Spaltenheaders
BOOL WINAPI MTblGetColumnHdrFont( HWND hWndCol, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Font ermitteln
	CMTblFont f;
	CMTblFont *pf = &f;

	CMTblCol * pCol = psc->m_Cols->FindHdr( hWndCol );
	if ( pCol )
		pf = &pCol->m_Font;

	// Font zurückgeben
	return pf->Get( phsName, plSize, plEnh );
}


// MTblGetColumnHdrImage
// Liefert das Bild eines Spaltenheaders
BOOL WINAPI MTblGetColumnHdrImage( HWND hWndCol, LPHANDLE phImage )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Bild ermitteln
	HANDLE hImage = NULL;
	
	CMTblCol * pCol = psc->m_Cols->FindHdr( hWndCol );
	if ( pCol )
	{
		hImage = pCol->m_Image.GetHandle( );
	}

	// In Parameter übertragen
	*phImage = hImage;

	return TRUE;
}


// MTblGetColumnHdrTextColor
// Liefert die Textfarbe eines Spaltenheaders
BOOL WINAPI MTblGetColumnHdrTextColor( HWND hWndCol, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblCol * pCol = psc->m_Cols->FindHdr( hWndCol );
	if ( pCol )
		dwColor = DWORD( pCol->m_TextColor.Get( ) );

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetColXCoord
// Liefert die X-Koordinaten ( links, rechts ) einer Spalte
int MTblGetColXCoord( HWND hWndTbl, HWND hWndCol, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return MTGR_RET_ERROR;

	return psc->GetColXCoord( hWndCol, NULL, ppt, pptVis );
}


// MTblGetColHdr
// Liefert das Spalten-Window-Handle, wenn die übergebenen Koordinaten innerhalb
// eines Spalten-Kopfes liegen, ansonsten NULL
HWND MTblGetColHdr( HWND hWndTbl, int iX, int iY )
{
	// CTD da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	DWORD dwFlags;
	HWND hWndCol = NULL;
	long lRow;

	if ( !Ctd.TblObjectsFromPoint( hWndTbl, iX, iY, &lRow, &hWndCol, &dwFlags ) ) return NULL;
	if ( !( dwFlags & TBL_YOverColumnHeader ) ) return NULL;

	return hWndCol;
}


// MTblGetColHdrGrp
// Liefert die Spaltenheader-Gruppe, zu der eine Spalte gehört
LPVOID WINAPI MTblGetColHdrGrp( HWND hWndCol )
{
	// Ctd da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	// Tabelle ermitteln
	HWND hWndTbl = Ctd.ParentWindow( hWndCol );

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_ColHdrGrps->GetGrpOfCol( hWndCol );
}


// MTblGetColHdrGrpBackColor
// Liefert die Hintergrundfarbe einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetColHdrGrpBackColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// In Parameter übertragen
	*pdwColor = p->m_BackColor.Get( );

	return TRUE;
}


// MTblGetColHdrGrpFont
// Liefert den Font einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetColHdrGrpFont( HWND hWndTbl, LPVOID pGrp, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Font zurückgeben
	return p->m_Font.Get( phsName, plSize, plEnh );
}


// MTblGetColHdrGrpImage
// Liefert das Bild einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetColHdrGrpImage( HWND hWndTbl, LPVOID pGrp, LPHANDLE phImage )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Bild ermitteln
	*phImage = p->m_Image.GetHandle( );

	return TRUE;
}


// MTblGetColHdrGrpText
// Liefert den Text einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetColHdrGrpText( HWND hWndTbl, LPVOID pGrp, LPHSTRING phsText )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Text ermitteln
	if ( !Ctd.InitLPHSTRINGParam( phsText, Ctd.BufLenFromStrLen( p->m_Text.size( ) ) ) ) return FALSE;
	long lLen;
	LPTSTR lpsBuf = Ctd.StringGetBuffer( *phsText, &lLen );
	lstrcpy( lpsBuf, p->m_Text.c_str( ) );
	Ctd.HStrUnlock(*phsText);

	return TRUE;
}


// MTblGetColHdrGrpTextColor
// Liefert die Textfarbe einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetColHdrGrpTextColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Farbe ermitteln
	*pdwColor = p->m_TextColor.Get( );

	return TRUE;
}


// MTblGetColHdrGrpTextLineCount
// Liefert die Anzahl der Textzeilen einer Spaltenheader-Gruppe
int WINAPI MTblGetColHdrGrpTextLineCount( HWND hWndTbl, LPVOID pGrp )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return 0;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return 0;

	return p->GetNrOfTextLines( );
}

// MTblGetEditHandle
// Liefert das Window-Handle des aktuellen Edit-Controls
HANDLE WINAPI MTblGetEditHandle(HWND hWndTbl)
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass(hWndTbl);
	if (!psc) return NULL;

	if (!IsWindow(psc->m_hWndEdit))
		return NULL;

	return psc->m_hWndEdit;
}


// MTblGetEffCellBackColor
// Liefert die effektive Hintergrundfarbe einer Zelle
BOOL WINAPI MTblGetEffCellBackColor( HWND hWndCol, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffCellBackColor( lRow, hWndCol );

	return TRUE;
}


// MTblGetEffCellFont
// Liefert den effektiven Font einer Zelle
BOOL WINAPI MTblGetEffCellFont( HWND hWndCol, long lRow, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// DC ermitteln
	HDC hDC = GetDC( psc->m_hWndTbl );
	if ( !hDC ) return FALSE;

	// Effektiven Font ermitteln
	BOOL bError = FALSE, bMustDelete;
	HFONT hFont = psc->GetEffCellFont( hDC, lRow, hWndCol, &bMustDelete );

	// Fontbeschreibung ermitteln
	CMTblFont f;
	if ( f.Set( hDC, hFont ) )
	{
		if ( !f.Get( phsName, plSize, plEnh ) )
			bError = TRUE;
	}
	else
		bError = TRUE;

	if ( bMustDelete )
		DeleteObject( hFont );

	// DC freigeben
	ReleaseDC( psc->m_hWndTbl, hDC );

	return !bError;
}


// MTblGetEffCellFontHandle
// Liefert das Handle des effektiven Fonts einer Zelle
HFONT WINAPI MTblGetEffCellFontHandle( HWND hWndCol, long lRow, LPBOOL pbMustDelete )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// DC ermitteln
	HDC hDC = GetDC( psc->m_hWndTbl );
	if ( !hDC ) return FALSE;

	// Effektiven Font ermitteln
	BOOL bError = FALSE, bMustDelete = FALSE;
	HFONT hFont = psc->GetEffCellFont( hDC, lRow, hWndCol, &bMustDelete );
	*pbMustDelete = bMustDelete;

	// DC freigeben
	ReleaseDC( psc->m_hWndTbl, hDC );

	return hFont;
}


// MTblGetEffCellIndent
// Liefert die effektive Einrückung einer Zelle
BOOL WINAPI MTblGetEffCellIndent( HWND hWndCol, long lRow, LPLONG plIndent )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Ermitteln
	*plIndent = psc->GetEffCellIndent( lRow, hWndCol );

	return TRUE;
}


// MTblGetEffCellImage
// Liefert das effektive Bild einer Zelle
HANDLE WINAPI MTblGetEffCellImage( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return NULL;

	// Bild ermitteln
	CMTblImg Img = psc->GetEffCellImage( lRow, hWndCol );

	return Img.GetHandle( );
}


// MTblGetEffCellTextColor
// Liefert die effektive Textfarbe einer Zelle
BOOL WINAPI MTblGetEffCellTextColor( HWND hWndCol, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffCellTextColor( lRow, hWndCol );

	return TRUE;
}


// MTblGetEffCellTextJustify
// Liefert die effektive horizontale Textausrichtung einer Zelle
int WINAPI MTblGetEffCellTextJustify( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return -1;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return -1;

	// Ermitteln
	return psc->GetEffCellTextJustify( lRow, hWndCol );
}


// MTblGetEffCellTextVAlign
// Liefert die effektive vertikale Textausrichtung einer Zelle
int WINAPI MTblGetEffCellTextVAlign( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return -1;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return -1;

	// Ermitteln
	return psc->GetEffCellTextVAlign( lRow, hWndCol );

	return TRUE;
}


// MTblGetEffColHdrGrpBackColor
// Liefert die effektive Hintergrundfarbe einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetEffColHdrGrpBackColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffColHdrGrpBackColor( p );

	return TRUE;
}


// MTblGetEffColHdrGrpFont
// Liefert den effektiven Font einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetEffColHdrGrpFont( HWND hWndTbl, LPVOID pGrp, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// DC ermitteln
	HDC hDC = GetDC( hWndTbl );
	if ( !hDC ) return FALSE;

	// Effektiven Font ermitteln
	BOOL bError = FALSE, bMustDelete;
	HFONT hFont = psc->GetEffColHdrGrpFont( hDC, p, &bMustDelete );

	// Fontbeschreibung ermitteln
	CMTblFont f;
	if ( f.Set( hDC, hFont ) )
	{
		if ( !f.Get( phsName, plSize, plEnh ) )
			bError = TRUE;
	}
	else
		bError = TRUE;

	if ( bMustDelete )
		DeleteObject( hFont );

	// DC freigeben
	ReleaseDC( hWndTbl, hDC );

	return !bError;
}


// MTblGetEffColHdrGrpFontHandle
// Liefert das Handle des effektiven Fonts einer Spaltenheader-Gruppe
HFONT WINAPI MTblGetEffColHdrGrpFontHandle( HWND hWndTbl, LPVOID pGrp, LPBOOL pbMustDelete )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return NULL;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return NULL;

	// DC ermitteln
	HDC hDC = GetDC( hWndTbl );
	if ( !hDC ) return NULL;

	// Effektiven Font ermitteln
	BOOL bError = FALSE, bMustDelete;
	HFONT hFont = psc->GetEffColHdrGrpFont( hDC, p, &bMustDelete );
	*pbMustDelete = bMustDelete;	
	
	// DC freigeben
	ReleaseDC( hWndTbl, hDC );

	return hFont;
}


// MTblGetEffColHdrGrpTextColor
// Liefert die effektive Textfarbe einer Spaltenheader-Gruppe
BOOL WINAPI MTblGetEffColHdrGrpTextColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffColHdrGrpTextColor( p );

	return TRUE;
}


// MTblGetEffColumnHdrBackColor
// Liefert die effektive Hintergrundfarbe eines Spaltenkopfes
BOOL WINAPI MTblGetEffColumnHdrBackColor( HWND hWndCol, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffColHdrBackColor( hWndCol );

	return TRUE;
}


// MTblGetEffColumnHdrFont
// Liefert den effektiven Font eines Spaltenkopfes
BOOL WINAPI MTblGetEffColumnHdrFont( HWND hWndCol, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// DC ermitteln
	HDC hDC = GetDC( psc->m_hWndTbl );
	if ( !hDC ) return FALSE;

	// Effektiven Font ermitteln
	BOOL bError = FALSE, bMustDelete;
	HFONT hFont = psc->GetEffColHdrFont( hDC, hWndCol, &bMustDelete );

	// Fontbeschreibung ermitteln
	CMTblFont f;
	if ( f.Set( hDC, hFont ) )
	{
		if ( !f.Get( phsName, plSize, plEnh ) )
			bError = TRUE;
	}
	else
		bError = TRUE;

	if ( bMustDelete )
		DeleteObject( hFont );

	// DC freigeben
	ReleaseDC( psc->m_hWndTbl, hDC );

	return !bError;
}


// MTblGetEffColumnHdrFontHandle
// Liefert das Handle des effektiven Fonts eines Spaltenkopfes
HFONT WINAPI MTblGetEffColumnHdrFontHandle( HWND hWndCol, LPBOOL pbMustDelete )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return NULL;

	// DC ermitteln
	HDC hDC = GetDC( psc->m_hWndTbl );
	if ( !hDC ) return FALSE;

	// Effektiven Font ermitteln
	BOOL bError = FALSE, bMustDelete;
	HFONT hFont = psc->GetEffColHdrFont( hDC, hWndCol, &bMustDelete );
	*pbMustDelete = bMustDelete;	
	
	// DC freigeben
	ReleaseDC( psc->m_hWndTbl, hDC );

	return hFont;
}


// MTblGetEffColumnHdrFreeBackColor
// Liefert die effektive Hintergrundfarbe des freien Spaltenheader-Bereichs
BOOL WINAPI MTblGetEffColumnHdrFreeBackColor( HWND hWndTbl, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffColHdrFreeBackColor( );

	return TRUE;
}


// MTblGetEffColumnHdrTextColor
// Liefert die effektive Textfarbe eines Spaltenkopfes
BOOL WINAPI MTblGetEffColumnHdrTextColor( HWND hWndCol, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffColHdrTextColor( hWndCol );

	return TRUE;
}


// MTblGetEffCornerBackColor
// Liefert die effektive Hintergrundfarbe des Corners
BOOL WINAPI MTblGetEffCornerBackColor( HWND hWndTbl, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffCornerBackColor( );

	return TRUE;
}


// MTblGetEffHighlighting
// Liefert die effektive Hervorhebung für ein Item
BOOL WINAPI MTblGetEffHighlighting( HUDV hUdvItem, HUDV hUdvHili )
{
	CMTblItem item;
	item.FromUdv( hUdvItem );

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem( &item );
	if ( !psc ) return FALSE;

	// UDV Highlighting dereferenzieren
	LPMTBLHILIUDV pUDV = (LPMTBLHILIUDV)Ctd.UdvDeref( hUdvHili );
	if ( !pUDV )
		return FALSE;

	// Eff. Hervorhebung ermitteln
	MTBLHILI hili;
	BOOL bOk = psc->GetEffItemHighlighting( item, hili );

	// Daten in UDV setzen
	Ctd.CvtULongToNumber( hili.dwBackColor, &pUDV->BackColor );
	Ctd.CvtULongToNumber( hili.dwTextColor, &pUDV->TextColor );
	Ctd.CvtLongToNumber( hili.lFontEnh, &pUDV->FontEnh );
	Ctd.CvtULongToNumber( (DWORD)hili.Img.GetHandle( ), &pUDV->Image );
	Ctd.CvtULongToNumber( (DWORD)hili.ImgExp.GetHandle( ), &pUDV->ImageExp );

	return bOk;
}


// MTblGetEffRowHdrBackColor
// Liefert die effektive Hintergrundfarbe eines Zeilenkopfes
BOOL WINAPI MTblGetEffRowHdrBackColor( HWND hWndTbl, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( lRow != TBL_Error && !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->GetEffRowHdrBackColor( lRow );

	return TRUE;
}


// MTblGetEffRowHeight
// Liefert die effektive Höhe einer Zeile
BOOL WINAPI MTblGetEffRowHeight( HWND hWndTbl, long lRow, LPLONG plHeight )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Höhe ermitteln
	long lHeight = psc->GetEffRowHeight( lRow );
	if ( lHeight < 0 )
		return FALSE;

	// Höhe setzen + return
	*plHeight = lHeight;
	return TRUE;
}


// MTblGetExportSubstColor
// Liefert eine Export-Ersatzfarbe
DWORD WINAPI MTblGetExportSubstColor( HWND hWndTbl, DWORD dwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return MTBL_COLOR_UNDEF;

	// Farbe suchen
	return psc->GetExportSubstColor( dwColor );
}

// MTblGetFirstChildRow
// Liefert die erste Kindzeile einer Zeile
long WINAPI MTblGetFirstChildRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	CMTblRow * pRow, * pChildRow;

	// Zeiger auf Zeile ermitteln
	pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( !pRow ) return TBL_Error;

	// Erste Kindzeile ermitteln
	BOOL bBottomUp = ( lRow >= 0 ? psc->m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP : psc->m_dwTreeFlags & MTBL_TREE_FLAG_SPLIT_BOTTOM_UP );
	pChildRow = psc->m_Rows->GetFirstChildRow( pRow, bBottomUp );
	if ( !pChildRow ) return TBL_Error;

	return pChildRow->GetNr( );
}


// MTblGetColumnHdrFreeBackColor
// Liefert die Hintergrundfarbe des freien Spaltenheader-Bereichs
BOOL WINAPI MTblGetColumnHdrFreeBackColor( HWND hWndTbl, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*pdwColor = psc->m_dwColHdrFreeBackColor;

	return TRUE;
}


// MTblGetFocusCol
// Liefert die Spalte die den Focus hat
HWND MTblGetFocusCol( HWND hWndTbl )
{
	HWND hWndCol = NULL;
	long lRow;
	Ctd.TblQueryFocus( hWndTbl, &lRow, &hWndCol );
	return hWndCol;
}


// MTblGetHeadersBackColor
// Liefert Hintergrundfarbe der Header
BOOL WINAPI MTblGetHeadersBackColor( HWND hWndTbl, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farbe ermitteln
	*pdwColor = psc->m_dwHdrsBackColor;

	return TRUE;
}


// MTblGetItem
// Liefert Daten zum einem Objekt
BOOL WINAPI MTblGetItem( HWND hWndTbl, LPARAM lID, HUDV hUdv )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Item-Struktur ermitteln
	LPMTBLITEM pItem = NULL;
	if ( lID == (LPARAM)psc->m_pELItem )
		pItem = psc->m_pELItem;
	if ( !pItem )
		return FALSE;

	// In UDV übertragen
	if ( !pItem->ToUdv( hUdv ) ) return FALSE;

	// Fertig
	return TRUE;
}


// MTblGetIndentPerLevel
// Liefert die Einrückung in Pixel pro Level
BOOL WINAPI MTblGetIndentPerLevel( HWND hWndTbl, LPLONG plIndentPerLevel )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Einrückung ermitteln
	*plIndentPerLevel = psc->m_lIndentPerLevel;

	return TRUE;
}


// MTblGetLanguageFile
BOOL WINAPI MTblGetLanguageFile( int iLanguage, LPHSTRING phsLanguageFile )
{
	LanguageMap::iterator it = glm.find( iLanguage );
	if ( it == glm.end( ) )
	{
		return FALSE;
	}
	else
	{
		BOOL bOk = TRUE;
		tstring sLanguageFile = (*it).second;

		if ( !Ctd.InitLPHSTRINGParam( phsLanguageFile, Ctd.BufLenFromStrLen( sLanguageFile.size( ) ) ) ) return FALSE;
		long lLen;
		LPTSTR lpsBuf = Ctd.StringGetBuffer( *phsLanguageFile, &lLen );
		if (lpsBuf)
			lstrcpy(lpsBuf, sLanguageFile.c_str());
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsLanguageFile);

		return bOk;
	}
}


// MTblGetLanguageString
BOOL WINAPI MTblGetLanguageString( int iLanguage, LPCTSTR lpsSection, LPCTSTR lpsKey, LPHSTRING phsString )
{
	if ( !lpsSection )
		return FALSE;

	if ( !lpsKey )
		return FALSE;

	LanguageMap::iterator it = glm.find( iLanguage );
	if ( it == glm.end( ) )
	{
		return FALSE;
	}
	else
	{		
		tstring sLanguageFile = (*it).second;

		TCHAR szString[8196] = _T("");
		DWORD dwLen = GetPrivateProfileString( lpsSection, lpsKey, _T(""), szString, sizeof( szString ) - 1, sLanguageFile.c_str( ) );
		if ( dwLen == 0 )
			return FALSE;

		if ( !Ctd.InitLPHSTRINGParam( phsString, Ctd.BufLenFromStrLen( lstrlen( szString ) ) ) ) return FALSE;

		BOOL bOk = TRUE;
		long lLen;
		LPTSTR lpsBuf = Ctd.StringGetBuffer( *phsString, &lLen );
		if (lpsBuf)
			lstrcpy(lpsBuf, szString);
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsString);

		return bOk;
	}
}

// MTblGetLanguageStringC
int WINAPI MTblGetLanguageStringC( int iLanguage, LPCTSTR lpsSection, LPCTSTR lpsKey, LPTSTR lpsString )
{
	if ( !lpsSection )
		return -1;

	if ( !lpsKey )
		return -1;

	LanguageMap::iterator it = glm.find( iLanguage );
	if ( it == glm.end( ) )
	{
		return -1;
	}
	else
	{		
		tstring sLanguageFile = (*it).second;

		TCHAR szString[8196] = _T("");
		DWORD dwLen = GetPrivateProfileString( lpsSection, lpsKey, _T(""), szString, sizeof( szString ) - 1, sLanguageFile.c_str( ) );

		if ( dwLen > 0 && lpsString )
			lstrcpy( lpsString, szString );

		return (int)dwLen;
	}
}


// MTblGetLastMergedCell
// Liefert die letzte verbundene Zelle
BOOL WINAPI MTblGetLastMergedCell( HWND hWndCol, long lRow, LPHWND phWndLastMergedCol, LPLONG plLastMergedRow )
{
	*phWndLastMergedCol = NULL;
	*plLastMergedRow = TBL_Error;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	return psc->GetLastMergedCellEx( hWndCol, lRow, *phWndLastMergedCol, *plLastMergedRow );
}


// MTblGetLastMergedRow
// Liefert die letzte verbundene Zeile
long WINAPI MTblGetLastMergedRow( HWND hWndTbl, long lRow )
{

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return TBL_Error;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return TBL_Error;

	return psc->GetLastMergedRow( lRow );
}


// MTblGetLastSortDef
// Liefert die letzte Sortierdefinition
long WINAPI MTblGetLastSortDef( HWND hWndTbl, HARRAY hWndaCols, HARRAY naSortFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	long lLower;

	// Spaltenarray füllen
	if ( !Ctd.ArrayGetLowerBound( hWndaCols, 1, &lLower ) ) return FALSE;
	HwndList::iterator ith;
	long lCols;
	for ( lCols = 0, ith = psc->m_LastSortCols->begin( ); ith != psc->m_LastSortCols->end( ); ++lCols, ++ith )
	{
		if ( !Ctd.MDArrayPutHandle( hWndaCols, (HANDLE)*ith, lLower + lCols ) ) return -1;
	}

	// Spaltenflag-Array füllen
	if ( !Ctd.ArrayGetLowerBound( naSortFlags, 1, &lLower ) ) return FALSE;
	long lFlags;
	LongList::iterator itl;
	NUMBER n;
	for ( lFlags = 0, itl = psc->m_LastSortFlags->begin( ); itl != psc->m_LastSortFlags->end( ); ++lFlags, ++itl )
	{
		if ( !Ctd.CvtLongToNumber( *itl, &n ) ) return -1;
		if ( !Ctd.MDArrayPutNumber( naSortFlags, &n, lLower + lFlags ) ) return -1;
	}

	if ( lCols != lFlags ) return -1;

	return lCols;
}


// MTblGetMergeCell
// Liefert Merge-Zelle
BOOL WINAPI MTblGetMergeCell( HWND hWndCol, long lRow, LPHWND phWndMergeCol, LPLONG plMergeRow )
{
	*phWndMergeCol = NULL;
	*plMergeRow = TBL_Error;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	return psc->GetMergeCellEx( hWndCol, lRow, *phWndMergeCol, *plMergeRow );
}


// MTblGetMergeCells
// Erzeugt Merge-Zellen-Liste und liefert Zeiger darauf zurück
LPVOID WINAPI MTblGetMergeCells( HWND hWndTbl, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return NULL;

	return (LPVOID)psc->GetMergeCells( TRUE, wRowFlagsOn, wRowFlagsOff, dwColFlagsOn, dwColFlagsOff );
}


// MTblGetMergeCol
// Liefert Merge-Spalte
HWND WINAPI MTblGetMergeCol( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return NULL;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return NULL;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return NULL;

	return psc->GetMergeCell( hWndCol, lRow );
}


// MTblGetMergeRow
// Liefert Merge-Zeile
long WINAPI MTblGetMergeRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return TBL_Error;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return TBL_Error;

	return psc->GetMergeRow( lRow );
}


// MTblGetMsgOffs
// Liefert Message-Offset
UINT WINAPI MTblGetMsgOffset( )
{
	return guiMsgOffs;
}


// MTblGetMWheelOption
// Liefert den Wert einer Mausrad-Scroll-Option
int WINAPI MTblGetMWheelOption( HWND hWndTbl, int iOption )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	int iRet = -1;

	// Wert ermitteln
	switch ( iOption )
	{
	case MTMW_ROWS_PER_NOTCH:
		iRet = psc->m_iMWRowsPerNotch;
		break;
	case MTMW_SPLITROWS_PER_NOTCH:
		iRet = psc->m_iMWSplitRowsPerNotch;
		break;
	case MTMW_PAGE_KEY:
		iRet = psc->m_iMWPageKey;
		break;
	}

	return iRet;
}


// MTblGetNextChildRow
// Liefert die nächste Kindzeile
long WINAPI MTblGetNextChildRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	CMTblRow * pRow, * pChildRow;

	// Zeiger auf Zeile ermitteln
	pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( !pRow ) return TBL_Error;

	// Nächste Kindzeile ermitteln
	BOOL bBottomUp = ( lRow >= 0 ? psc->m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP : psc->m_dwTreeFlags & MTBL_TREE_FLAG_SPLIT_BOTTOM_UP );
	pChildRow = psc->m_Rows->GetNextChildRow( pRow, bBottomUp );
	if ( !pChildRow ) return TBL_Error;

	return pChildRow->GetNr( );
}


// MTblGetNodeImages
BOOL WINAPI MTblGetNodeImages( HWND hWndTbl, LPHANDLE phImg, LPHANDLE phImgExp )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Bilder ermitteln
	*phImg = psc->m_hImgNode;
	*phImgExp = psc->m_hImgNodeExp;

	return TRUE;
}


// MTblGetNodeRect
// Liefert das Knotenrechteck
BOOL WINAPI MTblGetNodeRect( HWND hWndTbl, long lCellLeft, long lTextTop, int iFontHeight, int iLevel, int iCellLeadingX, int iCellLeadingY, LPRECT pr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->GetNodeRect( lCellLeft, lTextTop, iFontHeight, iLevel, iCellLeadingX, iCellLeadingY, pr );
}


// MTblGetOrigParentRow
// Liefert die ursprüngliche Elternzeile einer Zeile
long WINAPI MTblGetOrigParentRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	return psc->m_Rows->GetOrigParentRow( lRow );
}


// MTblGetOwnerDrawnItem
// Liefert Daten zum ODI
BOOL WINAPI MTblGetOwnerDrawnItem( HWND hWndTbl, LPARAM lID, HUDV hUdv )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// ODI-Struktur ermitteln
	LPMTBLODI pODI = NULL;
	if ( lID == (LPARAM)psc->m_pODI )
		pODI = psc->m_pODI;
	else if ( lID == (LPARAM)psc->m_pPrODI )
		pODI = psc->m_pPrODI;
	if ( !pODI )
		return FALSE;

	// UDV dereferenzieren
	LPMTBLODIUDV pODIUDV = (LPMTBLODIUDV)Ctd.UdvDeref( hUdv );
	if ( !pODIUDV )
		return FALSE;

	// Werte setzen
	Ctd.CvtLongToNumber( pODI->lType, &pODIUDV->Type );
	pODIUDV->HWndParam = pODI->hWndParam;
	Ctd.CvtLongToNumber( pODI->lParam, &pODIUDV->NParam );
	pODIUDV->HWndPart = pODI->hWndPart;
	Ctd.CvtLongToNumber( pODI->lPart, &pODIUDV->NPart );
	Ctd.CvtULongToNumber( (ULONG)pODI->hDC, &pODIUDV->DC );
	Ctd.CvtLongToNumber( pODI->r.left, &pODIUDV->Left );
	Ctd.CvtLongToNumber( pODI->r.top, &pODIUDV->Top );
	Ctd.CvtLongToNumber( pODI->r.right, &pODIUDV->Right );
	Ctd.CvtLongToNumber( pODI->r.bottom, &pODIUDV->Bottom );
	Ctd.CvtLongToNumber( pODI->lSelected, &pODIUDV->Selected );
	Ctd.CvtLongToNumber( pODI->lXLeading, &pODIUDV->XLeading );
	Ctd.CvtLongToNumber( pODI->lYLeading, &pODIUDV->YLeading );
	Ctd.CvtLongToNumber( pODI->lPrinting, &pODIUDV->Printing );
	Ctd.CvtDoubleToNumber( pODI->dYResFactor, &pODIUDV->YResFactor );
	Ctd.CvtDoubleToNumber( pODI->dXResFactor, &pODIUDV->XResFactor );

	return TRUE;
}


// MTblGetParentRow
// Liefert die Elternzeile einer Zeile
long WINAPI MTblGetParentRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	return psc->m_Rows->GetParentRow( lRow );
}


// MTblGetPasswordChar
// Liefert das akt. Passwort-Zeichen
HSTRING WINAPI MTblGetPasswordChar( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return 0;

	HSTRING hsPwd = 0;
	if ( !Ctd.InitLPHSTRINGParam( &hsPwd, Ctd.BufLenFromStrLen( 1 ) ) ) return 0;
	long lLen;
	LPTSTR lpsBuf = Ctd.StringGetBuffer( hsPwd, &lLen );
	lpsBuf[0] = psc->m_cPassword;
	lpsBuf[1] = 0;
	Ctd.HStrUnlock(hsPwd);

	return hsPwd;
}

// MTblGetPasswordCharVal
// Liefert das akt. Passwort-Zeichen als CHAR-Wert
TCHAR WINAPI MTblGetPasswordCharVal( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return 0;

	return psc->m_cPassword;
}


// MTblGetPatchLevel
// Liefert den Patch-Level

long WINAPI MTblGetPatchLevel( )
{
	return MTBL_PATCH_LEVEL;
}

// MTblGetPrevChildRow
// Liefert die vorherige Kindzeile
long WINAPI MTblGetPrevChildRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	CMTblRow * pRow, * pChildRow;

	// Zeiger auf Zeile ermitteln
	pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( !pRow ) return TBL_Error;

	// Nächste Kindzeile ermitteln
	BOOL bBottomUp = ( lRow >= 0 ? psc->m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP : psc->m_dwTreeFlags & MTBL_TREE_FLAG_SPLIT_BOTTOM_UP );
	pChildRow = psc->m_Rows->GetPrevChildRow( pRow, bBottomUp );
	if ( !pChildRow ) return TBL_Error;

	return pChildRow->GetNr( );
}


// MTblGetRowBackColor
// Liefert die Hintergrundfarbe einer Zeile
BOOL WINAPI MTblGetRowBackColor( HWND hWndTbl, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;
	
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
		dwColor = DWORD( pRow->m_BackColor->Get( ) );

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetRowFlagImage
// Liefert das Bild eines Row-Flags 
BOOL WINAPI MTblGetRowFlagImage( HWND hWndTbl, WORD wRowFlag, LPHANDLE phImage )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*phImage = psc->GetRowFlagImg( wRowFlag, TRUE );
	return TRUE;
}

// MTblGetRowFont
// Liefert den Font einer Zeile
BOOL WINAPI MTblGetRowFont( HWND hWndTbl, long lRow, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Font ermitteln
	CMTblFont f;
	CMTblFont *pf = &f;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
		pf = pRow->m_Font;

	// Font zurückgeben
	return pf->Get( phsName, plSize, plEnh );
}


// MTblGetRowFromID
// Liefert die Zeilennummer anhand der RowID ( = Zeiger auf interne Zeile )
long WINAPI MTblGetRowFromID( HWND hWndTbl, LPVOID lpRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return NULL;

	CMTblRow *pRow = (CMTblRow*)lpRow;
	if ( !psc->m_Rows->IsValidRowPtr( pRow ) ) return TBL_Error;
	if ( pRow->IsDelAdj( ) ) return TBL_Error;

	return pRow->GetNr( );
}


// MTblGetRowHdrBackColor
// Liefert die Header-Hintergrundfarbe einer Zeile
BOOL WINAPI MTblGetRowHdrBackColor( HWND hWndTbl, long lRow, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	DWORD dwColor = MTBL_COLOR_UNDEF;

	if ( lRow == TBL_Error )
	{
		// Farbe = Farbe freier Bereich
		dwColor = psc->m_dwRowHdrFreeBackColor;
	}
	else
	{
		// Gültige Zeile?
		if ( !psc->IsRow( lRow ) ) return FALSE;

		// Farbe ermitteln
		CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
		if ( pRow )
			dwColor = DWORD( pRow->m_HdrBackColor->Get( ) );
	}

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetRowHeight
// Liefert die eingestellte Höhe einer Zeile
BOOL WINAPI MTblGetRowHeight( HWND hWndTbl, long lRow, LPLONG plHeight )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Höhe ermitteln
	long lHeight = psc->m_Rows->GetRowHeight( lRow );
	if ( lHeight < 0 )
		return FALSE;

	// Höhe setzen + return
	*plHeight = lHeight;
	return TRUE;
}


// MTblGetRowID
// Liefert die RowID ( = Zeiger auf interne Zeile )
LPVOID WINAPI MTblGetRowID( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return NULL;

	//return psc->m_Rows->GetRowPtr( lRow );
	if ( !psc->IsRow( lRow ) ) return 0;

	return psc->m_Rows->EnsureValid( lRow );
}


// MTblGetRowMerge
// Liefert Zeilenverknüpfung
BOOL WINAPI MTblGetRowMerge( HWND hWndTbl, long lRow, LPINT piCount )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Ermitteln
	int iCount = 0;

	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
		iCount = pRow->m_Merged;

	// In Parameter übertragen
	*piCount = iCount;

	return TRUE;
}


// MTblGetRowLevel
// Liefert die Ebene einer Zeile
int WINAPI MTblGetRowLevel( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->m_Rows->GetRowLevel( lRow );
}


// MTblGetRowSelChanges
// Liefert die Änderungen der Zeilenselektion
long WINAPI MTblGetRowSelChanges( HWND hWndTbl, HARRAY haRows, HARRAY haChanges )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return -1;

	return psc->GetRowSelChanges( haRows, haChanges );
}


// MTblGetRowTextColor
// Liefert die Textfarbe einer Zeile
BOOL WINAPI MTblGetRowTextColor( HWND hWndTbl, long lRow, LPDWORD pdwColor )
{

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Farbe ermitteln
	DWORD dwColor = MTBL_COLOR_UNDEF;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
		dwColor = DWORD( pRow->m_TextColor->Get( ) );

	// In Parameter übertragen
	*pdwColor = dwColor;

	return TRUE;
}


// MTblGetRowUserData
// Liefert Benutzerdaten einer Zeile
HSTRING WINAPI MTblGetRowUserData( HWND hWndTbl, long lRow, LPTSTR lpsKey )
{
	// CTD da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return 0;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return 0;

	// Daten ermitteln
	tstring sKey( lpsKey );
	tstring sValue( psc->m_Rows->GetRowUserData( lRow, sKey ) );

	// Return
	BOOL bOk = TRUE;
	HSTRING hs = 0;
	Ctd.InitLPHSTRINGParam( &hs, Ctd.BufLenFromStrLen( sValue.size( ) ) );
	long lLen;
	LPTSTR lps = Ctd.StringGetBuffer( hs, &lLen );
	if (lps)
		lstrcpy(lps, sValue.c_str());
	else
		bOk = FALSE;
	Ctd.HStrUnlock(hs);

	return bOk ? hs : 0;
}


// MTblGetRowVisPos
// Liefert die sichtbare Position einer Zeile
long WINAPI MTblGetRowVisPos( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) )
		return TBL_Error;

	// Ermitteln
	return psc->m_Rows->GetRowVisPos( lRow );
}


// MTblGetRowYCoord
// Liefert die Y-Koordinaten ( oben, unten ) einer Zeile
int MTblGetRowYCoord( HWND hWndTbl, long lRow, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis, BOOL bInVisRangeOnly /*= TRUE*/ )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return MTGR_RET_ERROR;

	return psc->GetRowYCoord( lRow, NULL, ppt, pptVis, bInVisRangeOnly );
}


// MTblGetSelectionColors
// Liefert die Selektionsfarben
BOOL WINAPI MTblGetSelectionColors( HWND hWndTbl, LPDWORD pdwTextColor, LPDWORD pdwBackColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*pdwTextColor = psc->m_dwSelTextColor;
	*pdwBackColor = psc->m_dwSelBackColor;

	return TRUE;
}


// MTblGetSelectionDarkening
// Liefert die Selektionsverdunkelung
BOOL WINAPI MTblGetSelectionDarkening(HWND hWndTbl, LPBYTE pbySelDarkening)
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass(hWndTbl);
	if (!psc) return FALSE;

	*pbySelDarkening = psc->m_bySelDarkening;

	return TRUE;
}


// MTblGetSepCol
// Liefert das Spalten-Window-Handle, wenn die übergebenen Koordinaten innerhalb
// des Trennbereiches eines Spalten-Kopfes liegen, ansonsten NULL
HWND WINAPI MTblGetSepCol( HWND hWndTbl, int iX, int iY )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return NULL;

	return psc->GetSepCol( iX, iY );
}


// MTblGetSepRow
// Liefert die Zeilennummer, wenn die übergebenen Koordinaten innerhalb
// des Trennbereiches eines Zeilen-Kopfes liegen, ansonsten TBL_Error
long WINAPI MTblGetSepRow( HWND hWndTbl, int iX, int iY )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	POINT pt = {iX,iY};
	return psc->GetSepRow( pt );
}


// MTblGetSubClass
// Liefert Zeiger auf Subclass-Struktur
inline LPMTBLSUBCLASS MTblGetSubClass( HWND hWndTbl )
{
	SubClassMap::iterator it = gscm.find( hWndTbl );
	if ( it != gscm.end( ) )
	{
		return (LPMTBLSUBCLASS)(*it).second;
	}

	return NULL;
}


// MTblGetSubClassFromCol
// Liefert Zeiger auf Subclass-Struktur anhand des Spalten-Handles mit Anpassung Window-Handles
LPMTBLSUBCLASS MTblGetSubClassFromCol( HWND hWndCol )
{
	// CTD da?
	if ( !Ctd.IsPresent( ) ) return NULL;

	// Tabelle ermitteln
	HWND hWndTbl = Ctd.ParentWindow( hWndCol );

	// Subclass-Struktur ermitteln
	if ( IsWindow( hWndTbl ) )
		return MTblGetSubClass( hWndTbl );
	else
		return NULL;
}


// MTblGetSubClassFromItem
// Liefert Zeiger auf Subclass-Struktur für ein Item
LPMTBLSUBCLASS MTblGetSubClassFromItem( LPMTBLITEM pItem )
{
	if ( !pItem ) return NULL;

	LPMTBLSUBCLASS psc = NULL;

	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
	case MTBL_ITEM_COLUMN:
	case MTBL_ITEM_COLHDR:
		psc = MTblGetSubClassFromCol( pItem->WndHandle );
		break;
	default:
		psc = MTblGetSubClass( pItem->WndHandle );
	}

	return psc;
}


// MTblGetSubClass_GetTip
// Liefert Zeiger auf Subclass-Struktur für MTblGetTip* Funktionen
LPMTBLSUBCLASS MTblGetSubClass_GetTip( HWND hWndParam, long lParam, int iType )
{
	// Tip-Typ ermitteln
	BOOL bDefault = ( iType == MTBL_TIP_DEFAULT );
	BOOL bCell = ( iType == MTBL_TIP_CELL ||
		          iType == MTBL_TIP_CELL_TEXT ||
				  iType == MTBL_TIP_CELL_COMPLETETEXT ||
				  iType == MTBL_TIP_CELL_IMAGE ||
				  iType == MTBL_TIP_CELL_BTN );
	BOOL bColHdr = ( iType == MTBL_TIP_COLHDR );
	BOOL bColHdrGrp = ( iType == MTBL_TIP_COLHDRGRP );
	BOOL bCorner = ( iType == MTBL_TIP_CORNER );
	BOOL bRowHdr = ( iType == MTBL_TIP_ROWHDR );

	if ( !( bDefault || bCell || bColHdr || bColHdrGrp || bCorner || bRowHdr ) )
		return NULL;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = NULL;
	HWND hWndTbl = NULL;
	if ( bDefault || bColHdrGrp || bCorner || bRowHdr )
		psc = MTblGetSubClass( hWndParam );
	else
		psc = MTblGetSubClassFromCol( hWndParam );
	if ( !psc ) return NULL;

	// Gültige Zeile?
	if ( bCell || bRowHdr )
		if ( !psc->IsRow( lParam ) ) return NULL;

	return psc;
}


// MTblGetSubClass_SetTip
// Liefert Zeiger auf Subclass-Struktur für MTblSetTip* Funktionen
LPMTBLSUBCLASS MTblGetSubClass_SetTip( HWND hWndParam, long lParam, int iType )
{
	// Tip-Typ ermitteln
	BOOL bDefault = ( iType == MTBL_TIP_DEFAULT );
	BOOL bCell = ( iType == MTBL_TIP_CELL ||
		          iType == MTBL_TIP_CELL_TEXT ||
				  iType == MTBL_TIP_CELL_COMPLETETEXT ||
				  iType == MTBL_TIP_CELL_IMAGE ||
				  iType == MTBL_TIP_CELL_BTN );
	BOOL bColHdr = ( iType == MTBL_TIP_COLHDR || 
		             iType == MTBL_TIP_COLHDR_BTN );
	BOOL bColHdrGrp = ( iType == MTBL_TIP_COLHDRGRP );
	BOOL bCorner = ( iType == MTBL_TIP_CORNER );
	BOOL bRowHdr = ( iType == MTBL_TIP_ROWHDR );

	if ( !( bDefault || bCell || bColHdr || bColHdrGrp || bCorner || bRowHdr ) )
		return NULL;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = NULL;
	HWND hWndTbl = NULL;
	if ( bDefault || bColHdrGrp || bCorner || bRowHdr )
		psc = MTblGetSubClass( hWndParam );
	else
		psc = MTblGetSubClassFromCol( hWndParam );
	if ( !psc ) return NULL;

	if ( bCell || bRowHdr )
	{
		// Gültige Zeile?
		if ( !psc->IsRow( lParam ) ) return NULL;

		// Zeile mit TBL_Adjust gelöscht?
		if ( psc->m_Rows->IsRowDelAdj( lParam ) ) return NULL;	
	}

	return psc;
}


// MTblGetTipBackColor
// Liefert Tip-Hintergrundfarbe
BOOL WINAPI MTblGetTipBackColor( HWND hWndParam, long lParam, int iType, LPDWORD pdwColor )
{
	// Parameter checken
	if ( !pdwColor )
		return FALSE;
	
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Farbe ermitteln
	*pdwColor = pTip->GetBackColor( TRUE );

	return TRUE;
}


// MTblGetTipDelay
// Liefert Tip-Anzeigeverzögerung
BOOL WINAPI MTblGetTipDelay( HWND hWndParam, long lParam, int iType, LPNUMBER pnDelay )
{
	// Parameter checken
	if ( !pnDelay )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Delay ermitteln
	if ( pTip->m_iDelay == INT_MIN )
		pnDelay->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iDelay, pnDelay );

	return TRUE;
}


// MTblGetTipFadeInTime
// Liefert Tip-Einblendzeit
BOOL WINAPI MTblGetTipFadeInTime( HWND hWndParam, long lParam, int iType, LPNUMBER pnTime )
{
	// Parameter checken
	if ( !( pnTime ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Position ermitteln
	if ( pTip->m_iFadeInTime == INT_MIN )
		pnTime->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iFadeInTime, pnTime );

	return TRUE;
}


// MTblGetTipFont
// Liefert Tip-Font
BOOL WINAPI MTblGetTipFont( HWND hWndParam, long lParam, int iType, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh )
{
	// Parameter checken
	if ( !( phsName && plSize && plEnh ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Font ermitteln
	return pTip->m_Font.Get( phsName, plSize, plEnh );
}


// MTblGetTipFrameColor
// Liefert Tip-Rahmenfarbe
BOOL WINAPI MTblGetTipFrameColor( HWND hWndParam, long lParam, int iType, LPDWORD pdwColor )
{
	// Parameter checken
	if ( !pdwColor )
		return FALSE;
	
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Farbe ermitteln
	*pdwColor = pTip->GetFrameColor( TRUE );

	return TRUE;
}



// MTblGetTipOffset
// Liefert Tip-Maus-Offsets
BOOL WINAPI MTblGetTipOffset( HWND hWndParam, long lParam, int iType, LPNUMBER pnOffsX, LPNUMBER pnOffsY )
{
	// Parameter checken
	if ( !( pnOffsX && pnOffsY ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Position ermitteln
	if ( pTip->m_iOffsX == INT_MIN )
		pnOffsX->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iOffsX, pnOffsX );

	if ( pTip->m_iOffsY == INT_MIN )
		pnOffsY->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iOffsY, pnOffsY );

	return TRUE;
}


// MTblGetTipOpacity
// Liefert Tip-Opazität
BOOL WINAPI MTblGetTipOpacity( HWND hWndParam, long lParam, int iType, LPNUMBER pnOpacity )
{
	// Parameter checken
	if ( !( pnOpacity ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Position ermitteln
	if ( pTip->m_iOpacity == INT_MIN )
		pnOpacity->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iOpacity, pnOpacity );

	return TRUE;
}


// MTblGetTipPos
// Liefert Tip-Position
BOOL WINAPI MTblGetTipPos( HWND hWndParam, long lParam, int iType, LPNUMBER pnX, LPNUMBER pnY )
{
	// Parameter checken
	if ( !( pnX && pnY ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Position ermitteln
	if ( pTip->m_iPosX == INT_MIN )
		pnX->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iPosX, pnX );

	if ( pTip->m_iPosY == INT_MIN )
		pnY->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iPosY, pnY );

	return TRUE;
}


// MTblGetTipRCESize
// Liefert Tip-RCE-Größe ( RCE = RoundCornerEllipse )
BOOL WINAPI MTblGetTipRCESize( HWND hWndParam, long lParam, int iType, LPNUMBER pnWid, LPNUMBER pnHt )
{
	// Parameter checken
	if ( !( pnWid && pnHt ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Position ermitteln
	if ( pTip->m_iRCEWid == INT_MIN )
		pnWid->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iRCEWid, pnWid );

	if ( pTip->m_iRCEHt == INT_MIN )
		pnHt->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iRCEHt, pnHt );

	return TRUE;
}


// MTblGetTipSize
// Liefert Tip-Größe
BOOL WINAPI MTblGetTipSize( HWND hWndParam, long lParam, int iType, LPNUMBER pnWid, LPNUMBER pnHt )
{
	// Parameter checken
	if ( !( pnWid && pnHt ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Größe ermitteln
	if ( pTip->m_iWid == INT_MIN )
		pnWid->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iWid, pnWid );

	if ( pTip->m_iHt == INT_MIN )
		pnHt->numLength = 0;
	else
		Ctd.CvtIntToNumber( pTip->m_iHt, pnHt );

	return TRUE;
}


// MTblGetTipText
// Liefert Tip-Text
BOOL WINAPI MTblGetTipText( HWND hWndParam, long lParam, int iType, LPHSTRING phsText )
{
	// Parameter checken
	if ( !phsText )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Text ermitteln
	BOOL bOk = TRUE;
	long lLen = Ctd.BufLenFromStrLen( pTip->m_sText.size( ) );
	if ( !Ctd.InitLPHSTRINGParam( phsText, lLen ) )
		return FALSE;
	LPTSTR lpsBuf = Ctd.StringGetBuffer( *phsText, &lLen );
	if (lpsBuf)
		lstrcpy(lpsBuf, pTip->m_sText.c_str());
	else
		bOk = FALSE;
	Ctd.HStrUnlock(*phsText);

	return bOk;
}


// MTblGetTipTextColor
// Liefert Tip-Textfarbe
BOOL WINAPI MTblGetTipTextColor( HWND hWndParam, long lParam, int iType, LPDWORD pdwColor )
{
	// Parameter checken
	if ( !pdwColor )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	// Farbe ermitteln
	*pdwColor = pTip->GetTextColor( TRUE );

	return TRUE;
}


// MTblGetTreeNodeColors
BOOL WINAPI MTblGetTreeNodeColors( HWND hWndTbl, LPDWORD pdwOuterClr, LPDWORD pdwInnerClr, LPDWORD pdwBackClr )
{
	// Parameter checken
	if ( !( pdwOuterClr && pdwInnerClr && pdwBackClr ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Setzen
	*pdwOuterClr = psc->m_dwNodeColor;
	*pdwInnerClr = psc->m_dwNodeInnerColor;
	*pdwBackClr = psc->m_dwNodeBackColor;

	return TRUE;
}

/*
// MTblGetVersion
// Liefert die Version
HSTRING WINAPI MTblGetVersion( )
{
	HSTRING hsVer = 0;
	LPTSTR pBuf;
	long lLen;

	Ctd.InitLPHSTRINGParam( &hsVer, Ctd.BufLenFromStrLen( lstrlen( MTBL_VERSION ) ) );
	pBuf = Ctd.StringGetBuffer( hsVer, &lLen );
	lstrcpy( pBuf, MTBL_VERSION );
	return hsVer;
}
*/

// MTblGetVersion
// Liefert die Version
HSTRING WINAPI MTblGetVersion()
{
	TCHAR szFilename[MAX_PATH + 1] = { 0 };
	TCHAR szVersion[100] = { 0 };

	if (GetModuleFileName((HMODULE)ghInstance, szFilename, MAX_PATH))
	{
		DWORD dwTemp;
		DWORD dwSize = GetFileVersionInfoSize(szFilename, &dwTemp);
		if (dwSize > 0)
		{
			std::vector<BYTE> data(dwSize);
			if (GetFileVersionInfo(szFilename, NULL, dwSize, &data[0]))
			{
				UINT uiVerLen = 0;
				VS_FIXEDFILEINFO* pFixedInfo = 0;
				if (VerQueryValue(&data[0], TEXT("\\"), (void**)&pFixedInfo, (UINT *)&uiVerLen))
				{
					//_stprintf(szVersion, _T("%u.%u.%u.%u"), HIWORD(pFixedInfo->dwProductVersionMS), LOWORD(pFixedInfo->dwProductVersionMS), HIWORD(pFixedInfo->dwProductVersionLS), LOWORD(pFixedInfo->dwProductVersionLS));
					_stprintf(szVersion, _T("%u.%u.%u"), HIWORD(pFixedInfo->dwProductVersionMS), LOWORD(pFixedInfo->dwProductVersionMS), HIWORD(pFixedInfo->dwProductVersionLS));
				}
			}
		}
	}

	HSTRING hsVer = 0;
	LPTSTR pBuf;
	long lLen;

	Ctd.InitLPHSTRINGParam(&hsVer, Ctd.BufLenFromStrLen(lstrlen(szVersion)));
	pBuf = Ctd.StringGetBuffer(hsVer, &lLen);
	lstrcpy(pBuf, szVersion);
	Ctd.HStrUnlock(hsVer);
	return hsVer;
}


// MTblInsertChildRow
// Fügt eine Kindzeile ein
long WINAPI MTblInsertChildRow( HWND hWndTbl, LPLONG plParentRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	// Gültige Eltern-Zeile?
	if ( !plParentRow ) return TBL_Error;
	if ( !psc->IsRow( *plParentRow ) ) return TBL_Error;	
	CMTblRow *pParentRow = psc->m_Rows->EnsureValid( *plParentRow );
	if ( !pParentRow ) return TBL_Error;

	// Ist Elternzeile Teil eines Zeilenverbundes?
	LPMTBLMERGEROW pmr = psc->FindMergeRow( *plParentRow, FMR_ANY );

	// Split-Zeile?
	BOOL bSplitRow = ( *plParentRow < 0 );

	// Umgekehrter Baum?
	BOOL bBottomUp = FALSE;
	if ( *plParentRow < 0 )
		bBottomUp = psc->QueryTreeFlags( MTBL_TREE_FLAG_SPLIT_BOTTOM_UP );
	else
		bBottomUp = psc->QueryTreeFlags( MTBL_TREE_FLAG_BOTTOM_UP );

	// Einzufügende Zeile ermitteln
	long lInsRow = TBL_Error, lRow;

	if ( dwFlags & MTICR_FIRST )
	{
		if ( bBottomUp )
		{
			//lInsRow = *plParentRow;
			lInsRow = ( pmr ? pmr->lRowFrom : *plParentRow );
		}
		else
		{
			//lRow = *plParentRow;
			lRow = ( pmr ? pmr->lRowTo : *plParentRow );
			if ( Ctd.TblFindNextRow( hWndTbl, &lRow, 0, 0 ) && ( bSplitRow ? lRow < 0 : TRUE ) )
				lInsRow = lRow;
			else
				lInsRow = ( bSplitRow ? -1 : TBL_MaxRow );
		}
	}
	else
	{
		CMTblRow *pLastDescRow = NULL;
		if ( pParentRow->GetChildCount( ) > 0 )
			pLastDescRow = psc->m_Rows->GetLastDescendantRow( pParentRow, bBottomUp );

		if ( bBottomUp )
		{
			//lInsRow = ( pLastDescRow ? pLastDescRow->GetNr( ) : *plParentRow );
			lInsRow = ( pLastDescRow ? pLastDescRow->GetNr( ) : ( pmr ? pmr->lRowFrom : *plParentRow ) );
		}
		else
		{
			lRow = ( pLastDescRow ? pLastDescRow->GetNr( ) : ( pmr ? pmr->lRowTo : *plParentRow ) );
			if ( Ctd.TblFindNextRow( hWndTbl, &lRow, 0, 0 ) && ( bSplitRow ? lRow < 0 : TRUE ) )
				lInsRow = lRow;
			else
				lInsRow = ( bSplitRow ? -1 : TBL_MaxRow );
		}
	}

	lInsRow = Ctd.TblInsertRow( hWndTbl, lInsRow );

	if ( lInsRow != TBL_Error )
	{
		*plParentRow = pParentRow->GetNr( );
		if ( !MTblSetParentRow( hWndTbl, lInsRow, *plParentRow, ( dwFlags & MTICR_REDRAW ) ? MTSPR_REDRAW : 0 ) )
			*plParentRow = TBL_Error;
	}

	return lInsRow;
}


// MTblIsCellDisabled
// Ermittelt, ob eine Zelle gesperrt ist
BOOL WINAPI MTblIsCellDisabled( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	return psc->m_Rows->IsRowDisCol( lRow, hWndCol );
}


// MTblIsExtMsgsEnabled
// Ermittelt, ob die Generierung von erweiterten Nachrichten eingeschaltet ist
BOOL WINAPI MTblIsExtMsgsEnabled( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_bExtMsgs;
}


// MTblIsListOpen
// Ermittelt, ob die Drop Down Liste offen ist
BOOL MTblIsListOpen( HWND hWndTbl, HWND hWndCol )
{
	BOOL bListOpen = FALSE;

	HWND hWndList = NULL;
	int iCellType = Ctd.TblGetCellType( hWndCol );
	if ( iCellType == COL_CellType_DropDownList )
		hWndList = (HWND)SendMessage( hWndTbl, TIM_GetExtEdit, Ctd.TblQueryColumnID( hWndCol ) - 1, 0 );
	if ( hWndList )
	{
		HWND hWndListParent = GetParent( hWndList );
		if ( hWndListParent )
		{
			TCHAR cClassName[255];
			if ( GetClassName( hWndListParent, cClassName, 254 ) )
			{
				bListOpen = ( lstrcmp( cClassName, CLSNAME_DROPDOWN ) == 0 );
			}
		}
	}

	return bListOpen;
}


// MTblIsHighlighted
// Ermittelt, ob ein bestimmtes Item hervorgehoben ist
BOOL WINAPI MTblIsHighlighted( HUDV hUdvItem )
{
	CMTblItem item;
	item.FromUdv( hUdvItem );

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem( &item );
	if ( !psc ) return FALSE;

	// Ermitteln
	ItemList::iterator it = psc->FindHighlightedItem( item );
	return it != psc->m_pHLIL->end( );
}


// MTblIsMWheelScrollEnabled
// Ermittelt, ob das Scrollen bei WM_MOUSEWHEEL eingeschaltet ist
BOOL WINAPI MTblIsMWheelScrollEnabled( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_bMWheelScroll;
}


// MTblIsOwnerDrawnItem
// Ermittelt, ob ein Item benutzerdef. gezeichnet ist
BOOL WINAPI MTblIsOwnerDrawnItem( HWND hWndParam, long lParam, int iType )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = NULL;

	switch ( iType )
	{
	case MTBL_ODI_CELL:
	case MTBL_ODI_COLUMN:
	case MTBL_ODI_COLHDR:
		psc = MTblGetSubClassFromCol( hWndParam );
		break;
	case MTBL_ODI_COLHDRGRP:
	case MTBL_ODI_ROW:
	case MTBL_ODI_ROWHDR:
	case MTBL_ODI_CORNER:
		psc = MTblGetSubClass( hWndParam );
		break;
	default:
		return FALSE;
	}
	if ( !psc ) return FALSE;

	// Ermitteln
	return psc->IsOwnerDrawnItem( hWndParam, lParam, iType );
}

// MTblIsParentRow
// Ermittelt, ob eine Zeile Kindzeilen hat
BOOL WINAPI MTblIsParentRow( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_Rows->IsParentRow( lRow );
}


// MTblIsRow
// Ermittelt, ob eine Zeile gültig ist
BOOL WINAPI MTblIsRow( HWND hWndTbl, long lRow )
{
	if ( !Ctd.IsPresent( ) ) return FALSE;

	if ( lRow == TBL_Error ) return FALSE;

	long lContext = Ctd.TblQueryContext( hWndTbl );
	BOOL bOk = Ctd.TblSetContextOnly( hWndTbl, lRow );
	Ctd.TblSetContextOnly( hWndTbl, lContext );

	return bOk;
}


// MTblIsRowDescOf
// Ermittelt, ob eine Zeile Nachkomme einer Elternzeile ist
BOOL WINAPI MTblIsRowDescOf( HWND hWndTbl, long lRow, long lParentRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_Rows->IsRowDescendantOf( lRow, lParentRow );
}


// MTblIsRowDisabled
// Ermittelt, ob eine Zeile gesperrt ist
BOOL WINAPI MTblIsRowDisabled( HWND hWndTbl, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_Rows->IsRowDisabled( lRow );
}


// MTblIsSubClassed
// Ermittelt, ob das Subclassing für eine Tabelle bereits durchgeführt wurde
BOOL WINAPI MTblIsSubClassed( HWND hWndTbl )
{
	return ( MTblGetSubClass( hWndTbl ) != NULL );
}


// MTblIsTipTypeEnabled
// Ermittelt, ob ein Tip-Typ aktiviert ist
BOOL WINAPI MTblIsTipTypeEnabled( HWND hWndTbl, DWORD dwTipType )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->IsTipTypeEnabled( dwTipType );
}


// MTblLockPaint
// Blockiert das Neuzeichnen der Tabelle
long WINAPI MTblLockPaint( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	return psc->LockPaint( );
}

// MTblMustRedrawSelections
// Ermittelt, ob Selektionen generell neu gezeichnet werden müssen
BOOL WINAPI MTblMustRedrawSelections( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->MustRedrawSelections( );
}


// MTblObjFromPoint
// Liefert Informationen zu Objekten an der angegebenen Position
BOOL WINAPI MTblObjFromPoint( HWND hWndTbl, int iX, int iY, LPLONG plRow, LPHWND phWndCol, LPDWORD pdwFlags, LPDWORD pdwMFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	POINT pt;
	pt.x = iX;
	pt.y = iY;

	return psc->ObjFromPoint( pt, plRow, phWndCol, pdwFlags, pdwMFlags );
}


// MTblODIPrint
// Druck eines OwnerDrawnItem
BOOL WINAPI MTblODIPrint( HWND hWndTbl, LPVOID pODI )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Parameter-Check
	if ( !pODI )
		return FALSE;

	psc->ODIPrint( (LPMTBLODI)pODI );

	return TRUE;
}

// MTblPropStrToBool
BOOL MTblPropStrToBool( HSTRING hs )
{
	BOOL bRet = FALSE;

	LPTSTR lps = Ctd.HStrToLPTSTR( hs );
	if ( lps )
	{
		if ( lstrcmp( lps, MTPROP_VAL_BOOL_YES ) == 0 )
			bRet = TRUE;
	}

	return bRet;
}


// MTblQueryButtonCell
// Liefert Button-Definition einer Zelle
BOOL WINAPI MTblQueryButtonCell( HWND hWndCol, long lRow, LPBOOL pbShowBtn, LPINT piWidth, LPHSTRING phsText, LPHANDLE phImage, LPDWORD pdwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Button ermitteln
	CMTblBtn *pBtn = NULL;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pc = pRow->m_Cols->Find( hWndCol );
		if ( pc )
			pBtn = &pc->m_Btn;

	}

	BOOL bOk = TRUE;
	long lLen;
	LPTSTR pBuf = NULL;

	if( pBtn)
	{
		*pbShowBtn = pBtn->m_bShow;
		*piWidth = pBtn->m_iWidth;
		Ctd.InitLPHSTRINGParam( phsText, Ctd.BufLenFromStrLen( pBtn->m_sText.size( ) ) );
		pBuf = Ctd.StringGetBuffer( *phsText, &lLen );
		if (pBuf)
			lstrcpy(pBuf, pBtn->m_sText.c_str());
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsText);
		
		*phImage = pBtn->m_hImage;
		*pdwFlags = pBtn->m_dwFlags;
	}
	else
	{
		*pbShowBtn = FALSE;
		*piWidth = 0;
		Ctd.InitLPHSTRINGParam( phsText, Ctd.BufLenFromStrLen( 0 ) );
		pBuf = Ctd.StringGetBuffer( *phsText, &lLen );
		if (pBuf)
			lstrcpy(pBuf, _T(""));
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsText);
		
		*phImage = NULL;
		*pdwFlags = 0;

	}

	return bOk;
}


// MTblQueryButtonColumn
// Liefert Button-Definition einer Spalte
BOOL WINAPI MTblQueryButtonColumn( HWND hWndCol, LPBOOL pbShowBtn, LPINT piWidth, LPHSTRING phsText, LPHANDLE phImage, LPDWORD pdwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Ermitteln
	CMTblCol * pc = psc->m_Cols->Find( hWndCol );

	BOOL bOk = TRUE;
	long lLen;
	LPTSTR pBuf = NULL;

	if( pc )
	{
		*pbShowBtn = pc->m_Btn.m_bShow;
		*piWidth = pc->m_Btn.m_iWidth;
		Ctd.InitLPHSTRINGParam( phsText, Ctd.BufLenFromStrLen( pc->m_Btn.m_sText.size( ) ) );
		pBuf = Ctd.StringGetBuffer( *phsText, &lLen );
		if (pBuf)
			lstrcpy(pBuf, pc->m_Btn.m_sText.c_str());
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsText);
		
		*phImage = pc->m_Btn.m_hImage;
		*pdwFlags = pc->m_Btn.m_dwFlags;
	}
	else
	{
		*pbShowBtn = FALSE;
		*piWidth = 0;
		Ctd.InitLPHSTRINGParam( phsText, Ctd.BufLenFromStrLen( 0 ) );
		pBuf = Ctd.StringGetBuffer( *phsText, &lLen );
		if (pBuf)
			lstrcpy(pBuf, _T(""));
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsText);
		
		*phImage = NULL;
		*pdwFlags = 0;

	}

	return bOk;
}


// MTblQueryButtonHotkey
// Liefert Button-Hotkey
BOOL WINAPI MTblQueryButtonHotkey( HWND hWndTbl, LPINT piKey, LPBOOL pbCtrl, LPBOOL pbShift, LPBOOL pbAlt )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*piKey = psc->m_iBtnKey;
	*pbCtrl = psc->m_bBtnKeyCtrl;
	*pbShift = psc->m_bBtnKeyShift;
	*pbAlt = psc->m_bBtnKeyAlt;

	return TRUE;
}


// MTblQueryCellImageFlags
// Ermittelt Bild-Flags einer Zelle
BOOL WINAPI MTblQueryCellImageFlags( HWND hWndCol, long lRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
		{
			return pCol->m_Image.QueryFlags( dwFlags );
		}
	}

	return FALSE;
}


// MTblQueryCellImageExpFlags
// Ermittelt Bild-Flags einer Zelle in einer expandierten Zeile
BOOL WINAPI MTblQueryCellImageExpFlags( HWND hWndCol, long lRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
		if ( pCol )
		{
			return pCol->m_Image2.QueryFlags( dwFlags );
		}
	}

	return FALSE;
}


// MTblQueryCellFlags
// Ermittelt Flags einer Zelle
BOOL WINAPI MTblQueryCellFlags( HWND hWndCol, long lRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	return psc->QueryCellFlags( lRow, hWndCol, dwFlags );
}


// MTblQueryCellMode
// Ermittelt Zellen-Modus
BOOL WINAPI MTblQueryCellMode( HWND hWndTbl, LPBOOL pbActive, LPDWORD pdwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*pbActive = psc->m_bCellMode;
	*pdwFlags = psc->m_dwCellModeFlags;

	return TRUE;
}


// MTblQueryCellType
// Ermittelt Zelltyp
BOOL WINAPI MTblQueryCellType( HWND hWndCol, long lRow, LPINT piType, LPDWORD pdwFlags, LPINT piLines, LPHSTRING phsCheckedVal, LPHSTRING phsUnCheckedVal )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Definition ermitteln
	CMTblCellType ct, *pct = NULL;
	if ( lRow != TBL_Error )
	{
		// Gültige Zeile?
		if ( !psc->IsRow( lRow ) ) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

		// Zelltyp ermitteln
		
		CMTblRow *pRow = psc->m_Rows->GetRowPtr( lRow );
		if ( pRow )
		{
			CMTblCol *pCell = pRow->m_Cols->Find( hWndCol );
			if ( pCell && pCell->m_pCellType )
			{
				pct = pCell->m_pCellType;
			}
		}
	}
	else
	{
		CMTblCol *pCol = psc->m_Cols->Find( hWndCol );
		if ( pCol && pCol->m_pCellType )
		{
			pct = pCol->m_pCellType;
		}
	}

	if ( !pct )
		pct = &ct;

	// Werte ermitteln
	return pct->Get( piType, pdwFlags, piLines, phsCheckedVal, phsUnCheckedVal );
}


// MTblQueryColHdrGrpFlags
// Ermittelt, ob best. Spaltenheadergruppen-Flags gesetzt sind
BOOL WINAPI MTblQueryColHdrGrpFlags( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	return p->QueryFlags( dwFlags );
}


// MTblQueryColHdrGrpImageFlags
// Ermittelt Bild-Flags einer Spaltenheadergruppe
BOOL WINAPI MTblQueryColHdrGrpImageFlags( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	return p->m_Image.QueryFlags( dwFlags );
}


// MTblQueryColLines
// Ermittelt die Definition der Spaltenlinien
BOOL WINAPI MTblQueryColLines( HWND hWndTbl, LPINT piStyle, LPDWORD pclr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*piStyle = psc->m_pColsVertLineDef->Style;
	*pclr = psc->m_pColsVertLineDef->Color;

	return TRUE;
}


// MTblQueryColumnFlags
// Ermittelt, ob best. Spalten-Flags gesetzt sind
BOOL WINAPI MTblQueryColumnFlags( HWND hWndCol, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	return psc->QueryColumnFlags( hWndCol, dwFlags );
}


// MTblQueryColumnHdrFlags
// Ermittelt, ob best. Spaltenkopf-Flags gesetzt sind
BOOL WINAPI MTblQueryColumnHdrFlags( HWND hWndCol, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	return psc->QueryColumnHdrFlags( hWndCol, dwFlags );
}


// MTblQueryColumnHdrImageFlags
// Ermittelt Bild-Flags eines Spaltenheaders
BOOL WINAPI MTblQueryColumnHdrImageFlags( HWND hWndCol, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Header ermitteln
	CMTblCol * pCol = psc->m_Cols->FindHdr( hWndCol );
	if ( pCol )
	{
		return pCol->m_Image.QueryFlags( dwFlags );
	}

	return FALSE;
}


// MTblQueryCornerImageFlags
// Ermittelt Bild-Flags vom Corner
BOOL WINAPI MTblQueryCornerImageFlags( HWND hWndTbl, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->m_Corner->m_Image.QueryFlags( dwFlags );
}


// MTblQueryDropDownListContext
// Ermittelt den aktuellen Listen-Kontext einer Spalte
BOOL WINAPI MTblQueryDropDownListContext( HWND hWndCol, LPLONG plRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return TBL_Error;

	// Ermitteln
	return psc->GetDropDownListContext( hWndCol, plRow );
}


// MTblQueryEffCellImageFlags
// Ermittelt effektive Bild-Flags einer Zelle
BOOL WINAPI MTblQueryEffCellImageFlags( HWND hWndCol, long lRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return NULL;

	// Bild ermitteln
	CMTblImg Img = psc->GetEffCellImage( lRow, hWndCol );

	return Img.QueryFlags( dwFlags );
}


// MTblQueryEffCellPropDet
// Ermittelt die Definition der Ermittlung der effektiven Zellen-Eigenschaften
BOOL WINAPI MTblQueryEffCellPropDet( HWND hWndTbl, LPDWORD pdwOrder, LPDWORD pdwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*pdwOrder = psc->m_dwEffCellPropDetOrder;
	*pdwFlags = psc->m_dwEffCellPropDetFlags;

	return TRUE;
}


// MTblQueryEffItemLine
// Liefert die effektive Linien-Definition eines Items
BOOL WINAPI MTblQueryEffItemLine(HUDV hUdvItem, HUDV hUdvLineDef, int iLineType)
{
	MTBLITEM item;
	if (!item.FromUdv(hUdvItem))
		return FALSE;

	CMTblLineDef ld;
	if (!MTblQueryEffItemLineC(item, ld, iLineType))
		return FALSE;

	if (!ld.ToUdv(hUdvLineDef))
		return FALSE;

	// Fertig
	return TRUE;
}


// MTblQueryEffItemLineC
// Liefert die effektive Linien-Definition eines Items, Variante mit C-Klassen
BOOL WINAPI MTblQueryEffItemLineC(MTBLITEM& item, MTBLLINEDEF& LineDef, int iLineType)
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem(&item);
	if (!psc) return FALSE;

	// Zeile prüfen
	CMTblRow *pRow = NULL;
	switch (item.Type)
	{
	case MTBL_ITEM_CELL:
	case MTBL_ITEM_ROW:
	case MTBL_ITEM_ROWHDR:
		pRow = item.GetRowPtr();

		// Gültige Zeile?
		if (!pRow) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if (pRow->IsDelAdj()) return FALSE;

		break;
	}

	// Liniendefinition ermitteln
	if (!psc->GetEffItemLine(item, iLineType, LineDef))
		return FALSE;

	// Fertig
	return TRUE;
}


// MTblQueryExtMsgsOptions
// Ermittelt, ob eine oder mehrere Optionen für das Generieren erweiterter Nachrichten gesetzt sind
BOOL WINAPI MTblQueryExtMsgsOptions( HWND hWndTbl, DWORD dwOptions )
{
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return (BOOL)psc->m_dwExtMsgOpt & dwOptions;
}


// MTblQueryFlags
// Ermittelt, ob Flags gesetzt sind
BOOL WINAPI MTblQueryFlags( HWND hWndTbl, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Flags setzen
	BOOL bOk = psc->QueryFlags( dwFlags );

	return bOk;
}


// MTblQueryFocusCell
// Liefert die aktuelle Fokus-Zelle
BOOL WINAPI MTblQueryFocusCell( HWND hWndTbl, LPLONG plRow, LPHWND phWndCol )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Pointer checken
	if ( !plRow )
		return FALSE;
	if ( !phWndCol )
		return FALSE;

	// Ermitteln
	if ( psc->m_bCellMode )
		return psc->GetEffFocusCell( *plRow, *phWndCol );
	else
		return Ctd.TblQueryFocus( hWndTbl, plRow, phWndCol );
}


// MTblQueryFocusFrame
// Ermittelt Einstellungen Fokus-Rahmen
BOOL WINAPI MTblQueryFocusFrame( HWND hWndTbl, LPINT piThickness, LPINT piOffsX, LPINT piOffsY, LPDWORD pdwColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Ermitteln
	*piThickness = psc->m_lFocusFrameThick;
	*piOffsX = psc->m_lFocusFrameOffsX;
	*piOffsY = psc->m_lFocusFrameOffsY;
	*pdwColor = psc->m_dwFocusFrameColor;

	return TRUE;
}


// MTblQueryHighlighting
// Ermittelt die Hervorhebung für ein Item
BOOL WINAPI MTblQueryHighlighting( HWND hWndTbl, long lItem, long lPart, HUDV hUdv )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// UDV dereferenzieren
	LPMTBLHILIUDV pUDV = (LPMTBLHILIUDV)Ctd.UdvDeref( hUdv );
	if ( !pUDV )
		return FALSE;

	LPMTBLHILI phld;

	// Aus Map ermitteln
	HiLiDefMap::iterator it;
	it = psc->m_pHLDM->find( LONG_PAIR( lItem, lPart ) );
	if ( it != psc->m_pHLDM->end( ) )
		phld = &(*it).second;
	else
		return FALSE;

	// Daten in UDV setzen
	Ctd.CvtULongToNumber( phld->dwBackColor, &pUDV->BackColor );
	Ctd.CvtULongToNumber( phld->dwTextColor, &pUDV->TextColor );
	Ctd.CvtLongToNumber( phld->lFontEnh, &pUDV->FontEnh );
	Ctd.CvtULongToNumber( (DWORD)phld->Img.GetHandle( ), &pUDV->Image );
	Ctd.CvtULongToNumber( (DWORD)phld->ImgExp.GetHandle( ), &pUDV->ImageExp );

	return TRUE;
}


// MTblQueryHyperlinkCell
// Liefert Hyperlink-Definition einer Zelle
BOOL WINAPI MTblQueryHyperlinkCell( HWND hWndCol, long lRow, LPBOOL pbEnabled, LPDWORD pdwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Ermitteln
	CMTblHyLnk * pHyLnk = NULL;
	CMTblRow * pRow = psc->m_Rows->GetRowPtr( lRow );
	if ( pRow )
	{
		pHyLnk = pRow->m_Cols->GetHyLnk( hWndCol );
	}

	if( pHyLnk )
	{
		*pbEnabled = pHyLnk->IsEnabled( );
		*pdwFlags = pHyLnk->m_wFlags;
	}
	else
	{
		*pbEnabled = FALSE;
		*pdwFlags = 0;
	}

	return TRUE;
}


// MTblQueryHyperlinkHover
// Liefert Hyperlink-HOVER Definition
BOOL WINAPI MTblQueryHyperlinkHover( HWND hWndTbl, LPBOOL pbEnabled, LPLONG plAddFontEnh, LPDWORD pdwTextColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*pbEnabled = psc->m_bHLHover;
	*plAddFontEnh = psc->m_lHLHoverFontEnh;
	*pdwTextColor = psc->m_dwHLHoverTextColor;

	return TRUE;
}


// MTblQueryItemButtonC
// Liefert Button-Definition eines Items, Variante mit C-Strukturen
BOOL WINAPI MTblQueryItemButtonC( LPVOID pItemV, LPVOID pBtnDefV )
{
	// Check Parameter
	if ( !(pItemV && pBtnDefV) ) return FALSE;

	// Void-Pointer in Strukturpointer umwandeln
	LPMTBLITEM pItem = (LPMTBLITEM)pItemV;
	LPMTBLBTNDEF pBtnDef = (LPMTBLBTNDEF)pBtnDefV;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem( pItem );
	if ( !psc ) return FALSE;

	// Zeile prüfen
	CMTblRow *pRow = NULL;
	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
		pRow = pItem->GetRowPtr( );

		// Gültige Zeile?
		if ( !pRow ) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if ( pRow->IsDelAdj( ) ) return FALSE;	

		break;
	}

	// Button ermitteln
	CMTblBtn *pBtn = NULL;
	switch ( pItem->Type )
	{
	case MTBL_ITEM_CELL:
		pBtn = pRow->m_Cols->GetBtn( pItem->WndHandle );
		break;
	case MTBL_ITEM_COLUMN:
		pBtn = psc->m_Cols->GetBtn( pItem->WndHandle );
		break;
	case MTBL_ITEM_COLHDR:
		pBtn = psc->m_Cols->GetHdrBtn( pItem->WndHandle );
		break;
	}

	// Definitionsstruktur setzen
	if( pBtn)
	{
		pBtnDef->bVisible = pBtn->m_bShow;
		pBtnDef->iWidth = pBtn->m_iWidth;
		pBtnDef->sText = pBtn->m_sText;
		pBtnDef->hImage = pBtn->m_hImage;
		pBtnDef->dwFlags = pBtn->m_dwFlags;
	}
	else
	{
		pBtnDef->bVisible = FALSE;
		pBtnDef->iWidth = 0;
		pBtnDef->sText = _T("");
		pBtnDef->hImage = NULL;
		pBtnDef->dwFlags = 0;
	}

	// Fertig
	return TRUE;
}



// MTblQueryItemButton
// Liefert Button-Definition eines Items
BOOL WINAPI MTblQueryItemButton( HUDV hUdvItem, HUDV hUdvBtnDef )
{
	MTBLITEM item;
	if ( !item.FromUdv( hUdvItem ) )
		return FALSE;

	MTBLBTNDEF btndef;
	if ( !MTblQueryItemButtonC( &item, &btndef ) )
		return FALSE;

	if ( !btndef.ToUdv( hUdvBtnDef ) )
		return FALSE;

	// Fertig
	return TRUE;
}


// MTblQueryItemLine
// Liefert Linien-Definition eines Items
BOOL WINAPI MTblQueryItemLine(HUDV hUdvItem, HUDV hUdvLineDef, int iLineType)
{
	MTBLITEM item;
	if (!item.FromUdv(hUdvItem))
		return FALSE;

	CMTblLineDef ld;
	if (!MTblQueryItemLineC(&item, &ld, iLineType))
		return FALSE;

	if (!ld.ToUdv(hUdvLineDef))
		return FALSE;

	// Fertig
	return TRUE;
}


// MTblQueryItemLineC
// Liefert Linien-Definition eines Items, Variante mit C-Strukturen
BOOL WINAPI MTblQueryItemLineC(LPVOID pItemV, LPVOID pLineDefV, int iLineType)
{
	// Check Parameter
	if (!(pItemV && pLineDefV)) return FALSE;

	// Void-Pointer in Strukturpointer umwandeln
	LPMTBLITEM pItem = (LPMTBLITEM)pItemV;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem(pItem);
	if (!psc) return FALSE;

	// Prüfung Item
	CMTblRow* pRow = NULL;
	CMTblColHdrGrp* pCHG = NULL;
	switch (pItem->Type)
	{
	case MTBL_ITEM_CELL:
	case MTBL_ITEM_ROW:
	case MTBL_ITEM_ROWHDR:
		pRow = pItem->GetRowPtr();

		// Gültige Zeile?
		if (!pRow) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if (pRow->IsDelAdj()) return FALSE;

		break;
	case MTBL_ITEM_COLHDRGRP:
		pCHG = (CMTblColHdrGrp*)pItem->Id;

		if (!pCHG)
			return FALSE;

		if (!psc->m_ColHdrGrps->IsValidGrp(pCHG))
			return FALSE;
		break;
	}

	// Liniendefinition ermitteln
	LPMTBLLINEDEF pLineDef = psc->GetItemLine(*pItem, iLineType);
	if (pLineDef)
		*LPMTBLLINEDEF(pLineDefV) = *pLineDef;
	else
		LPMTBLLINEDEF(pLineDefV)->Init();

	// Fertig
	return TRUE;
}


// MTblQueryLastLockedColLine
// Ermittelt die Definition der Spaltenlinie der letzten verankerten Spalte
BOOL WINAPI MTblQueryLastLockedColLine( HWND hWndTbl, LPINT piStyle, LPDWORD pclr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*piStyle = psc->m_pLastLockedColVertLineDef->Style;
	*pclr = psc->m_pLastLockedColVertLineDef->Color;

	return TRUE;
}


// MTblQueryLockedColumns
// Ermittelt die tatsächliche Anzahl gelockter Spalten
int WINAPI MTblQueryLockedColumns( HWND hWndTbl )
{
	if ( !Ctd.IsPresent( ) ) return -1;

	CMTblMetrics tm( hWndTbl );
	int iLocked = 0;
	if ( tm.m_LockedColumns > 0 )
	{
		HWND hWndCol = NULL;
		int iColPos;
		//NUMBER n;
		POINT pt;

		pt.y = tm.m_RowHeaderRight;
		while ( iColPos = MTblFindNextCol( hWndTbl, &hWndCol, 0, 0 ) )
		{
			//if ( Ctd.IsWindowVisible( hWndCol ) )
			if ( SendMessage( hWndCol, TIM_QueryColumnFlags, 0, COL_Visible ) )
			{
				pt.x = pt.y;
				//Ctd.TblQueryColumnWidth( hWndCol, &n );
				//pt.y += Ctd.FormUnitsToPixels( hWndCol, n, FALSE );
				pt.y += SendMessage( hWndCol, TIM_QueryColumnWidth, 0, 0 );
				pt.y -= 1;

				if ( iColPos <= tm.m_LockedColumns && pt.y <= tm.m_LockedColumnsRight )
					++iLocked;
				else
					break;

				pt.y += 1;
			}
			else
				++iLocked;
		}
	}

	return iLocked;
}


// MTblQueryRowFlagImageFlags
// Ermittelt, ob ein Row-Flag-Image bestimmte Flags hat
BOOL WINAPI MTblQueryRowFlagImageFlags( HWND hWndTbl, DWORD dwRowFlag, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->QueryRowFlagImgFlags( dwRowFlag, dwFlags );
}

// MTblQueryRowFlags
// Ermittelt, ob eine Zeile bestimmte Flags hat
BOOL WINAPI MTblQueryRowFlags( HWND hWndTbl, long lRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	return psc->QueryRowFlags( lRow, dwFlags );
}


// MTblQueryRowLines
// Ermittelt die Definition der Zeilenlinien
BOOL WINAPI MTblQueryRowLines( HWND hWndTbl, LPINT piStyle, LPDWORD pclr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*piStyle = psc->m_pRowsHorzLineDef->Style;
	*pclr = psc->m_pRowsHorzLineDef->Color;

	return TRUE;
}


// MTblQueryTipFlags
// Ermittelt Tip-Flags
BOOL WINAPI MTblQueryTipFlags( HWND hWndParam, long lParam, int iType, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_GetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip DefTip;
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType );
	if ( !pTip )
		pTip = &DefTip;

	return pTip->QueryFlags( dwFlags, TRUE );
}


// MTblQueryTree
// Ermittelt die Tree-Definition
BOOL WINAPI MTblQueryTree( HWND hWndTbl, LPHWND phWndTreeCol, LPINT piNodeSize, LPINT piIndent )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*phWndTreeCol = psc->m_hWndTreeCol;
	*piNodeSize = psc->m_iNodeSize;
	*piIndent = psc->m_iIndent;

	return TRUE;
}


// MTblQueryTreeFlags
// Ermittelt, ob Tree-Flags gesetzt sind
BOOL WINAPI MTblQueryTreeFlags( HWND hWndTbl, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Flags setzen
	BOOL bOk = psc->QueryTreeFlags( dwFlags );

	return bOk;
}


// MTblQueryTreeLines
// Ermittelt die Definition der Tree-Linien
BOOL WINAPI MTblQueryTreeLines( HWND hWndTbl, LPINT piStyle, LPDWORD pclr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	*piStyle = psc->m_TreeLineDef.Style;
	*pclr = psc->m_TreeLineDef.Color;

	return TRUE;
}


// MTblQueryVisibleSplitRange
// Ermittelt die erste und letzte vollständig sichtbare Split-Zeile
BOOL WINAPI MTblQueryVisibleSplitRange( HWND hWndTbl, LPLONG plMin, LPLONG plMax )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->QueryVisRange( *plMin, *plMax, TRUE );
}


// MTblRegisterLanguage
BOOL WINAPI MTblRegisterLanguage( int iLanguage, LPTSTR lpsLanguageFile, BOOL bShowErrMsg /*= TRUE*/ )
{
	if ( !lpsLanguageFile )
	{
		if ( bShowErrMsg )
			MessageBox( NULL, _T("Parameter 'lpsLanguageFile' is NULL"), _T("MTblRegisterLanguage - Error"), MB_OK | MB_ICONERROR );
		return FALSE;
	}

	TCHAR szPath[_MAX_PATH] = _T("");
	tstring sLanguageFile = _T("");

	// Ermitteln, ob Pfad in Dateiname enthalten ist
	TCHAR szDrive[_MAX_DRIVE] = _T("");
	TCHAR szDir[_MAX_DIR] = _T("");
	TCHAR szFileName[_MAX_FNAME] = _T("");
	TCHAR szExt[_MAX_EXT] = _T("");
	_tsplitpath( lpsLanguageFile, szDrive, szDir, szFileName, szExt );
	BOOL bPathInFile = lstrlen( szDir ) > 0;

	// Wenn Pfad nicht in Dateiname, Pfad dieser DLL ermitteln
	if ( !bPathInFile )
	{
		if ( GetModuleFileName( (HMODULE)ghInstance, szPath, sizeof( szPath ) ) == 0 )
		{
			if ( bShowErrMsg )
				MessageBox( NULL, _T("GetModuleFileName returns 0"), _T("MTblRegisterLanguage - Error"), MB_OK | MB_ICONERROR );
			return FALSE;
		}
		_tsplitpath( szPath, szDrive, szDir, szFileName, szExt );
		sLanguageFile += szDrive;
		sLanguageFile += szDir;
		sLanguageFile += lpsLanguageFile;
	}
	else
	{
		sLanguageFile = lpsLanguageFile;
	}

	// Prüfen, ob Datei existiert
	FILE *pf = _tfopen( sLanguageFile.c_str( ), _T("r") );
	if ( !pf )
	{
		if ( bShowErrMsg )
		{
			tstring sError = _T("Could not open language file ");
			sError += sLanguageFile;		
			MessageBox( NULL, sError.c_str( ), _T("MTblRegisterLanguage - Error"), MB_OK | MB_ICONERROR );
		}
		return FALSE;
	}
	int iRet = fclose( pf );

	// Registrieren...
	LanguageMap::value_type v( iLanguage, sLanguageFile );
	LanguageMap::iterator it = glm.find( iLanguage );
	if ( it != glm.end( ) )
	{
		glm.erase( it );
	}
	glm.insert( v );

	return TRUE;
}


// MTblResetTip
BOOL WINAPI MTblResetTip(HWND hWndTbl)
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass(hWndTbl);
	if (!psc) return FALSE;

	// Tooltip zurücksetzen
	psc->ResetTip();

	return TRUE;
}

// MTblSetAltRowBackColors
BOOL WINAPI MTblSetAltRowBackColors( HWND hWndTbl, BOOL bSplitRows, DWORD dwColor1, DWORD dwColor2 )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( bSplitRows )
	{
		psc->m_dwSplitRowAlternateBackColor1 = dwColor1;
		psc->m_dwSplitRowAlternateBackColor2 = dwColor2;

	}
	else
	{
		psc->m_dwRowAlternateBackColor1 = dwColor1;
		psc->m_dwRowAlternateBackColor2 = dwColor2;
	}

	InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblSetCellBackColor
// Setzt die Hintergrundfarbe einer Zelle
BOOL WINAPI MTblSetCellBackColor( HWND hWndCol, long lRow, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc )
		return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) )
		return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) )
		return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow )
		return FALSE;

	// Zelle ermitteln
	CMTblCol * pCol = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCol )
		return FALSE;

	// Hintergrundfarbe setzen
	pCol->m_BackColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetCellFlags
// Setzt die Flags einer Zelle
BOOL WINAPI MTblSetCellFlags( HWND hWndCol, long lRow, DWORD dwFlags, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc )
		return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) )
		return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) )
		return FALSE;
	
	// Flags setzen
	DWORD dwDiff = 0;
	BOOL bOk = psc->SetCellFlags( lRow, hWndCol, NULL, dwFlags, bSet, TRUE, &dwDiff );

	// Ggf. neu zeichnen
	if ( dwDiff )
	{
		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE );
		}
	}

	return bOk;
}


// MTblSetCellIndentLevel
// Setzt die Einrückungs-Ebene einer Zelle
BOOL WINAPI MTblSetCellIndentLevel( HWND hWndCol, long lRow, long lIndentLevel )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc )
		return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) )
		return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) )
		return FALSE;
	
	BOOL bOk = psc->SetCellIndentLevel( lRow, hWndCol, lIndentLevel, NULL, TRUE );

	if ( bOk )
	{
		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE );
		}
	}

	return bOk;
}

// MTblSetCellFont
// Setzt den Font einer Zelle
BOOL WINAPI MTblSetCellFont( HWND hWndCol, long lRow, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Zelle ermitteln
	CMTblCol * pCol = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Font setzen
	if ( !pCol->m_Font.Set( lpsName, lSize, lEnh ) ) return FALSE;

	// Neu zeichnen?
	if ( dwFlags & MTSF_REDRAW )
	{
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE, FALSE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetCellImage
// Setzt das normale Bild einer Zelle
BOOL WINAPI MTblSetCellImage( HWND hWndCol, long lRow, HANDLE hImage, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Zelle ermitteln
	CMTblCol * pCol = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Bild setzen
	if ( !pCol->m_Image.SetHandle( hImage, psc->m_pCounter ) ) return FALSE;

	// Bild-Flags setzen
	pCol->m_Image.ClearFlags( );
	pCol->m_Image.SetFlags( dwFlags, TRUE );
	pCol->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	/*if ( hImage )
	{
		pCol->m_Image.SetFlags( dwFlags, TRUE );
		pCol->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	}*/

	// Neu zeichnen?
	if ( dwFlags & MTSI_REDRAW )
	{
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );
	}

	return TRUE;
}


// MTblSetCellImageExp
// Setzt das Bild "expandierte Zeile" einer Zelle
BOOL WINAPI MTblSetCellImageExp( HWND hWndCol, long lRow, HANDLE hImage, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Zelle ermitteln
	CMTblCol * pCol = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Bild setzen
	if ( !pCol->m_Image2.SetHandle( hImage, psc->m_pCounter ) ) return FALSE;

	// Bild-Flags setzen
	pCol->m_Image2.ClearFlags( );
	pCol->m_Image2.SetFlags( dwFlags, TRUE );
	pCol->m_Image2.SetFlags( MTSI_REDRAW, FALSE );
	/*if ( hImage )
	{
		pCol->m_Image2.SetFlags( dwFlags, TRUE );
		pCol->m_Image2.SetFlags( MTSI_REDRAW, FALSE );
	}*/

	// Neu zeichnen?
	if ( dwFlags & MTSI_REDRAW )
	{
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );
	}

	return TRUE;
}


// MTblSetCellMerge
// Verbindet eine Zelle mit x benachbarten
BOOL WINAPI MTblSetCellMerge( HWND hWndCol, long lRow, int iCount, DWORD dwFlags )
{
	/*
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Merge-Zellen ungültig
	psc->m_bMergeCellsInvalid = TRUE;

	// Merge setzen
	if ( !pRow->m_Cols->SetMerged( hWndCol, iCount ) ) return FALSE;

	// Fokus-Zelle
	psc->UpdateFocus( TBL_Error, NULL, (dwFlags & MTSM_REDRAW) ? UF_REDRAW_INVALIDATE : UF_REDRAW_AUTO );

	// Neu zeichnen?
	if ( dwFlags & MTSM_REDRAW )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
	*/

	return MTblSetCellMergeEx( hWndCol, lRow, iCount, 0, dwFlags );
}


// MTblSetCellMergeEx
// Setzt erweiterte Zellenverbindung ( Spalten + Zeilen )
BOOL WINAPI MTblSetCellMergeEx( HWND hWndCol, long lRow, int iColCount, int iRowCount, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Zelle ermitteln
	CMTblCol *pCell = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCell ) return FALSE;

	// Aktuelle Werte ermitteln
	int iColCountOld = pCell->m_Merged;
	int iRowCountOld = pCell->m_MergedRows;

	// Ende, wenn keine Änderung
	if ( iColCount == iColCountOld && iRowCount == iRowCountOld ) return TRUE;

	// Merge setzen
	pCell->SetMerged( iColCount );
	pCell->SetMergedRows( iRowCount );

	// Entscheidung, ob Merge-Zellen komplett neu aufgebaut werden oder ob hinzugefügt wird
	BOOL bSetInvalid = TRUE;
	
	if ( !psc->m_bMergeCellsInvalid && iColCountOld == 0 && iRowCountOld == 0 )
	{
		LPMTBLMERGECELL pmc = NULL;
		HWND hWndColCheck;
		long lRowCheck = lRow;
		long lRowMax = lRow < 0 ? TBL_MaxSplitRow : TBL_MaxRow;
		for( int r = 0; ; ++r )
		{
			hWndColCheck = hWndCol;
			for( int c = 0; ; ++c )
			{
				pmc = psc->FindMergeCell( hWndColCheck, lRowCheck, FMC_ANY );
				if ( pmc ) break;
				if ( c == iColCount || !psc->FindNextCol( &hWndColCheck, COL_Visible, 0 ) ) break;
			}

			if ( pmc ) break;
			if ( r == iRowCount || !psc->FindNextRow( &lRowCheck, 0, ROW_Hidden, lRowMax ) )
				break;
		}

		if ( !pmc )
			bSetInvalid = FALSE;
	}

	if ( bSetInvalid )
	{
		// Merge-Zellen ungültig
		psc->m_bMergeCellsInvalid = TRUE;
	}
	else
	{
		// Merge-Zelle hinzufügen
		psc->m_lRowAddToMergeCells = lRow;
		psc->m_hWndColAddToMergeCells = hWndCol;
	}

	// Fokus-Zelle
	psc->UpdateFocus( TBL_Error, NULL, (dwFlags & MTSM_REDRAW) ? UF_REDRAW_INVALIDATE : UF_REDRAW_AUTO );

	// Neu zeichnen?
	if ( dwFlags & MTSM_REDRAW )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblSetCellTextColor
// Setzt die Textfarbe einer Zelle
BOOL WINAPI MTblSetCellTextColor( HWND hWndCol, long lRow, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Zelle ermitteln
	CMTblCol * pCol = pRow->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Textfarbe setzen
	pCol->m_TextColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateCell( hWndCol, lRow, NULL, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol && lFocusRow == lRow )
				psc->InvalidateEdit( FALSE, TRUE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetColHdrGrpBackColor
// Setzt die Hintergrundfarbe einer Spaltenheader-Gruppe
BOOL WINAPI MTblSetColHdrGrpBackColor( HWND hWndTbl, LPVOID pGrp, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Hintergrundfarbe setzen
	p->m_BackColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateColHdr( FALSE );
	}

	return TRUE;
}


// MTblSetColHdrGrpCol
// Fügt eine Spalte einer Spaltenheader-Gruppe hinzu oder entfernt sie
BOOL WINAPI MTblSetColHdrGrpCol( HWND hWndTbl, LPVOID pGrp, HWND hWndCol, BOOL bSet, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Gruppe?
	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Setzen
	if ( !psc->m_ColHdrGrps->SetCol( p, hWndCol, bSet ) ) return FALSE;

	// Text der Spalte anpassen
	psc->AdaptColHdrText( hWndCol );

	// Neu zeichnen
	psc->InvalidateColHdr( FALSE );

	return TRUE;
}


// MTblSetColHdrGrpFlags
// Setzt Flags einer Spaltenheader-Gruppe
BOOL WINAPI MTblSetColHdrGrpFlags( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Flags setzen
	if ( p->SetFlags( dwFlags, bSet ) )
	{
		// Texte der zugehörigen Spalten-Header anpassen
		list<HWND>::iterator it, itEnd = p->m_Cols.end( );
		for ( it = p->m_Cols.begin( ); it != itEnd; ++it )
		{
			psc->AdaptColHdrText( *it );
		}

		psc->InvalidateColHdr( FALSE );
	}

	return TRUE;
}

// MTblSetColHdrGrpFont
// Setzt den Font einer Spaltenheader-Gruppe
BOOL WINAPI MTblSetColHdrGrpFont( HWND hWndTbl, LPVOID pGrp, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Font setzen
	if ( !p->m_Font.Set( lpsName, lSize, lEnh ) ) return FALSE;

	// Neu zeichnen?
	if ( dwFlags & MTSF_REDRAW )
	{
		psc->InvalidateColHdr( FALSE );
	}

	return TRUE;
}


// MTblSetColHdrGrpImage
// Setzt das Bild einer Spaltenheader-Gruppe
BOOL WINAPI MTblSetColHdrGrpImage( HWND hWndTbl, LPVOID pGrp, HANDLE hImage, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Bild setzen
	if ( !p->m_Image.SetHandle( hImage, psc->m_pCounter ) ) return FALSE;

	// Bild-Flags setzen
	p->m_Image.ClearFlags( );
	p->m_Image.SetFlags( dwFlags, TRUE );
	p->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	/*if ( hImage )
	{
		p->m_Image.SetFlags( dwFlags, TRUE );
		p->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	}*/
	
	// Neu zeichnen?
	if ( dwFlags & MTSI_REDRAW )
	{
		psc->InvalidateColHdr( FALSE );
	}

	return TRUE;
}


// MTblSetColHdrGrpText
// Setzt den Text einer Spalten-Header-Gruppe
BOOL WINAPI MTblSetColHdrGrpText( HWND hWndTbl, LPVOID pGrp, LPCTSTR lpsText, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Text setzen
	if ( lpsText )
		p->m_Text = lpsText;
	else
		p->m_Text = _T("");

	// Texte der zugehörigen Spalten-Header anpassen
	list<HWND>::iterator it, itEnd = p->m_Cols.end( );
	for ( it = p->m_Cols.begin( ); it != itEnd; ++it )
	{
		psc->AdaptColHdrText( *it );
	}

	// Neu zeichnen
	psc->InvalidateColHdr( FALSE );

	return TRUE;
}


// MTblSetColHdrGrpTextColor
// Setzt die Textfarbe einer Spaltenheader-Gruppe
BOOL WINAPI MTblSetColHdrGrpTextColor( HWND hWndTbl, LPVOID pGrp, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	CMTblColHdrGrp *p = (CMTblColHdrGrp*)pGrp;
	if ( !psc->m_ColHdrGrps->IsValidGrp( p ) ) return FALSE;

	// Hintergrundfarbe setzen
	p->m_TextColor.Set( dwColor );

	// Neu zeichnen
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateColHdr( FALSE );
	}

	return TRUE;
}


// MTblSetColumnBackColor
// Setzt die Hintergrundfarbe einer Spalte
BOOL WINAPI MTblSetColumnBackColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Hintergrundfarbe setzen
	pCol->m_BackColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateCol( hWndCol, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol )
				psc->InvalidateEdit( TRUE, TRUE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetColumnFlags
// Setzt Spalten-Flags
BOOL WINAPI MTblSetColumnFlags( HWND hWndCol, DWORD dwFlags, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Flags setzen
	DWORD dwDiff = pCol->SetFlags( dwFlags, bSet );

	// Read-Only-Status sicherstellen
	if ( dwDiff & MTBL_COL_FLAG_READ_ONLY )
		psc->EnsureFocusCellReadOnly( );

	// Evtl. neu zeichnen
	if ( dwDiff )
	{
		psc->InvalidateCol( hWndCol, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol )
				psc->InvalidateEdit( TRUE, TRUE );
		}

	}

	return TRUE;
}


// MTblSetColumnFont
// Setzt den Font einer Spalte
BOOL WINAPI MTblSetColumnFont( HWND hWndCol, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Font setzen
	if ( !pCol->m_Font.Set( lpsName, lSize, lEnh ) ) return FALSE;

	// Neu zeichnen?
	if ( dwFlags & MTSF_REDRAW )
	{
		psc->InvalidateCol( hWndCol, FALSE );
	}

	return TRUE;
}


// MTblSetColumnTextColor
// Setzt die Textdfarbe einer Spalte
BOOL WINAPI MTblSetColumnTextColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;

	// Hintergrundfarbe setzen
	pCol->m_TextColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateCol( hWndCol, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && hWndFocusCol == hWndCol )
				psc->InvalidateEdit( FALSE, TRUE, TRUE );
		}
	}

	return TRUE;
}

// MTblSetColumnTitle
// Setzt den Titel einer Spalte
BOOL WINAPI MTblSetColumnTitle( HWND hWndCol, LPCTSTR lpsTitle )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	return psc->SetColumnTitle( hWndCol, lpsTitle );
}


// MTblSetCornerBackColor
// Setzt Corner-Hintergrundfarbe
BOOL WINAPI MTblSetCornerBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farbe setzen
	psc->m_Corner->m_BackColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateCorner( FALSE ) ;
	}

	return TRUE;
}


// MTblSetCornerImage
// Setzt Corner-Bild
BOOL WINAPI MTblSetCornerImage( HWND hWndTbl, HANDLE hImage, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Bild setzen
	if ( !psc->m_Corner->m_Image.SetHandle( hImage, psc->m_pCounter ) ) return FALSE;

	// Bild-Flags setzen
	psc->m_Corner->m_Image.ClearFlags( );
	psc->m_Corner->m_Image.SetFlags( dwFlags, TRUE );
	psc->m_Corner->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	/*if ( hImage )
	{
		psc->m_Corner->m_Image.SetFlags( dwFlags, TRUE );
		psc->m_Corner->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	}*/

	// Neu zeichnen?
	if ( dwFlags & MTSI_REDRAW )
	{
		psc->InvalidateCorner( FALSE ) ;
	}

	return TRUE;

}


// MTblSetDropDownListContext
// Setzt Kontext für SalList* Funktionen
BOOL WINAPI MTblSetDropDownListContext( HWND hWndCol, long lRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	return psc->SetDropDownListContext( hWndCol, lRow );
}


// MTblSetExportSubstColor
// Setzt eine Export-Ersatzfarbe
BOOL WINAPI MTblSetExportSubstColor( HWND hWndTbl, DWORD dwColor, DWORD dwExportColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Parameter-Check
	if ( dwColor == MTBL_COLOR_UNDEF )
		return FALSE;

	// Farbe setzen
	if ( dwExportColor == MTBL_COLOR_UNDEF )
		psc->m_escm->erase( dwColor );
	else
	{
		SubstColorMap::value_type val( dwColor, dwExportColor );
		psc->m_escm->insert( val );
	}

	return TRUE;
}


// MTblSetExtMsgsOptions
// Setzt oder entfernt Optionen für die Generierung erweiterter Nachrichten
BOOL WINAPI MTblSetExtMsgsOptions( HWND hWndTbl, DWORD dwOptions, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	psc->m_dwExtMsgOpt = ( bSet ? ( psc->m_dwExtMsgOpt | dwOptions ) : ( ( psc->m_dwExtMsgOpt & dwOptions ) ^ psc->m_dwExtMsgOpt ) );

	return TRUE;
}


// MTblSetFlags
// Setzt Flags
BOOL WINAPI MTblSetFlags( HWND hWndTbl, DWORD dwFlags, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Evtl. Zeilenselektionen entfernen
	if ( dwFlags & MTBL_FLAG_SUPPRESS_ROW_SELECTION && bSet )
		Ctd.TblClearSelection( hWndTbl );

	// Evtl. Focus entfernen
	if ( dwFlags & MTBL_FLAG_SUPPRESS_FOCUS && bSet )
		//Ctd.TblKillFocus( hWndTbl );
		psc->KillFocusFrameI( );

	// Flags setzen
	DWORD dwOldFlags = psc->m_dwFlags;
	BOOL bOk = psc->SetFlags( dwFlags, bSet );

	// Variable Zeilenhöhe / Kein Column-Header
	if ( (dwFlags & MTBL_FLAG_VARIABLE_ROW_HEIGHT) || (dwFlags & MTBL_FLAG_NO_COLUMN_HEADER) )
	{
		BOOL bAdapt = FALSE;
		BOOL bWasSet;

		// Zeilenhöhen werden abgeschaltet und Zeilenhöhen vorhanden?
		bWasSet = (dwOldFlags & MTBL_FLAG_VARIABLE_ROW_HEIGHT) != 0;
		if ( (dwFlags & MTBL_FLAG_VARIABLE_ROW_HEIGHT) && !bSet && bWasSet && psc->m_pCounter->GetRowHeights( ) > 0 )
		{
			// Zeilenhöhen zurücksetzen
			psc->m_Rows->ResetRowHeights( FALSE );
			psc->m_Rows->ResetRowHeights( TRUE );
			psc->m_pCounter->InitRowHeights( );

			bAdapt = TRUE;
		}

		// Kein Column-Header wird umgeschaltet?
		bWasSet = (dwOldFlags & MTBL_FLAG_NO_COLUMN_HEADER) != 0;
		if ( (dwFlags & MTBL_FLAG_NO_COLUMN_HEADER) && bSet != bWasSet )
			bAdapt = TRUE;

		if ( bAdapt )
		{
			// Scrollbereiche anpassen
			psc->AdaptScrollRangeV( FALSE );
			psc->AdaptScrollRangeV( TRUE );

			// Fokus aktualisieren
			psc->UpdateFocus( );
		}
	}

	// List-Clip anpassen?
	if ( ( dwFlags & MTBL_FLAG_BUTTONS_OUTSIDE_CELL ) || ( dwFlags & MTBL_FLAG_ADAPT_LIST_WIDTH ) )
		psc->AdaptListClip( );

	// Neu zeichnen
	if ( psc->m_dwFlags != dwOldFlags )
	{
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && lFocusRow != TBL_Error )
				psc->InvalidateEdit( TRUE, TRUE );
		}
	}

	return bOk;
}


// MTblSetFocusCell
// Setzt die Fokus-Zelle
BOOL WINAPI MTblSetFocusCell( HWND hWndTbl, long lRow, HWND hWndCol, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->SetFocusCell( lRow, hWndCol );
}


// MTblSetColumnHdrFreeBackColor
// Setzt die Hintergrundfarbe des freien Spaltenheader-Bereichs
BOOL WINAPI MTblSetColumnHdrFreeBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	psc->m_dwColHdrFreeBackColor = dwColor;

	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateColHdr( FALSE );
	}

	return TRUE;
}


// MTblSetHeadersBackColor
// Setzt Hintergrundfarbe der Header
BOOL WINAPI MTblSetHeadersBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farbe setzen
	psc->m_dwHdrsBackColor = dwColor;

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateColHdr( FALSE );
		psc->InvalidateRowHdr( TBL_Error, FALSE );
		psc->InvalidateCorner( FALSE ) ;
	}

	return TRUE;
}


// MTblSetHighlighted
// Setzt oder entfernt Hervorhebung eines Items
BOOL WINAPI MTblSetHighlighted( HUDV hUdvItem, BOOL bSet )
{
	CMTblItem item;
	item.FromUdv( hUdvItem );

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromItem( &item );
	if ( !psc ) return FALSE;

	// Setzen
	if ( !psc->SetItemHighlighted( item, bSet ) )
		return FALSE;

	// Ggf. neu zeichnen
	HiLiDefMap::iterator it;
	it = psc->m_pHLDM->find( LONG_PAIR( item.Type, item.Part ) );
	if ( it != psc->m_pHLDM->end( ) && !(*it).second.IsUndef( ) )
	{
		long lFocusRow = TBL_Error;
		HWND hWndFocusCol = NULL;
		Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol );

		switch( item.Type )
		{
		case MTBL_ITEM_CELL:
			psc->InvalidateCell( item.WndHandle, item.GetRowNr( ), NULL, FALSE );
			if ( hWndFocusCol && hWndFocusCol == item.WndHandle && lFocusRow == item.GetRowNr( ) )
				psc->InvalidateEdit( TRUE, TRUE, TRUE, TRUE );
			break;
		case MTBL_ITEM_COLUMN:
			psc->InvalidateCol( item.WndHandle, FALSE );
			if ( hWndFocusCol && hWndFocusCol == item.WndHandle )
				psc->InvalidateEdit( TRUE, TRUE, TRUE, TRUE );
			break;
		case MTBL_ITEM_COLHDR:
			psc->InvalidateColHdr( item.WndHandle, FALSE );
			break;
		case MTBL_ITEM_COLHDRGRP:
			psc->InvalidateColHdr( FALSE ) ;
			break;
		case MTBL_ITEM_ROW:
			psc->InvalidateRow( item.GetRowNr( ), FALSE, TRUE );
			if ( hWndFocusCol && lFocusRow == item.GetRowNr( ) )
				psc->InvalidateEdit( TRUE, TRUE, TRUE, TRUE );
			break;
		case MTBL_ITEM_ROWHDR:
			psc->InvalidateRowHdr( item.GetRowNr( ), FALSE, TRUE ); 
			break;
		case MTBL_ITEM_CORNER:
			psc->InvalidateCorner( FALSE );
			break;
		}

		UpdateWindow( item.GetTableHandle( ) );
	}

	return TRUE;
}


// MTblSetIndentPerLevel
// Setzt die Einrückung in Pixel pro Level
BOOL WINAPI MTblSetIndentPerLevel( HWND hWndTbl, long lIndentPerLevel )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Einrückung setzen
	psc->m_lIndentPerLevel = max( lIndentPerLevel, 0 );

	// Neu zeichnen
	InvalidateRect( hWndTbl, NULL, FALSE );

	long lFocusRow;
	HWND hWndFocusCol;
	if ( Ctd.TblQueryFocus( psc->m_hWndTbl, &lFocusRow, &hWndFocusCol ) )
	{
		if ( hWndFocusCol )
			psc->InvalidateEdit( TRUE, TRUE );
	}

	return TRUE;
}


// MTblSetMsgOffest
// Setzt Offset von SAM_User der MTM_Nachrichten
BOOL WINAPI MTblSetMsgOffset( UINT uiOffs )
{
	if ( !gscm.empty( ) ) return FALSE;

	guiMsgOffs = uiOffs;
	return TRUE;
}


// MTblSetMWheelOption
// Setzt eine Mausrad-Scroll-Optionen
BOOL WINAPI MTblSetMWheelOption( HWND hWndTbl, int iOption, int iValue )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	BOOL bOk = FALSE;

	// Option setzen
	switch ( iOption )
	{
	case MTMW_ROWS_PER_NOTCH:
		psc->m_iMWRowsPerNotch = max( iValue, 0 );
		bOk = TRUE;
		break;
	case MTMW_SPLITROWS_PER_NOTCH:
		psc->m_iMWSplitRowsPerNotch = max( iValue, 0 );
		bOk = TRUE;
		break;
	case MTMW_PAGE_KEY:
		psc->m_iMWPageKey = max( iValue, 0 );
		bOk = TRUE;
		break;
	}

	return bOk;
}


// MTblSetNodeImages
BOOL WINAPI MTblSetNodeImages( HWND hWndTbl, HANDLE hImg, HANDLE hImgExp, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Bilder setzen
	if ( !psc->SetNodeImages( hImg, hImgExp ) )
		return FALSE;

	// Ggf. neu zeichnen
	if ( dwFlags & MTSI_REDRAW && psc->m_hWndTreeCol )
		psc->InvalidateCol( psc->m_hWndTreeCol, FALSE );

	return TRUE;
}


// MTblSetColumnHdrBackColor
// Setzt die Hintergrundfarbe eines Spaltenheaders
BOOL WINAPI MTblSetColumnHdrBackColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;

	// Hintergrundfarbe setzen
	pCol->m_BackColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		/*int iLeft, iTop, iRight, iBottom;
		if ( MTblGetColHdrRect( hWndCol, &iLeft, &iTop, &iRight, &iBottom ) > 0 )
		{
			RECT r;
			SetRect( &r, iLeft, iTop, iRight + 1, iBottom + 1 );
			InvalidateRect( psc->m_hWndTbl, &r, FALSE );
		}*/
		psc->InvalidateColHdr( hWndCol, FALSE );
	}

	return TRUE;
}


// MTblSetColumnHdrFlags
// Setzt Spaltenkopf-Flags
BOOL WINAPI MTblSetColumnHdrFlags( HWND hWndCol, DWORD dwFlags, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;

	// Flags setzen
	DWORD dwDiff = pCol->SetFlags( dwFlags, bSet );

	// Evtl. neu zeichnen
	if ( dwDiff )
		psc->InvalidateColHdr( FALSE );

	return TRUE;
}


// MTblSetColumnHdrFont
// Setzt den Font eines Spaltenheaders
BOOL WINAPI MTblSetColumnHdrFont( HWND hWndCol, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spaltenheader ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;

	// Font setzen
	if ( !pCol->m_Font.Set( lpsName, lSize, lEnh ) ) return FALSE;

	// Neu zeichnen?
	if ( dwFlags & MTSF_REDRAW )
	{
		/*int iLeft, iTop, iRight, iBottom;
		if ( MTblGetColHdrRect( hWndCol, &iLeft, &iTop, &iRight, &iBottom ) > 0 )
		{
			RECT r;
			SetRect( &r, iLeft, iTop, iRight + 1, iBottom + 1 );
			InvalidateRect( psc->m_hWndTbl, &r, FALSE );
		}*/
		psc->InvalidateColHdr( hWndCol, FALSE );
	}

	return TRUE;
}


// MTblSetColumnHdrImage
// Setzt das Bild eines Spaltenheaders
BOOL WINAPI MTblSetColumnHdrImage( HWND hWndCol, HANDLE hImage, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;

	// Bild setzen
	if ( !pCol->m_Image.SetHandle( hImage, psc->m_pCounter ) ) return FALSE;

	// Bild-Flags setzen
	pCol->m_Image.ClearFlags( );
	pCol->m_Image.SetFlags( dwFlags, TRUE );
	pCol->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	/*if ( hImage )
	{
		pCol->m_Image.SetFlags( dwFlags, TRUE );
		pCol->m_Image.SetFlags( MTSI_REDRAW, FALSE );
	}*/

	// Neu zeichnen?
	if ( dwFlags & MTSI_REDRAW )
	{
		/*int iLeft, iTop, iRight, iBottom;
		if ( MTblGetColHdrRect( hWndCol, &iLeft, &iTop, &iRight, &iBottom ) > 0 )
		{
			RECT r;
			SetRect( &r, iLeft, iTop, iRight + 1, iBottom + 1 );
			InvalidateRect( psc->m_hWndTbl, &r, FALSE );
		}*/
		psc->InvalidateColHdr( hWndCol, FALSE );
	}

	return TRUE;
}


// MTblSetColumnHdrTextColor
// Setzt die Textfarbe eines Spaltenheaders
BOOL WINAPI MTblSetColumnHdrTextColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClassFromCol( hWndCol );
	if ( !psc ) return FALSE;

	// Spalte ermitteln
	CMTblCol * pCol = psc->m_Cols->EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;

	// Hintergrundfarbe setzen
	pCol->m_TextColor.Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		/*int iLeft, iTop, iRight, iBottom;
		if ( MTblGetColHdrRect( hWndCol, &iLeft, &iTop, &iRight, &iBottom ) > 0 )
		{
			RECT r;
			SetRect( &r, iLeft, iTop, iRight + 1, iBottom + 1 );
			InvalidateRect( psc->m_hWndTbl, &r, FALSE );
		}*/
		psc->InvalidateColHdr( hWndCol, FALSE );
	}

	return TRUE;
}


// MTblSetOwnerDrawnItem
// Deklariert ein Item als benutzerdef. gezeichnet
BOOL WINAPI MTblSetOwnerDrawnItem( HWND hWndParam, long lParam, int iType, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = NULL;

	switch ( iType )
	{
	case MTBL_ODI_CELL:
	case MTBL_ODI_COLUMN:
	case MTBL_ODI_COLHDR:
		psc = MTblGetSubClassFromCol( hWndParam );
		break;
	case MTBL_ODI_COLHDRGRP:
	case MTBL_ODI_ROW:
	case MTBL_ODI_ROWHDR:
	case MTBL_ODI_CORNER:
		psc = MTblGetSubClass( hWndParam );
		break;
	default:
		return FALSE;
	}
	if ( !psc ) return FALSE;

	// Setzen
	return psc->SetOwnerDrawnItem( hWndParam, lParam, iType, bSet );
}


// MTblSetParentRow
// Setzt die Elternzeile einer Zeile
BOOL WINAPI MTblSetParentRow( HWND hWndTbl, long lRow, long lParentRow, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Dafür sorgen, dass beim Setzen die Elternzeile als ursprüngliche Elternzeile gesetzt wird
	dwFlags |= MTSPR_AS_ORIG;

	// Setzen
	return psc->SetParentRow( lRow, lParentRow, dwFlags );
}

// MTblSetPasswordChar
// Setzt das Passwort-Zeichen
BOOL WINAPI MTblSetPasswordChar( HWND hWndTbl, LPCTSTR lpsPwd )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( !lpsPwd ) return FALSE;
	if ( lstrlen( lpsPwd ) < 1 ) return FALSE;

	TCHAR cPwd = lpsPwd[0];

	if ( cPwd != psc->m_cPassword )
	{
		psc->m_cPassword = cPwd;
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );
		if ( IsWindow( psc->m_hWndEdit ) && SendMessage( psc->m_hWndEdit, EM_GETPASSWORDCHAR, 0, 0 ) )
		{
			SendMessage( psc->m_hWndEdit, EM_SETPASSWORDCHAR, (WPARAM)cPwd, 0 );
			InvalidateRect( psc->m_hWndEdit, NULL, 0 );
		}
	}

	return TRUE;
}


// MTblSetRowBackColor
// Setzt die Hintergrundfarbe einer Zeile
BOOL WINAPI MTblSetRowBackColor( HWND hWndTbl, long lRow, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Hintergrundfarbe setzen
	pRow->m_BackColor->Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateRow( lRow, FALSE, TRUE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE, TRUE );
		}

	}

	return TRUE;
}

// MTblSetRowFlagImage
BOOL WINAPI MTblSetRowFlagImage( HWND hWndTbl, WORD wRowFlag, HANDLE hImage, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;
	
	if ( !psc->SetRowFlagImg( wRowFlag, hImage, TRUE, dwFlags ) )
		return FALSE;

	if ( dwFlags & MTSI_REDRAW )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}

// MTblSetRowFlags
// Setzt oder entfernt Flags einer Zeile
BOOL WINAPI MTblSetRowFlags( HWND hWndTbl, long lRow, DWORD dwRowFlags, BOOL bSet, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeilenzeiger ermitteln
	CMTblRow *pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow )
		return FALSE;

	// Ggf. Hierarchie normalisieren
	BOOL bNormalized = FALSE;
	if ( psc->QueryTreeFlags( MTBL_TREE_FLAG_AUTO_NORM_HIER ) && (dwRowFlags & MTBL_ROW_HIDDEN) /*&& pRow->QueryFlags( MTBL_ROW_HIDDEN ) != bSet*/ )
	{
		if ( SendMessage( hWndTbl, MTM_QueryAutoNormHierarchy, lRow, bSet ? NORM_HIER_ROW_HIDE : NORM_HIER_ROW_SHOW ) != MTBL_NODEFAULTACTION )
		{
			if ( bSet && pRow->QueryFlags( MTBL_ROW_CANEXPAND ) && !pRow->IsParent( ) )
			{
				if ( psc->m_lLoadChildRows > 0 )
				{
					DWORD dwIntFlags = ROW_IFLAG_LOAD_CHILD_ROWS | ( dwFlags & MTSRF_REDRAW ? ROW_IFLAG_HIDE_REDRAW : ROW_IFLAG_HIDE );
					pRow->SetInternalFlags( dwIntFlags, TRUE );
					return TRUE;
				}
				else
				{
					psc->LoadChildRows( lRow );
					lRow = pRow->GetNr( );
				}
			}
			

			BOOL bNormalizeHierarchy = bSet ? psc->m_Rows->IsParentRow( lRow ) : TRUE;
			if ( bNormalizeHierarchy )
				bNormalized = psc->NormalizeHierarchy( lRow, bSet ? NORM_HIER_ROW_HIDE : NORM_HIER_ROW_SHOW );
		}
	}

	// Flags setzen
	DWORD dwDiff = psc->m_Rows->SetRowFlags( lRow, dwRowFlags, bSet );

	// Read-Only-Status sicherstellen
	if ( dwDiff & MTBL_ROW_READ_ONLY )
		psc->EnsureFocusCellReadOnly( );

	// Anzeigen/Verbergen	
	if ( dwDiff & MTBL_ROW_HIDDEN )
	{
		// ROW_Hidden setzen
		BOOL bSetSalFlag = TRUE;
		if ( !bSet )
		{
			CMTblRow *pParentRow = NULL;
			if ( ( pParentRow = pRow->GetParentRow( ) ) && !pParentRow->QueryFlags( MTBL_ROW_ISEXPANDED ) )
				bSetSalFlag = FALSE;
		}

		if ( bSetSalFlag )
			Ctd.TblSetRowFlags( hWndTbl, lRow, ROW_Hidden, bSet );
	}

	// Neu zeichnen?
	if ( ( dwFlags & MTSRF_REDRAW ) && ( dwDiff || bNormalized ) )
	{
		if ( bNormalized )
			InvalidateRect( hWndTbl, NULL, FALSE );
		else
		{
			psc->InvalidateRow( lRow, FALSE, TRUE );
			if ( dwDiff & MTBL_ROW_HIDDEN )
			{
				long lParentRow = psc->m_Rows->GetParentRow( lRow );
				if ( lParentRow != TBL_Error )
					psc->InvalidateRow( lParentRow, FALSE, TRUE );
			}
		}

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetRowFont
// Setzt den Font einer Zeile
BOOL WINAPI MTblSetRowFont( HWND hWndTbl, long lRow, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Font setzen
	if ( !pRow->m_Font->Set( lpsName, lSize, lEnh ) ) return FALSE;

	// Neu zeichnen?
	if ( dwFlags & MTSF_REDRAW )
	{
		psc->InvalidateRow( lRow, FALSE, TRUE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && lFocusRow == lRow )
				psc->InvalidateEdit( TRUE, TRUE, FALSE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetRowHdrBackColor
// Setzt die Header-Hintergrundfarbe einer Zeile
BOOL WINAPI MTblSetRowHdrBackColor( HWND hWndTbl, long lRow, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	if ( lRow == TBL_Error )
	{
		// Farbe freier Bereich setzen
		psc->m_dwRowHdrFreeBackColor = dwColor;
	}
	else
	{
		// Gültige Zeile?
		if ( !psc->IsRow( lRow ) ) return FALSE;

		// Zeile mit TBL_Adjust gelöscht?
		if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
		
		// Zeile ermitteln
		CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
		if ( !pRow ) return FALSE;

		// Hintergrundfarbe setzen
		pRow->m_HdrBackColor->Set( dwColor );
	}	

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
		psc->InvalidateRowHdr( lRow, FALSE, TRUE );

	return TRUE;
}


// MTblSetRowHeight
// Setzt die Höhe einer Zeile, 0 = Stanardhöhe
BOOL WINAPI MTblSetRowHeight( HWND hWndTbl, long lRow, long lHeight, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Variable Zeilenhöhe eingeschaltet?
	if ( !psc->QueryFlags( MTBL_FLAG_VARIABLE_ROW_HEIGHT ) )
		return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Höhe setzen
	if ( !psc->m_Rows->SetRowHeight( lRow, lHeight ) )
		return FALSE;

	// Scroll-Range anpassen
	psc->AdaptScrollRangeV( lRow < 0 );

	// List-Clip anpassen
	psc->AdaptListClip( );

	// Fokus
	psc->UpdateFocus( TBL_Error, NULL, (dwFlags & MTSRH_REDRAW) ? UF_REDRAW_INVALIDATE : UF_REDRAW_AUTO );

	// Neu zeichnen?
	if ( dwFlags & MTSRH_REDRAW )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblSetRowMerge
// Verbindet eine Zeile mit x Zeilen unterhalb
BOOL WINAPI MTblSetRowMerge( HWND hWndTbl, long lRow, int iCount, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Aktuellen Zellenverbund ermitteln + kopieren
	LPMTBLMERGEROW pmrOld = psc->FindMergeRow( lRow, FMR_MASTER );
	MTBLMERGEROW mr;
	if ( pmrOld )
	{
		pmrOld->CopyTo( mr );
		pmrOld = &mr;
	}

	// Zeilen löschen/einfügen?
	if ( dwFlags & MTSM_INS_DEL_ROWS )
	{
		int i;
		int iMergedOld = pmrOld ? pmrOld->iMergedRows : 0;

		if ( iCount != iMergedOld )
		{
			psc->m_bMergeRowsInvalid = TRUE;
			psc->m_Rows->SetMerged( lRow, 0 );

			if ( iCount < iMergedOld )
			{
				long lRowDel = pmrOld->lRowTo;
				for ( i = iMergedOld - iCount; i > 0; i-- )
				{
					Ctd.TblDeleteRow( hWndTbl, lRowDel, TBL_NoAdjust );
					if ( !psc->FindPrevRow( &lRowDel, 0, ROW_Hidden, pmrOld->lRowFrom ) )
						break;
				}
			}
			else if ( iCount > iMergedOld )
			{
				long lRowIns = (pmrOld ? pmrOld->lRowTo : lRow) + 1;
				for ( i = iCount - iMergedOld; i > 0; i-- )
				{
					lRowIns = Ctd.TblInsertRow( hWndTbl, lRowIns );
					if ( lRowIns == TBL_Error )
						break;

					lRowIns++;
				}
			}
		}
	}

	// Merge-Zeilen ungültig
	psc->m_bMergeRowsInvalid = TRUE;

	// Merge setzen
	if ( !psc->m_Rows->SetMerged( lRow, iCount ) ) return FALSE;

	// Automerge Zellen
	if ( !(dwFlags & MTSM_NO_CELLMERGE) )
		psc->AutoMergeCells( lRow, pmrOld, 0 );

	// Zeilenflags replizieren
	LPMTBLMERGEROW pmr = psc->FindMergeRow( lRow, FMR_MASTER );
	if ( pmr && !psc->m_bReplMergeRowsFlags )
	{
		BOOL bSet;
		DWORD dwFlagsCheck[] = {ROW_Selected,ROW_New,ROW_Edited,ROW_MarkDeleted,ROW_HideMarks,ROW_UserFlag1,ROW_UserFlag2,ROW_UserFlag3,ROW_UserFlag4,ROW_UserFlag5};
		
		psc->m_bReplMergeRowsFlags = TRUE;

		int iMax = sizeof(dwFlagsCheck) / sizeof(DWORD);
		if (psc->QueryFlags(MTBL_FLAG_NO_USER_ROW_FLAG_REPL))
			iMax -= 5;

		for ( int i = 0; i < iMax; i++ )
		{
			bSet = Ctd.TblQueryRowFlags( hWndTbl, lRow, dwFlagsCheck[i] );
			psc->SetMergeRowFlags( pmr, dwFlagsCheck[i], bSet, FALSE, lRow );			
		}

		psc->m_bReplMergeRowsFlags = FALSE;
	}

	// Fokus
	psc->UpdateFocus( TBL_Error, NULL, (dwFlags & MTSM_REDRAW) ? UF_REDRAW_INVALIDATE : UF_REDRAW_AUTO );

	// Neu zeichnen?
	if ( dwFlags & MTSM_REDRAW )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblSetRowTextColor
// Setzt die Textfarbe einer Zeile
BOOL WINAPI MTblSetRowTextColor( HWND hWndTbl, long lRow, DWORD dwColor, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;
	
	// Zeile ermitteln
	CMTblRow * pRow = psc->m_Rows->EnsureValid( lRow );
	if ( !pRow ) return FALSE;

	// Hintergrundfarbe setzen
	pRow->m_TextColor->Set( dwColor );

	// Neu zeichnen?
	if ( dwFlags & MTSC_REDRAW )
	{
		psc->InvalidateRow( lRow, FALSE, TRUE );

		long lFocusRow;
		HWND hWndFocusCol;
		if ( Ctd.TblQueryFocus( hWndTbl, &lFocusRow, &hWndFocusCol ) )
		{
			if ( hWndFocusCol && lFocusRow == lRow )
				psc->InvalidateEdit( FALSE, TRUE, TRUE );
		}
	}

	return TRUE;
}


// MTblSetRowUserData
// Setzt Benutzerdaten einer Zeile
BOOL WINAPI MTblSetRowUserData( HWND hWndTbl, long lRow, LPTSTR lpsKey, LPTSTR lpsValue )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Gültige Zeile?
	if ( !psc->IsRow( lRow ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( psc->m_Rows->IsRowDelAdj( lRow ) ) return FALSE;

	// Daten setzen
	tstring sKey( lpsKey );
	tstring sValue( lpsValue );
	return psc->m_Rows->SetRowUserData( lRow, sKey, sValue );
}


// MTblSetSelectionColors
// Setzt die Selektionsfarben
BOOL WINAPI MTblSetSelectionColors( HWND hWndTbl, DWORD dwTextColor, DWORD dwBackColor )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	BOOL bInvalidate = FALSE;

	if ( dwTextColor != psc->m_dwSelTextColor )
	{
		if ( dwTextColor == MTBL_COLOR_UNDEF  )
			psc->m_pCounter->DecNoSelInv( );
		else if ( psc->m_dwSelTextColor == MTBL_COLOR_UNDEF )
			psc->m_pCounter->IncNoSelInv( );
		bInvalidate = TRUE;
		psc->m_dwSelTextColor = dwTextColor;
	}

	if ( dwBackColor != psc->m_dwSelBackColor )
	{
		if ( dwBackColor == MTBL_COLOR_UNDEF  )
			psc->m_pCounter->DecNoSelInv( );
		else if ( psc->m_dwSelBackColor == MTBL_COLOR_UNDEF )
			psc->m_pCounter->IncNoSelInv( );
		bInvalidate = TRUE;
		psc->m_dwSelBackColor = dwBackColor;
	}

	if ( bInvalidate )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}

// MTblSetSelectionDarkening
// Setzt die Selektionsverdunkelung
BOOL WINAPI MTblSetSelectionDarkening(HWND hWndTbl, BYTE byDarkening)
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass(hWndTbl);
	if (!psc) return FALSE;

	BOOL bInvalidate = FALSE;

	if (byDarkening != psc->m_bySelDarkening)
	{
		if (byDarkening == 0)
			psc->m_pCounter->DecNoSelInv();
		else if (psc->m_bySelDarkening == 0)
			psc->m_pCounter->IncNoSelInv();
		bInvalidate = TRUE;
		psc->m_bySelDarkening = byDarkening;
	}

	if (bInvalidate)
		InvalidateRect(psc->m_hWndTbl, NULL, FALSE);

	return TRUE;
}


// MTblSetTipBackColor
// Setzt die Tip-Hintergrundfarbe
BOOL WINAPI MTblSetTipBackColor( HWND hWndParam, long lParam, int iType, DWORD dwColor )
{
	// Undefinierte Standard-Hintergrundfarbe nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && dwColor == MTBL_COLOR_UNDEF )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	pTip->m_BackColor.Set( dwColor );

	return TRUE;
}


// MTblSetTipDelay
// Setzt die Tip-Anzeigeverzögerung
BOOL WINAPI MTblSetTipDelay( HWND hWndParam, long lParam, int iType, NUMBER nDelay )
{
	// Undefinierte Standard-Anzeigeverzögerung nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && nDelay.numLength == 0 )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	// Wert setzen
	if ( nDelay.numLength == 0 )
		pTip->m_iDelay = INT_MIN;
	else
		pTip->m_iDelay = max( Ctd.NumToInt( nDelay ), 0 );

	return TRUE;
}


// MTblSetTipFadeInTime
// Setzt Tip-Einblendzeit
BOOL WINAPI MTblSetTipFadeInTime( HWND hWndParam, long lParam, int iType, NUMBER nTime )
{
	// Undefinierte Standard-Einblendzeit nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && nTime.numLength == 0 )
		return FALSE;

	// Wert OK?
	int iTime = Ctd.NumToInt( nTime );
	if ( iTime < 0 || iTime > 5000 )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	if ( nTime.numLength == 0 )
		pTip->m_iFadeInTime = INT_MIN;
	else
		pTip->m_iFadeInTime = max( Ctd.NumToInt( nTime ), 0 );

	return TRUE;
}


// MTblSetTipFlags
// Setzt Tip-Flags
BOOL WINAPI MTblSetTipFlags( HWND hWndParam, long lParam, int iType, DWORD dwFlags, BOOL bSet )
{
	// Standardflag MTBL_TIP_FLAG_NODEFFLAGS kann nicht gesetzt werden
	if ( iType == MTBL_TIP_DEFAULT && dwFlags & MTBL_TIP_FLAG_NODEFFLAGS && bSet )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	return pTip->SetFlags( dwFlags, bSet );
}


// MTblSetTipFont
// Setzt den Tip-Font
BOOL WINAPI MTblSetTipFont( HWND hWndParam, long lParam, int iType, LPTSTR lpsName, long lSize, long lEnh )
{
	// Undefinierte Standard-Fonteinstellungen nicht möglich
	CMTblFont f;
	f.Set( lpsName, lSize, lEnh );
	if ( iType == MTBL_TIP_DEFAULT && f.AnyUndef( ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	return pTip->m_Font.Set( lpsName, lSize, lEnh );
}


// MTblSetTipFrameColor
// Setzt die Tip-Rahmenfarbe
BOOL WINAPI MTblSetTipFrameColor( HWND hWndParam, long lParam, int iType, DWORD dwColor )
{
	// Undefinierte Standard-Rahmenfarbe nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && dwColor == MTBL_COLOR_UNDEF )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	pTip->m_FrameColor.Set( dwColor );

	return TRUE;
}


// MTblSetTipOffset
// Setzt die Tip-Offsets
BOOL WINAPI MTblSetTipOffset( HWND hWndParam, long lParam, int iType, NUMBER nOffsX, NUMBER nOffsY )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	if ( nOffsX.numLength == 0 )
		pTip->m_iOffsX = INT_MIN;
	else
		pTip->m_iOffsX = Ctd.NumToInt( nOffsX );

	if ( nOffsY.numLength == 0 )
		pTip->m_iOffsY = INT_MIN;
	else
		pTip->m_iOffsY = Ctd.NumToInt( nOffsY );

	return TRUE;
}


// MTblSetTipOpacity
// Setzt Tip-Opazität
BOOL WINAPI MTblSetTipOpacity( HWND hWndParam, long lParam, int iType, NUMBER nOpacity )
{
	// Undefinierte Standardopazität nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && nOpacity.numLength == 0 )
		return FALSE;

	// Wert OK?
	int iOpacity = Ctd.NumToInt( nOpacity );
	if ( iOpacity < 0 || iOpacity > 255 )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	if ( nOpacity.numLength == 0 )
		pTip->m_iOpacity = INT_MIN;
	else
		pTip->m_iOpacity = iOpacity;

	return TRUE;
}


// MTblSetTipPos
// Setzt die Tip-Position
BOOL WINAPI MTblSetTipPos( HWND hWndParam, long lParam, int iType, NUMBER nX, NUMBER nY )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	if ( nX.numLength == 0 )
		pTip->m_iPosX = INT_MIN;
	else
		pTip->m_iPosX = Ctd.NumToInt( nX );

	if ( nY.numLength == 0 )
		pTip->m_iPosY = INT_MIN;
	else
		pTip->m_iPosY = Ctd.NumToInt( nY );

	return TRUE;
}


// MTblSetTipRCESize
// Setzt die Tip-RCE-Größe ( RCE = RoundCornerEllipse )
BOOL WINAPI MTblSetTipRCESize( HWND hWndParam, long lParam, int iType, NUMBER nWid, NUMBER nHt )
{
	// Undefinierte Standard-Breite/Höhe nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && ( nWid.numLength == 0 || nHt.numLength == 0 ) )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	if ( nWid.numLength == 0 )
		pTip->m_iRCEWid = INT_MIN;
	else
		pTip->m_iRCEWid = Ctd.NumToInt( nWid );

	if ( nHt.numLength == 0 )
		pTip->m_iRCEHt = INT_MIN;
	else
		pTip->m_iRCEHt = Ctd.NumToInt( nHt );

	return TRUE;
}


// MTblSetTipSize
// Setzt die Tip-Größe
BOOL WINAPI MTblSetTipSize( HWND hWndParam, long lParam, int iType, NUMBER nWid, NUMBER nHt )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	if ( nWid.numLength == 0 )
		pTip->m_iWid = INT_MIN;
	else
		pTip->m_iWid = Ctd.NumToInt( nWid );

	if ( nHt.numLength == 0 )
		pTip->m_iHt = INT_MIN;
	else
		pTip->m_iHt = Ctd.NumToInt( nHt );

	return TRUE;
}


// MTblSetTipText
// Setzt den Tip-Text
BOOL WINAPI MTblSetTipText( HWND hWndParam, long lParam, int iType, LPCTSTR lpsText )
{
	// Setzen eines Standardtexts nicht möglich
	if ( iType == MTBL_TIP_DEFAULT )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	pTip->m_sText = ( lpsText ? lpsText : _T("") );

	return TRUE;
}


// MTblSetTipTextColor
// Setzt die Tip-Textfarbe
BOOL WINAPI MTblSetTipTextColor( HWND hWndParam, long lParam, int iType, DWORD dwColor )
{
	// Undefinierte Standard-Textfarbe nicht möglich
	if ( iType == MTBL_TIP_DEFAULT && dwColor == MTBL_COLOR_UNDEF )
		return FALSE;

	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass_SetTip( hWndParam, lParam, iType );
	if ( !psc ) return FALSE;

	// Tip ermitteln
	CMTblTip *pTip = psc->GetTip( lParam, hWndParam, iType, TRUE );
	if ( !pTip )
		return FALSE;

	pTip->m_TextColor.Set( dwColor );

	return TRUE;
}


// MTblSetTreeFlags
// Setzt Tree-Flags
BOOL WINAPI MTblSetTreeFlags( HWND hWndTbl, DWORD dwFlags, BOOL bSet )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Flags setzen
	DWORD dwOldTreeFlags = psc->m_dwTreeFlags;
	BOOL bOk = psc->SetTreeFlags( dwFlags, bSet );

	// Neu zeichnen?
	BOOL bRedraw = ( psc->m_dwTreeFlags != dwOldTreeFlags ) && IsWindow( psc->m_hWndTreeCol ) && Ctd.IsWindowVisible( psc->m_hWndTreeCol );
	if ( bRedraw )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return bOk;
}


// MTblSetTreeNodeColors
// Setzt Tree-Node-Farben
BOOL WINAPI MTblSetTreeNodeColors( HWND hWndTbl, DWORD dwOuterClr, DWORD dwInnerClr, DWORD dwBackClr )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	// Farben setzen
	BOOL bAnyChanged = FALSE;
	if ( dwOuterClr != psc->m_dwNodeColor )
	{
		psc->m_dwNodeColor = dwOuterClr;
		bAnyChanged = TRUE;
	}
	if ( dwInnerClr != psc->m_dwNodeInnerColor )
	{
		psc->m_dwNodeInnerColor = dwInnerClr;
		bAnyChanged = TRUE;
	}
	if ( dwBackClr != psc->m_dwNodeBackColor )
	{
		psc->m_dwNodeBackColor = dwBackClr;
		bAnyChanged = TRUE;
	}

	// Neu zeichnen
	BOOL bRedraw = bAnyChanged && IsWindow( psc->m_hWndTreeCol ) && Ctd.IsWindowVisible( psc->m_hWndTreeCol );
	if ( bRedraw )
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );

	return TRUE;
}


// MTblSort
// Sortiert eine Tabelle
BOOL WINAPI MTblSort( HWND hWndTbl, HARRAY hWndaCols, HARRAY naFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->Sort( hWndaCols, naFlags, TBL_Error, TBL_Error );
}


// MTblSortEx
// Erweiterte Sortierung einer Tabelle
BOOL WINAPI MTblSortEx( HWND hWndTbl, HARRAY hWndaCols, HARRAY naFlags, long lMinRow, long lMaxRow, long lFromLevel, long lToLevel, long lParentRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->Sort( hWndaCols, naFlags, lMinRow, lMaxRow, lFromLevel, lToLevel, lParentRow );
}


// MTblSortRange
// Sortiert einen Zeilenbereich einer Tabelle
BOOL WINAPI MTblSortRange( HWND hWndTbl, HARRAY hWndaCols, HARRAY naFlags, long lMinRow, long lMaxRow )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->Sort( hWndaCols, naFlags, lMinRow, lMaxRow );
}


// MTblSortTree
// Sortiert die Zeilen hierarchisch
BOOL WINAPI MTblSortTree( HWND hWndTbl, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	return psc->SortTree( dwFlags );
}


// MTblSubClass
// Führt Subclassing für eine Tabelle durch
BOOL WINAPI MTblSubClass( HWND hWndTbl )
{
	// CTD da?
	if ( !Ctd.IsPresent( ) ) return FALSE;

	long lType = 0;
	LPMTBLSUBCLASS psc = NULL;

	// Handle prüfen
	if ( !IsWindow( hWndTbl ) ) return FALSE;
	lType = Ctd.GetType( hWndTbl );
	if ( !(lType == TYPE_ChildTable || lType == TYPE_TableWindow) ) return FALSE;

	// Property schon gesetzt?
	if ( MTblIsSubClassed( hWndTbl ) ) return TRUE;

	// Subclass-Struktur erzeugen
	psc = (LPMTBLSUBCLASS)new MTBLSUBCLASS;
	if ( !psc ) return FALSE;

	// Der Subclass-Map hinzufügen
	SubClassMap::value_type val( hWndTbl, (LPVOID) psc );
	gscm.insert( val );

	// Subclass-Struktur initialisieren
	memset( LPVOID(psc), 0, sizeof( MTBLSUBCLASS ) );

	BOOL bError = FALSE;
	while ( TRUE )
	{
		psc->m_pCounter = new CMTblCounter;
		if ( !psc->m_pCounter )
		{
			bError = TRUE;
			break;
		}

		psc->m_DelAdjRows = (PLongList)new LongList;
		if ( !psc->m_DelAdjRows )
		{
			bError = TRUE;
			break;
		}

		psc->m_LastSortCols = (PHwndList)new HwndList;
		if ( !psc->m_LastSortCols )
		{
			bError = TRUE;
			break;
		}

		psc->m_LastSortFlags = (PLongList)new LongList;
		if ( !psc->m_LastSortFlags )
		{
			bError = TRUE;
			break;
		}

		psc->m_Rows = new CMTblRows( psc->m_pCounter );
		if ( !psc->m_Rows )
		{
			bError = TRUE;
			break;
		}

		psc->m_Cols = new CMTblCols( COL_ITYPE_COL, psc->m_pCounter );
		if ( !psc->m_Cols )
		{
			bError = TRUE;
			break;
		}

		psc->m_CurCols = new CCurCols( hWndTbl );
		if ( !psc->m_CurCols )
		{
			bError = TRUE;
			break;
		}

		psc->m_Corner = new CMTblCol;
		if ( !psc->m_Corner )
		{
			bError = TRUE;
			break;
		}

		psc->m_ColHdrGrps = new CMTblColHdrGrps;
		if ( !psc->m_ColHdrGrps )
		{
			bError = TRUE;
			break;
		}

		psc->m_DefTip = new CMTblTip;
		if ( !psc->m_DefTip )
		{
			bError = TRUE;
			break;
		}

		psc->m_ETTip = new CMTblTip;
		if ( !psc->m_ETTip )
		{
			bError = TRUE;
			break;
		}

		psc->m_scm = new SubClassMap;
		if ( !psc->m_scm )
		{
			bError = TRUE;
			break;
		}

		psc->m_escm = new SubstColorMap;
		if ( !psc->m_escm )
		{
			bError = TRUE;
			break;
		}

		psc->m_pHLDM = new HiLiDefMap;
		if ( !psc->m_pHLDM )
		{
			bError = TRUE;
			break;
		}

		psc->m_pHLIL = new ItemList;
		if ( !psc->m_pHLIL )
		{
			bError = TRUE;
			break;
		}

		psc->m_ppd = new MTBLPAINTDATA;
		if ( !psc->m_ppd )
		{
			bError = TRUE;
			break;
		}
		/*ZeroMemory( psc->m_ppd, sizeof( MTBLPAINTDATA ) );
		psc->m_ppd->clrTblBackColor = MTBL_COLOR_UNDEF;*/

		/*for ( int i = 0; i < MTBL_MAX_COLS; ++i )
		{
			psc->m_ppd->PaintCols[i] = new MTBLPAINTCOL;
			if ( !psc->m_ppd->PaintCols[i] )
			{
				bError = TRUE;
				break;
			}
		}*/
		/*ZeroMemory( psc->m_ppd->PaintCols, sizeof( LPMTBLPAINTCOL ) * MTBL_MAX_COLS );*/

		psc->m_RowFlagImages = new RowFlagImgMap;
		if ( !psc->m_RowFlagImages )
		{
			bError = TRUE;
			break;
		}

		psc->m_RowsLoadingChilds = new RowPtrStack;
		if ( !psc->m_RowsLoadingChilds )
		{
			bError = TRUE;
			break;
		}

		psc->m_pHdrLineDef = new CMTblLineDef;
		if (!psc->m_pHdrLineDef)
		{
			bError = TRUE;
			break;
		}

		psc->m_pColsVertLineDef = new CMTblLineDef;
		if (!psc->m_pColsVertLineDef)
		{
			bError = TRUE;
			break;
		}

		psc->m_pLastLockedColVertLineDef = new CMTblLineDef;
		if (!psc->m_pLastLockedColVertLineDef)
		{
			bError = TRUE;
			break;
		}

		psc->m_pRowsHorzLineDef = new CMTblLineDef;
		if (!psc->m_pRowsHorzLineDef)
		{
			bError = TRUE;
			break;
		}

		// Font setzen
		psc->m_hFont = (HFONT)SendMessage( hWndTbl, WM_GETFONT, 0, 0 );

		// Standard-Row-Flag-Bitmaps setzen
		psc->SetRowFlagBitmap( ROW_MarkDeleted, IDB_ROWFLAG_DELETED, RGB(255,255,255) );
		psc->SetRowFlagBitmap( ROW_New, IDB_ROWFLAG_NEW, RGB(255,255,255) );
		psc->SetRowFlagBitmap( ROW_Edited, IDB_ROWFLAG_EDITED, RGB(255,255,255) );

		// Linien-Definitionen
		#ifdef TD_SOLID_TBL_LINES
		psc->m_pRowsHorzLineDef->Style = psc->m_pColsVertLineDef->Style = MTLS_SOLID;
		#else
		psc->m_pRowsHorzLineDef->Style = psc->m_pColsVertLineDef->Style = MTLS_DOT;
		#endif

		psc->m_pHdrLineDef->Style = MTLS_SOLID;
		psc->m_pLastLockedColVertLineDef->Style = MTLS_SOLID;
		psc->m_TreeLineDef.Style = MTLS_INVISIBLE;

		psc->m_pHdrLineDef->Thickness = psc->m_pRowsHorzLineDef->Thickness = psc->m_pColsVertLineDef->Thickness = psc->m_pLastLockedColVertLineDef->Thickness = psc->m_TreeLineDef.Thickness = 1;

		// Node-Einstellungen
		psc->m_dwNodeColor = MTBL_COLOR_UNDEF;
		psc->m_dwNodeInnerColor = MTBL_COLOR_UNDEF;
		psc->m_dwNodeBackColor = MTBL_COLOR_UNDEF;

		// Hyperlink HOVER-Einstellungen
		psc->m_dwHLHoverTextColor = MTBL_COLOR_UNDEF;

		// Standardeinstellungen Tip
		psc->m_DefTip->m_TextColor.Set( 0 );
		psc->m_DefTip->m_BackColor.Set( RGB( 255, 255, 228 ) );
		psc->m_DefTip->m_FrameColor.Set( 0 );
		psc->m_DefTip->m_Font.Set( _T("Arial"), 8, 0 );
		psc->m_DefTip->m_iDelay = 400;
		psc->m_DefTip->m_iFadeInTime = 0;
		psc->m_DefTip->m_iOpacity = 255;
		psc->m_DefTip->m_iOffsX = 0;
		psc->m_DefTip->m_iOffsY = 20;
		psc->m_DefTip->m_iRCEWid = 0;
		psc->m_DefTip->m_iRCEHt = 0;
		psc->m_DefTip->SetFlags( MTBL_TIP_FLAG_PERMEABLE, TRUE );

		// Passwort-Zeichen
		psc->m_cPassword = ' ';

		// Farben Selektionen
		psc->m_dwSelBackColor = psc->m_dwSelTextColor = MTBL_COLOR_UNDEF;

		// Farben Header
		psc->m_dwHdrsBackColor = MTBL_COLOR_UNDEF;

		// Farben freie Header-Bereiche
		psc->m_dwColHdrFreeBackColor = MTBL_COLOR_UNDEF;
		psc->m_dwRowHdrFreeBackColor = MTBL_COLOR_UNDEF;

		// Abwechselnde Zeilen-Hintergrundfarben
		psc->m_dwRowAlternateBackColor1 = psc->m_dwRowAlternateBackColor2 = psc->m_dwSplitRowAlternateBackColor1 = psc->m_dwSplitRowAlternateBackColor2 = MTBL_COLOR_UNDEF;

		// Ermittlungs-Reihenfolgen
		if ( !psc->SetEffCellPropDetOrder( MTECPD_ORDER_CELL_COL_ROW ) )
		{
			bError = TRUE;
			break;
		}

		// Mousewheel-Optionen 
		psc->m_iMWRowsPerNotch = 3;
		psc->m_iMWSplitRowsPerNotch = 3;

		// Optionen erweiterte Nachrichten
		psc->m_dwExtMsgOpt = MTEM_COLHDRSEP_EXCLUSIVE;

		// Merging
		psc->m_bMergeCellsInvalid = TRUE;
		psc->m_bMergeRowsInvalid = TRUE;
		psc->m_lRowAddToMergeCells = TBL_Error;

		// Buttons
		psc->m_lRowBtnPushed = TBL_Error;

		// Veränderung Zeilengröße
		psc->m_lRowSizing = TBL_Error;

		break;
	}

	if ( bError )
	{
		delete psc->m_DelAdjRows;
		delete psc->m_LastSortCols;
		delete psc->m_LastSortFlags;
		delete psc->m_Rows;
		delete psc->m_Cols;
		delete psc->m_RowFlagImages;
		delete psc;
		return FALSE;
	}

	// Handle setzen
	psc->m_hWndTbl = hWndTbl;

	// TD-Version setzen
	psc->m_wTDVer = Ctd.GetVersion( );

	// Spaltenausrichtungs-Bug fixen?
	psc->m_bFixColJustifyBug = FALSE;
	if ( psc->m_wTDVer == 31 )
	{
		tstring sVer = _T("");
		if ( FileVerGetVersionStr( Ctd.m_szDLL, sVer ) )
		{
			if ( sVer.find( _T("PTF") ) == tstring::npos )
				psc->m_bFixColJustifyBug = TRUE;
		}
	}

	// Fensterprozedur setzen
	psc->m_wpOld = (WNDPROC)SetWindowLongPtr( hWndTbl, GWLP_WNDPROC, (LONG_PTR)MTblWndProc );

	// Focus killen
	Ctd.TblKillFocus( hWndTbl );

	// Button-Klasse registrieren
	psc->RegisterBtnClass( );

	// Spalten subclassen
	//psc->SubClassCols( );
	psc->AttachCols( );

	// Standardwerte Fokus
	psc->m_lFocusFrameThick = 3;
	psc->m_lFocusFrameOffsX = -1;
	psc->m_lFocusFrameOffsY = -1;
	psc->m_dwFocusFrameColor = MTBL_COLOR_UNDEF;

	// Standardwert Einzug pro Level
	psc->m_lIndentPerLevel = 8;

	// Gesplittet?
	BOOL bDragAdjust = FALSE;
	int iSplitRows = 0;
	Ctd.TblQuerySplitWindow( hWndTbl, &iSplitRows, &bDragAdjust );
	psc->m_bSplitted = ( iSplitRows > 0 );

	// Tip-Fenster erzeugen
	/*psc->CreateTipWnd( );*/

	// Timer setzen
	//psc->m_uiTimer = SetTimer( hWndTbl, 999, 100, psc->TimerProc );
	psc->StartInternalTimer( );

	// Properties
	MTblApplyProperties( hWndTbl );

	return TRUE;
}


// MTblSwapRows
BOOL WINAPI MTblSwapRows( HWND hWndTbl, long lRow1, long lRow2, DWORD dwFlags )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return FALSE;

	BOOL bOk;
	long lContext;
	//MTBLSWAPROWS sr;

	// Gültige Zeilen?
	lContext = Ctd.TblQueryContext( hWndTbl );
	bOk = SendMessage( hWndTbl, TIM_SetContext, 0, lRow1 ) && SendMessage( hWndTbl, TIM_SetContext, 0, lRow2 );
	SendMessage( hWndTbl, TIM_SetContext, 0, lContext );
	if ( !bOk ) return FALSE;
	
	// Editiermodus beenden
	Ctd.TblKillEdit( hWndTbl );

	// Zeilen vertauschen
	//sr.lRow1 = lRow1;
	//sr.lRow2 = lRow2;
	//bOk = (BOOL)SendMessage( hWndTbl, TIM_SwapRows, 0, (LPARAM)&sr );
	bOk = psc->SwapRows( lRow1, lRow2 );

	// Evtl. neu zeichnen
	if ( dwFlags & MTSR_REDRAW )
	{
		InvalidateRect( psc->m_hWndTbl, NULL, FALSE );
		UpdateWindow( psc->m_hWndTbl );
	}

	return bOk;
}


// MTblUnlockPaint
// Hebt das Blockieren des Neuzeichnens auf
long WINAPI MTblUnlockPaint( HWND hWndTbl )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return TBL_Error;

	return psc->UnlockPaint( );
}


// MTblWndProc
// Subclassing - Fensterprozedur
LRESULT CALLBACK MTblWndProc( HWND hWndTbl, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	// Subclass-Struktur ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
	if ( !psc ) return 0;

	BOOL bScrollFuncs;
	CMTblRow * pRow;
	LRESULT lRet = 0;
	WORD wScrollCode;

	// Alte Fensterprozedur prüfen
	if ( !psc->m_wpOld ) return 0;

	// Nachrichten verarbeiten
	switch( uMsg )
	{

	case EM_GETLINE:
		lRet = (LRESULT)MTblGetRowID(hWndTbl, lParam);
		break;

	case WM_CAPTURECHANGED:
		lRet = psc->OnCaptureChanged( wParam, lParam );
		break;

	case WM_CHAR:
		lRet = psc->OnChar( wParam, lParam );
		break;

	case WM_COMMAND:
		lRet = psc->OnCommand( wParam, lParam );
		break;

	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORSTATIC:
		lRet = psc->OnCtlColorEdit( wParam, lParam );
		break;

	case WM_GETDLGCODE:
		lRet = psc->OnGetDlgCode( wParam, lParam );
		break;

	case WM_HSCROLL:
		if ( psc->QueryFlags( MTBL_FLAG_THUMBTRACK_HSCROLL ) && LOWORD(wParam) == SB_THUMBTRACK )
			wParam = MAKELONG( SB_THUMBPOSITION, HIWORD( wParam ) );

		wScrollCode = LOWORD(wParam);
		bScrollFuncs = wScrollCode != SB_THUMBTRACK && wScrollCode != SB_ENDSCROLL;
		if ( bScrollFuncs )
			psc->BeforeScroll( );

		lRet = psc->CallOldWndProc( uMsg, wParam, lParam );

		if ( bScrollFuncs )
			psc->AfterScroll( );

		break;

	case WM_KEYDOWN:
		lRet = psc->OnKeyDown( wParam, lParam );
		break;

	case WM_KILLFOCUS:
		lRet = psc->OnKillFocus( wParam, lParam );
		break;

	case WM_LBUTTONDOWN:
		lRet = psc->OnLButtonDown( wParam, lParam );
		break;

	case WM_LBUTTONUP:
		lRet = psc->OnLButtonUp( wParam, lParam );
		break;

	case WM_LBUTTONDBLCLK:
		lRet = psc->OnLButtonDblClk( wParam, lParam );
		break;

	case WM_MOUSEMOVE:
		lRet = psc->OnMouseMove( wParam, lParam );
		break;

	case WM_MOUSEWHEEL:
		lRet = psc->OnMouseWheel( wParam, lParam );
		break;

	case WM_NCDESTROY:
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc( psc->m_wpOld, hWndTbl, uMsg, wParam, lParam);
		// Aufräumen
		psc->OnNCDestroy( wParam, lParam );
		// Speicher freigeben
		delete psc;
		break;

	case WM_PAINT:
		//lRet = psc->CallOldWndProc( uMsg, wParam, lParam );
		psc->HandlePaint( hWndTbl, wParam, lParam );
		break;

	case WM_PARENTNOTIFY:
		lRet = psc->CallOldWndProc( uMsg, wParam, lParam );
		psc->OnParentNotify( wParam, lParam );
		break;

	case WM_PASTE:
		lRet = psc->OnPaste( wParam, lParam );
		break;

	case WM_RBUTTONDOWN:
		lRet = psc->OnRButtonDown( wParam, lParam );
		break;

	case WM_RBUTTONUP:
		lRet = psc->OnRButtonUp( wParam, lParam );
		break;

	case WM_RBUTTONDBLCLK:
		lRet = psc->OnRButtonDblClk( wParam, lParam );
		break;

	case WM_SETCURSOR:
		lRet = psc->OnSetCursor( wParam, lParam );
		break;

	case WM_SETFOCUS:
		lRet = psc->OnSetFocus( wParam, lParam );
		break;

	case WM_SETFONT:
		lRet = psc->OnSetFont( wParam, lParam );
		break;

	case WM_SIZE:
		lRet = psc->OnSize( wParam, lParam );
		break;

	case WM_SYSKEYDOWN:
		lRet = psc->OnSysKeyDown( wParam, lParam );
		break;

	case WM_TIMER:
		lRet = psc->CallOldWndProc( WM_TIMER, wParam, lParam );
		break;

	case WM_VSCROLL:
		lRet = psc->OnVScroll( wParam, lParam );
		break;

	case TIM_ClearSelection:
		lRet = psc->OnClearSelection( wParam, lParam );
		break;

	case TIM_DefineSplitWindow:
		lRet = psc->OnDefineSplitWindow( wParam, lParam );
		break;

	case TIM_DeleteRow:
		lRet = psc->OnDeleteRow( wParam, lParam );
		break;

	case TIM_DestroyColumns:
		HWND hWndCol;
		hWndCol = NULL;
		//if ( !MTblFindNextCol( hWndTbl, &hWndCol, 0, COL_Dynamic ) )
		if ( !psc->FindNextCol( &hWndCol, 0, COL_Dynamic ) )
			SendMessage( hWndTbl, TIM_Reset, 0, 0 );

		lRet = psc->CallOldWndProc( uMsg, wParam, lParam );
		break;

	case TIM_EditModeChange:
		lRet = psc->OnEditModeChange( wParam, lParam );
		break;

	case TIM_FetchRow:
		lRet = psc->OnFetchRow( wParam, lParam );
		break;

	case TIM_GetMetrics:
		lRet = psc->OnGetMetrics( wParam, lParam );
		break;

	case TIM_KillEdit:
		lRet = psc->OnKillEdit( wParam, lParam );
		break;

	case TIM_KillFocusFrame:
		lRet = psc->OnKillFocusFrame( wParam, lParam );
		break;
	
	case TIM_InsertRow:
		lRet = psc->OnInsertRow( wParam, lParam );
		break;

	case TIM_ObjectsFromPoint:
		lRet = psc->OnObjectsFromPoint( wParam, lParam );
		break;

	case TIM_PosChanged:
		lRet = psc->OnPosChanged( wParam, lParam );
		break;

	case TIM_QueryVisRange:
		lRet = psc->OnQueryVisRange( wParam, lParam );
		break;

	case TIM_Reset:
		lRet = psc->OnReset( wParam, lParam );
		break;

	case TIM_RowSetContext:
		lRet = psc->OnRowSetContext( wParam, lParam );
		break;

	case TIM_Scroll:
		lRet = psc->OnScroll( wParam, lParam );
		break;

	case TIM_SetContext:
		lRet = psc->OnSetContext( wParam, lParam );
		break;

	case TIM_SetLinesPerRow:
		lRet = psc->OnSetLinesPerRow( wParam, lParam );
		break;

	case TIM_SetLockedColumns:
		lRet = psc->OnSetLockedColumns( wParam, lParam );
		break;

	case TIM_SetRowFlags:
		lRet = psc->OnSetRowFlags( wParam, lParam );
		break;

	case TIM_SetFocusCell:
		if ( psc->m_lNoSetFocusCellProcessing < 1 )
			lRet = psc->OnSetFocusCell( wParam, lParam );
		break;

	case TIM_SetFocusRow:
		lRet = psc->OnSetFocusRow( wParam, lParam );
		break;

	case TIM_SetTableFlags:
		lRet = psc->OnSetTableFlags( wParam, lParam );
		break;

	case TIM_SwapRows:
		lRet = psc->OnSwapRows( wParam, lParam );
		break;

	default:
		if ( uMsg == MTM_CollapseRow )
		{
			pRow = psc->m_Rows->GetRowPtr( lParam );
			lRet = psc->CallOldWndProc( uMsg, wParam, lParam );
			if ( lRet != MTBL_NODEFAULTACTION && pRow )
				psc->CollapseRow( pRow, wParam );
			SendMessage( psc->m_hWndTbl, MTM_CollapseRowDone, wParam, lParam );
		}
		else if ( uMsg == MTM_ExpandRow )
		{
			pRow = psc->m_Rows->GetRowPtr( lParam );
			lRet = psc->CallOldWndProc( uMsg, wParam, lParam );
			if ( lRet != MTBL_NODEFAULTACTION && pRow )
				psc->ExpandRow( pRow, wParam );
			SendMessage( psc->m_hWndTbl, MTM_ExpandRowDone, wParam, lParam );
		}
		else if ( uMsg == MTM_Internal )
		{
			lRet = psc->OnInternal( wParam, lParam );
		}
		else
		{
			// Alte Fensterprozedur aufrufen
			lRet = psc->CallOldWndProc( uMsg, wParam, lParam );
		}
	}	

	return lRet;
}
