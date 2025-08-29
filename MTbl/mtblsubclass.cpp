#include "mtblsubclass.h"

extern BOOL gbCtdHooked;
extern HANDLE ghInstance;
extern LanguageMap glm;
extern SubClassMap gscm;
extern UINT guiMsgOffs;
extern tstring gsEmpty;

// Funktionscontainer
extern CCtd Ctd;

// Funktionsvariablen
extern FLATSB_SETSCROLLPOS FlatSB_SetScrollPos_Hooked;
extern FLATSB_SETSCROLLRANGE FlatSB_SetScrollRange_Hooked;
extern INVERTRECT InvertRect_Hooked;
extern PATBLT PatBlt_Hooked;
extern SCROLLWINDOW ScrollWindow_Hooked;
extern SETSCROLLPOS SetScrollPos_Hooked;
extern SETSCROLLRANGE SetScrollRange_Hooked;

// MTblSortCompare
int MTblSortCompare(const void *lpv1, const void *lpv2)
{
	SORTROW *pr1 = (SORTROW*)lpv1;
	SORTROW *pr2 = (SORTROW*)lpv2;

	CMTblSortVal *psv1 = pr1->pSortVals;
	CMTblSortVal *psv2 = pr2->pSortVals;

	SORTCOLINFO *psi1 = pr1->pSortColInfos;
	SORTCOLINFO *psi2 = pr2->pSortColInfos;

	int iRet;
	int iTo = min(pr1->lCols, pr2->lCols);
	for (int i = 0; i < iTo; ++i)
	{
		iRet = psv1[i].Compare(psv2[i], psi1[i].bIgnoreCase, psi1[i].bStringSort);
		if (iRet)
		{
			if (psi1[i].iSortOrder == TBL_SortDecreasing)
				iRet *= -1;
			return iRet;
		}
	}

	return 0;
}

// ********************
// ****MTBLSUBCLASS****
// ********************

HHOOK MTBLSUBCLASS::m_hHookEditStyle = NULL;
BOOL MTBLSUBCLASS::m_bHookEditMultiline = FALSE;
int MTBLSUBCLASS::m_iHookEditJustify = -1;
void * MTBLSUBCLASS::m_pscHookEditStyle = NULL;

// AdaptCellRange
BOOL MTBLSUBCLASS::AdaptCellRange(MTBLCELLRANGE &cr)
{
	// Range prüfen
	if (!cr.IsValid())
		return FALSE;

	// Range normalisieren
	cr.Normalize();

	// Spaltenpositionen
	int iColPosFrom = Ctd.TblQueryColumnPos(cr.hWndColFrom);
	int iColPosTo = Ctd.TblQueryColumnPos(cr.hWndColTo);

	// Durchlaufen...
	BOOL bRepeat;
	HWND hWndCol;
	int iColPos;
	long lRow;
	LPMTBLMERGECELL pmc;

	// Sichtbare Spalten ermitteln
	HWND hWndaVisCols[MTBL_MAX_COLS];
	for (iColPos = 1; hWndCol = Ctd.TblGetColumnWindow(m_hWndTbl, iColPos, COL_GetPos); iColPos++)
	{
		if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible))
			hWndaVisCols[iColPos] = hWndCol;
		else
			hWndaVisCols[iColPos] = FALSE;
	}

	// Auf geht's...
	do
	{
		bRepeat = FALSE;

		lRow = cr.lRowFrom;
		while (lRow <= cr.lRowTo)
		{
			iColPos = iColPosFrom;
			while (iColPos <= iColPosTo)
			{
				hWndCol = hWndaVisCols[iColPos];
				if (hWndCol)
				{
					pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);
					if (pmc)
					{
						if (pmc->iColPosFrom < iColPosFrom)
						{
							iColPosFrom = pmc->iColPosFrom;
							bRepeat = TRUE;
						}

						if (pmc->iColPosTo > iColPosTo)
						{
							iColPosTo = pmc->iColPosTo;
							bRepeat = TRUE;
						}

						if (pmc->lRowFrom < cr.lRowFrom)
						{
							cr.lRowFrom = pmc->lRowFrom;
							bRepeat = TRUE;
						}

						if (pmc->lRowTo > cr.lRowTo)
						{
							cr.lRowTo = pmc->lRowTo;
							bRepeat = TRUE;
						}
					}

					if (bRepeat)
						break;
				}

				if (iColPos == iColPosFrom && iColPosTo > iColPosFrom && lRow != cr.lRowFrom && lRow != cr.lRowTo)
					iColPos = iColPosTo;
				else
					iColPos++;
			}

			if (bRepeat)
				break;

			if (!Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, ROW_Hidden))
				break;
		}
	} while (bRepeat);


	cr.hWndColFrom = Ctd.TblGetColumnWindow(m_hWndTbl, iColPosFrom, COL_GetPos);
	cr.hWndColTo = Ctd.TblGetColumnWindow(m_hWndTbl, iColPosTo, COL_GetPos);

	return TRUE;
}

// AdaptColHdrText
BOOL MTBLSUBCLASS::AdaptColHdrText(HWND hWndCol)
{
	BOOL bModified = FALSE, bOk = FALSE;

	// Text ermitteln
	tstring sText(_T(""));
	if (GetColHdrText(hWndCol, sText))
	{
		// Anzahl CRLFs ermitteln
		int iCRLF = GetCRLFCount(sText.c_str());

		// Anzahl anhängende CRLFs ermitteln
		int iTrCRLF = GetTrailingCRLFCount(sText.c_str());

		// Gruppe ermitteln
		CMTblColHdrGrp * pGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);

		// Schon alles in Ordnung?
		int iLines;
		if (pGrp)
		{
			// Anzahl Textlinien ermitteln
			iLines = pGrp->GetNrOfTextLines();

			if (pGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS) ? iLines == iCRLF + 1 : iLines == iTrCRLF) return TRUE;
		}

		// Anhängende CRLFs entfernen
		if (iTrCRLF > 0)
		{
			sText = sText.substr(0, sText.size() - (iTrCRLF * 2));
			iCRLF -= iTrCRLF;
			bModified = TRUE;
		}

		if (pGrp && iLines > 0)
		{
			// CRLFs anhängen
			int iNeeded;
			if (pGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS))
				iNeeded = iLines - iCRLF - 1;
			else
				iNeeded = iLines;
			TCHAR szCRLF[3] = { 13, 10, 0 };
			for (int i = 0; i < iNeeded; ++i)
			{
				sText += szCRLF;
			}
			bModified = TRUE;
		}

		// Neuen Text setzen
		if (bModified)
		{
			bOk = Ctd.TblSetColumnTitle(hWndCol, (LPTSTR)sText.c_str());
		}
		else
			bOk = TRUE;
	}

	return bOk;
}


// AdaptListClip
void MTBLSUBCLASS::AdaptListClip()
{
	if (!m_hWndListClip || m_bAdaptingListClip)
		return;

	// Akt. Focus ermitteln
	HWND hWndCol = NULL;

	long lRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol);

	if (hWndCol)
	{
		// Metriken ermitteln
		CMTblMetrics tm(m_hWndTbl);

		// Absolutes Rechteck der Zelle ermitteln
		RECT rCell;
		//int iCellVis = MTblGetCellRectEx( hWndCol, lRow, (LPINT)&rCell.left, (LPINT)&rCell.top, (LPINT)&rCell.right, (LPINT)&rCell.bottom, MTGR_INCLUDE_MERGED );
		int iCellVis = GetCellRect(hWndCol, lRow, &tm, &rCell, MTGR_INCLUDE_MERGED);
		if (iCellVis == MTGR_RET_ERROR)
			return;

		m_bAdaptingListClip = TRUE;

		long lRowsTop = (lRow < 0 ? tm.m_SplitRowsTop : tm.m_RowsTop);
		long lRowsBottom = (lRow < 0 ? tm.m_SplitRowsBottom : tm.m_RowsBottom);
		long lColsLeft = (IsLockedColumn(hWndCol) ? tm.m_LockedColumnsLeft : tm.m_ColumnsLeft);

		// Zelltyp ermitteln
		int iCellType = GetEffCellType(lRow, hWndCol);

		// Buttons außerhalb der Zelle?
		BOOL bButtonsOutsideCell = QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL);

		// Edit-Größe automatisch anpassen?
		BOOL bAutoSizeEdit = QueryCellFlags(lRow, hWndCol, MTBL_CELL_FLAG_AUTOSIZE_EDIT) || QueryColumnFlags(hWndCol, MTBL_COL_FLAG_AUTOSIZE_EDIT) || QueryRowFlags(lRow, MTBL_ROW_AUTOSIZE_EDIT);

		// Minimale Breite Edit-Bereich
		int iMinEditAreaWid = 0;
		if (iCellType == COL_CellType_CheckBox)
			iMinEditAreaWid = CMTblMetrics::QueryCheckBoxSize(m_hWndTbl);
		/*else
		iMinEditAreaWid = iCharWid;*/

		// M!Table-Button ermitteln
		BOOL bBtnLeft = FALSE;
		CMTblBtn *pBtn = NULL;
		int iBtnWid = 0;
		if (MustShowBtn(hWndCol, lRow))
		{
			pBtn = GetEffCellBtn(hWndCol, lRow);
			if (pBtn)
			{
				iBtnWid = CalcBtnWidth(pBtn, hWndCol, lRow);
				bBtnLeft = pBtn->IsLeft() && !bButtonsOutsideCell;
			}
		}

		// Nicht editierbaren Bereich auschneiden?
		int iLeftOffset = 0, iRightOffset = 0;
		BOOL bClipNonedit = QueryCellFlags(lRow, hWndCol, MTBL_CELL_FLAG_CLIP_NONEDIT) || QueryColumnFlags(hWndCol, MTBL_COL_FLAG_CLIP_NONEDIT) || QueryRowFlags(lRow, MTBL_ROW_CLIP_NONEDIT);
		if (bClipNonedit)
		{
			CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);
			//iLeftOffset = GetCellTextAreaLeft( lRow, hWndCol, rCell.left, tm.m_CellLeadingX, m_Rows->GetColImagePtr( lRow, hWndCol ), bBtnLeft ) - rCell.left;
			int iLeft = GetCellTextAreaLeft(lRow, hWndCol, rCell.left, tm.m_CellLeadingX, Img, bBtnLeft);
			if (iLeft + iMinEditAreaWid <= rCell.right)
				iLeftOffset = iLeft - rCell.left;
			else
				bClipNonedit = FALSE;
			//iRightOffset = rCell.right - GetCellTextAreaRight( lRow, hWndCol, rCell.right, tm.m_CellLeadingX, m_Rows->GetColImagePtr( lRow, hWndCol ) );
			int iRight = GetCellTextAreaRight(lRow, hWndCol, rCell.right, tm.m_CellLeadingX, Img);
			if (iRight - iMinEditAreaWid >= rCell.left)
				iRightOffset = rCell.right - iRight;
			else
				bClipNonedit = FALSE;
		}
		else if (bBtnLeft)
		{
			iLeftOffset = iBtnWid;
		}

		// Ränder
		WORD wLM = tm.m_CellLeadingX, wRM = tm.m_CellLeadingX;

		// Zelle Read-Only?
		BOOL bCellReadOnly = (IsCellReadOnly(hWndCol, lRow));

		// Edit: GDI-Objekte
		if (IsWindow(m_hWndEdit) && (m_hWndEdit != m_hWndLastEdit || m_bForceColorEdit || m_bForceEditFont))
		{
			// Font
			if (m_hWndEdit != m_hWndLastEdit || m_bForceEditFont)
			{
				// Alten evtl. löschen
				if (m_hEditFont && m_bDeleteEditFont)
					DeleteObject(m_hEditFont);
				// Zurücksetzen
				m_hEditFont = NULL;
				m_bDeleteEditFont = FALSE;

				// Erzeugen
				HDC hDC = GetDC(NULL);
				if (hDC)
				{
					DWORD dwFlags = 0;
					if (AnyHighlightedItems())
						dwFlags |= GEF_HILI;

					m_hEditFont = GetEffCellFont(hDC, lRow, hWndCol, &m_bDeleteEditFont, dwFlags);

					ReleaseDC(NULL, hDC);
				}
			}

			// Textfarbe/Hintergrundpinsel
			if (m_hWndEdit != m_hWndLastEdit || m_bForceColorEdit)
			{
				DWORD dwFlags = 0;
				if (AnyHighlightedItems())
					dwFlags |= GEC_HILI;

				// Textfarbe ermitteln
				m_clrEditText = GetEffCellTextColor(lRow, hWndCol, dwFlags);

				// Hintergrundfarbe
				m_clrEditBack = GetEffCellBackColor(lRow, hWndCol, dwFlags);

				// Brush
				if (m_hEditBrush)
				{
					DeleteObject(m_hEditBrush);
					m_hEditBrush = NULL;
				}
				m_hEditBrush = CreateSolidBrush(m_clrEditBack);
			}

			// Dieses Edit als letztes Edit merken
			m_hWndLastEdit = m_hWndEdit;
		}

		// Handles der relevanten Kindobjekte ermitteln
		HWND hWndDDBtn = FindFirstChildClass(m_hWndListClip, CLSNAME_DROPDOWNBTN);
		HWND hWndContainer = FindFirstChildClass(m_hWndListClip, CLSNAME_CONTAINER);
		HWND hWndCheckBox = FindFirstChildClass(m_hWndListClip, CLSNAME_CHECKBOX);
		HWND hWndBtn = FindFirstChildClass(hWndContainer ? hWndContainer : m_hWndListClip, MTBL_BTN_CLASS);

		// Höhe einer Textzeile / Breite eines Zeichens ermitteln
		int iCharWid = 0;
		int iTextLineHt = 0;
		HDC hDC = GetDC(NULL);
		if (hDC)
		{
			if (m_hEditFont)
			{
				HGDIOBJ hOldFont = SelectObject(hDC, m_hEditFont);
				SIZE s = { 0, 0 };
				GetTextExtentPoint32(hDC, _T("A"), 1, &s);
				iTextLineHt = s.cy;
				iCharWid = s.cy;
				SelectObject(hDC, hOldFont);
			}
			ReleaseDC(NULL, hDC);
		}
		if (!iTextLineHt)
			iTextLineHt = tm.m_CharBoxY;
		if (!iCharWid)
			iCharWid = tm.m_CharBoxX;

		// Benötigte logische Höhe ermitteln
		BOOL bWordWrap = GetCellWordWrap(hWndCol, lRow);

		int iRowHt = max(GetEffRowHeight(lRow, &tm), rCell.bottom - rCell.top + 1) - (tm.m_CellLeadingY * 2);

		int iHt, iTextLinesHt;
		if ( /*bAutoSizeEdit &&*/ bWordWrap && iCellType != COL_CellType_PopupEdit && iCellType != COL_CellType_CheckBox)
		{
			SIZE s = { 0, 0 };
			if (bAutoSizeEdit && GetCellTextExtent(hWndCol, lRow, &s, m_hEditFont, TRUE) && s.cy > 0)
				iTextLinesHt = s.cy;
			else
			{
				int iLines = max(iRowHt / iTextLineHt, 1);
				iTextLinesHt = iLines * iTextLineHt;
			}
		}
		else
			iTextLinesHt = iTextLineHt;

		iHt = max(iTextLinesHt, iRowHt);

		if (rCell.top + tm.m_CellLeadingY + iHt > lRowsBottom)
		{
			iHt = lRowsBottom - rCell.top - tm.m_CellLeadingY;
			int iLines = max(iHt / iTextLineHt, 1);
			iTextLinesHt = (iLines * iTextLineHt);
		}

		// Benötigte logische Breite ermitteln
		int iWid = 0;
		if (bAutoSizeEdit && !bWordWrap && iCellType != COL_CellType_PopupEdit && iCellType != COL_CellType_CheckBox && !iRightOffset && !hWndDDBtn && !(pBtn && !pBtn->IsLeft()))
		{
			iWid = GetBestCellWidth(hWndCol, lRow, tm.m_CellLeadingX, m_hEditFont);
			/*if ( hWndDDBtn )
			iWid += MTBL_DDBTN_WID;*/
			if (pBtn)
				iWid += iBtnWid;
			if (!bClipNonedit)
			{
				//int iSubtract = GetCellTextAreaLeft( lRow, hWndCol, rCell.left, tm.m_CellLeadingX, m_Rows->GetColImagePtr( lRow, hWndCol ), bBtnLeft ) - rCell.left;
				CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);
				int iSubtract = GetCellTextAreaLeft(lRow, hWndCol, rCell.left, tm.m_CellLeadingX, Img, bBtnLeft) - rCell.left;
				iSubtract -= (pBtn ? iBtnWid : 0);
				if (iSubtract > 0)
					iWid -= iSubtract;
			}

			iWid = max(iWid, (rCell.right - rCell.left));
		}
		else
		{
			iWid = rCell.right - rCell.left;

			if (bButtonsOutsideCell)
			{
				if (hWndDDBtn)
					iWid += MTBL_DDBTN_WID;
				if (pBtn && !bBtnLeft)
					iWid += iBtnWid;
			}
		}

		// Logisches Rechteck List-Clip
		RECT rLCLog;
		rLCLog.left = rCell.left;
		rLCLog.top = rCell.top + tm.m_CellLeadingY;
		rLCLog.right = rLCLog.left + iWid;
		rLCLog.bottom = rLCLog.top + iHt;

		// Logische Breite/Höhe List-Clip
		int iLCLogWid = rLCLog.right - rLCLog.left;
		int iLCLogHt = rLCLog.bottom - rLCLog.top;

		// Aktuelles Rechteck List-Clip
		RECT rLCCur;
		GetWindowRect(m_hWndListClip, &rLCCur);
		MapWindowPoints(NULL, m_hWndTbl, (LPPOINT)&rLCCur, 2);

		// Tatsächliches Rechteck List-Clip
		BOOL bInvisible = FALSE;
		RECT rLC;

		rLC.left = max(rLCLog.left, lColsLeft);
		rLC.top = max(rLCLog.top, lRowsTop);
		rLC.right = rLCLog.right;
		rLC.bottom = min(rLCLog.bottom, lRowsBottom);
		if (rLC.right <= rLC.left || rLC.bottom <= rLC.top)
		{
			bInvisible = TRUE;
			SetRect(&rLC, rLCCur.left, rLCCur.top, rLCCur.right, rLCCur.top);
		}

		int iLCLeftOffs = rLC.left - rLCLog.left;
		int iLCTopOffs = rLCLog.top - rLC.top;

		// Tatsächliche Breite/Höhe List-Clip
		int iLCWid = rLC.right - rLC.left;
		int iLCHt = rLC.bottom - rLC.top;

		// Aktuelle Breite/Höhe List-Clip
		int iLCCurWid = rLCCur.right - rLCCur.left;
		int iLCCurHt = rLCCur.bottom - rLCCur.top;

		// Position/Größe List-Clip setzen
		if (!EqualRect(&rLC, &rLCCur))
		{
			DWORD dwFlags = SWP_NOACTIVATE | SWP_NOZORDER;
			if (!bInvisible && !IsWindowVisible(m_hWndListClip))
				dwFlags |= SWP_SHOWWINDOW;
			if (rLCCur.left == rLC.left && rLCCur.top == rLC.top)
				dwFlags |= SWP_NOMOVE;
			if (iLCWid == iLCCurWid && iLCHt == iLCCurHt)
				dwFlags |= SWP_NOSIZE;

			SetWindowPos(m_hWndListClip, NULL, rLC.left, rLC.top, iLCWid, iLCHt, dwFlags);
		}

		// Positionen/Größen der Kindobjekte anpassen, evtl. Button erzeugen
		if (iLCWid > 0)
		{
			// Client-Rechteck des List-Clips ermitteln
			RECT rLCClient;
			GetClientRect(m_hWndListClip, &rLCClient);

			// Drop-Down-Button
			int iDDBtnWid = 0;
			int iDDBtnHt = 0;
			if (hWndDDBtn)
			{
				iDDBtnWid = MTBL_DDBTN_WID;
				iDDBtnHt = iTextLineHt;

				RECT r = { rLCLog.right - MTBL_DDBTN_WID, rLCLog.top, rLCLog.right, rLCLog.top + iTextLineHt };
				MapWindowPoints(m_hWndTbl, m_hWndListClip, (LPPOINT)&r, 2);

				RECT rCur;
				GetWindowRect(hWndDDBtn, &rCur);
				MapWindowPoints(NULL, m_hWndListClip, (LPPOINT)&rCur, 2);

				if (!EqualRect(&r, &rCur))
					SetWindowPos(hWndDDBtn, NULL, r.left, r.top, r.right - r.left, r.bottom - r.top, SWP_NOACTIVATE | SWP_NOZORDER);
			}

			// M!Table-Button: Höhe, Font, Position links
			HFONT hBtnFont = NULL;
			int iBtnHt = 0, iBtnLeft = 0;
			//m_hBtnFont = NULL;
			if (pBtn)
			{
				// Höhe ermitteln
				//m_iBtnHt = iTextLineHt;
				//iBtnHt = iTextLineHt;
				iBtnHt = iTextLineHt + 4;

				SIZE siImg = { 0, 0 };
				if (MImgIsValidHandle(pBtn->m_hImage))
				{
					MImgGetSize(pBtn->m_hImage, &siImg);
					//m_iBtnHt = min( rLCClient.bottom - rLCClient.top, siImg.cy + 4 );
					iBtnHt = max(siImg.cy + 4, iBtnHt);

				}
				//iBtnHt = m_iBtnHt;
				iBtnHt = min(rLCClient.bottom - rLCClient.top, iBtnHt);

				// Font
				//m_hBtnFont = (pBtn->m_dwFlags & MTBL_BTN_SHARE_FONT) ? m_hEditFont : (HFONT)SendMessage( m_hWndTbl, WM_GETFONT, 0, 0 );
				hBtnFont = (pBtn->m_dwFlags & MTBL_BTN_SHARE_FONT) ? m_hEditFont : (HFONT)SendMessage(m_hWndTbl, WM_GETFONT, 0, 0);

				// Position links
				if (bBtnLeft)
					iBtnLeft = bClipNonedit ? GetCellAreaLeft(lRow, hWndCol, rLCLog.left) : rLCLog.left;
			}

			// Offsets anpassen
			/*if ( iLeftOffset || iRightOffset )
			{
			int iWid = iLCWid - iBtnWid - iDDBtnWid;

			if ( iLeftOffset > iWid )
			iLeftOffset = iWid;

			if ( ( iWid - iLeftOffset - iRightOffset ) < iMinEditAreaWid )
			iRightOffset = max( 0, iWid - iLeftOffset - iMinEditAreaWid );
			}*/

			// Edit: Rechteck, Font, Passwort-Zeichen, Ränder, Read-Only
			if (IsWindow(m_hWndEdit))
			{
				// Edit-Ränder
				long lStyle = GetWindowLongPtr(m_hWndEdit, GWL_STYLE);
				if (lStyle & ES_RIGHT)
					wRM = max(wRM - 1, 0);
				else if (lStyle & ES_CENTER)
					wLM = wRM = 0;

				if (iLeftOffset)
					wLM += iLeftOffset;

				if (iRightOffset)
					wRM += iRightOffset;

				// Rechteck
				RECT r;

				int iEditWid = iLCLogWid - ((bBtnLeft ? 0 : iBtnWid) + iDDBtnWid);

				int iEditHt = iTextLinesHt;
				if (iCellType == COL_CellType_PopupEdit)
					iEditHt = tm.m_CharBoxY;

				r.left = rLCLog.left;
				r.top = rLCLog.top;
				r.right = r.left + iEditWid;
				r.bottom = r.top + iEditHt;
				MapWindowPoints(m_hWndTbl, m_hWndListClip, (LPPOINT)&r, 2);

				RECT rCur;
				GetWindowRect(m_hWndEdit, &rCur);
				MapWindowPoints(NULL, m_hWndListClip, (LPPOINT)&rCur, 2);

				if (!EqualRect(&r, &rCur))
					SetWindowPos(m_hWndEdit, NULL, r.left, r.top, iEditWid, iEditHt, SWP_NOACTIVATE | SWP_NOZORDER);

				// Font
				if (m_hEditFont && m_hEditFont != (HFONT)SendMessage(m_hWndEdit, WM_GETFONT, 0, 0))
					SendMessage(m_hWndEdit, WM_SETFONT, WPARAM(m_hEditFont), MAKELPARAM(FALSE, 0));

				// Passwort-Zeichen
				TCHAR cPwd = (TCHAR)SendMessage(m_hWndEdit, EM_GETPASSWORDCHAR, 0, 0);
				if (cPwd && cPwd != m_cPassword)
					SendMessage(m_hWndEdit, EM_SETPASSWORDCHAR, (WPARAM)m_cPassword, 0);

				// Edit-Rect bzw. Ränder
				if (lStyle & ES_MULTILINE)
				{
					RECT rEdit;
					rEdit.left = wLM;
					rEdit.top = 0;
					rEdit.right = iEditWid - wRM;
					rEdit.bottom = iEditHt;
					SendMessage(m_hWndEdit, EM_SETRECTNP, 0, (LPARAM)&rEdit);
				}
				else
				{
					SendMessage(m_hWndEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELONG(wLM, wRM));
				}

				// Read-Only
				BOOL bIsReadOnly = (GetWindowLongPtr(m_hWndEdit, GWL_STYLE) & ES_READONLY) != 0;
				if (bCellReadOnly != bIsReadOnly)
					SendMessage(m_hWndEdit, EM_SETREADONLY, bCellReadOnly, 0);
			}

			// Container
			if (hWndContainer)
			{
				// Rechteck
				RECT r;

				int iContainerWid = iLCLogWid;
				int iContainerHt = iLCLogHt;

				r.left = rLCLog.left;
				r.top = rLCLog.top;
				r.right = r.left + iContainerWid;
				r.bottom = r.top + iContainerHt;
				MapWindowPoints(m_hWndTbl, m_hWndListClip, (LPPOINT)&r, 2);

				RECT rCur;
				GetWindowRect(hWndContainer, &rCur);
				MapWindowPoints(NULL, m_hWndListClip, (LPPOINT)&rCur, 2);

				if (!EqualRect(&r, &rCur))
					SetWindowPos(hWndContainer, NULL, r.left, r.top, iContainerWid, iContainerHt, SWP_NOACTIVATE | SWP_NOZORDER);
			}

			// Checkbox
			if (hWndCheckBox && hWndContainer)
			{
				RECT rCell;
				int iRet = GetCellRect(hWndCol, lRow, &tm, &rCell, MTGR_INCLUDE_MERGED);

				if (iRet != MTGR_RET_ERROR)
				{
					int iJustify = GetEffCellTextJustify(lRow, hWndCol);
					int iVAlign = GetEffCellTextVAlign(lRow, hWndCol);

					RECT rCbArea = { 0, rCell.top, 0, rCell.bottom };

					CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);
					rCbArea.left = GetCellTextAreaLeft(lRow, hWndCol, rCell.left, tm.m_CellLeadingX, Img);
					rCbArea.right = GetCellTextAreaRight(lRow, hWndCol, rCell.right, tm.m_CellLeadingX, Img, iJustify != DT_CENTER);

					RECT r;
					CalcCheckBoxRect(iJustify, iVAlign, rCbArea, tm.m_CellLeadingX, tm.m_CellLeadingY, tm.m_CharBoxY, r);
					MapWindowPoints(m_hWndTbl, hWndContainer, (LPPOINT)&r, 2);

					RECT rCur;
					GetWindowRect(hWndCheckBox, &rCur);
					MapWindowPoints(NULL, hWndContainer, (LPPOINT)&rCur, 2);

					long lSize = r.bottom - r.top;

					if (!EqualRect(&r, &rCur))
						SetWindowPos(hWndCheckBox, NULL, r.left, r.top, lSize, lSize, SWP_NOACTIVATE | SWP_NOZORDER);
				}
			}

			// M!Table-Button erzeugen/anpassen/zerstören
			if (pBtn)
			{
				// Elternfenster für Button ermitteln
				HWND hWndParent = (hWndContainer ? hWndContainer : m_hWndListClip);

				// Rechteck ermitteln
				RECT r;
				if (bBtnLeft)
				{
					r.left = iBtnLeft;
				}
				else
				{
					r.left = rLCLog.right - iBtnWid;
					if (hWndDDBtn)
						r.left -= MTBL_DDBTN_WID;
				}
				r.top = rLCLog.top;
				r.right = r.left + iBtnWid;
				//r.bottom = r.top + m_iBtnHt;
				r.bottom = r.top + iBtnHt;
				MapWindowPoints(m_hWndTbl, hWndParent, (LPPOINT)&r, 2);

				// Erzeugen...
				if (!hWndBtn)
				{
					DWORD dwStyle = WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VISIBLE;
					if (pBtn->m_dwFlags & MTBL_BTN_DISABLED)
						dwStyle |= WS_DISABLED;

					hWndBtn = m_hWndBtn = CreateWindowEx(0, MTBL_BTN_CLASS, pBtn->m_sText.c_str(), dwStyle, r.left, r.top, r.right - r.left, r.bottom - r.top, hWndParent, HMENU(MTBL_BTN_ID), HINSTANCE(ghInstance), NULL);
					if (hWndBtn)
					{
						// Button-Daten setzen
						SetBtnWndData(pBtn, iBtnHt);

						// Über Edit positionieren
						if (m_hWndEdit)
							SetWindowPos(hWndBtn, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
					}
				}
				// ... oder anpassen
				else if (hWndBtn)
				{
					// Button-Daten
					SetBtnWndData(pBtn, iBtnHt);

					// Größe/Position
					RECT rCur;
					GetWindowRect(hWndBtn, &rCur);
					MapWindowPoints(NULL, GetParent(hWndBtn), (LPPOINT)&rCur, 2);

					if (!EqualRect(&r, &rCur))
						SetWindowPos(hWndBtn, NULL, r.left, r.top, r.right - r.left, r.bottom - r.top, SWP_NOACTIVATE | SWP_NOZORDER);

					// Text
					int iTextLen = GetWindowTextLength(hWndBtn);
					TCHAR *pText = new TCHAR[iTextLen + 1];
					pText[0] = 0;
					GetWindowText(hWndBtn, pText, iTextLen + 1);
					if (lstrcmp(pText, pBtn->m_sText.c_str()))
						SendMessage(hWndBtn, WM_SETTEXT, 0, (LPARAM)pBtn->m_sText.c_str());
					delete[]pText;

					// Enabled
					BOOL bBtnEnabled = IsWindowEnabled(hWndBtn);
					BOOL bEnable = !(pBtn->m_dwFlags & MTBL_BTN_DISABLED);
					if (bEnable != bBtnEnabled)
						EnableWindow(hWndBtn, bEnable);
				}
			}
			else if (hWndBtn)
			{
				DestroyWindow(hWndBtn);
				hWndBtn = NULL;
			}

			// Region setzen
			HRGN hRgnNew = NULL;

			if (QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL) && (hWndBtn || hWndDDBtn))
			{
				HRGN hRgn1 = CreateRectRgn(0, 0, iLCWid, iLCHt);
				if (hRgn1)
				{
					int iBtnsWid = iDDBtnWid;
					if (!bBtnLeft)
						iBtnsWid += iBtnWid;
					int iLeft = iLCWid - iBtnsWid;

					int iBtnsHt = max(iBtnHt, iDDBtnHt);
					int iBtnOffs = rLCLog.top - rLC.top;
					int iTop = max(0, iBtnOffs + iBtnsHt);

					HRGN hRgn2 = CreateRectRgn(iLeft, iTop, iLCWid, iLCHt);
					if (hRgn2)
					{
						if (hRgnNew = CreateRectRgn(0, 0, 0, 0))
						{
							if (CombineRgn(hRgnNew, hRgn1, hRgn2, RGN_DIFF) == ERROR)
							{
								DeleteObject(hRgnNew);
								hRgnNew = NULL;
							}
						}

						DeleteObject(hRgn2);
					}

					DeleteObject(hRgn1);
				}
			}

			if (iLeftOffset)
			{
				HRGN hRgn1 = CreateRectRgn(0, 0, iLCWid, iLCHt);
				if (hRgnNew && hRgn1)
				{
					if (CombineRgn(hRgn1, hRgnNew, NULL, RGN_COPY) != ERROR)
					{
						DeleteObject(hRgnNew);
						hRgnNew = NULL;
					}
				}

				if (hRgn1)
				{
					HRGN hRgn2 = CreateRectRgn(0, 0, iLeftOffset - iLCLeftOffs, iLCHt);
					if (hRgn2)
					{
						if (hRgnNew = CreateRectRgn(0, 0, 0, 0))
						{
							if (CombineRgn(hRgnNew, hRgn1, hRgn2, RGN_DIFF) == ERROR)
							{
								DeleteObject(hRgnNew);
								hRgnNew = NULL;
							}
						}

						DeleteObject(hRgn2);
					}

					DeleteObject(hRgn1);
				}

				if (bBtnLeft && hWndBtn && hRgnNew)
				{
					// Linken Button wieder im List-Clip sichtbar machen
					HRGN hRgn1 = CreateRectRgn(0, 0, 0, 0);
					if (hRgn1)
					{
						if (CombineRgn(hRgn1, hRgnNew, NULL, RGN_COPY) != ERROR)
						{
							DeleteObject(hRgnNew);
							hRgnNew = NULL;
						}

						RECT r;
						GetWindowRect(hWndBtn, &r);
						MapWindowPoints(NULL, m_hWndListClip, (LPPOINT)&r, 2);

						HRGN hRgn2 = CreateRectRgn(max(r.left, 0), max(r.top, 0), r.right, r.bottom);
						if (hRgn2)
						{
							if (hRgnNew = CreateRectRgn(0, 0, 0, 0))
							{
								if (CombineRgn(hRgnNew, hRgn1, hRgn2, RGN_OR) == ERROR)
								{
									DeleteObject(hRgnNew);
									hRgnNew = NULL;
								}
							}

							DeleteObject(hRgn2);
						}

						DeleteObject(hRgn1);
					}
				}
			}

			if (iRightOffset)
			{
				HRGN hRgn1 = CreateRectRgn(0, 0, iLCWid, iLCHt);
				if (hRgnNew && hRgn1)
				{
					if (CombineRgn(hRgn1, hRgnNew, NULL, RGN_COPY) != ERROR)
					{
						DeleteObject(hRgnNew);
						hRgnNew = NULL;
					}
				}

				if (hRgn1)
				{
					int iBtnsWid = /*iBtnWid +*/ iDDBtnWid;
					if (!bBtnLeft)
						iBtnsWid += iBtnWid;
					HRGN hRgn2 = CreateRectRgn(iLCWid - iBtnsWid - iRightOffset, 0, iLCWid - iBtnsWid, iLCHt);
					if (hRgn2)
					{
						if (hRgnNew = CreateRectRgn(0, 0, 0, 0))
						{
							if (CombineRgn(hRgnNew, hRgn1, hRgn2, RGN_DIFF) == ERROR)
							{
								DeleteObject(hRgnNew);
								hRgnNew = NULL;
							}
						}

						DeleteObject(hRgn2);
					}

					DeleteObject(hRgn1);
				}
			}

			if (hRgnNew != m_hRgnListClip)
				SetWindowRgn(m_hWndListClip, hRgnNew, TRUE);
			if (m_hRgnListClip)
				DeleteObject(m_hRgnListClip);
			m_hRgnListClip = hRgnNew;

			// Neu zeichnen
			if (IsWindowVisible(m_hWndListClip))
				RedrawWindow(m_hWndListClip, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN);
		}

		// Popup-Edit/Listbox: Breite, Position, Font, Read-Only
		HWND hWndExt = NULL;
		if (iCellType == COL_CellType_PopupEdit || iCellType == COL_CellType_DropDownList)
			hWndExt = (HWND)SendMessage(m_hWndTbl, TIM_GetExtEdit, Ctd.TblQueryColumnID(hWndCol) - 1, 0);
		if (hWndExt)
		{
			// Font
			BOOL bUseTblFont = (iCellType == COL_CellType_PopupEdit && QueryFlags(MTBL_FLAG_USE_TBL_FONT_PE)) ||
				(iCellType == COL_CellType_DropDownList && QueryFlags(MTBL_FLAG_USE_TBL_FONT_DDL));
			HFONT hFont;

			if (m_hEditFont && !bUseTblFont)
				hFont = m_hEditFont;
			else
				hFont = m_hFont;

			SendMessage(hWndExt, WM_SETFONT, WPARAM(hFont), MAKELPARAM(TRUE, 0));

			// Read-Only
			if (iCellType == COL_CellType_PopupEdit)
			{
				BOOL bIsReadOnly = (GetWindowLongPtr(hWndExt, GWL_STYLE) & ES_READONLY) != 0;
				BOOL bReadOnly = bCellReadOnly;

#ifdef TD_POPUP_EDIT_ALWAYS_SHOWN
				BOOL bDisabled = IsCellDisabled(hWndCol, lRow);
				bReadOnly = bReadOnly || bDisabled;
#endif

				if (bReadOnly != bIsReadOnly)
					SendMessage(hWndExt, EM_SETREADONLY, bReadOnly, 0);
			}
			else
			{
				BOOL bIsReadOnly = !IsWindowEnabled(hWndExt);
				if (bCellReadOnly != bIsReadOnly)
					EnableWindow(hWndExt, !bCellReadOnly);
			}

			// Breite
			int iWid = iLCLogWid + 1;
			if (iCellType == COL_CellType_PopupEdit)
				iWid += GetSystemMetrics(SM_CXVSCROLL);

			if (iCellType == COL_CellType_DropDownList && QueryFlags(MTBL_FLAG_ADAPT_LIST_WIDTH))
			{
				int iCount = Ctd.ListQueryCount(hWndCol);
				if (iCount > 0)
				{
					HDC hDC;
					if (hDC = GetDC(hWndExt))
					{
						HGDIOBJ hOldFont = SelectObject(hDC, hFont);
						HSTRING hs = 0;
						int iLen;
						LPTSTR lps;
						long lLen, lMaxTxtWid = 0;
						SIZE siTxt;
						for (int i = 0; i < iCount; i++)
						{
							if ((iLen = Ctd.ListQueryText(hWndCol, i, &hs)) > 0)
							{
								if (lps = Ctd.StringGetBuffer(hs, &lLen))
								{
									if (GetTextExtentPoint32(hDC, lps, iLen, &siTxt))
										lMaxTxtWid = max(lMaxTxtWid, siTxt.cx);
								}
								Ctd.HStrUnlock(hs);
							}
						}
						SelectObject(hDC, hOldFont);

						ReleaseDC(hWndExt, hDC);

						if (hs)
							Ctd.HStringUnRef(hs);

						if (lMaxTxtWid > 0)
						{
							long lBestWid = lMaxTxtWid + (tm.m_CellLeadingX * 2);

							SCROLLINFO si;
							si.cbSize = sizeof(SCROLLINFO);
							si.fMask = SIF_RANGE | SIF_PAGE;
							if (FlatSB_GetScrollInfo(hWndExt, SB_VERT, &si))
							{
								if (si.nMax - si.nMin + 1 > si.nPage)
									lBestWid += GetSystemMetrics(SM_CXVSCROLL);
							}

							int iScrLeft, iScrTop, iScrWid, iScrHt;
							GetScreenMetrics(iScrLeft, iScrTop, iScrWid, iScrHt);

							POINT ptLC = { 0, 0 };
							ClientToScreen(m_hWndListClip, &ptLC);

							if (ptLC.x + lBestWid > iScrLeft + iScrWid)
							{
								lBestWid = (iScrLeft + iScrWid) - ptLC.x;
							}

							if (lBestWid > iWid)
								iWid = lBestWid;
						}
					}
				}
			}

			RECT r, r2;
			GetWindowRect(hWndExt, &r);
			if (iWid != r.right - r.left)
			{
				SetWindowPos(hWndExt, NULL, 0, 0, iWid, r.bottom - r.top, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
				if (iCellType == COL_CellType_DropDownList)
				{
					HWND hWndDropDown = GetParent(hWndExt);
					if (hWndDropDown)
						SetWindowPos(hWndDropDown, NULL, 0, 0, iWid, r.bottom - r.top, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
				}
			}

			if (FALSE)
			{
				// Position setzen lassen
				SendMessage(m_hWndTbl, WM_USER + 128, 0, 0);

				// Position überprüfen und Position ggfs. nochmal setzen lassen
				BOOL bSendPosMsgAgain = FALSE;

				if (iCellType == COL_CellType_DropDownList)
				{
					// Zum  oberen Bildschirmrand "gesprungen"?
					/*GetWindowRect( hWndExt, &r );
					if ( r.top == 0 )
					bSendPosMsgAgain = TRUE;*/

					// "Weggesprungen"?
					GetWindowRect(hWndExt, &r);
					GetWindowRect(m_hWndEdit, &r2);
					if (r.top != r2.bottom && r.bottom != r2.top)
						bSendPosMsgAgain = TRUE;
				}

				if (bSendPosMsgAgain)
				{
					SendMessage(m_hWndTbl, WM_USER + 128, 0, 0);
				}
			}
		}

		m_bAdaptingListClip = FALSE;
	}
}


// AdaptScrollRangeV
void MTBLSUBCLASS::AdaptScrollRangeV(BOOL bSplitBar /*= FALSE*/, int iScrPos /*=TBL_Error*/)
{
	CMTblMetrics tm(m_hWndTbl);

	HWND hWndScr = NULL;
	int iScrCtl;
	if (GetScrollBarV(&tm, bSplitBar, &hWndScr, &iScrCtl))
	{
		int iScrPosBefore = FlatSB_GetScrollPos(hWndScr, iScrCtl);

		int iMaxPos, iMinPos;
		FlatSB_GetScrollRange(hWndScr, iScrCtl, &iMinPos, &iMaxPos);
		OnSetScrollRange(hWndScr, iScrCtl, iMinPos, iMaxPos, TRUE, iScrPos, TRUE);

		FlatSB_GetScrollRange(hWndScr, iScrCtl, &iMinPos, &iMaxPos);
		if (iScrPosBefore > iMaxPos)
		{
			InvalidateRect(m_hWndTbl, NULL, FALSE);
		}
	}

}


// AfterKillEdit
void MTBLSUBCLASS::AfterKillEdit()
{
	m_pCounter->DecNoFocusUpdate();

	HWND hWndCol = NULL;
	long lRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol);

	if (MustSuppressRowFocusCTD())
	{
		//Ctd.TblKillFocus( m_hWndTbl );
		KillFocusFrameI();
		InvalidateRowFocus(lRow, FALSE);
		//UpdateWindow( m_hWndTbl );
	}

	if ((m_bCellMode && !hWndCol) || QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION) /*|| m_bKillEditSuppressSelection*/)
		Ctd.TblSetRowFlags(m_hWndTbl, lRow, ROW_Selected, FALSE);

	if (m_bTrappingRowSelChanges)
	{
		EndRowSelTrap();
		if (!m_bKillEditStartedRowSelTrap)
			BeginRowSelTrap();
	}

	// Update Focus
	if (hWndCol)
		m_bNoFocusPaint = FALSE;
	int iRedrawMode = (m_bCellMode || hWndCol) ? UF_REDRAW_INVALIDATE : UF_REDRAW_AUTO;
	int iSelChangeMode = (m_bCellMode && !hWndCol && !(m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION)) ? UF_SELCHANGE_FOCUS_ADD : UF_SELCHANGE_NONE;
	UpdateFocus(TBL_Error, NULL, iRedrawMode, iSelChangeMode);
}


// AfterScroll
void MTBLSUBCLASS::AfterScroll()
{
	UpdateFocus();
	m_bScrolling = FALSE;
}


// AfterSetFocusCell
void MTBLSUBCLASS::AfterSetFocusCell(HWND hWndCol, long lRow, BOOL bMouseClickSim /*=FALSE*/)
{
	// Prüfen, ob Focus wirklich gesetzt wurde
	HWND hWndFocusCol = NULL;
	long lFocusRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);
	BOOL bFocusSet = (hWndCol == hWndFocusCol && lRow == lFocusRow) && IsMyWindow(GetFocus());

	if (bFocusSet)
	{
		//int iCellType = Ctd.TblGetCellType( hWndCol );
		int iCellType = GetEffCellType(lRow, hWndCol);

		if (bMouseClickSim && iCellType == COL_CellType_Standard && IsWindow(m_hWndEdit))
		{
			// Selektion im Edit korrigieren
			POINT pt;
			GetCursorPos(&pt);

			POINT ptEdit = pt;
			ScreenToClient(m_hWndEdit, &ptEdit);

			RECT r;
			GetWindowRect(m_hWndEdit, &r);

			RECT rEdit;
			SendMessage(m_hWndEdit, EM_GETRECT, 0, (LPARAM)&rEdit);
			r.right = r.left + rEdit.right;

			long lPos = 0;
			if (pt.x >= r.left && pt.x <= r.right)
			{
				//ScreenToClient( m_hWndEdit, &pt );
				lPos = LOWORD(SendMessage(m_hWndEdit, EM_CHARFROMPOS, 0, MAKELPARAM(ptEdit.x, ptEdit.y)));
			}
			else if (pt.x > r.right)
			{
				lPos = SendMessage(m_hWndEdit, WM_GETTEXTLENGTH, 0, 0);
			}

			DWORD dwCurStart, dwCurEnd;
			SendMessage(m_hWndEdit, EM_GETSEL, (WPARAM)&dwCurStart, (LPARAM)&dwCurEnd);
			if (dwCurStart != lPos || dwCurEnd != lPos)
			{
				SendMessage(m_hWndEdit, EM_SETSEL, lPos, lPos);
				SendMessage(m_hWndEdit, EM_SCROLLCARET, 0, 0);
			}

			// WM_LBUTTONDOWN senden (wenn linke Maustaste gedrückt), damit Textmarkierung mit der Maus funktioniert
			SHORT ks = GetKeyState(VK_LBUTTON);
			if (HIBYTE(ks))
				SendMessage(m_hWndEdit, WM_LBUTTONDOWN, 0, MAKELONG(ptEdit.x, ptEdit.y));
		}
		else if (bMouseClickSim && iCellType == COL_CellType_CheckBox)
		{
			// Checkbox "anklicken"
			HWND hWnd = NULL;
			POINT pt;
			GetCursorPos(&pt);

			//hWnd = WindowFromPoint( pt );
			//if ( hWnd && hWnd == FindFirstChildClass( m_hWndListClip, CLSNAME_CHECKBOX ) )
			if (hWnd = FindFirstChildClass(m_hWndListClip, CLSNAME_CHECKBOX))
			{
				SendMessage(hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0);
				/*SendMessage( hWnd, WM_LBUTTONUP, 0, 0 );*/
			}
		}
		else if (bMouseClickSim && iCellType == COL_CellType_DropDownList && MustShowDDBtn(hWndFocusCol, lFocusRow, TRUE, TRUE) && !MTblIsListOpen(m_hWndTbl, hWndFocusCol))
		{
			// Drop-Down-Button "anklicken"
			POINT pt;
			GetCursorPos(&pt);

			HWND hWnd = WindowFromPoint(pt);
			if (hWnd && hWnd == FindFirstChildClass(m_hWndListClip, CLSNAME_DROPDOWNBTN))
			{
				ScreenToClient(hWnd, &pt);
				SendMessage(hWnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELONG(pt.x, pt.y));
			}
		}

		// List-Clip incl. aller Kindobjekte neu zeichnen
		if (m_hWndListClip)
		{
			RedrawWindow(m_hWndListClip, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN);
			if (iCellType == COL_CellType_CheckBox)
			{
				HWND hWnd = FindFirstChildClass(m_hWndListClip, CLSNAME_CHECKBOX);
				InvalidateRect(hWnd, NULL, FALSE);
			}
		}

		// Focus wieder zeichnen
		m_bNoFocusPaint = FALSE;
	}
	else
		m_pscHookEditStyle = NULL;

	m_bSettingFocus = FALSE;
	//m_bNoFocusPaint = FALSE;

	if (MustSuppressRowFocusCTD())
	{
		KillFocusFrameI();
		InvalidateRowFocus(lFocusRow, FALSE);
	}

	// Update Focus
	int iRedrawMode = bFocusSet && m_bCellMode ? UF_REDRAW_INVALIDATE : UF_REDRAW_AUTO;
	int iSelChangeMode = bFocusSet && m_bCellMode ? UF_SELCHANGE_CLEAR_CELLS : UF_SELCHANGE_NONE;
	UpdateFocus(TBL_Error, NULL, iRedrawMode, iSelChangeMode);
}


// AnyCellsSelected
BOOL MTBLSUBCLASS::AnyCellSelected()
{
	if (!m_bCellMode)
		return FALSE;

	CMTblMetrics tm(m_hWndTbl);
	CMTblRow *pRow;
	CMTblRow **aRows[2];
	long aLastValidPos[2], lUpperBound, lLastValidPos;
	MTblColList::iterator it, itEnd;

	aRows[0] = m_Rows->GetRowArray(0, &lUpperBound, &lLastValidPos);
	aLastValidPos[0] = lLastValidPos;

	aRows[1] = m_Rows->GetRowArray(-1, &lUpperBound, &lLastValidPos);
	aLastValidPos[1] = lLastValidPos;

	for (int i = 0; i < 2; i++)
	{
		if (aRows[i])
		{
			for (int j = 0; j <= aLastValidPos[i]; j++)
			{
				if ((pRow = aRows[i][j]) && !pRow->IsDelAdj())
				{
					itEnd = pRow->m_Cols->m_List.end();

					for (it = pRow->m_Cols->m_List.begin(); it != itEnd; ++it)
					{
						if ((*it).second.QueryFlags(MTBL_CELL_FLAG_SELECTED))
							return TRUE;
					}
				}
			}
		}
	}

	return FALSE;
}


// AnyHighlightedItems
BOOL MTBLSUBCLASS::AnyHighlightedItems()
{
	return m_pHLIL->size() > 0;
}


// AnyMergeRowSelected
BOOL MTBLSUBCLASS::AnyMergeRowSelected(LPMTBLMERGEROW pmr)
{
	if (!pmr)
		return FALSE;

	BOOL bAnySelected = Ctd.TblQueryRowFlags(m_hWndTbl, pmr->lRowFrom, ROW_Selected);
	if (!bAnySelected)
	{
		long lRow = pmr->lRowFrom;
		if (FindNextRow(&lRow, ROW_Selected, ROW_Hidden, pmr->lRowTo))
			bAnySelected = TRUE;
	}

	return bAnySelected;
}


// AreRowsOfSameArea
BOOL MTBLSUBCLASS::AreRowsOfSameArea(long lRow1, long lRow2)
{
	if (lRow1 == TBL_Error || lRow2 == TBL_Error)
		return FALSE;

	return (lRow1 >= 0 && lRow2 >= 0) || (lRow1 < 0 && lRow2 < 0);
}


// AttachCol
BOOL MTBLSUBCLASS::AttachCol(HWND hWndCol)
{
	if (!SubClassCol(hWndCol))
		return FALSE;

	m_CurCols->SetInvalid();

	CMTblCol *pCol = m_Cols->EnsureValid(hWndCol);
	if (!pCol)
		return FALSE;

	if (!pCol->m_pCellType)
	{
		pCol->m_pCellType = new CMTblCellType;
		pCol->m_pCellTypeCurr = pCol->m_pCellType;
		if (!pCol->m_pCellType->DefineFromColumn(hWndCol))
			return FALSE;
	}

	int iCellType = Ctd.TblGetCellType(hWndCol);
	if (iCellType == COL_CellType_DropDownList)
	{
		HWND hWndExt = (HWND)SendMessage(m_hWndTbl, TIM_GetExtEdit, Ctd.TblQueryColumnID(hWndCol) - 1, 0);
		if (hWndExt)
			SubClassExtEdit(hWndCol, hWndExt);
	}

	return TRUE;
}


// AttachCols
BOOL MTBLSUBCLASS::AttachCols()
{
	HWND hWndCol = NULL;
	/*while( MTblFindNextCol( m_hWndTbl, &hWndCol, 0, 0 ) )
	{
	if ( !AttachCol( hWndCol ) )
	return FALSE;
	}*/

	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	int iCurCols = CurCols.size();
	for (int i = 0; i < iCurCols; ++i)
	{
		hWndCol = CurCols[i];
		if (!AttachCol(hWndCol))
			return FALSE;
	}

	return TRUE;
}


// AutoMergeCells
BOOL MTBLSUBCLASS::AutoMergeCells(long lRow, LPMTBLMERGEROW pmrOld, DWORD dwFlags)
{
	long lMergedOld = 0;
	MTBLMERGEROW mrOld;
	if (pmrOld)
	{
		lMergedOld = pmrOld->iMergedRows;

		pmrOld->CopyTo(mrOld);
		pmrOld = &mrOld;
	}


	long lMergedNew = 0;
	LPMTBLMERGEROW pmr = FindMergeRow(lRow, FMR_ANY);

	if (pmr && pmr->lRowFrom != lRow)
		return FALSE;
	else if (pmr && pmr->lRowFrom == lRow)
		lMergedNew = pmr->iMergedRows;

	BOOL bBeginsAtOldRowBegin, bEndsAtOldRowEnd;
	CMTblRow *pRow;
	HWND hWndCol = NULL, hWndColMerge;
	int iPos;
	long lCols, lRows, lRowMerge;
	LPMTBLMERGECELL pmc;
	//while( iPos = MTblFindNextCol( m_hWndTbl, &hWndCol, COL_Visible, 0 ) )
	while (iPos = FindNextCol(&hWndCol, COL_Visible, 0))
	{
		hWndColMerge = hWndCol;
		lRowMerge = lRow;
		lCols = lRows = -1;

		pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);

		if (pmc)
		{
			// Beginnt Zellenverbund in der ersten Zeile des alten Zeilenverbundes?
			bBeginsAtOldRowBegin = FALSE;
			if (pmrOld)
				bBeginsAtOldRowBegin = pmc->lRowFrom == pmrOld->lRowFrom;
			else
				bBeginsAtOldRowBegin = pmc->lRowFrom == lRow;


			// Endet Zellenverbund in der letzten Zeile des alten Zeilenverbundes?
			bEndsAtOldRowEnd = FALSE;
			if (pmrOld)
				bEndsAtOldRowEnd = pmc->lRowTo == pmrOld->lRowTo;
			else
				bEndsAtOldRowEnd = pmc->lRowTo == lRow;

			// Neuen Zellenverbund ermitteln
			if (bBeginsAtOldRowBegin && pmc->hWndColFrom == hWndCol && lMergedNew != lMergedOld)
			{
				lRows = pmc->iMergedRows + (lMergedNew - lMergedOld);
			}
			else if (bEndsAtOldRowEnd && lMergedNew != lMergedOld)
			{
				lRows = pmc->iMergedRows + (lMergedNew - lMergedOld);
				lRowMerge = pmc->lRowFrom;
			}
		}
		else
		{
			lRows = lMergedNew;
		}

		pRow = m_Rows->GetRowPtr(lRowMerge);
		if (!pRow)
			continue;

		if (lCols >= 0)
		{

			pRow->m_Cols->SetMerged(hWndColMerge, lCols);
			m_bMergeCellsInvalid = TRUE;
		}

		if (lRows >= 0)
		{
			pRow->m_Cols->SetMergedRows(hWndColMerge, lRows);
			m_bMergeCellsInvalid = TRUE;
		}
	}

	return TRUE;
}


// BeforeKillEdit
void MTBLSUBCLASS::BeforeKillEdit()
{
	if (MustTrapRowSelChanges())
	{
		BeginRowSelTrap();
		m_bKillEditStartedRowSelTrap = TRUE;
	}
	else
		m_bKillEditStartedRowSelTrap = FALSE;

	m_bNoFocusPaint = FALSE;

	m_pCounter->IncNoFocusUpdate();
}


// BeforeSetFocusCell
int MTBLSUBCLASS::BeforeSetFocusCell(HWND hWndCol, long lRow, BOOL bSendValidateMsg /*=FALSE*/)
{
	m_bSettingFocus = TRUE;

	int iRet = VALIDATE_Ok;

	// Aktuellen Focus ermitteln
	HWND hWndColCur = NULL;
	long lRowCur = lRow;
	Ctd.TblQueryFocus(m_hWndTbl, &lRowCur, &hWndColCur);

	// Validate-Nachricht schicken?
	if (bSendValidateMsg)
		iRet = Validate(hWndColCur, hWndCol) ? VALIDATE_Ok : VALIDATE_Cancel;

	// Focus nicht zeichnen?
	//if ( lRowCur != lRow )
	//	m_bNoFocusPaint = TRUE;

	// Aktuelle Focus-Zelle neu zeichnen?
	if (hWndColCur && IsMergeCell(hWndColCur, lRowCur))
		InvalidateCell(hWndColCur, lRowCur, NULL, FALSE);

	// Wenn Validierung Ok...
	if (iRet != VALIDATE_Cancel)
	{
		// Hat Tabelle das EditLeftJustify-Flag?
		m_bEditLeftJustify = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_EditLeftJustify);

		// Spalten-Typ ermitteln/anpassen
		int iCellType = 0;

		CMTblCellType *pctCurr = NULL;
		CMTblCol *pCol = m_Cols->Find(hWndCol);
		if (pCol)
			pctCurr = pCol->m_pCellTypeCurr;

		CMTblCellType *pctCell;
		pctCell = GetEffCellTypePtr(lRow, hWndCol);

		if (pctCell && pctCell != pctCurr)
		{
			iCellType = pctCell->m_iType;

			//m_bKillEditSuppressSelection = TRUE;

			if (pctCell->IsDropDownList())
				SetDropDownListContext(hWndCol, lRow);
			else
				SetCurrentCellType(hWndCol, pctCell);

			//m_bKillEditSuppressSelection = FALSE;
		}
		else
			iCellType = Ctd.TblGetCellType(hWndCol);

		// Evtl. dafür sorgen, dass bei Erzeugung des List-Clips der Edit-Stil-Hook gesetzt wird
		// und dadurch im Hook der gewünschte Stil gesetzt wird
		if ( /*!m_bEditLeftJustify &&*/ iCellType != COL_CellType_CheckBox && iCellType != COL_CellType_PopupEdit)
		{
			// ***Text-Ausrichtung***
			int iNeededJustify = -1;

			// Effektive Text-Ausrichtung ermitteln
			int iEffJustify = GetEffCellEditJustify(lRow, hWndCol);

			// Text zu breit für zentrierte/rechtsbündige Ausrichtung?
			BOOL bTooWide = FALSE;
			if (iEffJustify != DT_LEFT)
			{
				int iLeft, iTop, iRight, iBottom, iWid;
				int iRet = MTblGetCellRectEx(hWndCol, lRow, &iLeft, &iTop, &iRight, &iBottom, MTGR_INCLUDE_MERGED);
				if (iRet != MTGR_RET_ERROR)
				{
					iWid = iRight - iLeft;
				}
				else
					iWid = SendMessage(hWndCol, TIM_QueryColumnWidth, 0, 0);

				// Ausmaße des Zellentexts ermitteln
				SIZE s = { 0, 0 };
				GetCellTextExtent(hWndCol, lRow, &s);

				// Zur Verfügung stehenden Platz ermitteln...

				// Standard-Metriken ermitteln
				MTBLGETMETRICS tm;
				SendMessage(m_hWndTbl, TIM_GetMetrics, 0, LPARAM(&tm));

				// Nicht editierbaren Bereich ausschneiden?
				BOOL bClipNonedit = QueryCellFlags(lRow, hWndCol, MTBL_CELL_FLAG_CLIP_NONEDIT) || QueryColumnFlags(hWndCol, MTBL_COL_FLAG_CLIP_NONEDIT) || QueryRowFlags(lRow, MTBL_ROW_CLIP_NONEDIT);

				// Buttons außerhalb?
				BOOL bButtonsOutsideCell = QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL);

				// Müssen Buttons angezeigt werden?
				BOOL bBtn = MustShowBtn(hWndCol, lRow);
				BOOL bDDBtn = MustShowDDBtn(hWndCol, lRow);

				// Effektiven Button ermitteln
				CMTblBtn *pBtn = NULL;
				if (bBtn)
					pBtn = GetEffCellBtn(hWndCol, lRow);

				// Platz verringern...
				if (bClipNonedit)
				{
					// Bild ermitteln
					CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);

					// Textbereich links ermitteln
					BOOL bBtnLeft = pBtn && pBtn->IsLeft() && !bButtonsOutsideCell;
					int iAreaLeft = GetCellTextAreaLeft(lRow, hWndCol, iLeft, tm.ptCellLeading.x, Img, bBtnLeft);

					// Differenz links abziehen
					iWid -= (iAreaLeft - iLeft);

					// Textbereich rechts ermitteln
					BOOL bBtnRight = pBtn && !pBtn->IsLeft() && !bButtonsOutsideCell;
					int iAreaRight = GetCellTextAreaRight(lRow, hWndCol, iRight, tm.ptCellLeading.x, Img, bBtnRight, bDDBtn);

					// Differenz rechts abziehen
					iWid -= (iRight - iAreaRight);
				}
				else
				{
					if (!bButtonsOutsideCell)
					{
						// Button-Breite abziehen
						if (pBtn)
							iWid -= CalcBtnWidth(pBtn, hWndCol, lRow);

						// Drop-Down-Button Breite abziehen
						if (bDDBtn)
							iWid -= MTBL_DDBTN_WID;
					}
				}

				// Ränder abziehen
				iWid -= (2 * tm.ptCellLeading.x);

				// Zu breit?
				if (s.cx > iWid)
					bTooWide = TRUE;
			}

			// Ausrichtung der Spalte ermitteln
			int iColJustify = GetColJustify(hWndCol);

			// Benötigte Ausrichtung ermitteln
			if (bTooWide)
				iNeededJustify = DT_LEFT;
			else
				iNeededJustify = iEffJustify;

			// Ggfs. Benötigte Ausrichtung für Hook setzen und Hook aktivieren
			if (iNeededJustify != iColJustify)
			{
				m_iHookEditJustify = iNeededJustify;
				m_pscHookEditStyle = (void*)this;
			}

			// ***Multiline***
			// Ggfs. Multiline-Flag für Hook setzen und Hook aktivieren
			if (GetCellWordWrap(hWndCol, lRow))
			{
				m_bHookEditMultiline = TRUE;
				m_pscHookEditStyle = (void*)this;
			}
		}
	}

	return iRet;
}


// BeforeScroll
void MTBLSUBCLASS::BeforeScroll()
{
	m_bScrolling = TRUE;

	LRESULT lRet = SendMessage(m_hWndTbl, TIM_QueryPixelOrigin, 0, 0);
	m_iColOriginBeforeScroll = (int)LOWORD(lRet);
	m_iRowOriginBeforeScroll = (int)HIWORD(lRet);
}

// BeginColSelTrap
void MTBLSUBCLASS::BeginColSelTrap(HWND hWndCol /*= NULL*/)
{
	m_bTrappingColSelChanges = TRUE;
	GetColSel(GETSEL_BEFORE, hWndCol);
}

// BeginRowSelTrap
void MTBLSUBCLASS::BeginRowSelTrap(long lRow /*= TBL_Error*/, RowPtrVector *pvRows /* = NULL*/)
{
	m_bTrappingRowSelChanges = TRUE;
	GetRowSel(GETSEL_BEFORE, lRow, pvRows);
}


// BtnWndProc
LRESULT CALLBACK MTBLSUBCLASS::BtnWndProc(HWND hWndBtn, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_CAPTURECHANGED:
		return OnBtnCaptureChanged(hWndBtn, wParam, lParam);
	case WM_DESTROY:
		return OnBtnDestroy(hWndBtn, wParam, lParam);
	case WM_LBUTTONDOWN:
		return OnBtnLButtonDown(hWndBtn, wParam, lParam);
	case WM_LBUTTONUP:
		return OnBtnLButtonUp(hWndBtn, wParam, lParam);
	case WM_MOUSEMOVE:
		return OnBtnMouseMove(hWndBtn, wParam, lParam);
	case WM_PAINT:
		return OnBtnPaint(hWndBtn, wParam, lParam);
	default:
		return DefWindowProc(hWndBtn, uMsg, wParam, lParam);
	}
}


// CalcBtnRect
BOOL MTBLSUBCLASS::CalcBtnRect(CMTblBtn *pb, HWND hWndCol, long lRow, CMTblMetrics *pTblMetrics, LPRECT prBounds, HDC hDC, HFONT hFont, LPRECT prBtn)
{
	if (!pb)
		return FALSE;
	if (!prBounds)
		return FALSE;
	if (!prBtn)
		return FALSE;

	if (!pTblMetrics)
	{
		CMTblMetrics tm(m_hWndTbl);
		pTblMetrics = &tm;
	}


	BOOL bDelFont = FALSE;
	if (!hFont)
		hFont = GetBtnFont(pb, lRow, hWndCol, &bDelFont);

	int iWid = CalcBtnWidth(pb, hWndCol, lRow, hFont);
	int iHt = 0;

	BOOL bRelDC = FALSE;
	if (!hDC)
	{
		hDC = GetDC(m_hWndTbl);
		bRelDC = TRUE;

	}
	if (hDC)
	{
		HGDIOBJ hOldFont = SelectObject(hDC, hFont);
		SIZE s = { 0, 0 };
		GetTextExtentPoint32(hDC, _T("A"), 1, &s);
		SelectObject(hDC, hOldFont);

		if (bRelDC)
			ReleaseDC(m_hWndTbl, hDC);

		iHt = s.cy + 4;
	}
	if (bDelFont)
		DeleteObject(hFont);


	if (MImgIsValidHandle(pb->m_hImage))
	{
		SIZE siImg = { 0, 0 };
		MImgGetSize(pb->m_hImage, &siImg);
		iHt = max(siImg.cy + 4, iHt);
	}

	if (pb->IsLeft())
		prBtn->left = GetCellAreaLeft(lRow, hWndCol, prBounds->left);
	else
		prBtn->left = max(prBounds->right - iWid, prBounds->left);

	prBtn->top = prBounds->top + pTblMetrics->m_CellLeadingY;
	prBtn->right = pb->IsLeft() ? min(prBtn->left + iWid, prBounds->right) : prBounds->right;
	prBtn->bottom = min(prBtn->top + iHt, prBounds->bottom);

	return TRUE;
}


// CalcBtnWidth
int MTBLSUBCLASS::CalcBtnWidth(CMTblBtn * pb, HWND hWndCol, long lRow, HFONT hFont /*=NULL*/)
{
	if (!pb) return 0;

	CMTblBtn & b = *pb;

	int iBtnWid = -1;
	if (b.m_iWidth < 0)
	{
		// Breite anpassen
		int iImgWid = 0, iTxtWid = 0;
		SIZE s;

		// Bildbreite ermitteln
		if (MImgIsValidHandle(b.m_hImage))
		{
			if (MImgGetSize(b.m_hImage, &s))
				iImgWid = s.cx;
		}

		// Textbreite ermitteln
		if (b.m_sText.size() > 0)
		{
			HDC hDC = GetDC(m_hWndTbl);
			if (hDC)
			{
				BOOL bDelFont = FALSE;
				if (!hFont)
					hFont = GetBtnFont(pb, lRow, hWndCol, &bDelFont);
				HGDIOBJ hOldFont = SelectObject(hDC, hFont);
				if (GetTextExtentPoint32(hDC, b.m_sText.c_str(), b.m_sText.size(), &s))
				{
					iTxtWid = s.cx;
				}
				SelectObject(hDC, hOldFont);
				ReleaseDC(m_hWndTbl, hDC);
				if (bDelFont)
					DeleteObject(hFont);
			}
		}

		if (iImgWid > 0 || iTxtWid > 0)
		{
			iBtnWid = iImgWid + (iImgWid > 0 && iTxtWid > 0 ? 2 : 0) + iTxtWid + 4;
		}

		if (iBtnWid < 0)
			iBtnWid = MTBL_DDBTN_WID;
	}
	else
		iBtnWid = b.m_iWidth;

	return iBtnWid;
}


// CalcCellTextAreaXCoord
// Berechnet die tatsächlichen X-Koordinaten des Texts
void MTBLSUBCLASS::CalcCellTextAreaXCoord(int iTextAreaLeft, int iTextAreaRight, int iCellLeadingX, int iJustify, BOOL bWordWrap, POINT &pt)
{
	pt.x = iTextAreaLeft;
	pt.y = iTextAreaRight;

	if (iJustify == DT_LEFT || bWordWrap)
		pt.x += iCellLeadingX;

	if (iJustify == DT_RIGHT || bWordWrap)
		pt.y -= iCellLeadingX;
}


// CalcCellTextRect
// Berechnet das tatsächliche Text-Rechteck
BOOL MTBLSUBCLASS::CalcCellTextRect(long lCellTop, long lCellBottom, long lTextAreaLeft, long lTextAreaRight, long lCellLeadingX, long lCellLeadingY, int iJustify, int iVAlign, SIZE &siText, LPRECT prTxt)
{
	if (!prTxt)
		return FALSE;

	if (iJustify == DT_CENTER)
	{
		int iDiff = 0;
		if (lTextAreaRight > lTextAreaLeft)
			iDiff = (lTextAreaRight - lTextAreaLeft) - siText.cx;
		prTxt->left = lTextAreaLeft + (iDiff / 2);
		if (prTxt->left <= lTextAreaLeft && iDiff % 2)
			--prTxt->left;
	}
	else if (iJustify == DT_RIGHT)
		prTxt->left = lTextAreaRight - lCellLeadingX - siText.cx;
	else
		prTxt->left = lTextAreaLeft + lCellLeadingX;

	prTxt->right = prTxt->left + siText.cx;

	prTxt->top = lCellTop + lCellLeadingY;
	int iCellHt = lCellBottom - lCellTop;
	if (iCellHt > siText.cy)
	{
		if (iVAlign == DT_VCENTER)
			prTxt->top = max(prTxt->top, lCellTop + ((iCellHt - siText.cy) / 2));
		else if (iVAlign == DT_BOTTOM)
			prTxt->top = lCellBottom - lCellLeadingY - siText.cy;
	}

	prTxt->bottom = prTxt->top + siText.cy;

	return TRUE;
}


// CalcCheckBoxRect
// Berechnet das CheckBox-Rechteck
void MTBLSUBCLASS::CalcCheckBoxRect(int iJustify, int iVAlign, const RECT &rCbArea, long lCellLeadingX, long lCellLeadingY, long lCharBoxY, RECT &r)
{
	int iSize = CMTblMetrics::QueryCheckBoxSize(m_hWndTbl);

	long lAreaWid = rCbArea.right - rCbArea.left;

	if (iJustify == DT_CENTER)
		r.left = max(rCbArea.left, rCbArea.left + ((lAreaWid + 1 - iSize) / 2));
	else if (iJustify == DT_LEFT)
		//r.left = rCbArea.left;
		r.left = rCbArea.left + lCellLeadingX;
	else
		//r.left = max( rCbArea.left, rCbArea.right - iSize );
		r.left = max(rCbArea.left, rCbArea.right - lCellLeadingX - iSize);

	long lCbLeadingTop = lCellLeadingY + max(0, ((lCellLeadingY + lCharBoxY + 1 - iSize) / 2));
	if (iVAlign == DT_TOP)
		r.top = rCbArea.top + lCbLeadingTop;
	else if (iVAlign == DT_VCENTER)
		r.top = max(rCbArea.top + lCbLeadingTop, rCbArea.top + ((rCbArea.bottom - rCbArea.top - iSize) / 2));
	else
		r.top = max(rCbArea.top + lCbLeadingTop, rCbArea.bottom - iSize - 1);

	r.right = r.left + iSize;
	r.bottom = r.top + iSize;
}


// CalcDDBtnRect
BOOL MTBLSUBCLASS::CalcDDBtnRect(HWND hWndCol, long lRow, CMTblMetrics *pTblMetrics, LPRECT prBounds, HDC hDC, HFONT hFont, LPRECT prBtn)
{
	if (!prBounds)
		return FALSE;
	if (!prBtn)
		return FALSE;

	if (!pTblMetrics)
	{
		CMTblMetrics tm(m_hWndTbl);
		pTblMetrics = &tm;
	}


	BOOL bDelFont = FALSE;
	if (!hFont)
		hFont = MTblGetEffCellFontHandle(hWndCol, lRow, &bDelFont);

	int iWid = MTBL_DDBTN_WID;
	int iHt = 0;

	BOOL bRelDC = FALSE;
	if (!hDC)
	{
		hDC = GetDC(m_hWndTbl);
		bRelDC = TRUE;
	}
	if (hDC)
	{
		HGDIOBJ hOldFont = SelectObject(hDC, hFont);
		SIZE s = { 0, 0 };
		GetTextExtentPoint32(hDC, _T("A"), 1, &s);
		SelectObject(hDC, hOldFont);

		if (bRelDC)
			ReleaseDC(m_hWndTbl, hDC);

		iHt = s.cy;
	}
	if (bDelFont)
		DeleteObject(hFont);


	prBtn->left = max(prBounds->right - iWid, prBounds->left);
	prBtn->top = prBounds->top + pTblMetrics->m_CellLeadingY;
	prBtn->right = prBounds->right;
	prBtn->bottom = min(prBtn->top + iHt, prBounds->bottom);

	return TRUE;
}


// CalcMaxScrollRangeV
long MTBLSUBCLASS::CalcMaxScrollRangeV(BOOL bSplitBar /*= FALSE*/, int iScrPos /*= TBL_Error*/)
{
	// Metriken...
	CMTblMetrics tm(m_hWndTbl);

	// Zeilenarray ermitteln
	long lLastValidPos, lUpperBound;
	CMTblRow **RowArr = m_Rows->GetRowArray(bSplitBar ? -1 : 0, &lUpperBound, &lLastValidPos);

	// Ggf. aktuelle Scrollposition ermitteln
	if (iScrPos == TBL_Error)
	{
		HWND hWndScr;
		int iScrCtl;
		if (!GetScrollBarV(&tm, bSplitBar, &hWndScr, &iScrCtl))
			return TBL_Error;
		iScrPos = FlatSB_GetScrollPos(hWndScr, iScrCtl);
	}

	// Max. Zeile ermitteln
	long lMaxRow = TBL_Error;
	if (bSplitBar)
	{
		long lRow = 0;
		if (Ctd.TblFindPrevRow(m_hWndTbl, &lRow, 0, 0))
			lMaxRow = lRow;
	}
	else
	{
		long lDummy;
		Ctd.TblQueryScroll(m_hWndTbl, &lDummy, &lDummy, &lMaxRow);
		// Defect #186
		if (lMaxRow < 0)
			return TBL_Error;
	}

	// Anzahl Zeilen ermitteln, welche nicht im sichtbaren Bereich liegen
	long lAdd = 0;
	long lRowsNotInVisRange = 0;
	if (lMaxRow != TBL_Error)
	{
		long lRow = m_Rows->GetVisPosRow(bSplitBar ? iScrPos - TBL_MinSplitRow : iScrPos, bSplitBar);
		if (lRow != TBL_Error)
		{
			long lFirstRow = lRow, lLastPartlyVisRow = TBL_Error;
			long lPixPartlyNotVis, lPixEmptyRows;
			long lRowsBottom = bSplitBar ? tm.m_SplitRowsBottom : tm.m_RowsBottom;
			long lRowsTop = bSplitBar ? tm.m_SplitRowsTop : tm.m_RowsTop;
			long lRowBottom, lRowTop = lRowsTop;

			CMTblRow *pRow;
			long lEffRowHt, lLastPartlyVisPos = -1;
			long lPos = bSplitBar ? m_Rows->GetSplitRowPos(lRow) : lRow;
			long lMaxPos = bSplitBar ? m_Rows->GetSplitRowPos(lMaxRow) : lMaxRow;
			long lFirstPos = lPos;

			for (; lPos <= lMaxPos; ++lPos)
			{
				pRow = lPos <= lLastValidPos ? RowArr[lPos] : NULL;
				if (pRow && !pRow->IsVisible())
					continue;

				lEffRowHt = pRow ? GetEffRowHeight(pRow->GetNr(), &tm) : tm.m_RowHeight;
				lRowBottom = lRowTop + (lEffRowHt - 1);
				if (lRowBottom > lRowsBottom)
				{
					if (lRowTop <= lRowsBottom)
					{
						lLastPartlyVisPos = lPos;
						lPixPartlyNotVis = lRowBottom - lRowsBottom;
					}
					else
					{
						lRowsNotInVisRange++;
						for (lPos += 1; lPos <= lMaxPos; ++lPos)
						{
							if (lPos > lLastValidPos)
							{
								lRowsNotInVisRange += ((lMaxPos - lPos) + 1);
								break;
							}
							else
							{
								pRow = RowArr[lPos];
								if (pRow && !pRow->IsVisible())
									continue;
								lRowsNotInVisRange++;
							}
						}
						break;
					}
				}
				lRowTop = lRowBottom + 1;
			}

			if (lLastPartlyVisPos >= 0)
			{
				long lPix = 0;
				for (lPos = lFirstPos; lPos < lLastPartlyVisPos; ++lPos)
				{
					pRow = lPos <= lLastValidPos ? RowArr[lPos] : NULL;
					if (pRow && !pRow->IsVisible())
						continue;

					lEffRowHt = pRow ? GetEffRowHeight(pRow->GetNr(), &tm) : tm.m_RowHeight;
					lPix += lEffRowHt;
					++lAdd;
					if (lPix >= lPixPartlyNotVis)
						break;
				}
			}
			else
			{
				lPixEmptyRows = 0;
				while (lRowTop <= lRowsBottom)
				{
					lRowBottom = lRowTop + (tm.m_RowHeight - 1);
					if (lRowBottom > lRowsBottom)
						lPixEmptyRows += (lRowsBottom - lRowTop);
					else
						lPixEmptyRows += tm.m_RowHeight;

					lRowTop = lRowBottom + 1;
				}

				if (lPixEmptyRows > 0)
				{
					long lPix = 0;
					long lMinRow = bSplitBar ? TBL_MinSplitRow : 0;
					for (lPos = lFirstPos - 1; lPos >= 0; --lPos)
					{
						pRow = lPos <= lLastValidPos ? RowArr[lPos] : NULL;
						if (pRow && !pRow->IsVisible())
							continue;

						lEffRowHt = pRow ? GetEffRowHeight(pRow->GetNr(), &tm) : tm.m_RowHeight;
						lPix += lEffRowHt;
						if (lPix > lPixEmptyRows)
							break;
						--lAdd;
					}
				}
			}

			lAdd += lRowsNotInVisRange;
		}
	}

	int iMinScrPos = bSplitBar ? TBL_MinSplitRow : 0;
	int iMaxRange = max(iMinScrPos, (iScrPos + lAdd));
	return iMaxRange;
}


// CallOldWndProc
// Ruft die "alte" Fensterprozedur auf
LRESULT MTBLSUBCLASS::CallOldWndProc(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	return CallWindowProc(m_wpOld, m_hWndTbl, uMsg, wParam, lParam);
}


// CellSelRangeChanged
BOOL MTBLSUBCLASS::CellSelRangeChanged(MTBLCELLRANGE &crOld, MTBLCELLRANGE &crNew, BOOL bInvalidate)
{
	if (!crOld.IsEqualTo(crNew))
	{
		BOOL bCellInOld, bCellInNew, bColInOld, bColInNew, bRowInOld, bRowInNew, bSelect;
		CMTblMetrics tm(m_hWndTbl);
		int iColPos, iColPosFrom, iColPosFromOld, iColPosFromNew, iColPosTo, iColPosToOld, iColPosToNew;
		long lRow, lRowFrom, lRowTo;

		iColPosFromOld = crOld.GetColFromPos();
		iColPosToOld = crOld.GetColToPos();

		iColPosFromNew = crNew.GetColFromPos();
		iColPosToNew = crNew.GetColToPos();

		iColPosFrom = min(iColPosFromOld, iColPosFromNew);
		iColPosTo = max(iColPosToOld, iColPosToNew);

		lRowFrom = min(crOld.lRowFrom, crNew.lRowFrom);
		lRowTo = max(crOld.lRowTo, crNew.lRowTo);

		struct COLINFO
		{
			HWND hWnd;
			BOOL bVisible;
		};

		COLINFO *paCols = new COLINFO[iColPosTo + 1];
		for (int i = iColPosFrom; i <= iColPosTo; i++)
		{
			paCols[i].hWnd = Ctd.TblGetColumnWindow(m_hWndTbl, i, COL_GetPos);
			paCols[i].bVisible = Ctd.TblQueryColumnFlags(paCols[i].hWnd, COL_Visible);
		}

		for (lRow = lRowFrom; lRow <= lRowTo; lRow++)
		{
			if (!IsRow(lRow) || Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Hidden))
				continue;

			bRowInOld = (lRow >= crOld.lRowFrom && lRow <= crOld.lRowTo);
			bRowInNew = (lRow >= crNew.lRowFrom && lRow <= crNew.lRowTo);

			for (iColPos = iColPosFrom; iColPos <= iColPosTo; iColPos++)
			{
				if (!paCols[iColPos].bVisible)
					continue;

				bColInOld = (iColPos >= iColPosFromOld && iColPos <= iColPosToOld);
				bColInNew = (iColPos >= iColPosFromNew && iColPos <= iColPosToNew);

				bCellInOld = (bRowInOld && bColInOld);
				bCellInNew = (bRowInNew && bColInNew);

				if (bCellInOld != bCellInNew)
				{
					bSelect = bCellInNew;
					SetCellFlags(lRow, paCols[iColPos].hWnd, &tm, MTBL_CELL_FLAG_SELECTED, bSelect, bInvalidate);
				}
			}
		}

		delete[]paCols;
	}

	return TRUE;
}


// CanSetFocusCell
BOOL MTBLSUBCLASS::CanSetFocusCell(HWND hWndCol, long lRow, BOOL bAllowMerged /*=TRUE*/)
{
	HWND hWndMergeCol;
	long lMergeRow;
	if (GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
	{
		if (!bAllowMerged)
			return FALSE;

		hWndCol = hWndMergeCol;
		lRow = lMergeRow;
	}

#ifdef TD_POPUP_EDIT_ALWAYS_SHOWN
	int iCellType = GetEffCellType(lRow, hWndCol);
	if (iCellType == COL_CellType_PopupEdit)
		return TRUE;
#endif

	/*if (m_Rows->IsRowDisabled(lRow)) return FALSE;
	if (!Ctd.IsWindowEnabled(hWndCol)) return FALSE;
	if (m_Rows->IsRowDisCol(lRow, hWndCol)) return FALSE;*/
	if (IsCellDisabled(hWndCol, lRow)) return FALSE;
	if (IsCheckBoxCell(hWndCol, lRow) && IsCellReadOnly(hWndCol, lRow)) return FALSE;

	return TRUE;
}


// CanSetFocusToCell
BOOL MTBLSUBCLASS::CanSetFocusToCell(HWND hWndCol, long lRow)
{
	if (!hWndCol) return FALSE;
	if (lRow == TBL_Error) return FALSE;

	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);

	HWND hWndColSend = pmc ? pmc->hWndColFrom : hWndCol;
	long lRowSend = pmc ? pmc->lRowFrom : lRow;
	if (SendMessage(m_hWndTbl, MTM_QueryNoFocusOnCell, (WPARAM)hWndColSend, lRowSend) != 0)
		return FALSE;
	else
		return TRUE;
}


// ClearCellSelection
BOOL MTBLSUBCLASS::ClearCellSelection(BOOL bInvalidate /*=TRUE*/)
{
	CMTblMetrics tm(m_hWndTbl);
	CMTblRow *pRow;
	CMTblRow **aRows[2];
	DWORD dwDiff;
	long aLastValidPos[2], lUpperBound, lLastValidPos;
	MTblColList::iterator it, itEnd;

	aRows[0] = m_Rows->GetRowArray(0, &lUpperBound, &lLastValidPos);
	aLastValidPos[0] = lLastValidPos;

	aRows[1] = m_Rows->GetRowArray(-1, &lUpperBound, &lLastValidPos);
	aLastValidPos[1] = lLastValidPos;

	for (int i = 0; i < 2; i++)
	{
		if (aRows[i])
		{
			for (int j = 0; j <= aLastValidPos[i]; j++)
			{
				if ((pRow = aRows[i][j]) && !pRow->IsDelAdj())
				{
					itEnd = pRow->m_Cols->m_List.end();

					for (it = pRow->m_Cols->m_List.begin(); it != itEnd; ++it)
					{
						dwDiff = (*it).second.SetFlags(MTBL_CELL_FLAG_SELECTED, FALSE);
						if (dwDiff && bInvalidate)
						{
							InvalidateCell((*it).second.m_hWnd, pRow->GetNr(), &tm, FALSE);
						}
					}
				}
			}
		}
	}

	return TRUE;
}


// ClearColumnSelection
BOOL MTBLSUBCLASS::ClearColumnSelection()
{
	HWND hWndCol = NULL;
	/*while ( MTblFindNextCol( m_hWndTbl, &hWndCol, 0, 0 ) )
	{
	Ctd.TblSetColumnFlags( hWndCol, COL_Selected, FALSE );
	}*/

	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	int iCurCols = CurCols.size();
	for (int i = 0; i < iCurCols; ++i)
	{
		hWndCol = CurCols[i];
		Ctd.TblSetColumnFlags(hWndCol, COL_Selected, FALSE);
	}

	return TRUE;
}


// CollapseRow
BOOL MTBLSUBCLASS::CollapseRow(CMTblRow * pRow, DWORD dwFlags)
{
	// Zeile ermitteln
	if (!pRow) return FALSE;
	long lRow = pRow->GetNr();

	// Arraydaten ermitteln
	long lLastValidPos, lUpperBound;
	CMTblRow ** RowArr = m_Rows->GetRowArray(lRow, &lUpperBound, &lLastValidPos);
	if (!RowArr) return FALSE;

	// Zeichnen unterdrücken
	LockPaint();

	// Kindzeilen verbergen
	long l;
	for (l = 0; l <= lLastValidPos; ++l)
	{
		if (RowArr[l])
		{
			if (RowArr[l]->IsDescendantOf(pRow))
			{
				BOOL bOk = Ctd.TblSetRowFlags(m_hWndTbl, RowArr[l]->GetNr(), ROW_Hidden, TRUE);
				RowArr[l]->SetFlags(MTBL_ROW_ISEXPANDED, FALSE);
			}
		}
	}

	// Flags setzen
	pRow->SetFlags(MTBL_ROW_ISEXPANDED, FALSE);

	// Zeichnen wieder erlaubt
	UnlockPaint();

	// Tabelle neu zeichnen
	if (dwFlags & MTCR_REDRAW)
		InvalidateRect(m_hWndTbl, NULL, FALSE);

	return TRUE;
}

// ColWndProc
LRESULT CALLBACK MTBLSUBCLASS::ColWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	BOOL bLockUpdate;
	CMTblCol *pCell;
	CMTblRow *pRow;
	ItemList::iterator it;
	long lNewPos, lOldPos, lRow, lWid;
	LRESULT lRet;
	LPMTBLCOLSUBCLASS psc = NULL;
	LPMTBLSUBCLASS psct = MTblGetSubClass(Ctd.ParentWindow(hWnd));

	if (psct)
		psc = psct->GetColSubClass(hWnd);

	if (psc)
	{
		switch (uMsg)
		{
		case LB_ADDSTRING:
		case LB_DELETESTRING:
		case LB_GETITEMDATA:
		case LB_INSERTSTRING:
		case LB_SETITEMDATA:
		{
			// Wenn Spalte Listbox-Nachrichten bekommt, dann wurde wohl eine VisList-Funktion aufgerufen
			lRet = LB_ERR;

			HWND hWndExt = (HWND)SendMessage(psct->m_hWndTbl, TIM_GetExtEdit, Ctd.TblQueryColumnID(hWnd) - 1, 0);
			if (hWndExt)
			{
				TCHAR cClassName[255];
				if (GetClassName(hWndExt, cClassName, 254) && lstrcmp(cClassName, _T("ListBox")) == 0)
					lRet = SendMessage(hWndExt, uMsg, wParam, lParam);
			}

			break;
		}

		case WM_NCDESTROY:
			// Abkoppeln
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)psc->wpOld);
			psct->m_scm->erase(hWnd);
			delete psc;

			// Ggf. Enter/Leave-Spalten zurücksetzen
			if (psct->m_hWndELCol == hWnd)
				psct->m_hWndELCol = NULL;
			if (psct->m_hWndELCol_Cell == hWnd)
				psct->m_hWndELCol_Cell = NULL;
			if (psct->m_hWndELCol_Btn == hWnd)
				psct->m_hWndELCol_Btn = NULL;
			if (psct->m_hWndELCol_DDBtn == hWnd)
				psct->m_hWndELCol_DDBtn = NULL;

			// Hervorgehobene Items in Spalte löschen
			for (it = psct->m_pHLIL->begin(); it != psct->m_pHLIL->end();)
			{
				if ((*it).WndHandle == hWnd)
				{
					it = psct->m_pHLIL->erase(it);
					if (it == psct->m_pHLIL->end()) break;
				}
				else
					it++;
			}

			// Aktuelle Spalten ungültig
			psct->m_CurCols->SetInvalid();

			// Merge-Zellen ungültig
			psct->m_bMergeCellsInvalid = TRUE;

			break;

		case WM_SETTEXT:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			psct->UpdateFocus();
			break;

		case CIM_DefineCellType:
			if (wParam)
			{
				if (wParam == 1)
					psc->bSettingCellType = TRUE;
				else
					psc->bSettingCellType = FALSE;
				lRet = 1;
			}
			else
			{
				// Kill Edit
				//Ctd.TblKillEdit( psct->m_hWndTbl );
				psct->KillEditNeutral();

				const int LOCK_NOTHING = 0;
				const int LOCK_UPDATE = 1;
				const int LOCK_PAINT = 2;

				// Start LockUpdate/LockPaint
				int iLock = LOCK_NOTHING;
				if (!psc->bSettingCellType)
				{
					if (LockWindowUpdate(psc->hWndTbl))
						iLock = LOCK_UPDATE;
				}
				else if (psct->m_bSettingFocus)
				{
					psct->LockPaint();
					iLock = LOCK_PAINT;
				}

				// Alte Prozedur aufrufen
				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

				// Neue Typdefinition übernehmen
				if (!psc->bSettingCellType)
				{
					CMTblCol *pCol = psct->m_Cols->Find(hWnd);
					if (pCol && pCol->m_pCellType)
					{
						pCol->m_pCellType->DefineFromColumn(hWnd);
						pCol->m_pCellTypeCurr = pCol->m_pCellType;
						if (pCol->m_pCellType->IsDropDownList())
						{
							pCol->m_DropDownListContext = TBL_Error;
							psct->SubClassExtEdit(hWnd, NULL);
						}
					}
				}

				// Ende LockUpdate/LockPaint
				if (iLock == LOCK_UPDATE)
					LockWindowUpdate(NULL);
				else if (iLock == LOCK_PAINT)
					psct->UnlockPaint();
			}

			break;

		case CIM_DefineCellTypeParamMiddle:
		case CIM_DefineCellTypeParamLast:
			// Alte Prozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Neue Typdefinition in Zelltyp-Objekt der Spalte übernehmen
			if (!psc->bSettingCellType)
			{
				CMTblCol *pCol = psct->m_Cols->Find(hWnd);
				if (pCol && pCol->m_pCellType)
				{
					pCol->m_pCellType->DefineFromColumn(hWnd);
				}
			}

			break;

		case CIM_GetExtEdit:
		{
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			break;
		}

		case CIM_SetEditType:
			// Kill Edit
			//Ctd.TblKillEdit( psct->m_hWndTbl );
			psct->KillEditNeutral();

			// Alte Prozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			break;

		case CIM_SetPos:
			lOldPos = Ctd.TblQueryColumnPos(hWnd);
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			if (lRet)
			{
				lNewPos = Ctd.TblQueryColumnPos(hWnd);
				if (lOldPos != lNewPos)
				{
					// Aktuelle Spalten ungültig
					psct->m_CurCols->SetInvalid();

					// Merge-Zellen ungültig
					psct->m_bMergeCellsInvalid = TRUE;

					// Bereich Zellenselektion/Tastatur zurücksetzen
					if (hWnd == psct->m_CellSelRangeKey.hWndColFrom)
						psct->m_CellSelRangeKey.SetEmpty();

					// Evtl. Erweiterte Nachricht schicken
					if (psct->m_bExtMsgs)
						PostMessage(psct->m_hWndTbl, MTM_ColPosChanged, WPARAM(hWnd), MAKELPARAM(lNewPos, lOldPos));
				}
			}
			break;

		case CIM_SetFlags:
			if (lParam & COL_Visible)
			{
				// Merge-Zellen ungültig
				psct->m_bMergeCellsInvalid = TRUE;
			}

			// Werden Flags gesetzt oder entfernt?
			BOOL bSet;
			bSet = (wParam != 0);

			// Wird sich COL_Selected ändern?			
			BOOL bSelFlagWillChange;
			bSelFlagWillChange = (lParam & COL_Selected) && Ctd.TblQueryColumnFlags(hWnd, COL_Selected) != bSet;

			// Wird sich COL_Editable ändern?
			BOOL bEditableFlagWillChange;
			bEditableFlagWillChange = (lParam & COL_Editable) && Ctd.TblQueryColumnFlags(hWnd, COL_Editable) != bSet;

			// Wird sich COL_Visible ändern?
			BOOL bVisibleFlagWillChange;
			bVisibleFlagWillChange = (lParam & COL_Visible) && Ctd.TblQueryColumnFlags(hWnd, COL_Visible) != bSet;

			// Wird sich COL_MultilineCell ändern?
			BOOL bMultiFlagWillChange;
			bMultiFlagWillChange = (lParam & COL_MultilineCell) && Ctd.TblQueryColumnFlags(hWnd, COL_MultilineCell) != bSet;

			// Wird sich COL_Invisible ändern?
			BOOL bInvisibleFlagWillChange;
			bInvisibleFlagWillChange = (lParam & COL_Invisible) && Ctd.TblQueryColumnFlags(hWnd, COL_Invisible) != bSet;

			// Vorsorglich KillEdit durchführen?
			if (bSelFlagWillChange || bEditableFlagWillChange || bVisibleFlagWillChange || bMultiFlagWillChange || bInvisibleFlagWillChange)
				Ctd.TblKillEdit(psct->m_hWndTbl);

			// Wird sich die Selektion ändern?
			BOOL bSelWillChange;
			bSelWillChange = bSelFlagWillChange && Ctd.TblQueryColumnFlags(hWnd, COL_Visible);

			BOOL bTrapColSelChanges;
			bTrapColSelChanges = bSelWillChange && psct->MustTrapColSelChanges();

			if (bTrapColSelChanges)
				psct->BeginColSelTrap(hWnd);

			bLockUpdate = bSelWillChange && psct->MustRedrawSelections() && !gbCtdHooked;
			//bLockUpdate = ( lParam & COL_Selected ) && psct->MustRedrawSelections( ) && Ctd.TblQueryColumnFlags( hWnd, COL_Visible );
			if (bLockUpdate)
				bLockUpdate = LockWindowUpdate(psc->hWndTbl);

			// Justify-Fix
			BOOL bCenter, bRight, bRemCenter, bRemRight, bRestoreCenter, bRestoreRight;
			if (!psc->bFixingJustify)
			{
				bCenter = Ctd.TblQueryColumnFlags(hWnd, COL_CenterJustify);
				bRemCenter = (wParam == 0 && lParam & COL_CenterJustify);
				bRestoreCenter = (bCenter && !bRemCenter);
				bRight = Ctd.TblQueryColumnFlags(hWnd, COL_RightJustify);
				bRemRight = (wParam == 0 && lParam & COL_RightJustify);
				bRestoreRight = (bRight && !bRemRight);

				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

				bRestoreCenter = (bRestoreCenter && !Ctd.TblQueryColumnFlags(hWnd, COL_CenterJustify));
				bRestoreRight = (bRestoreRight && !Ctd.TblQueryColumnFlags(hWnd, COL_RightJustify));

				if (bRestoreCenter || bRestoreRight)
				{
					psc->bFixingJustify = TRUE;
					DWORD dwFlags = 0;
					if (bRestoreCenter)
						dwFlags |= COL_CenterJustify;
					if (bRestoreRight)
						dwFlags |= COL_RightJustify;
					Ctd.TblSetColumnFlags(hWnd, dwFlags, TRUE);
					psc->bFixingJustify = FALSE;
				}
			}
			else
				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);


			if (bLockUpdate)
				LockWindowUpdate(NULL);

			if (bTrapColSelChanges)
				psct->EndColSelTrap(hWnd);

			if (bVisibleFlagWillChange && !bSet)
				psct->m_Rows->SetCellFlagsAnyRows(hWnd, MTBL_CELL_FLAG_SELECTED, FALSE);

			if (bEditableFlagWillChange)
				psct->InvalidateCol(hWnd, FALSE);

			if (lParam & COL_Visible)
				psct->UpdateFocus();

			break;

		case CIM_SetText:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			lRow = Ctd.TblQueryContext(psc->hWndTbl);

			if (psct->QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER) || psct->m_pCounter->GetRowHeights() > 0 || psct->IsMergeCell(hWnd, lRow))
				psct->InvalidateCell(hWnd, lRow, NULL, FALSE);

			break;

		case CIM_SetTextColor:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			if (lRet)
			{
				lRow = Ctd.TblQueryContext(psc->hWndTbl);
				if (psct->m_Rows->IsRowDelAdj(lRow)) return lRet;

				pRow = psct->m_Rows->EnsureValid(lRow);
				if (!pRow) return lRet;

				pCell = pRow->m_Cols->EnsureValid(hWnd);
				if (!pCell) return lRet;

				pCell->m_TextColor.Set(lParam);

				if (psct->QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER) || psct->m_pCounter->GetRowHeights() > 0 || psct->IsMergeCell(hWnd, lRow))
					psct->InvalidateCell(hWnd, lRow, NULL, FALSE);
			}
			break;

		case CIM_SetWidth:
			if (psct->IsInEditMode())
			{
				lWid = SendMessage(hWnd, TIM_QueryColumnWidth, 0, 0);
				if (lWid != lParam)
					Ctd.TblKillEdit(psct->m_hWndTbl);
			}

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			if (lRet)
			{
				psct->m_CurCols->UpdateWidth(hWnd);
				psct->UpdateFocus();
			}
			break;

		default:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		}
	}
	else
		lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);

	return lRet;
}


// CombineHighlighting
BOOL MTBLSUBCLASS::CombineHighlighting(int iItem, MTBLHILI &hili, int iItemAdd, MTBLHILI &hiliAdd)
{
	BOOL bCombFontEnh = (m_dwEffCellPropDetFlags & MTECPD_FLAG_COMB_FONT_ENH);

	if (hiliAdd.dwBackColor != MTBL_COLOR_UNDEF)
		hili.dwBackColor = hiliAdd.dwBackColor;

	if (hiliAdd.dwTextColor != MTBL_COLOR_UNDEF)
		hili.dwTextColor = hiliAdd.dwTextColor;

	if (iItemAdd != MTBL_ITEM_ROW && iItemAdd != MTBL_ITEM_COLUMN)
	{
		if (hiliAdd.lFontEnh != MTBL_FONT_UNDEF_ENH)
		{
			if (bCombFontEnh && hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
				hili.lFontEnh |= hiliAdd.lFontEnh;
			else
				hili.lFontEnh = hiliAdd.lFontEnh;
		}

		HANDLE hImg;

		hImg = hiliAdd.Img.GetHandle();
		if (hImg)
			hili.Img.SetHandle(hImg, NULL);

		hImg = hiliAdd.ImgExp.GetHandle();
		if (hImg)
			hili.ImgExp.SetHandle(hImg, NULL);
	}

	return TRUE;
}


// ContainerWndProc
LRESULT CALLBACK MTBLSUBCLASS::ContainerWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	long lRet = 0, lRow;
	POINT pt;
	RECT r;

	// Subclass-Struktur ermitteln
	LPMTBLCONTAINERSUBCLASS psc = (LPMTBLCONTAINERSUBCLASS)GetProp(hWnd, MTBL_CONTAINER_SUBCLASS_PROP);
	if (!psc) return 0;

	// Subclass-Struktur der Tabelle ermitteln
	LPMTBLSUBCLASS psct = MTblGetSubClass(psc->hWndTbl);
	if (!psct) return 0;

	// Alte Fensterprozedur prüfen
	if (!psc->wpOld) return 0;

	// Nachrichten verarbeiten
	switch (uMsg)
	{
	case WM_MOUSEMOVE:
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		// Ggf. internen Timer starten
		if (!psct->m_uiTimer)
			psct->StartInternalTimer(TRUE);

		break;

	case WM_NCDESTROY:
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		// Alte Fensterprozedur setzen + merken
		SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)psc->wpOld);
		// Property entfernen
		RemoveProp(hWnd, MTBL_CONTAINER_SUBCLASS_PROP);
		// Bereiche in der Tabelle neu zeichnen
		if (GetWindowRect(hWnd, &r) && MapWindowPoints(NULL, psct->m_hWndTbl, (LPPOINT)&r, 2))
		{
			// Bereich des Containers
			InvalidateRect(psct->m_hWndTbl, &r, FALSE);
			// Bereich des Row-Headers
			pt.x = r.left;
			pt.y = r.top;
			lRow = psct->RowFromPoint(pt);
			if (lRow != TBL_Error)
				psct->InvalidateRowHdr(lRow, FALSE, TRUE);
		}
		// Speicher freigeben
		delete psc;
		break;

	case WM_PAINT:
		// Focus ermitteln
		HWND hWndCol;
		long lRow;
		if (Ctd.TblQueryFocus(psc->hWndTbl, &lRow, &hWndCol))
		{
			HRGN hRgn = CreateRectRgn(0, 0, 0, 0);
			GetUpdateRgn(hWnd, hRgn, FALSE);

			// Zeichnen beginnen
			PAINTSTRUCT ps;
			BeginPaint(hWnd, &ps);

			// Zeichnen
			HDC hDC = ps.hdc;
			if (hDC)
			{
				// Effektive Hintergrundfarbe der Zelle ermitteln
				COLORREF clrBack = psct->GetEffCellBackColor(lRow, hWndCol);
				// Pinsel erzeugen
				HBRUSH hBrush = CreateSolidBrush(clrBack);
				if (hBrush)
				{
					// Hintergrund zeichnen
					FillRgn(hDC, hRgn, hBrush);

					// Pinsel löschen
					DeleteObject(hBrush);
				}
			}
			// Zeichnen beenden
			EndPaint(hWnd, &ps);

			DeleteObject(hRgn);
		}

		break;

	case WM_PARENTNOTIFY:
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		psct->OnContainerParentNotify(psct->m_hWndTbl, hWnd, wParam, lParam);

		break;

	default:
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
	}

	return lRet;
}


// CopyCurrentSelection
BOOL MTBLSUBCLASS::CopyCurrentSelection()
{
	typedef struct
	{
		HWND hWnd;
		int iID;
		BOOL bSelected;
		BOOL bAnyCellCopied;
	}
	COLINFO;

	BOOL bOk = TRUE;

	// Spalten ermitteln
	BOOL bAnyColSelected = FALSE;
	COLINFO aCols[MTBL_MAX_COLS];
	HWND hWndCol = NULL;
	int iCols = 0;
	while (FindNextCol(&hWndCol, COL_Visible, 0))
	{
		aCols[iCols].hWnd = hWndCol;
		aCols[iCols].iID = Ctd.TblQueryColumnID(hWndCol);
		if (aCols[iCols].bSelected = Ctd.TblQueryColumnFlags(hWndCol, COL_Selected))
			bAnyColSelected = TRUE;
		aCols[iCols].bAnyCellCopied = FALSE;

		++iCols;
		if (iCols == MTBL_MAX_COLS)
			break;
	}

	// Irgendwelche Zeilen markiert?
	BOOL bAnyRowSelected = Ctd.TblAnyRows(m_hWndTbl, ROW_Selected, ROW_Hidden);

	// Irgendwelche Zellen markiert?
	BOOL bAnyCellSelected = AnyCellSelected();

	// Kontext merken
	long lContext = Ctd.TblQueryContext(m_hWndTbl);

	// Fokus ermitteln
	long lFocusRow = m_pRowFocus ? m_pRowFocus->GetNr() : TBL_Error;
	HWND hWndFocusCol = m_bCellMode ? m_hWndColFocus : NULL;

	// Werte ermitteln
	BOOL bAnyCellCopied;
	BOOL bCopyCell;
	BOOL bCellSelected;
	BOOL bRowSelected;
	HSTRING hs = 0;
	list<vector<tstring>> value_list;
	tstring s;
	vector<tstring> values;

	// Zeilen und Split-Zeilen durchlaufen
	BOOL bSplitRows;
	CMTblRow *pRow;
	CMTblRow **Rows;
	int iColFrom = -1, iColTo = -1;
	long lRow;

	for (int iRowArea = 0; iRowArea < 2; iRowArea++)
	{
		bSplitRows = iRowArea > 0;
		lRow = bSplitRows ? TBL_MaxRow - 1 : -1;
		while (pRow = bSplitRows ? m_Rows->FindNextSplitRow(&lRow) : m_Rows->FindNextRow(&lRow))
		{
			// Versteckte Zeilen überspringen
			if (pRow->QueryHiddenFlag())
				continue;

			// Kontext setzen
			if (!Ctd.TblSetContextEx(m_hWndTbl, lRow))
			{
				bOk = FALSE;
				break;
			}

			// Zeile markiert?
			bRowSelected = Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Selected);

			// Spalten durchlaufen
			bAnyCellCopied = FALSE;
			values.clear();
			values.reserve(iCols);

			for (int iCol = 0; iCol < iCols; iCol++)
			{
				// Ermitteln, ob der Inhalt der Zelle kopiert werden muss
				bCopyCell = FALSE;

				if (!GetMergeCell(aCols[iCol].hWnd, lRow))
				{
					// Zelle selektiert?
					bCellSelected = m_bCellMode ? QueryCellFlags(lRow, aCols[iCol].hWnd, MTBL_CELL_FLAG_SELECTED) : FALSE;

					if (bCellSelected)
						bCopyCell = TRUE;
					else
					{
						// Sind Zeilen und Spalten markiert?
						if (bAnyRowSelected && bAnyColSelected)
							// Kopieren, wenn Zeile und Spalte markiert sind
							bCopyCell = bRowSelected && aCols[iCol].bSelected;
						// Sind Zeilen markiert?
						else if (bAnyRowSelected)
							// Kopieren, wenn die Zeile markiert ist
							bCopyCell = bRowSelected;
						// Sind Spalten markiert?
						else if (bAnyColSelected)
							// Kopieren, wenn Spalte markiert ist
							bCopyCell = aCols[iCol].bSelected;
						// Keine Zeilen oder Spalten markiert?
						else
							// Kopieren, wenn die Zeile bzw. die Zelle den Fokus hat
							bCopyCell = (lRow == lFocusRow) && (m_bCellMode ? aCols[iCol].hWnd == hWndFocusCol && !bAnyCellSelected: TRUE);
					}
				}

				// Inhalt der Zelle kopieren?
				if (bCopyCell)
				{
					// Zeile und Spalte berücksichtigen
					bAnyCellCopied = TRUE;
					aCols[iCol].bAnyCellCopied = TRUE;

					// Minimale und maximale Spalte mit kopierten Werten setzen
					if (iColFrom < 0 || iCol < iColFrom)
						iColFrom = iCol;
					if (iCol > iColTo)
						iColTo = iCol;

					// Text ermitteln
					if (!Ctd.TblGetColumnText(m_hWndTbl, aCols[iCol].iID, &hs))
					{
						bOk = FALSE;
						break;
					}

					// In tstring umwandeln
					if (Ctd.HStrToTStr(hs, s))
					{
						// TAB durch Leerzeichen ersetzen
						tstring::size_type nPos = 0;
						while ((nPos = s.find(_T("	"), nPos)) != tstring::npos)
						{
							s.replace(nPos, 1, _T(" "));
							nPos += 1;
						}

						// Zeilenumbruch durch Leerzeichen ersetzen
						nPos = 0;
						while ((nPos = s.find(_T("\r\n"), nPos)) != tstring::npos)
						{
							s.replace(nPos, 2, _T(" "));
							nPos += 1;
						}

						// Dem Wertearry hinzufügen
						values.push_back(s);
					}
					else
					{
						bOk = FALSE;
						break;
					}
				}
				else
					// Dem Wertearry einen Leerstring hinzufügen
					values.push_back(_T(""));
			}

			if (!bOk)
				break;

			// Zeile berücksichtigen?
			if (bAnyCellCopied)
			{
				// Wertearray der Werteliste hinzufügen
				value_list.push_back(values);
			}
		}

		if (!bOk)
			break;
	}

	// HSTRING dereferenzieren
	if (hs)
		Ctd.HStringUnRef(hs);

	// Ursprüngliche Kontext wiederherstellen
	Ctd.TblSetContextEx(m_hWndTbl, lContext);

	// Werte vorhanden?
	if (bOk && value_list.size() > 0)
	{
		// String Zwischenablage bauen
		long lGrow = 2097152;
		tstring sPaste(_T(""));

		list<vector<tstring>>::iterator it, itEnd = value_list.end();
		list<vector<tstring>>::iterator itLast = --value_list.end();

		// Werteliste (= Zeilen) durchlaufen
		for (it = value_list.begin(); it != itEnd; ++it)
		{
			// Wertearray (= Spalten) durchlaufen
			for (int iCol = iColFrom; iCol <= iColTo; iCol++)
			{
				// Spalte berücksichtigen?
				if (aCols[iCol].bAnyCellCopied)
				{
					// Wert ermitteln
					s = (*it)[iCol];

					// Wenn das Anhängen des Wertes die Kapazität des Zwischenablage-Strings überschreiten würde, Kapazität des Zwischenablage-Strings erhöhen
					if (sPaste.size() + s.size() + 2 > sPaste.capacity())
						sPaste.reserve(sPaste.capacity() + lGrow);

					// Wert anhängen
					sPaste += s;

					// Letzte Spalte?
					if (iCol == iColTo)
					{
						// Nicht die letzte Zeile?
						if (it != itLast)
							// Zeilenumbruch anhängen
							sPaste += _T("\r\n");
					}
					else
						// TAB anhängen
						sPaste += _T("	");
				}
			}
		}

		// In Zwischenablage kopieren
		if (sPaste.size() > 0)
			bOk = CopyToClipboard(sPaste);
	}

	return bOk;
}


// CopyRows
BOOL MTBLSUBCLASS::CopyRows(WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff, BOOL bHeaders /*=FALSE*/)
{
	typedef struct
	{
		HWND hWnd;
		int iID;
	}
	COLINFO;

	BOOL bOk = TRUE;

	long lContext = Ctd.TblQueryContext(m_hWndTbl), lLen;

	COLINFO aCols[MTBL_MAX_COLS];
	HWND hWndCol = NULL;
	int iCols = 0;
	//while ( MTblFindNextCol( m_hWndTbl, &hWndCol, dwColFlagsOn, dwColFlagsOff ) )
	while (FindNextCol(&hWndCol, dwColFlagsOn, dwColFlagsOff))
	{
		aCols[iCols].hWnd = hWndCol;
		aCols[iCols].iID = Ctd.TblQueryColumnID(hWndCol);
		++iCols;
		if (iCols == MTBL_MAX_COLS)
			break;
	}

	HSTRING hs = 0;
	int iCol, iMax;
	long lGrow = 2097152, lRow = TBL_MinRow;
	tstring s;
	tstring sPaste(_T(""));
	sPaste.reserve(lGrow);

	tstring sCRLF(_T("\r\n"));
	tstring sTab(_T("	"));

	if (bHeaders)
	{
		for (iCol = 0, iMax = iCols - 1; iCol <= iMax; ++iCol)
		{
			if (!GetColumnTitle(aCols[iCol].hWnd, &hs))
			{
				bOk = FALSE;
				break;
			}

			if (Ctd.HStrToTStr(hs, s))
			{

				tstring::size_type nPos = 0;
				while ((nPos = s.find(sTab, nPos)) != tstring::npos)
				{
					s.replace(nPos, 1, _T(" "));
					nPos += 1;
				}
				nPos = 0;
				while ((nPos = s.find(sCRLF, nPos)) != tstring::npos)
				{
					s.replace(nPos, 2, _T(" "));
					nPos += 1;
				}

				if (sPaste.size() + s.size() + 2 > sPaste.capacity())
					sPaste.reserve(sPaste.capacity() + lGrow);

				sPaste += s;
				if (iCol == iMax)
					sPaste += sCRLF;
				else
					sPaste += sTab;
			}
			else
			{
				bOk = FALSE;
				break;
			}			
		}
	}

	while (FindNextRow(&lRow, wRowFlagsOn, wRowFlagsOff))
	{
		if (!bOk)
			break;

		if (!Ctd.TblSetContextEx(m_hWndTbl, lRow))
		{
			bOk = FALSE;
			break;
		}

		for (iCol = 0, iMax = iCols - 1; iCol <= iMax; ++iCol)
		{
			if (!Ctd.TblGetColumnText(m_hWndTbl, aCols[iCol].iID, &hs))
			{
				bOk = FALSE;
				break;
			}

			if (Ctd.HStrToTStr(hs, s))
			{
				tstring::size_type nPos = 0;
				while ((nPos = s.find(sTab, nPos)) != tstring::npos)
				{
					s.replace(nPos, 1, _T(" "));
					nPos += 1;
				}
				nPos = 0;
				while ((nPos = s.find(sCRLF, nPos)) != tstring::npos)
				{
					s.replace(nPos, 2, _T(" "));
					nPos += 1;
				}

				if (sPaste.size() + s.size() + 2 > sPaste.capacity())
					sPaste.reserve(sPaste.capacity() + lGrow);

				sPaste += s;
				if (iCol == iMax)
					sPaste += sCRLF;
				else
					sPaste += sTab;
			}
			else
			{
				bOk = FALSE;
				break;
			}			
		}
	}

	if (hs)
		Ctd.HStringUnRef(hs);

	Ctd.TblSetContextEx(m_hWndTbl, lContext);

	if (bOk)
		bOk = CopyToClipboard(sPaste);

	return bOk;
}


// CopyShortcutPressed
int MTBLSUBCLASS::CopyShortcutPressed()
{
	if (SendMessage(m_hWndTbl, MTM_Copy, 0, 0) == MTBL_NODEFAULTACTION)
		return MTBL_NODEFAULTACTION;

	CopyCurrentSelection();

	return 0;
}


// CopyToClipboard
BOOL MTBLSUBCLASS::CopyToClipboard(tstring &s)
{
	BOOL bOk = FALSE;

	if (OpenClipboard(m_hWndTbl))
	{
		EmptyClipboard();

		HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, (s.size() * sizeof(TCHAR)) + sizeof(TCHAR));
		if (hGlobal)
		{
			LPTSTR pBuf = (LPTSTR)GlobalLock(hGlobal);
			lstrcpy(pBuf, s.c_str());
			GlobalUnlock(hGlobal);

			#ifdef UNICODE
				UINT uiFormat = CF_UNICODETEXT;
			#else
				UINT uiFormat = CF_TEXT;
			#endif

			if (SetClipboardData(uiFormat, hGlobal))
			{
				CloseClipboard();
				bOk = TRUE;
			}
			else
			{
				CloseClipboard();
				GlobalFree(hGlobal);
			}
		}
	}

	return bOk;
}


// CreateTipWnd
void MTBLSUBCLASS::CreateTipWnd()
{
	WNDCLASSEX wc;
	wc.cbSize = sizeof(WNDCLASSEX);

	BOOL bRegistered = GetClassInfoEx((HINSTANCE)ghInstance, MTBL_TIPWND_CLASS, &wc);

	if (!bRegistered)
	{
		// Fensterklasse registrieren
		wc.style = CS_OWNDC | CS_HREDRAW | CS_VREDRAW;
		if (WinVerIsXP_OG())
			wc.style |= CS_DROPSHADOW;
		wc.lpfnWndProc = TipWndProc;
		wc.cbClsExtra = 0;
		wc.cbWndExtra = 4;
		wc.hInstance = (HINSTANCE)ghInstance;
		wc.hIcon = NULL;
		wc.hCursor = LoadCursor(NULL, IDC_ARROW);
		wc.hbrBackground = NULL;
		wc.lpszMenuName = NULL;
		wc.lpszClassName = MTBL_TIPWND_CLASS;
		wc.hIconSm = NULL;

		if (!RegisterClassEx(&wc)) return;
	}

	if (IsWindow(m_hWndTbl) && !IsWindow(m_hWndTip))
	{
		m_hWndTip = CreateWindowEx(0, MTBL_TIPWND_CLASS, _T(""), WS_POPUP | WS_CLIPSIBLINGS, 0, 0, 100, 40, m_hWndTbl, NULL, (HINSTANCE)ghInstance, (LPVOID)this);
	}
}


// DarkenColor
COLORREF MTBLSUBCLASS::DarkenColor(COLORREF cr, UCHAR ucDarkening)
{
	if (!ucDarkening)
		return cr;

	UCHAR ucRed = GetRValue(cr);
	UCHAR ucGreen = GetGValue(cr);
	UCHAR ucBlue = GetBValue(cr);

	ucRed = ucRed - min(ucRed, ucDarkening);
	ucGreen = ucGreen - min(ucGreen, ucDarkening);
	ucBlue = ucBlue - min(ucBlue, ucDarkening);

	return RGB(ucRed, ucGreen, ucBlue);
}


// DeleteDescRows
int MTBLSUBCLASS::DeleteDescRows(CMTblRow * pRow, WORD wMethod)
{
	// Zeile ermitteln
	if (!pRow) return -1;
	long lRow = pRow->GetNr();

	int iDel = 0;
	if (pRow->IsParent())
	{
		// Arraydaten ermitteln
		long lLastValidPos, lUpperBound;
		CMTblRow ** RowArr = m_Rows->GetRowArray(lRow, &lUpperBound, &lLastValidPos);
		if (!RowArr) return FALSE;

		// Baum von unten nach oben?
		BOOL bBottomUp = FALSE;
		if (lRow < 0)
			bBottomUp = QueryTreeFlags(MTBL_TREE_FLAG_SPLIT_BOTTOM_UP);
		else
			bBottomUp = QueryTreeFlags(MTBL_TREE_FLAG_BOTTOM_UP);

		// Zu löschende Zeilen ermitteln, je nach Baumrichtung
		list<CMTblRow*> DelRows;
		CMTblRow * pCurRow;
		long l;
		if (bBottomUp)
		{
			for (l = 0; l <= lLastValidPos; ++l)
			{
				if (pCurRow = RowArr[l])
				{
					if (pCurRow->IsDescendantOf(pRow))
					{
						DelRows.push_back(pCurRow);
					}
				}
			}
		}
		else
		{
			for (l = lLastValidPos; l >= 0; --l)
			{
				if (pCurRow = RowArr[l])
				{
					if (pCurRow->IsDescendantOf(pRow))
					{
						DelRows.push_back(pCurRow);
					}
				}
			}
		}

		// Zeilen löschen
		list<CMTblRow*>::iterator it, itEnd = DelRows.end();
		for (it = DelRows.begin(); it != itEnd; ++it)
		{
			pCurRow = (*it);
			if (!Ctd.TblDeleteRow(m_hWndTbl, pCurRow->GetNr(), wMethod))
				return -1;
			++iDel;
		}
	}

	return iDel;
}


// DeleteRowFlagImages
void MTBLSUBCLASS::DeleteRowFlagImages()
{
	if (m_RowFlagImages)
	{
		// Nicht benutzerdefinierte Images und CMTblImg-Objekte löschen
		CMTblImg *pImg, *pImgUD;
		HANDLE hImg;
		RowFlagImgMap::iterator it, itEnd = m_RowFlagImages->end();
		for (it = m_RowFlagImages->begin(); it != itEnd; ++it)
		{
			pImg = (*it).second.first;
			pImgUD = (*it).second.second;

			if (pImg)
			{
				hImg = pImg->GetHandle();
				if (hImg)
					MImgDelete(hImg);

				delete pImg;
			}

			if (pImgUD)
				delete pImgUD;
		}

		delete m_RowFlagImages;
	}
}

// DotlineProc
void CALLBACK MTBLSUBCLASS::DotlineProc(int iX, int iY, LPARAM lParam)
{
	if (lParam)
	{
		LPMTBLDOTLINE pl = (LPMTBLDOTLINE)lParam;
		int iMod;
		if (pl->bVert)
			iMod = (pl->rCoord.top - iY) % 2;
		else
			iMod = (pl->rCoord.left - iX) % 2;
		BOOL bClr1 = (pl->bOdd ? iMod != 0 : iMod == 0);
		COLORREF clr = (bClr1 ? pl->clr1 : pl->clr2);
		if (clr != MTBL_COLOR_UNDEF)
			SetPixelV(pl->hDC, iX, iY, clr);
	}
}

// DragDropAutoStart
BOOL MTBLSUBCLASS::DragDropAutoStart(WPARAM wParam, LPARAM lParam, HWND hWndDragDrop /* = NULL*/)
{
	m_uiDragDropTimer = SetTimer(m_hWndTbl, 998, 200, (TIMERPROC)TimerProc);
	if (m_uiDragDropTimer)
	{
		m_hWndDragDrop = hWndDragDrop ? hWndDragDrop : m_hWndTbl;
		m_DragDropAutoStart_wParam = wParam;
		m_DragDropAutoStart_lParam = lParam;
		m_bDragCanAutoStart = TRUE;
		return TRUE;
	}
	else
		return FALSE;

}

// DragDropAutoStartCancel
BOOL MTBLSUBCLASS::DragDropAutoStartCancel()
{
	KillTimer(m_hWndTbl, 998);
	m_uiDragDropTimer = 0;
	SendMessage(m_hWndDragDrop, WM_LBUTTONDOWN, m_DragDropAutoStart_wParam, m_DragDropAutoStart_lParam);
	m_bDragCanAutoStart = FALSE;
	return TRUE;
}

// DragDropBegin
BOOL MTBLSUBCLASS::DragDropBegin()
{
	KillTimer(m_hWndTbl, 998);
	m_uiDragDropTimer = 0;
	m_bDragCanAutoStart = FALSE;
	BOOL bOk = Ctd.DragDropStart(m_hWndTbl);
	return bOk;
}

// DrawLinesPerRowIndicators
void MTBLSUBCLASS::DrawLinesPerRowIndicators(int iLinesPerRow)
{
	if (iLinesPerRow < 1) return;

	HDC hDC = GetDC(m_hWndTbl);
	if (!hDC) return;

	CMTblMetrics tm(m_hWndTbl);
	long lDist = tm.CalcRowHeight(iLinesPerRow);
	long lBottom, lLeft, lPos, lTop, lWidth;
	vector<LONG_PAIR> v;

	lLeft = tm.m_RowHeaderRight + 1;
	lWidth = tm.m_ClientRect.right - lLeft;

	v.push_back(vector<LONG_PAIR>::value_type(LONG_PAIR(tm.m_RowsTop, tm.m_RowsBottom)));
	if (tm.m_SplitRowsTop > 0)
		v.push_back(vector<LONG_PAIR>::value_type(LONG_PAIR(tm.m_SplitRowsTop, tm.m_SplitRowsBottom)));

	vector<LONG_PAIR>::iterator it;
	for (it = v.begin(); it != v.end(); ++it)
	{
		lTop = (*it).first;
		lBottom = (*it).second;

		for (lPos = lTop + lDist; lPos <= lBottom; lPos += lDist)
		{
			PatBlt_Hooked(hDC, lLeft, lPos, lWidth, 1, DSTINVERT);
		}
	}

	ReleaseDC(m_hWndTbl, hDC);
}

// DrawTip
void MTBLSUBCLASS::DrawTip(HWND hWnd, HDC hDC)
{
	if (!m_pTip)
		return;

	// Client-Rechteck ermitteln
	RECT r;
	GetClientRect(hWnd, &r);

	// Hintergrund
	HBRUSH hBr = CreateSolidBrush(GetDrawColor(m_pTip->GetBackColor()));
	if (hBr)
	{
		RECT rFill = r;
		BOOL bGradient = m_pTip->QueryFlags(MTBL_TIP_FLAG_GRADIENT) != 0;
		if (bGradient)
		{
			bGradient = GradientRect(hDC, rFill, RGB(255, 255, 255), GetDrawColor(m_pTip->GetBackColor()), GRFM_TB);
		}

		if (!bGradient)
		{
			rFill.right += 1;
			rFill.bottom += 1;
			FillRect(hDC, &rFill, hBr);
			DeleteObject(hBr);
		}
	}

	// Rahmen
	if (!m_pTip->QueryFlags(MTBL_TIP_FLAG_NOFRAME))
	{
		HRGN hRgn = CreateRoundRectRgn(r.left, r.top, r.right + 1, r.bottom + 1, m_pTip->GetRCEWid(), m_pTip->GetRCEHt());
		if (hRgn)
		{
			HBRUSH hBr = CreateSolidBrush(GetDrawColor(m_pTip->GetFrameColor()));
			if (hBr)
			{
				FrameRgn(hDC, hRgn, hBr, 1, 1);
				DeleteObject(hBr);
			}
			DeleteObject(hRgn);
		}
	}

	// Text
	if (m_pTip->GetText().size() > 0)
	{
		HFONT hFont = m_pTip->CreateFont(hDC);
		if (hFont)
		{
			HGDIOBJ hOldFont = SelectObject(hDC, hFont);
			SetBkMode(hDC, TRANSPARENT);
			COLORREF clrOld = SetTextColor(hDC, GetDrawColor(m_pTip->GetTextColor()));

			r.left += 4;
			r.top += 2;
			r.right -= 4;
			r.bottom -= 2;

			DrawText(hDC, m_pTip->GetText().c_str(), -1, &r, m_dwTipDrawFlags);

			SetTextColor(hDC, clrOld);
			SelectObject(hDC, hOldFont);
			DeleteObject(hFont);
		}
	}
}


// EmptyColWndProc
LRESULT CALLBACK MTBLSUBCLASS::EmptyColWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LPMTBLSUBCLASS psc = NULL;
	long lRet = 0;

	switch (uMsg)
	{
	case WM_CREATE:
		psc = (LPMTBLSUBCLASS)LPCREATESTRUCT(lParam)->lpCreateParams;
		SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)psc);
		lRet = 0;
		break;

	case WM_USER:
		psc = (LPMTBLSUBCLASS)GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (psc)
			lRet = (LRESULT)psc->m_hWndTbl;
		break;

	case TIM_QueryColumnFlags:
		if (lParam & COL_Visible)
			lRet = 1;
		break;

	default:
		lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
	}

	return lRet;
}


// EnableTipType
BOOL MTBLSUBCLASS::EnableTipType(DWORD dwType, BOOL bEnable)
{
	if (!dwType)
		return FALSE;

	m_dwEnabledTipTypes = (bEnable ? (m_dwEnabledTipTypes | dwType) : ((m_dwEnabledTipTypes & dwType) ^ m_dwEnabledTipTypes));
	return TRUE;
}


// EndColSelTrap
void MTBLSUBCLASS::EndColSelTrap(HWND hWndCol /*= NULL*/)
{
	GetColSel(GETSEL_AFTER, hWndCol);
	m_bTrappingColSelChanges = FALSE;

	BOOL bMustRedrawSelections = MustRedrawSelections();

	HRGN hRgn = NULL;
	//if ( bMustRedrawSelections || m_bCellMode )
	hRgn = GetColSelChangesRgn(hWndCol);

	if (bMustRedrawSelections && hRgn)
	{
		if (InvalidateRgn(m_hWndTbl, hRgn, FALSE))
		{
			UpdateWindow(m_hWndTbl);
		}
	}

	if (hRgn && !IsRectEmpty(&m_rFocus))
	{
		HRGN hRgnFC = CreateRectRgnIndirect(&m_rFocus);
		if (hRgnFC)
		{
			int iRet = CombineRgn(hRgn, hRgn, hRgnFC, RGN_AND);
			if (iRet != NULLREGION)
				UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);
			DeleteObject(hRgnFC);
		}
	}

	if (hRgn)
		DeleteObject(hRgn);

	if (m_bExtMsgs)
		GenerateColSelChangeMsg();
}


// EndRowSelTrap
void MTBLSUBCLASS::EndRowSelTrap(long lRow /*=TBL_Error*/, RowPtrVector *pvRows /*= NULL*/, BOOL bRemoveNew /*=FALSE*/)
{
	// Selektion "nachher" ermitteln
	int iMode = GETSEL_AFTER;
	if (bRemoveNew)
		iMode |= GETSEL_REMOVE_NEW;
	GetRowSel(iMode, lRow, pvRows);

	// Ende Verfolgung Zeilenselektionsänderungen
	m_bTrappingRowSelChanges = FALSE;

	// Müssen Selektionen neu gezeichnet werden?
	BOOL bMustRedrawSelections = MustRedrawSelections();

	// Region der Zeilenselektionsänderungen ermitteln
	HRGN hRgn = NULL;
	hRgn = GetRowSelChangesRgn(lRow);

	// Ggf. Region neu zeichnen
	if (bMustRedrawSelections && hRgn)
	{
		if (InvalidateRgn(m_hWndTbl, hRgn, FALSE))
		{
			UpdateWindow(m_hWndTbl);
		}
	}

	// Ggf. Fokus neu zeichnen
	if (hRgn && !IsRectEmpty(&m_rFocus))
	{
		HRGN hRgnFC = CreateRectRgnIndirect(&m_rFocus);
		if (hRgnFC)
		{
			int iRet = CombineRgn(hRgn, hRgn, hRgnFC, RGN_AND);
			if (iRet != NULLREGION)
				UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);
			DeleteObject(hRgnFC);
		}
	}

	// Region löschen
	if (hRgn)
		DeleteObject(hRgn);

	// Änerungsmeldungen für Zeilenselektion generieren
	if (m_bExtMsgs)
		GenerateRowSelChangeMsg();
}

// EnsureCellReadOnly
BOOL MTBLSUBCLASS::EnsureFocusCellReadOnly()
{
	HWND hWndCol = NULL;
	long lRow = TBL_Error;

	if (!Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol))
		return FALSE;

	if (hWndCol)
	{
		BOOL bCheckBoxCell = IsCheckBoxCell(hWndCol, lRow);
		BOOL bCellReadOnly = IsCellReadOnly(hWndCol, lRow);

		if (bCheckBoxCell)
			Ctd.TblKillEdit(m_hWndTbl);
		else
		{
			if (IsWindow(m_hWndEdit))
				SendMessage(m_hWndEdit, EM_SETREADONLY, bCellReadOnly, 0);

			HWND hWndExt = NULL;
			//int iCellType = Ctd.TblGetCellType( hWndCol );
			int iCellType = GetEffCellType(lRow, hWndCol);
			if (iCellType == COL_CellType_PopupEdit || iCellType == COL_CellType_DropDownList)
				hWndExt = (HWND)SendMessage(m_hWndTbl, TIM_GetExtEdit, Ctd.TblQueryColumnID(hWndCol) - 1, 0);
			if (hWndExt)
			{
				if (iCellType == COL_CellType_PopupEdit)
				{
					BOOL bReadOnly = bCellReadOnly;

#ifdef TD_POPUP_EDIT_ALWAYS_SHOWN
					BOOL bDisabled = IsCellDisabled(hWndCol, lRow);
					bReadOnly = bReadOnly || bDisabled;
#endif

					SendMessage(hWndExt, EM_SETREADONLY, bReadOnly, 0);
				}
					
				else
					EnableWindow(hWndExt, !bCellReadOnly);
			}
		}
	}

	return TRUE;
}


// ExpandRow
BOOL MTBLSUBCLASS::ExpandRow(CMTblRow * pRow, DWORD dwFlags)
{
	// Zeile ermitteln
	if (!pRow) return FALSE;
	long lRow = pRow->GetNr();

	// Split-Zeile?
	BOOL bSplitRow = (lRow < 0);

	// Hierarchie automatisch normalisieren?
	BOOL bAutoNormHierarchy = QueryTreeFlags(MTBL_TREE_FLAG_AUTO_NORM_HIER);

	// Zeichnen unterdrücken
	LockPaint();

	// Evtl. Laden der Kindzeilen anfordern
	if (pRow->QueryFlags(MTBL_ROW_CANEXPAND) && !pRow->IsParent())
	{
		LoadChildRows(lRow);
	}

	if (pRow->IsParent())
	{
		long l, lCurRow, lLastValidPos, lUpperBound;
		CMTblRow *pCurRow;
		CMTblRow **RowArr = NULL;

		// Rekursiv?
		BOOL bRecursiveDesc = (dwFlags & MTER_RECURSIVE);
		BOOL bRecursiveChilds = !bRecursiveDesc && bAutoNormHierarchy;
		if (bRecursiveDesc || bRecursiveChilds)
		{
			// Laden der Kindzeilen anfordern
			m_Rows->SetRowsInternalFlags(bSplitRow, ROW_IFLAG_LOADCHILDVISIT, FALSE);
			BOOL bLoadChildRows, bAutoNormHierarchyThisRow, bRepeat, bVisited;
			for (;;)
			{
				RowArr = m_Rows->GetRowArray(lRow, &lUpperBound, &lLastValidPos);
				if (!RowArr) return FALSE;

				bRepeat = FALSE;
				for (l = 0; l <= lLastValidPos && !bRepeat; ++l)
				{
					pCurRow = RowArr[l];
					if (pCurRow)
					{
						bVisited = pCurRow->QueryInternalFlags(ROW_IFLAG_LOADCHILDVISIT);
						if (!bVisited)
						{
							if (bRecursiveChilds ? pCurRow->GetParentRow() == pRow : pCurRow->IsDescendantOf(pRow))
							{
								if (pCurRow->QueryFlags(MTBL_ROW_CANEXPAND) && !pCurRow->IsParent())
								{
									bAutoNormHierarchyThisRow = bAutoNormHierarchy && pCurRow->QueryFlags(MTBL_ROW_HIDDEN) && SendMessage(m_hWndTbl, MTM_QueryAutoNormHierarchy, pCurRow->GetNr(), NORM_HIER_ROW_HIDE) != MTBL_NODEFAULTACTION;
									bLoadChildRows = bRecursiveDesc || bAutoNormHierarchyThisRow;

									if (bLoadChildRows)
									{
										lCurRow = pCurRow->GetNr();
										LoadChildRows(lCurRow);
										if (bRecursiveDesc)
											bRepeat = TRUE;

										if (bAutoNormHierarchyThisRow && pCurRow->IsParent())
										{
											NormalizeHierarchy(pCurRow->GetNr(), NORM_HIER_ROW_HIDE);
										}
									}
								}

								pCurRow->SetInternalFlags(ROW_IFLAG_LOADCHILDVISIT, TRUE);
								if (bRepeat)
									break;
							}
						}
					}
				}

				if (!bRepeat)
					break;
			}
		}
		else
		{
			RowArr = m_Rows->GetRowArray(lRow, &lUpperBound, &lLastValidPos);
			if (!RowArr) return FALSE;
		}

		// Kindzeilen anzeigen und Flags setzen
		for (l = 0; l <= lLastValidPos; ++l)
		{
			pCurRow = RowArr[l];
			if (pCurRow)
			{
				lCurRow = pCurRow->GetNr();

				if (bRecursiveDesc)
				{
					if (pCurRow->IsDescendantOf(pRow))
					{
						if (!pCurRow->QueryFlags(MTBL_ROW_HIDDEN))
							Ctd.TblSetRowFlags(m_hWndTbl, lCurRow, ROW_Hidden, FALSE);

						if (pCurRow->HasVisChildRows())
							pCurRow->SetFlags(MTBL_ROW_ISEXPANDED, TRUE);
					}
				}
				else
				{
					if (pCurRow->GetParentRow() == pRow)
					{
						if (!pCurRow->QueryFlags(MTBL_ROW_HIDDEN))
							Ctd.TblSetRowFlags(m_hWndTbl, lCurRow, ROW_Hidden, FALSE);
					}
				}
			}
		}

	}

	// Flags setzen
	if (pRow->HasVisChildRows())
		pRow->SetFlags(MTBL_ROW_ISEXPANDED, TRUE);
	else
		pRow->SetFlags(MTBL_ROW_CANEXPAND, FALSE);

	// Zeichnen wieder erlaubt
	UnlockPaint();

	// Tabelle neu zeichnen
	if (dwFlags & MTER_REDRAW)
		InvalidateRect(m_hWndTbl, NULL, FALSE);

	return TRUE;
}


// 
int MTBLSUBCLASS::ExportRowToHTML(LPMTBLERTHPARAMS p)
{
	BOOL bDeleteFont, bRowHeight;
	CMTblCol *pCol;
	CMTblRow *pRow;
	DWORD dwSubstColor, dwWritten;
	HDC hDC;
	HFONT hFont;
	int i, iIndentLevel, iRet = 1;
	long lColSpan, lRowSpan;
	LPMTBLMERGECELL pmc;
	tstring s, sTag;

	// Flags prüfen und Kontext setzen
	if (p->wRowFlagsOn && !Ctd.TblQueryRowFlags(m_hWndTbl, p->lRow, p->wRowFlagsOn))
		goto END;
	if (p->wRowFlagsOff && Ctd.TblQueryRowFlags(m_hWndTbl, p->lRow, p->wRowFlagsOff))
		goto END;

	// Optionen
	bRowHeight = (p->dwHTMLFlags & MTE_HTML_ROW_HEIGHT);

	// Kontext setzen
	Ctd.TblSetContextEx(m_hWndTbl, p->lRow);

	// Zeilenzeiger ermitteln
	pRow = m_Rows->GetRowPtr(p->lRow);

	// Einleitendes Zeilen-Tag schreiben
	s = _T("\r\n<tr valign=\"top\">");
	if (!WriteFile(p->hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
	{
		iRet = MTE_ERR_FILE;
		goto END;
	}

	for (i = 0; i < p->iCols; ++i)
	{
		MTBLCOLINFO &ci = p->Cols[i];

		p->pColTag->Init();

		// Evtl. Konvertierung in HTML-Spezialzeichen ausschalten
		if (QueryColumnFlags(ci.hWnd, MTBL_COL_FLAG_EXPORT_AS_TEXT))
			p->pColTag->UseSpecialCharConversion(FALSE);

		// Merged-Zelle
		if (p->pMergeCells)
			pmc = FindMergeCell(ci.hWnd, p->lRow, FMC_SLAVE, p->pMergeCells);
		else
			pmc = NULL;
		if (pmc)
		{
			continue;
		}

		// Zellenzeiger ermitteln
		if (pRow)
			pCol = pRow->m_Cols->Find(ci.hWnd);
		else
			pCol = NULL;

		// Spannings ermitteln
		lColSpan = lRowSpan = 1;
		if (p->pMergeCells)
			pmc = FindMergeCell(ci.hWnd, p->lRow, FMC_MASTER, p->pMergeCells);
		else
			pmc = NULL;
		if (pmc)
		{
			lColSpan += pmc->iMergedCols;
			lRowSpan += pmc->iMergedRows;
		}

		// Höhe
		if (bRowHeight && i == 0)
		{
			p->pColTag->m_lHeight = p->iHt * lRowSpan;
		}

		// Text
		if (!(pCol && pCol->QueryFlags(MTBL_CELL_FLAG_HIDDEN_TEXT)))
			GetEffCellText(p->lRow, ci.hWnd, p->pColTag->m_sText);

		// Font
		if (p->dwExpFlags & MTE_FONTS)
		{
			hDC = GetDC(m_hWndTbl);
			if (hDC)
			{
				hFont = GetEffCellFont(hDC, p->lRow, ci.hWnd, &bDeleteFont, 0);

				p->pColTag->m_Font.Set(hDC, hFont);

				if (bDeleteFont)
					DeleteObject(hFont);

				ReleaseDC(m_hWndTbl, hDC);
			}
		}

		// Farben
		if (p->dwExpFlags & MTE_COLORS)
		{
			p->pColTag->m_dwTextColor = GetEffCellTextColor(p->lRow, ci.hWnd);
			dwSubstColor = GetExportSubstColor(p->pColTag->m_dwTextColor);
			if (dwSubstColor != MTBL_COLOR_UNDEF)
				p->pColTag->m_dwTextColor = dwSubstColor;

			p->pColTag->m_dwBackColor = GetEffCellBackColor(p->lRow, ci.hWnd);
			dwSubstColor = GetExportSubstColor(p->pColTag->m_dwBackColor);
			if (dwSubstColor != MTBL_COLOR_UNDEF)
				p->pColTag->m_dwBackColor = dwSubstColor;
		}

		// Alignments
		int iJustify = GetEffCellTextJustify(p->lRow, ci.hWnd);
		if (iJustify == DT_RIGHT)
			p->pColTag->m_sAlign = _T("right");
		else if (iJustify == DT_CENTER)
			p->pColTag->m_sAlign = _T("center");

		int iVAlign = GetEffCellTextVAlign(p->lRow, ci.hWnd);
		if (iVAlign == DT_BOTTOM)
			p->pColTag->m_sVAlign = _T("bottom");
		else if (iVAlign == DT_VCENTER)
			p->pColTag->m_sVAlign = _T("middle");

		// Spannings
		if (lColSpan > 1)
			p->pColTag->m_lColSpan = lColSpan;
		if (lRowSpan > 1)
			p->pColTag->m_lRowSpan = lRowSpan;

		// Indent
		if (ci.hWnd == m_hWndTreeCol && m_iIndent > 0)
		{
			if (pRow)
			{
				if (QueryTreeFlags(MTBL_TREE_FLAG_FLAT_STRUCT))
					iIndentLevel = 0;
				else
					iIndentLevel = pRow->GetLevel();
			}
			else
				iIndentLevel = 0;
			p->pColTag->m_lTextIndent = iIndentLevel * m_iIndent;
		}
		else
			p->pColTag->m_lTextIndent = GetEffCellIndent(p->lRow, ci.hWnd);

		// Spalten-Tag schreiben
		s = _T("\r\n");
		p->pColTag->GetString(sTag);
		s += sTag;
		if (!WriteFile(p->hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
		{
			iRet = MTE_ERR_FILE;
			goto END;
		}

		// Zähler anpassen
		i = i + lColSpan - 1;
	}

	// Abschließendes Zeilen-Tag schreiben
	s = _T("\r\n</tr>");
	if (!WriteFile(p->hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
	{
		iRet = MTE_ERR_FILE;
		goto END;
	}

END:
	return iRet;
}

// ExportToHTML
int MTBLSUBCLASS::ExportToHTML(LPCTSTR lpsFile, DWORD dwHTMLFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff)
{
	BOOL bAnyColHdrGrp = FALSE, bAnyColHdrGrpWCH = FALSE, bDeleteFont, bWriteEnd = FALSE, bWriteTableEnd = FALSE;
	CMTblColHdrGrp *pColHdrGrp;
	CMTblColTag ColTag;
	CMTblFont f, fTbl;
	CMTblMetrics tm(m_hWndTbl);
	DWORD dwBackColorTbl, dwColorTbl, dwSubstColor, dwWritten;
	HANDLE hFile;
	HDC hDC;
	HFONT hFont;
	HSTRING hsText = 0;
	HWND hWndCol;
	int i, iCols, iPos, iRet = 1, iRetOpen, iRetRowExp, j;
	long lBorderWid, lCellPadding, lCellSpacing, lColSpan, lContext = TBL_Error, lLen, lRow;
	LPMTBLMERGECELLS pMergeCells = NULL;
	LPTSTR lpsText;
	MTBLERTHPARAMS erp;
	tstring s, sCSS, sCSSBorderCollapse, sCSSColors, sCSSFont, sTag;
	TCHAR szBuf[2048];

	// Cell-Padding/Spacing festlegen
	lCellPadding = 2;
	lCellSpacing = 0;

	// Datei erzeugen
	if (!lpsFile || lstrlen(lpsFile) < 1)
		return MTE_ERR_FILE;

	hFile = CreateFile(lpsFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		return MTE_ERR_FILE;

	// Beginn schreiben
	s = _T("");
#ifdef UNICODE
	// "Header" UTF-16, little endian
	s += 0xFEFF;
#endif
	s += _T("<html>\r\n<head>\r\n");
	s += _T("</head>\r\n<body bgcolor=\"#FFFFFF\">");
	if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
	{
		iRet = MTE_ERR_FILE;
		goto END;
	}
	bWriteEnd = TRUE;

	// Spalten ermitteln
	MTBLCOLINFO Cols[MTBL_MAX_COLS];

	hWndCol = NULL;
	iCols = 0;
	//while ( iPos = MTblFindNextCol( m_hWndTbl, &hWndCol, dwColFlagsOn, dwColFlagsOff ) )
	while (iPos = FindNextCol(&hWndCol, dwColFlagsOn, dwColFlagsOff))
	{
		if (iCols == MTBL_MAX_COLS)
			break;

		MTBLCOLINFO &ci = Cols[iCols];
		ZeroMemory(&ci, sizeof(MTBLCOLINFO));

		ci.hWnd = hWndCol;
		ci.iID = Ctd.TblQueryColumnID(hWndCol);
		//ci.lType = Ctd.TblGetCellType( hWndCol );
		ci.lType = GetEffCellType(TBL_Error, hWndCol);
		ci.lWidth = SendMessage(hWndCol, TIM_QueryColumnWidth, 0, 0);
		ci.lJustify = GetColJustify(hWndCol);
		ci.pColHdrGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
		if (ci.pColHdrGrp)
		{
			bAnyColHdrGrp = TRUE;
			if (!ci.pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS))
			{
				ci.bColHdrGrpCols = TRUE;
				bAnyColHdrGrpWCH = TRUE;
			}
		}

		iCols++;
	}

	if (iCols == 0)
		goto END;

	// Merge-Zellen ermitteln
	pMergeCells = GetMergeCells(TRUE, wRowFlagsOn, wRowFlagsOff, dwColFlagsOn, dwColFlagsOff);

	// Tabellen-Font ermitteln
	sCSSFont = _T("");
	if (dwExpFlags & MTE_FONTS)
	{
		hDC = GetDC(m_hWndTbl);
		if (hDC)
		{
			hFont = (HFONT)SendMessage(m_hWndTbl, WM_GETFONT, 0, 0);
			if (hFont)
			{
				fTbl.Set(hDC, hFont);
				GetCSSFont(fTbl, sCSSFont);
				ColTag.m_FontTbl = fTbl;
			}

			ReleaseDC(m_hWndTbl, hDC);
		}
	}

	// Tabellen-Farben ermitteln
	sCSSColors = _T("");
	if (dwExpFlags & MTE_COLORS)
	{
		dwColorTbl = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindowText);
		dwSubstColor = GetExportSubstColor(dwColorTbl);
		if (dwSubstColor != MTBL_COLOR_UNDEF)
			dwColorTbl = dwSubstColor;

		dwBackColorTbl = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);
		dwSubstColor = GetExportSubstColor(dwBackColorTbl);
		if (dwSubstColor != MTBL_COLOR_UNDEF)
			dwBackColorTbl = dwSubstColor;

		GetCSSColors(dwColorTbl, dwBackColorTbl, sCSSColors);

		ColTag.m_dwTextColorTbl = dwColorTbl;
		ColTag.m_dwBackColorTbl = dwBackColorTbl;
	}

	// Border-Collapse
	sCSSBorderCollapse = _T("");
	if (dwExpFlags & MTE_GRID)
		sCSSBorderCollapse = _T("border-collapse: collapse");

	// Einleitendes Table-Tag schreiben
	if (dwExpFlags & MTE_GRID)
		lBorderWid = 1;
	else
		lBorderWid = 0;
	wsprintf(szBuf, _T("\r\n<table border=\"%i\" cellspacing=\"%i\" cellpadding=\"%i\""), lBorderWid, lCellSpacing, lCellPadding);
	s = szBuf;
	if (!sCSSFont.empty() || !sCSSColors.empty() || !sCSSBorderCollapse.empty())
	{
		sCSS = _T(" style=\"");

		if (!sCSSFont.empty())
			sCSS.append(sCSSFont).append(_T("; "));
		if (!sCSSColors.empty())
			sCSS.append(sCSSColors).append(_T("; "));
		if (!sCSSBorderCollapse.empty())
			sCSS.append(sCSSBorderCollapse).append(_T("; "));

		sCSS.append(_T("\""));

		s.append(sCSS);
	}
	s.append(_T(">"));
	if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
	{
		iRet = MTE_ERR_FILE;
		goto END;
	}
	bWriteTableEnd = TRUE;

	// Colgroup-Tag schreiben
	s = _T("\r\n<colgroup>");
	for (i = 0; i < iCols; ++i)
	{
		wsprintf(szBuf, _T("\r\n<col width=\"%i\">"), Cols[i].lWidth);
		s += szBuf;
	}
	s += _T("\r\n</colgroup>");
	if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
	{
		iRet = MTE_ERR_FILE;
		goto END;
	}

	// Column-Header schreiben
	if (dwExpFlags & MTE_COL_HEADERS)
	{
		if (bAnyColHdrGrpWCH)
		{
			// Einleitendes Zeilen-Tag schreiben
			s = _T("\r\n<tr valign=\"top\">");
			if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
			{
				iRet = MTE_ERR_FILE;
				goto END;
			}

			pColHdrGrp = NULL;
			for (i = 0; i < iCols; ++i)
			{
				MTBLCOLINFO &ci = Cols[i];

				ColTag.Init();

				if (ci.pColHdrGrp)
				{
					pColHdrGrp = ci.pColHdrGrp;

					lColSpan = 1;
					for (j = i + 1; j < iCols; ++j)
					{
						if (Cols[j].pColHdrGrp != pColHdrGrp)
							break;
						++lColSpan;
					}

					// Text ermitteln
					ColTag.m_sText = pColHdrGrp->m_Text;

					// Horizontale Ausrichtung ermitteln
					if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_TXTALIGN_LEFT))
						ColTag.m_sAlign = _T("left");
					else if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_TXTALIGN_RIGHT))
						ColTag.m_sAlign = _T("right");
					else
						ColTag.m_sAlign = _T("center");

					// Font ermitteln
					if (dwExpFlags & MTE_FONTS)
					{
						hDC = GetDC(m_hWndTbl);
						if (hDC)
						{
							hFont = GetEffColHdrGrpFont(hDC, pColHdrGrp, &bDeleteFont);
							ColTag.m_Font.Set(hDC, hFont);

							if (bDeleteFont)
								DeleteObject(hFont);

							ReleaseDC(m_hWndTbl, hDC);
						}
					}

					// Farben ermitteln
					if (dwExpFlags & MTE_COLORS)
					{
						ColTag.m_dwTextColor = GetEffColHdrGrpTextColor(pColHdrGrp);
						dwSubstColor = GetExportSubstColor(ColTag.m_dwTextColor);
						if (dwSubstColor != MTBL_COLOR_UNDEF)
							ColTag.m_dwTextColor = dwSubstColor;

						ColTag.m_dwBackColor = GetEffColHdrGrpBackColor(pColHdrGrp);
						dwSubstColor = GetExportSubstColor(ColTag.m_dwBackColor);
						if (dwSubstColor != MTBL_COLOR_UNDEF)
							ColTag.m_dwBackColor = dwSubstColor;
					}

					// Spannings
					if (lColSpan > 1)
						ColTag.m_lColSpan = lColSpan;

					if (!ci.bColHdrGrpCols)
						ColTag.m_lRowSpan = 2;

					// Zähler anpassen
					i = i + lColSpan - 1;
				}
				else
				{
					// Text ermitteln
					if (!GetColumnTitle(ci.hWnd, &hsText))
					{
						iRet = MTE_ERR_EXCEPTION;
						goto END;
					}

					if (!Ctd.HStrToTStr(hsText, ColTag.m_sText))
					{
						iRet = MTE_ERR_EXCEPTION;
						goto END;
					}					

					// Horizontale Ausrichtung ermitteln
					if (QueryColumnHdrFlags(ci.hWnd, MTBL_COLHDR_FLAG_TXTALIGN_LEFT))
						ColTag.m_sAlign = _T("left");
					else if (QueryColumnHdrFlags(ci.hWnd, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT))
						ColTag.m_sAlign = _T("right");
					else
						ColTag.m_sAlign = _T("center");

					// Font ermitteln
					if (dwExpFlags & MTE_FONTS)
					{
						hDC = GetDC(m_hWndTbl);
						if (hDC)
						{
							hFont = GetEffColHdrFont(hDC, ci.hWnd, &bDeleteFont);
							ColTag.m_Font.Set(hDC, hFont);

							if (bDeleteFont)
								DeleteObject(hFont);

							ReleaseDC(m_hWndTbl, hDC);
						}
					}

					// Farben ermitteln
					if (dwExpFlags & MTE_COLORS)
					{
						ColTag.m_dwTextColor = GetEffColHdrTextColor(ci.hWnd);
						dwSubstColor = GetExportSubstColor(ColTag.m_dwTextColor);
						if (dwSubstColor != MTBL_COLOR_UNDEF)
							ColTag.m_dwTextColor = dwSubstColor;

						ColTag.m_dwBackColor = GetEffColHdrBackColor(ci.hWnd);
						dwSubstColor = GetExportSubstColor(ColTag.m_dwBackColor);
						if (dwSubstColor != MTBL_COLOR_UNDEF)
							ColTag.m_dwBackColor = dwSubstColor;
					}

					// Spannings
					ColTag.m_lRowSpan = 2;
				}

				// Spalten-Tag schreiben
				s = _T("\r\n");
				ColTag.GetString(sTag);
				s += sTag;
				if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
				{
					iRet = MTE_ERR_FILE;
					goto END;
				}
			}

			// Abschließendes Zeilen-Tag schreiben
			s = _T("\r\n</tr>");
			if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
			{
				iRet = MTE_ERR_FILE;
				goto END;
			}

		}

		// Einleitendes Zeilen-Tag schreiben
		s = _T("\r\n<tr align=\"center\" valign=\"top\">");
		if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
		{
			iRet = MTE_ERR_FILE;
			goto END;
		}

		pColHdrGrp = NULL;
		for (i = 0; i < iCols; ++i)
		{
			MTBLCOLINFO &ci = Cols[i];

			// Spalte auslassen?
			if (bAnyColHdrGrpWCH && (!ci.pColHdrGrp || (ci.pColHdrGrp && !ci.bColHdrGrpCols)))
				continue;

			// Column-Tag initialisieren
			ColTag.Init();

			if (ci.pColHdrGrp && !bAnyColHdrGrpWCH)
			{
				pColHdrGrp = ci.pColHdrGrp;

				lColSpan = 1;
				for (j = i + 1; j < iCols; ++j)
				{
					if (Cols[j].pColHdrGrp != pColHdrGrp)
						break;
					++lColSpan;
				}

				// Text ermitteln
				ColTag.m_sText = pColHdrGrp->m_Text;

				// Horizontale Ausrichtung ermitteln
				if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_TXTALIGN_LEFT))
					ColTag.m_sAlign = _T("left");
				else if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_TXTALIGN_RIGHT))
					ColTag.m_sAlign = _T("right");
				else
					ColTag.m_sAlign = _T("center");

				// Font ermitteln
				if (dwExpFlags & MTE_FONTS)
				{
					hDC = GetDC(m_hWndTbl);
					if (hDC)
					{
						hFont = GetEffColHdrGrpFont(hDC, pColHdrGrp, &bDeleteFont);
						ColTag.m_Font.Set(hDC, hFont);

						if (bDeleteFont)
							DeleteObject(hFont);

						ReleaseDC(m_hWndTbl, hDC);
					}
				}

				// Farben ermitteln
				if (dwExpFlags & MTE_COLORS)
				{
					ColTag.m_dwTextColor = GetEffColHdrGrpTextColor(pColHdrGrp);
					dwSubstColor = GetExportSubstColor(ColTag.m_dwTextColor);
					if (dwSubstColor != MTBL_COLOR_UNDEF)
						ColTag.m_dwTextColor = dwSubstColor;

					ColTag.m_dwBackColor = GetEffColHdrGrpBackColor(pColHdrGrp);
					dwSubstColor = GetExportSubstColor(ColTag.m_dwBackColor);
					if (dwSubstColor != MTBL_COLOR_UNDEF)
						ColTag.m_dwBackColor = dwSubstColor;
				}

				// Spannings
				if (lColSpan > 1)
					ColTag.m_lColSpan = lColSpan;

				// Zähler anpassen
				i = i + lColSpan - 1;
			}
			else
			{
				// Text ermitteln
				if (!GetColumnTitle(ci.hWnd, &hsText))
				{
					iRet = MTE_ERR_EXCEPTION;
					goto END;
				}
			
				if (!Ctd.HStrToTStr(hsText, ColTag.m_sText))
				{
					iRet = MTE_ERR_EXCEPTION;
					goto END;
				}			

				// Horizontale Ausrichtung ermitteln
				if (QueryColumnHdrFlags(ci.hWnd, MTBL_COLHDR_FLAG_TXTALIGN_LEFT))
					ColTag.m_sAlign = _T("left");
				else if (QueryColumnHdrFlags(ci.hWnd, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT))
					ColTag.m_sAlign = _T("right");
				else
					ColTag.m_sAlign = _T("center");

				// Font ermitteln
				if (dwExpFlags & MTE_FONTS)
				{
					hDC = GetDC(m_hWndTbl);
					if (hDC)
					{
						hFont = GetEffColHdrFont(hDC, ci.hWnd, &bDeleteFont);

						ColTag.m_Font.Set(hDC, hFont);

						if (bDeleteFont)
							DeleteObject(hFont);

						ReleaseDC(m_hWndTbl, hDC);
					}
				}

				// Farben ermitteln
				if (dwExpFlags & MTE_COLORS)
				{
					ColTag.m_dwTextColor = GetEffColHdrTextColor(ci.hWnd);
					dwSubstColor = GetExportSubstColor(ColTag.m_dwTextColor);
					if (dwSubstColor != MTBL_COLOR_UNDEF)
						ColTag.m_dwTextColor = dwSubstColor;

					ColTag.m_dwBackColor = GetEffColHdrBackColor(ci.hWnd);
					dwSubstColor = GetExportSubstColor(ColTag.m_dwBackColor);
					if (dwSubstColor != MTBL_COLOR_UNDEF)
						ColTag.m_dwBackColor = dwSubstColor;
				}

			}

			// Spalten-Tag schreiben
			s = _T("\r\n");
			ColTag.GetString(sTag);
			s += sTag;
			if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
			{
				iRet = MTE_ERR_FILE;
				goto END;
			}
		}

		// Abschließendes Zeilen-Tag schreiben
		s = _T("\r\n</tr>");
		if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
		{
			iRet = MTE_ERR_FILE;
			goto END;
		}
	}

	// Zeilen schreiben
	lContext = Ctd.TblQueryContext(m_hWndTbl);

	erp.hFile = hFile;
	erp.wRowFlagsOn = wRowFlagsOn;
	erp.wRowFlagsOff = wRowFlagsOff;
	erp.dwHTMLFlags = dwHTMLFlags;
	erp.dwExpFlags = dwExpFlags;
	//erp.iHt = tm.m_RowHeight;
	erp.Cols = Cols;
	erp.iCols = iCols;
	erp.pColTag = &ColTag;
	erp.pMergeCells = pMergeCells;

	// Normale Zeilen schreiben
	lRow = -1;
	while (Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, 0))
	{
		erp.lRow = lRow;
		erp.iHt = GetEffRowHeight(lRow, &tm);
		if ((iRetRowExp = ExportRowToHTML(&erp)) < 1)
		{
			iRet = iRetRowExp;
			goto END;
		}
	}

	// Split-Zeilen schreiben
	if (dwExpFlags & MTE_SPLIT_ROWS)
	{
		lRow = TBL_MinRow;
		while (Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, 0) && lRow < 0)
		{
			erp.lRow = lRow;
			if ((iRetRowExp = ExportRowToHTML(&erp)) < 1)
			{
				iRet = iRetRowExp;
				goto END;
			}
		}
	}

END:
	if (hFile != INVALID_HANDLE_VALUE)
	{
		// Tabellenende schreiben
		if (bWriteTableEnd)
		{
			s = _T("\r\n</table>");
			if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
				iRet = MTE_ERR_FILE;
		}

		// Dateiende schreiben
		if (bWriteEnd)
		{
			s = _T("\r\n</body>\r\n</html>");
			if (!WriteFile(hFile, LPCVOID(s.c_str()), s.size() * sizeof(TCHAR), &dwWritten, NULL))
				iRet = MTE_ERR_FILE;
		}

		// Datei schließen
		CloseHandle(hFile);

		// Anzeigen
		if (dwHTMLFlags & MTE_HTML_OPEN_FILE)
		{
			iRetOpen = (int)ShellExecute(m_hWndTbl, _T("open"), lpsFile, NULL, NULL, SW_SHOW);
			if (iRetOpen <= 32)
				iRet = MTE_ERR_OPEN;
		}
	}

	// Original-Kontext setzen
	if (lContext != TBL_Error)
		Ctd.TblSetContextEx(m_hWndTbl, lContext);

	// Ressourcen freigeben
	if (hsText)
		Ctd.HStringUnRef(hsText);

	if (pMergeCells)
		FreeMergeCells(pMergeCells);

	return iRet;
}


// ColWndProc
LRESULT CALLBACK MTBLSUBCLASS::ExtEditWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lRet;
	LPMTBLEXTEDITSUBCLASS psc = NULL;
	LPMTBLSUBCLASS psct = MTblGetSubClass((HWND)GetProp(hWnd, _T("m_hWndTbl")));

	if (psct)
		psc = psct->GetExtEditSubClass(hWnd);

	if (psc)
	{
		switch (uMsg)
		{
		case LB_ADDSTRING:
		case LB_INSERTSTRING:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			if (lRet >= 0)
			{
				long lRow;
				if (!psct->GetDropDownListContext(psc->hWndCol, &lRow))
					return LB_ERR;

				CMTblCellType *pct = psct->GetEffCellTypePtr(lRow, psc->hWndCol);

				if (pct && pct->IsDropDownList() && !pct->m_bSettingColumn)
				{
					CellTypeListEntries::value_type val;
					val.first = (LPCTSTR)lParam;
					val.second = 0;
					pct->m_ListEntries.insert(pct->m_ListEntries.begin() + lRet, val);
				}
			}
			break;

		case LB_DELETESTRING:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			if (lRet >= 0)
			{
				long lRow;
				if (!psct->GetDropDownListContext(psc->hWndCol, &lRow))
					return LB_ERR;

				CMTblCellType *pct = psct->GetEffCellTypePtr(lRow, psc->hWndCol);

				if (pct && pct->IsDropDownList() && !pct->m_bSettingColumn)
				{
					pct->m_ListEntries.erase(pct->m_ListEntries.begin() + wParam);
				}
			}
			break;

		case LB_RESETCONTENT:
		{
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			long lRow;
			if (!psct->GetDropDownListContext(psc->hWndCol, &lRow))
				return LB_ERR;

			CMTblCellType *pct = psct->GetEffCellTypePtr(lRow, psc->hWndCol);

			if (pct && pct->IsDropDownList() && !pct->m_bSettingColumn)
			{
				pct->m_ListEntries.clear();
			}

			break;
		}

		case LB_SETITEMDATA:
		{
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			if (lRet >= 0)
			{
				long lRow;
				if (!psct->GetDropDownListContext(psc->hWndCol, &lRow))
					return LB_ERR;

				CMTblCellType *pct = psct->GetEffCellTypePtr(lRow, psc->hWndCol);

				if (pct && pct->IsDropDownList() && !pct->m_bSettingColumn)
				{
					if (wParam < pct->m_ListEntries.size())
					{
						CellTypeListEntries::iterator it = pct->m_ListEntries.begin() + wParam;
						(*it).second = lParam;
					}
				}
			}

			break;
		}

		case WM_DESTROY:
		{
			RemoveProp(hWnd, _T("m_hWndTbl"));

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)psc->wpOld);
			psct->m_scm->erase(hWnd);
			delete psc;

			break;
		}

		case WM_MOUSEWHEEL:
		{
			TCHAR szClsName[1024] = _T("");
			WNDCLASSEX wcex;

			if (psct->m_bMWheelScroll && GetClassName(hWnd, szClsName, sizeof(szClsName) / sizeof(TCHAR) && lstrcmp(szClsName, _T("ListBox")) == 0 && GetClassInfoEx((HINSTANCE)ghInstance, szClsName, &wcex)))
			{
				lRet = CallWindowProc(wcex.lpfnWndProc, hWnd, uMsg, wParam, lParam);
			}
			else
				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			break;
		}

		default:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		}
	}
	else
		lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);

	return lRet;
}


// FetchMissingRows
int MTBLSUBCLASS::FetchMissingRows()
{
	int iFetched = 0;
	long lRow = -1;
	while (Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, ROW_InCache))
	{
		SendMessage(m_hWndTbl, TIM_FetchRow, 0, lRow);
		++iFetched;
	}
	return iFetched;
}


// FindChildClassProc
BOOL CALLBACK MTBLSUBCLASS::FindChildClassProc(HWND hWnd, LPARAM lParam)
{
	TCHAR cBuf[255];
	if (GetClassName(hWnd, cBuf, 254))
	{
		LPMTBLFSC pfsc = LPMTBLFSC(lParam);
		if (lstrcmp(cBuf, pfsc->lpsClass) == 0)
		{
			pfsc->hWndFound = hWnd;
			return FALSE;
		}
	}

	return TRUE;
}


// FindFirstChildClass
HWND MTBLSUBCLASS::FindFirstChildClass(HWND hWndParent, LPTSTR lpsClassName)
{
	MTBLFSC fsc;

	fsc.lpsClass = lpsClassName;
	fsc.hWndFound = NULL;

	EnumChildWindows(hWndParent, FindChildClassProc, LPARAM(&fsc));

	return fsc.hWndFound;
}


ItemList::iterator MTBLSUBCLASS::FindHighlightedItem(CMTblItem &item)
{
	// Suchen...
	ItemList::iterator it, itRet = m_pHLIL->end();

	for (it = m_pHLIL->begin(); it != m_pHLIL->end(); it++)
	{
		CMTblItem &item_list = (*it);
		if (*it == item)
		{
			itRet = it;
			break;
		}
	}

	return itRet;
}


// FindMergeCell
LPMTBLMERGECELL MTBLSUBCLASS::FindMergeCell(HWND hWndCol, long lRow, int iMode, LPMTBLMERGECELLS pMergeCells /* = NULL*/)
{
	MTBLMERGECELL_VECTOR v;
	long lFound = FindMergeCells(hWndCol, lRow, iMode, v, 1, pMergeCells);
	if (lFound > 0)
		return v[0];
	else
		return NULL;
}


// FindMergeCells
long MTBLSUBCLASS::FindMergeCells(HWND hWndCol, long lRow, int iMode, MTBLMERGECELL_VECTOR &v, int iFindMax /* = 0*/, LPMTBLMERGECELLS pMergeCells /* = NULL*/)
{
	if (!pMergeCells)
	{
		if (!GetMergeCells())
			return 0;

		if (!m_pMergeCells)
			return 0;

		pMergeCells = m_pMergeCells;
	}

	BOOL bIgnoreCol = (!hWndCol), bIgnoreRow = (lRow == TBL_Error);
	if (bIgnoreCol && bIgnoreRow)
		return 0;

	int iColPos = 0;
	if (!bIgnoreCol)
	{
		if (pMergeCells->dwColFlagsOn && !Ctd.TblQueryColumnFlags(hWndCol, pMergeCells->dwColFlagsOn))
			return 0;

		if (pMergeCells->dwColFlagsOff && Ctd.TblQueryColumnFlags(hWndCol, pMergeCells->dwColFlagsOff))
			return 0;

		iColPos = Ctd.TblQueryColumnPos(hWndCol);
		if (iColPos < 1)
			return 0;
	}

	if (!bIgnoreRow)
	{
		if (pMergeCells->wRowFlagsOn && !Ctd.TblQueryRowFlags(m_hWndTbl, lRow, pMergeCells->wRowFlagsOn))
			return 0;

		if (pMergeCells->wRowFlagsOff && Ctd.TblQueryRowFlags(m_hWndTbl, lRow, pMergeCells->wRowFlagsOff))
			return 0;
	}

	BOOL bColInRange, bMaster, bMasterCol, bMasterRow, bOk, bRowInRange;
	int iFound = 0;
	LPMTBLMERGECELL pmc;
	if (bIgnoreCol || bIgnoreRow)
	{
		MTBLMERGECELL_MAP::iterator it, itEnd = pMergeCells->mcm.end();
		for (it = pMergeCells->mcm.begin(); it != itEnd; ++it)
		{
			pmc = (*it).second;

			bOk = FALSE;

			bMasterCol = bIgnoreCol ? FALSE : (iColPos == pmc->iColPosFrom);
			bMasterRow = bIgnoreRow ? FALSE : (lRow == pmc->lRowFrom);

			if (iMode == FMC_MASTER)
				bOk = (bIgnoreCol || bMasterCol) && (bIgnoreRow || bMasterRow);
			else
			{
				bColInRange = bIgnoreCol ? FALSE : (iColPos >= pmc->iColPosFrom && iColPos <= pmc->iColPosTo);
				bRowInRange = bIgnoreRow ? FALSE : (lRow >= pmc->lRowFrom && lRow <= pmc->lRowTo);

				if (iMode == FMC_SLAVE)
					bOk = (bIgnoreCol || bColInRange) && (bIgnoreRow || bRowInRange) && !(bMasterCol && bMasterRow);
				else if (iMode == FMC_ANY)
					bOk = (bIgnoreCol || bColInRange) && (bIgnoreRow || bRowInRange);
			}

			if (bOk)
			{
				MTBLMERGECELL_VECTOR::iterator itv = find(v.begin(), v.end(), pmc);
				if (itv != v.end())
				{
					v.push_back(pmc);
					iFound++;
					if (iFindMax > 0 && iFound == iFindMax)
						break;
				}
			}
		}
	}
	else
	{
		MTBLMERGECELL_MAP::iterator it = pMergeCells->mcm.find(LONG_PAIR(lRow, iColPos));
		if (it != pMergeCells->mcm.end())
		{
			pmc = (*it).second;

			bOk = FALSE;
			if (iMode == FMC_ANY)
				bOk = TRUE;
			else
			{
				bMaster = (lRow == pmc->lRowFrom) && (iColPos == pmc->iColPosFrom);
				if (iMode == FMC_MASTER)
					bOk = bMaster;
				else if (iMode == FMC_SLAVE)
					bOk = !bMaster;
			}

			if (bOk)
			{
				v.push_back(pmc);
				iFound++;
			}
		}
	}

	return iFound;
}


// FindMergeRow
LPMTBLMERGEROW MTBLSUBCLASS::FindMergeRow(long lRow, int iMode, LPMTBLMERGEROWS pMergeRows /* = NULL*/)
{
	MTBLMERGEROW_VECTOR v;
	long lFound = FindMergeRows(lRow, iMode, v, 1, pMergeRows);
	if (lFound > 0)
		return v[0];
	else
		return NULL;
}


// FindMergeRows
long MTBLSUBCLASS::FindMergeRows(long lRow, int iMode, MTBLMERGEROW_VECTOR &v, int iFindMax /* = 0*/, LPMTBLMERGEROWS pMergeRows /* = NULL*/)
{
	if (!pMergeRows)
	{
		if (!GetMergeRows())
			return 0;

		if (!m_pMergeRows)
			return 0;

		pMergeRows = m_pMergeRows;
	}


	if (pMergeRows->wRowFlagsOn && !Ctd.TblQueryRowFlags(m_hWndTbl, lRow, pMergeRows->wRowFlagsOn))
		return 0;

	if (pMergeRows->wRowFlagsOff && Ctd.TblQueryRowFlags(m_hWndTbl, lRow, pMergeRows->wRowFlagsOff))
		return 0;

	BOOL bMaster, bOk;
	int iFound = 0;
	LPMTBLMERGEROW pmr;
	MTBLMERGEROW_MAP::iterator it, itEnd = pMergeRows->mrm.end();
	it = pMergeRows->mrm.find(lRow);
	if (it != itEnd)
	{
		pmr = (*it).second;

		bOk = FALSE;
		if (iMode == FMC_ANY)
			bOk = TRUE;
		else
		{
			bMaster = (lRow == pmr->lRowFrom);
			if (iMode == FMC_MASTER)
				bOk = bMaster;
			else if (iMode == FMC_SLAVE)
				bOk = !bMaster;
		}

		if (bOk)
		{
			v.push_back(pmr);
			iFound++;
		}
	}

	return iFound;
}


// FindNextCell
BOOL MTBLSUBCLASS::FindNextCell(LPLONG plRow, LPHWND phWndCol, int iDirection, BOOL bCheckFocusCell /*=FALSE*/)
{
	HWND hWndCol;
	long lRow;

	if (!plRow)
		lRow = TBL_Error;
	else
		lRow = *plRow;

	if (!phWndCol)
		hWndCol = NULL;
	else
		hWndCol = *phWndCol;


	BOOL bFound = FALSE;
	int iRet;
	LPMTBLMERGECELL pmc, pmcs;

	if (hWndCol && lRow != TBL_Error)
		pmcs = FindMergeCell(hWndCol, lRow, FMC_ANY);
	else
		pmcs = NULL;

	switch (iDirection)
	{
	case VK_RIGHT:
		while (!bFound)
		{
			//iRet = MTblFindNextCol( m_hWndTbl, &hWndCol, COL_Visible, 0 );
			iRet = FindNextCol(&hWndCol, COL_Visible, 0);
			if (iRet > 0)
			{
				if (lRow != TBL_Error)
					pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);
				else
					pmc = NULL;

				if (pmcs && pmc)
					bFound = (pmcs->hWndColFrom != pmc->hWndColFrom) || (pmcs->lRowFrom != pmc->lRowFrom);
				else
					bFound = TRUE;

				if (bFound && bCheckFocusCell && m_bCellMode && lRow != TBL_Error)
				{
					if (!CanSetFocusToCell(hWndCol, lRow))
						bFound = FALSE;
				}
			}

			if (!bFound)
			{
				if (iRet < 1)
				{
					hWndCol = NULL;
					long lNextRow = (lRow == TBL_Error ? -1 : lRow);
					if (Ctd.TblFindNextRow(m_hWndTbl, &lNextRow, 0, ROW_Hidden))
					{
						if (lRow != TBL_Error && !AreRowsOfSameArea(lRow, lNextRow))
							break;
						lRow = lNextRow;
					}
					else
						break;
				}
			}
		}
		break;

	case VK_LEFT:
		while (!bFound)
		{
			//iRet = MTblFindPrevCol( m_hWndTbl, &hWndCol, COL_Visible, 0 );
			iRet = FindPrevCol(&hWndCol, COL_Visible, 0);
			if (iRet > 0)
			{
				if (lRow != TBL_Error)
					pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);
				else
					pmc = NULL;

				if (pmcs && pmc)
					bFound = (pmcs->hWndColFrom != pmc->hWndColFrom) || (pmcs->lRowFrom != pmc->lRowFrom);
				else
					bFound = TRUE;

				if (bFound && bCheckFocusCell && m_bCellMode && lRow != TBL_Error)
				{
					if (!CanSetFocusToCell(hWndCol, lRow))
						bFound = FALSE;
				}
			}

			if (!bFound)
			{
				if (iRet < 1)
				{
					hWndCol = NULL;
					long lNextRow = (lRow == TBL_Error ? -1 : lRow);
					if (Ctd.TblFindPrevRow(m_hWndTbl, &lNextRow, 0, ROW_Hidden))
					{
						if (lRow != TBL_Error && !AreRowsOfSameArea(lRow, lNextRow))
							break;
						lRow = lNextRow;
					}
					else
						break;
				}
			}
		}
		break;

	case VK_DOWN:
		while (!bFound)
		{
			long lNextRow = (lRow == TBL_Error ? -1 : lRow);
			if (Ctd.TblFindNextRow(m_hWndTbl, &lNextRow, 0, ROW_Hidden))
			{
				if (lRow != TBL_Error && !AreRowsOfSameArea(lRow, lNextRow))
					break;
				lRow = lNextRow;
			}
			else
				break;

			if (hWndCol)
				pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);
			else
				pmc = NULL;

			if (pmcs && pmc)
				bFound = (pmcs->lRowFrom != pmc->lRowFrom);
			else
				bFound = TRUE;

			if (bFound && bCheckFocusCell && m_bCellMode && hWndCol)
			{
				if (!CanSetFocusToCell(hWndCol, lRow))
					bFound = FALSE;
			}
		}
		break;

	case VK_UP:
		while (!bFound)
		{
			long lNextRow = (lRow == TBL_Error ? -1 : lRow);
			if (Ctd.TblFindPrevRow(m_hWndTbl, &lNextRow, 0, ROW_Hidden))
			{
				if (lRow != TBL_Error && !AreRowsOfSameArea(lRow, lNextRow))
					break;
				lRow = lNextRow;
			}
			else
				break;

			if (hWndCol)
				pmc = FindMergeCell(hWndCol, lRow, FMC_ANY);
			else
				pmc = NULL;

			if (pmcs && pmc)
				bFound = (pmcs->lRowFrom != pmc->lRowFrom) && (pmc->lRowFrom == lRow);
			else if (pmc)
				bFound = (pmc->lRowFrom == lRow);
			else
				bFound = TRUE;

			if (bFound && bCheckFocusCell && m_bCellMode && hWndCol)
			{
				if (!CanSetFocusToCell(hWndCol, lRow))
					bFound = FALSE;
			}
		}
		break;
	}

	if (bFound)
	{
		*plRow = lRow;
		*phWndCol = hWndCol;
	}

	return bFound;
}


// FindNextCol
int MTBLSUBCLASS::FindNextCol(LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff)
{
	BOOL bFlagsOnOk, bFlagsOffOk, bFound = FALSE;
	int iPos;
	HWND hWndCol;

	if (*phWndCol)
	{
		iPos = m_CurCols->GetPos(*phWndCol);
		if (iPos < 0) return 0;
		iPos++;
	}
	else
		iPos = 1;

	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	int iCurCols = CurCols.size();
	for (int i = iPos - 1; i < iCurCols; ++i)
	{
		hWndCol = CurCols[i];
		if (!hWndCol)
			break;

		bFlagsOnOk = dwFlagsOn ? SendMessage(hWndCol, TIM_QueryColumnFlags, 0, dwFlagsOn) : TRUE;
		bFlagsOffOk = dwFlagsOff ? !SendMessage(hWndCol, TIM_QueryColumnFlags, 0, dwFlagsOff) : TRUE;

		if (bFlagsOnOk && bFlagsOffOk)
		{
			bFound = TRUE;
			iPos = i + 1;
			break;
		}
	}

	if (bFound)
	{
		*phWndCol = hWndCol;
		return iPos;
	}
	else
		return 0;
}


// FindNextListClip
HWND MTBLSUBCLASS::FindNextListClip(HWND hWndChildAfter)
{
	// Korrekten Klassennamen des List-Clips ermitteln
	LPTSTR lpsClipClass;
	long lType = Ctd.GetType(m_hWndTbl);
	if (lType == TYPE_ChildTable)
		lpsClipClass = CLSNAME_LISTCLIP;
	else if (lType == TYPE_TableWindow)
		lpsClipClass = CLSNAME_LISTCLIP_TBW;
	else
		return NULL;

	// Suchen
	return FindWindowEx(m_hWndTbl, hWndChildAfter, lpsClipClass, NULL);
}


// FindNextRow
BOOL MTBLSUBCLASS::FindNextRow(LPLONG plRow, WORD wFlagsOn, WORD wFlagsOff, long lMaxRow)
{
	long lRow = *plRow;
	BOOL bRet = Ctd.TblFindNextRow(m_hWndTbl, &lRow, wFlagsOn, wFlagsOff);
	if (bRet)
	{
		if (lRow > lMaxRow)
			return FALSE;
		*plRow = lRow;
	}

	return bRet;
}


// FindPaintCol
LPMTBLPAINTCOL MTBLSUBCLASS::FindPaintCol(HWND hWndCol)
{
	LPMTBLPAINTCOL pc;
	for (int i = 0; i < m_ppd->iCols; ++i)
	{
		pc = m_ppd->PaintCols[i];
		if (pc->hWnd == hWndCol)
			return pc;
	}

	return NULL;
}


// FindPaintMergeCell
LPMTBLPAINTMERGECELL MTBLSUBCLASS::FindPaintMergeCell(HWND hWndCol, long lRow, int iMode)
{
	int iPos = Ctd.TblQueryColumnPos(hWndCol);
	if (iPos < 1)
		return NULL;

	LPMTBLPAINTMERGECELL ppmc;
	vector<LPMTBLPAINTMERGECELL>::size_type i, iMax = m_ppd->PaintMergeCells.size();
	for (i = 0; i < iMax; ++i)
	{
		ppmc = m_ppd->PaintMergeCells[i];

		if (iMode == FMC_MASTER)
		{
			if (iPos == ppmc->pColFrom->iPos && lRow == ppmc->pRowFrom->GetNr())
				return ppmc;
		}
		else if (iMode == FMC_SLAVE)
		{
			if ((iPos > ppmc->pColFrom->iPos && iPos <= ppmc->pColTo->iPos) && (lRow >= ppmc->pRowFrom->GetNr() && lRow <= ppmc->pRowTo->GetNr()))
				return ppmc;
		}
		else if (iMode == FMC_ANY)
		{
			if ((iPos >= ppmc->pColFrom->iPos && iPos <= ppmc->pColTo->iPos) && (lRow >= ppmc->pRowFrom->GetNr() && lRow <= ppmc->pRowTo->GetNr()))
				return ppmc;
		}
	}

	return NULL;
}


// FindPaintMergeRowHdr
LPMTBLPAINTMERGEROWHDR MTBLSUBCLASS::FindPaintMergeRowHdr(long lRow, int iMode)
{
	LPMTBLPAINTMERGEROWHDR ppmrh;
	vector<LPMTBLPAINTMERGEROWHDR>::size_type i, iMax = m_ppd->PaintMergeRowHdrs.size();
	for (i = 0; i < iMax; ++i)
	{
		ppmrh = m_ppd->PaintMergeRowHdrs[i];

		if (iMode == FMC_MASTER)
		{
			if (lRow == ppmrh->pRowFrom->GetNr())
				return ppmrh;
		}
		else if (iMode == FMC_SLAVE)
		{
			if (lRow >= ppmrh->pRowFrom->GetNr() && lRow <= ppmrh->pRowTo->GetNr())
				return ppmrh;
		}
		else if (iMode == FMC_ANY)
		{
			if (lRow >= ppmrh->pRowFrom->GetNr() && lRow <= ppmrh->pRowTo->GetNr())
				return ppmrh;
		}
	}

	return NULL;
}


// FindPrevCol
int MTBLSUBCLASS::FindPrevCol(LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff)
{
	BOOL bFlagsOnOk, bFlagsOffOk, bFound = FALSE;
	int iPos;
	HWND hWndCol;

	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	int iCurCols = CurCols.size();

	if (*phWndCol)
	{
		iPos = m_CurCols->GetPos(*phWndCol);
		if (iPos < 0) return 0;
		iPos--;
	}
	else
	{
		iPos = CurCols.size();
	}

	for (int i = iPos - 1; i >= 0; --i)
	{
		hWndCol = CurCols[i];
		if (!hWndCol)
			break;

		bFlagsOnOk = dwFlagsOn ? SendMessage(hWndCol, TIM_QueryColumnFlags, 0, dwFlagsOn) : TRUE;
		bFlagsOffOk = dwFlagsOff ? !SendMessage(hWndCol, TIM_QueryColumnFlags, 0, dwFlagsOff) : TRUE;

		if (bFlagsOnOk && bFlagsOffOk)
		{
			bFound = TRUE;
			iPos = i + 1;
			break;
		}
	}

	if (bFound)
	{
		*phWndCol = hWndCol;
		return iPos;
	}
	else
		return 0;
}


// FindPrevRow
BOOL MTBLSUBCLASS::FindPrevRow(LPLONG plRow, WORD wFlagsOn, WORD wFlagsOff, long lMinRow)
{
	long lRow = *plRow;
	BOOL bRet = Ctd.TblFindPrevRow(m_hWndTbl, &lRow, wFlagsOn, wFlagsOff);
	if (bRet)
	{
		if (lRow < lMinRow)
			return FALSE;
		*plRow = lRow;
	}

	return bRet;
}


// FixSplitWindow
BOOL MTBLSUBCLASS::FixSplitWindow()
{
	int iLinesPerRow = 0;
	Ctd.TblQueryLinesPerRow(m_hWndTbl, &iLinesPerRow);
	if (iLinesPerRow > 1)
	{
		CMTblMetrics tm(m_hWndTbl);
		if (tm.m_SplitBarTop > 0)
		{
			int iSplitHt = tm.m_SplitRowsBottom - tm.m_SplitRowsTop + 1;
			if (iSplitHt < tm.m_RowHeight)
			{
				BOOL bDragAdjust = FALSE;
				int iSplitRows = 0;
				/*Ctd.TblQuerySplitWindow( m_hWndTbl, &iSplitRows, &bDragAdjust );
				Ctd.TblDefineSplitWindow( m_hWndTbl, 0, FALSE );
				Ctd.TblDefineSplitWindow( m_hWndTbl, iSplitRows, bDragAdjust );*/
			}
		}
	}

	return TRUE;
}


// FreeMergeCells
void MTBLSUBCLASS::FreeMergeCells(LPMTBLMERGECELLS pMergeCells /* = NULL*/)
{
	BOOL bInternal = FALSE;
	if (!pMergeCells)
	{
		if (!m_pMergeCells)
			return;
		else
		{
			pMergeCells = m_pMergeCells;
			bInternal = TRUE;
		}
	}

	LPMTBLMERGECELL pmc;
	MTBLMERGECELL_MAP::iterator it, itEnd = pMergeCells->mcm.end(), itLast, it2;
	MTBLMERGECELL_SET dset;
	for (it = pMergeCells->mcm.begin(); it != itEnd; ++it)
	{
		pmc = (*it).second;
		if (pmc)
			dset.insert(pmc);
	}

	MTBLMERGECELL_SET::iterator itd, itdEnd = dset.end();
	for (itd = dset.begin(); itd != itdEnd; ++itd)
	{
		pmc = *itd;
		if (pmc)
			delete pmc;
	}

	dset.clear();
	pMergeCells->mcm.clear();

	delete pMergeCells;
	if (bInternal)
		m_pMergeCells = NULL;
}


// FreeMergeRows
void MTBLSUBCLASS::FreeMergeRows(LPMTBLMERGEROWS pMergeRows /* = NULL*/)
{
	BOOL bInternal = FALSE;
	if (!pMergeRows)
	{
		if (!m_pMergeRows)
			return;
		else
		{
			pMergeRows = m_pMergeRows;
			bInternal = TRUE;
		}
	}

	LPMTBLMERGEROW pmr;
	MTBLMERGEROW_MAP::iterator it, itEnd = pMergeRows->mrm.end();
	MTBLMERGEROW_SET dset;
	for (it = pMergeRows->mrm.begin(); it != itEnd; ++it)
	{
		pmr = (*it).second;
		if (pmr)
			dset.insert(pmr);
	}

	MTBLMERGEROW_SET::iterator itd, itdEnd = dset.end();
	for (itd = dset.begin(); itd != itdEnd; ++itd)
	{
		pmr = *itd;
		if (pmr)
			delete pmr;
	}

	dset.clear();
	pMergeRows->mrm.clear();

	delete pMergeRows;
	if (bInternal)
		m_pMergeRows = NULL;
}


// FreePaintMergeCells
void MTBLSUBCLASS::FreePaintMergeCells()
{
	BOOL bLog = FALSE;

	if (bLog)
		LogLine(_T("c:\\temp\\mtbl.log"), _T("Beginn FreePaintMergeCells"));
	LPMTBLPAINTMERGECELL ppmc;
	PaintMergeCellsVector::iterator it, itEnd = m_ppd->PaintMergeCells.end();
	for (it = m_ppd->PaintMergeCells.begin(); it != itEnd; ++it)
	{
		ppmc = *it;
		if (ppmc)
		{
			if (bLog)
				LogLine(_T("c:\\temp\\mtbl.log"), _T("Vor delete ppmc"));
			delete ppmc;
		}
	}

	if (bLog)
		LogLine(_T("c:\\temp\\mtbl.log"), _T("Vor clear"));
	m_ppd->PaintMergeCells.clear();
}


// FreePaintMergeRowHdrs
void MTBLSUBCLASS::FreePaintMergeRowHdrs()
{
	BOOL bLog = FALSE;

	if (bLog)
		LogLine(_T("c:\\temp\\mtbl.log"), _T("Beginn FreePaintMergeRowHdrs"));
	LPMTBLPAINTMERGEROWHDR ppmrh;
	PaintMergeRowHdrsVector::iterator it, itEnd = m_ppd->PaintMergeRowHdrs.end();
	for (it = m_ppd->PaintMergeRowHdrs.begin(); it != itEnd; ++it)
	{
		ppmrh = *it;
		if (ppmrh)
		{
			if (bLog)
				LogLine(_T("c:\\temp\\mtbl.log"), _T("Vor delete ppmrh"));
			delete ppmrh;
		}
	}

	if (bLog)
		LogLine(_T("c:\\temp\\mtbl.log"), _T("Vor clear"));
	m_ppd->PaintMergeRowHdrs.clear();
}


// GetBestCellHeight
// Liefert die optimale Höhe einer Zelle
int MTBLSUBCLASS::GetBestCellHeight(HWND hWndCol, long lRow, long lCellLeadingY, long lCharBoxY, HFONT hUseFont /*=NULL*/, LPMTBLMERGECELLS pMergeCells /*=NULL*/, BOOL bCheckboxText /*=FALSE*/, BOOL bIgnoreButtons /*=FALSE*/)
{
	int iHeight;

	// Beste Höhe abfragen
	iHeight = SendMessage(m_hWndTbl, MTM_QueryBestCellHeight, (WPARAM)hWndCol, (LPARAM)lRow);
	if (iHeight > 0)
		return iHeight;
	else
		iHeight = -1;

	// Spalten-Typ ermitteln
	// int iCellType = Ctd.TblGetCellType( hWndCol );
	int iCellType = GetEffCellType(lRow, hWndCol);

	// Checkbox?
	BOOL bCheckbox = FALSE;
	if (iCellType == COL_CellType_CheckBox && !bCheckboxText)
		bCheckbox = TRUE;

	// Merge-Zelle ermitteln
	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_ANY, pMergeCells);

	// Erste und letzte Spalte setzen
	HWND hWndFirstCol, hWndLastCol;
	if (pmc)
	{
		hWndFirstCol = pmc->hWndColFrom;
		hWndLastCol = pmc->hWndColTo;
	}
	else
	{
		hWndFirstCol = hWndLastCol = hWndCol;
	}

	// Effektive Zeile setzen
	long lEffRow;
	if (pmc)
		lEffRow = pmc->lRowFrom;
	else
		lEffRow = lRow;

	// Höhe des Texts bzw. Checkbox
	SIZE s;
	CMTblCol *pCell = m_Rows->GetColPtr(lEffRow, hWndFirstCol);
	if (pCell && pCell->QueryFlags(MTBL_CELL_FLAG_HIDDEN_TEXT))
	{
		s.cx = 0;
		s.cy = 0;
	}
	else
	{
		if (bCheckbox)
		{
			s.cy = CMTblMetrics::QueryCheckBoxSize(m_hWndTbl) /*+ lCellLeadingY*/;
			if (s.cy <= lCharBoxY)
				s.cy = lCharBoxY;
			else
			{
				long lExtraLeading = max(0, ((lCellLeadingY + lCharBoxY + 1 - s.cy) / 2));
				s.cy += lExtraLeading;
			}
		}
		else
		{
			int iLinesPerRow = 0;
			Ctd.TblQueryLinesPerRow(m_hWndTbl, &iLinesPerRow);

			int iMergedRows = pmc ? pmc->iMergedRows : 0;

			BOOL bWordbreak = SendMessage(hWndFirstCol, TIM_QueryColumnFlags, 0, COL_MultilineCell) && (iLinesPerRow > 1 || iMergedRows > 0 || QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT));
			BOOL bForceMultiline = bWordbreak || QueryFlags(MTBL_FLAG_EXPAND_CRLF);
			if (!GetCellTextExtent(hWndFirstCol, lEffRow, &s, hUseFont, bWordbreak, bForceMultiline, FALSE, bIgnoreButtons, NULL, -1, pMergeCells))
				return -1;
		}
	}

	// Node-Höhe
	if (m_hWndTreeCol && hWndFirstCol == m_hWndTreeCol && QueryRowFlags(lEffRow, MTBL_ROW_CANEXPAND))
	{
		s.cy = max(m_iNodeSize/* + 2*/, s.cy);
	}

	// Leading...
	s.cy += (2 * lCellLeadingY);

	// Höhe des Bildes
	SIZE si = { 0, 0 };
	CMTblImg Img = GetEffCellImage(lEffRow, hWndFirstCol, GEI_HILI);
	if (Img.GetHandle())
	{
		if (!Img.GetSize(si))
			return -1;

		if (Img.QueryFlags(MTSI_NOLEAD_TOP))
			si.cy += lCellLeadingY;
	}

	iHeight = max(s.cy, si.cy);

	if (iHeight > 0)
	{
		// + 1 Pixel für Zeilenlinie
		++iHeight;

		// Aufrunden
		if (iHeight % 2)
			++iHeight;

		// Ggf. Aufteilung
		if (pmc)
		{
			int iMod = iHeight % (pmc->iMergedRows + 1);
			iHeight = iHeight / (pmc->iMergedRows + 1);
			if (iMod)
			{
				int iRowsChecked = 0;
				long lRowCheck = pmc->lRowFrom;
				while (TRUE)
				{
					if (lRowCheck == lRow)
					{
						++iHeight;
						break;
					}

					++iRowsChecked;
					if (iRowsChecked == iMod)
						break;

					if (!Ctd.TblFindNextRow(m_hWndTbl, &lRowCheck, 0, ROW_Hidden))
						break;
				}
			}
		}
	}

	return iHeight;
}


// GetBestCellWidth
// Liefert die optimale Breite einer Zelle
int MTBLSUBCLASS::GetBestCellWidth(HWND hWndCol, long lRow, long lCellLeadingX, HFONT hUseFont /*=NULL*/, LPMTBLMERGECELLS pMergeCells /*=NULL*/, BOOL bCheckboxText /*=FALSE*/)
{
	int iWidth;

	// Beste Breite abfragen
	iWidth = SendMessage(m_hWndTbl, MTM_QueryBestCellWidth, (WPARAM)hWndCol, (LPARAM)lRow);
	if (iWidth > 0)
		return iWidth;
	else
		iWidth = -1;

	// Spalten-Typ ermitteln
	//int iCellType = Ctd.TblGetCellType( hWndCol );
	int iCellType = GetEffCellType(lRow, hWndCol);

	// Checkbox?
	BOOL bCheckbox = FALSE;
	if (iCellType == COL_CellType_CheckBox && !bCheckboxText)
		bCheckbox = TRUE;

	// Merge-Zelle ermitteln
	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_ANY, pMergeCells);

	// Erste und letzte Spalte setzen
	HWND hWndFirstCol, hWndLastCol;
	if (pmc)
	{
		hWndFirstCol = pmc->hWndColFrom;
		hWndLastCol = pmc->hWndColTo;
	}
	else
	{
		hWndFirstCol = hWndLastCol = hWndCol;
	}

	// Effektive Zeile setzen
	long lEffRow;
	if (pmc)
		lEffRow = pmc->lRowFrom;
	else
		lEffRow = lRow;

	// Breite des Texts bzw. Checkbox ermitteln
	SIZE s;
	CMTblCol *pCell = m_Rows->GetColPtr(lEffRow, hWndFirstCol);
	if (pCell && pCell->QueryFlags(MTBL_CELL_FLAG_HIDDEN_TEXT))
	{
		s.cx = 0;
		s.cy = 0;
	}
	else
	{
		if (bCheckbox)
			s.cx = CMTblMetrics::QueryCheckBoxSize(m_hWndTbl);
		else
		{
			if (!GetCellTextExtent(hWndFirstCol, lEffRow, &s, hUseFont, FALSE, FALSE, FALSE, FALSE, NULL, -1, pMergeCells))
				return -1;
		}
	}

	// Bild und dessen Breite ermitteln
	SIZE si = { 0, 0 };
	CMTblImg Img = GetEffCellImage(lEffRow, hWndFirstCol, GEI_HILI);
	if (Img.GetHandle())
	{
		if (!Img.GetSize(si))
			return -1;
		++si.cx;
	}

	// Text und Bild nebeneinander?
	BOOL bBeside = TRUE;
	if (Img.GetHandle())
	{
		if (Img.QueryFlags(MTSI_ALIGN_HCENTER) || Img.QueryFlags(MTSI_ALIGN_TILE) || Img.QueryFlags(MTSI_NOTEXTADJ) || s.cx == 0)
			bBeside = FALSE;
	}

	if (bBeside)
	{
		iWidth = s.cx;
		iWidth += GetCellTextAreaLeft(lEffRow, hWndFirstCol, 0, lCellLeadingX, Img);
		iWidth += abs(GetCellTextAreaRight(lEffRow, hWndFirstCol, 0, lCellLeadingX, Img));
		iWidth += ((bCheckbox ? 1 : 2) * lCellLeadingX);
	}
	else
	{
		if (s.cx >= si.cx)
		{
			iWidth = s.cx;
			iWidth += GetCellAreaLeft(lEffRow, hWndFirstCol, 0);
			iWidth += (2 * lCellLeadingX);
		}
		else
		{
			iWidth = si.cx;
			iWidth += GetCellAreaLeft(lEffRow, hWndFirstCol, 0);
			if (Img.GetHandle() && !Img.QueryFlags(MTSI_NOLEAD_LEFT))
				iWidth += lCellLeadingX;
		}
	}

	if (s.cx > 0)
	{
		if (pmc)
			iWidth = iWidth / (pmc->iMergedCols + 1);
		iWidth += 1;
	}

	// Breite der Spaltenlinie berücksichtigen
	CMTblCol *pCol = m_Cols->Find(hWndLastCol);
	if (pCol && pCol->m_pVertLineDef && pCol->m_pVertLineDef->Thickness > 1)
	{
		iWidth += (pCol->m_pVertLineDef->Thickness - 1);
	}

	return iWidth;
}


// GetBestColHdrWidth
int MTBLSUBCLASS::GetBestColHdrWidth(HWND hWndCol, long lCellLeadingX, HFONT hUseFont /*=NULL*/)
{
	int iWidth;

	// Beste Breite abfragen
	iWidth = SendMessage(m_hWndTbl, MTM_QueryBestColHdrWidth, (WPARAM)hWndCol, 0);
	if (iWidth > 0)
		return iWidth;
	else
		iWidth = -1;

	// Breite des Texts ermitteln
	SIZE s;
	if (!GetColHdrTextExtent(hWndCol, &s, hUseFont)) return -1;

	// Bild und dessen Breite ermitteln
	SIZE si;
	CMTblImg Img = GetEffColHdrImage(hWndCol, GEI_HILI);
	if (Img.GetHandle())
	{
		if (!Img.GetSize(si))
			return -1;
		++si.cx;

		// Text und Bild nebeneinander?
		BOOL bBeside = TRUE;
		if (Img.QueryFlags(MTSI_ALIGN_HCENTER) || Img.QueryFlags(MTSI_ALIGN_TILE) || s.cx == 0)
			bBeside = FALSE;

		if (bBeside)
		{
			iWidth = s.cx + si.cx + (2 * lCellLeadingX);
			if (!Img.QueryFlags(MTSI_NOLEAD_LEFT))
				iWidth += lCellLeadingX;
		}
		else
		{
			if (s.cx >= si.cx)
				iWidth = s.cx + (2 * lCellLeadingX);
			else
			{
				iWidth = si.cx;
				if (!Img.QueryFlags(MTSI_NOLEAD_LEFT))
					iWidth += lCellLeadingX;
			}
		}
	}
	else
		iWidth = s.cx + (2 * lCellLeadingX);


	// Breite des Gruppentextes ermitteln, wenn das die einzige sichtbare Spalte der Gruppe ist
	CMTblColHdrGrp *pGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
	if (pGrp && IsStandaloneColHdrGrpCol(hWndCol))
	{
		SIZE siGrpText;
		if (GetColHdrGrpTextExtent(pGrp, &siGrpText, hUseFont))
		{
			int iGrpWidth = siGrpText.cx + (2 * lCellLeadingX);
			iWidth = max(iWidth, iGrpWidth);
		}

	}

	return iWidth;
}


// GetBtnBestHeight
int MTBLSUBCLASS::GetBtnBestHeight(CMTblBtn * pb, long lRow, HWND hWndCol)
{
	if (!pb) return 0;

	CMTblBtn & b = *pb;

	int iBtnHt = 0, iImgHt = 0, iTxtHt = 0;
	SIZE s;

	// Bildhöhe ermitteln
	if (MImgIsValidHandle(b.m_hImage))
	{
		if (MImgGetSize(b.m_hImage, &s))
			iImgHt = s.cy;
	}

	// Texthöhe ermitteln
	if (b.m_sText.size() > 0)
	{
		HDC hDC = GetDC(m_hWndTbl);
		if (hDC)
		{
			/*HFONT hFont;
			if ( ( b.m_dwFlags & MTBL_BTN_SHARE_FONT ) && m_hEditFont )
			hFont = m_hEditFont;
			else
			hFont = HFONT( SendMessage( m_hWndTbl, WM_GETFONT, 0, 0 ) );*/
			BOOL bDelFont = FALSE;
			HFONT hFont = GetBtnFont(pb, lRow, hWndCol, &bDelFont);
			HGDIOBJ hOldFont = SelectObject(hDC, hFont);
			if (GetTextExtentPoint32(hDC, b.m_sText.c_str(), b.m_sText.size(), &s))
			{
				iTxtHt = s.cy;
			}
			SelectObject(hDC, hOldFont);
			ReleaseDC(m_hWndTbl, hDC);

			if (bDelFont)
				DeleteObject(hFont);
		}
	}

	iBtnHt = max(iImgHt, iTxtHt) + /*2*/4;

	return iBtnHt;
}


// GetBtnFont
HFONT MTBLSUBCLASS::GetBtnFont(CMTblBtn * pb, long lRow, HWND hWndCol, LPBOOL pbMustDelete)
{
	if (!pb) return NULL;
	if (!pbMustDelete) return NULL;

	*pbMustDelete = FALSE;

	HFONT hBtnFont = NULL;

	if (pb->m_dwFlags & MTBL_BTN_SHARE_FONT)
	{
		HDC hDC = GetDC(m_hWndTbl);
		if (hDC)
		{
			if (lRow != TBL_Error)
				hBtnFont = GetEffCellFont(hDC, lRow, hWndCol, pbMustDelete);
			else
				hBtnFont = GetEffColHdrFont(hDC, hWndCol, pbMustDelete);
			ReleaseDC(m_hWndTbl, hDC);
		}
	}
	else
	{
		hBtnFont = (HFONT)SendMessage(m_hWndTbl, WM_GETFONT, 0, 0);
	}

	return hBtnFont;
}


// GetBtnHeight
/*int MTBLSUBCLASS::GetBtnHeight( HWND hWndTbl )
{
LPMTBLSUBCLASS psc = MTblGetSubClass( hWndTbl );
if ( psc )
{
return psc->m_iBtnHt;
}
else
{
MTBLGETMETRICS m;
SendMessage( hWndTbl, TIM_GetMetrics, 0, (LPARAM)&m );
return m.ptCharBox.y;
}
}*/


// GetBtnRect
BOOL MTBLSUBCLASS::GetBtnRect(HWND hWndCol, long lRow, CMTblMetrics * ptm, LPRECT pRect)
{
	BOOL bCellBtn = hWndCol && lRow != TBL_Error;
	BOOL bColHdrBtn = hWndCol && lRow == TBL_Error;

	CMTblBtn *pb = NULL;

	if (bCellBtn)
		pb = GetEffCellBtn(hWndCol, lRow);
	else if (bColHdrBtn)
		pb = m_Cols->GetHdrBtn(hWndCol);

	if (!pb) return FALSE;

	CMTblMetrics tm;
	if (!ptm)
	{
		tm.Get(m_hWndTbl);
		ptm = &tm;
	}

	RECT rBounds;

	if (bCellBtn)
	{
		RECT rCell;
		int iRet = GetCellRect(hWndCol, lRow, ptm, &rCell, MTGR_INCLUDE_MERGED);
		if (iRet == MTGR_RET_ERROR) return FALSE;

		rBounds = rCell;

		RECT rDDBtn;
		if (GetDDBtnRect(hWndCol, lRow, ptm, &rDDBtn))
			rBounds.right = rDDBtn.left;
	}
	else if (bColHdrBtn)
	{
		RECT rColHdr;
		int iRet = GetColHdrRect(hWndCol, rColHdr, MTGR_EXCLUDE_COLHDRGRP);
		if (iRet == MTGR_RET_ERROR) return FALSE;

		rBounds = rColHdr;
	}

	return CalcBtnRect(pb, hWndCol, lRow, ptm, &rBounds, NULL, NULL, pRect);
}


// GetCellAreaLeft
int MTBLSUBCLASS::GetCellAreaLeft(long lRow, HWND hWndCol, long lCellLeft)
{
	int iIndent = GetEffCellIndent(lRow, hWndCol);
	if (iIndent)
		lCellLeft += iIndent;

	return lCellLeft;
}


// GetCellFocusRect
BOOL MTBLSUBCLASS::GetCellFocusRect(long lRow, HWND hWndCol, LPRECT pRect)
{
	if (!pRect)
		return FALSE;

	SetRectEmpty(pRect);

	int iLeft = 0, iTop = 0, iRight = 0, iBottom = 0;
	int iRet;
	if (hWndCol && lRow != TBL_Error)
	{
		LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_SLAVE);
		if (pmc)
		{
			hWndCol = pmc->hWndColFrom;
			lRow = pmc->lRowFrom;
		}

		iRet = MTblGetCellRectEx(hWndCol, lRow, &iLeft, &iTop, &iRight, &iBottom, MTGR_INCLUDE_MERGED);
		if (iRet == MTGR_RET_ERROR || iRet == MTGR_RET_INVISIBLE)
			return FALSE;

	}
	else if (!hWndCol && lRow == TBL_Error)
	{
		/*if ( Ctd.TblAnyRows( m_hWndTbl, 0, ROW_Hidden ) )
		return FALSE;*/

		HWND hWndCol = NULL;
		//if ( !MTblFindNextCol( m_hWndTbl, &hWndCol, COL_Visible, 0 ) )
		if (!FindNextCol(&hWndCol, COL_Visible, 0))
			return FALSE;

		CMTblMetrics tm(m_hWndTbl);
		POINT pt, ptVis;

		//iRet = MTblGetColXCoord( m_hWndTbl, hWndCol, &tm, &pt, &ptVis );
		iRet = GetColXCoord(hWndCol, &tm, &pt, &ptVis);
		if (iRet == MTGR_RET_ERROR || iRet == MTGR_RET_INVISIBLE)
			return FALSE;

		iLeft = pt.x;
		iRight = pt.y;
		iTop = tm.m_RowsTop;
		iBottom = iTop + tm.m_RowHeight - 1;
	}

	SetRect(pRect, iLeft - m_lFocusFrameOffsX - m_lFocusFrameThick, iTop - m_lFocusFrameOffsY - m_lFocusFrameThick, iRight + m_lFocusFrameOffsX + m_lFocusFrameThick, iBottom + m_lFocusFrameOffsY + m_lFocusFrameThick);
	return TRUE;
}


// GetCellTextAreaLeft
int MTBLSUBCLASS::GetCellTextAreaLeft(long lRow, HWND hWndCol, long lCellLeft, long lCellLeadingX, CMTblImg & Image, BOOL bButton /*=FALSE*/)
{
	long lLeft = GetCellAreaLeft(lRow, hWndCol, lCellLeft);
	if (bButton)
	{
		CMTblBtn * pBtn = GetEffCellBtn(hWndCol, lRow);
		if (pBtn && pBtn->IsLeft())
		{
			int iBtnWid = CalcBtnWidth(pBtn, hWndCol, lRow);
			lLeft += iBtnWid;
		}
	}

	if (Image.GetHandle() && !(Image.QueryFlags(MTSI_ALIGN_HCENTER) || Image.QueryFlags(MTSI_ALIGN_RIGHT) || Image.QueryFlags(MTSI_ALIGN_TILE) || Image.QueryFlags(MTSI_NOTEXTADJ)))
	{
		SIZE s;
		if (Image.GetSize(s))
		{
			int iOffset;
			if (Image.QueryFlags(MTSI_NOLEAD_LEFT) && !IsIndented(hWndCol, lRow))
				iOffset = 0;
			else
				iOffset = lCellLeadingX;

			lLeft += (iOffset + s.cx + 1);
		}
	}

	return lLeft;
}


// GetCellTextAreaRight
int MTBLSUBCLASS::GetCellTextAreaRight(long lRow, HWND hWndCol, long lCellRight, long lCellLeadingX, CMTblImg & Image, BOOL bButton /*=FALSE*/, BOOL bDDButton /*=FALSE*/)
{
	BOOL bMargin = FALSE;
	long lRight = lCellRight;

	if (bDDButton)
	{
		int iCellType = GetEffCellType(lRow, hWndCol);
		if (iCellType == COL_CellType_DropDownList)
		{
			lRight -= MTBL_DDBTN_WID;
			bMargin = TRUE;
		}
	}

	if (bButton)
	{
		CMTblBtn * pBtn = GetEffCellBtn(hWndCol, lRow);
		if (pBtn && !pBtn->IsLeft())
		{
			int iBtnWid = CalcBtnWidth(pBtn, hWndCol, lRow);
			lRight -= iBtnWid;
			bMargin = TRUE;
		}
	}

	if (Image.GetHandle() && Image.QueryFlags(MTSI_ALIGN_RIGHT) && !(Image.QueryFlags(MTSI_ALIGN_TILE) || Image.QueryFlags(MTSI_NOTEXTADJ)))
	{
		SIZE s;
		if (Image.GetSize(s))
		{
			lRight -= s.cx;
			bMargin = TRUE;
		}
	}

	if (bMargin)
		lRight -= 1;

	return lRight;
}


// GetCellTextAreaXCoord
// Liefert Zellentext-Koordinaten der X-Achse ( links / rechts ). Berücksichtigt Merging.
BOOL MTBLSUBCLASS::GetCellTextAreaXCoord(long lRow, HWND hWndCol, CMTblMetrics *ptm, LPMTBLMERGECELLS pMergeCells, LPPOINT ppt, LPPOINT pptVis, BOOL bIgnoreButtons /*=FALSE*/)
{
	// Parameter-Check?
	if (!(ppt && pptVis))
		return FALSE;

	CMTblMetrics tm;
	if (!ptm)
	{
		tm.Get(m_hWndTbl);
		ptm = &tm;
	}

	// Koordinaten ermitteln
	POINT ptColX, ptColVisX;
	int iRet = GetColXCoord(hWndCol, ptm, &ptColX, &ptColVisX);

	CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);

	BOOL bButton = !bIgnoreButtons && MustShowBtn(hWndCol, lRow, TRUE, TRUE);
	BOOL bDDButton = !bIgnoreButtons && MustShowDDBtn(hWndCol, lRow, TRUE, TRUE);

	int iLeft = GetCellTextAreaLeft(lRow, hWndCol, ptColX.x, ptm->m_CellLeadingX, Img, bButton);
	HWND hWndLastMergedCol = GetLastMergedCell(hWndCol, lRow, pMergeCells);
	if (hWndLastMergedCol)
	{
		POINT ptLColX, ptLColVisX;
		iRet = GetColXCoord(hWndLastMergedCol, ptm, &ptLColX, &ptLColVisX);
		ptColX.y = ptLColX.y;
		ptColVisX.y = ptLColVisX.y;
	}

	int iRight = GetCellTextAreaRight(lRow, hWndCol, ptColX.y, ptm->m_CellLeadingX, Img, bButton, bDDButton);

	BOOL bWordWrap = GetCellWordWrap(hWndCol, lRow);
	int iJustify = GetEffCellTextJustify(lRow, hWndCol);

	CalcCellTextAreaXCoord(iLeft, iRight, ptm->m_CellLeadingX, iJustify, bWordWrap, *ppt);

	pptVis->x = max(ppt->x, ptColVisX.x);
	pptVis->y = min(ppt->y, ptColVisX.y);

	return TRUE;
}


// GetCellTextExtent
// Liefert die Ausmaße eines Zellentexts
BOOL MTBLSUBCLASS::GetCellTextExtent(HWND hWndCol, long lRow, LPSIZE pSize, HFONT hUseFont /*= NULL*/, BOOL bWordbreak /*= FALSE*/, BOOL bForceMultiline /*=FALSE*/, BOOL bEndEllipsis /*= FALSE*/, BOOL bIgnoreButtons /*=FALSE*/, LPCTSTR lpsUseText /*=NULL*/, long lTextAreaWid /*=-1*/, LPMTBLMERGECELLS pMergeCells /*= NULL*/)
{
	// Effektiven Text ermitteln
	long lLen;
	LPCTSTR lpsText;
	tstring sText;

	if (lpsUseText)
	{
		lpsText = lpsUseText;
		lLen = lstrlen(lpsText);
	}
	else
	{
		if (!GetEffCellText(lRow, hWndCol, sText)) return FALSE;
		lpsText = sText.c_str();
		lLen = sText.size();
	}


	// Ausmaße ermitteln
	if (lpsText && lLen > 0)
	{
		// Gerätekontext ermitteln
		HDC hDC = GetDC(m_hWndTbl);
		if (!hDC) return FALSE;

		// Font ermitteln
		BOOL bDeleteFont = FALSE;
		HFONT hFont;
		if (hUseFont)
			hFont = hUseFont;
		else
		{
			DWORD dwFlags = 0;
			if (hWndCol == m_hWndHLCol && lRow == m_lHLRow)
				dwFlags |= GEF_HYPERLINK;
			if (AnyHighlightedItems())
				dwFlags |= GEF_HILI;

			hFont = GetEffCellFont(hDC, lRow, hWndCol, &bDeleteFont, dwFlags);
		}

		// Zellenfont selektieren
		HGDIOBJ hOldFont = SelectObject(hDC, hFont);

		// Ausmaße ermitteln
		UINT uiFlags = DT_NOPREFIX | DT_CALCRECT;
		BOOL bMultiline;
		if (bForceMultiline)
			bMultiline = TRUE;
		else
			bMultiline = QueryFlags(MTBL_FLAG_EXPAND_CRLF) || GetCellWordWrap(hWndCol, lRow);
		if (!bMultiline)
			uiFlags |= DT_SINGLELINE;
		if (QueryFlags(MTBL_FLAG_EXPAND_TABS))
			uiFlags |= DT_EXPANDTABS;

		RECT r = { 0, 0, 0, 0 };
		if (bWordbreak && !(uiFlags & DT_SINGLELINE))
		{
			if (lTextAreaWid < 0)
			{
				POINT pt, ptVis;
				if (GetCellTextAreaXCoord(lRow, hWndCol, NULL, pMergeCells, &pt, &ptVis, bIgnoreButtons))
					lTextAreaWid = (pt.y - pt.x);
			}
			if (lTextAreaWid >= 0)
			{
				r.right = lTextAreaWid;
				uiFlags |= DT_WORDBREAK;
			}
		}
		if (bEndEllipsis)
		{
			if (lTextAreaWid < 0)
			{
				POINT pt, ptVis;
				if (GetCellTextAreaXCoord(lRow, hWndCol, NULL, pMergeCells, &pt, &ptVis, bIgnoreButtons))
					lTextAreaWid = (pt.y - pt.x);
			}
			if (lTextAreaWid >= 0)
			{
				RECT rEE = { 0, 0, 0, 0 };
				rEE.right = lTextAreaWid;
				long lLenEE = lLen + 4 + 1;
				LPTSTR lpsTextEE = new (nothrow)TCHAR[lLenEE];
				if (lpsTextEE)
				{
					lstrcpy(lpsTextEE, lpsText);
					DRAWTEXTPARAMS dtp = { sizeof(DRAWTEXTPARAMS), 8, 0, 0, 0 };
					DrawTextEx(hDC, lpsTextEE, lLen, &rEE, uiFlags | DT_END_ELLIPSIS | DT_MODIFYSTRING, &dtp);
					sText = lpsTextEE;
					lLen = sText.size();
					for (; lLen < dtp.uiLengthDrawn; ++lLen)
					{
						sText += _T(".");
					}
					lpsText = sText.c_str();
					delete[] lpsTextEE;
				}
			}
		}

		DrawText(hDC, lpsText, lLen, &r, uiFlags);
		pSize->cx = r.right - r.left;
		pSize->cy = r.bottom - r.top;

		// Alten Font selektieren
		SelectObject(hDC, hOldFont);

		// Zellenfont löschen?
		if (bDeleteFont)
			DeleteObject(hFont);

		// Gerätekontext freigeben
		ReleaseDC(m_hWndTbl, hDC);
	}
	else if (lpsText)
		pSize->cx = pSize->cy = 0;
	else
		return FALSE;

	return TRUE;
}


// GetColJustify
int MTBLSUBCLASS::GetColJustify(HWND hWndCol)
{
	if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_RightJustify))
		return DT_RIGHT;
	else if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_CenterJustify))
		return DT_CENTER;
	else
		return DT_LEFT;
}


// GetCellFontMetrics
// Liefert die Metriken eines Zellenfonts
BOOL MTBLSUBCLASS::GetCellFontMetrics(long lRow, HWND hWndCol, LPTEXTMETRIC pMetrics)
{
	BOOL bOk = FALSE;
	HDC hDC = GetDC(m_hWndTbl);
	if (hDC)
	{
		// Font der Zelle ermitteln
		BOOL bDeleteCellFont = FALSE;
		HFONT hCellFont = GetEffCellFont(hDC, lRow, hWndCol, &bDeleteCellFont);

		// Font selektieren
		HGDIOBJ hOldFont = SelectObject(hDC, hCellFont);

		// Metriken ermitteln
		bOk = GetTextMetrics(hDC, pMetrics);

		// Alten Font selektieren
		SelectObject(hDC, hOldFont);

		// Font löschen?
		if (bDeleteCellFont)
			DeleteObject(hCellFont);

		// Gerätekontext freigeben
		ReleaseDC(m_hWndTbl, hDC);
	}

	return bOk;
}


// GetCellIndentLevel
// Liefert Einrückungs-Ebene einer Zelle
long MTBLSUBCLASS::GetCellIndentLevel(long lRow, HWND hWndCol)
{
	CMTblCol *pCell = m_Rows->GetColPtr(lRow, hWndCol);
	return (pCell ? pCell->m_IndentLevel : 0);
}



// GetCellRect
int MTBLSUBCLASS::GetCellRect(HWND hWndCol, long lRow, CMTblMetrics * ptm, LPRECT pRect, DWORD dwFlags)
{
	// Gültige Zeile?
	if (!IsRow(lRow)) return MTGR_RET_ERROR;

	// Tabellenmaße ok?
	CMTblMetrics tm;
	if (!ptm)
	{
		tm.Get(m_hWndTbl);
		ptm = &tm;
	}

	// Nur sichtbaren Teil?
	BOOL bVisiblePart = dwFlags & MTGR_VISIBLE_PART;

	// Inkl. Merge-Zellen?
	HWND hWndLastMCol = NULL;
	//int iMergedCols = 0, iMergedRows = 0;
	long lLastMRow = TBL_Error;
	BOOL bInclMergedCols = FALSE, bInclMergedRows = FALSE;
	if (dwFlags & MTGR_INCLUDE_MERGED /*&& IsMergeCell( hWndCol, lRow )*/)
	{
		LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_MASTER);
		if (pmc)
		{
			hWndLastMCol = pmc->hWndColTo;
			lLastMRow = pmc->lRowTo;

			bInclMergedCols = (pmc->iMergedCols > 0);
			bInclMergedRows = (pmc->iMergedRows > 0);
		}
		/*if ( GetLastMergedCellEx( hWndCol, lRow, hWndLastMCol, lLastMRow, &iMergedCols, &iMergedRows ) )
		{
		bInclMergedCols = ( iMergedCols > 0 );
		bInclMergedRows = ( iMergedRows > 0 );
		}*/
	}

	// Oben und unten ermitteln
	POINT ptRow = { 0, 0 }, ptRowVis = { 0, 0 };

	int iRowsRet;
	//int iRowRet = MTblGetRowYCoord( m_hWndTbl, lRow, ptm, &ptRow, &ptRowVis, bInclMergedRows ? FALSE : bVisiblePart );
	int iRowRet = GetRowYCoord(lRow, ptm, &ptRow, &ptRowVis, bInclMergedRows ? FALSE : bVisiblePart);
	if (iRowRet == MTGR_RET_ERROR)
		return MTGR_RET_ERROR;

	if (bInclMergedRows)
	{
		POINT pt, ptVis;
		//int iRowRet2 = MTblGetRowYCoord( m_hWndTbl, lLastMRow, ptm, &pt, &ptVis, FALSE );
		int iRowRet2 = GetRowYCoord(lLastMRow, ptm, &pt, &ptVis, FALSE);
		if (iRowRet2 == MTGR_RET_ERROR)
			return MTGR_RET_ERROR;

		long lRowsTop = (lRow < 0) ? ptm->m_SplitRowsTop : ptm->m_RowsTop;
		long lRowsBottom = (lRow < 0) ? ptm->m_SplitRowsBottom : ptm->m_RowsBottom;

		ptRow.y = pt.y;

		if (ptRow.x >= lRowsBottom || ptRow.y < lRowsTop)
			iRowsRet = MTGR_RET_INVISIBLE;
		else
		{
			// Sichtbaren Bereich setzen
			ptRowVis.x = max(ptRow.x, lRowsTop);
			ptRowVis.y = min(ptRow.y, lRowsBottom - 1);

			// Nur teilweise sichtbar?
			if (ptRowVis.x != ptRow.x || ptRowVis.y != ptRow.y)
				iRowsRet = MTGR_RET_PARTLY_VISIBLE;
			else
				iRowsRet = MTGR_RET_VISIBLE;
		}
	}
	else
		iRowsRet = iRowRet;

	if (iRowsRet == MTGR_RET_INVISIBLE && bVisiblePart)
		return MTGR_RET_INVISIBLE;

	// Links und Rechts ermitteln
	int iColRet;
	POINT ptCol = { 0, 0 }, ptColVis = { 0, 0 };
	if (bInclMergedCols)
	{
		BOOL bAllFullyVisible = TRUE, bVisibleFound = FALSE;
		HWND hWndMCol = hWndCol;
		int iRet;
		POINT ptMCol = { 0, 0 }, ptMColVis = { 0, 0 };
		while (TRUE)
		{
			//iRet = MTblGetColXCoord( m_hWndTbl, hWndMCol, ptm, &ptMCol, &ptMColVis );
			iRet = GetColXCoord(hWndMCol, ptm, &ptMCol, &ptMColVis);
			if (iRet == MTGR_RET_ERROR)
				return MTGR_RET_ERROR;

			if (iRet != MTGR_RET_INVISIBLE)
			{
				if (!bVisibleFound)
					ptColVis.x = ptMColVis.x;
				ptColVis.y = ptMColVis.y;

				bVisibleFound = TRUE;
				if (iRet != MTGR_RET_VISIBLE)
					bAllFullyVisible = FALSE;
			}
			else
				bAllFullyVisible = FALSE;

			if (hWndMCol == hWndCol)
				ptCol.x = ptMCol.x;

			if (hWndMCol == hWndLastMCol)
			{
				ptCol.y = ptMCol.y;
				break;
			}

			//if ( !MTblFindNextCol( m_hWndTbl, &hWndMCol, COL_Visible, 0 ) )
			if (!FindNextCol(&hWndMCol, COL_Visible, 0))
				return MTGR_RET_ERROR;
		}

		if (!bVisibleFound)
			iColRet = MTGR_RET_INVISIBLE;
		else
			iColRet = bAllFullyVisible ? MTGR_RET_VISIBLE : MTGR_RET_PARTLY_VISIBLE;
	}
	else
	{
		//iColRet = MTblGetColXCoord( m_hWndTbl, hWndCol, ptm, &ptCol, &ptColVis );
		iColRet = GetColXCoord(hWndCol, ptm, &ptCol, &ptColVis);
		if (iColRet == MTGR_RET_ERROR)
			return MTGR_RET_ERROR;
	}

	if (iColRet == MTGR_RET_INVISIBLE && bVisiblePart)
		return MTGR_RET_INVISIBLE;

	if (bVisiblePart)
	{
		pRect->left = ptColVis.x;
		pRect->top = ptRowVis.x;
		pRect->right = ptColVis.y;
		pRect->bottom = ptRowVis.y;
	}
	else
	{
		pRect->left = ptCol.x;
		pRect->top = ptRow.x;
		pRect->right = ptCol.y;
		pRect->bottom = ptRow.y;
	}

	if (iColRet == MTGR_RET_VISIBLE && iRowsRet == MTGR_RET_VISIBLE)
		return MTGR_RET_VISIBLE;
	else if (iColRet == MTGR_RET_INVISIBLE || iRowsRet == MTGR_RET_INVISIBLE)
		return MTGR_RET_INVISIBLE;
	else
		return MTGR_RET_PARTLY_VISIBLE;
}


// GetCellWordWrap
// Ermittelt, ob eine Zelle einen automatischen Zeilenumbruch hat
BOOL MTBLSUBCLASS::GetCellWordWrap(HWND hWndCol, long lRow, CMTblMetrics *ptm /*= NULL*/)
{
	if (!hWndCol)
		return FALSE;

	if (!SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_MultilineCell))
		return FALSE;

	int iLinesPerRow = 0;
	if (ptm)
		iLinesPerRow = ptm->m_LinesPerRow;
	else
		Ctd.TblQueryLinesPerRow(m_hWndTbl, &iLinesPerRow);

	if (iLinesPerRow > 1 || QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT))
		return TRUE;
	else
	{
		LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_MASTER);
		if (pmc && pmc->iMergedRows > 0)
			return TRUE;
		else
			return FALSE;
	}
}


// GetColHdrFontMetrics
// Liefert die Metriken eines Column-Header-Fonts
BOOL MTBLSUBCLASS::GetColHdrFontMetrics(HWND hWndCol, LPTEXTMETRIC pMetrics)
{
	BOOL bOk = FALSE;
	HDC hDC = GetDC(m_hWndTbl);
	if (hDC)
	{
		// Font ermitteln
		BOOL bDeleteFont = FALSE;
		DWORD dwFlags = 0;
		if (AnyHighlightedItems())
			dwFlags |= GEF_HILI;
		HFONT hFont = GetEffColHdrFont(hDC, hWndCol, &bDeleteFont, dwFlags);

		// Font selektieren
		HGDIOBJ hOldFont = SelectObject(hDC, hFont);

		// Metriken ermitteln
		bOk = GetTextMetrics(hDC, pMetrics);

		// Alten Font selektieren
		SelectObject(hDC, hOldFont);

		// Font löschen?
		if (bDeleteFont)
			DeleteObject(hFont);

		// Gerätekontext freigeben
		ReleaseDC(m_hWndTbl, hDC);
	}

	return bOk;
}


// GetColHdrGrpHeight
// Liefert die Höhe einer Spalten-Header-Gruppe
int MTBLSUBCLASS::GetColHdrGrpHeight(CMTblColHdrGrp *pGrp)
{
	int iHt = 0;

	if (pGrp)
	{
		if (pGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS))
		{
			CMTblMetrics tm(m_hWndTbl);
			iHt = tm.m_HeaderHeight - 1;
		}
		else
		{
			int iLines = pGrp->GetNrOfTextLines();
			HDC hDC = GetDC(m_hWndTbl);
			if (hDC)
			{
				HGDIOBJ hOldFont = SelectObject(hDC, m_hFont);
				SIZE si;
				if (GetTextExtentPoint32(hDC, _T("A"), 1, &si))
				{
					iHt = (iLines * si.cy) + 2;
				}
				SelectObject(hDC, hOldFont);
				ReleaseDC(m_hWndTbl, hDC);
			}
		}
	}

	return iHt;
}


// GetColHdrGrpTextExtent
// Liefert die Ausmaße eines Spaltenheadergruppentexts
BOOL MTBLSUBCLASS::GetColHdrGrpTextExtent(CMTblColHdrGrp *pGrp, LPSIZE pSize, HFONT hUseFont /*=NULL*/)
{
	if (!pGrp)
		return FALSE;

	// Text ermitteln
	tstring sText(pGrp->m_Text);

	// Gerätekontext ermitteln
	HDC hDC = GetDC(m_hWndTbl);
	if (!hDC) return FALSE;

	// Font ermitteln
	BOOL bDeleteFont = FALSE;
	HFONT hFont;
	if (hUseFont)
		hFont = hUseFont;
	else
	{
		DWORD dwFlags = 0;
		if (AnyHighlightedItems())
			dwFlags |= GEF_HILI;
		hFont = GetEffColHdrGrpFont(hDC, pGrp, &bDeleteFont, dwFlags);
	}

	// Font selektieren
	HGDIOBJ hOldFont = SelectObject(hDC, hFont);

	// Ausmaße ermitteln
	UINT uiFlags = DT_NOPREFIX | DT_CALCRECT;
	if (QueryFlags(MTBL_FLAG_EXPAND_TABS))
		uiFlags |= DT_EXPANDTABS;
	RECT r;
	SetRectEmpty(&r);
	DrawText(hDC, sText.c_str(), sText.size(), &r, uiFlags);
	pSize->cx = r.right - r.left;
	pSize->cy = r.bottom - r.top;

	// Alten Font selektieren
	SelectObject(hDC, hOldFont);

	// Font löschen?
	if (bDeleteFont)
		DeleteObject(hFont);

	// Gerätekontext freigeben
	ReleaseDC(m_hWndTbl, hDC);


	return TRUE;
}


// GetColHdrRect
// Liefert das rechteck eines Spaltenkopfes
int MTBLSUBCLASS::GetColHdrRect(HWND hWndCol, RECT &r, DWORD dwFlags)
{
	// Spalte unsichtbar?
	if (!SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible)) return MTGR_RET_INVISIBLE;

	// Tabellen-Maße holen
	CMTblMetrics tm(m_hWndTbl);

	// Links und Rechts ermitteln
	POINT ptCol = { 0, 0 }, ptColVis = { 0, 0 };
	int iColRet = GetColXCoord(hWndCol, &tm, &ptCol, &ptColVis);
	if (iColRet == MTGR_RET_ERROR) return MTGR_RET_ERROR;

	// Oben und unten ermitteln
	int iRowRet;
	POINT ptHdr = { 0, 0 }, ptHdrVis = { 0, 0 };
	if (!QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER) && !Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_HideColumnHeaders))
	{
		if (dwFlags & MTGR_EXCLUDE_COLHDRGRP)
		{
			CMTblColHdrGrp *pGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
			if (pGrp)
			{
				int iHt = GetColHdrGrpHeight(pGrp);
				ptHdr.x = iHt + 1;
				ptHdrVis.x = min(ptHdr.x, tm.m_ClientRect.bottom - 1);
			}
		}
		ptHdr.y = tm.m_HeaderHeight - 1;
		ptHdrVis.y = min(ptHdr.y, tm.m_ClientRect.bottom - 1);

		if (ptHdrVis.y < 0 || ptHdrVis.x < ptHdr.x)
			iRowRet = MTGR_RET_INVISIBLE;
		else if (ptHdrVis.y < ptHdr.y || ptHdrVis.x > ptHdr.x)
			iRowRet = MTGR_RET_PARTLY_VISIBLE;
		else
			iRowRet = MTGR_RET_VISIBLE;
	}
	else
		iRowRet = MTGR_RET_INVISIBLE;

	if (dwFlags & MTGR_VISIBLE_PART)
	{
		if (iColRet == MTGR_RET_INVISIBLE || iRowRet == MTGR_RET_INVISIBLE)
			return MTGR_RET_INVISIBLE;
		r.left = ptColVis.x;
		r.top = ptHdrVis.x;
		r.right = ptColVis.y;
		r.bottom = ptHdrVis.y;
	}
	else
	{
		r.left = ptCol.x;
		r.top = ptHdr.x;
		r.right = ptCol.y;
		r.bottom = ptHdr.y;
	}

	if (iColRet == MTGR_RET_VISIBLE && iRowRet == MTGR_RET_VISIBLE)
		return MTGR_RET_VISIBLE;
	else if (iColRet == MTGR_RET_INVISIBLE || iRowRet == MTGR_RET_INVISIBLE)
		return MTGR_RET_INVISIBLE;
	else
		return MTGR_RET_PARTLY_VISIBLE;
}


// GetColHdrText
// Liefert den Text eines Spaltenheaders
BOOL MTBLSUBCLASS::GetColHdrText(HWND hWndCol, tstring &sText, int iMaxLen)
{
	BOOL bOk = FALSE;
	HSTRING hsText = 0;

	sText = _T("");

	Ctd.TblGetColumnTitle(hWndCol, &hsText, iMaxLen);
	bOk = Ctd.HStrToTStr(hsText, sText);

	if (hsText)
		Ctd.HStringUnRef(hsText);

	return bOk;
}


// GetColHdrTextExtent
// Liefert die Ausmaße eines Spaltenheadertexts
BOOL MTBLSUBCLASS::GetColHdrTextExtent(HWND hWndCol, LPSIZE pSize, HFONT hUseFont /*= NULL*/)
{
	// Text ermitteln
	tstring sText;
	if (!GetColHdrText(hWndCol, sText))
		return FALSE;

	// Gerätekontext ermitteln
	HDC hDC = GetDC(m_hWndTbl);
	if (!hDC) return FALSE;

	// Font ermitteln
	BOOL bDeleteFont = FALSE;
	HFONT hFont;
	if (hUseFont)
		hFont = hUseFont;
	else
	{
		DWORD dwFlags = 0;
		if (AnyHighlightedItems())
			dwFlags |= GEF_HILI;
		hFont = GetEffColHdrFont(hDC, hWndCol, &bDeleteFont, dwFlags);
	}

	// Font selektieren
	HGDIOBJ hOldFont = SelectObject(hDC, hFont);

	// Ausmaße ermitteln
	UINT uiFlags = DT_NOPREFIX | DT_CALCRECT;
	if (QueryFlags(MTBL_FLAG_EXPAND_TABS))
		uiFlags |= DT_EXPANDTABS;
	RECT r;
	SetRectEmpty(&r);
	DrawText(hDC, sText.c_str(), sText.size(), &r, uiFlags);
	pSize->cx = r.right - r.left;
	pSize->cy = r.bottom - r.top;

	// Alten Font selektieren
	SelectObject(hDC, hOldFont);

	// Font löschen?
	if (bDeleteFont)
		DeleteObject(hFont);

	// Gerätekontext freigeben
	ReleaseDC(m_hWndTbl, hDC);


	return TRUE;
}


// GetColRect
int MTBLSUBCLASS::GetColRect(HWND hWndCol, LPRECT pr)
{
	// Tabellenmaße
	CMTblMetrics tm(m_hWndTbl);

	// X-Koordinaten der Spalte ermitteln
	POINT pt, ptVis;
	//int iRet = MTblGetColXCoord( m_hWndTbl, hWndCol, &tm, &pt, &ptVis );
	int iRet = GetColXCoord(hWndCol, &tm, &pt, &ptVis);
	if (iRet <= MTGR_RET_INVISIBLE) return iRet;

	pr->left = ptVis.x;
	pr->top = tm.m_ClientRect.top;
	pr->right = ptVis.y;
	pr->bottom = tm.m_ClientRect.bottom;

	return iRet;
}



// GetColSel
long MTBLSUBCLASS::GetColSel(int iMode, HWND hWndCol /*= NULL*/)
{
	CMTblCol *pCol;
	DWORD dwFlag = ((iMode & GETSEL_BEFORE) ? COL_IFLAG_SELECTED_BEFORE : COL_IFLAG_SELECTED_AFTER);
	long lCount = 0;

	m_Cols->SetColsInternalFlags(dwFlag, FALSE);

	if (!hWndCol)
	{
		int iPos = 1;
		while (hWndCol = Ctd.TblGetColumnWindow(m_hWndTbl, iPos, COL_GetPos))
		{
			if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Selected))
			{
				pCol = m_Cols->EnsureValid(hWndCol);
				if (!pCol)
					return -1;
				pCol->SetInternalFlags(dwFlag, TRUE);
				++lCount;
			}
			++iPos;
		}
	}
	else
	{
		pCol = m_Cols->EnsureValid(hWndCol);
		if (!pCol)
			return -1;
		if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Selected))
		{
			pCol->SetInternalFlags(dwFlag, TRUE);
			lCount = 1;
		}
	}

	return lCount;
}


// GetColSelChanges
// Liefert die Änderungen der Spaltenselektion
long MTBLSUBCLASS::GetColSelChanges(HARRAY haCols, HARRAY haChanges)
{
	// Check Arrays
	long lDims;
	if (!Ctd.ArrayDimCount(haCols, &lDims)) return -1;
	if (lDims > 1) return -1;

	if (!Ctd.ArrayDimCount(haChanges, &lDims)) return -1;
	if (lDims > 1) return -1;

	// Untergrenzen der Arrays ermitteln
	long lLowCols;
	if (!Ctd.ArrayGetLowerBound(haCols, 1, &lLowCols)) return -1;

	long lLowChanges;
	if (!Ctd.ArrayGetLowerBound(haChanges, 1, &lLowChanges)) return -1;

	// Änderungen ermitteln
	BOOL bSel, bSelAfter, bSelBefore;
	CMTblCol *pCol;
	long lCount = 0;
	NUMBER n;

	// Ermitteln
	MTblColList::iterator it, itEnd = m_Cols->m_List.end();

	for (it = m_Cols->m_List.begin(); it != itEnd; ++it)
	{
		pCol = &(*it).second;
		bSelBefore = pCol->QueryInternalFlags(COL_IFLAG_SELECTED_BEFORE) != 0;
		bSelAfter = pCol->QueryInternalFlags(COL_IFLAG_SELECTED_AFTER) != 0;
		if (bSelBefore != bSelAfter)
		{
			if (!Ctd.MDArrayPutHandle(haCols, pCol->m_hWnd, lCount + lLowCols))
				return -1;

			if (bSelBefore && !bSelAfter)
				bSel = FALSE;
			else
				bSel = TRUE;

			if (!Ctd.CvtLongToNumber(bSel, &n))
				return -1;
			if (!Ctd.MDArrayPutNumber(haChanges, &n, lCount + lLowChanges))
				return -1;

			++lCount;
		}
	}

	return lCount;
}


// GetColSelChangesRgn
HRGN MTBLSUBCLASS::GetColSelChangesRgn(HWND hWndCol /*= NULL*/)
{
	BOOL bAny = FALSE, bSelAfter, bSelBefore;
	int iRet;
	MTblColList::iterator it, itEnd = m_Cols->m_List.end();
	RECT r;

	HRGN hRgn = CreateRectRgn(0, 0, 0, 0);
	if (!hRgn)
		return NULL;

	if (!hWndCol)
	{
		for (it = m_Cols->m_List.begin(); it != itEnd; ++it)
		{
			bSelBefore = (*it).second.QueryInternalFlags(COL_IFLAG_SELECTED_BEFORE) != 0;
			bSelAfter = (*it).second.QueryInternalFlags(COL_IFLAG_SELECTED_AFTER) != 0;
			if (bSelBefore != bSelAfter)
			{
				iRet = GetColRect((*it).second.m_hWnd, &r);
				if (iRet != MTGR_RET_ERROR && iRet != MTGR_RET_INVISIBLE)
				{
					r.right += 1;
					r.bottom += 1;
					HRGN hRgnCol = CreateRectRgnIndirect(&r);
					if (hRgnCol)
					{
						iRet = CombineRgn(hRgn, hRgn, hRgnCol, RGN_OR);
						DeleteObject(hRgnCol);
						bAny = TRUE;
					}
				}
			}
		}
	}
	else
	{
		CMTblCol *pCol = m_Cols->Find(hWndCol);
		if (pCol)
		{
			bSelBefore = pCol->QueryInternalFlags(COL_IFLAG_SELECTED_BEFORE) != 0;
			bSelAfter = pCol->QueryInternalFlags(COL_IFLAG_SELECTED_AFTER) != 0;
			if (bSelBefore != bSelAfter)
			{
				iRet = GetColRect(hWndCol, &r);
				if (iRet != MTGR_RET_ERROR && iRet != MTGR_RET_INVISIBLE)
				{
					r.right += 1;
					r.bottom += 1;
					HRGN hRgnCol = CreateRectRgnIndirect(&r);
					if (hRgnCol)
					{
						iRet = CombineRgn(hRgn, hRgn, hRgnCol, RGN_OR);
						DeleteObject(hRgnCol);
						bAny = TRUE;
					}
				}
			}

		}
	}

	if (bAny)
		return hRgn;
	else
	{
		DeleteObject(hRgn);
		return NULL;
	}
}


// GetColSubClass
LPMTBLCOLSUBCLASS MTBLSUBCLASS::GetColSubClass(HWND hWndCol)
{
	SubClassMap::iterator it = m_scm->find(hWndCol);
	if (it != m_scm->end())
	{
		return (LPMTBLCOLSUBCLASS)(*it).second;
	}

	return NULL;
}


int MTBLSUBCLASS::GetColXCoord(HWND hWndCol, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis)
{
	// Tabellenmaße ok?
	CMTblMetrics tm;
	if (!ptm)
	{
		tm.Get(m_hWndTbl);
		ptm = &tm;
	}

	// Aktuelle Spalten ermitteln
	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	CURCOLS_WIDTH_VECTOR & CurColsWid = m_CurCols->GetVectorWidth();
	int iCurCols = CurCols.size();

	// Koordinaten der Pseudo-Spalte rechts neben der letzten echten Spalte ermitteln?
	if (!hWndCol)
	{
		POINT ptpc = { -1, ptm->m_ColumnsRight };

		HWND hWndLastCol = NULL;
		//MTblFindPrevCol( m_hWndTbl, &hWndLastCol, COL_Visible, 0 );
		FindPrevCol(&hWndLastCol, COL_Visible, 0);
		if (hWndLastCol)
		{
			POINT pt, ptVis;
			int iRet = GetColXCoord(hWndLastCol, ptm, &pt, &ptVis);
			if ((iRet == MTGR_RET_VISIBLE || iRet == MTGR_RET_PARTLY_VISIBLE) && pt.y < ptpc.y)
				ptpc.x = pt.y + 1;
		}
		else
		{
			ptpc.x = tm.m_ColumnsLeft;
		}

		if (ptpc.x >= 0 && ptpc.x < ptpc.y)
		{
			*ppt = ptpc;
			*pptVis = ptpc;
			return MTGR_RET_VISIBLE;
		}
		else
			return MTGR_RET_INVISIBLE;
	}

	// Spalte sichtbar?
	if (!SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible)) return MTGR_RET_INVISIBLE;

	// Ermitteln
	BOOL bFound = FALSE;
	HWND hWndCurCol = NULL;
	int iColPos;
	int iCurCol;
	ppt->x = ptm->m_RowHeaderRight;
	ppt->y = ppt->x;
	/*while( iColPos = MTblFindNextCol( hWndTbl, &hWndCurCol, COL_Visible, 0 ) )*/
	for (iCurCol = 0; iCurCol < iCurCols; ++iCurCol)
	{
		hWndCurCol = CurCols[iCurCol];

		if (!SendMessage(hWndCurCol, TIM_QueryColumnFlags, 0, COL_Visible))
			continue;

		ppt->x = ppt->y;
		//ppt->y += SendMessage( hWndCurCol, TIM_QueryColumnWidth, 0, 0 );
		ppt->y += CurColsWid[iCurCol];

		if (hWndCurCol == hWndCol)
		{
			bFound = TRUE;
			iColPos = iCurCol + 1;
			break;
		}
	}
	if (!bFound) return MTGR_RET_INVISIBLE;
	ppt->y -= 1;

	// Verankerte Spalte?
	BOOL bLocked = (iColPos <= ptm->m_LockedColumns && ppt->y <= ptm->m_LockedColumnsRight);

	// Keine verankerte Spalte?
	if (!bLocked)
	{
		// 1 Pixel nach rechts "schieben", wenn verankerte Spalten sichtbar sind
		if (ptm->m_LockedColumnsWidth > 0)
		{
			// Spalte davor verankert?
			HWND hWndColBef = hWndCurCol;
			//int iColPosBef = MTblFindPrevCol( m_hWndTbl, &hWndColBef, COL_Visible, 0 );
			int iColPosBef = FindPrevCol(&hWndColBef, COL_Visible, 0);
			if (iColPosBef > 0 && iColPosBef <= ptm->m_LockedColumns)
			{
				ppt->y += 1;
			}
			else
			{
				ppt->x += 1;
				ppt->y += 1;
			}
		}
	}

	// Scroll-Offset abziehen
	if (!bLocked)
	{
		ppt->x -= ptm->m_ColumnOrigin;
		ppt->y -= ptm->m_ColumnOrigin;

		// Nur sichtbaren Teil ermitteln
		pptVis->x = max(ppt->x, ptm->m_ColumnsLeft);
		pptVis->y = min(ppt->y, ptm->m_ColumnsRight - 1);

		// Überhaupt ein Teil sichtbar?
		if (pptVis->x >= ptm->m_ColumnsRight) return MTGR_RET_INVISIBLE;
		if (pptVis->y < ptm->m_ColumnsLeft) return MTGR_RET_INVISIBLE;

		// Nur teilweise sichtbar?
		if (pptVis->x != ppt->x || pptVis->y != ppt->y) return MTGR_RET_PARTLY_VISIBLE;
	}
	else
	{
		// Nur sichtbaren Teil ermitteln
		pptVis->x = max(ppt->x, ptm->m_LockedColumnsLeft);
		pptVis->y = min(ppt->y, ptm->m_LockedColumnsRight);

		// Überhaupt ein Teil sichtbar?
		if (pptVis->x >= ptm->m_LockedColumnsRight) return MTGR_RET_INVISIBLE;
		if (pptVis->y < ptm->m_LockedColumnsLeft) return MTGR_RET_INVISIBLE;

		// Nur teilweise sichtbar?
		if (pptVis->x != ppt->x || pptVis->y != ppt->y) return MTGR_RET_PARTLY_VISIBLE;
	}

	return MTGR_RET_VISIBLE;
}


// GetColumnTitle
// Liefert den Titel einer Spalte
BOOL MTBLSUBCLASS::GetColumnTitle(HWND hWndCol, LPHSTRING phsTitle)
{
	// Textlänge ermitteln
	long lLen = SendMessage(hWndCol, WM_GETTEXTLENGTH, 0, 0);

	// Text ermitteln
	tstring sText;
	if (!GetColHdrText(hWndCol, sText, lLen)) return FALSE;

	// Anhängende CRLFs entfernen, wenn Gruppenspalte
	CMTblColHdrGrp *pGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
	if (pGrp)
	{
		int iTrCRLF = GetTrailingCRLFCount(sText.c_str());
		if (iTrCRLF > 0)
		{
			sText = sText.substr(0, sText.size() - (iTrCRLF * 2));
		}
	}

	// HString setzen
	if (!Ctd.InitLPHSTRINGParam(phsTitle, Ctd.BufLenFromStrLen(sText.size()))) return FALSE;

	BOOL bOk = TRUE;
	LPTSTR lpsBuf = Ctd.StringGetBuffer(*phsTitle, &lLen);
	if (lpsBuf)
		lstrcpy(lpsBuf, sText.c_str());
	else
		bOk = FALSE;
	Ctd.HStrUnlock(*phsTitle);	

	return bOk;
}


// GetCornerRect
// Liefert das Corner-Rechteck
BOOL MTBLSUBCLASS::GetCornerRect(LPRECT pr)
{
	CMTblMetrics tm(m_hWndTbl);

	pr->left = 0;
	pr->top = 0;
	pr->right = tm.m_RowHeaderRight;
	pr->bottom = tm.m_HeaderHeight;

	return TRUE;
}


// GetCRLFCount
int MTBLSUBCLASS::GetCRLFCount(LPCTSTR lpsStr)
{
	int iCount = 0;

	if (lpsStr)
	{
		int iLen = lstrlen(lpsStr);
		if (iLen > 0)
		{
			int iPos = iLen - 2;
			while (iPos >= 0)
			{
				if (lpsStr[iPos] == 13 && lpsStr[iPos + 1] == 10)
					++iCount;
				iPos -= 2;
			}
		}
	}

	return iCount;
}


// GetCSSColors
BOOL MTBLSUBCLASS::GetCSSColors(DWORD dwColor, DWORD dwBackColor, tstring &sCSS)
{
	TCHAR szBuf[255];

	sCSS = _T("");

	sCSS.append(_T("color: #"));
	wsprintf(szBuf, _T("%02X%02X%02X"), GetRValue(dwColor), GetGValue(dwColor), GetBValue(dwColor));
	sCSS.append(szBuf);

	sCSS.append(_T("; background-color: #"));
	wsprintf(szBuf, _T("%02X%02X%02X"), GetRValue(dwBackColor), GetGValue(dwBackColor), GetBValue(dwBackColor));
	sCSS.append(szBuf);

	return TRUE;
}


// GetCSSFont
BOOL MTBLSUBCLASS::GetCSSFont(CMTblFont &f, tstring &sCSS)
{
	TCHAR szBuf[255];
	long lFontSize, lFontEnh;
	tstring sFontName;

	sCSS = _T("");

	f.Get(sFontName, lFontSize, lFontEnh);

	sCSS.append(_T("font-family: "));
	sCSS.append(sFontName);

	sCSS.append(_T("; "));
	wsprintf(szBuf, _T("font-size: %dpt"), lFontSize);
	sCSS.append(szBuf);

	if (lFontEnh & FONT_EnhItalic)
		sCSS.append(_T("; font-style: italic"));

	if (lFontEnh & FONT_EnhBold)
		sCSS.append(_T("; font-weight: bold"));

	BOOL bUnderline = lFontEnh & FONT_EnhUnderline;
	BOOL bStrikeOut = lFontEnh & FONT_EnhStrikeOut;
	if (bUnderline || bStrikeOut)
	{
		sCSS.append(_T("; text-decoration:"));
		if (bUnderline)
			sCSS.append(_T(" underline"));
		if (bStrikeOut)
			sCSS.append(_T(" line-through"));
	}

	return TRUE;
}


// GetDDBtnRect
BOOL MTBLSUBCLASS::GetDDBtnRect(HWND hWndCol, long lRow, CMTblMetrics * ptm, LPRECT pRect)
{
	if (!MustShowDDBtn(hWndCol, lRow, TRUE, TRUE)) return FALSE;

	CMTblMetrics tm;
	if (!ptm)
	{
		tm.Get(m_hWndTbl);
		ptm = &tm;
	}

	RECT rCell;
	int iRet = GetCellRect(hWndCol, lRow, ptm, &rCell, MTGR_INCLUDE_MERGED);
	if (iRet == MTGR_RET_ERROR) return FALSE;

	return CalcDDBtnRect(hWndCol, lRow, ptm, &rCell, NULL, NULL, pRect);
}


// GetDefaultColor
COLORREF MTBLSUBCLASS::GetDefaultColor(int iColorIndex)
{
	COLORREF clr = RGB(0, 0, 0);

#ifdef TDTHEMES

	if (QueryFlags(MTBL_FLAG_IGNORE_TD_THEME_COLORS))
	{
		clr = GetDefaultColorNoTdTheme(iColorIndex);
	}
	else
	{
		int iTheme = Ctd.ThemeGet();

		CTheme &th = theTheme();
		BOOL bWinTheme = (th.CanUse() && th.IsAppThemed());

		switch (iColorIndex)
		{
		case GDC_HDR_BKGND:
			switch (iTheme)
			{
			case THEME_Office2007_R1:
				clr = RGB(205, 205, 205);
				break;
			case THEME_Office2007_R2_LunaBlue:
			case THEME_Office2007_R3_LunaBlue:
				clr = RGB(191, 219, 255);
				break;
			case THEME_Office2007_R2_Obsidian:
			case THEME_Office2007_R3_Obsidian:
				clr = RGB(173, 174, 189);
				break;
			case THEME_Office2007_R2_Silver:
			case THEME_Office2007_R3_Silver:
				clr = RGB(244, 247, 251);
				break;
#ifdef TDTHEMES_OFFICE2010
			case THEME_Office2010_R1:
				clr = RGB(205, 205, 205);
				break;
			case THEME_Office2010_R2_Blue:
				clr = RGB(192, 211, 235);
				break;
			case THEME_Office2010_R2_Silver:
				clr = RGB(233, 235, 238);
				break;
			case THEME_Office2010_R2_Black:
				clr = RGB(188, 188, 188);
				break;
#endif
#ifdef TDTHEMES_OFFICE2013
			case THEME_Office2013:
				clr = RGB(194, 213, 242);
				break;
#endif
#ifdef TDTHEMES_OFFICE2016
			case THEME_Office2016_R2_Metro:
				clr = RGB(231, 231, 235);
				break;
#endif
			default:
				clr = GetDefaultColorNoTdTheme(iColorIndex);
			}
			break;

		case GDC_HDR_LINE:
			switch (iTheme)
			{
			case THEME_Office2007_R1:
				clr = RGB(124, 101, 101);
				break;
			case THEME_Office2007_R2_LunaBlue:
			case THEME_Office2007_R3_LunaBlue:
				clr = RGB(0, 107, 245);
				break;
			case THEME_Office2007_R2_Obsidian:
			case THEME_Office2007_R3_Obsidian:
				clr = RGB(79, 82, 119);
				break;
			case THEME_Office2007_R2_Silver:
			case THEME_Office2007_R3_Silver:
				clr = RGB(74, 127, 197);
				break;
#ifdef TDTHEMES_OFFICE2010
			case THEME_Office2010_R1:
				clr = RGB(124, 101, 101);
				break;
			case THEME_Office2010_R2_Blue:
				clr = RGB(50, 109, 183);
				break;
			case THEME_Office2010_R2_Silver:
				clr = RGB(102, 124, 156);
				break;
			case THEME_Office2010_R2_Black:
				clr = RGB(113, 93, 93);
				break;
#endif
#ifdef TDTHEMES_OFFICE2013
			case THEME_Office2013:
				clr = RGB(102, 124, 156);
				break;
#endif
#ifdef TDTHEMES_OFFICE2016
			case THEME_Office2016_R2_Metro:
				clr = RGB(102, 124, 156);
				break;
#endif
			default:
				clr = GetDefaultColorNoTdTheme(iColorIndex);
			}
			break;

		case GDC_HDR_3DHL:
			switch (iTheme)
			{
			case THEME_Office2007_R1:
				clr = RGB(242, 242, 242);
				break;
			case THEME_Office2007_R2_LunaBlue:
			case THEME_Office2007_R3_LunaBlue:
				clr = RGB(239, 246, 255);
				break;
			case THEME_Office2007_R2_Obsidian:
			case THEME_Office2007_R3_Obsidian:
				clr = RGB(234, 234, 238);
				break;
			case THEME_Office2007_R2_Silver:
			case THEME_Office2007_R3_Silver:
				clr = RGB(252, 253, 254);
				break;
#ifdef TDTHEMES_OFFICE2010
			case THEME_Office2010_R1:
				clr = RGB(242, 242, 242);
				break;
			case THEME_Office2010_R2_Blue:
				clr = RGB(239, 243, 250);
				break;
			case THEME_Office2010_R2_Silver:
				clr = RGB(249, 250, 250);
				break;
			case THEME_Office2010_R2_Black:
				clr = RGB(238, 238, 238);
				break;
#endif
#ifdef TDTHEMES_OFFICE2013
			case THEME_Office2013:
				clr = RGB(249, 250, 250);
				break;
#endif
#ifdef TDTHEMES_OFFICE2016
			case THEME_Office2016_R2_Metro:
				clr = RGB(249, 250, 250);
				break;
#endif
			default:
				clr = GetDefaultColorNoTdTheme(iColorIndex);
			}
			break;

		case GDC_CELL_LINE:
			switch (iTheme)
			{
			case THEME_Office2000:
			case THEME_NativeXP:
				if (bWinTheme)
					clr = GetDefaultColorNoTdTheme(iColorIndex);
				else
					clr = RGB(128, 128, 128);
				break;
			case THEME_Office2007_R1:
				clr = RGB(153, 153, 153);
				break;
			case THEME_Office2007_R2_LunaBlue:
			case THEME_Office2007_R3_LunaBlue:
				clr = RGB(79, 156, 254);
				break;
			case THEME_Office2007_R2_Obsidian:
			case THEME_Office2007_R3_Obsidian:
				clr = RGB(122, 124, 148);
				break;
			case THEME_Office2007_R2_Silver:
			case THEME_Office2007_R3_Silver:
				clr = RGB(153, 180, 217);
				break;
#ifdef TDTHEMES_OFFICE2010
			case THEME_Office2010_R1:
				clr = RGB(153, 153, 153);
				break;
			case THEME_Office2010_R2_Blue:
				clr = RGB(110, 154, 209);
				break;
			case THEME_Office2010_R2_Silver:
				clr = RGB(166, 174, 186);
				break;
			case THEME_Office2010_R2_Black:
				clr = RGB(141, 141, 141);
				break;
#endif
#ifdef TDTHEMES_OFFICE2013
			case THEME_Office2013:
				clr = RGB(166, 174, 186);
				break;
#endif
#ifdef TDTHEMES_OFFICE2016
			case THEME_Office2016_R2_Metro:
				clr = RGB(166, 174, 186);
				break;
#endif
			default:
				if (bWinTheme)
					clr = GetDefaultColorNoTdTheme(iColorIndex);
				else
					clr = RGB(112, 112, 112);
			}
			break;
		}
	}

#else
	clr = GetDefaultColorNoTdTheme(iColorIndex);
#endif;

	return clr;
}


// GetDefaultColorNoTdTheme
COLORREF MTBLSUBCLASS::GetDefaultColorNoTdTheme(int iColorIndex)
{
	COLORREF clr = RGB(0, 0, 0);

	HTHEME hTheme = NULL;
	CTheme &th = theTheme();
	if (th.CanUse() && th.IsAppThemed() && !QueryFlags(MTBL_FLAG_IGNORE_WIN_THEME))
	{
		hTheme = th.OpenThemeData(m_hWndTbl, L"Button");
	}

	switch (iColorIndex)
	{
	case GDC_HDR_BKGND:
		if (hTheme)
			clr = th.GetThemeSysColor(hTheme, COLOR_3DFACE);
		else
			clr = (COLORREF)GetSysColor(COLOR_3DFACE);
		break;

	case GDC_HDR_LINE:
		if (hTheme)
			clr = th.GetThemeSysColor(hTheme, COLOR_3DSHADOW);
		else
			clr = (COLORREF)GetSysColor(COLOR_3DSHADOW);
		break;

	case GDC_HDR_3DHL:
		if (hTheme)
			clr = th.GetThemeSysColor(hTheme, COLOR_3DHILIGHT);
		else
			clr = (COLORREF)GetSysColor(COLOR_3DHILIGHT);
		break;

	case GDC_CELL_LINE:
		if (hTheme)
			clr = th.GetThemeSysColor(hTheme, COLOR_3DSHADOW);
		else
			clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindowText);
		break;
	}

	if (hTheme)
		th.CloseThemeData(hTheme);

	return clr;
}


// GetDescendantMergeRows
long MTBLSUBCLASS::GetDescendantMergeRows(long lRow, vector<long> &v, LPMTBLMERGECELLS pMergeCells /*= NULL*/)
{
	long lCurRow, lDescRows = 0, lLastRow, lMergeCells;
	vector<LPMTBLMERGECELL> vmc;

	for (lCurRow = lLastRow = lRow; lCurRow <= lLastRow; lCurRow++)
	{
		vmc.clear();
		lMergeCells = FindMergeCells(NULL, lCurRow, FMC_MASTER, vmc, 0, pMergeCells);
		for (long l = 0; l < lMergeCells; l++)
		{
			if (vmc[l]->lRowTo > lLastRow)
				lLastRow = vmc[l]->lRowTo;
		}
	}

	if (lLastRow > lRow)
	{
		if (!pMergeCells)
			pMergeCells = m_pMergeCells;

		BOOL bAdd;
		for (lCurRow = lRow + 1; lCurRow <= lLastRow; lCurRow++)
		{
			if (pMergeCells->wRowFlagsOn && !Ctd.TblQueryRowFlags(m_hWndTbl, lCurRow, pMergeCells->wRowFlagsOn))
				bAdd = FALSE;
			else if (pMergeCells->wRowFlagsOff && Ctd.TblQueryRowFlags(m_hWndTbl, lCurRow, pMergeCells->wRowFlagsOff))
				bAdd = FALSE;
			else
				bAdd = TRUE;

			if (bAdd)
			{
				v.push_back(lCurRow);
				lDescRows++;
			}
		}
	}

	return lDescRows;
}


// GetDrawColor
// Liefert die zum Zeichnen zu verwendende Farbe
COLORREF MTBLSUBCLASS::GetDrawColor(COLORREF clr)
{
	if (QueryFlags(MTBL_FLAG_NO_DITHER))
		return PALETTERGB(GetRValue(clr), GetGValue(clr), GetBValue(clr));
	else
		return clr;
}


// GetDropDownListContext
BOOL MTBLSUBCLASS::GetDropDownListContext(HWND hWndCol, LPLONG plRow)
{
	if (!plRow)
		return FALSE;

	CMTblCol *pCol = m_Cols->Find(hWndCol);
	if (!pCol)
		return FALSE;

	CMTblCellType *pct = NULL;
	if (pCol->m_DropDownListContext != TBL_Error)
	{
		if (!IsRow(pCol->m_DropDownListContext))
			return FALSE;

		if (m_Rows->IsRowDelAdj(pCol->m_DropDownListContext))
			return FALSE;

		CMTblRow *pRow = m_Rows->GetRowPtr(pCol->m_DropDownListContext);
		if (!pRow)
			return FALSE;

		CMTblCol *pCell = pRow->m_Cols->Find(hWndCol);
		if (!pCell)
			return FALSE;

		pct = pCell->m_pCellType;
	}
	else
		pct = pCol->m_pCellType;

	if (pct && pct->IsDropDownList())
	{
		*plRow = pCol->m_DropDownListContext;
		return TRUE;
	}
	else
		return FALSE;
}


// GetEffCellBackColor
// Liefert die effektive Hintergrundfarbe einer Zelle
COLORREF MTBLSUBCLASS::GetEffCellBackColor(long lRow, HWND hWndCol, DWORD dwFlags /*= 0*/)
{
	// M!Table
	CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
	CMTblCol *pCol = m_Cols->Find(hWndCol);
	CMTblCol *pCell = (pRow ? pRow->m_Cols->Find(hWndCol) : NULL);

	COLORREF clr = GetEffCellBackColorMTbl(pCell, pCol, pRow, dwFlags);
	if (clr == MTBL_COLOR_UNDEF)
		clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);

	return clr;
}


// GetEffCellBackColorMTbl
// Liefert die effektive M!Table-Hintergrundfarbe einer Zelle
COLORREF MTBLSUBCLASS::GetEffCellBackColorMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow, DWORD dwFlags /*=0*/)
{
	if (dwFlags & GEC_EMPTY_COL && !QueryFlags(MTBL_FLAG_COLOR_ENTIRE_ROW))
		return MTBL_COLOR_UNDEF;
	if (dwFlags & GEC_EMPTY_ROW && !QueryFlags(MTBL_FLAG_COLOR_ENTIRE_COLUMN))
		return MTBL_COLOR_UNDEF;

	if (dwFlags & GEC_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;
		item.DefineAsCell(pCol ? pCol->m_hWnd : NULL, pRow);
		if (GetEffItemHighlighting(item, hili) && hili.dwBackColor != MTBL_COLOR_UNDEF)
			return hili.dwBackColor;
	}
	else if (dwFlags & GEC_HILI_PD && m_ppd->pc.Hili.dwBackColor != MTBL_COLOR_UNDEF)
		return m_ppd->pc.Hili.dwBackColor;

	DWORD dwColors[3] = { MTBL_COLOR_UNDEF, MTBL_COLOR_UNDEF, MTBL_COLOR_UNDEF };

	// Zelle
	if (pCell)
	{
		if (!pCell->m_BackColor.IsUndef())
			dwColors[m_byEffCellPropCellPos] = pCell->m_BackColor.Get();
	}

	// Spalte
	if (pCol)
	{
		if (!pCol->m_BackColor.IsUndef())
			dwColors[m_byEffCellPropColPos] = pCol->m_BackColor.Get();
	}

	// Zeile / Zeile abwechselnd
	if (pRow)
	{
		if (!pRow->m_BackColor->IsUndef())
			dwColors[m_byEffCellPropRowPos] = pRow->m_BackColor->Get();
		else
			dwColors[m_byEffCellPropRowPos] = GetRowAlternateBackColor(pRow->GetNr());
	}

	DWORD dwRet = MTBL_COLOR_UNDEF;
	for (int i = 0; i < 3; ++i)
	{
		if (dwColors[i] != MTBL_COLOR_UNDEF)
		{
			dwRet = dwColors[i];
			break;
		}

	}

	return dwRet;
}


// GetEffCellBtn
// Liefert den effektiven Button einer Zelle
CMTblBtn * MTBLSUBCLASS::GetEffCellBtn(HWND hWndCol, long lRow)
{
	BOOL bCellDef = FALSE;
	CMTblCol *pc = NULL;

	if (lRow != TBL_Error)
	{
		CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
		if (pRow)
		{
			pc = pRow->m_Cols->Find(hWndCol);
			if (pc && pc->m_Btn.m_bShow)
				bCellDef = TRUE;
		}
	}

	if (!bCellDef)
		pc = m_Cols->Find(hWndCol);

	return (pc != NULL ? &pc->m_Btn : NULL);
}


// GetEffCellEditJustify
int MTBLSUBCLASS::GetEffCellEditJustify(long lRow, HWND hWndCol)
{
	if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_EditLeftJustify))
		return DT_LEFT;
	else
	{
		int iJfy = GetEffCellEditJustifyMTbl(m_Rows->GetColPtr(lRow, hWndCol), m_Cols->Find(hWndCol), m_Rows->GetRowPtr(lRow));
		if (iJfy < 0)
			iJfy = GetColJustify(hWndCol);
		return iJfy;
	}
}


// GetEffCellEditJustifyMTbl
int MTBLSUBCLASS::GetEffCellEditJustifyMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow)
{
	int iJustifies[3] = { -1, -1, -1 };

	// Zelle
	if (pCell)
	{
		if (pCell->QueryFlags(MTBL_CELL_FLAG_EDITALIGN_LEFT))
			iJustifies[m_byEffCellPropCellPos] = DT_LEFT;
		else if (pCell->QueryFlags(MTBL_CELL_FLAG_EDITALIGN_CENTER))
			iJustifies[m_byEffCellPropCellPos] = DT_CENTER;
		else if (pCell->QueryFlags(MTBL_CELL_FLAG_EDITALIGN_RIGHT))
			iJustifies[m_byEffCellPropCellPos] = DT_RIGHT;
	}

	// Spalte
	if (pCol)
	{
		if (pCol->QueryFlags(MTBL_COL_FLAG_EDITALIGN_LEFT))
			iJustifies[m_byEffCellPropColPos] = DT_LEFT;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_EDITALIGN_CENTER))
			iJustifies[m_byEffCellPropColPos] = DT_CENTER;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_EDITALIGN_RIGHT))
			iJustifies[m_byEffCellPropColPos] = DT_RIGHT;
	}

	// Zeile
	if (pRow)
	{
		if (pRow->QueryFlags(MTBL_ROW_EDITALIGN_LEFT))
			iJustifies[m_byEffCellPropRowPos] = DT_LEFT;
		else if (pRow->QueryFlags(MTBL_ROW_EDITALIGN_CENTER))
			iJustifies[m_byEffCellPropRowPos] = DT_CENTER;
		else if (pRow->QueryFlags(MTBL_ROW_EDITALIGN_RIGHT))
			iJustifies[m_byEffCellPropRowPos] = DT_RIGHT;
	}

	int iRet = -1;
	for (int i = 0; i < 3; ++i)
	{
		if (iJustifies[i] != -1)
		{
			iRet = iJustifies[i];
			break;
		}

	}
	return iRet;
}


// GetEffCellFont
// Liefert den effektiven Font einer Zelle
HFONT MTBLSUBCLASS::GetEffCellFont(HDC hDC, long lRow, HWND hWndCol, LPBOOL pbMustDelete, DWORD dwFlags /*=0*/)
{
	return GetEffCellFont(hDC, m_Rows->GetColPtr(lRow, hWndCol), m_Cols->Find(hWndCol), m_Rows->GetRowPtr(lRow), pbMustDelete, dwFlags);
}

HFONT MTBLSUBCLASS::GetEffCellFont(HDC hDC, CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow, LPBOOL pbMustDelete, DWORD dwFlags /*=0*/)
{
	*pbMustDelete = FALSE;
	if (!hDC) return NULL;

	BOOL bCombEnh = (m_dwEffCellPropDetFlags & MTECPD_FLAG_COMB_FONT_ENH);
	CMTblFont f, *pFonts[3] = { NULL, NULL, NULL };

	// Zelle
	if (pCell)
		pFonts[m_byEffCellPropCellPos] = &pCell->m_Font;

	// Spalte
	if (pCol)
		pFonts[m_byEffCellPropColPos] = &pCol->m_Font;

	// Zeile
	if (pRow)
		pFonts[m_byEffCellPropRowPos] = pRow->m_Font;

	for (int i = 0; i < 3; ++i)
	{
		if (pFonts[i])
		{
			f.SubstUndef(*pFonts[i]);
			if (bCombEnh && !pFonts[i]->IsEnhUndef())
				f.AddEnh(pFonts[i]->GetEnh());
		}

		if (!bCombEnh && !f.AnyUndef())
			break;
	}

	// Tabelle
	if (f.IsUndef())
	{
		if (dwFlags & GEF_HYPERLINK || dwFlags & GEF_HILI || dwFlags & GEF_HILI_PD)
			f.Set(hDC, m_hFont);
		else
			return m_hFont;
	}
	else if (f.AnyUndef() || bCombEnh)
	{
		CMTblFont fSubst;
		fSubst.Set(hDC, m_hFont);
		f.SubstUndef(fSubst);
		if (bCombEnh && !fSubst.IsEnhUndef())
			f.AddEnh(fSubst.GetEnh());
	}

	// Modifizierungen
	if (dwFlags & GEF_HYPERLINK)
		f.AddEnh(m_lHLHoverFontEnh);
	else if (dwFlags & GEF_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsCell(pCol->m_hWnd, pRow);
		if (GetEffItemHighlighting(item, hili) && hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
			f.AddEnh(hili.lFontEnh);
	}
	else if (dwFlags & GEF_HILI_PD)
	{
		if (m_ppd->pc.Hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
			f.AddEnh(m_ppd->pc.Hili.lFontEnh);
	}

	// Font erzeugen
	HFONT hFont = f.Create(hDC);
	*pbMustDelete = TRUE;

	return hFont;
}


// GetEffCellImage
// Liefert das effektive Bild einer Zelle
CMTblImg MTBLSUBCLASS::GetEffCellImage(long lRow, HWND hWndCol, DWORD dwFlags /*= 0*/)
{
	CMTblImg Img;

	CMTblRow * pRow = m_Rows->GetRowPtr(lRow);
	if (!pRow)
		return Img;

	BOOL bExpanded = pRow->QueryFlags(MTBL_ROW_ISEXPANDED);

	CMTblCol * pCell = pRow->m_Cols->Find(hWndCol);
	if (pCell)
		Img = bExpanded && pCell->m_Image2.GetHandle() ? pCell->m_Image2 : pCell->m_Image;

	if (dwFlags & GEI_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsCell(hWndCol, pRow);
		if (GetEffItemHighlighting(item, hili))
		{
			CMTblImg &ImgHili = bExpanded ? hili.ImgExp : hili.Img;

			if (HANDLE hImg = ImgHili.GetHandle())
			{
				Img.SetHandle(hImg, NULL);
			}
		}
	}
	else if (dwFlags & GEI_HILI_PD)
	{
		CMTblImg &ImgHili = bExpanded ? m_ppd->pc.Hili.ImgExp : m_ppd->pc.Hili.Img;

		if (HANDLE hImg = ImgHili.GetHandle())
		{
			Img.SetHandle(hImg, NULL);
		}
	}

	return Img;
}


// GetEffCellIndent
int MTBLSUBCLASS::GetEffCellIndent(long lRow, HWND hWndCol)
{
	int iIndent = 0, iIndentLevel = 0;

	if ((m_hWndTreeCol != NULL) && (hWndCol == m_hWndTreeCol) && (lRow != TBL_Error) &&
		(m_Rows->QueryRowFlags(lRow, MTBL_ROW_CANEXPAND) || m_Rows->GetParentRow(lRow) != TBL_Error || (m_Rows->GetRowLevel(lRow) == 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_ALL))) &&
		!(m_Rows->GetRowLevel(lRow) == 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE))
		)
	{
		if (QueryTreeFlags(MTBL_TREE_FLAG_FLAT_STRUCT))
			iIndentLevel = 1;
		else
			iIndentLevel = m_Rows->GetRowLevel(lRow) + (QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE) ? 0 : 1);

		iIndent = iIndentLevel * m_iIndent;
	}
	else if (lRow != TBL_Error && hWndCol)
		iIndent = GetCellIndentLevel(lRow, hWndCol) * m_lIndentPerLevel;

	return iIndent;
}


// GetEffCellText
// Liefert effektiven ( = sichtbaren ) Text einer Zelle
BOOL MTBLSUBCLASS::GetEffCellText(long lRow, HWND hWndCol, tstring &sText)
{
	// Spalten-ID ermitteln
	int iColID = Ctd.TblQueryColumnID(hWndCol);
	if (iColID < 0) return FALSE;

	// Invisible?
	//BOOL bInvisible = Ctd.TblQueryColumnFlags( hWndCol, COL_Invisible );
	BOOL bInvisible = SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Invisible);

	// Kontext merken
	long lContext = Ctd.TblQueryContext(m_hWndTbl);

	// Kontext setzen
	if (!Ctd.TblSetContextEx(m_hWndTbl, lRow)) return FALSE;

	// Text ermitteln
	HSTRING hsText = 0;
	Ctd.TblGetColumnText(m_hWndTbl, iColID, &hsText);
	long lLen;
	LPTSTR lpsText = Ctd.StringGetBuffer(hsText, &lLen);
	lLen = Ctd.StrLenFromBufLen(lLen);

	// Kontext wiederherstellen
	Ctd.TblSetContextEx(m_hWndTbl, lContext);

	// Evtl. modifizieren
	if (lpsText)
	{
		// Invisible?
		if (bInvisible && !IsCheckBoxCell(hWndCol, lRow))
		{
			for (int i = 0; i < lLen; ++i)
			{
				lpsText[i] = m_cPassword;
			}
		}

		sText = lpsText;
	}
	else
		sText = _T("");
	Ctd.HStrUnlock(hsText);

	// HString dereferenzieren
	if (hsText)
		Ctd.HStringUnRef(hsText);

	return TRUE;
}


// GetEffCellTextColor
// Liefert die effektive Textfarbe einer Zelle
COLORREF MTBLSUBCLASS::GetEffCellTextColor(long lRow, HWND hWndCol, DWORD dwFlags /*= 0*/)
{
	// M!Table
	CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
	CMTblCol *pCol = m_Cols->Find(hWndCol);
	CMTblCol *pCell = (pRow ? pRow->m_Cols->Find(hWndCol) : NULL);

	COLORREF clr = GetEffCellTextColorMTbl(pCell, pCol, pRow, dwFlags);
	if (clr == MTBL_COLOR_UNDEF)
	{
		if (pCell)
			clr = pCell->m_TextColorSal.Get();
		if (clr == MTBL_COLOR_UNDEF)
			clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindowText);
	}

	return clr;
}


// GetEffCellTextColorMTbl
// Liefert die effektive M!Table-Textfarbe einer Zelle
COLORREF MTBLSUBCLASS::GetEffCellTextColorMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow, DWORD dwFlags /*= 0*/)
{
	if (dwFlags & GEC_HYPERLINK && m_dwHLHoverTextColor != MTBL_COLOR_UNDEF)
		return m_dwHLHoverTextColor;
	else if (dwFlags & GEC_HILI && pCol && pRow)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsCell(pCol->m_hWnd, pRow);
		if (GetEffItemHighlighting(item, hili) && hili.dwTextColor != MTBL_COLOR_UNDEF)
			return hili.dwTextColor;
	}
	else if (dwFlags & GEC_HILI_PD)
	{
		if (m_ppd->pc.Hili.dwTextColor != MTBL_COLOR_UNDEF)
			return m_ppd->pc.Hili.dwTextColor;
	}

	DWORD dwColors[3] = { MTBL_COLOR_UNDEF, MTBL_COLOR_UNDEF, MTBL_COLOR_UNDEF };

	// Zelle
	if (pCell)
	{
		if (!pCell->m_TextColor.IsUndef())
			dwColors[m_byEffCellPropCellPos] = pCell->m_TextColor.Get();
	}

	// Spalte
	if (pCol)
	{
		if (!pCol->m_TextColor.IsUndef())
			dwColors[m_byEffCellPropColPos] = pCol->m_TextColor.Get();
	}

	// Zeile
	if (pRow)
	{
		if (!pRow->m_TextColor->IsUndef())
			dwColors[m_byEffCellPropRowPos] = pRow->m_TextColor->Get();
	}

	DWORD dwRet = MTBL_COLOR_UNDEF;
	for (int i = 0; i < 3; ++i)
	{
		if (dwColors[i] != MTBL_COLOR_UNDEF)
		{
			dwRet = dwColors[i];
			break;
		}

	}

	return dwRet;
}


// GetEffCellTextJustify
int MTBLSUBCLASS::GetEffCellTextJustify(long lRow, HWND hWndCol)
{
	int iJfy = GetEffCellTextJustifyMTbl(m_Rows->GetColPtr(lRow, hWndCol), m_Cols->Find(hWndCol), m_Rows->GetRowPtr(lRow));
	if (iJfy < 0)
		iJfy = GetColJustify(hWndCol);
	return iJfy;
}


// GetEffCellTextJustifyMTbl
int MTBLSUBCLASS::GetEffCellTextJustifyMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow)
{
	int iJustifies[3] = { -1, -1, -1 };

	// Zelle
	if (pCell)
	{
		if (pCell->QueryFlags(MTBL_CELL_FLAG_TXTALIGN_LEFT))
			iJustifies[m_byEffCellPropCellPos] = DT_LEFT;
		else if (pCell->QueryFlags(MTBL_CELL_FLAG_TXTALIGN_CENTER))
			iJustifies[m_byEffCellPropCellPos] = DT_CENTER;
		else if (pCell->QueryFlags(MTBL_CELL_FLAG_TXTALIGN_RIGHT))
			iJustifies[m_byEffCellPropCellPos] = DT_RIGHT;
	}

	// Spalte
	if (pCol)
	{
		if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_LEFT))
			iJustifies[m_byEffCellPropColPos] = DT_LEFT;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_CENTER))
			iJustifies[m_byEffCellPropColPos] = DT_CENTER;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_RIGHT))
			iJustifies[m_byEffCellPropColPos] = DT_RIGHT;
	}

	// Zeile
	if (pRow)
	{
		if (pRow->QueryFlags(MTBL_ROW_TXTALIGN_LEFT))
			iJustifies[m_byEffCellPropRowPos] = DT_LEFT;
		else if (pRow->QueryFlags(MTBL_ROW_TXTALIGN_CENTER))
			iJustifies[m_byEffCellPropRowPos] = DT_CENTER;
		else if (pRow->QueryFlags(MTBL_ROW_TXTALIGN_RIGHT))
			iJustifies[m_byEffCellPropRowPos] = DT_RIGHT;
	}

	int iRet = -1;
	for (int i = 0; i < 3; ++i)
	{
		if (iJustifies[i] != -1)
		{
			iRet = iJustifies[i];
			break;
		}

	}
	return iRet;

}


// GetEffCellTextVAlign
int MTBLSUBCLASS::GetEffCellTextVAlign(long lRow, HWND hWndCol)
{
	int iVA = GetEffCellTextVAlignMTbl(m_Rows->GetColPtr(lRow, hWndCol), m_Cols->Find(hWndCol), m_Rows->GetRowPtr(lRow));
	if (iVA < 0)
		iVA = DT_TOP;
	return iVA;
}


// GetEffCellTextVAlignMTbl
int MTBLSUBCLASS::GetEffCellTextVAlignMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow)
{
	int iVAligns[3] = { -1, -1, -1 };

	// Zelle
	if (pCell)
	{
		if (pCell->QueryFlags(MTBL_CELL_FLAG_TXTALIGN_TOP))
			iVAligns[m_byEffCellPropCellPos] = DT_TOP;
		else if (pCell->QueryFlags(MTBL_CELL_FLAG_TXTALIGN_VCENTER))
			iVAligns[m_byEffCellPropCellPos] = DT_VCENTER;
		else if (pCell->QueryFlags(MTBL_CELL_FLAG_TXTALIGN_BOTTOM))
			iVAligns[m_byEffCellPropCellPos] = DT_BOTTOM;
	}

	// Spalte
	if (pCol)
	{
		if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_TOP))
			iVAligns[m_byEffCellPropColPos] = DT_TOP;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_VCENTER))
			iVAligns[m_byEffCellPropColPos] = DT_VCENTER;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_BOTTOM))
			iVAligns[m_byEffCellPropColPos] = DT_BOTTOM;
	}

	// Zeile
	if (pRow)
	{
		if (pRow->QueryFlags(MTBL_ROW_TXTALIGN_TOP))
			iVAligns[m_byEffCellPropRowPos] = DT_TOP;
		else if (pRow->QueryFlags(MTBL_ROW_TXTALIGN_VCENTER))
			iVAligns[m_byEffCellPropRowPos] = DT_VCENTER;
		else if (pRow->QueryFlags(MTBL_ROW_TXTALIGN_BOTTOM))
			iVAligns[m_byEffCellPropRowPos] = DT_BOTTOM;
	}

	int iRet = -1;
	for (int i = 0; i < 3; ++i)
	{
		if (iVAligns[i] != -1)
		{
			iRet = iVAligns[i];
			break;
		}

	}
	return iRet;
}


// GetEffCellType
// Liefert den effektiven Typ einer Zelle
int MTBLSUBCLASS::GetEffCellType(long lRow, HWND hWndCol)
{
	CMTblCellType *pct;
	pct = GetEffCellTypePtr(lRow, hWndCol);
	if (pct)
		return pct->m_iType;
	else
		return Ctd.TblGetCellType(hWndCol);
}

// GetEffCellTypePtr
CMTblCellType * MTBLSUBCLASS::GetEffCellTypePtr(long lRow, HWND hWndCol)
{
	CMTblCellType *pct = NULL;

	if (lRow != TBL_Error && hWndCol)
	{
		CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
		if (pRow)
		{
			CMTblCol *pCell = pRow->m_Cols->Find(hWndCol);
			if (pCell && pCell->m_pCellType)
			{
				pct = pCell->m_pCellType;
			}
		}
	}

	if (!pct)
	{
		CMTblCol *pCol = m_Cols->Find(hWndCol);
		if (pCol && pCol->m_pCellType)
		{
			pct = pCol->m_pCellType;
		}
	}

	return pct;
}


// GetEffColHdrBackColor
// Liefert die effektive Hintergrundfarbe eines Spaltenheaders
COLORREF MTBLSUBCLASS::GetEffColHdrBackColor(HWND hWndCol, DWORD dwFlags /*= 0*/)
{
	COLORREF clr = MTBL_COLOR_UNDEF;

	// Konkrete Spalte?
	if (hWndCol)
	{
		// Header?
		CMTblCol *pCol = m_Cols->FindHdr(hWndCol);
		if (pCol)
		{
			clr = pCol->m_BackColor.Get();

			if (dwFlags & GEC_HILI)
			{
				CMTblItem item;
				MTBLHILI hili;

				item.DefineAsColHdr(hWndCol);
				if (GetEffItemHighlighting(item, hili) && hili.dwBackColor != MTBL_COLOR_UNDEF)
					clr = hili.dwBackColor;
			}
			else if (dwFlags & GEC_HILI_PD && m_ppd->pc.Hili.dwBackColor != MTBL_COLOR_UNDEF)
				clr = m_ppd->pc.Hili.dwBackColor;
		}

		if (clr == MTBL_COLOR_UNDEF)
		{
			// Headers?
			clr = m_dwHdrsBackColor;

			if (clr == MTBL_COLOR_UNDEF)
			{
				// Standard
				if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders))
					//clr = COLORREF( GetSysColor( COLOR_3DFACE ) );
					clr = GetDefaultColor(GDC_HDR_BKGND);
				else
					clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);
			}
		}
	}
	// Oder freier Bereich?
	else
		clr = GetEffColHdrFreeBackColor();

	return clr;
}


// GetEffColHdrFont
// Liefert den effektiven Font eines Spaltenheaders
HFONT MTBLSUBCLASS::GetEffColHdrFont(HDC hDC, HWND hWndCol, LPBOOL pbMustDelete, DWORD dwFlags /*= 0*/)
{
	*pbMustDelete = FALSE;
	if (!hDC) return NULL;

	BOOL bAnyHili = (dwFlags & GEF_HILI) || (dwFlags & GEF_HILI_PD);

	CMTblFont f;
	for (;;)
	{
		// Header
		CMTblCol *pCol = m_Cols->FindHdr(hWndCol);
		if (pCol)
		{
			f = pCol->m_Font;
		}

		// Tabelle
		if (f.IsUndef() && !bAnyHili)
		{
			return m_hFont;
		}
		else if (f.AnyUndef())
		{
			CMTblFont fSubst;
			fSubst.Set(hDC, m_hFont);
			f.SubstUndef(fSubst);
		}

		break;
	}

	// Highlighting
	if (dwFlags & GEF_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsColHdr(hWndCol);
		if (GetEffItemHighlighting(item, hili) && hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
			f.AddEnh(hili.lFontEnh);
	}
	else if (dwFlags & GEF_HILI_PD)
	{
		if (m_ppd->pc.Hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
			f.AddEnh(m_ppd->pc.Hili.lFontEnh);
	}


	// Font erzeugen
	HFONT hFont = f.Create(hDC);
	*pbMustDelete = TRUE;

	return hFont;
}


// GetEffColHdrFreeBackColor
// Liefert die effektive Hintergrundfarbe des freien Spaltenheader-Bereichs
COLORREF MTBLSUBCLASS::GetEffColHdrFreeBackColor()
{
	COLORREF clr = m_dwColHdrFreeBackColor;

	if (clr == MTBL_COLOR_UNDEF)
	{
		// Headers?
		clr = m_dwHdrsBackColor;

		if (clr == MTBL_COLOR_UNDEF)
		{
			// Standard
			if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders))
				//clr = COLORREF( GetSysColor( COLOR_3DFACE ) );
				clr = GetDefaultColor(GDC_HDR_BKGND);
			else
				clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);
		}
	}

	return clr;
}


// GetEffColHdrGrpBackColor
// Liefert die effektive Hintergrundfarbe einer Spalten-Header-Gruppe
COLORREF MTBLSUBCLASS::GetEffColHdrGrpBackColor(CMTblColHdrGrp *pGrp, DWORD dwFlags /*= 0*/)
{
	COLORREF clr = MTBL_COLOR_UNDEF;

	// Gruppe?
	if (pGrp)
	{
		clr = pGrp->m_BackColor.Get();

		if (dwFlags & GEC_HILI)
		{
			CMTblItem item;
			MTBLHILI hili;

			item.DefineAsColHdrGrp(m_hWndTbl, pGrp);
			if (GetEffItemHighlighting(item, hili) && hili.dwBackColor != MTBL_COLOR_UNDEF)
				clr = hili.dwBackColor;
		}
		else if (dwFlags & GEC_HILI_PD)
		{
			if (m_ppd->pc.Hili.dwBackColor != MTBL_COLOR_UNDEF)
				clr = m_ppd->pc.Hili.dwBackColor;
		}
	}

	if (clr == MTBL_COLOR_UNDEF)
	{
		// Headers
		clr = m_dwHdrsBackColor;

		if (clr == MTBL_COLOR_UNDEF)
		{
			// Standard
			if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders))
				//clr = COLORREF( GetSysColor( COLOR_3DFACE ) );
				clr = GetDefaultColor(GDC_HDR_BKGND);
			else
				clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);
		}
	}

	return clr;
}


// GetEffColHdrGrpFont
// Liefert den effektiven Font einer Spalten-Header-Gruppe
HFONT MTBLSUBCLASS::GetEffColHdrGrpFont(HDC hDC, CMTblColHdrGrp * pGrp, LPBOOL pbMustDelete, DWORD dwFlags /*= 0*/)
{
	*pbMustDelete = FALSE;
	if (!hDC) return NULL;

	BOOL bAnyMod = (dwFlags & GEF_HILI) || (dwFlags & GEF_HILI_PD);

	CMTblFont f;
	for (;;)
	{
		// Gruppe
		if (pGrp)
		{
			f = pGrp->m_Font;
		}

		// Tabelle
		if (f.IsUndef() && !bAnyMod)
		{
			return m_hFont;
		}
		else if (f.AnyUndef())
		{
			CMTblFont fSubst;
			fSubst.Set(hDC, m_hFont);
			f.SubstUndef(fSubst);
		}

		break;
	}

	// Modifikation
	if (dwFlags & GEF_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsColHdrGrp(m_hWndTbl, pGrp);
		if (GetEffItemHighlighting(item, hili) && hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
			f.AddEnh(hili.lFontEnh);
	}
	else if (dwFlags & GEF_HILI_PD)
	{
		if (m_ppd->pc.Hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
			f.AddEnh(m_ppd->pc.Hili.lFontEnh);
	}

	// Font erzeugen
	HFONT hFont = f.Create(hDC);
	*pbMustDelete = TRUE;

	return hFont;
}


// GetEffColHdrGrpImage
// Liefert das effektive Bild einer Spalten-Header-Gruppe
CMTblImg MTBLSUBCLASS::GetEffColHdrGrpImage(CMTblColHdrGrp * pGrp, DWORD dwFlags /*= 0*/)
{
	CMTblImg Img;

	if (!pGrp)
		return Img;

	Img = pGrp->m_Image;

	if (dwFlags & GEI_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsColHdrGrp(m_hWndTbl, pGrp);
		if (GetEffItemHighlighting(item, hili))
		{
			if (HANDLE hImg = hili.Img.GetHandle())
				Img.SetHandle(hImg, NULL);
		}
	}
	else if (dwFlags & GEI_HILI_PD)
	{
		if (HANDLE hImg = m_ppd->pc.Hili.Img.GetHandle())
			Img.SetHandle(hImg, NULL);
	}


	return Img;
}


// GetEffColHdrGrpTextColor
// Liefert die effektive Textfarbe einer Spalten-Header-Gruppe
COLORREF MTBLSUBCLASS::GetEffColHdrGrpTextColor(CMTblColHdrGrp *pGrp, DWORD dwFlags /*= 0*/)
{
	COLORREF clr = MTBL_COLOR_UNDEF;

	// Gruppe?
	if (pGrp)
	{
		clr = pGrp->m_TextColor.Get();

		if (dwFlags & GEC_HILI)
		{
			CMTblItem item;
			MTBLHILI hili;

			item.DefineAsColHdrGrp(m_hWndTbl, pGrp);
			if (GetEffItemHighlighting(item, hili) && hili.dwTextColor != MTBL_COLOR_UNDEF)
				clr = hili.dwTextColor;
		}
		else if (dwFlags & GEC_HILI_PD)
		{
			if (m_ppd->pc.Hili.dwTextColor != MTBL_COLOR_UNDEF)
				clr = m_ppd->pc.Hili.dwTextColor;
		}
	}

	// Standard
	if (clr == MTBL_COLOR_UNDEF)
		clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindowText);

	return clr;
}


// GetEffColHdrImage
// Liefert das effektive Bild eines Spaltenheaders
CMTblImg MTBLSUBCLASS::GetEffColHdrImage(HWND hWndCol, DWORD dwFlags /*= 0*/)
{
	CMTblImg Img;

	CMTblCol * pColHdr = m_Cols->FindHdr(hWndCol);
	if (pColHdr)
		Img = pColHdr->m_Image;

	if (dwFlags & GEI_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsColHdr(hWndCol);
		if (GetEffItemHighlighting(item, hili))
		{
			if (HANDLE hImg = hili.Img.GetHandle())
				Img.SetHandle(hImg, NULL);
		}
	}
	else if (dwFlags & GEI_HILI_PD)
	{
		if (HANDLE hImg = m_ppd->pc.Hili.Img.GetHandle())
			Img.SetHandle(hImg, NULL);
	}

	return Img;
}



// GetEffColHdrTextColor
// Liefert die effektive Textfarbe eines Spaltenheaders
COLORREF MTBLSUBCLASS::GetEffColHdrTextColor(HWND hWndCol, DWORD dwFlags /*= 0*/)
{
	COLORREF clr = MTBL_COLOR_UNDEF;

	// Header?
	CMTblCol *pCol = m_Cols->FindHdr(hWndCol);
	if (pCol)
	{
		clr = pCol->m_TextColor.Get();

		if (dwFlags & GEC_HILI)
		{
			CMTblItem item;
			MTBLHILI hili;

			item.DefineAsColHdr(hWndCol);
			if (GetEffItemHighlighting(item, hili) && hili.dwTextColor != MTBL_COLOR_UNDEF)
				clr = hili.dwTextColor;
		}
		else if (dwFlags & GEC_HILI_PD)
		{
			if (m_ppd->pc.Hili.dwTextColor != MTBL_COLOR_UNDEF)
				clr = m_ppd->pc.Hili.dwTextColor;
		}
	}

	// Standard
	if (clr == MTBL_COLOR_UNDEF)
		clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindowText);

	return clr;
}


// GetEffColTextJustify
// Liefert die effektive Textausrichtung einer Spalte
int MTBLSUBCLASS::GetEffColTextJustify(HWND hWndCol)
{
	CMTblCol *pCol = m_Cols->Find(hWndCol);
	if (pCol)
	{
		if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_LEFT))
			return DT_LEFT;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_CENTER))
			return DT_CENTER;
		else if (pCol->QueryFlags(MTBL_COL_FLAG_TXTALIGN_RIGHT))
			return DT_RIGHT;
	}

	return GetColJustify(hWndCol);
}


// GetEffCornerBackColor
// Liefert die effektive Hintergrundfarbe des Corners
COLORREF MTBLSUBCLASS::GetEffCornerBackColor(DWORD dwFlags /*= 0*/)
{
	COLORREF clr = m_Corner->m_BackColor.Get();

	if (dwFlags & GEC_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsCorner(m_hWndTbl);
		if (GetEffItemHighlighting(item, hili) && hili.dwBackColor != MTBL_COLOR_UNDEF)
			clr = hili.dwBackColor;
	}
	else if (dwFlags & GEC_HILI_PD && m_ppd->pc.Hili.dwBackColor != MTBL_COLOR_UNDEF)
		clr = m_ppd->pc.Hili.dwBackColor;

	if (clr == MTBL_COLOR_UNDEF)
	{
		// Headers?
		clr = m_dwHdrsBackColor;

		if (clr == MTBL_COLOR_UNDEF)
		{
			// Standard
			if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders))
				clr = GetDefaultColor(GDC_HDR_BKGND);
			else
				clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);
		}
	}

	return clr;
}


// GetEffCornerImage
// Liefert das effektive Bild des Corners
CMTblImg MTBLSUBCLASS::GetEffCornerImage(DWORD dwFlags /*= 0*/)
{
	CMTblImg Img = m_Corner->m_Image;

	if (dwFlags & GEI_HILI)
	{
		CMTblItem item;
		MTBLHILI hili;

		item.DefineAsCorner(m_hWndTbl);
		if (GetEffItemHighlighting(item, hili))
		{
			if (HANDLE hImg = hili.Img.GetHandle())
			{
				Img.SetHandle(hImg, NULL);
			}
		}
	}
	else if (dwFlags & GEI_HILI_PD)
	{
		if (HANDLE hImg = m_ppd->pc.Hili.Img.GetHandle())
			Img.SetHandle(hImg, NULL);
	}


	return Img;
}


// GetEffFocusCell
// Liefert die effektive Focus-Zelle im Cell-Mode
BOOL MTBLSUBCLASS::GetEffFocusCell(long &lRow, HWND &hWndCol)
{
	if (!m_bCellMode)
		return FALSE;

	lRow = GetRowNrFocus();
	hWndCol = m_hWndColFocus;

	if (lRow != TBL_Error && hWndCol)
	{
		GetMasterCell(hWndCol, lRow);
	}

	return TRUE;
}


// GetEffItemBtn
// Liefert den effektiven Button eines Items
CMTblBtn * MTBLSUBCLASS::GetEffItemBtn(CMTblItem &item)
{
	CMTblBtn* pBtn = NULL;
	if (item.IsCell())
	{
		BOOL bCellDef = FALSE;
		CMTblCol *pc = NULL;

		CMTblRow* pRow = item.GetRowPtr();
		if (pRow)
		{
			pc = pRow->m_Cols->Find(item.WndHandle);
			if (pc && pc->m_Btn.m_bShow)
				bCellDef = TRUE;
		}

		if (!bCellDef)
			pc = m_Cols->Find(item.WndHandle);

		pBtn = (pc ? &pc->m_Btn : NULL);
	}
	else if (item.IsColHdr())
	{
		pBtn = m_Cols->GetHdrBtn(item.WndHandle);
	}

	return pBtn;
}


// GetEffItemHighlighting
// Liefert die effektive Hervorhebung eines Items
BOOL MTBLSUBCLASS::GetEffItemHighlighting(CMTblItem &item, MTBLHILI &hili)
{
	hili.Init();

	CMTblItem item2;
	LPMTBLHILI pahili[3] = { NULL, NULL, NULL };
	int iaItem[3] = { 0, 0, 0 };

	switch (item.Type)
	{
	case MTBL_ITEM_CELL:
		pahili[m_byEffCellPropCellPos] = GetItemHighlighting(item);
		iaItem[m_byEffCellPropCellPos] = item.Type;

		item2.DefineAsRow(m_hWndTbl, item.Id);
		pahili[m_byEffCellPropRowPos] = GetItemHighlighting(item2);
		iaItem[m_byEffCellPropRowPos] = item2.Type;

		item2.DefineAsCol(item.WndHandle);
		pahili[m_byEffCellPropColPos] = GetItemHighlighting(item2);
		iaItem[m_byEffCellPropColPos] = item2.Type;
		break;
	case MTBL_ITEM_COLHDR:
		pahili[0] = GetItemHighlighting(item);
		iaItem[0] = item.Type;

		item2.DefineAsCol(item.WndHandle);
		pahili[1] = GetItemHighlighting(item2);
		iaItem[1] = item2.Type;
		break;
	case MTBL_ITEM_COLHDRGRP:
		if (item.WndHandle == m_hWndTbl)
		{
			pahili[0] = GetItemHighlighting(item);
			iaItem[0] = item.Type;
		}
		else
		{
			item2 = item;
			item2.WndHandle = m_hWndTbl;
			pahili[0] = GetItemHighlighting(item2);
			iaItem[0] = item2.Type;
		}

		if (item.WndHandle != m_hWndTbl)
		{
			item2.DefineAsCol(item.WndHandle);
			pahili[1] = GetItemHighlighting(item2);
			iaItem[1] = item2.Type;
		}
		break;
	case MTBL_ITEM_ROWHDR:
		pahili[0] = GetItemHighlighting(item);
		iaItem[0] = item.Type;

		item2.DefineAsRow(m_hWndTbl, item.Id);
		pahili[1] = GetItemHighlighting(item2);
		iaItem[1] = item2.Type;
		break;
	default:
		pahili[0] = GetItemHighlighting(item);
		iaItem[0] = item.Type;
		break;
	}

	for (int i = 2; i >= 0; i--)
	{
		if (pahili[i])
		{
			CombineHighlighting(item.Type, hili, iaItem[i], *pahili[i]);
		}
	}

	// Ermitteln, ob es sich um einen verbunden Teil handelt. Falls ja, Highlighting für das Master-Objekt ermitteln und mit dem bisher ermittelten Highlighting kombinieren.
	LPMTBLHILI phili = NULL;
	if (item.IsCell())
	{
		LPMTBLMERGECELL pmc = FindMergeCell(item.WndHandle, item.GetRowNr(), FMC_SLAVE);
		if (pmc)
		{
			item2 = item;
			item2.WndHandle = pmc->hWndColFrom;
			item2.Id = m_Rows->GetRowPtr(pmc->lRowFrom);
			phili = GetItemHighlighting(item2);
		}
	}
	else if (item.IsRowHdr())
	{
		LPMTBLMERGEROW pmr = FindMergeRow(item.GetRowNr(), FMR_SLAVE);
		if (pmr)
		{
			item2 = item;
			item2.Id = m_Rows->GetRowPtr(pmr->lRowFrom);
			phili = GetItemHighlighting(item2);
		}
	}

	if (phili)
		CombineHighlighting(item.Type, hili, item.Type, *phili);

	return TRUE;
}

// GetEffItemLine
// Liefert die effektive Liniendefinition eines Items
BOOL MTBLSUBCLASS::GetEffItemLine(CMTblItem &item, int iLineType, CMTblLineDef &ld)
{
	CMTblItem it;
	CMTblLineDef dld;
	CMTblLineDef* plda[4] = { NULL, NULL, NULL, NULL };
	HWND hWndCol, hWndLastLockedCol;
	int i;
	const int iDef = 3;
	long lRow;

	if (item.IsCell())
	{
		// Letzte verbundene Spalte/Zeile ermitteln
		if (!GetLastMergedCellEx(item.WndHandle, item.GetRowNr(), hWndCol, lRow))
		{
			hWndCol = item.WndHandle;
			lRow = item.GetRowNr();
		}

		// Letzte verankerte Spalte ermitteln
		hWndLastLockedCol = GetLastLockedColumn();

		// Lininedefinition Zelle
		plda[m_byEffCellPropCellPos] = GetItemLine(item, iLineType);

		// Liniendefinition Zeile
		it.DefineAsRow(m_hWndTbl, lRow);
		plda[m_byEffCellPropRowPos] = GetItemLine(it, iLineType);

		// Liniendefinition Spalte
		it.DefineAsCol(hWndCol);
		plda[m_byEffCellPropColPos] = GetItemLine(it, iLineType);

		// Lininedefinition Standard
		if (iLineType == MTLT_HORZ)
		{
			it.DefineAsRows(m_hWndTbl);
			if (GetEffItemLine(it, iLineType, dld))
				plda[iDef] = &dld;
		}
			
		else if (iLineType == MTLT_VERT)
		{
			if (hWndCol == hWndLastLockedCol)
			{
				it.DefineAsLastLockedCol(m_hWndTbl);
				if (GetEffItemLine(it, iLineType, dld))
					plda[iDef] = &dld;
			}
			else
			{
				it.DefineAsCols(m_hWndTbl);
				if (GetEffItemLine(it, iLineType, dld))
					plda[iDef] = &dld;
			}
		}			
	}
	else if (item.IsRow())
	{
		// Liniendefinition Zeile
		it.DefineAsRow(m_hWndTbl, item.GetRowPtr());
		plda[m_byEffCellPropRowPos] = GetItemLine(it, iLineType);

		// Liniendefinition Spalte
		it.DefineAsCol(item.WndHandle);
		plda[m_byEffCellPropColPos] = GetItemLine(it, iLineType);

		// Lininedefinition Standard
		if (iLineType == MTLT_HORZ)
		{
			it.DefineAsRows(m_hWndTbl);
			if (GetEffItemLine(it, iLineType, dld))
				plda[iDef] = &dld;
		}

		else if (iLineType == MTLT_VERT)
		{
			it.DefineAsCols(m_hWndTbl);
			if (GetEffItemLine(it, iLineType, dld))
				plda[iDef] = &dld;
		}
	}
	else if (item.IsRows())
	{
		if (iLineType == MTLT_HORZ)
		{
			dld = *m_pRowsHorzLineDef;
			if (dld.Color == MTBL_COLOR_UNDEF)
				dld.Color = GetDefaultColor(GDC_CELL_LINE);
			plda[iDef] = &dld;
		}
		else
			return FALSE;
	}
	else if (item.IsRowHdr())
	{
		// Lininedefinition Row-Header
		plda[0] = GetItemLine(item, iLineType);

		// Lininedefinition Standard
		dld = *m_pHdrLineDef;
		if (dld.Color == MTBL_COLOR_UNDEF)
			dld.Color = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders) ? GetDefaultColor(GDC_HDR_LINE) : 0;
		plda[iDef] = &dld;
	}
	else if (item.IsCol())
	{
		hWndCol = item.WndHandle;

		// Letzte verankerte Spalte ermitteln
		hWndLastLockedCol = GetLastLockedColumn();

		// Lininedefinition Spalte
		plda[0] = GetItemLine(item, iLineType);

		// Lininedefinition Standard
		if (iLineType == MTLT_HORZ)
		{
			it.DefineAsRows(m_hWndTbl);
			if (GetEffItemLine(it, iLineType, dld))
				plda[iDef] = &dld;
		}

		else if (iLineType == MTLT_VERT)
		{
			if (hWndCol == hWndLastLockedCol)
			{
				it.DefineAsLastLockedCol(m_hWndTbl);
				if (GetEffItemLine(it, iLineType, dld))
					plda[iDef] = &dld;
			}
			else
			{
				it.DefineAsCols(m_hWndTbl);
				if (GetEffItemLine(it, iLineType, dld))
					plda[iDef] = &dld;
			}
		}
	}
	else if (item.IsLastLockedCol())
	{
		if (iLineType == MTLT_VERT)
		{
			dld = *m_pLastLockedColVertLineDef;
			if (dld.Color == MTBL_COLOR_UNDEF)
				dld.Color = GetDefaultColor(GDC_CELL_LINE);
			plda[iDef] = &dld;
		}
		else
			return FALSE;
	}
	else if (item.IsCols())
	{
		if (iLineType == MTLT_VERT)
		{
			dld = *m_pColsVertLineDef;
			if (dld.Color == MTBL_COLOR_UNDEF)
				dld.Color = GetDefaultColor(GDC_CELL_LINE);
			plda[iDef] = &dld;
		}
		else
			return FALSE;
	}
	else if (item.IsColHdr() || item.IsColHdrGrp() || item.IsCorner())
	{
		// Lininedefinition Spalten-Header
		plda[0] = GetItemLine(item, iLineType);

		// Lininedefinition Standard
		CMTblLineDef dld(*m_pHdrLineDef);
		if (dld.Color == MTBL_COLOR_UNDEF)
			dld.Color = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders) ? GetDefaultColor(GDC_HDR_LINE) : 0;
		plda[iDef] = &dld;
	}

	// Effektive Liniendefinition ermitteln
	ld.Init();
	for (i = 0; i < 4; i++)
	{
		if (plda[i])
			ld.ReplaceUndefinedFrom(*plda[i]);
	}

	return TRUE;
}


// GetEffRowHdrBackColor
// Liefert die effektive Hintergrundfarbe eines Zeilenkopfes
COLORREF MTBLSUBCLASS::GetEffRowHdrBackColor(long lRow, DWORD dwFlags /*= 0*/)
{
	COLORREF clr = MTBL_COLOR_UNDEF;

	if (lRow == TBL_Error)
	{
		// Freier Bereich
		clr = m_dwRowHdrFreeBackColor;
	}
	else
	{
		// Farbe ermitteln
		CMTblRow * pRow = m_Rows->GetRowPtr(lRow);
		if (pRow)
		{
			clr = pRow->m_HdrBackColor->Get();

			if (dwFlags & GEC_HILI)
			{
				CMTblItem item;
				MTBLHILI hili;

				item.DefineAsRowHdr(m_hWndTbl, pRow);
				if (GetEffItemHighlighting(item, hili) && hili.dwBackColor != MTBL_COLOR_UNDEF)
					clr = hili.dwBackColor;
			}
			else if (dwFlags & GEC_HILI_PD && m_ppd->pc.Hili.dwBackColor != MTBL_COLOR_UNDEF)
				clr = m_ppd->pc.Hili.dwBackColor;
		}
	}

	// Keine Farbe gefunden = Standard
	if (clr == MTBL_COLOR_UNDEF)
	{
		//Headers?
		clr = m_dwHdrsBackColor;

		if (clr == MTBL_COLOR_UNDEF)
		{
			// Standard
			if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders))
				//clr = GetSysColor( COLOR_3DFACE );
				clr = GetDefaultColor(GDC_HDR_BKGND);
			else
				clr = Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow);
		}
	}

	return clr;
}


// GetEffRowHeight
// Liefert die effektive Höhe einer Zeile
long MTBLSUBCLASS::GetEffRowHeight(long lRow, CMTblMetrics *ptm /*= NULL*/)
{
	long lHeight = m_Rows->GetRowHeight(lRow);
	if (lHeight == 0) // Standardhöhe
	{
		CMTblMetrics tm;
		if (!ptm)
		{
			tm.Get(m_hWndTbl);
			ptm = &tm;
		}

		lHeight = ptm->m_RowHeight;
	}

	return lHeight;
}

// GetExportSubstColor
DWORD MTBLSUBCLASS::GetExportSubstColor(DWORD dwColor)
{
	SubstColorMap::iterator it = m_escm->find(dwColor);
	if (it != m_escm->end())
		return (*it).second;
	else
		return MTBL_COLOR_UNDEF;
}


// GetExtEditSubClass
LPMTBLEXTEDITSUBCLASS MTBLSUBCLASS::GetExtEditSubClass(HWND hWndExtEdit)
{
	SubClassMap::iterator it = m_scm->find(hWndExtEdit);
	if (it != m_scm->end())
	{
		return (LPMTBLEXTEDITSUBCLASS)(*it).second;
	}

	return NULL;
}


// GetFocusCellRect
BOOL MTBLSUBCLASS::GetFocusCellRect(LPRECT pRect)
{
	long lRowFocus = GetRowNrFocus();
	return GetCellFocusRect(lRowFocus, m_hWndColFocus, pRect);
}


// GetFocusRgn
HRGN MTBLSUBCLASS::GetFocusRgn(LPRECT pRect, int iType, BOOL bCellMode)
{
	if (!pRect)
		return NULL;

	switch (iType)
	{
	case GFR_OUTER:
		return CreateRectRgnIndirect(pRect);
		break;
	case GFR_INNER:
		return CreateRectRgn(pRect->left + (bCellMode ? m_lFocusFrameThick : 0), pRect->top + +m_lFocusFrameThick, pRect->right - (bCellMode ? m_lFocusFrameThick : 0), pRect->bottom - m_lFocusFrameThick);
		break;
	case GFR_FRAME:
		HRGN rgnFrame, rgnInner;
		int iRet;

		rgnFrame = CreateRectRgnIndirect(pRect);
		rgnInner = CreateRectRgn(pRect->left + (bCellMode ? m_lFocusFrameThick : 0), pRect->top + +m_lFocusFrameThick, pRect->right - (bCellMode ? m_lFocusFrameThick : 0), pRect->bottom - m_lFocusFrameThick);
		iRet = CombineRgn(rgnFrame, rgnFrame, rgnInner, RGN_DIFF);

		DeleteObject(rgnInner);

		if (iRet != ERROR)
		{
			return rgnFrame;
		}
		else
		{
			DeleteObject(rgnFrame);
			return NULL;
		}

		break;
	}

	return NULL;
}


// GetHTMLText
BOOL MTBLSUBCLASS::GetHTMLText(tstring &sText)
{
	if (sText.empty())
	{
		sText = _T("<p>&nbsp</p>");
		return TRUE;
	}
	else
	{
		tstring sModText = _T("<p>");
		tstring::iterator it, itEnd = sText.end();
		for (it = sText.begin(); it != itEnd; ++it)
		{
			switch (*it)
			{
			case _T('ä'):
				sModText.append(_T("&auml"));
				break;
			case _T('Ä'):
				sModText.append(_T("&Auml"));
				break;
			case _T('ö'):
				sModText.append(_T("&ouml"));
				break;
			case _T('Ö'):
				sModText.append(_T("&Ouml"));
				break;
			case _T('ü'):
				sModText.append(_T("&uuml"));
				break;
			case _T('Ü'):
				sModText.append(_T("&Uuml"));
				break;
			case _T('ß'):
				sModText.append(_T("&szlig"));
				break;
			case _T('\r'):
				sModText.append(_T("<br>"));
				break;
			case _T('\n'):
				break;
			default:
				sModText.append(1, *it);
			}
		}
		sModText.append(_T("</p>"));
		sText = sModText;
	}

	return TRUE;
}


// GenerateColSelChangeMsgs
void MTBLSUBCLASS::GenerateColSelChangeMsg()
{
	BOOL bSelAfter, bSelBefore;
	MTblColList::iterator it, itEnd = m_Cols->m_List.end();

	for (it = m_Cols->m_List.begin(); it != itEnd; ++it)
	{
		bSelBefore = (*it).second.QueryInternalFlags(COL_IFLAG_SELECTED_BEFORE) != 0;
		bSelAfter = (*it).second.QueryInternalFlags(COL_IFLAG_SELECTED_AFTER) != 0;
		if (bSelBefore != bSelAfter)
		{
			SendMessage(m_hWndTbl, MTM_ColSelChanged, 0, 0);
			return;
		}
	}
}


// GenerateExtMsgs
void MTBLSUBCLASS::GenerateExtMsgs(UINT uMsg)
{
	// CTD da?
	if (!Ctd.IsPresent()) return;

	DWORD dwMFlags;
	LPARAM lParam;
	//NUMBER n;
	POINT pt;
	WPARAM wParam;

	switch (uMsg)
	{
	case WM_LBUTTONDOWN:
		// Wo ist die Maus?
		pt.x = iLDownX;
		pt.y = iLDownY;
		ObjFromPoint(pt, &lLDownRow, &hWndLDownCol, &dwLDownFlags, &dwMFlags);
		hWndLDownSepCol = GetSepCol(iLDownX, iLDownY);
		lLDownSepRow = GetSepRow(pt);
		bLDownCornerSep = (dwMFlags & MTOFP_OVERCORNERSEP);
		// Beginn Änderung Zeilenhöhe/LinesPerRow?
		bLDownBeginRowSizing = FALSE;
		if (lLDownSepRow != TBL_Error && m_bRowSizing)
		{
			bLDownBeginRowSizing = TRUE;
			if (QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT))
				lLDownRowHeight = GetEffRowHeight(lLDownSepRow, NULL);
			else
				Ctd.TblQueryLinesPerRow(m_hWndTbl, &iLDownLinesPerRow);
		}
		// Position der Spalte merken
		iLDownColPos = Ctd.TblQueryColumnPos(hWndLDownCol);
		// Tabellen-Maße ermitteln
		SendMessage(m_hWndTbl, TIM_GetMetrics, 0, (LPARAM)&mLDownMetrics);
		// Nachrichten generieren
		if (bLDownCornerSep)
		{
			PostMessage(m_hWndTbl, MTM_CornerSepLBtnDown, 0, 0);
			if (m_dwExtMsgOpt & MTEM_CORNERSEP_EXCLUSIVE)
				goto endlbtndown;
		}

		if (dwLDownFlags & TBL_XOverRowHeader && dwLDownFlags & TBL_YOverColumnHeader)
		{
			PostMessage(m_hWndTbl, MTM_CornerLBtnDown, 0, 0);
		}
		else if (dwLDownFlags & TBL_YOverColumnHeader)
		{
			if (hWndLDownSepCol)
			{
				PostMessage(m_hWndTbl, MTM_ColHdrSepLBtnDown, (WPARAM)hWndLDownSepCol, 0);
				dLDownSepColWidth = SendMessage(hWndLDownSepCol, TIM_QueryColumnWidth, 0, 0);
				if (m_dwExtMsgOpt & MTEM_COLHDRSEP_EXCLUSIVE)
					goto endlbtndown;
			}

			wParam = WPARAM(hWndLDownCol);
			if (hWndLDownCol && (dwMFlags & MTOFP_OVERCOLHDRGRP))
				lParam = LPARAM(m_ColHdrGrps->GetGrpOfCol(hWndLDownCol));
			else
				lParam = 0;
			PostMessage(m_hWndTbl, MTM_ColHdrLBtnDown, wParam, lParam);
		}
		else if (dwLDownFlags & TBL_XOverRowHeader)
		{
			if (lLDownSepRow != TBL_Error)
			{
				PostMessage(m_hWndTbl, MTM_RowHdrSepLBtnDown, 0, lLDownSepRow);
				if (m_dwExtMsgOpt & MTEM_ROWHDRSEP_EXCLUSIVE)
					goto endlbtndown;
			}
			PostMessage(m_hWndTbl, MTM_RowHdrLBtnDown, 0, lLDownRow);
		}
		else if (dwLDownFlags & TBL_YOverSplitBar)
		{
			PostMessage(m_hWndTbl, MTM_SplitBarLBtnDown, 0, 0);
		}
		else
		{
			PostMessage(m_hWndTbl, MTM_AreaLBtnDown, (WPARAM)hWndLDownCol, lLDownRow);
		}

	endlbtndown:
		break;

	case WM_LBUTTONUP:
		// Wo ist die Maus?
		pt.x = iLUpX;
		pt.y = iLUpY;
		ObjFromPoint(pt, &lLUpRow, &hWndLUpCol, &dwLUpFlags, &dwMFlags);
		hWndLUpSepCol = GetSepCol(iLUpX, iLUpY);
		lLUpSepRow = GetSepRow(pt);
		bLUpCornerSep = (dwMFlags & MTOFP_OVERCORNERSEP);
		// Tabellen-Maße ermitteln
		SendMessage(m_hWndTbl, TIM_GetMetrics, 0, (LPARAM)&mLUpMetrics);
		// Nachrichten generieren
		if (bLUpCornerSep)
		{
			PostMessage(m_hWndTbl, MTM_CornerSepLBtnUp, 0, 0);
			if (m_dwExtMsgOpt & MTEM_CORNERSEP_EXCLUSIVE)
				goto endlbtnup;
		}

		if (dwLUpFlags & TBL_XOverRowHeader && dwLUpFlags & TBL_YOverColumnHeader)
		{
			PostMessage(m_hWndTbl, MTM_CornerLBtnUp, 0, 0);
		}
		else if (dwLUpFlags & TBL_YOverColumnHeader)
		{
			if (hWndLUpSepCol)
			{
				PostMessage(m_hWndTbl, MTM_ColHdrSepLBtnUp, (WPARAM)hWndLDownSepCol, 0);
				if (m_dwExtMsgOpt & MTEM_COLHDRSEP_EXCLUSIVE)
					goto endlbtnup;
			}

			wParam = WPARAM(hWndLUpCol);
			if (hWndLUpCol && (dwMFlags & MTOFP_OVERCOLHDRGRP))
				lParam = LPARAM(m_ColHdrGrps->GetGrpOfCol(hWndLUpCol));
			else
				lParam = 0;
			PostMessage(m_hWndTbl, MTM_ColHdrLBtnUp, wParam, lParam);
		}
		else if (dwLUpFlags & TBL_XOverRowHeader)
		{
			if (lLUpSepRow != TBL_Error)
			{
				PostMessage(m_hWndTbl, MTM_RowHdrSepLBtnUp, 0, lLUpSepRow);
				if (m_dwExtMsgOpt & MTEM_ROWHDRSEP_EXCLUSIVE)
					goto endlbtnup;
			}
			PostMessage(m_hWndTbl, MTM_RowHdrLBtnUp, 0, lLUpRow);
		}
		else if (dwLUpFlags & TBL_YOverSplitBar)
		{
			PostMessage(m_hWndTbl, MTM_SplitBarLBtnUp, 0, 0);
		}
		else
		{
			PostMessage(m_hWndTbl, MTM_AreaLBtnUp, (WPARAM)hWndLUpCol, lLUpRow);
		}

	endlbtnup:

		if (hWndLDownCol)
		{
			// Position geändert?
			iLUpColPos = Ctd.TblQueryColumnPos(hWndLDownCol);
			if (iLUpColPos != iLDownColPos && uLastLDownMsg == WM_LBUTTONDOWN)
			{
				PostMessage(m_hWndTbl, MTM_ColPosChanged, (WPARAM)hWndLDownCol, MAKELPARAM(iLDownColPos, iLUpColPos));
				PostMessage(m_hWndTbl, MTM_ColMoved, (WPARAM)hWndLDownCol, MAKELPARAM(iLDownColPos, iLUpColPos));
			}
		}

		if (hWndLDownSepCol)
		{
			// Breite geändert?
			dLUpSepColWidth = SendMessage(hWndLDownSepCol, TIM_QueryColumnWidth, 0, 0);
			if (dLUpSepColWidth != dLDownSepColWidth && uLastLDownMsg == WM_LBUTTONDOWN)
			{
				PostMessage(m_hWndTbl, MTM_ColSized, (WPARAM)hWndLDownSepCol, 0);
			}
		}

		if (dwLDownFlags & TBL_YOverSplitBar && dwLUpFlags && TBL_YOverSplitBar)
		{
			if (mLDownMetrics.iSplitBarTop != mLUpMetrics.iSplitBarTop && uLastLDownMsg == WM_LBUTTONDOWN)
			{
				PostMessage(m_hWndTbl, MTM_SplitBarMoved, 0, 0);
			}
		}

		if (bLDownBeginRowSizing)
		{
			if (QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT))
			{
				lLUpRowHeight = GetEffRowHeight(lLDownSepRow, NULL);
				if (lLUpRowHeight != lLDownRowHeight)
					PostMessage(m_hWndTbl, MTM_RowSized, lLDownRowHeight, lLDownSepRow);
			}
			else
			{
				// Anzahl LinesPerRow geändert?
				if (Ctd.TblQueryLinesPerRow(m_hWndTbl, &iLUpLinesPerRow) && iLUpLinesPerRow != iLDownLinesPerRow)
					PostMessage(m_hWndTbl, MTM_LinesPerRowChanged, iLDownLinesPerRow, iLUpLinesPerRow);
			}

			bLDownBeginRowSizing = FALSE;
		}

		break;

	case WM_LBUTTONDBLCLK:
		// Wo ist die Maus?
		pt.x = iLDblClkX;
		pt.y = iLDblClkY;
		ObjFromPoint(pt, &lLDblClkRow, &hWndLDblClkCol, &dwLDblClkFlags, &dwMFlags);
		hWndLDblClkSepCol = GetSepCol(iLDblClkX, iLDblClkY);
		lLDblClkSepRow = GetSepRow(pt);
		bLDblClkCornerSep = (dwMFlags & MTOFP_OVERCORNERSEP);
		// Nachrichten generieren
		if (bLDblClkCornerSep)
		{
			PostMessage(m_hWndTbl, MTM_CornerSepLBtnDblClk, 0, 0);
			if (m_dwExtMsgOpt & MTEM_CORNERSEP_EXCLUSIVE)
				goto endlbtndblclk;
		}

		if (dwLDblClkFlags & TBL_XOverRowHeader && dwLDblClkFlags & TBL_YOverColumnHeader)
		{
			PostMessage(m_hWndTbl, MTM_CornerLBtnDblClk, 0, 0);
		}
		else if (dwLDblClkFlags & TBL_YOverColumnHeader)
		{
			if (hWndLDblClkSepCol)
			{
				PostMessage(m_hWndTbl, MTM_ColHdrSepLBtnDblClk, (WPARAM)hWndLDblClkSepCol, 0);
				if (m_dwExtMsgOpt & MTEM_COLHDRSEP_EXCLUSIVE)
					goto endlbtndblclk;
			}

			wParam = WPARAM(hWndLDblClkCol);
			if (hWndLDblClkCol && (dwMFlags & MTOFP_OVERCOLHDRGRP))
				lParam = LPARAM(m_ColHdrGrps->GetGrpOfCol(hWndLDblClkCol));
			else
				lParam = 0;
			PostMessage(m_hWndTbl, MTM_ColHdrLBtnDblClk, wParam, lParam);
		}
		else if (dwLDblClkFlags & TBL_XOverRowHeader)
		{
			if (lLDblClkSepRow != TBL_Error)
			{
				PostMessage(m_hWndTbl, MTM_RowHdrSepLBtnDblClk, 0, lLDblClkSepRow);
				if (m_dwExtMsgOpt & MTEM_ROWHDRSEP_EXCLUSIVE)
					goto endlbtndblclk;
			}
			PostMessage(m_hWndTbl, MTM_RowHdrLBtnDblClk, 0, lLDblClkRow);
		}
		else if (dwLDblClkFlags & TBL_YOverSplitBar)
		{
			PostMessage(m_hWndTbl, MTM_SplitBarLBtnDblClk, 0, 0);
		}
		else
		{
			PostMessage(m_hWndTbl, MTM_AreaLBtnDblClk, (WPARAM)hWndLDblClkCol, lLDblClkRow);
		}

	endlbtndblclk:
		break;

	case WM_RBUTTONDOWN:
		// Wo ist die Maus?
		pt.x = iRDownX;
		pt.y = iRDownY;
		ObjFromPoint(pt, &lRDownRow, &hWndRDownCol, &dwRDownFlags, &dwMFlags);
		hWndRDownSepCol = GetSepCol(iRDownX, iRDownY);
		lRDownSepRow = GetSepRow(pt);
		bRDownCornerSep = (dwMFlags & MTOFP_OVERCORNERSEP);
		// Nachrichten generieren
		if (bRDownCornerSep)
		{
			PostMessage(m_hWndTbl, MTM_CornerSepRBtnDown, 0, 0);
			if (m_dwExtMsgOpt & MTEM_CORNERSEP_EXCLUSIVE)
				goto endrbtndown;
		}

		if (dwRDownFlags & TBL_XOverRowHeader && dwRDownFlags & TBL_YOverColumnHeader)
		{
			PostMessage(m_hWndTbl, MTM_CornerRBtnDown, 0, 0);
		}
		else if (dwRDownFlags & TBL_YOverColumnHeader)
		{
			if (hWndRDownSepCol)
			{
				PostMessage(m_hWndTbl, MTM_ColHdrSepRBtnDown, (WPARAM)hWndRDownSepCol, 0);
				if (m_dwExtMsgOpt & MTEM_COLHDRSEP_EXCLUSIVE)
					goto endrbtndown;
			}

			wParam = WPARAM(hWndRDownCol);
			if (hWndRDownCol && (dwMFlags & MTOFP_OVERCOLHDRGRP))
				lParam = LPARAM(m_ColHdrGrps->GetGrpOfCol(hWndRDownCol));
			else
				lParam = 0;
			PostMessage(m_hWndTbl, MTM_ColHdrRBtnDown, wParam, lParam);
		}
		else if (dwRDownFlags & TBL_XOverRowHeader)
		{
			if (lRDownSepRow != TBL_Error)
			{
				PostMessage(m_hWndTbl, MTM_RowHdrSepRBtnDown, 0, lRDownSepRow);
				if (m_dwExtMsgOpt & MTEM_ROWHDRSEP_EXCLUSIVE)
					goto endrbtndown;
			}
			PostMessage(m_hWndTbl, MTM_RowHdrRBtnDown, 0, lRDownRow);
		}
		else if (dwRDownFlags & TBL_YOverSplitBar)
		{
			PostMessage(m_hWndTbl, MTM_SplitBarRBtnDown, 0, 0);
		}
		else
		{
			PostMessage(m_hWndTbl, MTM_AreaRBtnDown, (WPARAM)hWndRDownCol, lRDownRow);
		}

	endrbtndown:
		break;

	case WM_RBUTTONUP:
		// Wo ist die Maus?
		pt.x = iRUpX;
		pt.y = iRUpY;
		ObjFromPoint(pt, &lRUpRow, &hWndRUpCol, &dwRUpFlags, &dwMFlags);
		hWndRUpSepCol = GetSepCol(iRUpX, iRUpY);
		lRUpSepRow = GetSepRow(pt);
		bRUpCornerSep = (dwMFlags & MTOFP_OVERCORNERSEP);
		// Nachrichten generieren
		if (bRUpCornerSep)
		{
			PostMessage(m_hWndTbl, MTM_CornerSepRBtnUp, 0, 0);
			if (m_dwExtMsgOpt & MTEM_CORNERSEP_EXCLUSIVE)
				goto endrbtnup;
		}

		if (dwRUpFlags & TBL_XOverRowHeader && dwRUpFlags & TBL_YOverColumnHeader)
		{
			PostMessage(m_hWndTbl, MTM_CornerRBtnUp, 0, 0);
		}
		else if (dwRUpFlags & TBL_YOverColumnHeader)
		{
			if (hWndRUpSepCol)
			{
				PostMessage(m_hWndTbl, MTM_ColHdrSepRBtnUp, (WPARAM)hWndLDownSepCol, 0);
				if (m_dwExtMsgOpt & MTEM_COLHDRSEP_EXCLUSIVE)
					goto endrbtnup;
			}

			wParam = WPARAM(hWndRUpCol);
			if (hWndRUpCol && (dwMFlags & MTOFP_OVERCOLHDRGRP))
				lParam = LPARAM(m_ColHdrGrps->GetGrpOfCol(hWndRUpCol));
			else
				lParam = 0;
			PostMessage(m_hWndTbl, MTM_ColHdrRBtnUp, wParam, lParam);

		}
		else if (dwRUpFlags & TBL_XOverRowHeader)
		{
			if (lRUpSepRow != TBL_Error)
			{
				PostMessage(m_hWndTbl, MTM_RowHdrSepRBtnUp, 0, lRUpSepRow);
				if (m_dwExtMsgOpt & MTEM_ROWHDRSEP_EXCLUSIVE)
					goto endrbtnup;
			}
			PostMessage(m_hWndTbl, MTM_RowHdrRBtnUp, 0, lRUpRow);
		}
		else if (dwRUpFlags & TBL_YOverSplitBar)
		{
			PostMessage(m_hWndTbl, MTM_SplitBarRBtnUp, 0, 0);
		}
		else
		{
			PostMessage(m_hWndTbl, MTM_AreaRBtnUp, (WPARAM)hWndRUpCol, lRUpRow);
		}

	endrbtnup:
		break;

	case WM_RBUTTONDBLCLK:
		// Wo ist die Maus?
		pt.x = iRDblClkX;
		pt.y = iRDblClkY;
		ObjFromPoint(pt, &lRDblClkRow, &hWndRDblClkCol, &dwRDblClkFlags, &dwMFlags);
		hWndRDblClkSepCol = MTblGetSepCol(m_hWndTbl, iRDblClkX, iRDblClkY);
		lRDblClkSepRow = GetSepRow(pt);
		bRDblClkCornerSep = (dwMFlags & MTOFP_OVERCORNERSEP);
		// Nachrichten generieren
		if (bRDblClkCornerSep)
		{
			PostMessage(m_hWndTbl, MTM_CornerSepRBtnDblClk, 0, 0);
			if (m_dwExtMsgOpt & MTEM_CORNERSEP_EXCLUSIVE)
				goto endrbtndblclk;
		}

		if (dwRDblClkFlags & TBL_XOverRowHeader && dwRDblClkFlags & TBL_YOverColumnHeader)
		{
			PostMessage(m_hWndTbl, MTM_CornerRBtnDblClk, 0, 0);
		}
		else if (dwRDblClkFlags & TBL_YOverColumnHeader)
		{
			if (hWndRDblClkSepCol)
			{
				PostMessage(m_hWndTbl, MTM_ColHdrSepRBtnDblClk, (WPARAM)hWndRDblClkSepCol, 0);
				if (m_dwExtMsgOpt & MTEM_COLHDRSEP_EXCLUSIVE)
					goto endrbtndblclk;
			}

			wParam = WPARAM(hWndRDblClkCol);
			if (hWndRDblClkCol && (dwMFlags & MTOFP_OVERCOLHDRGRP))
				lParam = LPARAM(m_ColHdrGrps->GetGrpOfCol(hWndRDblClkCol));
			else
				lParam = 0;
			PostMessage(m_hWndTbl, MTM_ColHdrRBtnDblClk, wParam, lParam);
		}
		else if (dwRDblClkFlags & TBL_XOverRowHeader)
		{
			if (lRDblClkSepRow != TBL_Error)
			{
				PostMessage(m_hWndTbl, MTM_RowHdrSepRBtnDblClk, 0, lRDblClkSepRow);
				if (m_dwExtMsgOpt & MTEM_ROWHDRSEP_EXCLUSIVE)
					goto endrbtndblclk;
			}
			PostMessage(m_hWndTbl, MTM_RowHdrRBtnDblClk, 0, lRDblClkRow);
		}
		else if (dwRDblClkFlags & TBL_YOverSplitBar)
		{
			PostMessage(m_hWndTbl, MTM_SplitBarRBtnDblClk, 0, 0);
		}
		else
		{
			PostMessage(m_hWndTbl, MTM_AreaRBtnDblClk, (WPARAM)hWndRDblClkCol, lRDblClkRow);
		}

	endrbtndblclk:
		break;
	}
}


// GenerateRowSelChangeMsgs
void MTBLSUBCLASS::GenerateRowSelChangeMsg()
{
	BOOL bSelAfter, bSelBefore;
	CMTblRow *pRow;
	// Normale Zeilen
	long lRow = -1;
	while (pRow = m_Rows->FindNextRow(&lRow))
	{
		bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
		bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
		if (bSelBefore != bSelAfter)
		{
			SendMessage(m_hWndTbl, MTM_RowSelChanged, 0, 0);
			return;
		}
	}

	// Split-Zeilen
	lRow = TBL_MinRow;
	while (pRow = m_Rows->FindNextSplitRow(&lRow))
	{
		bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
		bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
		if (bSelBefore != bSelAfter)
		{
			SendMessage(m_hWndTbl, MTM_RowSelChanged, 0, 0);
			return;
		}
	}
}


// GetImgRect
BOOL MTBLSUBCLASS::GetImgRect(CMTblImg &Img, long lRow, HWND hWndCol, long lCellLeft, long lCellTop, long lCellWid, long lCellHt, long lTextTop, int iFontHeight, long lCellLeadingX, long lCellLeadingY, BOOL bButton, BOOL bDDButton, LPRECT pr)
{
	if (!Img.GetHandle()) return FALSE;

	SIZE siImg;
	if (!Img.GetSize(siImg)) return FALSE;

	BOOL bClipFocusRect = QueryFlags(MTBL_FLAG_CLIP_FOCUS_AREA);
	BOOL bClipFocusRectTop = (bClipFocusRect && lRow != TBL_Error);

	BOOL bTile = Img.QueryFlags(MTSI_ALIGN_TILE);

	BOOL bHCenter = Img.QueryFlags(MTSI_ALIGN_HCENTER);
	BOOL bRight = Img.QueryFlags(MTSI_ALIGN_RIGHT);
	BOOL bLeft = !(bHCenter || bRight);

	BOOL bTop = Img.QueryFlags(MTSI_ALIGN_TOP);
	BOOL bVCenter = Img.QueryFlags(MTSI_ALIGN_VCENTER);
	BOOL bBottom = Img.QueryFlags(MTSI_ALIGN_BOTTOM);

	BOOL bNoLeadLeft = Img.QueryFlags(MTSI_NOLEAD_LEFT);
	BOOL bNoLeadTop = Img.QueryFlags(MTSI_NOLEAD_TOP);

	long lCellRight = lCellLeft + lCellWid;
	long lCellBottom = lCellTop + lCellHt;

	if (bDDButton)
	{
		int iCellType = GetEffCellType(lRow, hWndCol);
		if (iCellType == COL_CellType_DropDownList)
		{
			lCellRight -= MTBL_DDBTN_WID;
		}
	}

	CMTblBtn * pBtn = NULL;
	int iBtnWid = 0;
	if (bButton)
	{
		if (hWndCol && lRow != TBL_Error)
			pBtn = GetEffCellBtn(hWndCol, lRow);
		else if (hWndCol && lRow == TBL_Error)
			pBtn = m_Cols->GetHdrBtn(hWndCol);

		if (pBtn)
		{
			iBtnWid = CalcBtnWidth(pBtn, hWndCol, lRow);
			if (!pBtn->IsLeft())
				lCellRight -= iBtnWid;
		}
	}

	if (hWndCol && lRow != TBL_Error)
		lCellLeft = GetCellAreaLeft(lRow, hWndCol, lCellLeft);

	long lLeft = lCellLeft;
	if (pBtn && pBtn->IsLeft())
		lLeft += iBtnWid;
	if (!bNoLeadLeft || IsIndented(hWndCol, lRow))
		lLeft += lCellLeadingX;

	long lTop = lCellTop;
	if (!bNoLeadTop)
		lTop += (bClipFocusRectTop ? max(lCellLeadingY, 1) : lCellLeadingY);
	else if (bClipFocusRectTop)
		++lTop;

	long lRight = lCellRight;

	long lBottom = lCellBottom;
	if (bClipFocusRect)
		--lBottom;

	if (bTile)
	{
		pr->left = lLeft;
		pr->top = lTop;
		pr->right = lRight;
		pr->bottom = lBottom;
	}
	else
	{
		if (bLeft)
		{
			pr->left = lLeft;
		}
		else if (bHCenter)
		{
			pr->left = max(lLeft, lCellLeft + (((lCellRight - lCellLeft) - siImg.cx) / 2));
		}
		else if (bRight)
		{
			pr->left = max(lLeft, lCellRight - siImg.cx);
		}

		if (bTop)
		{
			pr->top = lTop;
		}
		else if (bVCenter)
		{
			pr->top = max(lTop, lCellTop + (((lCellBottom - lCellTop) - siImg.cy) / 2));
		}
		else if (bBottom)
		{
			pr->top = max(lTop, lCellBottom - siImg.cy);
		}
		else // MTSI_ALIGN_VTEXT
		{
			pr->top = max(lTop, lTextTop + ((iFontHeight - siImg.cy) / 2));
		}

		pr->right = pr->left + siImg.cx;
		pr->bottom = pr->top + siImg.cy;
	}

	return TRUE;
}


// GetItemAreaLeft
int MTBLSUBCLASS::GetItemAreaLeft(CMTblItem &item, long lItemLeft)
{
	if (item.IsCell())
	{
		int iIndent = GetEffCellIndent(item.GetRowNr(), item.WndHandle);
		if (iIndent)
			lItemLeft += iIndent;
	}

	return lItemLeft;
}


// GetItemBtnFont
HFONT MTBLSUBCLASS::GetItemBtnFont(LPMTBLITEM pItem, CMTblBtn * pBtn, LPBOOL pbMustDelete)
{
	// Check Parameter
	if (!(pItem && pBtn && pbMustDelete)) return NULL;

	// Lösch-KZ initialisieren
	*pbMustDelete = FALSE;

	// Font ermitteln
	HFONT hBtnFont = NULL;
	if (pBtn->m_dwFlags & MTBL_BTN_SHARE_FONT)
	{
		HDC hDC = GetDC(m_hWndTbl);
		if (hDC)
		{
			switch (pItem->Type)
			{
			case MTBL_ITEM_CELL:
				hBtnFont = GetEffCellFont(hDC, pItem->GetRowNr(), pItem->WndHandle, pbMustDelete);
				break;
			case MTBL_ITEM_COLHDR:
				hBtnFont = GetEffColHdrFont(hDC, pItem->WndHandle, pbMustDelete);
				break;
			}
			ReleaseDC(m_hWndTbl, hDC);
		}
	}
	else
		hBtnFont = (HFONT)SendMessage(m_hWndTbl, WM_GETFONT, 0, 0);

	// Fertig
	return hBtnFont;

}


// GetItemHighlighting
LPMTBLHILI MTBLSUBCLASS::GetItemHighlighting(CMTblItem &item)
{
	LPMTBLHILI phili = NULL;

	ItemList::iterator it = FindHighlightedItem(item);
	if (it != m_pHLIL->end())
	{
		HiLiDefMap::iterator it = m_pHLDM->find(LONG_PAIR(item.Type, item.Part));
		if (it != m_pHLDM->end())
		{
			phili = &(*it).second;
		}
	}

	return phili;
}

// GetItemLine
LPMTBLLINEDEF MTBLSUBCLASS::GetItemLine(CMTblItem &item, int iLineType)
{
	CMTblColHdrGrp* pCHG = NULL;
	CMTblLineDef* pLineDef = NULL;
	CMTblRow* pRow = NULL;

	switch (item.Type)
	{
	case MTBL_ITEM_CELL:
		pRow = item.GetRowPtr();
		if (pRow)
			pLineDef = pRow->m_Cols->GetLineDef(item.WndHandle, iLineType);
		break;
	case MTBL_ITEM_ROW:
		pRow = item.GetRowPtr();
		if (pRow)
			pLineDef = pRow->GetLineDef(iLineType);
		break;
	case MTBL_ITEM_ROWS:
		if (iLineType == MTLT_HORZ)
			pLineDef = m_pRowsHorzLineDef;
		break;
	case MTBL_ITEM_ROWHDR:
		pRow = item.GetRowPtr();
		if (pRow)
			pLineDef = pRow->GetHdrLineDef(iLineType);
		break;
	case MTBL_ITEM_COLUMN:
		pLineDef = m_Cols->GetLineDef(item.WndHandle, iLineType);
		break;
	case MTBL_ITEM_LAST_LOCKED_COLUMN:
		if (iLineType == MTLT_VERT)
			pLineDef = m_pLastLockedColVertLineDef;
		break;
	case MTBL_ITEM_COLUMNS:
		if (iLineType == MTLT_VERT)
			pLineDef = m_pColsVertLineDef;
		break;
	case MTBL_ITEM_COLHDR:
		pLineDef = m_Cols->GetHdrLineDef(item.WndHandle, iLineType);
		break;
	case MTBL_ITEM_COLHDRGRP:
		pCHG = (CMTblColHdrGrp*)item.Id;
		if (pCHG)
			pLineDef = pCHG->GetLineDef(iLineType);
		break;
	case MTBL_ITEM_CORNER:
		pLineDef = m_Corner->GetLineDef(iLineType);
		break;
	}

	return pLineDef;
}


// GetItemTextAreaLeft
int MTBLSUBCLASS::GetItemTextAreaLeft(CMTblItem& item, long lItemLeft, long lItemLeadingX, CMTblImg & Image, BOOL bButton /*= FALSE*/)
{
	long lLeft = lItemLeft;
	if (item.IsCell())
		GetCellAreaLeft(item.GetRowNr(), item.WndHandle, lItemLeft);

	if (bButton)
	{
		CMTblBtn * pBtn = GetEffItemBtn(item);
		if (pBtn && pBtn->IsLeft())
		{
			int iBtnWid = CalcBtnWidth(pBtn, item.WndHandle, item.GetRowNr());
			lLeft += iBtnWid;
		}
	}

	if (Image.GetHandle() && !(Image.QueryFlags(MTSI_ALIGN_HCENTER) || Image.QueryFlags(MTSI_ALIGN_RIGHT) || Image.QueryFlags(MTSI_ALIGN_TILE) || Image.QueryFlags(MTSI_NOTEXTADJ)))
	{
		SIZE s;
		if (Image.GetSize(s))
		{
			BOOL bIndented = FALSE;
			if (item.IsCell())
				bIndented = IsIndented(item.WndHandle, item.GetRowNr());

			int iOffset;
			if (Image.QueryFlags(MTSI_NOLEAD_LEFT) && !bIndented)
				iOffset = 0;
			else
				iOffset = lItemLeadingX;

			lLeft += (iOffset + s.cx + 1);
		}
	}

	return lLeft;
}


// GetItemTextAreaRight
int MTBLSUBCLASS::GetItemTextAreaRight(CMTblItem& item, long lItemRight, long lItemLeadingX, CMTblImg & Image, BOOL bButton /*= FALSE*/, BOOL bDDButton /*= FALSE*/)
{
	BOOL bMargin = FALSE;
	long lRight = lItemRight;

	if (bDDButton && item.IsCell())
	{
		int iCellType = GetEffCellType(item.GetRowNr(), item.WndHandle);
		if (iCellType == COL_CellType_DropDownList)
		{
			lRight -= MTBL_DDBTN_WID;
			bMargin = TRUE;
		}
	}

	if (bButton)
	{
		CMTblBtn * pBtn = GetEffItemBtn(item);
		if (pBtn && !pBtn->IsLeft())
		{
			int iBtnWid = CalcBtnWidth(pBtn, item.WndHandle, item.GetRowNr());
			lRight -= iBtnWid;
			bMargin = TRUE;
		}
	}

	if (Image.GetHandle() && Image.QueryFlags(MTSI_ALIGN_RIGHT) && !(Image.QueryFlags(MTSI_ALIGN_TILE) || Image.QueryFlags(MTSI_NOTEXTADJ)))
	{
		SIZE s;
		if (Image.GetSize(s))
		{
			lRight -= s.cx;
			bMargin = TRUE;
		}
	}

	if (bMargin)
		lRight -= 1;

	return lRight;
}


// GetLastLockedColumn
HWND MTBLSUBCLASS::GetLastLockedColumn()
{
	HWND hWndCol = NULL;
	int iLocked = Ctd.TblQueryLockedColumns(m_hWndTbl);
	if (iLocked > 0)
	{
		CMTblMetrics tm(m_hWndTbl);
		CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
		CURCOLS_WIDTH_VECTOR & CurColsWid = m_CurCols->GetVectorWidth();
		int iCurCol = 0, iCurCols = CurCols.size();
		int iColPos = 1, iColRight = tm.m_RowHeaderRight;

		for (; iCurCol < iCurCols; ++iCurCol)
		{
			iColRight += (CurColsWid[iCurCol] - 1);
			if (iColPos <= tm.m_LockedColumns && iColRight <= tm.m_LockedColumnsRight)
				hWndCol = CurCols[iCurCol];
			else
				break;
		}
	}

	return hWndCol;
}

// GetLastMergedCell
HWND MTBLSUBCLASS::GetLastMergedCell(HWND hWndCol, long lRow, LPMTBLMERGECELLS pMergeCells /*=NULL*/)
{
	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_MASTER, pMergeCells);
	if (pmc)
		return pmc->hWndColTo;
	else
		return NULL;
}


// GetLastMergedCellEx
BOOL MTBLSUBCLASS::GetLastMergedCellEx(HWND hWndCol, long lRow, HWND &hWndLastMergedCol, long &lLastMergedRow, LPINT piCols /*=NULL*/, LPINT piRows /*=NULL*/, LPMTBLMERGECELLS pMergeCells /*=NULL*/)
{
	hWndLastMergedCol = NULL;
	lLastMergedRow = TBL_Error;
	if (piCols)
		*piCols = 0;
	if (piRows)
		*piRows = 0;

	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_MASTER, pMergeCells);
	if (pmc)
	{
		hWndLastMergedCol = pmc->hWndColTo;
		lLastMergedRow = pmc->lRowTo;
		if (piCols)
			*piCols = pmc->iMergedCols;
		if (piRows)
			*piRows = pmc->iMergedRows;
		return TRUE;
	}
	else
		return FALSE;
}


// GetLastMergedRow
long MTBLSUBCLASS::GetLastMergedRow(long lRow, LPMTBLMERGEROWS pMergeRows /*=NULL*/)
{
	LPMTBLMERGEROW pmr = FindMergeRow(lRow, FMR_MASTER, pMergeRows);
	if (pmr)
		return pmr->lRowTo;
	else
		return TBL_Error;
}


// GetMasterCell
BOOL MTBLSUBCLASS::GetMasterCell(HWND &hWndCol, long &lRow)
{
	if (!hWndCol || lRow == TBL_Error)
		return FALSE;

	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_SLAVE);
	if (pmc)
	{
		hWndCol = pmc->hWndColFrom;
		lRow = pmc->lRowFrom;
	}

	return TRUE;
}


// GetMergeCell
HWND MTBLSUBCLASS::GetMergeCell(HWND hWndCol, long lRow)
{
	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_SLAVE);
	if (pmc)
		return pmc->hWndColFrom;
	else
		return NULL;
}


// GetMergeCellEx
BOOL MTBLSUBCLASS::GetMergeCellEx(HWND hWndCol, long lRow, HWND &hWndMergeCol, long &lMergeRow)
{
	hWndMergeCol = NULL;
	lMergeRow = TBL_Error;

	LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_SLAVE);
	if (pmc)
	{
		hWndMergeCol = pmc->hWndColFrom;
		lMergeRow = pmc->lRowFrom;
		return TRUE;
	}
	else
		return FALSE;
}


// GetMergeCells
LPMTBLMERGECELLS MTBLSUBCLASS::GetMergeCells(BOOL bCreateExternal /* = FALSE*/, WORD wRowFlagsOn /* = 0*/, WORD wRowFlagsOff /* = ROW_Hidden*/, DWORD dwColFlagsOn /* = COL_Visible */, DWORD dwColFlagsOff /* = 0 */)
{
	LPMTBLMERGECELLS pMergeCells = NULL;

	BOOL bAdd = FALSE;
	long lRowAdd = TBL_Error;
	HWND hWndColAdd = NULL;

	if (bCreateExternal)
	{
		// Return, wenn keine Zellenverbindungen existieren
		if (m_pCounter->GetMergeCells() < 1)
			return NULL;

		pMergeCells = new MTBLMERGECELLS;
		//pMergeCells->mcv.reserve( 100 );
	}
	else
	{
		// Hinzufügen?
		lRowAdd = m_lRowAddToMergeCells;
		m_lRowAddToMergeCells = TBL_Error;

		hWndColAdd = m_hWndColAddToMergeCells;
		m_hWndColAddToMergeCells = NULL;

		bAdd = (lRowAdd != TBL_Error && hWndColAdd && !m_bMergeCellsInvalid);

		if (!bAdd)
		{
			// Return, wenn bestehende Zellenverbindungen noch gültig sind
			if (!m_bMergeCellsInvalid)
				return m_pMergeCells;

			// Löschen
			FreeMergeCells();

			// Return, wenn keine Zellenverbindungen existieren
			if (m_pCounter->GetMergeCells() < 1)
			{
				m_bMergeCellsInvalid = FALSE;
				return NULL;
			}
		}

		if (!m_pMergeCells)
		{
			m_pMergeCells = new MTBLMERGECELLS;
			//m_pMergeCells->mcv.reserve( 100 );
		}

		pMergeCells = m_pMergeCells;
	}

	// Flags setzen
	pMergeCells->wRowFlagsOn = wRowFlagsOn;
	pMergeCells->wRowFlagsOff = wRowFlagsOff;
	pMergeCells->dwColFlagsOn = dwColFlagsOn;
	pMergeCells->dwColFlagsOff = dwColFlagsOff;

	// Tabellen-Metriken
	CMTblMetrics tm(m_hWndTbl);

	// Spalten ermitteln
	vector<LPMTBLCOLINFO2> Cols;
	Cols.reserve(100);

	BOOL bAddColFound = FALSE, bUseCol;
	HWND hWndCol;
	int iColPos = 1, iColRight = tm.m_RowHeaderRight, iCols = 0;
	LPMTBLCOLINFO2 pColInfo;

	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	CURCOLS_WIDTH_VECTOR & CurColsWid = m_CurCols->GetVectorWidth();
	int iAddColPos, iCurCol = 0, iCurCols = CurCols.size();
	/*while ( hWndCol = Ctd.TblGetColumnWindow( m_hWndTbl, iColPos, COL_GetPos ) )*/
	if (bAdd)
	{
		// Spalten nur ab der hinzuzufügenden Spalte ermitteln
		iAddColPos = m_CurCols->GetPos(hWndColAdd);
		if (iAddColPos > 0)
		{
			iColPos = iAddColPos;
			iCurCol = iAddColPos - 1;
		}
	}
	for (; iCurCol < iCurCols; ++iCurCol)
	{
		hWndCol = CurCols[iCurCol];

		bUseCol = TRUE;
		if (dwColFlagsOn)
		{
			if (!SendMessage(hWndCol, TIM_QueryColumnFlags, 0, dwColFlagsOn))
				bUseCol = FALSE;
		}

		if (dwColFlagsOff)
		{
			if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, dwColFlagsOff))
				bUseCol = FALSE;
		}

		if (bUseCol)
		{
			//iColRight += ( SendMessage( hWndCol, TIM_QueryColumnWidth, 0, 0 ) - 1 );
			iColRight += (CurColsWid[iCurCol] - 1);

			pColInfo = new MTBLCOLINFO2;
			pColInfo->hWnd = hWndCol;
			pColInfo->iPos = iColPos;
			pColInfo->iCellType = GetEffCellType(TBL_Error, hWndCol);
			pColInfo->bLocked = (iColPos <= tm.m_LockedColumns && iColRight <= tm.m_LockedColumnsRight);

			Cols.push_back(pColInfo);
			++iCols;
		}
		++iColPos;
	}

	// Aktuellen Zeilenkontext ermitteln
	long lContext = Ctd.TblQueryContext(m_hWndTbl);

	// Zeilen durchlaufen...
	BOOL bBreak, bSplitRows;
	CMTblCol *pCell;
	CMTblRow *pri, *prj, *pRowTo;
	CMTblRow **Rows;
	int i = 0, iMax = 2, iMaxColPos, iMaxRowPos, cd, cd2, cd3, ci, cj, ck, cl, cm, cm2, rd, ri, rj, rm, rp;
	long lUpperBound, lLastValidPos;
	LPMTBLCOLINFO2 pci, pcj, pck, pcl, pColTo;
	LPMTBLMERGECELL pmc;

	if (bAdd)
	{
		m_Rows->EnsureValid(lRowAdd);

		if (lRowAdd >= 0)
			// Nur normale Zeilen
			iMax = 1;
		else
			// Nur Split-Zeilen
			i = 1;
	}

	for (; i < iMax; ++i)
	{
		// Split-Zeilen?
		bSplitRows = (i > 0);

		// Zeilenarray ermitteln
		Rows = m_Rows->GetRowArray(bSplitRows ? -1 : 0, &lUpperBound, &lLastValidPos);
		if (!Rows)
			continue;

		// Start und Ende festlegen		
		if (bAdd)
		{
			// Erste sichtbare Position = sichtbare Position der hinzuzufügenden Zeile
			rp = m_Rows->GetRowVisPos(lRowAdd);
			// Nur die hinzuzufügende Zeile
			ri = (bSplitRows ? m_Rows->GetSplitRowPos(lRowAdd) : lRowAdd);
			iMaxRowPos = ri;
		}
		else
		{
			// Erste sichtbare Position = 0
			rp = 0;
			// Alle Zeilen bis zur letzten gültigen Position
			ri = 0;
			iMaxRowPos = lLastValidPos;
		}
		for (; ri <= iMaxRowPos; ++ri)
		{
			if (!(pri = Rows[ri]))
				continue;

			if (pri->IsDelAdj())
				continue;

			if (wRowFlagsOn)
			{
				if (wRowFlagsOn == ROW_Hidden && !pri->QueryHiddenFlag())
					continue;
				else
				{
					if (!Ctd.TblQueryRowFlags(m_hWndTbl, pri->GetNr(), wRowFlagsOn))
						continue;
				}
			}

			if (wRowFlagsOff)
			{
				if (wRowFlagsOff == ROW_Hidden && pri->QueryHiddenFlag())
					continue;
				else
				{
					if (Ctd.TblQueryRowFlags(m_hWndTbl, pri->GetNr(), wRowFlagsOff))
						continue;
				}
			}


			// Spalten durchlaufen
			if (bAdd)
				// Nur die hinzuzufügende Spalte ( = erste ermittelte Spalte )
				iMaxColPos = 0;
			else
				// Alle Spalten
				iMaxColPos = iCols - 1;
			for (ci = 0; ci <= iMaxColPos; ++ci)
			{
				pci = Cols[ci];

				// Letzte verbundene Spalte ermitteln
				cd = 0;
				pColTo = pci;
				if (cm = pri->m_Cols->GetMerged(pci->hWnd))
				{
					for (cj = ci + 1; cj < iCols && cd < cm; ++cj)
					{
						pcj = Cols[cj];

						if (!pcj->hWnd || pcj->bLocked != pci->bLocked /*|| pri->m_Cols->GetMerged( pcj->hWnd ) || pri->m_Cols->GetMergedRows( pcj->hWnd )*/)
							break;

						pCell = pri->m_Cols->Find(pcj->hWnd);
						if (pCell && (pCell->m_Merged || pCell->m_MergedRows))
							break;

						pColTo = pcj;
						++cd;
					}
				}

				// Letzte verbundene Zeile ermitteln
				rd = 0;
				pRowTo = pri;
				if (rm = pri->m_Cols->GetMergedRows(pci->hWnd))
				{
					for (rj = ri + 1; rd < rm; ++rj)
					{
						prj = NULL;

						if (rj <= lLastValidPos)
							prj = Rows[rj];
						else
						{
							if (bSplitRows || !Ctd.TblSetContextOnly(m_hWndTbl, rj))
								break;
						}

						if (!prj)
						{
							prj = m_Rows->EnsureValid(bSplitRows ? m_Rows->GetSplitPosRow(rj) : rj);
							if (!prj)
								break;
						}

						if (prj->IsDelAdj())
							continue;

						if (wRowFlagsOn)
						{
							if (wRowFlagsOn == ROW_Hidden && !prj->QueryHiddenFlag())
								continue;
							else
							{
								if (!Ctd.TblQueryRowFlags(m_hWndTbl, prj->GetNr(), wRowFlagsOn))
									continue;
							}
						}

						if (wRowFlagsOff)
						{
							if (wRowFlagsOff == ROW_Hidden && prj->QueryHiddenFlag())
								continue;
							else
							{
								if (Ctd.TblQueryRowFlags(m_hWndTbl, prj->GetNr(), wRowFlagsOff))
									continue;
							}
						}

						if (prj->GetLevel() != pri->GetLevel())
							break;

						bBreak = FALSE;
						for (ck = ci, cd2 = 0; cd2 <= cd; ++ck, ++cd2)
						{
							pck = Cols[ck];

							/*if ( prj->m_Cols->GetMerged( pck->hWnd ) || prj->m_Cols->GetMergedRows( pck->hWnd ) )
							{
							bBreak = TRUE;
							break;
							}*/

							pCell = prj->m_Cols->Find(pck->hWnd);
							if (pCell && (pCell->m_Merged || pCell->m_MergedRows))
							{
								bBreak = TRUE;
								break;
							}


							for (cl = ck - 1, cd3 = 1; cl >= 0; --cl, ++cd3)
							{
								pcl = Cols[cl];

								if (!pcl->hWnd || pcl->bLocked != pck->bLocked)
									break;

								if (cm2 = prj->m_Cols->GetMerged(pcl->hWnd))
								{
									if (cm2 >= cd3)
										bBreak = TRUE;
									break;
								}
							}
						}
						if (bBreak)
							break;

						pRowTo = prj;
						++rd;
					}
				}

				// Evtl. Objekt erzeugen und hinzufügen
				if (cd || rd)
				{
					pmc = new MTBLMERGECELL;

					pmc->iType = rd ? MC_EXTENDED : MC_SIMPLE;

					pmc->hWndColFrom = pci->hWnd;
					pmc->hWndColTo = pColTo->hWnd;

					pmc->iColPosFrom = pci->iPos;
					pmc->iColPosTo = pColTo->iPos;

					pmc->lRowFrom = pri->GetNr();
					pmc->lRowTo = pRowTo->GetNr();

					pmc->lRowVisPosFrom = rp;
					pmc->lRowVisPosTo = rp + rd;

					pmc->iMergedCols = cd;
					pmc->iMergedRows = rd;

					//pMergeCells->mcv.push_back( pmc );

					// Zur Map hinzufügen
					for (long lr = pmc->lRowFrom; lr <= pmc->lRowTo; ++lr)
					{
						for (long lc = pmc->iColPosFrom; lc <= pmc->iColPosTo; ++lc)
						{
							pMergeCells->mcm.insert(MTBLMERGECELL_MAP::value_type(LONG_PAIR(lr, lc), pmc));
						}
					}
					//pMergeCells->mcm.insert( MTBLMERGECELL_MAP::value_type( LONG_PAIR_PAIR( LONG_PAIR(pmc->lRowFrom, pmc->lRowTo), LONG_PAIR(pmc->iColPosFrom, pmc->iColPosTo) ), pmc ) );
					//pMergeCells->mcm.insert( MTBLMERGECELL_MAP::value_type( LONG_PAIR(pmc->lRowFrom, pmc->iColPosFrom), pmc ) );
				}
			}
			++rp;
		}
	}

	// Zeilenkontext wiederherstellen
	Ctd.TblSetContextOnly(m_hWndTbl, lContext);

	// Spalten freigeben
	vector<LPMTBLCOLINFO2>::size_type s, sMax = Cols.size();
	for (s = 0; s < sMax; ++s)
	{
		pColInfo = Cols[s];
		if (pColInfo)
		{
			delete pColInfo;
			pColInfo = NULL;
		}
	}

	if (sMax)
		Cols.erase(Cols.begin(), Cols.end());

	// Fertig
	if (!bCreateExternal)
		m_bMergeCellsInvalid = FALSE;
	return pMergeCells;
}


// GetMergeCellType
int MTBLSUBCLASS::GetMergeCellType(HWND hWndCol, long lRow)
{
	CMTblRow * pRow = m_Rows->GetRowPtr(lRow);
	if (!pRow) return MC_NONE;
	if (pRow->QueryHiddenFlag()) return MC_NONE;
	if (pRow->IsDelAdj()) return MC_NONE;

	if (!SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible)) return MC_NONE;

	long lMergedCols = pRow->m_Cols->GetMerged(hWndCol);
	long lMergedRows = pRow->m_Cols->GetMergedRows(hWndCol);
	if (lMergedRows > 0)
		return MC_EXTENDED;
	else if (lMergedCols > 0)
		return MC_SIMPLE;

	return MC_NONE;
}


// GetMergeRow
long MTBLSUBCLASS::GetMergeRow(long lRow)
{
	LPMTBLMERGEROW pmr = FindMergeRow(lRow, FMR_SLAVE);
	if (pmr)
		return pmr->lRowFrom;
	else
		return TBL_Error;
}


// GetMergeRows
LPMTBLMERGEROWS MTBLSUBCLASS::GetMergeRows(BOOL bCreateExternal /* = FALSE*/, WORD wRowFlagsOn /* = 0*/, WORD wRowFlagsOff /* = ROW_Hidden*/)
{
	LPMTBLMERGEROWS pMergeRows = NULL;

	if (bCreateExternal)
	{
		// Return, wenn keine Zellenverbindungen existieren
		if (m_pCounter->GetMergeRows() < 1)
			return NULL;

		pMergeRows = new MTBLMERGEROWS;
		//pMergeRows->mrv.reserve( 100 );
	}
	else
	{
		// Return, wenn bestehende noch gültig
		if (!m_bMergeRowsInvalid)
			return m_pMergeRows;

		// Array löschen
		FreeMergeRows();

		// Return, wenn keine Zellenverbindungen existieren
		if (m_pCounter->GetMergeRows() < 1)
		{
			m_bMergeRowsInvalid = FALSE;
			return NULL;
		}

		m_pMergeRows = new MTBLMERGEROWS;
		//m_pMergeRows->mrv.reserve( 100 );
		pMergeRows = m_pMergeRows;
	}

	// Flags setzen
	pMergeRows->wRowFlagsOn = wRowFlagsOn;
	pMergeRows->wRowFlagsOff = wRowFlagsOff;

	// Tabellen-Metriken
	CMTblMetrics tm(m_hWndTbl);

	// Aktuellen Zeilenkontext ermitteln
	long lContext = Ctd.TblQueryContext(m_hWndTbl);

	BOOL bSplitRows;
	CMTblRow *pri, *prj, *pRowTo;
	CMTblRow **Rows;
	int i, rd, ri, rj, rm, rp;
	long lUpperBound, lLastValidPos;
	LPMTBLMERGEROW pmr;

	for (i = 0; i < 2; ++i)
	{
		bSplitRows = (i > 0);

		Rows = m_Rows->GetRowArray(bSplitRows ? -1 : 0, &lUpperBound, &lLastValidPos);
		if (!Rows)
			continue;

		rp = 0;
		for (ri = 0; ri <= lLastValidPos; ++ri)
		{
			if (!(pri = Rows[ri]))
				continue;

			if (pri->IsDelAdj())
				continue;

			if (wRowFlagsOn)
			{
				if (wRowFlagsOn == ROW_Hidden && !pri->QueryHiddenFlag())
					continue;
				else
				{
					if (!Ctd.TblQueryRowFlags(m_hWndTbl, pri->GetNr(), wRowFlagsOn))
						continue;
				}
			}

			if (wRowFlagsOff)
			{
				if (wRowFlagsOff == ROW_Hidden && pri->QueryHiddenFlag())
					continue;
				else
				{
					if (Ctd.TblQueryRowFlags(m_hWndTbl, pri->GetNr(), wRowFlagsOff))
						continue;
				}
			}

			// Letzte verknüpfte Zeile ermitteln
			rd = 0;
			pRowTo = pri;
			if (rm = pri->m_Merged)
			{
				for (rj = ri + 1; rd < rm; ++rj)
				{
					prj = NULL;

					if (rj <= lLastValidPos)
						prj = Rows[rj];
					else
					{
						if (bSplitRows || !Ctd.TblSetContextOnly(m_hWndTbl, rj))
							break;
					}

					if (!prj)
					{
						prj = m_Rows->EnsureValid(bSplitRows ? m_Rows->GetSplitPosRow(rj) : rj);
						if (!prj)
							break;
					}

					if (prj->IsDelAdj())
						continue;

					if (wRowFlagsOn)
					{
						if (wRowFlagsOn == ROW_Hidden && !prj->QueryHiddenFlag())
							continue;
						else
						{
							if (!Ctd.TblQueryRowFlags(m_hWndTbl, prj->GetNr(), wRowFlagsOn))
								continue;
						}
					}

					if (wRowFlagsOff)
					{
						if (wRowFlagsOff == ROW_Hidden && prj->QueryHiddenFlag())
							continue;
						else
						{
							if (Ctd.TblQueryRowFlags(m_hWndTbl, prj->GetNr(), wRowFlagsOff))
								continue;
						}
					}

					if (prj->GetLevel() != pri->GetLevel())
						break;

					if (prj->m_Merged > 0)
						break;

					pRowTo = prj;
					++rd;
				}
			}

			// Evtl. Objekt erzeugen und hinzufügen
			if (rd)
			{
				pmr = new MTBLMERGEROW;

				pmr->lRowFrom = pri->GetNr();
				pmr->lRowTo = pRowTo->GetNr();

				pmr->lRowVisPosFrom = rp;
				pmr->lRowVisPosTo = rp + rd;

				pmr->iMergedRows = rd;

				// Zur Map hinzufügen
				for (long lr = pmr->lRowFrom; lr <= pmr->lRowTo; ++lr)
				{
					pMergeRows->mrm.insert(MTBLMERGEROW_MAP::value_type(lr, pmr));
				}

			}

			++rp;
		}
	}

	// Zeilenkontext wiederherstellen
	Ctd.TblSetContextOnly(m_hWndTbl, lContext);

	// Fertig
	m_bMergeRowsInvalid = FALSE;
	return pMergeRows;
}


// GetMergeRowRange
BOOL MTBLSUBCLASS::GetMergeRowRange(long lRow, DWORD dwFlags, RowRange & rr, LPMTBLMERGECELLS pMergeCells /*= NULL*/)
{
	typedef pair< long, DWORD > ROWSEARCH;

	BOOL bRange = FALSE;
	long l, lMergedCells, lRowFrom, lRowTo;
	LPMTBLMERGECELL pmc;
	LPMTBLMERGEROW pmr;
	stack<ROWSEARCH> st;
	vector<LPMTBLMERGECELL> vmc;

	st.push(ROWSEARCH(lRow, dwFlags));
	lRowFrom = lRowTo = lRow;

	while (!st.empty())
	{
		ROWSEARCH rsc = st.top();
		st.pop();

		pmr = FindMergeRow(rsc.first, FMR_ANY);
		if (pmr)
		{
			if (rsc.second & GCMRR_UP && pmr->lRowFrom < lRowFrom)
			{
				lRowFrom = pmr->lRowFrom;
				st.push(ROWSEARCH(lRowFrom, GCMRR_UP));
				bRange = TRUE;
			}

			if (rsc.second & GCMRR_DOWN && pmr->lRowTo > lRowTo)
			{
				lRowTo = pmr->lRowTo;
				st.push(ROWSEARCH(lRowTo, GCMRR_DOWN));
				bRange = TRUE;
			}
		}

		vmc.clear();
		lMergedCells = FindMergeCells(NULL, rsc.first, FMC_ANY, vmc, 0, pMergeCells);
		for (l = 0; l < lMergedCells; l++)
		{
			pmc = vmc[l];

			if (rsc.second & GCMRR_UP && pmc->lRowFrom < lRowFrom)
			{
				lRowFrom = pmc->lRowFrom;
				st.push(ROWSEARCH(lRowFrom, GCMRR_UP));
				bRange = TRUE;
			}

			if (rsc.second & GCMRR_DOWN && pmc->lRowTo > lRowTo)
			{
				lRowTo = pmc->lRowTo;
				st.push(ROWSEARCH(lRowTo, GCMRR_DOWN));
				bRange = TRUE;
			}
		}
	}

	if (bRange)
	{
		rr.first = lRowFrom;
		rr.second = lRowTo;
	}

	return bRange;
}


// GetNodeRect
BOOL MTBLSUBCLASS::GetNodeRect(long lCellLeft, long lTextTop, int iFontHeight, int iLevel, int iCellLeadingX, int iCellLeadingY, LPRECT pr)
{
	if (QueryTreeFlags(MTBL_TREE_FLAG_FLAT_STRUCT))
		iLevel = 0;

	if (QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE) && iLevel > 0)
		iLevel--;

	//pr->left = lCellLeft + iCellLeadingX + ( iLevel * m_iIndent ) + ( ( m_iIndent - m_iNodeSize ) / 2 );
	pr->left = lCellLeft + (iCellLeadingX - 1) + (iLevel * m_iIndent) + ((m_iIndent - m_iNodeSize) / 2);
	pr->top = lTextTop + (max(1, (iFontHeight - m_iNodeSize) / 2));
	pr->right = pr->left + m_iNodeSize;
	pr->bottom = pr->top + m_iNodeSize;

	return TRUE;
}


// GetNormalColor
// Liefert die normale Farbe ( entfernt 0x02000000 )
COLORREF MTBLSUBCLASS::GetNormalColor(COLORREF clr)
{
	return RGB(GetRValue(clr), GetGValue(clr), GetBValue(clr));
}


// GetPaintMergeCells
void MTBLSUBCLASS::GetPaintMergeCells()
{
	// Vorhandene löschen
	FreePaintMergeCells();

	// Merge-Zellen ermitteln
	if (!GetMergeCells())
		return;

	// Merge-Zellen-Array da?
	if (!m_pMergeCells)
		return;

	// Anzahl Merge-Zellen ermitteln
	//int iMergeCells = m_pMergeCells->mcv.size( );
	int iMergeCells = m_pMergeCells->mcm.size();
	if (iMergeCells < 1)
		return;

	// Platz reservieren
	m_ppd->PaintMergeCells.reserve(iMergeCells);

	// Durchlaufen
	BOOL bSplit;
	int iScrPosV;
	long lRowsTop;
	LPMTBLMERGECELL pmc;
	LPMTBLPAINTCOL pColFrom, pColTo;
	LPMTBLPAINTMERGECELL ppmc;
	RECT r;
	MTBLMERGECELL_MAP::iterator it, itEnd = m_pMergeCells->mcm.end();
	for (it = m_pMergeCells->mcm.begin(); it != itEnd; ++it)
	{
		pmc = (*it).second;
		if (!pmc)
			continue;

		bSplit = pmc->lRowFrom < 0;
		lRowsTop = bSplit ? m_ppd->tm.m_SplitRowsTop : m_ppd->tm.m_RowsTop;
		iScrPosV = bSplit ? m_ppd->iScrPosVSplit : m_ppd->iScrPosV;

		pColFrom = FindPaintCol(pmc->hWndColFrom);
		pColTo = FindPaintCol(pmc->hWndColTo);

		// Rechteck ermitteln
		long lLogX1, lLogX2;

		r.left = pColFrom->ptX.x;
		r.right = pColTo->ptX.y;

		if (m_pCounter->GetRowHeights() <= 0)
		{
			r.top = lRowsTop + ((pmc->lRowVisPosFrom - iScrPosV) * m_ppd->tm.m_RowHeight);
			r.bottom = r.top + ((pmc->iMergedRows + 1) * m_ppd->tm.m_RowHeight) - 1;
		}
		else
		{
			// Zeile ermitteln, deren sichtbare Position der Scrollposition entspricht ( = oberste Zeile im sichtbaren Bereich )
			long lRowScrPos = m_Rows->GetVisPosRow(iScrPosV, bSplit);

			// Logische X-Position der obersten sichtbaren Zeile ermitteln
			long lxRowTop = m_Rows->GetRowLogX(lRowScrPos, m_ppd->tm.m_RowHeight);

			// Logische X-Position der Von-Zeile ermitteln
			lLogX1 = m_Rows->GetRowLogX(pmc->lRowFrom, m_ppd->tm.m_RowHeight);

			// Logische X-Position der Bis-Zeile ermitteln
			if (pmc->iMergedRows > 0)
				lLogX2 = m_Rows->GetRowLogX(pmc->lRowTo, m_ppd->tm.m_RowHeight);
			else
				lLogX2 = lLogX1;

			// Physische X-Position berechnen
			r.top = lRowsTop + (lLogX1 - lxRowTop);

			// Position Y
			r.bottom = lRowsTop + (lLogX2 - lxRowTop) + GetEffRowHeight(pmc->lRowTo, &m_ppd->tm) - 1;
		}

		if (!(r.left >= m_ppd->rUpd.right || r.right < m_ppd->rUpd.left || r.top >= m_ppd->rUpd.bottom || r.bottom < m_ppd->rUpd.top))
		{
			ppmc = new MTBLPAINTMERGECELL;

			ppmc->iType = pmc->iType;

			ppmc->pColFrom = pColFrom;
			ppmc->pColTo = pColTo;

			ppmc->pRowFrom = m_Rows->GetRowPtr(pmc->lRowFrom);
			ppmc->pRowTo = m_Rows->GetRowPtr(pmc->lRowTo);

			ppmc->iMergedCols = pmc->iMergedCols;
			ppmc->iMergedRows = pmc->iMergedRows;

			if (m_pCounter->GetRowHeights() <= 0)
			{
				ppmc->dwTopLog = lRowsTop + (pmc->lRowVisPosFrom * m_ppd->tm.m_RowHeight);
				ppmc->dwBottomLog = ppmc->dwTopLog + ((pmc->iMergedRows + 1) * m_ppd->tm.m_RowHeight) - 1;
			}
			else
			{
				ppmc->dwTopLog = lRowsTop + lLogX1;
				ppmc->dwBottomLog = lRowsTop + lLogX2 + GetEffRowHeight(pmc->lRowTo, &m_ppd->tm) - 1;
			}

			ppmc->r = r;

			m_ppd->PaintMergeCells.push_back(ppmc);
		}
	}
}


// GetPaintMergeRowHdrs
void MTBLSUBCLASS::GetPaintMergeRowHdrs()
{
	// Vorhandene löschen
	FreePaintMergeRowHdrs();

	// Merge-Zeilen ermitteln
	if (!GetMergeRows())
		return;

	// Merge-Zeilen-Array da?
	if (!m_pMergeRows)
		return;

	// Anzahl Merge-Zeilen ermitteln
	//int iMergeRows = m_pMergeRows->mrv.size( );
	int iMergeRows = m_pMergeRows->mrm.size();
	if (iMergeRows < 1)
		return;

	// Paintcol für Rowheader ermitteln
	LPMTBLPAINTCOL pCol = FindPaintCol(NULL);
	if (!pCol)
		return;

	// Platz reservieren
	m_ppd->PaintMergeRowHdrs.reserve(iMergeRows);

	// Durchlaufen
	BOOL bSplit;
	int iScrPosV;
	long lRowsTop;
	LPMTBLMERGEROW pmr;
	LPMTBLPAINTMERGEROWHDR ppmrh;
	RECT r;

	/*for ( int i = 0; i < iMergeRows; ++i )*/
	MTBLMERGEROW_MAP::iterator it, itEnd = m_pMergeRows->mrm.end();
	for (it = m_pMergeRows->mrm.begin(); it != itEnd; ++it)
	{
		//pmr = m_pMergeRows->mrv[i];
		pmr = (*it).second;
		if (!pmr)
			continue;

		bSplit = pmr->lRowFrom < 0;
		lRowsTop = bSplit ? m_ppd->tm.m_SplitRowsTop : m_ppd->tm.m_RowsTop;
		iScrPosV = bSplit ? m_ppd->iScrPosVSplit : m_ppd->iScrPosV;

		// Rechteck ermitteln
		long lLogX1, lLogX2;

		r.left = m_ppd->tm.m_RowHeaderLeft;
		r.right = m_ppd->tm.m_RowHeaderRight;

		if (m_pCounter->GetRowHeights() <= 0)
		{
			r.top = lRowsTop + ((pmr->lRowVisPosFrom - iScrPosV) * m_ppd->tm.m_RowHeight);
			r.bottom = r.top + ((pmr->iMergedRows + 1) * m_ppd->tm.m_RowHeight) - 1;
		}
		else
		{
			// Zeile ermitteln, deren sichtbare Position der Scrollposition entspricht ( = oberste Zeile im sichtbaren Bereich )
			long lRowScrPos = m_Rows->GetVisPosRow(iScrPosV, bSplit);

			// Logische X-Position der obersten sichtbaren Zeile ermitteln
			long lxRowTop = m_Rows->GetRowLogX(lRowScrPos, m_ppd->tm.m_RowHeight);

			// Logische X-Position der Von-Zeile ermitteln
			lLogX1 = m_Rows->GetRowLogX(pmr->lRowFrom, m_ppd->tm.m_RowHeight);

			// Logische X-Position der Bis-Zeile ermitteln
			if (pmr->iMergedRows > 0)
				lLogX2 = m_Rows->GetRowLogX(pmr->lRowTo, m_ppd->tm.m_RowHeight);
			else
				lLogX2 = lLogX1;

			// Physische X-Position berechnen
			r.top = lRowsTop + (lLogX1 - lxRowTop);

			// Position Y
			r.bottom = lRowsTop + (lLogX2 - lxRowTop) + GetEffRowHeight(pmr->lRowTo, &m_ppd->tm) - 1;
		}

		if (!(r.left >= m_ppd->rUpd.right || r.right < m_ppd->rUpd.left || r.top >= m_ppd->rUpd.bottom || r.bottom < m_ppd->rUpd.top))
		{
			ppmrh = new MTBLPAINTMERGEROWHDR;

			ppmrh->pCol = pCol;

			ppmrh->pRowFrom = m_Rows->GetRowPtr(pmr->lRowFrom);
			ppmrh->pRowTo = m_Rows->GetRowPtr(pmr->lRowTo);

			ppmrh->iMergedRows = pmr->iMergedRows;

			if (m_pCounter->GetRowHeights() <= 0)
			{
				ppmrh->dwTopLog = lRowsTop + (pmr->lRowVisPosFrom * m_ppd->tm.m_RowHeight);
				ppmrh->dwBottomLog = ppmrh->dwTopLog + ((pmr->iMergedRows + 1) * m_ppd->tm.m_RowHeight) - 1;
			}
			else
			{
				ppmrh->dwTopLog = lRowsTop + lLogX1;
				ppmrh->dwBottomLog = lRowsTop + lLogX2 + GetEffRowHeight(pmr->lRowTo, &m_ppd->tm) - 1;
			}

			ppmrh->r = r;

			m_ppd->PaintMergeRowHdrs.push_back(ppmrh);
		}
	}
}


// GetRowAlternateBackColor
DWORD MTBLSUBCLASS::GetRowAlternateBackColor(long lRow)
{
	DWORD dwColor1 = (lRow >= 0) ? m_dwRowAlternateBackColor1 : m_dwSplitRowAlternateBackColor1;
	DWORD dwColor2 = (lRow >= 0) ? m_dwRowAlternateBackColor2 : m_dwSplitRowAlternateBackColor2;

	if (dwColor1 != MTBL_COLOR_UNDEF || dwColor2 != MTBL_COLOR_UNDEF)
	{
		long lVisPos = m_Rows->GetRowVisPos(lRow);
		if (lVisPos != TBL_Error)
			return (lVisPos % 2) ? dwColor2 : dwColor1;
	}

	return MTBL_COLOR_UNDEF;
}


// GetRowFlagImg
HANDLE MTBLSUBCLASS::GetRowFlagImg(DWORD dwFlag, BOOL bUserDefined)
{
	if (dwFlag)
	{
		RowFlagImgMap::iterator it;
		it = m_RowFlagImages->find(dwFlag);
		if (it != m_RowFlagImages->end())
		{
			if (bUserDefined)
				return (*it).second.second->GetHandle();
			else
				return (*it).second.first->GetHandle();
		}
	}

	return NULL;
}


// GetRowFocusRect
BOOL MTBLSUBCLASS::GetRowFocusRect(long lRow, LPRECT pRect)
{
	if (!pRect)
		return FALSE;

	SetRectEmpty(pRect);

	// Tabellenmaße
	CMTblMetrics tm(m_hWndTbl);

	// Y-Koordinaten der Zeile ermitteln
	POINT pt;
	if (lRow != TBL_Error)
	{
		POINT ptVis;

		LPMTBLMERGEROW pmr = FindMergeRow(lRow, FMR_ANY);
		if (pmr)
		{
			POINT pt1, pt2;

			//long lRet1 = MTblGetRowYCoord( m_hWndTbl, pmr->lRowFrom, &tm, &pt1, &ptVis, FALSE );
			long lRet1 = GetRowYCoord(pmr->lRowFrom, &tm, &pt1, &ptVis, FALSE);
			if (lRet1 == MTGR_RET_ERROR)
				return FALSE;

			// long lRet2 = MTblGetRowYCoord( m_hWndTbl, pmr->lRowTo, &tm, &pt2, &ptVis, FALSE );
			long lRet2 = GetRowYCoord(pmr->lRowTo, &tm, &pt2, &ptVis, FALSE);
			if (lRet2 == MTGR_RET_ERROR)
				return FALSE;

			if (lRet1 == MTGR_RET_INVISIBLE && lRet2 == MTGR_RET_INVISIBLE)
				return FALSE;

			pt.x = pt1.x;
			pt.y = pt2.y;

		}
		else
		{
			//long lRet = MTblGetRowYCoord( m_hWndTbl, lRow, &tm, &pt, &ptVis );
			long lRet = GetRowYCoord(lRow, &tm, &pt, &ptVis);
			if (lRet == MTGR_RET_ERROR || lRet == MTGR_RET_INVISIBLE)
				return FALSE;
		}


	}
	else
	{
		/*if ( Ctd.TblAnyRows( m_hWndTbl, 0, ROW_Hidden ) )
		return FALSE;*/

		pt.x = tm.m_RowsTop;
		pt.y = pt.x + tm.m_RowHeight - 1;
	}

	pRect->left = tm.m_ClientRect.left;
	pRect->top = pt.x - m_lFocusFrameOffsY - m_lFocusFrameThick;
	pRect->right = tm.m_ClientRect.right;
	pRect->bottom = pt.y + m_lFocusFrameOffsY + m_lFocusFrameThick;

	return TRUE;
}


// GetRowHdrFlags
WORD MTBLSUBCLASS::GetRowHdrFlags()
{
	HSTRING hsText = 0;
	HWND hWndCol;
	int iWid;
	WORD wFlags = 0;
	Ctd.TblQueryRowHeader(m_hWndTbl, &hsText, 254, &iWid, &wFlags, &hWndCol);

	if (hsText)
		Ctd.HStringUnRef(hsText);

	return wFlags;
}


// GetRowHdrRect
int MTBLSUBCLASS::GetRowHdrRect(long lRow, LPRECT pr, DWORD dwFlags /*=0*/)
{
	int iRet;
	// Tabellenmaße
	CMTblMetrics tm(m_hWndTbl);

	if (lRow == TBL_Error)
	{
		// Client-Rechteck ermitteln
		RECT r;
		if (!GetClientRect(m_hWndTbl, &r))
			return MTGR_RET_ERROR;

		// Setzen
		pr->left = tm.m_RowHeaderLeft;
		pr->top = r.top;
		pr->right = tm.m_RowHeaderRight;
		pr->bottom = r.bottom;

		iRet = MTGR_RET_VISIBLE;

	}
	else
	{
		BOOL bInclMerged = (dwFlags & MTGR_INCLUDE_MERGED && m_pCounter->GetMergeRows() > 0);
		int iRowRet;
		LPMTBLMERGEROW pmr = NULL;
		POINT pt, pt2, ptVis;

		if (bInclMerged)
			pmr = FindMergeRow(lRow, FMR_ANY);

		// Teil eines Zeilenverbundes?
		if (pmr)
		{
			// Y-Koordinaten der ersten Zeile ermitteln
			//iRowRet = MTblGetRowYCoord( m_hWndTbl, pmr->lRowFrom, &tm, &pt, &ptVis, FALSE );
			iRowRet = GetRowYCoord(pmr->lRowFrom, &tm, &pt, &ptVis, FALSE);
			if (iRowRet == MTGR_RET_ERROR) return iRowRet;

			// Y-Koordinaten der letzten Zeile ermitteln
			//iRowRet = MTblGetRowYCoord( m_hWndTbl, pmr->lRowTo, &tm, &pt2, &ptVis, FALSE );
			iRowRet = GetRowYCoord(pmr->lRowTo, &tm, &pt2, &ptVis, FALSE);
			if (iRowRet == MTGR_RET_ERROR) return iRowRet;

			// Oben/Unten Zeilenbereich ermitteln
			long lRowsTop = (lRow < 0) ? tm.m_SplitRowsTop : tm.m_RowsTop;
			long lRowsBottom = (lRow < 0) ? tm.m_SplitRowsBottom : tm.m_RowsBottom;

			// Nicht sichtbar?
			if (pt.x >= lRowsBottom || pt2.y < lRowsTop)
				return MTGR_RET_INVISIBLE;
			else
			{
				// Nur teilweise sichtbar?
				if (pt.x < lRowsTop || pt2.y > lRowsBottom - 1)
					iRet = MTGR_RET_PARTLY_VISIBLE;
				else
					iRet = MTGR_RET_VISIBLE;

				// Oben/Unten setzen
				pr->top = max(pt.x, lRowsTop);
				pr->bottom = min(pt2.y, lRowsBottom - 1);
			}
		}
		else
		{
			// Y-Koordinaten der Zeile ermitteln
			//iRowRet = MTblGetRowYCoord( m_hWndTbl, lRow, &tm, &pt, &ptVis, TRUE );
			iRowRet = GetRowYCoord(lRow, &tm, &pt, &ptVis, TRUE);
			if (iRowRet == MTGR_RET_ERROR || iRowRet == MTGR_RET_INVISIBLE)
				return iRowRet;
			else
				iRet = iRowRet;

			// Oben/Unten setzen
			pr->top = pt.x;
			pr->bottom = pt.y;
		}

		// Links/Rechts setzen
		pr->left = tm.m_RowHeaderLeft;
		pr->right = tm.m_RowHeaderRight;
	}

	return iRet;
}

// GetRowNrFocus
long MTBLSUBCLASS::GetRowNrFocus()
{
	if (m_pRowFocus)
		return m_pRowFocus->GetNr();

	return TBL_Error;
}


// GetRowRect
int MTBLSUBCLASS::GetRowRect(long lRow, LPRECT pr, DWORD dwFlags /* = 0 */)
{
	// Tabellenmaße
	CMTblMetrics tm(m_hWndTbl);

	BOOL bInclMerged = (dwFlags & MTGR_INCLUDE_MERGED && (m_pCounter->GetMergeCells() > 0 || m_pCounter->GetMergeRows() > 0));
	BOOL bToEndOfArea = (dwFlags & MTGR_TO_END_OF_AREA);
	POINT pt, ptVis;
	int iRet;

	if (bToEndOfArea)
	{
		if (lRow >= 0 && lRow < tm.m_MinVisibleRow)
		{
			pr->left = tm.m_ClientRect.left;
			pr->top = tm.m_RowsTop;
			pr->right = tm.m_ClientRect.right;
			pr->bottom = tm.m_RowsBottom;

			return MTGR_RET_PARTLY_VISIBLE;
		}
	}

	// Y-Koordinaten der Zeile ermitteln
	//iRet = MTblGetRowYCoord( m_hWndTbl, lRow, &tm, &pt, &ptVis );
	iRet = GetRowYCoord(lRow, &tm, &pt, &ptVis);
	if (!bInclMerged && iRet <= MTGR_RET_INVISIBLE) return iRet;

	pr->left = tm.m_ClientRect.left;
	pr->top = ptVis.x;
	pr->right = tm.m_ClientRect.right;
	pr->bottom = ptVis.y;

	if (bInclMerged)
	{
		iRet = MTGR_RET_INVISIBLE;

		if (m_pCounter->GetMergeCells() > 0)
		{
			BOOL bOneVisFound = FALSE;
			HWND hWndCol = NULL;
			int iBottom, iLeft, iRetCell, iRight, iTop;
			//while ( MTblFindNextCol( m_hWndTbl, &hWndCol, COL_Visible, 0 ) )
			while (FindNextCol(&hWndCol, COL_Visible, 0))
			{
				iRetCell = MTblGetCellRectEx(hWndCol, lRow, &iLeft, &iTop, &iRight, &iBottom, MTGR_VISIBLE_PART | MTGR_INCLUDE_MERGED);
				if (iRetCell > MTGR_RET_INVISIBLE)
				{
					if (iTop < pr->top)
						pr->top = iTop;

					if (iBottom > pr->bottom)
						pr->bottom = iBottom;

					bOneVisFound = TRUE;

					if (iRet == MTGR_RET_INVISIBLE || (iRet == MTGR_RET_VISIBLE && iRetCell == MTGR_RET_PARTLY_VISIBLE))
						iRet = iRetCell;
				}
				else if (iRetCell == MTGR_RET_INVISIBLE && bOneVisFound)
				{
					iRet = MTGR_RET_PARTLY_VISIBLE;
					break;
				}
				else if (iRetCell == MTGR_RET_ERROR)
				{
					iRet = MTGR_RET_ERROR;
					break;
				}
			}
		}

		if (m_pCounter->GetMergeRows() > 0 && iRet != MTGR_RET_ERROR)
		{
			LPMTBLMERGEROW pmr = FindMergeRow(lRow, FMR_ANY);
			if (pmr)
			{
				BOOL bOneVisFound = FALSE;
				int iRetRow;
				long lRow = pmr->lRowFrom;
				vector<long> vRows;
				vector<long>::iterator it;

				vRows.push_back(lRow);
				if (m_ppd->ppmrh)
					long lRow = m_ppd->lEffPaintRow;
				while (FindNextRow(&lRow, 0, ROW_Hidden, pmr->lRowTo))
					vRows.push_back(lRow);

				for (it = vRows.begin(); it != vRows.end(); it++)
				{
					//iRetRow = MTblGetRowYCoord( m_hWndTbl, *it, &tm, &pt, &ptVis );
					iRetRow = GetRowYCoord(*it, &tm, &pt, &ptVis);
					if (iRetRow > MTGR_RET_INVISIBLE)
					{
						if (ptVis.x < pr->top)
							pr->top = ptVis.x;

						if (ptVis.y > pr->bottom)
							pr->bottom = ptVis.y;

						bOneVisFound = TRUE;

						if (iRet == MTGR_RET_INVISIBLE || (iRet == MTGR_RET_VISIBLE && iRetRow == MTGR_RET_PARTLY_VISIBLE))
							iRet = iRetRow;
					}
					else if (iRetRow == MTGR_RET_INVISIBLE && bOneVisFound)
					{
						iRet = MTGR_RET_PARTLY_VISIBLE;
						break;
					}
					else if (iRetRow == MTGR_RET_ERROR)
					{
						iRet = MTGR_RET_ERROR;
						break;
					}
				}
			}
		}
	}

	if (bToEndOfArea && iRet > MTGR_RET_INVISIBLE)
	{
		long lRowsBottom = lRow < 0 ? tm.m_SplitRowsBottom : tm.m_RowsBottom;
		pr->bottom = max(pr->bottom, (lRowsBottom - 1));
	}

	return iRet;
}


// GetScrollBarV
BOOL MTBLSUBCLASS::GetScrollBarV(CMTblMetrics * ptm, BOOL bSplit, LPHWND phWndScr, LPINT piScrCtl)
{
	if (!ptm || !phWndScr || !piScrCtl)
		return FALSE;

	int iScrCtl;
	HWND hWndScr = NULL;

	// Tabelle geteilt?
	if (ptm->m_SplitBarTop > 0)
	{
		TCHAR szBuf[255];
		RECT r;
		// ScrollBars suchen
		HWND hWnd = GetWindow(m_hWndTbl, GW_CHILD);
		while (hWnd)
		{
			if (!GetClassName(hWnd, szBuf, 254)) return -1;
			if (lstrcmp(szBuf, _T("ScrollBar")) == 0)
			{
				if (!GetWindowRect(hWnd, &r)) return -1;
				if (!MapWindowPoints(NULL, m_hWndTbl, (LPPOINT)&r, 2)) return -1;
				if (bSplit ? r.top >= ptm->m_SplitBarTop : r.top < ptm->m_SplitBarTop)
				{
					hWndScr = hWnd;
					iScrCtl = SB_CTL;
					break;
				}
			}
			hWnd = GetWindow(hWnd, GW_HWNDNEXT);
		}
	}
	else
	{
		if (bSplit)
			return FALSE;
		iScrCtl = SB_VERT;
		hWndScr = m_hWndTbl;
	}

	if (!hWndScr)
		return FALSE;

	*phWndScr = hWndScr;
	*piScrCtl = iScrCtl;
	return TRUE;
}


// GetScrollPosV
int MTBLSUBCLASS::GetScrollPosV(CMTblMetrics * ptm, BOOL bSplit)
{
	int iScrCtl;
	HWND hWndScr = NULL;

	if (!GetScrollBarV(ptm, bSplit, &hWndScr, &iScrCtl))
		return -1;

	SCROLLINFO si;
	si.cbSize = sizeof(SCROLLINFO);
	//si.fMask = SIF_POS;
	si.fMask = SIF_ALL;
	if (!FlatSB_GetScrollInfo(hWndScr, iScrCtl, &si))
	{
		DWORD dwError = GetLastError();
		if (dwError == ERROR_NO_SCROLLBARS)
			return 0;
		else
			return -1;
	}

	if (bSplit)
		return si.nPos - TBL_MinSplitRow;
	else
		return si.nPos;
}


// GetRowSel
long MTBLSUBCLASS::GetRowSel(int iMode, long lRow /*= TBL_Error*/, RowPtrVector *pvRows /* = NULL*/)
{
	BOOL bRemoveNew = FALSE, bSet;
	CMTblRow *pRow;
	DWORD dwFlag = ((iMode & GETSEL_BEFORE) ? ROW_IFLAG_SELECTED_BEFORE : ROW_IFLAG_SELECTED_AFTER);
	long lCount = 0;
	LPMTBLMERGEROW pmr;
	RowPtrList liRows;
	RowPtrList::iterator it;

	if (iMode & GETSEL_AFTER)
		bRemoveNew = (iMode & GETSEL_REMOVE_NEW);

	if (pvRows)
		pvRows->clear();

	// Ggf. Selektionen innerhalb von Zeilenverbünden replizieren
	if ((iMode & GETSEL_AFTER) && m_pCounter->GetMergeRows() > 0)
	{
		// Zeilenverbünde ermitteln
		if (GetMergeRows())
		{
			// Single-Selection?
			BOOL bSingleSelection = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection);

			// Zellenverbünde durchlaufen
			BOOL bSetSel;
			/*MTBLMERGEROW_VECTOR::size_type i, iMax = m_pMergeRows->mrv.size( );
			for ( i = 0; i < iMax; ++i )*/
			MTBLMERGEROW_MAP::iterator it, itEnd = m_pMergeRows->mrm.end();
			for (it = m_pMergeRows->mrm.begin(); it != itEnd; ++it)
			{
				// Zellenverbund ermitteln
				//pmr = m_pMergeRows->mrv[i];
				pmr = (*it).second;

				// Ermitteln, ob Selektion gesetzt oder entfernt werden muss
				bSetSel = AnyMergeRowSelected(pmr);
				// Zeile Teil des Verbundes?
				if (lRow != TBL_Error && lRow >= pmr->lRowFrom && lRow <= pmr->lRowTo)
					bSetSel = Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Selected);
				// Top-Down-Mausselektion?
				else if (m_iRowSelByMouse == 1 && !bSingleSelection)
					bSetSel = Ctd.TblQueryRowFlags(m_hWndTbl, pmr->lRowFrom, ROW_Selected);
				// Bottom-Up-Mausselektion?
				else if (m_iRowSelByMouse == 2 && !bSingleSelection)
					bSetSel = Ctd.TblQueryRowFlags(m_hWndTbl, pmr->lRowTo, ROW_Selected);

				// ROW_Selected replizieren
				SetMergeRowFlags(pmr, ROW_Selected, bSetSel, FALSE);
			}
		}
	}

	// Internes Selektionsflag in allen Zeilen zurücksetzen
	/*m_Rows->SetRowsInternalFlags( FALSE, dwFlag, FALSE );
	m_Rows->SetRowsInternalFlags( TRUE, dwFlag, FALSE );*/

	if (lRow == TBL_Error)
	{
		// Internes Selektionsflag in allen Zeilen zurücksetzen
		m_Rows->SetRowsInternalFlags(FALSE, dwFlag, FALSE);
		m_Rows->SetRowsInternalFlags(TRUE, dwFlag, FALSE);

		// Alle selektierten Zeilen ermitteln
		lRow = TBL_MinRow;
		while (Ctd.TblFindNextRow(m_hWndTbl, &lRow, ROW_Selected, 0))
		{
			pRow = m_Rows->EnsureValid(lRow);
			if (!pRow)
				return -1;

			liRows.push_back(pRow);
		}
	}
	else
	{
		/*if ( Ctd.TblQueryRowFlags( m_hWndTbl, lRow, ROW_Selected ) )
		{
		pmr = FindMergeRow( lRow, FMR_ANY );
		if ( pmr )
		{
		for ( long l = pmr->lRowFrom; l<= pmr->lRowTo; l++ )
		{
		if ( !Ctd.TblQueryRowFlags( m_hWndTbl, l, ROW_Hidden ) )
		{
		pRow = m_Rows->EnsureValid( l );
		if ( !pRow )
		return -1;

		liRows.push_back( pRow );
		}
		}
		}
		else
		{
		pRow = m_Rows->EnsureValid( lRow );
		if ( !pRow )
		return -1;

		liRows.push_back( pRow );
		}
		}*/

		BOOL bSelected = Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Selected);

		pmr = FindMergeRow(lRow, FMR_ANY);
		if (pmr)
		{
			for (long l = pmr->lRowFrom; l <= pmr->lRowTo; l++)
			{
				pRow = m_Rows->EnsureValid(l);
				if (!pRow)
					return -1;

				pRow->SetInternalFlags(dwFlag, FALSE);

				if (bSelected && !Ctd.TblQueryRowFlags(m_hWndTbl, l, ROW_Hidden))
					liRows.push_back(pRow);
			}
		}
		else
		{
			pRow = m_Rows->EnsureValid(lRow);
			if (!pRow)
				return -1;

			pRow->SetInternalFlags(dwFlag, FALSE);

			if (bSelected)
				liRows.push_back(pRow);
		}
	}

	// Internes Selektionsflag bei markierten Zeilen setzen
	if (pvRows)
		pvRows->reserve(liRows.size());

	for (it = liRows.begin(); it != liRows.end(); it++)
	{
		pRow = *it;
		lRow = pRow->GetNr();

		bSet = TRUE;

		// Neue Selektion entfernen?
		if (bRemoveNew)
		{
			if (!pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE))
			{
				Ctd.TblSetRowFlags(m_hWndTbl, lRow, ROW_Selected, FALSE);
				bSet = FALSE;
			}
		}

		if (bSet)
		{
			pRow->SetInternalFlags(dwFlag, TRUE);

			if (pvRows)
				pvRows->push_back(pRow);

			++lCount;
		}
	}

	return lCount;
}


// GetRowSelChangesRgn
// Ermittelt die Region der Zeilen, bei denen sich die Selektion geändert hat
HRGN MTBLSUBCLASS::GetRowSelChangesRgn(long lRow /*= TBL_Error*/)
{
	BOOL bAny = FALSE, bSelAfter, bSelBefore;
	CMTblRow *pRow;
	int iRet;
	RECT r;

	HRGN hRgn = CreateRectRgn(0, 0, 0, 0);
	if (!hRgn)
		return NULL;

	// Single-Selection?
	BOOL bSingleSelection = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection);

	if (lRow == TBL_Error)
	{
		long lRowFrom, lRowTo;

		// Normale Zeilen
		if (QueryVisRange(lRowFrom, lRowTo))
		{
			Ctd.TblFindNextRow(m_hWndTbl, &lRowTo, 0, ROW_Hidden);

			long lRow = /*-1*/lRowFrom - 1;
			//while ( pRow = m_Rows->FindNextRow( &lRow ) )
			while ((pRow = m_Rows->FindNextRow(&lRow)) && lRow <= lRowTo)
			{
				bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
				bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
				if (bSelBefore != bSelAfter)
				{
					iRet = GetRowRect(lRow, &r, bSingleSelection ? MTGR_INCLUDE_MERGED : 0);
					if (iRet != MTGR_RET_ERROR && iRet != MTGR_RET_INVISIBLE)
					{
						r.right += 1;
						r.bottom += 1;
						HRGN hRgnRow = CreateRectRgnIndirect(&r);
						if (hRgnRow)
						{
							iRet = CombineRgn(hRgn, hRgn, hRgnRow, RGN_OR);
							DeleteObject(hRgnRow);
							bAny = TRUE;
						}
					}
				}
			}
		}

		// Split-Zeilen
		if (QueryVisRange(lRowFrom, lRowTo, TRUE))
		{
			Ctd.TblFindNextRow(m_hWndTbl, &lRowTo, 0, ROW_Hidden);

			lRow = /*TBL_MinRow*/lRowFrom - 1;
			//while ( pRow = m_Rows->FindNextSplitRow( &lRow ) )
			while ((pRow = m_Rows->FindNextSplitRow(&lRow)) && lRow <= lRowTo)
			{
				bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
				bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
				if (bSelBefore != bSelAfter)
				{
					iRet = GetRowRect(lRow, &r, bSingleSelection ? MTGR_INCLUDE_MERGED : 0);
					if (iRet != MTGR_RET_ERROR && iRet != MTGR_RET_INVISIBLE)
					{
						r.right += 1;
						r.bottom += 1;
						HRGN hRgnRow = CreateRectRgnIndirect(&r);
						if (hRgnRow)
						{
							iRet = CombineRgn(hRgn, hRgn, hRgnRow, RGN_OR);
							DeleteObject(hRgnRow);
							bAny = TRUE;
						}
					}
				}
			}
		}
	}
	else
	{
		pRow = m_Rows->GetRowPtr(lRow);
		if (pRow)
		{
			bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
			bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
			if (bSelBefore != bSelAfter)
			{
				iRet = GetRowRect(lRow, &r, MTGR_INCLUDE_MERGED);
				if (iRet != MTGR_RET_ERROR && iRet != MTGR_RET_INVISIBLE)
				{
					r.right += 1;
					r.bottom += 1;
					HRGN hRgnRow = CreateRectRgnIndirect(&r);
					if (hRgnRow)
					{
						iRet = CombineRgn(hRgn, hRgn, hRgnRow, RGN_OR);
						DeleteObject(hRgnRow);
						bAny = TRUE;
					}
				}
			}
		}
	}

	if (bAny)
		return hRgn;
	else
	{
		DeleteObject(hRgn);
		return NULL;
	}
}


// GetRowYCoord
int MTBLSUBCLASS::GetRowYCoord(long lRow, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis, BOOL bInVisRangeOnly /*= TRUE*/)
{
	// Gültige Zeile?
	if (!IsRow(lRow))
		return MTGR_RET_ERROR;

	// Tabellenmaße ok?
	CMTblMetrics tm;
	if (!ptm)
	{
		tm.Get(m_hWndTbl);
		ptm = &tm;
	}

	// Zeile sichtbar?
	if (Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Hidden))
		return MTGR_RET_INVISIBLE;

	// Split-Zeile?
	BOOL bSplitRow = (lRow < 0);

	if (bInVisRangeOnly && !bSplitRow)
	{
		if (lRow < ptm->m_MinVisibleRow || lRow > ptm->m_MaxVisibleRow)
			return MTGR_RET_INVISIBLE;
	}

	// Auf geht's
	int iScrPosV = GetScrollPosV(ptm, bSplitRow);
	if (iScrPosV < 0)
		return MTGR_RET_ERROR;

	long lRowVisPos = m_Rows->GetRowVisPos(lRow);
	if (lRowVisPos == TBL_Error)
		return MTGR_RET_ERROR;

	POINT pt;
	long lRowsTop = bSplitRow ? ptm->m_SplitRowsTop : ptm->m_RowsTop;
	long lRowsBottom = bSplitRow ? ptm->m_SplitRowsBottom : ptm->m_RowsBottom;

	if (m_pCounter->GetRowHeights() <= 0) // Keine individuellen Zeilenhöhen
	{
		pt.x = lRowsTop + ((lRowVisPos - iScrPosV) * ptm->m_RowHeight);
		pt.y = pt.x + ptm->m_RowHeight - 1;
	}
	else // Mind. eine Zeile mit individueller Höhe
	{
		if (!bSplitRow && lRow == ptm->m_MinVisibleRow)
			pt.x = lRowsTop;
		else
		{
			// Zeile ermitteln, deren sichtbare Position der Scrollposition entspricht ( = oberste Zeile im sichtbaren Bereich )
			long lRowScrPos = m_Rows->GetVisPosRow(iScrPosV, bSplitRow);
			if (lRowScrPos == TBL_Error)
				return MTGR_RET_ERROR;

			// Logische X-Position der obersten sichtbaren Zeile ermitteln
			long lxRowTop = m_Rows->GetRowLogX(lRowScrPos, ptm->m_RowHeight);

			// Logische X-Position der Zeile ermitteln
			long lxRow = m_Rows->GetRowLogX(lRow, ptm->m_RowHeight);

			// Physische X-Position berechnen
			pt.x = lRowsTop + (lxRow - lxRowTop);
		}

		// Position Y
		pt.y = pt.x + GetEffRowHeight(lRow, ptm) - 1;
	}

	// Überhaupt ein Teil sichtbar?
	BOOL bInvisible = (pt.x >= lRowsBottom || pt.y < lRowsTop);
	if (bInVisRangeOnly && bInvisible)
		return MTGR_RET_INVISIBLE;

	// Bereich setzen
	*ppt = pt;

	// Sichtbaren Bereich setzen
	if (!bInvisible)
	{
		pptVis->x = max(ppt->x, lRowsTop);
		pptVis->y = min(ppt->y, lRowsBottom - 1);

		// Nur teilweise sichtbar?
		if (pptVis->x != ppt->x || pptVis->y != ppt->y)
			return MTGR_RET_PARTLY_VISIBLE;
	}

	return bInvisible ? MTGR_RET_INVISIBLE : MTGR_RET_VISIBLE;
}

// GetRowSelChanges
// Liefert die Änderungen der Zeilenselektion
long MTBLSUBCLASS::GetRowSelChanges(HARRAY haRows, HARRAY haChanges)
{
	// Check Arrays
	long lDims;
	if (!Ctd.ArrayDimCount(haRows, &lDims)) return -1;
	if (lDims > 1) return -1;

	if (!Ctd.ArrayDimCount(haChanges, &lDims)) return -1;
	if (lDims > 1) return -1;

	// Untergrenzen der Arrays ermitteln
	long lLowRows;
	if (!Ctd.ArrayGetLowerBound(haRows, 1, &lLowRows)) return -1;

	long lLowChanges;
	if (!Ctd.ArrayGetLowerBound(haChanges, 1, &lLowChanges)) return -1;

	// Änderungen ermitteln
	BOOL bSel, bSelAfter, bSelBefore;
	CMTblRow *pRow;
	long lCount = 0;
	NUMBER n;

	// Normale Zeilen
	long lRow = -1;
	while (pRow = m_Rows->FindNextRow(&lRow))
	{
		bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
		bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
		if (bSelBefore != bSelAfter)
		{
			if (!Ctd.CvtLongToNumber(lRow, &n))
				return -1;
			if (!Ctd.MDArrayPutNumber(haRows, &n, lCount + lLowRows))
				return -1;

			if (bSelBefore && !bSelAfter)
				bSel = FALSE;
			else
				bSel = TRUE;

			if (!Ctd.CvtLongToNumber(bSel, &n))
				return -1;
			if (!Ctd.MDArrayPutNumber(haChanges, &n, lCount + lLowChanges))
				return -1;

			++lCount;
		}
	}

	// Split-Zeilen
	lRow = TBL_MinRow;
	while (pRow = m_Rows->FindNextSplitRow(&lRow))
	{
		bSelBefore = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_BEFORE) != 0;
		bSelAfter = pRow->QueryInternalFlags(ROW_IFLAG_SELECTED_AFTER) != 0;
		if (bSelBefore != bSelAfter)
		{
			if (!Ctd.CvtLongToNumber(lRow, &n))
				return -1;
			if (!Ctd.MDArrayPutNumber(haRows, &n, lCount + lLowRows))
				return -1;

			if (bSelBefore && !bSelAfter)
				bSel = FALSE;
			else
				bSel = TRUE;

			if (!Ctd.CvtLongToNumber(bSel, &n))
				return -1;
			if (!Ctd.MDArrayPutNumber(haChanges, &n, lCount + lLowChanges))
				return -1;

			++lCount;
		}
	}

	return lCount;
}


// GetScreenMetrics
BOOL MTBLSUBCLASS::GetScreenMetrics(int &iLeft, int &iTop, int &iWid, int &iHt)
{
	if (WinVerIs2000_OG() || WinVerIs98_OG())
	{
		iLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
		iTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
		iWid = GetSystemMetrics(SM_CXVIRTUALSCREEN);
		iHt = GetSystemMetrics(SM_CYVIRTUALSCREEN);
	}
	else
	{
		iLeft = 0;
		iTop = 0;
		iWid = GetSystemMetrics(SM_CXSCREEN);
		iHt = GetSystemMetrics(SM_CYSCREEN);
	}

	return TRUE;
}

// GetScreenRectFromPoint
BOOL MTBLSUBCLASS::GetScreenRectFromPoint(POINT &pt, RECT &rMetrics)
{
	BOOL bOk = FALSE;

	HMONITOR hMonitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONULL);
	if (hMonitor)
	{
		MONITORINFO mi;
		mi.cbSize = sizeof(MONITORINFO);
		if (GetMonitorInfo(hMonitor, &mi))
		{
			CopyRect(&rMetrics, &mi.rcMonitor);
			bOk = TRUE;
		}
	}

	return bOk;
}


// GetSepCol
HWND MTBLSUBCLASS::GetSepCol(int iX, int iY)
{
	HWND hWndCol = NULL, hWndCol1;

	hWndCol1 = MTblGetColHdr(m_hWndTbl, iX - 4, iY);
	if (hWndCol1)
	{
		CMTblMetrics tm(m_hWndTbl);
		POINT ptc, ptcVis;

		GetColXCoord(hWndCol1, &tm, &ptc, &ptcVis);
		if (iX >= ptc.y - 3 && iX <= ptc.y + 3)
			hWndCol = hWndCol1;
	}

	return hWndCol;
}


// GetSepRow
long MTBLSUBCLASS::GetSepRow(POINT pt)
{
	DWORD dwFlags;
	HWND hWndCol;
	long lRow1 = TBL_Error, lRow = TBL_Error;

	for (int iOffs = 0; iOffs <= 4; ++iOffs)
	{
		Ctd.TblObjectsFromPoint(m_hWndTbl, pt.x, pt.y - iOffs, &lRow1, &hWndCol, &dwFlags);
		if (lRow1 != TBL_Error && dwFlags & TBL_XOverRowHeader && !(dwFlags & TBL_YOverColumnHeader))
		{
			CMTblMetrics tm(m_hWndTbl);
			POINT ptc, ptcVis;
			GetRowYCoord(lRow1, &tm, &ptc, &ptcVis);
			if (pt.y >= ptc.y - 2 && pt.y <= ptc.y + 4)
			{
				DWORD dwFlags2;
				long lRow2;
				BOOL bOk = Ctd.TblObjectsFromPoint(m_hWndTbl, pt.x, pt.y, &lRow2, &hWndCol, &dwFlags2);
				if (!(bOk && dwFlags2 & TBL_YOverSplitBar) && !(bOk && dwFlags2 & TBL_YOverSplitRows && lRow1 >= 0))
				{
					lRow = lRow1;
					break;
				}
			}
		}
	}

	return lRow;
}

// GetSystemTimeI64
__int64 MTBLSUBCLASS::GetSystemTimeI64()
{
	FILETIME ft;
	GetSystemTimeAsFileTime(&ft);
	LARGE_INTEGER liTime;
	liTime.LowPart = ft.dwLowDateTime;
	liTime.HighPart = ft.dwHighDateTime;
	return liTime.QuadPart;
}

// GetTip
CMTblTip * MTBLSUBCLASS::GetTip(long lParam, HWND hWndParam, int iType, BOOL bEnsureValid /*= FALSE*/)
{
	CMTblTip *pTip = NULL;
	CMTblCol *pCol = NULL;
	CMTblRow *pRow = NULL;

	if (iType == MTBL_TIP_DEFAULT)
	{
		pTip = m_DefTip;
	}
	else if (iType == MTBL_TIP_CELL ||
		iType == MTBL_TIP_CELL_TEXT ||
		iType == MTBL_TIP_CELL_COMPLETETEXT ||
		iType == MTBL_TIP_CELL_IMAGE ||
		iType == MTBL_TIP_CELL_BTN)
	{
		if (bEnsureValid)
			pRow = m_Rows->EnsureValid(lParam);
		else
			pRow = m_Rows->GetRowPtr(lParam);

		if (pRow)
		{
			if (bEnsureValid)
				pCol = pRow->m_Cols->EnsureValid(hWndParam);
			else
				pCol = pRow->m_Cols->Find(hWndParam);
		}

		if (pCol)
		{
			switch (iType)
			{
			case MTBL_TIP_CELL:
				if (bEnsureValid && !pCol->m_pTip)
					pCol->m_pTip = new CMTblTip;

				pTip = pCol->m_pTip;
				break;
			case MTBL_TIP_CELL_TEXT:
				if (bEnsureValid && !pCol->m_pTip_Text)
					pCol->m_pTip_Text = new CMTblTip;

				pTip = pCol->m_pTip_Text;
				break;
			case MTBL_TIP_CELL_COMPLETETEXT:
				if (bEnsureValid && !pCol->m_pTip_CompleteText)
					pCol->m_pTip_CompleteText = new CMTblTip;

				pTip = pCol->m_pTip_CompleteText;
				break;
			case MTBL_TIP_CELL_IMAGE:
				if (bEnsureValid && !pCol->m_pTip_Image)
					pCol->m_pTip_Image = new CMTblTip;

				pTip = pCol->m_pTip_Image;
				break;
			case MTBL_TIP_CELL_BTN:
				if (bEnsureValid && !pCol->m_pTip_Button)
					pCol->m_pTip_Button = new CMTblTip;

				pTip = pCol->m_pTip_Button;
				break;
			}
		}
	}
	else if (iType == MTBL_TIP_COLHDR ||
		iType == MTBL_TIP_COLHDR_BTN)
	{
		if (bEnsureValid)
			pCol = m_Cols->EnsureValidHdr(hWndParam);
		else
			pCol = m_Cols->FindHdr(hWndParam);

		if (pCol)
		{
			switch (iType)
			{
			case MTBL_TIP_COLHDR:
				if (bEnsureValid && !pCol->m_pTip)
					pCol->m_pTip = new CMTblTip;
				pTip = pCol->m_pTip;
				break;
			case MTBL_TIP_COLHDR_BTN:
				if (bEnsureValid && !pCol->m_pTip_Button)
					pCol->m_pTip_Button = new CMTblTip;

				pTip = pCol->m_pTip_Button;
				break;
			}
		}
	}
	else if (iType == MTBL_TIP_COLHDRGRP)
	{
		CMTblColHdrGrp *pColHdrGrp = (CMTblColHdrGrp*)lParam;
		if (m_ColHdrGrps->IsValidGrp(pColHdrGrp))
		{
			if (bEnsureValid && !pColHdrGrp->m_pTip)
				pColHdrGrp->m_pTip = new CMTblTip;

			pTip = pColHdrGrp->m_pTip;
		}
	}
	else if (iType == MTBL_TIP_CORNER)
	{
		if (bEnsureValid && !m_Corner->m_pTip)
			m_Corner->m_pTip = new CMTblTip;

		pTip = m_Corner->m_pTip;
	}
	else if (iType == MTBL_TIP_ROWHDR)
	{
		if (bEnsureValid)
			pRow = m_Rows->EnsureValid(lParam);
		else
			pRow = m_Rows->GetRowPtr(lParam);

		if (pRow)
		{
			if (bEnsureValid && !pRow->m_HdrTip)
				pRow->m_HdrTip = new CMTblTip;

			pTip = pRow->m_HdrTip;
		}
	}

	if (pTip && pTip != m_DefTip)
		pTip->SetDefTip(m_DefTip);

	return pTip;
}


// GetTrailingCRLFCount
int MTBLSUBCLASS::GetTrailingCRLFCount(LPCTSTR lpsStr)
{
	int iCount = 0;

	if (lpsStr)
	{
		int iLen = lstrlen(lpsStr);
		if (iLen > 0)
		{
			int iPos = iLen - 2;
			while (iPos >= 0)
			{
				if (lpsStr[iPos] == 13 && lpsStr[iPos + 1] == 10)
					++iCount;
				else
					break;
				iPos -= 2;
			}
		}
	}

	return iCount;
}


// HandlePaint
LRESULT MTBLSUBCLASS::HandlePaint(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	BOOL bLog = FALSE;
	UINT start, stop;
	tstring sLogFile = _T("c:\\temp\\mtbl.log");

	PAINTSTRUCT ps;

	if (bLog)
		LogLine(sLogFile.c_str(), _T("Call to HandlePaint..."));

	if (!m_bPainting)
	{
		if (SendMessage(m_hWndTbl, MTM_Paint, wParam, lParam) == MTBL_NODEFAULTACTION)
			return 0;

		m_bPainting = TRUE;

		if (m_lLockPaint > 0)
		{
			if (bLog)
				LogLine(sLogFile.c_str(), _T("m_lLockPaint > 0"));
			BeginPaint(hWnd, &ps);
			ValidateRect(hWnd, &ps.rcPaint);
			EndPaint(hWnd, &ps);
		}
		else
		{
			GetUpdateRect(hWnd, &m_ppd->rUpd, FALSE);
			HDC hDC = BeginPaint(hWnd, &ps);
			if (hDC)
			{
				if (bLog)
				{
					start = GetTickCount();

					TCHAR sz[255];
					tstring sLine = _T("PaintTable, Upd-Rect=");
					_ltot(m_ppd->rUpd.left, sz, 10);
					sLine += sz;
					sLine += _T(",");
					_ltot(m_ppd->rUpd.top, sz, 10);
					sLine += sz;
					sLine += _T(",");
					_ltot(m_ppd->rUpd.right, sz, 10);
					sLine += sz;
					sLine += _T(",");
					_ltot(m_ppd->rUpd.bottom, sz, 10);
					sLine += sz;
					sLine += _T(",");

					LogLine(sLogFile.c_str(), sLine.c_str());
				}

				m_ppd->hDC = m_ppd->hDCTbl = hDC;
				PaintTable();

				if (bLog)
				{
					stop = GetTickCount();
					UINT ms = stop - start;

					TCHAR sz[255];
					tstring sLine = _T("PaintTable done in ");

					_stprintf(sz, _T("%d"), ms);
					sLine += sz;
					sLine += _T(" ms");

					LogLine(sLogFile.c_str(), sLine.c_str());
				}
			}
			EndPaint(hWnd, &ps);
		}

		m_bPainting = FALSE;
	}
	else
	{
		if (bLog)
			LogLine(sLogFile.c_str(), _T("m_bPainting = TRUE"));
		BeginPaint(hWnd, &ps);
		ValidateRect(hWnd, &ps.rcPaint);
		EndPaint(hWnd, &ps);
	}

	return 0;
}

// HookProc
LRESULT CALLBACK MTBLSUBCLASS::HookProcEditStyle(int iCode, WPARAM wParam, LPARAM lParam)
{
	if (iCode == HC_ACTION)
	{
		CWPSTRUCT *pcwp = (CWPSTRUCT*)lParam;
		if (!pcwp)
			goto end;

		if (pcwp->message == WM_NCCREATE)
		{
			CREATESTRUCT *pcs = (CREATESTRUCT*)pcwp->lParam;
			if (!pcs)
				goto end;

			TCHAR cBuf[255] = _T("");
			if (!GetClassName(pcwp->hwnd, cBuf, 255))
				goto end;

			if (!lstrcmpi(cBuf, _T("Edit")))
			{
				if (!GetClassName(pcs->hwndParent, cBuf, 255))
					goto end;

				if (!lstrcmp(cBuf, CLSNAME_LISTCLIP) || !lstrcmp(cBuf, CLSNAME_LISTCLIP_TBW))
				{
					HWND hWndTbl = GetParent(pcs->hwndParent);
					if ((void*)MTblGetSubClass(hWndTbl) == m_pscHookEditStyle)
					{
						long lStyle = GetWindowLongPtr(pcwp->hwnd, GWL_STYLE);
						long lNewStyle = lStyle;

						// Text-Ausrichtung
						if (m_iHookEditJustify != -1)
						{
							if (m_iHookEditJustify != DT_CENTER && (lStyle & ES_CENTER))
								lNewStyle ^= ES_CENTER;
							else if (m_iHookEditJustify == DT_CENTER && !(lStyle & ES_CENTER))
								lNewStyle |= ES_CENTER;

							if (m_iHookEditJustify != DT_RIGHT && (lStyle & ES_RIGHT))
								lNewStyle ^= ES_RIGHT;
							else if (m_iHookEditJustify == DT_RIGHT && !(lStyle & ES_RIGHT))
								lNewStyle |= ES_RIGHT;
						}

						// Multiline
						if (m_bHookEditMultiline)
						{
							/*if ( !(lStyle & ES_MULTILINE) )
							lNewStyle |= ES_MULTILINE;*/
							if (!(lStyle & ES_AUTOVSCROLL))
								lNewStyle |= ES_AUTOVSCROLL;
							if (lStyle & ES_AUTOHSCROLL)
								lNewStyle ^= ES_AUTOHSCROLL;
						}
						else
						{
							/*if ( lStyle & ES_MULTILINE )
							lNewStyle ^= ES_MULTILINE;*/
							if (lStyle & ES_AUTOVSCROLL)
								lNewStyle ^= ES_AUTOVSCROLL;
							if (!(lStyle & ES_AUTOHSCROLL))
								lNewStyle |= ES_AUTOHSCROLL;
						}

						if (lNewStyle != lStyle)
							SetWindowLongPtr(pcwp->hwnd, GWL_STYLE, lNewStyle);
					}
				}
			}
		}
	}

	end:
	return CallNextHookEx(m_hHookEditStyle, iCode, wParam, lParam);
}


// IsContainerSubClassed
BOOL MTBLSUBCLASS::IsContainerSubClassed(HWND hWnd)
{
	return (GetProp(hWnd, MTBL_CONTAINER_SUBCLASS_PROP) != NULL);
}


// IsDropDownBtn
BOOL MTBLSUBCLASS::IsDropDownBtn(HWND hWnd)
{
	TCHAR szClsName[255];
	return (GetClassName(hWnd, szClsName, (sizeof(szClsName) / sizeof(TCHAR)) - 1) && !lstrcmp(szClsName, CLSNAME_DROPDOWNBTN));
}


// IsEnabledHyperlink
BOOL MTBLSUBCLASS::IsEnabledHyperlinkCell(HWND hWndCol, long lRow)
{
	CMTblRow * pRow = m_Rows->GetRowPtr(lRow);
	if (!pRow) return FALSE;

	CMTblHyLnk * pHyLnk = pRow->m_Cols->GetHyLnk(hWndCol);
	if (!pHyLnk) return FALSE;

	return pHyLnk->IsEnabled();
}


// IsFocusCell
BOOL MTBLSUBCLASS::IsFocusCell(HWND hWndCol, long lRow)
{
	HWND hWndColFocus = NULL;
	long lRowFocus = TBL_Error;

	Ctd.TblQueryFocus(m_hWndTbl, &lRowFocus, &hWndColFocus);
	BOOL bIs = (hWndColFocus && lRowFocus != TBL_Error && hWndColFocus == hWndCol && lRowFocus == lRow);

	return bIs;
}


// IsGroupedColHdr
BOOL MTBLSUBCLASS::IsGroupedColHdr(HWND hWndCol)
{
	return m_ColHdrGrps->GetGrpOfCol(hWndCol) != NULL;
}


// IsIndented
BOOL MTBLSUBCLASS::IsIndented(HWND hWndCol, long lRow)
{
	/*return ( m_hWndTreeCol != NULL ) &&
	( hWndCol == m_hWndTreeCol ) &&
	( lRow != TBL_Error ) &&
	(
	m_Rows->QueryRowFlags( lRow, MTBL_ROW_CANEXPAND ) ||
	m_Rows->GetParentRow( lRow ) != TBL_Error ||
	( m_Rows->GetRowLevel( lRow ) == 0 && m_dwTreeFlags & MTBL_TREE_FLAG_INDENT_ALL )
	);*/
	return (GetEffCellIndent(lRow, hWndCol) != 0);
}


// IsInEditMode
BOOL MTBLSUBCLASS::IsInEditMode()
{
	HWND hWndCol = NULL;
	long lRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol);
	return hWndCol != NULL;
}


// IsListClipSubClassed
BOOL MTBLSUBCLASS::IsListClipSubClassed(HWND hWnd)
{
	return (GetProp(hWnd, MTBL_LISTCLIP_SUBCLASS_PROP) != NULL);
}


// IsLockedColumn
BOOL MTBLSUBCLASS::IsLockedColumn(HWND hWndCol)
{
	int iLocked = MTblQueryLockedColumns(m_hWndTbl);

	int iPos = Ctd.TblQueryColumnPos(hWndCol);
	if (iPos < 1) return FALSE;

	return (iPos <= iLocked);
}


// IsMergeCell
BOOL MTBLSUBCLASS::IsMergeCell(HWND hWndCol, long lRow)
{
	return FindMergeCell(hWndCol, lRow, FMC_MASTER) != NULL;
}


// IsMergedCell
BOOL MTBLSUBCLASS::IsMergedCell(HWND hWndCol, long lRow)
{

	return FindMergeCell(hWndCol, lRow, FMC_SLAVE) != NULL;
}


// IsMyWindow
BOOL MTBLSUBCLASS::IsMyWindow(HWND hWnd)
{
	if (!hWnd)
		return FALSE;

	if (hWnd == m_hWndTbl)
		return TRUE;

	HWND hWndParent = hWnd;
	while (hWndParent)
	{
		hWndParent = GetParent(hWndParent);
		if (hWndParent == m_hWndTbl)
			return TRUE;
	}

	HWND hWndCol = NULL;
	long lRow = TBL_Error;
	if (Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol) && hWndCol)
	{
		//int iCellType = Ctd.TblGetCellType( hWndCol );
		int iCellType = GetEffCellType(lRow, hWndCol);
		if (iCellType == COL_CellType_PopupEdit || iCellType == COL_CellType_DropDownList)
		{
			if (hWnd == (HWND)SendMessage(m_hWndTbl, TIM_GetExtEdit, Ctd.TblQueryColumnID(hWndCol) - 1, 0))
				return TRUE;
		}
	}

	return FALSE;
}


// IsListClipCtrlSubClassed
BOOL MTBLSUBCLASS::IsListClipCtrlSubClassed(HWND hWnd)
{
	return (GetProp(hWnd, MTBL_LISTCLIP_CTRL_SUBCLASS_PROP) != NULL);
}


// IsOverCornerSep
BOOL MTBLSUBCLASS::IsOverCornerSep(int iX, int iY)
{
	BOOL bIs = FALSE;
	DWORD dwFlags = 0;
	HWND hWndCol = NULL;
	int iToLeft = 4, iToRight = 3;
	long lRow;

	Ctd.TblObjectsFromPoint(m_hWndTbl, iX - iToLeft, iY, &lRow, &hWndCol, &dwFlags);
	if (dwFlags & TBL_XOverRowHeader && dwFlags & TBL_YOverColumnHeader && !hWndCol)
	{
		hWndCol = MTblGetColHdr(m_hWndTbl, iX + iToRight, iY);
		if (hWndCol)
		{
			bIs = TRUE;
		}
	}

	return bIs;
}


// IsOverHyperlink
BOOL MTBLSUBCLASS::IsOverHyperlink(HWND hWndCol, long lRow, DWORD dwMFlags)
{
	HWND hWndCheckCol = hWndCol, hWndMergeCol;
	long lCheckRow = lRow, lMergeRow;
	if ((dwMFlags & MTOFP_OVERMERGEDCELL) && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
	{
		hWndCheckCol = hWndMergeCol;
		lCheckRow = lMergeRow;
	}
	return ((dwMFlags & MTOFP_OVERTEXT) &&
		IsEnabledHyperlinkCell(hWndCheckCol, lCheckRow) &&
		!CanSetFocusCell(hWndCheckCol, lCheckRow) &&
		GetEffCellType(lRow, hWndCol) != COL_CellType_CheckBox
		);
}


// IsOwnerDrawnItem
BOOL MTBLSUBCLASS::IsOwnerDrawnItem(HWND hWndParam, long lParam, int iType)
{
	CMTblCol *pCol;
	CMTblColHdrGrp *pColHdrGrp;
	CMTblRow *pRow;

	switch (iType)
	{
	case MTBL_ODI_CELL:
		pCol = m_Rows->GetColPtr(lParam, hWndParam);
		//if ( pCol && pCol->m_bOwnerdrawn ) return TRUE;
		if (pCol && pCol->IsOwnerdrawn()) return TRUE;
		if (IsOwnerDrawnItem(hWndParam, TBL_Error, MTBL_ODI_COLUMN)) return TRUE;
		if (IsOwnerDrawnItem(m_hWndTbl, lParam, MTBL_ODI_ROW)) return TRUE;
		return FALSE;
	case MTBL_ODI_COLUMN:
		pCol = m_Cols->Find(hWndParam);
		if (!pCol) return FALSE;
		//return pCol->m_bOwnerdrawn;
		return pCol->IsOwnerdrawn();
	case MTBL_ODI_COLHDR:
		pCol = m_Cols->FindHdr(hWndParam);
		//if ( pCol && pCol->m_bOwnerdrawn ) return TRUE;
		if (pCol && pCol->IsOwnerdrawn()) return TRUE;
		if (IsOwnerDrawnItem(hWndParam, TBL_Error, MTBL_ODI_COLUMN)) return TRUE;
		return FALSE;
	case MTBL_ODI_COLHDRGRP:
		pColHdrGrp = (CMTblColHdrGrp*)lParam;
		if (!m_ColHdrGrps->IsValidGrp(pColHdrGrp)) return FALSE;
		return pColHdrGrp->IsOwnerdrawn();
	case MTBL_ODI_ROW:
		pRow = m_Rows->GetRowPtr(lParam);
		if (!pRow) return FALSE;
		return pRow->IsOwnerdrawn();
	case MTBL_ODI_ROWHDR:
		pRow = m_Rows->GetRowPtr(lParam);
		if (pRow && pRow->IsHdrOwnerdrawn()) return TRUE;
		if (IsOwnerDrawnItem(hWndParam, lParam, MTBL_ODI_ROW)) return TRUE;
		return FALSE;
	case MTBL_ODI_CORNER:
		return m_bCornerOwnerdrawn;
	default:
		return FALSE;
	}
}


// IsRow
BOOL MTBLSUBCLASS::IsRow(long lRow)
{
	return Ctd.TblIsRow(m_hWndTbl, lRow);
}


// IsScrollBarV
int MTBLSUBCLASS::IsScrollBarV(HWND hWndScr, int iScrCtl, CMTblMetrics *ptm /*=NULL*/)
{
	BOOL bVert = FALSE;
	BOOL bSplitBar = FALSE;

	if (iScrCtl == SB_CTL)
	{
		CMTblMetrics tm;
		if (!ptm)
		{
			tm.Get(m_hWndTbl);
			ptm = &tm;
		}

		HWND hWndScrGet = NULL;
		int iScrCtlGet;

		// Vertikale Scrollbar?
		GetScrollBarV(ptm, FALSE, &hWndScrGet, &iScrCtlGet);
		if (hWndScrGet == hWndScr)
			bVert = TRUE;
		else
		{
			hWndScrGet = NULL;
			GetScrollBarV(ptm, TRUE, &hWndScrGet, &iScrCtlGet);
			if (hWndScrGet == hWndScr)
			{
				bVert = TRUE;
				bSplitBar = TRUE;
			}
		}
	}
	else if (iScrCtl == SB_VERT)
		bVert = TRUE;

	if (bVert && bSplitBar)
		return 2;
	else if (bVert)
		return 1;
	else
		return 0;
}


// IsSplitted
BOOL MTBLSUBCLASS::IsSplitted()
{
	return m_bSplitted;
}


// IsStandaloneColHdrGrpCol
BOOL MTBLSUBCLASS::IsStandaloneColHdrGrpCol(HWND hWndCol)
{
	//if ( !Ctd.TblQueryColumnFlags( hWndCol, COL_Visible ) ) return FALSE;
	if (!SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible)) return FALSE;

	CMTblColHdrGrp *pGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
	if (!pGrp) return FALSE;

	HWND hWnd, hWndPrevCol = NULL, hWndNextCol = NULL;
	hWnd = hWndCol;
	//if ( MTblFindPrevCol( m_hWndTbl, &hWnd, COL_Visible, 0 ) )
	if (FindPrevCol(&hWnd, COL_Visible, 0))
		hWndPrevCol = hWnd;
	hWnd = hWndCol;
	//if ( MTblFindNextCol( m_hWndTbl, &hWnd, COL_Visible, 0 ) )
	if (FindNextCol(&hWnd, COL_Visible, 0))
		hWndNextCol = hWnd;
	if (!hWndPrevCol && !hWndNextCol)
		return TRUE;

	BOOL bLocked = IsLockedColumn(hWndCol);
	CMTblColHdrGrp *pPrevGrp, *pNextGrp;
	pPrevGrp = m_ColHdrGrps->GetGrpOfCol(hWndPrevCol);
	if (pPrevGrp == pGrp && IsLockedColumn(hWndPrevCol) != bLocked)
		pPrevGrp = NULL;
	pNextGrp = m_ColHdrGrps->GetGrpOfCol(hWndNextCol);
	if (pNextGrp == pGrp && IsLockedColumn(hWndNextCol) != bLocked)
		pNextGrp = NULL;
	if (pPrevGrp != pGrp && pNextGrp != pGrp)
		return TRUE;

	return FALSE;
}


// IsTipTypeEnabled
BOOL MTBLSUBCLASS::IsTipTypeEnabled(DWORD dwTipType)
{
	return m_dwEnabledTipTypes & dwTipType;
}


// KillEdit
long MTBLSUBCLASS::KillEditEsc(HWND hWnd)
{
	LPMTBLLISTCLIPCTRLSUBCLASS psc = NULL;
	long lRet = 0;

	if (hWnd == NULL)
		hWnd = GetFocus();

	if (IsMyWindow(hWnd) && (psc = (LPMTBLLISTCLIPCTRLSUBCLASS)GetProp(hWnd, MTBL_LISTCLIP_CTRL_SUBCLASS_PROP)))
	{
		if (Validate(NULL, NULL))
		{
			// Defekt #189: Hier nochmal prüfen, ob das Edit noch existiert, denn bei der Validierung könnte der Editier-Modus evtl. beendet worden sein
			if (IsWindow(hWnd))
			{
				BeforeKillEdit();
				lRet = CallWindowProc(psc->wpOld, hWnd, WM_KEYDOWN, VK_ESCAPE, 0);
				AfterKillEdit();
			}			
		}
	}
	return lRet;
}


// KillEditNeutral
BOOL MTBLSUBCLASS::KillEditNeutral()
{
	if (m_hWndListClip)
	{
		BeforeKillEdit();
		DestroyWindow(m_hWndListClip);
		UpdateWindow(m_hWndTbl);
		AfterKillEdit();
	}

	return TRUE;
}


// KillFocusFrameI
BOOL MTBLSUBCLASS::KillFocusFrameI()
{
	m_bKillFocusFrameI = TRUE;
	BOOL bRet = Ctd.TblKillFocus(m_hWndTbl);
	m_bKillFocusFrameI = FALSE;

	return bRet;
}


// InvalidateBtn
BOOL MTBLSUBCLASS::InvalidateBtn(HWND hWndCol, long lRow, CMTblMetrics *ptm, BOOL bErase)
{
	RECT r;
	if (!GetBtnRect(hWndCol, lRow, ptm, &r)) return FALSE;

	r.right += 1;
	r.bottom += 1;
	return InvalidateRect(m_hWndTbl, &r, bErase);
}


// InvalidateCell
BOOL MTBLSUBCLASS::InvalidateCell(HWND hWndCol, long lRow, CMTblMetrics *ptm, BOOL bErase)
{
	int iRet;
	RECT r;

	iRet = GetCellRect(hWndCol, lRow, ptm, &r, MTGR_VISIBLE_PART | MTGR_INCLUDE_MERGED);

	if (iRet == MTGR_RET_ERROR)
		return FALSE;

	if (iRet == MTGR_RET_VISIBLE || iRet == MTGR_RET_PARTLY_VISIBLE)
	{
		r.right += 1;
		r.bottom += 1;
		InvalidateRect(m_hWndTbl, &r, bErase);
	}

	if (IsColMirrored(hWndCol))
		InvalidateRowHdr(lRow, bErase, TRUE);

	return TRUE;
}


// InvalidateCol
BOOL MTBLSUBCLASS::InvalidateCol(HWND hWndCol, BOOL bErase)
{
	int iRet;
	RECT r;

	if (hWndCol && m_Rows->AnyMergeCellInCol(hWndCol) || IsColMirrored(hWndCol))
	{
		if (GetClientRect(m_hWndTbl, &r))
			iRet = MTGR_RET_VISIBLE;
		else
			iRet = MTGR_RET_ERROR;
	}
	else
		iRet = GetColRect(hWndCol, &r);

	if (iRet == MTGR_RET_ERROR || iRet == MTGR_RET_INVISIBLE)
		return FALSE;

	r.right += 1;
	r.bottom += 1;
	return InvalidateRect(m_hWndTbl, &r, bErase);

	return TRUE;
}

// InvalidateColHdr
BOOL MTBLSUBCLASS::InvalidateColHdr(BOOL bErase)
{
	CMTblMetrics tm(m_hWndTbl);

	RECT r;
	CopyRect(&r, &tm.m_ClientRect);
	r.bottom = tm.m_HeaderHeight;

	return InvalidateRect(m_hWndTbl, &r, bErase);
}

BOOL MTBLSUBCLASS::InvalidateColHdr(HWND hWndCol, BOOL bErase)
{
	return InvalidateColHdr(hWndCol, FALSE, bErase);
}

BOOL MTBLSUBCLASS::InvalidateColHdr(HWND hWndCol, BOOL bInclColHdrGrp, BOOL bErase)
{
	RECT r;
	DWORD dwFlags = MTGR_VISIBLE_PART;
	if (!bInclColHdrGrp)
		dwFlags |= MTGR_EXCLUDE_COLHDRGRP;
	if (GetColHdrRect(hWndCol, r, dwFlags) > MTGR_RET_INVISIBLE)
	{
		r.right += 1;
		r.bottom += 1;
		return InvalidateRect(m_hWndTbl, &r, bErase);
	}		
	else
		return FALSE;
}


// InvalidateColSelChanges
BOOL MTBLSUBCLASS::InvalidateColSelChanges(HWND hWndCol /*= NULL*/, BOOL bErase /*= FALSE*/)
{
	BOOL bOk = FALSE;

	HRGN hRgn = GetColSelChangesRgn(hWndCol);
	if (hRgn)
	{
		bOk = InvalidateRgn(m_hWndTbl, hRgn, bErase);
		DeleteObject(hRgn);
	}

	return bOk;
}
/*
BOOL MTBLSUBCLASS::InvalidateColSelChanges( HWND hWndCol, BOOL bErase )
{
BOOL bAnyInv = FALSE, bSelAfter, bSelBefore;
MTblColList::iterator it, itEnd = m_Cols->m_List.end( );

if ( !hWndCol )
{
for ( it = m_Cols->m_List.begin( ); it != itEnd; ++it )
{
bSelBefore = (*it).second.QueryInternalFlags( COL_IFLAG_SELECTED_BEFORE ) != 0;
bSelAfter = (*it).second.QueryInternalFlags( COL_IFLAG_SELECTED_AFTER ) != 0;
if ( bSelBefore != bSelAfter )
{
if ( InvalidateCol( (*it).second.m_hWnd, bErase ) )
bAnyInv = TRUE;
}
}
}
else
{
CMTblCol *pCol = m_Cols->Find( hWndCol );
if ( pCol )
{
bSelBefore = pCol->QueryInternalFlags( COL_IFLAG_SELECTED_BEFORE ) != 0;
bSelAfter = pCol->QueryInternalFlags( COL_IFLAG_SELECTED_AFTER ) != 0;
if ( bSelBefore != bSelAfter )
{
if ( InvalidateCol( pCol->m_hWnd, bErase ) )
bAnyInv = TRUE;
}

}
}

return bAnyInv;
}
*/

// InvalidateCorner
BOOL MTBLSUBCLASS::InvalidateCorner(BOOL bErase)
{
	RECT r;
	GetCornerRect(&r);
	return InvalidateRect(m_hWndTbl, &r, bErase);
}

// InvalidateEdit
BOOL MTBLSUBCLASS::InvalidateEdit(BOOL bErase, BOOL bForceAdaptListClip /*=FALSE*/, BOOL bForceColorEdit /*=FALSE*/, BOOL bForceEditFont /*=FALSE*/)
{
	BOOL bOk = FALSE;
	if (bForceAdaptListClip)
	{
		m_bForceColorEdit = bForceColorEdit;
		m_bForceEditFont = bForceEditFont;

		AdaptListClip();

		m_bForceColorEdit = FALSE;
		m_bForceEditFont = FALSE;
	}

	BOOL bEdit = IsWindow(m_hWndEdit);
	BOOL bListClip = IsWindow(m_hWndListClip);

	if (bListClip)
		InvalidateRect(m_hWndListClip, NULL, bErase);
	if (bEdit)
		InvalidateRect(m_hWndEdit, NULL, bErase);

	if (bForceAdaptListClip)
	{
		if (bListClip)
			UpdateWindow(m_hWndListClip);
		if (bEdit)
			UpdateWindow(m_hWndEdit);
	}
	return bOk;
}

// InvalidateRow
BOOL MTBLSUBCLASS::InvalidateRow(long lRow, BOOL bErase, BOOL bInclMerged /*=FALSE*/, BOOL bToEndOfArea /*=FALSE*/)
{
	int iRet;
	RECT r;

	if (lRow == TBL_Error)
		return FALSE;

	DWORD dwFlags = 0;
	if (bInclMerged)
		dwFlags |= MTGR_INCLUDE_MERGED;
	if (bToEndOfArea)
		dwFlags |= MTGR_TO_END_OF_AREA;

	iRet = GetRowRect(lRow, &r, dwFlags);
	if (iRet == MTGR_RET_ERROR || iRet == MTGR_RET_INVISIBLE)
		return FALSE;
	r.right += 1;
	r.bottom += 1;

	return InvalidateRect(m_hWndTbl, &r, bErase);
}

// InvalidateRowFocus
BOOL MTBLSUBCLASS::InvalidateRowFocus(long lRow, BOOL bErase)
{
	int iRet;
	RECT r;

	iRet = GetRowRect(lRow, &r);

	if (iRet == MTGR_RET_ERROR || iRet == MTGR_RET_INVISIBLE)
		return FALSE;

	RECT rInv;
	SetRect(&rInv, r.left, r.top - 2, r.right, r.top + 2);
	BOOL bOk1 = InvalidateRect(m_hWndTbl, &rInv, bErase);

	SetRect(&rInv, r.left, r.bottom - 2, r.right, r.bottom + 2);
	BOOL bOk2 = InvalidateRect(m_hWndTbl, &rInv, bErase);

	return bOk1 || bOk2;
}


// InvalidateRowHdr
BOOL MTBLSUBCLASS::InvalidateRowHdr(long lRow, BOOL bErase, BOOL bInclMerged /*=FALSE*/)
{
	int iRet;
	RECT r;

	iRet = GetRowHdrRect(lRow, &r, bInclMerged ? MTGR_INCLUDE_MERGED : 0);
	if (iRet == MTGR_RET_ERROR || iRet == MTGR_RET_INVISIBLE)
		return FALSE;
	r.right += 1;
	r.bottom += 1;

	return InvalidateRect(m_hWndTbl, &r, bErase);
}


// InvalidateRowSelChanges
BOOL MTBLSUBCLASS::InvalidateRowSelChanges(long lRow /*= TBL_Error*/, BOOL bErase /*= FALSE*/)
{
	BOOL bOk = FALSE;

	HRGN hRgn = GetRowSelChangesRgn(lRow);
	if (hRgn)
	{
		bOk = InvalidateRgn(m_hWndTbl, hRgn, bErase);
		DeleteObject(hRgn);
	}

	return bOk;
}


// InvalidateTreeCol
BOOL MTBLSUBCLASS::InvalidateTreeCol(BOOL bErase)
{
	if (!m_hWndTreeCol) return TRUE;

	return InvalidateCol(m_hWndTreeCol, bErase);
}

// IsButtonHotkeyPressed
BOOL MTBLSUBCLASS::IsButtonHotkeyPressed()

{
	if (!m_iBtnKey) return FALSE;

	BOOL bDown, bPressed = FALSE;
	short shAlt, shCtrl, shKey, shShift;
	for (;;)
	{

		shKey = GetAsyncKeyState(m_iBtnKey);
		bDown = (shKey & 0x8000);
		if (!bDown)
			break;

		if (m_iBtnKey != VK_SHIFT)
		{
			shShift = GetAsyncKeyState(VK_SHIFT);
			bDown = (shShift & 0x8000);
			if (m_bBtnKeyShift && !bDown)
				break;
		}

		if (m_iBtnKey != VK_CONTROL)
		{
			shCtrl = GetAsyncKeyState(VK_CONTROL);
			bDown = (shCtrl & 0x8000);
			if (m_bBtnKeyCtrl && !bDown)
				break;
		}

		if (m_iBtnKey != VK_MENU)
		{
			shAlt = GetAsyncKeyState(VK_MENU);
			bDown = (shAlt & 0x8000);
			if (m_bBtnKeyAlt && !bDown)
				break;
		}

		bPressed = TRUE;
		break;
	}

	return bPressed;
}

// IsCellDisabled
BOOL MTBLSUBCLASS::IsCellDisabled(HWND hWndCol, long lRow)
{
	if (!hWndCol || lRow == TBL_Error)
		return FALSE;

	if (m_Rows->IsRowDisabled(lRow)) return TRUE;
	if (!Ctd.IsWindowEnabled(hWndCol)) return TRUE;
	if (m_Rows->IsRowDisCol(lRow, hWndCol)) return TRUE;
	return FALSE;
}

// IsCellReadOnly
BOOL MTBLSUBCLASS::IsCellReadOnly(HWND hWndCol, long lRow)
{
	if (!hWndCol || lRow == TBL_Error)
		return FALSE;

	CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
	if (pRow)
	{
		if (pRow->QueryFlags(MTBL_ROW_READ_ONLY))
			return TRUE;

		CMTblCol *pCell = pRow->m_Cols->Find(hWndCol);
		if (pCell && pCell->QueryFlags(MTBL_CELL_FLAG_READ_ONLY))
			return TRUE;
	}

	CMTblCol *pCol = m_Cols->Find(hWndCol);
	if (pCol && pCol->QueryFlags(MTBL_COL_FLAG_READ_ONLY))
		return TRUE;

	return FALSE;
}


// IsCheckBox
BOOL MTBLSUBCLASS::IsCheckBox(HWND hWnd)
{
	TCHAR szClsName[255];
	return (GetClassName(hWnd, szClsName, (sizeof(szClsName) / sizeof(TCHAR)) - 1) && !lstrcmp(szClsName, CLSNAME_CHECKBOX));
}


// IsCheckBoxCell
BOOL MTBLSUBCLASS::IsCheckBoxCell(HWND hWndCol, long lRow)
{
	return GetEffCellType(lRow, hWndCol) == COL_CellType_CheckBox;
}


// IsCheckBoxCol
BOOL MTBLSUBCLASS::IsCheckBoxCol(HWND hWndCol)
{
	return GetEffCellType(TBL_Error, hWndCol) == COL_CellType_CheckBox;
}


// IsColMirrored
BOOL MTBLSUBCLASS::IsColMirrored(HWND hWndCol)
{
	if (!hWndCol)
		return FALSE;

	BOOL bIs = FALSE;
	HSTRING hsTitle = 0;
	HWND hWndMirrorCol;
	int iWid;
	WORD wFlags;

	if (Ctd.TblQueryRowHeader(m_hWndTbl, &hsTitle, 0, &iWid, &wFlags, &hWndMirrorCol))
	{
		if (hWndMirrorCol && hWndMirrorCol == hWndCol && !(wFlags & TBL_RowHdr_MarkEdits) && (wFlags & TBL_RowHdr_Visible))
			bIs = TRUE;
	}

	if (hsTitle)
		Ctd.HStringUnRef(hsTitle);

	return bIs;
}


// ListClipCtrlWndProc
LRESULT CALLBACK MTBLSUBCLASS::ListClipCtrlWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	BOOL bAdaptListClip, bBack, bListOpen, bMultiline, bSetFocusRet, bSetFocusToNextCell, bUp;
	HWND hWndCol = NULL, hWndColFirst, hWndTbl;
	int iLinesPerRow, iPos, iPosFirst;
	long lRet = 0, lMinRow, lMaxRow, lRow = TBL_Error, lRowFirst, lRowNext;
	WNDPROC wpOld;

	// Subclass-Struktur ermitteln
	LPMTBLLISTCLIPCTRLSUBCLASS psc = (LPMTBLLISTCLIPCTRLSUBCLASS)GetProp(hWnd, MTBL_LISTCLIP_CTRL_SUBCLASS_PROP);
	if (!psc) return 0;

	// Subclass-Struktur der Tabelle ermitteln
	LPMTBLSUBCLASS psct = MTblGetSubClass(psc->hWndTbl);
	if (!psct) return 0;

	// Alte Fensterprozedur prüfen
	if (!psc->wpOld) return 0;

	hWndTbl = psct->m_hWndTbl;

	// Typ ermitteln
	//TCHAR szClsName[255];
	// Drop-Down-Button?
	/*BOOL bDDBtn = FALSE;
	if ( GetClassName( hWnd, szClsName, ( sizeof( szClsName ) / sizeof(TCHAR) ) - 1 ) && !lstrcmp( szClsName, CLSNAME_DROPDOWNBTN ) )
	bDDBtn = TRUE;*/
	BOOL bDDBtn = psct->IsDropDownBtn(hWnd);
	// Checkbox?
	/*BOOL bCheckbox = FALSE;
	if ( GetClassName( hWnd, szClsName, ( sizeof( szClsName ) / sizeof(TCHAR) ) - 1 ) && !lstrcmp( szClsName, CLSNAME_CHECKBOX ) )
	bCheckbox = TRUE;*/
	BOOL bCheckbox = psct->IsCheckBox(hWnd);

	// Windows-Theme verwenden?
	CTheme &th = theTheme();
	BOOL bUseWinTheme = th.CanUse() && th.IsAppThemed() && !psct->QueryFlags(MTBL_FLAG_IGNORE_WIN_THEME);

	// Variable Zeilenhöhen eingeschaltet?
	BOOL bVarRowHt = psct->QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT);

	// Zeilenhöhen?
	BOOL bRowHeights = psct->m_pCounter->GetRowHeights() > 0;

	// Kein Spaltenkopf?
	BOOL bNoColHdr = psct->QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	// Nachrichten verarbeiten...
	if (bDDBtn) // Drop-Down-Button
	{
		switch (uMsg)
		{
		case WM_MOUSEMOVE:
			// Alte Fensterprozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			// Ggf. internen Timer starten
			if (!psct->m_uiTimer)
				psct->StartInternalTimer(TRUE);
			break;
		case WM_NCDESTROY:
			// Alte Fensterprozedur setzen + merken
			SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)psc->wpOld);
			wpOld = psc->wpOld;
			// Property entfernen
			RemoveProp(hWnd, MTBL_LISTCLIP_CTRL_SUBCLASS_PROP);
			// Speicher freigeben
			delete psc;
			// Alte Fensterprozedur aufrufen
			return CallWindowProc(wpOld, hWnd, uMsg, wParam, lParam);
			break;

		case WM_PAINT:
			//return CallWindowProc( psc->wpOld, hWnd, uMsg, wParam, lParam);
			HDC hDC;
			PAINTSTRUCT ps;

			hDC = BeginPaint(hWnd, &ps);
			if (hDC)
			{
				BOOL bPushed = FALSE;

				POINT pt;
				if (GetCapture() == hWnd && GetCursorPos(&pt) && WindowFromPoint(pt) == hWnd)
				{
					SHORT ks = GetKeyState(VK_LBUTTON);
					bPushed = HIBYTE(ks);

					Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol);
					if (!MTblIsListOpen(hWndTbl, hWndCol))
					{
						bPushed = FALSE;
					}
				}

				MTBLPAINTDDBTN pb;
				pb.bPushed = bPushed;
				pb.hDC = hDC;
				GetClientRect(hWnd, &pb.rBtn);

				psct->PaintDDBtn(&pb);

				EndPaint(hWnd, &ps);
				return 0;
			}
			else
				return CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			break;
		default:
			return CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		}
	}
	else // Alle außer Drop-Down-Button...
	{
		switch (uMsg)
		{
		case BM_SETCHECK:
		case BM_SETSTATE:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			if (bCheckbox && bUseWinTheme)
				InvalidateRect(hWnd, NULL, FALSE);

			// Ggf. Zeilenkopf neu zeichnen
			if ((bRowHeights || bNoColHdr) && Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol))
				psct->InvalidateRowHdr(lRow, FALSE, TRUE);
			break;

		case WM_CHAR:
			// Focus ermitteln
			if (!Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol))
				return CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Spalte benachrichtigen
			if (SendMessage(hWndCol, MTM_Char, wParam, lParam) == MTBL_NODEFAULTACTION)
				return 0;

			// Button-Hotkey?
			if (psct->m_iBtnKey && psct->MustShowBtn(hWndCol, lRow, TRUE) && psct->IsButtonHotkeyPressed())
			{
				CMTblBtn *pbtn = psct->GetEffCellBtn(hWndCol, lRow);
				if (!(pbtn && pbtn->IsDisabled()))
					return 0;
			}

			// Read-Only?
			/*if ( psct->IsCellReadOnly( hWndCol, lRow ) )
			{
			BOOL bDown;
			BOOL bReturnDef = TRUE;
			SHORT shKey;

			// [Strg] + [C]?
			shKey = GetAsyncKeyState( 0x43 );
			bDown = (shKey & 0x8000);
			if ( bDown )
			{
			shKey = GetAsyncKeyState( VK_CONTROL );
			bDown = (shKey & 0x8000);
			if ( bDown )
			bReturnDef = FALSE;
			}

			if ( bReturnDef )
			return DefWindowProc( hWnd, uMsg, wParam, lParam );
			else
			return CallWindowProc( psc->wpOld, hWnd, uMsg, wParam, lParam);
			}*/

			// Alte Fensterprozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Ggf. Listclip wegen Größe anpassen
			bAdaptListClip = IsWindow(psct->m_hWndEdit);
			if (bAdaptListClip)
				psct->AdaptListClip();

			// Ggf. Zeilenkopf neu zeichnen
			if (bRowHeights || bNoColHdr)
				psct->InvalidateRowHdr(lRow, FALSE, TRUE);

			return lRet;

		case WM_KEYDOWN:
			// Focus ermitteln + merken
			if (!Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol))
				return CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			hWndColFirst = hWndCol;
			lRowFirst = lRow;

			// Spalte benachrichtigen
			if (SendMessage(hWndCol, MTM_KeyDown, wParam, lParam) == MTBL_NODEFAULTACTION)
				return 0;

			// Button-Hotkey?
			if (psct->m_iBtnKey && wParam == psct->m_iBtnKey && psct->MustShowBtn(hWndCol, lRow, TRUE) && psct->IsButtonHotkeyPressed())
			{
				CMTblBtn *pbtn = psct->GetEffCellBtn(hWndCol, lRow);
				if (!(pbtn && pbtn->IsDisabled()))
				{
					PostMessage(hWndTbl, MTM_BtnClick, WPARAM(hWndCol), lRow);
					return 0;
				}
			}

			switch (wParam)
			{
			case VK_BACK:
				// Read-Only?
				/*if ( psct->IsCellReadOnly( hWndCol, lRow ) )
				return DefWindowProc( hWnd, uMsg, wParam, lParam );*/

				// Ggf. Zeilenkopf neu zeichnen
				if (bRowHeights || bNoColHdr)
					psct->InvalidateRowHdr(lRow, FALSE, TRUE);

				break;

			case VK_ESCAPE:
				return psct->KillEditEsc(hWnd);

			case VK_DELETE:
				// Read-Only?
				/*if ( psct->IsCellReadOnly( hWndCol, lRow ) )
				return DefWindowProc( hWnd, uMsg, wParam, lParam );*/

				// Alte Fensterprozedur aufrufen
				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

				// Ggf. Listclip wegen Größe anpassen
				bAdaptListClip = IsWindow(psct->m_hWndEdit);
				if (bAdaptListClip)
					psct->AdaptListClip();

				// Ggf. Zeilenkopf neu zeichnen
				if (bRowHeights || bNoColHdr)
					psct->InvalidateRowHdr(lRow, FALSE, TRUE);

				return lRet;

			case VK_TAB:
				// Geht's zurück?
				bBack = GetKeyState(VK_SHIFT) < 0;

				// Validierung
				if (!psct->Validate(hWndColFirst, hWndColFirst))
					return 0;


				// Ausgangposition merken
				iPosFirst = Ctd.TblQueryColumnPos(hWndColFirst);
				// Position der Spalte ermitteln			
				iPos = iPosFirst;

				while (TRUE)
				{
					// Weiter geht's
					bBack ? --iPos : ++iPos;
					// Spalte ermitteln
					hWndCol = Ctd.TblGetColumnWindow(hWndTbl, iPos, COL_GetPos);

					if (!hWndCol)
					{
						if (bBack)
						{
							// Letzte Spalte suchen
							hWndCol = NULL;
							//while ( MTblFindNextCol( hWndTbl, &hWndCol, 0, 0 ) );
							psct->FindPrevCol(&hWndCol, 0, 0);
							if (!hWndCol) break;
							iPos = Ctd.TblQueryColumnPos(hWndCol);
							// Vorherige Zeile suchen
							lRowNext = lRow;
							lMinRow = (lRow >= 0 ? 0 : TBL_MinSplitRow);
							if (!psct->FindPrevRow(&lRowNext, 0, ROW_Hidden, lMinRow)) break;
							lRow = lRowNext;
						}
						else
						{
							// Erste Spalte ermitteln
							iPos = 1;
							hWndCol = Ctd.TblGetColumnWindow(hWndTbl, iPos, COL_GetPos);
							if (!hWndCol) break;
							// Nächste Zeile ermitteln
							lMaxRow = (lRow >= 0 ? TBL_MaxRow : TBL_MaxSplitRow);
							lRowNext = lRow;

							LPMTBLMERGECELL pmc = psct->FindMergeCell(hWndCol, lRow, FMC_ANY);
							if (pmc && pmc->lRowTo > lRowNext)
								lRowNext = pmc->lRowTo;

							LPMTBLMERGEROW pmr = psct->FindMergeRow(lRow, FMR_ANY);
							if (pmr && pmr->lRowTo > lRowNext)
								lRowNext = pmr->lRowTo;

							if (!psct->FindNextRow(&lRowNext, 0, ROW_Hidden, lMaxRow))
							{
								if (SendMessage(hWndTbl, SAM_EndCellTab, 0, lRow))
									return 0;
								else
								{
									PostMessage(hWndTbl, MTM_Internal, IA_KILL_EDIT_ESC, (LPARAM)hWnd);
									return 0;
								}
							}
							lRow = lRowNext;
						}
					}

					if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible))
					{
						if (psct->CanSetFocusCell(hWndCol, lRow))
						{
							BOOL bSet = TRUE;
							HWND hWndColSet = hWndCol, hWndMergeCol;
							long lRowSet = lRow, lMergeRow;

							if (psct->GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
							{
								if (!bBack && hWndMergeCol == hWndColFirst)
									bSet = FALSE;
								else if (psct->CanSetFocusCell(hWndMergeCol, lMergeRow))
								{
									hWndColSet = hWndMergeCol;
									lRowSet = lMergeRow;
								}
								else
									bSet = FALSE;
							}

							if (bSet)
							{
								bSetFocusRet = Ctd.TblSetFocusCell(hWndTbl, lRowSet, hWndColSet, 0, -1);
								if (bSetFocusRet)
								{
									PostMessage(hWndColSet, SAM_Click, 0, lRowSet);
									PostMessage(hWndTbl, SAM_Click, 0, lRowSet);
								}
								return 0;
							}
						}
					}
				}

				// Keine Aktion durchgeführt - jetzt ist KillEdit angesagt!
				// PostMessage( hWnd, WM_KEYDOWN, VK_ESCAPE, 0 );
				PostMessage(hWndTbl, MTM_Internal, IA_KILL_EDIT_ESC, (LPARAM)hWnd);
				return 0;

			case VK_NEXT:
			case VK_PRIOR:
				// Ermitteln, ob Liste offen ist
				bListOpen = MTblIsListOpen(hWndTbl, hWndCol);

				// Return, wenn Liste offen und Zelle Read-Only
				if (bListOpen && psct->IsCellReadOnly(hWndCol, lRow))
					return 0;
				break;

			case VK_UP:
			case VK_DOWN:
				// Ermitteln, ob Liste offen ist
				bListOpen = MTblIsListOpen(hWndTbl, hWndCol);

				// Return, wenn Liste offen und Zelle Read-Only
				if (bListOpen && psct->IsCellReadOnly(hWndCol, lRow))
					return 0;

				// Multiline?
				bMultiline = psct->GetCellWordWrap(hWndCol, lRow);

				// Fokus zur nächsten Zelle?
				bSetFocusToNextCell = !(bListOpen || bMultiline);

				if (psct->QueryFlags(MTBL_FLAG_MOVE_INP_FOCUS_UD_EX) && bMultiline && !bListOpen && IsWindow(psct->m_hWndEdit))
				{
					DWORD dwEnd = 0;
					SendMessage(psct->m_hWndEdit, EM_GETSEL, NULL, (LPARAM)&dwEnd);

					long lLine = SendMessage(psct->m_hWndEdit, EM_LINEFROMCHAR, (WPARAM)dwEnd, 0);

					if (wParam == VK_UP && lLine == 0)
						bSetFocusToNextCell = TRUE;
					else if (wParam == VK_DOWN)
					{
						long lLines = SendMessage(psct->m_hWndEdit, EM_GETLINECOUNT, 0, 0);
						if (lLine == lLines - 1)
							bSetFocusToNextCell = TRUE;
					}
				}

				// Wenn Fokus zur nächsten Zelle gesetzt werden soll...
				if ( /*!( bListOpen || bMultiline )*/ bSetFocusToNextCell)
				{
					// Validierung
					if (!psct->Validate(hWndCol, hWndCol))
						return 0;

					// Rauf oder runter?
					bUp = (wParam == VK_UP);
					while (TRUE)
					{
						// Nächste Zeile suchen
						lRowNext = lRow;
						if (bUp)
						{
							if (!psct->FindPrevRow(&lRowNext, 0, ROW_Hidden))
								break;
						}
						else
						{
							if (!psct->FindNextRow(&lRowNext, 0, ROW_Hidden))
								break;
						}

						// Übersprung von Split-Bereich in normalen Bereich und umgekehrt verhindern
						if ((lRow < 0 && lRowNext >= 0) || (lRow >= 0 && lRowNext < 0))
							break;

						lRow = lRowNext;

						if (psct->CanSetFocusCell(hWndCol, lRow))
						{
							BOOL bSet = TRUE;
							HWND hWndColSet = hWndCol, hWndMergeCol;
							long lRowSet = lRow, lMergeRow;

							if (psct->GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
							{
								if (!bUp && lMergeRow == lRowFirst)
									bSet = FALSE;
								else if (psct->CanSetFocusCell(hWndMergeCol, lMergeRow))
								{
									hWndColSet = hWndMergeCol;
									lRowSet = lMergeRow;
								}
								else
									bSet = FALSE;
							}

							if (bSet)
							{
								bSetFocusRet = Ctd.TblSetFocusCell(hWndTbl, lRowSet, hWndColSet, 0, -1);
								if (bSetFocusRet)
								{
									PostMessage(hWndColSet, SAM_Click, 0, lRow);
									PostMessage(hWndTbl, SAM_Click, 0, lRow);
								}
								return 0;
							}
						}
					}
					// Keine Aktion durchgeführt - jetzt ist KillEdit angesagt!
					// PostMessage( hWnd, WM_KEYDOWN, VK_ESCAPE, 0 );
					PostMessage(hWndTbl, MTM_Internal, IA_KILL_EDIT_ESC, (LPARAM)hWnd);
					return 0;
				}
				// ... wenn hingegen Multiline...
				else if (bMultiline /*&& !bListOpen*/)
				{
					// Multiline nicht Standard?
					if (Ctd.TblQueryLinesPerRow(hWndTbl, &iLinesPerRow) && iLinesPerRow < 2)
					{
						// Fensterprozedur der Klasse "Edit" aufrufen
						WNDCLASS wc;
						if (GetClassInfo((HINSTANCE)ghInstance, _T("Edit"), &wc) && wc.lpfnWndProc)
							return CallWindowProc(wc.lpfnWndProc, hWnd, uMsg, wParam, lParam);
						else
							return DefWindowProc(hWnd, uMsg, wParam, lParam);
					}
				}

				break;
			}

			// Alte Fensterprozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			break;

		case WM_KILLFOCUS:
			if (!psct->m_bSettingFocus)
				psct->m_bNoFocusPaint = TRUE;

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			if (!psct->m_bSettingFocus)
				psct->UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);

			break;

		case WM_LBUTTONDOWN:
			POINT pt;
			GetCursorPos(&pt);
			ScreenToClient(psct->m_hWndTbl, &pt);

			if (!psct->m_bDragCanAutoStart && SendMessage(psct->m_hWndTbl, MTM_DragCanAutoStart, pt.x, pt.y))
			{
				psct->DragDropAutoStart(wParam, lParam, hWnd);
				lRet = 0;
			}
			else
				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			break;

		case WM_LBUTTONUP:
			// DragDrop
			if (psct->m_bDragCanAutoStart)
				psct->DragDropAutoStartCancel();

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			break;

		case WM_MOUSEMOVE:
			// DragDrop
			if (psct->m_bDragCanAutoStart)
				psct->DragDropAutoStartCancel();
			// Alte Fensterprozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			// Ggf. internen Timer starten
			if (!psct->m_uiTimer)
				psct->StartInternalTimer(TRUE);

			break;

		case WM_MOUSEWHEEL:
			PostMessage(hWndTbl, WM_MOUSEWHEEL, wParam, lParam);
			return 0;

		case WM_NCDESTROY:
			// Alte Fensterprozedur setzen + merken
			SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)psc->wpOld);
			wpOld = psc->wpOld;
			// Property entfernen
			RemoveProp(hWnd, MTBL_LISTCLIP_CTRL_SUBCLASS_PROP);
			// Speicher freigeben
			delete psc;
			// Edit?
			if (hWnd == psct->m_hWndEdit)
			{
				//psct->m_hWndLastEdit = psct->m_hWndEdit;
				psct->m_hWndEdit = NULL;
			}
			// Alte Fensterprozedur aufrufen
			return CallWindowProc(wpOld, hWnd, uMsg, wParam, lParam);

		case WM_PAINT:
			// Checkbox?
			if (bCheckbox && bUseWinTheme)
			{
				PAINTSTRUCT ps;
				BeginPaint(hWnd, &ps);

				RECT r;
				GetClientRect(hWnd, &r);

				HTHEME hTheme = th.OpenThemeData(hWnd, L"Button");
				int iPart = BP_CHECKBOX;

				int iState = 0;
				long lCheckState = SendMessage(hWnd, BM_GETSTATE, 0, 0);
				if (lCheckState & BST_CHECKED)
				{
					if (lCheckState & BST_PUSHED)
						iState = CBS_CHECKEDPRESSED;
					else
						iState = CBS_CHECKEDNORMAL;
				}
				else if (lCheckState & BST_INDETERMINATE)
					iState = CBS_MIXEDNORMAL;
				else
				{
					if (lCheckState & BST_PUSHED)
						iState = CBS_UNCHECKEDPRESSED;
					else
						iState = CBS_UNCHECKEDNORMAL;
				}

				th.DrawThemeBackground(hTheme, ps.hdc, iPart, iState, &r, NULL);
				th.CloseThemeData(hTheme);

				EndPaint(hWnd, &ps);
			}
			else
				lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
			break;

		case WM_PASTE:
			// Focus ermitteln
			if (!Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol))
				return CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Read-Only?
			/*if ( psct->IsCellReadOnly( hWndCol, lRow ) )
			lRet = DefWindowProc( hWnd, uMsg, wParam, lParam );
			else
			{
			lRet = CallWindowProc( psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Ggf. Listclip wegen Größe anpassen
			bAdaptListClip = IsWindow( psct->m_hWndEdit );
			if ( bAdaptListClip )
			psct->AdaptListClip( );
			}*/

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Ggf. Listclip wegen Größe anpassen
			bAdaptListClip = IsWindow(psct->m_hWndEdit);
			if (bAdaptListClip)
				psct->AdaptListClip();

			// Ggf. Zeilenkopf neu zeichnen
			if (bRowHeights || bNoColHdr)
				psct->InvalidateRowHdr(lRow, FALSE, TRUE);

			break;

		case WM_SETFOCUS:
			if (!psct->m_bSettingFocus)
				psct->m_bNoFocusPaint = FALSE;

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			if (!psct->m_bSettingFocus)
			{
				if (psct->MustSuppressRowFocusCTD())
					psct->KillFocusFrameI();

				psct->UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);
			}

			break;

		case WM_SETTEXT:
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Ggf. Zeilenkopf neu zeichnen
			if ((bRowHeights || bNoColHdr) && Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol))
				psct->InvalidateRowHdr(lRow, FALSE, TRUE);

			break;

		case WM_SYSKEYDOWN:
			// Focus ermitteln
			if (!Ctd.TblQueryFocus(hWndTbl, &lRow, &hWndCol))
				return CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			// Button-Hotkey?
			if (psct->m_iBtnKey && wParam == psct->m_iBtnKey && psct->MustShowBtn(hWndCol, lRow, TRUE) && psct->IsButtonHotkeyPressed())
			{
				CMTblBtn *pbtn = psct->GetEffCellBtn(hWndCol, lRow);
				if (!(pbtn && pbtn->IsDisabled()))
				{
					PostMessage(hWndTbl, MTM_BtnClick, WPARAM(hWndCol), lRow);
					return 0;
				}
			}

			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);

			break;

		default:
			// Alte Fensterprozedur aufrufen
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		}
	}

	return lRet;
}


// ListClipWndProc
LRESULT CALLBACK MTBLSUBCLASS::ListClipWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	long lRet = 0;
	WNDPROC wpOld;

	// Subclass-Struktur ermitteln
	LPMTBLLISTCLIPSUBCLASS psc = (LPMTBLLISTCLIPSUBCLASS)GetProp(hWnd, MTBL_LISTCLIP_SUBCLASS_PROP);
	if (!psc) return 0;

	// Subclass-Struktur der Tabelle ermitteln
	LPMTBLSUBCLASS psct = MTblGetSubClass(psc->hWndTbl);
	if (!psct) return 0;

	// Alte Fensterprozedur prüfen
	if (!psc->wpOld) return 0;

	// Nachrichten verarbeiten
	switch (uMsg)
	{
	case WM_MOUSEMOVE:
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		// Ggf. internen Timer starten
		if (!psct->m_uiTimer)
			psct->StartInternalTimer(TRUE);

		break;
	case WM_MOVE:
		lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
		PostMessage(psct->m_hWndTbl, MTM_Internal, IA_ADAPT_LISTCLIP, 0);
		break;

	case WM_NCDESTROY:
		// Alte Fensterprozedur setzen + merken
		SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)psc->wpOld);
		wpOld = psc->wpOld;
		// Property entfernen
		RemoveProp(hWnd, MTBL_LISTCLIP_SUBCLASS_PROP);
		// Speicher freigeben
		delete psc;
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc(wpOld, hWnd, uMsg, wParam, lParam);

		break;

	case WM_PAINT:
		// Focus ermitteln
		HWND hWndCol;
		long lRow;
		if (Ctd.TblQueryFocus(psc->hWndTbl, &lRow, &hWndCol))
		{
			HRGN hRgn = CreateRectRgn(0, 0, 0, 0);
			int iRet = GetUpdateRgn(hWnd, hRgn, FALSE);

			// Zeichnen beginnen
			PAINTSTRUCT ps;
			BeginPaint(hWnd, &ps);
			HDC hDC = ps.hdc;
			if (hDC)
			{
				// Effektive Hintergrundfarbe der Zelle ermitteln
				COLORREF clrBack = psct->GetEffCellBackColor(lRow, hWndCol);
				// Pinsel erzeugen
				HBRUSH hBrush = CreateSolidBrush(clrBack);
				if (hBrush)
				{
					// Hintergrund zeichnen
					FillRgn(hDC, hRgn, hBrush);

					// Pinsel löschen
					DeleteObject(hBrush);
				}
			}

			// Zeichnen beenden
			EndPaint(hWnd, &ps);

			DeleteObject(hRgn);
		}

		break;

	case WM_PARENTNOTIFY:
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		psct->OnListClipParentNotify(psct->m_hWndTbl, hWnd, wParam, lParam);
		break;

	case WM_SETCURSOR:
		BOOL bBtn, bDDBtn;
		TCHAR cClass[255];
		HWND hWndPoint;
		POINT pt;

		GetCursorPos(&pt);
		hWndPoint = WindowFromPoint(pt);
		GetClassName(hWndPoint, cClass, 254);
		bBtn = !lstrcmp(cClass, MTBL_BTN_CLASS);
		bDDBtn = !lstrcmp(cClass, CLSNAME_DROPDOWNBTN);
		if (bBtn || bDDBtn)
		{
			SetCursor(LoadCursor(NULL, IDC_ARROW));
			lRet = TRUE;
		}
		else
			lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		break;

	case WM_SHOWWINDOW:
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
		if (wParam)
			psct->AdaptListClip();
		break;

	default:
		// Alte Fensterprozedur aufrufen
		lRet = CallWindowProc(psc->wpOld, hWnd, uMsg, wParam, lParam);
	}

	return lRet;
}


// LoadChildRows
void MTBLSUBCLASS::LoadChildRows(long lRow)
{
	CMTblRow *pRow = m_Rows->EnsureValid(lRow);
	if (pRow && m_RowsLoadingChilds)
	{
		++m_lLoadChildRows;
		m_RowsLoadingChilds->push(pRow);

		SendMessage(m_hWndTbl, MTM_LoadChildRows, 0, lRow);

		--m_lLoadChildRows;
		m_RowsLoadingChilds->pop();

		if (!m_lLoadChildRows)
		{
			CMTblRow **Rows, *pRow;
			DWORD dwFlags;
			long lLastValidPos, lUpperBound;
			RowArrLongVector v;

			Rows = m_Rows->GetRowArray(0, &lUpperBound, &lLastValidPos);
			if (Rows)
				v.push_back(RowArrLong(Rows, lUpperBound));

			Rows = m_Rows->GetRowArray(-1, &lUpperBound, &lLastValidPos);
			if (Rows)
				v.push_back(RowArrLong(Rows, lUpperBound));

			for (int i = 0; i < v.size(); i++)
			{
				Rows = v[i].first;
				lLastValidPos = v[i].second;

				for (int j = 0; j < lLastValidPos; j++)
				{
					if ((pRow = Rows[j]) && !pRow->IsDelAdj())
					{
						if (pRow->QueryInternalFlags(ROW_IFLAG_LOAD_CHILD_ROWS))
						{
							pRow->SetInternalFlags(ROW_IFLAG_LOAD_CHILD_ROWS, FALSE);
							LoadChildRows(pRow->GetNr());
							return;
						}

						if (pRow->QueryInternalFlags(ROW_IFLAG_HIDE | ROW_IFLAG_HIDE_REDRAW))
						{
							dwFlags = pRow->QueryInternalFlags(ROW_IFLAG_HIDE_REDRAW) ? MTSRF_REDRAW : 0;
							pRow->SetInternalFlags(ROW_IFLAG_HIDE | ROW_IFLAG_HIDE_REDRAW, FALSE);
							MTblSetRowFlags(m_hWndTbl, pRow->GetNr(), MTBL_ROW_HIDDEN, TRUE, dwFlags);
						}

						if (pRow->QueryInternalFlags(ROW_IFLAG_DELETE | ROW_IFLAG_DELETE_ADJUST))
						{
							dwFlags = pRow->QueryInternalFlags(ROW_IFLAG_DELETE_ADJUST) ? TBL_Adjust : TBL_NoAdjust;
							pRow->SetInternalFlags(ROW_IFLAG_DELETE | ROW_IFLAG_DELETE_ADJUST, FALSE);
							Ctd.TblDeleteRow(m_hWndTbl, pRow->GetNr(), dwFlags);
						}
					}
				}
			}
		}
	}
}


// LogLine
void MTBLSUBCLASS::LogLine(LPCTSTR lpsFile, LPCTSTR lpsText)
{
	tstring sLine(_T(""));

	HSTRING hsItemName = 0;
	Ctd.GetItemName(m_hWndTbl, &hsItemName);

	Ctd.HStrToTStr(hsItemName, sLine);

	sLine += _T(": ");
	//Ctd.HStrUnlock(hsItemName);

	if (lpsText)
		sLine += lpsText;

	if (hsItemName)
		Ctd.HStringUnRef(hsItemName);

	Ctd.LogLine(LPTSTR(lpsFile), LPTSTR(sLine.c_str()));
}


// LogLineFocusPaint
void MTBLSUBCLASS::LogLineFocusPaint(LPTSTR lpsText)
{
	tstring sText(_T(""));

	if (lpsText)
		sText = lpsText;

	sText += _T(" m_bNoFocusPaint = ");
	if (m_bNoFocusPaint)
		sText += _T("TRUE");
	else
		sText += _T("FALSE");

	LogLine(_T("focuspaint.log"), LPTSTR(sText.c_str()));
}


// LockPaint
long MTBLSUBCLASS::LockPaint(void)
{
	++m_lLockPaint;
	return m_lLockPaint;
}

// MustPaintRowFocus
BOOL MTBLSUBCLASS::MustPaintRowFocus()
{
	//return !m_bNoFocusPaint;
	return FALSE;
}

// MustRedrawSelections
BOOL MTBLSUBCLASS::MustRedrawSelections()
{
	long lSelInv = m_pCounter->GetNoSelInv();
	//long lMergeRows = m_pCounter->GetMergeRows( );
	return lSelInv > 0 /*|| lMergeRows > 0*/ || m_dwSelBackColor != MTBL_COLOR_UNDEF || m_dwSelTextColor != MTBL_COLOR_UNDEF || QueryFlags(MTBL_FLAG_FULL_OVERLAP_SELECTION) || gbCtdHooked;
}


// MustShowBtn
BOOL MTBLSUBCLASS::MustShowBtn(HWND hWndCol, long lRow, BOOL bNoWidthCheck /*=FALSE*/, BOOL bPermanent /*=FALSE*/)
{
	CMTblBtn *pBtn = GetEffCellBtn(hWndCol, lRow);
	if (!pBtn)
		return FALSE;

	if (!pBtn->m_bShow)
		return FALSE;

	if (bPermanent)
	{
		if (QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL))
			return FALSE;

		if (!QueryFlags(MTBL_FLAG_BUTTONS_PERMANENT) && !QueryColumnFlags(hWndCol, MTBL_COL_FLAG_BUTTONS_PERMANENT) && !QueryCellFlags(lRow, hWndCol, MTBL_CELL_FLAG_BUTTONS_PERMANENT))
			return FALSE;

		CMTblCol *pCol = m_Cols->Find(hWndCol);
		if (pCol && pCol->QueryFlags(MTBL_COL_FLAG_HIDE_PERM_BTN))
			return FALSE;
		BOOL bShowPermBtnNe_Col = pCol && pCol->QueryFlags(MTBL_COL_FLAG_SHOW_PERM_BTN_NE);

		BOOL bShowPermBtnNe_Cell = FALSE;
		CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
		if (pRow)
		{
			CMTblCol *pCell = pRow->m_Cols->Find(hWndCol);
			if (pCell && pCell->QueryFlags(MTBL_CELL_FLAG_HIDE_PERM_BTN))
				return FALSE;
			bShowPermBtnNe_Cell = pCell && pCell->QueryFlags(MTBL_CELL_FLAG_SHOW_PERM_BTN_NE);
		}

		if (QueryFlags(MTBL_FLAG_HIDE_PERM_BTN_NE) && !(bShowPermBtnNe_Cell || bShowPermBtnNe_Col))
		{
			if (!CanSetFocusCell(hWndCol, lRow) || IsCellReadOnly(hWndCol, lRow))
				return FALSE;
		}
	}

	// int iCellType = Ctd.TblGetCellType( hWndCol );
	int iCellType = GetEffCellType(lRow, hWndCol);
	if (iCellType == COL_CellType_PopupEdit && !bPermanent)
		return FALSE;

	if (!bNoWidthCheck && !(!bPermanent && QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL)))
	{
		int iBtnWid = CalcBtnWidth(pBtn, hWndCol, lRow);
		if (iBtnWid < 1)
			return FALSE;

		int iColWid, iLeft, iTop, iRight, iBottom;
		int iRet = MTblGetCellRectEx(hWndCol, lRow, &iLeft, &iTop, &iRight, &iBottom, MTGR_INCLUDE_MERGED);
		if (iRet != MTGR_RET_ERROR)
			iColWid = iRight - iLeft;
		else
			iColWid = SendMessage(hWndCol, TIM_QueryColumnWidth, 0, 0);

		CMTblMetrics tm(m_hWndTbl);

		if (iCellType == COL_CellType_CheckBox)
		{
			int iJustify = GetEffCellTextJustify(lRow, hWndCol);
			int iVAlign = GetEffCellTextVAlign(lRow, hWndCol);

			//int iAreaLeft = GetCellTextAreaLeft( lRow, hWndCol, iLeft, tm.m_CellLeadingX, m_Rows->GetColImagePtr( lRow, hWndCol ) );
			//int iAreaRight = GetCellTextAreaRight( lRow, hWndCol, iRight, tm.m_CellLeadingX, m_Rows->GetColImagePtr( lRow, hWndCol ), iJustify != DT_CENTER );
			CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);
			int iAreaLeft = GetCellTextAreaLeft(lRow, hWndCol, iLeft, tm.m_CellLeadingX, Img);
			int iAreaRight = GetCellTextAreaRight(lRow, hWndCol, iRight, tm.m_CellLeadingX, Img, iJustify != DT_CENTER);

			RECT rCb = { 0, 0, 0, 0 };
			RECT rCbArea = { iAreaLeft, iTop, iAreaRight, iBottom };

			CalcCheckBoxRect(iJustify, iVAlign, rCbArea, tm.m_CellLeadingX, tm.m_CellLeadingY, tm.m_CharBoxY, rCb);

			int iBtnLeft = iRight - iBtnWid;
			if (iBtnLeft < rCb.right)
				return FALSE;
		}
		else
		{
			int iEditWid = iColWid - iBtnWid;
			if (MustShowDDBtn(hWndCol, lRow))
				iEditWid -= MTBL_DDBTN_WID;


			int iMinWid = tm.m_CharBoxY;
			HDC hDC = GetDC(NULL);
			if (hDC)
			{
				BOOL bDelFont = FALSE;
				HFONT hFont = GetEffCellFont(hDC, lRow, hWndCol, &bDelFont);

				HGDIOBJ hOldFont = SelectObject(hDC, hFont);

				TEXTMETRIC txtm;
				if (GetTextMetrics(hDC, &txtm))
					iMinWid = txtm.tmAveCharWidth;

				SelectObject(hDC, hOldFont);
				if (bDelFont)
					DeleteObject(hFont);

				ReleaseDC(NULL, hDC);
			}

			iMinWid += (tm.m_CellLeadingX * 2);

			if (iEditWid < iMinWid)
				return FALSE;
		}
	}

	return TRUE;
}


// MustShowDDBtn
BOOL MTBLSUBCLASS::MustShowDDBtn(HWND hWndCol, long lRow, BOOL bNoWidthCheck /*=FALSE*/, BOOL bPermanent /*=FALSE*/)
{
	if (bPermanent)
	{
		if (QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL))
			return FALSE;

		if (!QueryFlags(MTBL_FLAG_BUTTONS_PERMANENT) && !QueryColumnFlags(hWndCol, MTBL_COL_FLAG_BUTTONS_PERMANENT) && !QueryCellFlags(lRow, hWndCol, MTBL_CELL_FLAG_BUTTONS_PERMANENT))
			return FALSE;

		CMTblCol *pCol = m_Cols->Find(hWndCol);
		if (pCol && pCol->QueryFlags(MTBL_COL_FLAG_HIDE_PERM_BTN))
			return FALSE;
		BOOL bShowPermBtnNe_Col = pCol && pCol->QueryFlags(MTBL_COL_FLAG_SHOW_PERM_BTN_NE);

		BOOL bShowPermBtnNe_Cell = FALSE;
		CMTblRow *pRow = m_Rows->GetRowPtr(lRow);
		if (pRow)
		{
			CMTblCol *pCell = pRow->m_Cols->Find(hWndCol);
			if (pCell && pCell->QueryFlags(MTBL_CELL_FLAG_HIDE_PERM_BTN))
				return FALSE;
			bShowPermBtnNe_Cell = pCell && pCell->QueryFlags(MTBL_CELL_FLAG_SHOW_PERM_BTN_NE);
		}

		if (!CanSetFocusCell(hWndCol, lRow))
			return FALSE;

		if (QueryFlags(MTBL_FLAG_HIDE_PERM_BTN_NE) && !(bShowPermBtnNe_Cell || bShowPermBtnNe_Col))
		{
			if (IsCellReadOnly(hWndCol, lRow))
				return FALSE;
		}
	}

	int iCellType = GetEffCellType(lRow, hWndCol);
	if (iCellType != COL_CellType_DropDownList)
		return FALSE;

	if (!bNoWidthCheck && !(!bPermanent && QueryFlags(MTBL_FLAG_BUTTONS_OUTSIDE_CELL)))
	{
		int iColWid/*, iLeft, iTop, iRight, iBottom*/;
		/*int iRet = MTblGetCellRectEx( hWndCol, lRow, &iLeft, &iTop, &iRight, &iBottom, MTGR_INCLUDE_MERGED );
		if ( iRet != MTGR_RET_ERROR )
		iColWid = iRight - iLeft;
		else*/
		iColWid = SendMessage(hWndCol, TIM_QueryColumnWidth, 0, 0);

		if (iColWid < (MTBL_DDBTN_WID + 10))
			return FALSE;
	}

	return TRUE;
}


// MustSuppressRowFocusCTD
BOOL MTBLSUBCLASS::MustSuppressRowFocusCTD()
{
	return /*m_bCellMode || QueryFlags( MTBL_FLAG_SUPPRESS_FOCUS )*/ gbCtdHooked ? FALSE : TRUE;
}


// MustTrapColSelChanges
BOOL MTBLSUBCLASS::MustTrapColSelChanges()
{
	return !m_bTrappingColSelChanges/* && ( m_bExtMsgs || m_bCellMode || MustRedrawSelections( ) )*/;
}


// MustTrapRowSelChanges
BOOL MTBLSUBCLASS::MustTrapRowSelChanges()
{
	return !m_bTrappingRowSelChanges && !QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION)/* && ( m_bExtMsgs || m_bCellMode || MustRedrawSelections( ) )*/;
}


// MustUseEndEllipsis
BOOL MTBLSUBCLASS::MustUseEndEllipsis(HWND hWndCol)
{
	//if ( Ctd.TblQueryColumnFlags( hWndCol, COL_IndicateOverflow | COL_Invisible ) )
	//	return FALSE;
	if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_IndicateOverflow | COL_Invisible))
		return FALSE;

	if (QueryFlags(MTBL_FLAG_END_ELLIPSIS))
		return TRUE;

	CMTblCol *pCol = m_Cols->Find(hWndCol);
	if (!pCol)
		return FALSE;

	return pCol->QueryFlags(MTBL_COL_FLAG_END_ELLIPSIS);
}


// NormalizeHierarchy
// Normalisiert die Hierarchie
BOOL MTBLSUBCLASS::NormalizeHierarchy(long lRow, int iReason)
{
	// Zeilenpointer ermitteln
	CMTblRow *pRow;
	if (!(pRow = m_Rows->GetRowPtr(lRow))) return FALSE;

	// Zeilenarray ermitteln
	long lLastValidPos, lUpperBound;
	CMTblRow **RowArr = m_Rows->GetRowArray(lRow, &lUpperBound, &lLastValidPos);
	if (!RowArr)
		return FALSE;

	// Auf geht's...
	BOOL bRowHidden = pRow->QueryFlags(MTBL_ROW_HIDDEN);
	BOOL bNormalized = FALSE;
	CMTblRow *pCurRow, *pParentRow;
	long l;
	// Zeile wird gelöscht oder verborgen und ist eine Elternzeile?
	if ((iReason == NORM_HIER_ROW_DELETE || iReason == NORM_HIER_ROW_HIDE) && pRow->IsParent())
	{
		for (l = 0; l <= lLastValidPos; ++l)
		{
			if (pCurRow = RowArr[l])
			{
				// Kindzeile?
				if (pCurRow->GetParentRow() == pRow)
				{
					// Auf Elterneile umhängen
					pParentRow = pRow->GetParentRow();
					SetParentRow(pCurRow->GetNr(), pParentRow ? pParentRow->GetNr() : TBL_Error, 0);
					bNormalized = TRUE;
				}
			}
		}
	}
	// Zeile wird sichtbar oder Löschen ist fehlgeschlagen?
	else if (iReason == NORM_HIER_ROW_SHOW || iReason == NORM_HIER_ROW_DELETE_FAIL)
	{
		// Auf die erste sichtbare ursprünglich übergeordnete Zeile umhängen
		if (pRow->HasParentRowChanged())
		{
			pParentRow = pRow->GetOrigParentRow();
			while (pParentRow && pParentRow->QueryFlags(MTBL_ROW_HIDDEN))
			{
				pParentRow = pParentRow->GetOrigParentRow();
			}

			if (pParentRow != pRow->GetParentRow())
			{
				SetParentRow(pRow->GetNr(), pParentRow ? pParentRow->GetNr() : TBL_Error, 0);
				bNormalized = TRUE;
			}
		}


		// Alle ursprünglich abhängigen Zeilen umhängen
		for (l = 0; l <= lLastValidPos; ++l)
		{
			if ((pCurRow = RowArr[l]) && pCurRow->HasParentRowChanged())
			{
				// Ursprüngliche Kindzeile?
				if ((pParentRow = pCurRow->GetOrigParentRow()) == pRow)
				{
					// Auf Zeile umhängen
					if (pParentRow != pCurRow->GetParentRow())
					{
						SetParentRow(pCurRow->GetNr(), pParentRow ? pParentRow->GetNr() : TBL_Error, 0);
						bNormalized = TRUE;
					}
				}
				// Ursprünglich abhängige Zeile?
				else if (pCurRow->IsOrigDescendantOf(pRow))
				{
					// Auf die erste sichtbare ursprünglich übergeordnete Zeile umhängen
					pParentRow = pCurRow->GetOrigParentRow();
					while (pParentRow && pParentRow != pRow && pParentRow->QueryFlags(MTBL_ROW_HIDDEN))
					{
						pParentRow = pParentRow->GetOrigParentRow();
					}

					if (pParentRow != pCurRow->GetParentRow())
					{
						SetParentRow(pCurRow->GetNr(), pParentRow ? pParentRow->GetNr() : TBL_Error, 0);
						bNormalized = TRUE;
					}
				}
			}
		}
	}

	return bNormalized;
}


// ObjFromPoint
BOOL MTBLSUBCLASS::ObjFromPoint(POINT pt, LPLONG plRow, LPHWND phWndCol, LPDWORD pdwFlags, LPDWORD pdwMFlags)
{
	if (!Ctd.TblObjectsFromPoint(m_hWndTbl, pt.x, pt.y, plRow, phWndCol, pdwFlags)) return FALSE;

	// Parameter initialisieren
	*pdwMFlags = 0;

	// Tabellen-Metriken ermitteln
	CMTblMetrics tm(m_hWndTbl);

	// Focus ermitteln
	HWND hWndFocusCol = NULL;
	long lFocusRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);

	// Über Zelle?
	BOOL bOverCell = *phWndCol && *plRow != TBL_Error;
	if (bOverCell)
	{
		HWND hWndCol = *phWndCol;
		long lRow = *plRow;

		// Spaltentyp ermitteln
		int iCellType = GetEffCellType(lRow, hWndCol);

		// Checkbox?
		BOOL bCheckbox = (iCellType == COL_CellType_CheckBox);

		// Merge
		BOOL bMerged = FALSE;
		HWND hWndFirstCol = NULL, hWndLastCol = NULL;
		long lFirstRow = TBL_Error, lLastRow = TBL_Error;
		if (GetMergeCellEx(hWndCol, lRow, hWndFirstCol, lFirstRow))
			bMerged = TRUE;
		else
		{
			hWndFirstCol = hWndCol;
			lFirstRow = lRow;
		}

		if (!GetLastMergedCellEx(hWndFirstCol, lFirstRow, hWndLastCol, lLastRow))
		{
			hWndLastCol = hWndFirstCol;
			lLastRow = lFirstRow;
		}

		if (bMerged)
			*pdwMFlags |= MTOFP_OVERMERGEDCELL;


		// Zellenpointer ermitteln
		CMTblCol *pCell = m_Rows->GetColPtr(lFirstRow, hWndFirstCol);

		// Zellenfont-Metriken ermitteln
		TEXTMETRIC tmCellFont;
		if (!GetCellFontMetrics(lFirstRow, hWndFirstCol, &tmCellFont))
			return FALSE;

		// Bild ermitteln
		CMTblImg Img = GetEffCellImage(lFirstRow, hWndFirstCol, GEI_HILI);

		// Buttons?
		BOOL bPermanent = TRUE;
		if (hWndFocusCol && hWndFirstCol == hWndFocusCol && lFirstRow == lFocusRow)
			bPermanent = FALSE;
		BOOL bButton = MustShowBtn(hWndFirstCol, lFirstRow, TRUE, bPermanent);
		BOOL bDDButton = MustShowDDBtn(hWndFirstCol, lFirstRow, TRUE, bPermanent);
		// Koordinaten ermitteln
		POINT ptCol, ptColVis, ptRow, ptRowVis;
		//MTblGetColXCoord( m_hWndTbl, hWndFirstCol, &tm, &ptCol, &ptColVis );
		GetColXCoord(hWndFirstCol, &tm, &ptCol, &ptColVis);
		if (hWndFirstCol != hWndLastCol)
		{
			POINT pt, ptVis;
			//MTblGetColXCoord( m_hWndTbl, hWndLastCol, &tm, &pt, &ptVis );
			GetColXCoord(hWndLastCol, &tm, &pt, &ptVis);
			ptCol.y = pt.y;
			ptColVis.y = ptVis.y;
		}

		//int iRowRet = MTblGetRowYCoord( m_hWndTbl, lFirstRow, &tm, &ptRow, &ptRowVis, lFirstRow == lLastRow );
		int iRowRet = GetRowYCoord(lFirstRow, &tm, &ptRow, &ptRowVis, lFirstRow == lLastRow);
		if (lFirstRow != lLastRow)
		{
			POINT pt, ptVis;
			//int iRowRet2 = MTblGetRowYCoord( m_hWndTbl, lLastRow, &tm, &pt, &ptVis, FALSE );
			int iRowRet2 = GetRowYCoord(lLastRow, &tm, &pt, &ptVis, FALSE);

			ptRow.y = pt.y;

			long lRowsTop = (lFirstRow < 0) ? tm.m_SplitRowsTop : tm.m_RowsTop;
			long lRowsBottom = (lFirstRow < 0) ? tm.m_SplitRowsBottom : tm.m_RowsBottom;

			if (iRowRet == MTGR_RET_INVISIBLE)
				ptRowVis.x = lRowsTop;

			if (iRowRet2 == MTGR_RET_INVISIBLE)
				ptRowVis.y = lRowsBottom - 1;
			else
				ptRowVis.y = ptVis.y;
		}

		// Effektive Ausrichtungen der Zelle ermittln
		int iJustify = GetEffCellTextJustify(lFirstRow, hWndFirstCol);
		int iVAlign = GetEffCellTextVAlign(lFirstRow, hWndFirstCol);

		// Text-Rechteck ermitteln
		//BOOL bWordWrap = tm.m_LinesPerRow > 1 && SendMessage( hWndFirstCol, TIM_QueryColumnFlags, 0, COL_MultilineCell );
		BOOL bWordWrap = GetCellWordWrap(hWndFirstCol, lFirstRow);
		BOOL bEndEllipsis = MustUseEndEllipsis(hWndFirstCol);

		tstring sText;
		GetEffCellText(lFirstRow, hWndFirstCol, sText);
		LPCTSTR lpsText = sText.empty() ? _T("A") : sText.c_str();

		SIZE si;
		if (!GetCellTextExtent(hWndFirstCol, lFirstRow, &si, NULL, bWordWrap, FALSE, bEndEllipsis, FALSE, lpsText))
			return FALSE;

		BOOL bCalcBtn = bButton && !(bCheckbox && iJustify == DT_CENTER);
		BOOL bCalcDDBtn = bDDButton;

		int iTextAreaLeft = GetCellTextAreaLeft(lFirstRow, hWndFirstCol, ptCol.x, tm.m_CellLeadingX, Img, bCalcBtn);
		int iTextAreaRight = GetCellTextAreaRight(lFirstRow, hWndFirstCol, ptCol.y, tm.m_CellLeadingX, Img, bCalcBtn, bCalcDDBtn);

		RECT rText;
		CalcCellTextRect(ptRow.x, ptRow.y, iTextAreaLeft, iTextAreaRight, tm.m_CellLeadingX, tm.m_CellLeadingY, iJustify, iVAlign, si, &rText);

		long lTextTop = rText.top;

		// Tree-Spalte?
		BOOL bTreeCol = (hWndFirstCol == m_hWndTreeCol && m_hWndTreeCol);

		// Über Knoten?
		if (bTreeCol && m_iNodeSize > 0 && m_Rows->QueryRowFlags(lFirstRow, MTBL_ROW_CANEXPAND) && !(m_Rows->GetRowLevel(lFirstRow) == 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE)))
		{
			// Ebene der Zeile ermitteln
			int iRowLevel = m_Rows->GetRowLevel(lFirstRow);

			// Rechteck des Knotens ermitteln
			RECT rNode;
			SetRect(&rNode, 0, 0, 0, 0);
			if (GetNodeRect(ptCol.x, lTextTop, tmCellFont.tmHeight, iRowLevel, tm.m_CellLeadingX, tm.m_CellLeadingY, &rNode))
			{
				rNode.right += 1;
				rNode.bottom += 1;

				// Punkt im Knoten?
				if (PtInRect(&rNode, pt))
				{
					*pdwMFlags |= MTOFP_OVERNODE;
				}
			}
		}

		if (!(*pdwMFlags & MTOFP_OVERNODE))
		{
			// Über Zellentext bzw. Checkbox?
			BOOL bHiddenText = (pCell && pCell->QueryFlags(MTBL_CELL_FLAG_HIDDEN_TEXT));
			if (!bHiddenText && !sText.empty())
			{
				RECT rVis;
				if (bCheckbox)
				{
					RECT rCbArea = { iTextAreaLeft, ptRow.x, iTextAreaRight, ptRow.y };
					RECT rCb;
					CalcCheckBoxRect(iJustify, iVAlign, rCbArea, tm.m_CellLeadingX, tm.m_CellLeadingY, tm.m_CharBoxY, rCb);

					rVis.left = max(rCb.left, max(ptColVis.x, iTextAreaLeft));
					rVis.top = max(rCb.top, ptRowVis.x);
					rVis.right = min(rCb.right, ptColVis.y);
					rVis.bottom = min(rCb.bottom, ptRowVis.y);

				}
				else
				{
					rVis.left = max(rText.left, max(ptColVis.x, iTextAreaLeft));
					rVis.top = max(rText.top, ptRowVis.x);
					rVis.right = min(rText.right, min(ptColVis.y, iTextAreaRight));
					rVis.bottom = min(rText.bottom, ptRowVis.y);
				}

				if (PtInRect(&rVis, pt))
					*pdwMFlags |= MTOFP_OVERTEXT;
			}

			// Über Bild?
			if (Img.GetHandle())
			{
				// Rechteck des Bildes ermitteln
				RECT rImg;
				SetRect(&rImg, 0, 0, 0, 0);
				long lHt = ptRow.y - ptRow.x, lWid = ptCol.y - ptCol.x;
				if (GetImgRect(Img, lFirstRow, hWndFirstCol, ptCol.x, ptRow.x, lWid, lHt, lTextTop, tmCellFont.tmHeight, tm.m_CellLeadingX, tm.m_CellLeadingY, bButton, bDDButton, &rImg))
				{
					// Punkt im Bild?
					if (PtInRect(&rImg, pt))
						*pdwMFlags |= MTOFP_OVERIMAGE;
				}
			}

			// Buttons
			RECT rBounds = { ptCol.x, ptRow.x, ptCol.y, ptRow.y };
			RECT rDDBtn = { 0, 0, 0, 0 };

			// Über Drop-Down-Button?
			if (bDDButton)
			{
				if (CalcDDBtnRect(hWndFirstCol, lFirstRow, &tm, &rBounds, NULL, NULL, &rDDBtn))
				{
					if (PtInRect(&rDDBtn, pt))
						*pdwMFlags |= MTOFP_OVERDDBTN;
				}
			}

			// Über Button?
			if (bButton)
			{
				if (bDDButton && !IsRectEmpty(&rDDBtn))
				{
					rBounds.right = rDDBtn.left;
				}

				RECT rBtn;
				CMTblBtn *pBtn = GetEffCellBtn(hWndFirstCol, lFirstRow);
				if (CalcBtnRect(pBtn, hWndFirstCol, lFirstRow, &tm, &rBounds, NULL, NULL, &rBtn))
				{
					if (PtInRect(&rBtn, pt))
						*pdwMFlags |= MTOFP_OVERBTN;
				}
			}
		}
	}
	else
	{
		// Über Corner-Begrenzung ( vertikal )?
		if (*pdwFlags & TBL_YOverColumnHeader)
		{
			if (IsOverCornerSep(pt.x, pt.y))
				*pdwMFlags |= MTOFP_OVERCORNERSEP;
		}

		// Über Spaltenbegrenzung?
		if (*pdwFlags & TBL_YOverColumnHeader)
		{
			HWND hWndSepCol = GetSepCol(pt.x, pt.y);
			if (hWndSepCol)
				*pdwMFlags |= MTOFP_OVERCOLHDRSEP;
		}

		// Über Spalten-Header?
		BOOL bOverColHdr = (*phWndCol && (*pdwFlags & TBL_YOverColumnHeader));
		if (bOverColHdr)
		{
			// Über Spalten-Header-Gruppe?
			CMTblColHdrGrp *pGrp = m_ColHdrGrps->GetGrpOfCol(*phWndCol);
			if (pGrp && pt.y <= GetColHdrGrpHeight(pGrp))
			{
				*pdwMFlags |= MTOFP_OVERCOLHDRGRP;
			}
			else
			{
				// Über Button?
				CMTblBtn* pBtn = m_Cols->GetHdrBtn(*phWndCol);
				if (pBtn && pBtn->m_bShow)
				{
					RECT rBounds;
					GetColHdrRect(*phWndCol, rBounds, MTGR_EXCLUDE_COLHDRGRP);
					RECT rBtn;
					if (CalcBtnRect(pBtn, *phWndCol, TBL_Error, &tm, &rBounds, NULL, NULL, &rBtn))
					{
						if (PtInRect(&rBtn, pt))
							*pdwMFlags |= MTOFP_OVERBTN;
					}
				}

				// Über Bild?
				if (!(*pdwMFlags & MTOFP_OVERBTN))
				{
					CMTblImg Img = GetEffColHdrImage(*phWndCol, GEI_HILI);
					if (Img.GetHandle())
					{
						// Rechteck des Spalten-Headers ermitteln
						RECT rBounds;
						GetColHdrRect(*phWndCol, rBounds, MTGR_EXCLUDE_COLHDRGRP);

						// Metriken des Fonts ermitteln
						TEXTMETRIC tmFont;
						GetColHdrFontMetrics(*phWndCol, &tmFont);

						// Button vorhanden?
						CMTblBtn* pBtn = m_Cols->GetHdrBtn(*phWndCol);
						BOOL bButton = (pBtn && pBtn->m_bShow);

						// Rechteck des Bildes ermitteln
						RECT rImg;
						SetRect(&rImg, 0, 0, 0, 0);
						if (GetImgRect(Img, TBL_Error, *phWndCol, rBounds.left, rBounds.top, rBounds.right - rBounds.left, rBounds.bottom - rBounds.top, rBounds.top + tm.m_CellLeadingY, tmFont.tmHeight, tm.m_CellLeadingX, tm.m_CellLeadingY, bButton, FALSE, &rImg))
						{
							// Punkt im Bild?
							if (PtInRect(&rImg, pt))
								*pdwMFlags |= MTOFP_OVERIMAGE;
						}
					}
				}
			}

		}
		else
		{
			// Über Zeilenbegrenzung?
			if ((*pdwFlags & TBL_XOverRowHeader) && !(*pdwFlags & TBL_YOverColumnHeader) && !(*pdwFlags & TBL_YOverSplitBar))
			{
				long lSepRow = GetSepRow(pt);
				if (lSepRow != TBL_Error)
				{
					LPMTBLMERGEROW pmr = FindMergeRow(lSepRow, FMR_ANY);
					if (!pmr || pmr->lRowTo == lSepRow)
						*pdwMFlags |= MTOFP_OVERROWHDRSEP;
					else if (pmr && pmr->lRowTo != lSepRow)
						*pdwMFlags |= MTOFP_OVERMERGEDROWHDRSEP;
				}
			}

			BOOL bOverRowHdr = (*plRow != TBL_Error && (*pdwFlags & TBL_XOverRowHeader) && !(*pdwFlags & TBL_YOverSplitBar));
			if (bOverRowHdr)
			{
				// Effektive Zeile ermitteln
				long lRow;
				LPMTBLMERGEROW pmr = FindMergeRow(*plRow, FMR_ANY);
				if (pmr)
					lRow = pmr->lRowFrom;
				else
					lRow = *plRow;

				// Über Knoten?
				if (m_iNodeSize > 0 && m_Rows->QueryRowFlags(lRow, MTBL_ROW_CANEXPAND))
				{
					HSTRING hsTitle = 0;
					HWND hWndCol = NULL;
					int iWid = 0;
					WORD wFlags = 0;
					if (Ctd.TblQueryRowHeader(m_hWndTbl, &hsTitle, 0, &iWid, &wFlags, &hWndCol))
					{
						if (hWndCol && hWndCol == m_hWndTreeCol && !(wFlags & TBL_RowHdr_MarkEdits))
						{
							// Zellenfont-Metriken ermitteln
							TEXTMETRIC tmCellFont;
							if (!GetCellFontMetrics(lRow, hWndCol, &tmCellFont))
								return FALSE;

							// Bild ermitteln
							CMTblImg Img = GetEffCellImage(lRow, hWndCol, GEI_HILI);

							// Koordinaten ermitteln
							POINT ptRow, ptRowVis;
							//MTblGetRowYCoord( m_hWndTbl, lRow, &tm, &ptRow, &ptRowVis);
							GetRowYCoord(lRow, &tm, &ptRow, &ptRowVis);

							// Text-Rechteck ermitteln
							//BOOL bWordWrap = tm.m_LinesPerRow > 1 && SendMessage( hWndCol, TIM_QueryColumnFlags, 0, COL_MultilineCell );
							BOOL bWordWrap = GetCellWordWrap(hWndCol, lRow);

							int iTextAreaLeft = GetCellTextAreaLeft(lRow, hWndCol, 0, tm.m_CellLeadingX, Img);
							int iTextAreaRight = GetCellTextAreaRight(lRow, hWndCol, iWid, tm.m_CellLeadingX, Img);
							int iJustify = GetEffCellTextJustify(lRow, hWndCol);

							POINT ptText;
							CalcCellTextAreaXCoord(iTextAreaLeft, iTextAreaRight, tm.m_CellLeadingX, iJustify, bWordWrap, ptText);

							tstring sText;
							GetEffCellText(lRow, hWndCol, sText);
							LPCTSTR lpsText = sText.empty() ? _T("A") : sText.c_str();

							BOOL bEndEllipsis = MustUseEndEllipsis(hWndCol);
							SIZE si;
							if (!GetCellTextExtent(hWndCol, lRow, &si, NULL, bWordWrap, FALSE, bEndEllipsis, FALSE, lpsText, ptText.y - ptText.x))
								return FALSE;

							int iVAlign = GetEffCellTextVAlign(lRow, hWndCol);
							RECT rText;
							CalcCellTextRect(ptRow.x, ptRow.y, iTextAreaLeft, iTextAreaRight, tm.m_CellLeadingX, tm.m_CellLeadingY, iJustify, iVAlign, si, &rText);

							long lTextTop = rText.top;

							// Ebene der Zeile ermitteln
							int iRowLevel = m_Rows->GetRowLevel(lRow);

							// Rechteck des Knotens ermitteln
							RECT rNode;
							SetRect(&rNode, 0, 0, 0, 0);
							if (GetNodeRect(0, lTextTop, tmCellFont.tmHeight, iRowLevel, tm.m_CellLeadingX, tm.m_CellLeadingY, &rNode))
							{
								rNode.right += 1;
								rNode.bottom += 1;

								// Punkt im Knoten?
								if (PtInRect(&rNode, pt))
								{
									*pdwMFlags |= MTOFP_OVERNODE;
								}
							}
						}
					}
					if (hsTitle)
						Ctd.HStringUnRef(hsTitle);
				}
			}
		}
	}

	return TRUE;
}


// ODIPrint
void MTBLSUBCLASS::ODIPrint(LPMTBLODI pODI)
{
	m_pPrODI = pODI;
	SendMessage(m_hWndTbl, MTM_DrawItem, m_pPrODI->lType, (LPARAM)m_pPrODI);
	m_pPrODI = NULL;
}

// OnBtnCaptureChanged
LRESULT MTBLSUBCLASS::OnBtnCaptureChanged(HWND hWndBtn, WPARAM wParam, LPARAM lParam)
{
	LPMTBLBTN pb = LPMTBLBTN(GetWindowLongPtr(hWndBtn, GWLP_USERDATA));
	if (pb)
	{
		pb->bPushed = FALSE;
		pb->bCaptured = FALSE;
		InvalidateRect(hWndBtn, NULL, FALSE);
	}

	return 0;
}


// OnBtnDestroy
LRESULT MTBLSUBCLASS::OnBtnDestroy(HWND hWndBtn, WPARAM wParam, LPARAM lParam)
{
	LPMTBLBTN pb = LPMTBLBTN(GetWindowLongPtr(hWndBtn, GWLP_USERDATA));
	if (pb)
	{
		LPMTBLSUBCLASS psc = MTblGetSubClass(pb->hWndTbl);
		if (psc)
			psc->m_hWndBtn = NULL;

		delete pb;
		SetWindowLongPtr(hWndBtn, GWLP_USERDATA, 0);
	}

	return 0;
}


// OnBtnLButtonDown
LRESULT MTBLSUBCLASS::OnBtnLButtonDown(HWND hWndBtn, WPARAM wParam, LPARAM lParam)
{
	LPMTBLBTN pb = LPMTBLBTN(GetWindowLongPtr(hWndBtn, GWLP_USERDATA));
	if (pb)
	{
		RECT r;
		GetWindowRect(hWndBtn, &r);
		r.bottom = r.top + pb->iHeight;

		POINT pt;
		GetCursorPos(&pt);

		if (PtInRect(&r, pt))
		{
			SetCapture(hWndBtn);
			pb->bCaptured = TRUE;

			pb->bPushed = TRUE;
			InvalidateRect(hWndBtn, NULL, FALSE);
		}
	}

	return 0;
}


// OnBtnLButtonUp
LRESULT MTBLSUBCLASS::OnBtnLButtonUp(HWND hWndBtn, WPARAM wParam, LPARAM lParam)
{
	LPMTBLBTN pb = LPMTBLBTN(GetWindowLongPtr(hWndBtn, GWLP_USERDATA));
	if (pb)
	{
		if (pb->bCaptured)
		{
			ReleaseCapture();

			POINT pt;
			GetCursorPos(&pt);

			if (WindowFromPoint(pt) == hWndBtn)
			{
				RECT r;
				GetWindowRect(hWndBtn, &r);
				r.bottom = r.top + pb->iHeight;

				if (PtInRect(&r, pt))
				{
					HWND hWndCol = NULL;
					long lRow = TBL_Error;
					Ctd.TblQueryFocus(pb->hWndTbl, &lRow, &hWndCol);
					PostMessage(pb->hWndTbl, MTM_BtnClick, WPARAM(hWndCol), lRow);
				}
			}
		}
	}

	return 0;
}


// OnBtnMouseMove
LRESULT MTBLSUBCLASS::OnBtnMouseMove(HWND hWndBtn, WPARAM wParam, LPARAM lParam)
{
	LPMTBLBTN pb = LPMTBLBTN(GetWindowLongPtr(hWndBtn, GWLP_USERDATA));
	if (pb)
	{
		if (pb->bCaptured)
		{
			POINT pt;
			GetCursorPos(&pt);
			BOOL bOver = FALSE;
			if (WindowFromPoint(pt) == hWndBtn)
			{
				RECT r;
				GetWindowRect(hWndBtn, &r);
				r.bottom = r.top + pb->iHeight;

				if (PtInRect(&r, pt))
					bOver = TRUE;
			}

			if (bOver && !pb->bPushed)
			{
				pb->bPushed = TRUE;
				InvalidateRect(hWndBtn, NULL, FALSE);
			}
			else if (!bOver && pb->bPushed)
			{
				pb->bPushed = FALSE;
				InvalidateRect(hWndBtn, NULL, FALSE);
			}
		}

		LPMTBLSUBCLASS psc = MTblGetSubClass(pb->hWndTbl);
		if (psc && !psc->m_uiTimer)
		{
			psc->StartInternalTimer(TRUE);
		}
	}

	return 0;
}


// OnBtnPaint
LRESULT MTBLSUBCLASS::OnBtnPaint(HWND hWndBtn, WPARAM wParam, LPARAM lParam)
{
	LPMTBLBTN pb = LPMTBLBTN(GetWindowLongPtr(hWndBtn, GWLP_USERDATA));
	if (pb)
	{
		// Tabellen-Subclass ermitteln
		LPMTBLSUBCLASS psc = MTblGetSubClass(pb->hWndTbl);
		if (!psc) return 0;

		// Zeichnen beginnen
		PAINTSTRUCT ps;
		BeginPaint(hWndBtn, &ps);
		HDC hDC = ps.hdc;

		// Rechteck ermitteln
		RECT r;
		GetWindowRect(hWndBtn, &r);
		ScreenToClient(hWndBtn, LPPOINT(&r));
		ScreenToClient(hWndBtn, LPPOINT(&r) + 1);

		// Stil festlegen
		int iStyle = DFCS_BUTTONPUSH;
		if (pb->bPushed)
			iStyle |= DFCS_PUSHED;

		// Hintergrund füllen
		RECT rBack;
		CopyRect(&rBack, &r);
		rBack.right += 1;
		rBack.bottom += 1;

		HWND hWndCol;
		long lRow;
		Ctd.TblQueryFocus(pb->hWndTbl, &lRow, &hWndCol);
		COLORREF crBackColor = psc->GetDrawColor(psc->GetEffCellBackColor(lRow, hWndCol));

		HBRUSH hBrBack = CreateSolidBrush(crBackColor);
		FillRect(hDC, &rBack, hBrBack);
		DeleteObject(hBrBack);

		BOOL bMustDeleteFont = FALSE;
		//HFONT hBtnFont = (pb->dwFlags & MTBL_BTN_SHARE_FONT) ? psc->m_hEditFont : (HFONT)SendMessage( pb->hWndTbl, WM_GETFONT, 0, 0 );
		HFONT hBtnFont = (pb->dwFlags & MTBL_BTN_SHARE_FONT) ? psc->GetEffCellFont(hDC, lRow, hWndCol, &bMustDeleteFont) : (HFONT)SendMessage(pb->hWndTbl, WM_GETFONT, 0, 0);

		LPTSTR lpsTxt = NULL;
		int iTxtLen = GetWindowTextLength(hWndBtn);
		if (iTxtLen > 0)
		{
			lpsTxt = new (nothrow)TCHAR[iTxtLen + 1];
			if (lpsTxt)
			{
				GetWindowText(hWndBtn, lpsTxt, iTxtLen + 1);
			}
		}

		// Button zeichnen
		MTBLPAINTBTN b;

		b.hDC = hDC;
		b.hFont = hBtnFont;
		b.hImage = pb->hImage;
		b.lpsText = lpsTxt;
		b.dwFlags = pb->dwFlags;
		b.bPushed = pb->bPushed;
		b.bDrawInverted = FALSE;
		b.bRestoreClipRgn = FALSE;

		CopyRect(&b.rBtn, &r);
		b.rBtn.bottom = b.rBtn.top + pb->iHeight;

		psc->PaintBtn(&b);

		if (bMustDeleteFont)
			DeleteObject(hBtnFont);

		// Zeichnen beenden
		EndPaint(hWndBtn, &ps);
	}

	return 0;
}


// OnCaptureChanged
LRESULT MTBLSUBCLASS::OnCaptureChanged(WPARAM wParam, LPARAM lParam)
{
	if (GetCapture() != m_hWndTbl)
	{
		// Handling aktuell gedrückter permanenter Button
		if (m_hWndColBtnPushed && m_lRowBtnPushed != TBL_Error)
		{
			CMTblRow *pRow = m_Rows->GetRowPtr(m_lRowBtnPushed);
			if (pRow)
			{
				pRow->m_Cols->SetBtnPushed(m_hWndColBtnPushed, FALSE);
				InvalidateBtn(m_hWndColBtnPushed, m_lRowBtnPushed, NULL, FALSE);

				m_hWndColBtnPushed = NULL;
				m_lRowBtnPushed = TBL_Error;
			}
		}
		else if (m_hWndColBtnPushed && m_lRowBtnPushed == TBL_Error)
		{
			m_Cols->SetHdrBtnPushed(m_hWndColBtnPushed, FALSE);
			InvalidateBtn(m_hWndColBtnPushed, TBL_Error, NULL, FALSE);
			m_hWndColBtnPushed = NULL;
		}
	}

	return CallOldWndProc(WM_CAPTURECHANGED, wParam, lParam);
}


// OnChar
LRESULT MTBLSUBCLASS::OnChar(WPARAM wParam, LPARAM lParam)
{
	long lRet = 0;

	if (SendMessage(m_hWndTbl, MTM_Char, wParam, lParam) != MTBL_NODEFAULTACTION)
		lRet = CallOldWndProc(WM_CHAR, wParam, lParam);

	// Tastenkombination kopieren? ( [STRG] + C )
	if (wParam == 3)
	{
		SHORT shState = GetKeyState(VK_CONTROL);
		if (shState & 0x8000)
			CopyShortcutPressed();
	}

	return lRet;
}


// OnClearSelection
LRESULT MTBLSUBCLASS::OnClearSelection(WPARAM wParam, LPARAM lParam)
{
	BOOL bTrapRowSelChanges = MustTrapRowSelChanges();
	if (bTrapRowSelChanges)
		BeginRowSelTrap();

	long lRet = CallOldWndProc(TIM_ClearSelection, wParam, lParam);

	if (bTrapRowSelChanges)
		EndRowSelTrap();

	return lRet;
}


// OnCommand
LRESULT MTBLSUBCLASS::OnCommand(WPARAM wParam, LPARAM lParam)
{
	BOOL bOldProc = TRUE;
	HWND hWndCtl = (HWND)lParam;
	long lRet = 0;
	WORD wNotifyCode = HIWORD(wParam);
	WORD wID = LOWORD(wParam);

	// Zelle geändert?
	if (hWndCtl && wNotifyCode == EN_CHANGE)
	{
		// Subklasse der Tabelle ermitteln
		HWND hWnd = hWndCtl;
		LPMTBLSUBCLASS psc;
		while (hWnd)
		{
			psc = MTblGetSubClass(hWnd);
			if (psc)
				break;

			hWnd = GetParent(hWnd);
		}

		// Subklasse gefunden..
		if (psc)
		{
			// Fokus ermitteln
			HWND hWndFocusCol = NULL;
			long lFocusRow = TBL_Error;
			Ctd.TblQueryFocus(psc->m_hWndTbl, &lFocusRow, &hWndFocusCol);

			//Wenn gültiger Fokus...
			if (lFocusRow != TBL_Error)
			{
				// Zeile eines Zellenverbundes?
				LPMTBLMERGEROW pmr = FindMergeRow(lFocusRow, FMR_ANY);
				if (pmr)
				{
					// ROW_Edited vor und nach Standardverarbeitung ermitteln
					BOOL bRowEditedBefore = Ctd.TblQueryRowFlags(psc->m_hWndTbl, lFocusRow, ROW_Edited);
					lRet = CallOldWndProc(WM_COMMAND, wParam, lParam);
					bOldProc = FALSE;
					BOOL bRowEditedAfter = Ctd.TblQueryRowFlags(psc->m_hWndTbl, lFocusRow, ROW_Edited);

					// Wenn sich ROW_Edited gesetzt wurde...
					if (bRowEditedBefore != bRowEditedAfter && bRowEditedAfter)
					{
						// ROW_Edited nochmal explizit setzen, damit bei allen Zellen des Verbundes ROW_Edited gesetzt wird
						Ctd.TblSetRowFlags(psc->m_hWndTbl, lFocusRow, ROW_Edited, TRUE);
					}
				}
			}
		}
	}

	if (bOldProc)
		lRet = CallOldWndProc(WM_COMMAND, wParam, lParam);

	return lRet;
}


// OnContainerParentNotify
LRESULT MTBLSUBCLASS::OnContainerParentNotify(HWND hWndTbl, HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	int iEvent = int(LOWORD(wParam));

	if (iEvent == WM_CREATE || iEvent == WM_DESTROY)
	{
		TCHAR szBuf[255];
		HWND hWndChild = HWND(lParam);

		if (GetClassName(hWndChild, szBuf, (sizeof(szBuf) / sizeof(TCHAR)) - 1))
		{
			BOOL bButton = !lstrcmp(szBuf, _T("Button"));

			if (iEvent == WM_CREATE)
			{
				if (bButton)
				{
					m_hWndCheckBox = hWndChild;
					PostMessage(m_hWndTbl, MTM_Internal, IA_SUBCLASS_LISTCLIP_CTRL, (LPARAM)hWndChild);
				}
			}
			else if (iEvent == WM_DESTROY)
			{
				if (bButton)
				{
					m_hWndCheckBox = NULL;
				}
			}
		}
	}

	return 0;
}


// OnCtlColorEdit
LRESULT MTBLSUBCLASS::OnCtlColorEdit(WPARAM wParam, LPARAM lParam)
{
	HDC hDC = HDC(wParam);
	HWND hWnd = HWND(lParam);

	if (hDC && hWnd == m_hWndEdit)
	{
		// Textfarbe setzen
		SetTextColor(hDC, m_clrEditText);

		// Hintergrundfarbe setzen
		SetBkColor(hDC, m_clrEditBack);
	}

	return LRESULT(m_hEditBrush);
}


// OnDefineSplitWindow
LRESULT MTBLSUBCLASS::OnDefineSplitWindow(WPARAM wParam, LPARAM lParam)
{
	m_bDefiningSplitWindow = TRUE;
	long lRet = CallOldWndProc(TIM_DefineSplitWindow, wParam, lParam);
	m_bDefiningSplitWindow = FALSE;
	if (lRet)
	{
		// Internes Flag setzen
		m_bSplitted = (lParam > 0);

		// Ggf. vertikale Scrollbar verbergen/anzeigen
		if (m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
		{
			if (m_bSplitted)
				FlatSB_SetScrollRange_Hooked(m_hWndTbl, SB_VERT, 0, 0, TRUE);
			/*else
			{
			CMTblMetrics tm( m_hWndTbl );
			int iScrPos = GetScrollPosV( &tm, FALSE );
			long lRow = m_Rows->GetVisPosRow( iScrPos, FALSE );
			if ( lRow != TBL_Error )
			Ctd.TblScroll( m_hWndTbl, lRow, NULL, TBL_ScrollTop );
			}*/
		}

		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		// Focus wieder zeichnen?
		HWND hWndCol;
		long lFocusRow;
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndCol);

		long lMinVisRow, lMaxVisRow;
		Ctd.TblQueryVisibleRange(m_hWndTbl, &lMinVisRow, &lMaxVisRow);
		if (lFocusRow != lMinVisRow)
		{
			m_bNoFocusPaint = FALSE;
		}
		else
		{
			if (lParam > 0 && IsMyWindow(GetFocus()))
			{
				//m_bNoFocusPaint = TRUE;
			}
			else if (lParam == 0)
			{
				m_bNoFocusPaint = FALSE;
			}
		}

		if (MustSuppressRowFocusCTD())
			//Ctd.TblKillFocus( m_hWndTbl );
			KillFocusFrameI();
	}

	return lRet;
}


// OnDeleteRow
LRESULT MTBLSUBCLASS::OnDeleteRow(WPARAM wParam, LPARAM lParam)
{
	// Focus ermitteln
	HWND hWndCol = NULL;
	long lFocusRow = 0;
	Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndCol);

	BOOL bRowHeights = m_pCounter->GetRowHeights() > 0;
	BOOL bNoColHdr = QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	// Zu löschende Zeile...
	long lRowToDelete = *((LPLONG)lParam);
	CMTblRow *pRowToDelete = m_Rows->EnsureValid(lRowToDelete);
	if (!pRowToDelete)
		return 0;

	// Ggf. Hirarchie normalisieren
	BOOL bNormalized = FALSE;
	if (QueryTreeFlags(MTBL_TREE_FLAG_AUTO_NORM_HIER))
	{
		if (SendMessage(m_hWndTbl, MTM_QueryAutoNormHierarchy, lRowToDelete, NORM_HIER_ROW_DELETE) != MTBL_NODEFAULTACTION)
		{
			if (pRowToDelete->QueryFlags(MTBL_ROW_CANEXPAND) && !pRowToDelete->IsParent())
			{
				if (m_lLoadChildRows > 0)
				{
					DWORD dwIntFlags = ROW_IFLAG_LOAD_CHILD_ROWS | (wParam == TBL_Adjust ? ROW_IFLAG_DELETE_ADJUST : ROW_IFLAG_DELETE);
					pRowToDelete->SetInternalFlags(dwIntFlags, TRUE);
					return TRUE;
				}
				else
				{
					LoadChildRows(lRowToDelete);
					lRowToDelete = pRowToDelete->GetNr();
				}
			}

			BOOL bNormalizeHierarchy = pRowToDelete->IsParent();
			if (bNormalizeHierarchy)
				bNormalized = NormalizeHierarchy(lRowToDelete, NORM_HIER_ROW_DELETE);
		}
	}

	// Wenn notwendig verbergen, da sonst ggf. GPF in tablixx.dll
	BOOL bHiddenBefore = FALSE;
	if ((bRowHeights || bNoColHdr) && !Ctd.TblQueryRowFlags(m_hWndTbl, lRowToDelete, ROW_Hidden))
		bHiddenBefore = Ctd.TblSetRowFlags(m_hWndTbl, lRowToDelete, ROW_Hidden, TRUE);

	// Alte Fensterprozedur aufrufen
	m_pCounter->IncNoFocusUpdate();
	long lRet = CallOldWndProc(TIM_DeleteRow, wParam, lParam);
	m_pCounter->DecNoFocusUpdate();

	if (lRet)
	{
		long lDeletedRow = *((LPLONG)lParam);
		CMTblRow *pDeletedRow = m_Rows->GetRowPtr(lDeletedRow);

		// Evtl. Bereich Zellenselektion/Tastatur zurücksetzen
		if (lDeletedRow == m_CellSelRangeKey.lRowFrom)
			m_CellSelRangeKey.SetEmpty();

		// Focus wieder zeichnen?
		if (IsMyWindow(GetFocus()))
		{
			m_bNoFocusPaint = FALSE;
		}
		else
		{
			if (lFocusRow == lDeletedRow)
			{
				m_bNoFocusPaint = FALSE;
			}
			else
			{
				long lMinVisRow, lMaxVisRow;
				Ctd.TblQueryVisibleRange(m_hWndTbl, &lMinVisRow, &lMaxVisRow);
				if (lFocusRow != lMinVisRow)
				{
					m_bNoFocusPaint = FALSE;
				}
			}
		}

		if (MustSuppressRowFocusCTD())
			KillFocusFrameI();

		if (wParam == TBL_Adjust)
		{
			// Liste der mit TBL_Adjust gelöschten Zeilen erweitern
			m_DelAdjRows->push_back(lDeletedRow);
		}

		// Elternzeile ermitteln
		long lParentRow = m_Rows->GetParentRow(lDeletedRow);

		// Gelöschte Zeile = Zeile Fokus-Zelle?
		if (pDeletedRow == m_pRowFocus)
			m_pRowFocus = NULL;

		// Ggf. Enter/Leave-Zeilen zurücksetzen
		if (pDeletedRow == m_pELRow)
			m_pELRow = NULL;
		if (pDeletedRow == m_pELRow_Hdr)
			m_pELRow_Hdr = NULL;
		if (pDeletedRow == m_pELRow_Cell)
			m_pELRow_Cell = NULL;
		if (pDeletedRow == m_pELRow_Btn)
			m_pELRow_Btn = NULL;
		if (pDeletedRow == m_pELRow_DDBtn)
			m_pELRow_DDBtn = NULL;

		// Ggf. hervorgehobene Items in Zeile löschen
		ItemList::iterator it;
		for (it = m_pHLIL->begin(); it != m_pHLIL->end(); it++)
		{
			if ((*it).Id == pDeletedRow)
			{
				it = m_pHLIL->erase(it);
				it--;
			}
		}

		// Zeile löschen
		m_Rows->DeleteRow(lDeletedRow, wParam);

		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		// Merge-Zeilen ungültig
		m_bMergeRowsInvalid = TRUE;

		// Elternzeile neu zeichnen, wenn sie jetzt keine Kinder mehr hat
		if (lParentRow != TBL_Error && !m_Rows->IsParentRow(lParentRow))
			InvalidateRow(lParentRow, FALSE, TRUE);

		// Gelöschte Zeile Teil eines Zeilen-Verbundes?
		LPMTBLMERGEROW pmr = FindMergeRow(lDeletedRow, FMR_ANY);
		if (pmr)
		{
			// Zeilenflags replizieren, da jetzt eine andere Zeile Teil des Verbundes sein könnte
			if (pmr && !m_bReplMergeRowsFlags)
			{
				BOOL bSet;
				DWORD dwFlagsCheck[] = { ROW_Selected, ROW_New, ROW_Edited, ROW_MarkDeleted, ROW_HideMarks, ROW_UserFlag1, ROW_UserFlag2, ROW_UserFlag3, ROW_UserFlag4, ROW_UserFlag5 };

				m_bReplMergeRowsFlags = TRUE;

				int iMax = sizeof(dwFlagsCheck) / sizeof(DWORD);
				if (QueryFlags(MTBL_FLAG_NO_USER_ROW_FLAG_REPL))
					iMax -= 5;

				for (int i = 0; i < iMax; i++)
				{
					bSet = Ctd.TblQueryRowFlags(m_hWndTbl, pmr->lRowFrom, dwFlagsCheck[i]);
					SetMergeRowFlags(pmr, dwFlagsCheck[i], bSet, FALSE, pmr->lRowFrom);
				}

				m_bReplMergeRowsFlags = FALSE;
			}
		}

		if (!Ctd.TblAnyRows(m_hWndTbl, 0, 0))
		{
			// Liste der mit TBL_Adjust gelöschten Zeilen zurücksetzen
			m_DelAdjRows->clear();

			// Zeilen zurücksetzen
			m_Rows->Reset();
		}

		// Ggf. neu zeichnen
		if (bRowHeights || bNoColHdr || bNormalized)
			InvalidateRect(m_hWndTbl, NULL, FALSE);

		// Fokus-Zelle aktualisieren
		UpdateFocus();

		// Erweiterte Nachricht senden
		if (m_bExtMsgs)
			PostMessage(m_hWndTbl, MTM_RowDeleted, wParam, lDeletedRow);
	}
	else
	{
		// Hoppla, Zeile konnte nicht gelöscht werden
		if (bHiddenBefore)
			Ctd.TblSetRowFlags(m_hWndTbl, lRowToDelete, ROW_Hidden, FALSE);

		if (bNormalized)
		{
			NormalizeHierarchy(lRowToDelete, NORM_HIER_ROW_DELETE_FAIL);
			InvalidateRect(m_hWndTbl, NULL, FALSE);
		}
	}

	return lRet;
}


// OnEditModeChange
LRESULT MTBLSUBCLASS::OnEditModeChange(WPARAM wParam, LPARAM lParam)
{
	long lRet = CallOldWndProc(TIM_EditModeChange, wParam, lParam);

	HWND hWndCol = NULL;
	long lRow = TBL_Error;
	if (Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol) && hWndCol)
	{
		int iCellType = GetEffCellType(lRow, hWndCol);
		if (iCellType == COL_CellType_DropDownList && !MTblIsListOpen(m_hWndTbl, hWndCol))
			PostMessage(m_hWndTbl, MTM_Internal, IA_CHECK_LIST_OPEN, (LPARAM)hWndCol);
	}

	return lRet;
}


// OnEnterItem
void MTBLSUBCLASS::OnEnterItem()
{
	if (!m_pELItem) return;
	if (m_bExtMsgs)
		SendMessage(m_hWndTbl, MTM_MouseEnterItem, (WPARAM)m_pELItem->Type, (LPARAM)m_pELItem);
}


// OnLeaveItem
void MTBLSUBCLASS::OnLeaveItem()
{
	if (!m_pELItem) return;

	if (m_bExtMsgs)
		SendMessage(m_hWndTbl, MTM_MouseLeaveItem, (WPARAM)m_pELItem->Type, (LPARAM)m_pELItem);
}


// OnFetchRow
LRESULT MTBLSUBCLASS::OnFetchRow(WPARAM wParam, LPARAM lParam)
{
	BOOL bUpdateFocus = !Ctd.TblQueryRowFlags(m_hWndTbl, lParam, ROW_InCache) || ( m_bCellMode && !m_hWndColFocus );

	long lRet = CallOldWndProc(TIM_FetchRow, wParam, lParam);
	if (lRet)
	{
		// Gültige Zeile sicherstellen
		m_Rows->EnsureValid(lParam);

		// Fokus aktualisieren
		if (bUpdateFocus)
			UpdateFocus();
	}

	return lRet;
}

// OnGetDlgCode
LRESULT MTBLSUBCLASS::OnGetDlgCode(WPARAM wParam, LPARAM lParam)
{
	// Handling für Hotkeys permanente Buttons
	if (m_iBtnKey && m_bCellMode && lParam)
	{
		MSG *pMsg = (MSG*)lParam;
		if (pMsg->message == WM_KEYDOWN && pMsg->wParam == m_iBtnKey)
		{
			HWND hWndFocusColCM = NULL;
			long lFocusRowCM = TBL_Error;
			if (GetEffFocusCell(lFocusRowCM, hWndFocusColCM) && MustShowBtn(hWndFocusColCM, lFocusRowCM, TRUE, TRUE) && IsButtonHotkeyPressed())
			{
				CMTblBtn *pbtn = GetEffCellBtn(hWndFocusColCM, lFocusRowCM);
				if (!(pbtn && pbtn->IsDisabled()))
					return DLGC_WANTMESSAGE;
			}
		}
	}

	return CallOldWndProc(WM_GETDLGCODE, wParam, lParam);
}


// OnGetMetrics
LRESULT MTBLSUBCLASS::OnGetMetrics(WPARAM wParam, LPARAM lParam)
{
	long lRet = CallOldWndProc(TIM_GetMetrics, wParam, lParam);

	if (QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
	{
		LPMTBLGETMETRICS pm = (LPMTBLGETMETRICS)lParam;
		pm->iHeaderHeight = 0;
	}

	return lRet;
}


// OnInsertRow
LRESULT MTBLSUBCLASS::OnInsertRow(WPARAM wParam, LPARAM lParam)
{
	// Ist an der Position eine mit TBL_Adjust gelöschte Zeile?
	if (m_Rows->IsRowDelAdj(lParam))
	{
		// Einfügeposition anpassen
		long lInsRow = lParam;
		if (lInsRow >= 0)
		{
			if (FindNextRow(&lInsRow, 0, 0))
				lParam = lInsRow;
			else
				lParam = TBL_MaxRow;
		}
		else
		{
			if (FindNextRow(&lInsRow, 0, 0, TBL_MaxSplitRow))
				lParam = lInsRow;
			else
				lParam = TBL_MaxSplitRow;
		}
	}

	// Focus killen?
	HWND hWndCol = NULL;
	long lRow;
	Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol);
	if (hWndCol)
	{
		Ctd.TblKillEdit(m_hWndTbl);
	}

	BOOL bAdaptScrollRangeV = m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	// Alte Fensterprozedur aufrufen
	m_pCounter->IncNoFocusUpdate();
	if (bAdaptScrollRangeV)
		m_bNoVScrolling = TRUE;
	long lRet = CallOldWndProc(TIM_InsertRow, wParam, lParam);
	if (bAdaptScrollRangeV)
		m_bNoVScrolling = FALSE;
	m_pCounter->DecNoFocusUpdate();

	if (lRet >= TBL_MinSplitRow && lRet <= TBL_MaxRow)
	{
		// Zeile einfügen
		m_Rows->InsertRow(lRet);

		// Gültige Zeile sicherstellen
		m_Rows->EnsureValid(lRet);

		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		// Merge-Zeilen ungültig
		m_bMergeRowsInvalid = TRUE;

		// Vertikalen Scrollbereich anpassen
		if (bAdaptScrollRangeV)
			AdaptScrollRangeV(lRet < 0);

		// Erweiterte Nachricht senden
		if (m_bExtMsgs)
			PostMessage(m_hWndTbl, MTM_RowInserted, 0, lRet);

		// Focus zeichnen?
		HWND hWndFocusCol;
		long lFocusRow;
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);

		int iSplitRows = 0;
		long lMinVisRow, lMaxVisRow;
		/*if ( lRet >= 0 )
		Ctd.TblQueryVisibleRange( m_hWndTbl, &lMinVisRow, &lMaxVisRow );
		else
		{
		MTblQueryVisibleSplitRange( m_hWndTbl, &lMinVisRow, &lMaxVisRow );
		}*/
		QueryVisRange(lMinVisRow, lMaxVisRow, lRet < 0);

		if ((lRet == lMinVisRow && lFocusRow == lMinVisRow && lParam != TBL_MaxRow) || (lRet < 0 && (lFocusRow < 0 || !IsSplitted())))
		{
			//m_bNoFocusPaint = TRUE;
		}
		else
		{
			if (lRet >= 0 && lRet < lFocusRow)
			{
				long lPos, lMinRow, lMaxRow;
				Ctd.TblQueryScroll(m_hWndTbl, &lPos, &lMinRow, &lMaxRow);
				if (lRet <= lMaxRow)
				{
					m_bNoFocusPaint = FALSE;
				}
			}
		}

		if (MustSuppressRowFocusCTD())
			//Ctd.TblKillFocus( m_hWndTbl );
			KillFocusFrameI();

		BOOL bTblInvalidated = FALSE;
		BOOL bRowInvalidated = FALSE;

		// Teil irgendeines Zellen-Verbundes?
		LPMTBLMERGECELL pmc = FindMergeCell(NULL, lRet, FMC_ANY);
		if (pmc)
		{
			// Ganze Tabelle neu zeichnen
			InvalidateRect(m_hWndTbl, NULL, FALSE);
			bTblInvalidated = TRUE;
		}

		// Teil eines Zeilen-Verbundes?
		LPMTBLMERGEROW pmr = FindMergeRow(lRet, FMR_ANY);
		if (pmr)
		{
			// Zeilenflags der Master-Zeile auf die neue Zeile replizieren
			if (pmr && !m_bReplMergeRowsFlags)
			{
				BOOL bSet;
				DWORD dwFlagsCheck[] = { ROW_Selected, ROW_New, ROW_Edited, ROW_MarkDeleted, ROW_HideMarks, ROW_UserFlag1, ROW_UserFlag2, ROW_UserFlag3, ROW_UserFlag4, ROW_UserFlag5 };

				m_bReplMergeRowsFlags = TRUE;
				int iMax = sizeof(dwFlagsCheck) / sizeof(DWORD);
				for (int i = 0; i < iMax; i++)
				{
					bSet = Ctd.TblQueryRowFlags(m_hWndTbl, pmr->lRowFrom, dwFlagsCheck[i]);
					Ctd.TblSetRowFlags(m_hWndTbl, lRet, dwFlagsCheck[i], bSet);
				}
				m_bReplMergeRowsFlags = FALSE;
			}

			if (!bTblInvalidated)
			{
				// Zeile inkl. aller verbundenen Zeilen neu zeichnen
				InvalidateRow(lRet, FALSE, TRUE);
				bRowInvalidated = TRUE;
			}
		}

		// Ggf. noch neu zeichnen
		if (!bTblInvalidated)
		{
			if (bAdaptScrollRangeV)
				// Von Zeile bis Ende des Zeilenbereiches neu zeichnen
				InvalidateRow(lRet, FALSE, FALSE, TRUE);
			else if (!bRowInvalidated)
				// Zeile neu zeichnen
				InvalidateRow(lRet, FALSE);
		}

		// Fokus
		UpdateFocus();
	}
	else
		lRet = lRet;

	return lRet;
}


// OnInternal
LRESULT MTBLSUBCLASS::OnInternal(WPARAM wParam, LPARAM lParam)
{
	HWND hWndCol;
	long lRow;

	switch (wParam)
	{
	case IA_KILL_EDIT_ESC:
		KillEditEsc((HWND)lParam);
		break;

	case IA_REMOVE_FOCUS_ROW_SELECTION:
		if (Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol) && Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Selected))
			Ctd.TblSetRowFlags(m_hWndTbl, lRow, ROW_Selected, FALSE);
		break;

	case IA_REMOVE_FOCUS:
		//Ctd.TblKillFocus( m_hWndTbl );
		KillFocusFrameI();
		break;

	case IA_SUBCLASS_LISTCLIP_CTRL:
		SubClassListClipCtrl((HWND)lParam);
		break;

	case IA_ADAPT_LISTCLIP:
		AdaptListClip();
		break;

	case IA_CHECK_LIST_OPEN:
		if (MTblIsListOpen(m_hWndTbl, (HWND)lParam))
			AdaptListClip();
		break;

	case IA_UPDATE_WINDOW:
		UpdateWindow(m_hWndTbl);
		break;

	case IA_AFTER_SCROLL:
		AfterScroll();
		break;

	case IA_GET_SPLITTED:
		return m_bSplitted ? 1 : -1;
		break;
	}
	return 0;
}

// OnInternalTimer
LRESULT MTBLSUBCLASS::OnInternalTimer(WPARAM wParam, LPARAM lParam)
{
	// Akt. Zeit ermitteln
	m_iiTimer = GetSystemTimeI64();

	// Alte Cursorposition merken
	POINT ptTimerPosOld = m_ptTimerPos;

	// Cursorposition ermitteln
	if (!GetCursorPos(&m_ptTimerPos)) return 0;
	POINT pt = m_ptTimerPos;
	if (!ScreenToClient(m_hWndTbl, &pt)) return 0;

	// Fenster unter Cursor ermitteln
	HWND hWndFromPoint = WindowFromPoint(m_ptTimerPos);
	BOOL bCursorOverTbl = (hWndFromPoint == m_hWndTbl);
	BOOL bCursorOverBtn = (hWndFromPoint == m_hWndBtn);
	BOOL bCursorOverDDBtn = IsDropDownBtn(hWndFromPoint);
	BOOL bCursorOverTip = (hWndFromPoint == m_hWndTip);
	if (bCursorOverTip && m_pTip && m_pTip->QueryFlags(MTBL_TIP_FLAG_PERMEABLE))
	{
		bCursorOverTip = FALSE;
		bCursorOverTbl = TRUE;
	}

	// WM_SETCURSOR fossieren
	/*if ( bCursorOverTbl && !m_bDragCanAutoStart )
	SetCursorPos( m_ptTimerPos.x, m_ptTimerPos.y );*/

	// Cursorposition gleich geblieben?
	BOOL bSameCursorPos = (ptTimerPosOld.x == m_ptTimerPos.x && ptTimerPosOld.y == m_ptTimerPos.y);

	// Focus ermitteln
	HWND hWndFocusCol = NULL;
	long lFocusRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);

	// Capture ermitteln
	HWND hWndCapture = GetCapture();

	// Alte Werte merken
	DWORD dwTimerFlagsOld = m_dwTimerFlags, dwTimerMFlagsOld = m_dwTimerMFlags;
	HWND hWndTimerColOld = m_hWndTimerCol;
	long lTimerRowOld = m_lTimerRow;

	// Werte ermitteln
	BOOL bReset = FALSE;
	if (hWndCapture)
		bReset = TRUE;
	else if (!(bCursorOverTbl && ObjFromPoint(pt, &m_lTimerRow, &m_hWndTimerCol, &m_dwTimerFlags, &m_dwTimerMFlags)))
		bReset = TRUE;
	else
	{
		if (hWndFocusCol && hWndFocusCol == m_hWndTimerCol && lFocusRow == m_lTimerRow)
			bReset = TRUE;
		if (m_dwTimerFlags & TBL_YOverSplitBar)
			bReset = TRUE;
	}

	if (bReset)
	{
		m_lTimerRow = TBL_Error;
		m_hWndTimerCol = NULL;
		m_dwTimerFlags = 0;
		m_dwTimerMFlags = 0;
	}

	// Analysieren und Aktionen durchführen

	// Merge-Ermittlung
	BOOL bMergeCell = FALSE;
	HWND hWndMergeCol = NULL, hWndLastMergedCol = NULL;
	long lMergeRow = TBL_Error, lLastMergedRow = TBL_Error;
	if (m_dwTimerMFlags & MTOFP_OVERMERGEDCELL)
		bMergeCell = GetMergeCellEx(m_hWndTimerCol, m_lTimerRow, hWndMergeCol, lMergeRow);
	else
	{
		bMergeCell = GetMergeCellType(m_hWndTimerCol, m_lTimerRow) != MC_NONE;
		if (bMergeCell)
		{
			hWndMergeCol = m_hWndTimerCol;
			lMergeRow = m_lTimerRow;
		}
	}

	if (bMergeCell)
	{
		if (GetLastMergedCellEx(hWndMergeCol, lMergeRow, hWndLastMergedCol, lLastMergedRow))
		{
			if (hWndMergeCol == hWndLastMergedCol && lMergeRow == lLastMergedRow)
				bMergeCell = FALSE;
		}
		else
			bMergeCell = FALSE;

		if (!bMergeCell)
		{
			hWndMergeCol = hWndLastMergedCol = NULL;
			lMergeRow = lLastMergedRow = TBL_Error;
		}
	}

	// Enter/Leave-Benachrichtigung...
	if (!hWndCapture || (hWndCapture && hWndCapture == m_hWndEdit))
	{
		m_pELItem = new CMTblItem();
		if (m_pELItem)
		{
			// Alte Werte merken
			DWORD dwELFlagsOld = m_dwELFlags;
			DWORD dwELMFlagsOld = m_dwELMFlags;

			HWND hWndELColOld = m_hWndELCol;
			HWND hWndELCol_BtnOld = m_hWndELCol_Btn;
			HWND hWndELCol_CellOld = m_hWndELCol_Cell;
			HWND hWndELCol_DDBtnOld = m_hWndELCol_DDBtn;
			CMTblRow *pELRowOld = m_pELRow;
			CMTblRow *pELRow_BtnOld = m_pELRow_Btn;
			CMTblRow *pELRow_CellOld = m_pELRow_Cell;
			CMTblRow *pELRow_DDBtnOld = m_pELRow_DDBtn;
			CMTblRow *pELRow_HdrOld = m_pELRow_Hdr;
			CMTblColHdrGrp * pELColHdrGrpOld = m_pELColHdrGrp;

			// Neue Werte ermitteln...

			// Flags
			m_dwELFlags = m_dwTimerFlags;
			m_dwELMFlags = m_dwTimerMFlags;

			// Column/Row
			m_hWndELCol = m_hWndTimerCol;
			m_pELRow = IsRow(m_lTimerRow) ? m_Rows->EnsureValid(m_lTimerRow) : NULL;

			// Row-Header
			LPMTBLMERGEROW pmr = NULL;
			if (m_pELRow && (pmr = FindMergeRow(m_pELRow->GetNr(), FMR_ANY)))
				m_pELRow_Hdr = m_Rows->EnsureValid(pmr->lRowFrom);
			else
				m_pELRow_Hdr = m_pELRow;

			// Cell
			LPMTBLMERGECELL pmc = NULL;
			if (m_hWndELCol && m_pELRow && (pmc = FindMergeCell(m_hWndELCol, m_pELRow->GetNr(), FMC_ANY)))
			{
				m_hWndELCol_Cell = pmc->hWndColFrom;
				m_pELRow_Cell = m_Rows->EnsureValid(pmc->lRowFrom);
			}
			else
			{
				m_hWndELCol_Cell = m_hWndELCol;
				m_pELRow_Cell = m_pELRow;
			}

			// Button
			if (bCursorOverBtn)
			{
				m_hWndELCol_Btn = hWndFocusCol;
				m_pELRow_Btn = m_Rows->EnsureValid(lFocusRow);
			}
			else if (m_dwELMFlags & MTOFP_OVERBTN)
			{
				// Zelle?
				if (m_pELRow_Cell)
				{
					m_hWndELCol_Btn = m_hWndELCol_Cell;
					m_pELRow_Btn = m_pELRow_Cell;
				}
				// Column-Header?
				if (m_hWndELCol && !(m_dwELMFlags & MTOFP_OVERCOLHDRGRP))
				{
					m_hWndELCol_Btn = m_hWndELCol;
					m_pELRow_Btn = NULL;
				}
			}
			else
			{
				m_hWndELCol_Btn = NULL;
				m_pELRow_Btn = NULL;
			}

			// Drop-Down-Button
			if (bCursorOverDDBtn)
			{
				m_hWndELCol_DDBtn = hWndFocusCol;
				m_pELRow_DDBtn = m_Rows->EnsureValid(lFocusRow);
			}
			else if (m_dwELMFlags & MTOFP_OVERDDBTN)
			{
				m_hWndELCol_DDBtn = m_hWndELCol_Cell;
				m_pELRow_DDBtn = m_pELRow_Cell;
			}
			else
			{
				m_hWndELCol_DDBtn = NULL;
				m_pELRow_DDBtn = NULL;
			}

			// Column-Header-Group
			if (m_dwELMFlags & MTOFP_OVERCOLHDRGRP)
				m_pELColHdrGrp = m_ColHdrGrps->GetGrpOfCol(m_hWndELCol);
			else
				m_pELColHdrGrp = NULL;

			// Benachrichtigungen generieren...

			BOOL bIsOver, bWasOver;

			// Corner
			bWasOver = (dwELFlagsOld & TBL_YOverColumnHeader && dwELFlagsOld & TBL_XOverRowHeader);
			bIsOver = (m_dwELFlags & TBL_YOverColumnHeader && m_dwELFlags & TBL_XOverRowHeader);
			if (bIsOver != bWasOver)
			{
				m_pELItem->DefineAsCorner(m_hWndTbl);
				if (bWasOver)
					OnLeaveItem();

				if (bIsOver)
					OnEnterItem();

			}

			// Row
			BOOL bWasOverRow = (pELRowOld != NULL);
			BOOL bIsOverRow = (m_pELRow != NULL);
			if (m_pELRow != pELRowOld)
			{
				if (bWasOverRow)
				{
					m_pELItem->DefineAsRow(m_hWndTbl, pELRowOld);
					OnLeaveItem();
				}

				if (bIsOverRow)
				{
					m_pELItem->DefineAsRow(m_hWndTbl, m_pELRow);
					OnEnterItem();
				}
			}

			// Row-Header
			BOOL bWasOverRowHdr = (pELRow_HdrOld != NULL && dwELFlagsOld & TBL_XOverRowHeader);
			BOOL bIsOverRowHdr = (m_pELRow_Hdr != NULL && m_dwELFlags & TBL_XOverRowHeader);
			if (m_pELRow_Hdr != pELRow_HdrOld || bWasOverRowHdr != bIsOverRowHdr)
			{
				if (bWasOverRowHdr)
				{
					m_pELItem->DefineAsRowHdr(m_hWndTbl, pELRow_HdrOld);
					OnLeaveItem();
				}

				if (bIsOverRowHdr)
				{
					m_pELItem->DefineAsRowHdr(m_hWndTbl, m_pELRow_Hdr);
					OnEnterItem();
				}
			}

			// Column
			BOOL bWasOverCol = bWasOver = (hWndELColOld != NULL);
			BOOL bIsOverCol = bIsOver = (m_hWndELCol != NULL);
			if (m_hWndELCol != hWndELColOld)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsCol(hWndELColOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsCol(m_hWndELCol);
					OnEnterItem();
				}
			}

			// Column-Header
			bWasOver = (bWasOverCol && dwELFlagsOld & TBL_YOverColumnHeader && !(dwELMFlagsOld & MTOFP_OVERCOLHDRGRP));
			bIsOver = (bIsOverCol && m_dwELFlags & TBL_YOverColumnHeader && !(m_dwELMFlags & MTOFP_OVERCOLHDRGRP));
			if (m_hWndELCol != hWndELColOld || bWasOver != bIsOver)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsColHdr(hWndELColOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsColHdr(m_hWndELCol);
					OnEnterItem();
				}
			}

			// Column-Header-Group
			bWasOver = (pELColHdrGrpOld != NULL);
			bIsOver = (m_pELColHdrGrp != NULL);
			if (m_pELColHdrGrp != pELColHdrGrpOld)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsColHdrGrp(m_hWndTbl, pELColHdrGrpOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsColHdrGrp(m_hWndTbl, m_pELColHdrGrp);
					OnEnterItem();
				}
			}

			// Cell
			BOOL bWasOverCell = bWasOver = (hWndELCol_CellOld && pELRow_CellOld);
			BOOL bIsOverCell = bIsOver = (m_hWndELCol_Cell && m_pELRow_Cell);

			if (m_hWndELCol_Cell != hWndELCol_CellOld || m_pELRow_Cell != pELRow_CellOld)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsCell(hWndELCol_CellOld, pELRow_CellOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsCell(m_hWndELCol_Cell, m_pELRow_Cell);
					OnEnterItem();
				}
			}

			// Node (Zelle)
			bWasOver = (bWasOverCell && dwELMFlagsOld & MTOFP_OVERNODE);
			bIsOver = (bIsOverCell && m_dwELMFlags & MTOFP_OVERNODE);
			if (m_hWndELCol_Cell != hWndELCol_CellOld || m_pELRow_Cell != pELRow_CellOld || bWasOver != bIsOver)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsCellNode(hWndELCol_CellOld, pELRow_CellOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsCellNode(m_hWndELCol_Cell, m_pELRow_Cell);
					OnEnterItem();
				}
			}

			// Node (Row-Header)
			bWasOver = (bWasOverRowHdr && dwELMFlagsOld & MTOFP_OVERNODE);
			bIsOver = (bIsOverRowHdr && m_dwELMFlags & MTOFP_OVERNODE);
			if (m_pELRow != pELRowOld || bWasOver != bIsOver)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsRowHdrNode(m_hWndTbl, pELRowOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsRowHdrNode(m_hWndTbl, m_pELRow);
					OnEnterItem();
				}
			}

			// Button
			bWasOver = (hWndELCol_BtnOld != NULL /*&& pELRow_BtnOld*/);
			bIsOver = (m_hWndELCol_Btn != NULL /*&& m_pELRow_Btn*/);
			if (m_hWndELCol_Btn != hWndELCol_BtnOld || m_pELRow_Btn != pELRow_BtnOld || bWasOver != bIsOver)
			{
				if (bWasOver)
				{
					if (pELRow_BtnOld)
						m_pELItem->DefineAsCellBtn(hWndELCol_BtnOld, pELRow_BtnOld);
					else if (hWndELCol_BtnOld)
						m_pELItem->DefineAsColHdrBtn(hWndELCol_BtnOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					if (m_pELRow_Btn)
						m_pELItem->DefineAsCellBtn(m_hWndELCol_Btn, m_pELRow_Btn);
					else if (m_hWndELCol_Btn)
						m_pELItem->DefineAsColHdrBtn(m_hWndELCol_Btn);
					OnEnterItem();
				}
			}

			// Drop-Down-Button
			bWasOver = (hWndELCol_DDBtnOld && pELRow_DDBtnOld);
			bIsOver = (m_hWndELCol_DDBtn && m_pELRow_DDBtn);
			if (m_hWndELCol_DDBtn != hWndELCol_DDBtnOld || m_pELRow_DDBtn != pELRow_DDBtnOld || bWasOver != bIsOver)
			{
				if (bWasOver)
				{
					m_pELItem->DefineAsCellDDBtn(hWndELCol_DDBtnOld, pELRow_DDBtnOld);
					OnLeaveItem();
				}

				if (bIsOver)
				{
					m_pELItem->DefineAsCellDDBtn(m_hWndELCol_DDBtn, m_pELRow_DDBtn);
					OnEnterItem();
				}
			}

			// Aufräumen
			delete m_pELItem;
			m_pELItem = NULL;
		}
	}


	// HOVER Hyperlink
	HWND hWndHLColOld = m_hWndHLCol;
	long lHLRowOld = m_lHLRow;

	BOOL bIsOverHyperlink = IsOverHyperlink(m_hWndTimerCol, m_lTimerRow, m_dwTimerMFlags);
	if (m_bHLHover && bIsOverHyperlink)
	{
		m_hWndHLCol = (bMergeCell ? hWndMergeCol : m_hWndTimerCol);
		m_lHLRow = (bMergeCell ? lMergeRow : m_lTimerRow);
	}
	else
	{
		m_hWndHLCol = NULL;
		m_lHLRow = TBL_Error;
	}

	BOOL bHLChanged = (m_hWndHLCol != hWndHLColOld || m_lHLRow != lHLRowOld);
	if (bHLChanged)
	{
		InvalidateCell(m_hWndHLCol, m_lHLRow, NULL, FALSE);
		InvalidateCell(hWndHLColOld, lHLRowOld, NULL, FALSE);
	}

	// Tip
	BOOL bDestroyTip = FALSE, bShowTip = FALSE;

	BOOL bCornerTipOld = m_bCornerTip;
	CMTblColHdrGrp *pColHdrGrpTipObjOld = m_pColHdrGrpTipObj;
	HWND hWndBtnTipColOld = m_hWndBtnTipCol, hWndCellTipColOld = m_hWndCellTipCol, hWndColHdrTipColOld = m_hWndColHdrTipCol;
	int iCellTipTypeOld = m_iTipType;
	long lBtnTipRowOld = m_lBtnTipRow, lCellTipRowOld = m_lCellTipRow, lRowHdrTipRowOld = m_lRowHdrTipRow;

	m_hWndCellTipCol = (bMergeCell ? hWndMergeCol : m_hWndTimerCol);
	m_lCellTipRow = (bMergeCell ? lMergeRow : m_lTimerRow);

	m_hWndColHdrTipCol = (m_hWndTimerCol && m_dwTimerFlags & TBL_YOverColumnHeader) ? m_hWndTimerCol : NULL;
	if (m_hWndColHdrTipCol && m_dwTimerMFlags & MTOFP_OVERCOLHDRGRP)
	{
		if (m_pColHdrGrpTipObj = m_ColHdrGrps->GetGrpOfCol(m_hWndColHdrTipCol))
			m_hWndColHdrTipCol = NULL;
	}
	else
		m_pColHdrGrpTipObj = NULL;

	m_hWndTipBtn = (bCursorOverBtn ? hWndFromPoint : NULL);
	if (m_hWndTipBtn)
	{
		m_hWndBtnTipCol = hWndFocusCol;
		m_lBtnTipRow = lFocusRow;
	}
	else if (m_dwTimerMFlags & MTOFP_OVERBTN)
	{
		// Zelle?
		if (m_lTimerRow != TBL_Error)
		{
			m_hWndBtnTipCol = (bMergeCell ? hWndMergeCol : m_hWndTimerCol);
			m_lBtnTipRow = (bMergeCell ? lMergeRow : m_lTimerRow);
		}
		// Column-Header?
		else if (m_hWndTimerCol && !m_pColHdrGrpTipObj)
		{
			m_hWndBtnTipCol = m_hWndTimerCol;
			m_lBtnTipRow = TBL_Error;
		}
	}
	else
	{
		m_hWndBtnTipCol = NULL;
		m_lBtnTipRow = TBL_Error;
	}

	m_bCornerTip = (m_dwTimerFlags & TBL_YOverColumnHeader && m_dwTimerFlags & TBL_XOverRowHeader);

	if (m_lTimerRow != TBL_Error && m_dwTimerFlags & TBL_XOverRowHeader)
		m_lRowHdrTipRow = m_lTimerRow;
	else
		m_lRowHdrTipRow = TBL_Error;

	BOOL bOverBtn = (m_hWndBtnTipCol != NULL /*&& m_lBtnTipRow != TBL_Error*/);
	BOOL bOverCell = (m_hWndCellTipCol && m_lCellTipRow != TBL_Error);
	BOOL bOverColHdr = (m_hWndColHdrTipCol != NULL);
	BOOL bOverColHdrGrp = (m_pColHdrGrpTipObj != NULL);
	BOOL bOverCorner = m_bCornerTip;
	BOOL bOverRowHdr = (m_lRowHdrTipRow != TBL_Error);

	if (bOverBtn || bOverCell || bOverColHdr || bOverColHdrGrp || bOverCorner || bOverRowHdr)
	{
		BOOL bQueryTip = FALSE;
		BOOL bOtherCell = FALSE;
		BOOL bOtherCellItem = FALSE;
		BOOL bOtherItem = FALSE;
		BOOL bOverBtnOld = FALSE;
		BOOL bOverImage = FALSE;
		BOOL bOverImageOld = FALSE;
		BOOL bOverText = FALSE;
		BOOL bOverTextOld = FALSE;
		BOOL bOtherTip = FALSE;

		if (bOverBtn)
		{
			bOtherItem = (m_hWndBtnTipCol != hWndBtnTipColOld || m_lBtnTipRow != lBtnTipRowOld);
		}
		else if (bOverCell)
		{
			bOverText = (m_dwTimerMFlags & MTOFP_OVERTEXT);
			bOverImage = (m_dwTimerMFlags & MTOFP_OVERIMAGE);

			bOverTextOld = (dwTimerMFlagsOld & MTOFP_OVERTEXT);
			bOverImageOld = (dwTimerMFlagsOld & MTOFP_OVERIMAGE);

			bOverBtnOld = (hWndBtnTipColOld && lBtnTipRowOld != TBL_Error);

			bOtherCell = (m_hWndCellTipCol != hWndCellTipColOld || m_lCellTipRow != lCellTipRowOld);
			bOtherCellItem = (bOverText != bOverTextOld || bOverImage != bOverImageOld || bOverBtn != bOverBtnOld);
			bOtherItem = bOtherCell || (!bOtherCell && bOtherCellItem);
		}
		else if (bOverColHdr)
		{
			bOtherItem = (m_hWndColHdrTipCol != hWndColHdrTipColOld);
			if (!bOtherItem)
			{
				bOverBtnOld = hWndBtnTipColOld != NULL;
				bOtherItem = bOverBtn != bOverBtnOld;
			}
		}
		else if (bOverColHdrGrp)
		{
			bOtherItem = (m_pColHdrGrpTipObj != pColHdrGrpTipObjOld);
		}
		else if (bOverCorner)
		{
			bOtherItem = (m_bCornerTip != bCornerTipOld);
		}
		else if (bOverRowHdr)
		{
			bOtherItem = (m_lRowHdrTipRow != lRowHdrTipRowOld);
		}

		if (bOtherItem)
		{
			bQueryTip = TRUE;
		}

		if (bQueryTip)
		{
			long lMsgRet;

			m_pTip = NULL;

			// Zelle?
			if (bOverBtn && /*bOverCell*/!bOverColHdr && IsTipTypeEnabled(MTBL_TIP_CELL_BTN))
			{
				lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellBtnTip, (WPARAM)m_hWndBtnTipCol, (LPARAM)m_lBtnTipRow);
				if (lMsgRet != MTBL_NODEFAULTACTION)
				{
					m_iTipType = MTBL_TIP_CELL_BTN;
					m_pTip = GetTip(m_lBtnTipRow, m_hWndBtnTipCol, m_iTipType);
				}
			}
			// Column-Header?
			else if (bOverBtn && bOverColHdr && IsTipTypeEnabled(MTBL_TIP_COLHDR_BTN))
			{
				lMsgRet = SendMessage(m_hWndTbl, MTM_ShowColHdrBtnTip, (WPARAM)m_hWndBtnTipCol, (LPARAM)0);
				if (lMsgRet != MTBL_NODEFAULTACTION)
				{
					m_iTipType = MTBL_TIP_COLHDR_BTN;
					m_pTip = GetTip(m_lBtnTipRow, m_hWndBtnTipCol, m_iTipType);
				}
			}
			else if (bOverCell)
			{
				if (bOverText)
				{
					BOOL bShowComplTextTip = FALSE;
					if (IsTipTypeEnabled(MTBL_TIP_CELL_COMPLETETEXT) && GetEffCellType(m_lCellTipRow, m_hWndCellTipCol) != COL_CellType_CheckBox)
					{
						CMTblMetrics tm(m_hWndTbl);

						// Textausmaße ermitteln
						//BOOL bWordWrap = tm.m_LinesPerRow > 1 && SendMessage( m_hWndCellTipCol, TIM_QueryColumnFlags, 0, COL_MultilineCell );
						BOOL bWordWrap = GetCellWordWrap(m_hWndCellTipCol, m_lCellTipRow);
						SIZE siText;
						GetCellTextExtent(m_hWndCellTipCol, m_lCellTipRow, &siText, NULL, bWordWrap);

						// X-Koordinaten des Textbereichs ermitteln
						POINT ptTAX, ptTAVisX;
						GetCellTextAreaXCoord(m_lCellTipRow, m_hWndCellTipCol, &tm, NULL, &ptTAX, &ptTAVisX);

						// Y-Koordinaten ermitteln
						POINT ptRowY, ptRowVisY;
						//int iRowYRet = MTblGetRowYCoord( m_hWndTbl, m_lCellTipRow, &tm, &ptRowY, &ptRowVisY, !bMergeCell );
						int iRowYRet = GetRowYCoord(m_lCellTipRow, &tm, &ptRowY, &ptRowVisY, !bMergeCell);
						if (bMergeCell && lLastMergedRow != m_lCellTipRow)
						{
							POINT ptY, ptVisY;
							//int iRowYRet2 = MTblGetRowYCoord( m_hWndTbl, lLastMergedRow, &tm, &ptY, &ptVisY, FALSE );
							int iRowYRet2 = GetRowYCoord(lLastMergedRow, &tm, &ptY, &ptVisY, FALSE);

							ptRowY.y = ptY.y;

							long lRowsTop = (m_lCellTipRow < 0) ? tm.m_SplitRowsTop : tm.m_RowsTop;
							long lRowsBottom = (m_lCellTipRow < 0) ? tm.m_SplitRowsBottom : tm.m_RowsBottom;

							if (iRowYRet == MTGR_RET_INVISIBLE)
								ptRowVisY.x = lRowsTop;

							if (iRowYRet2 == MTGR_RET_INVISIBLE)
								ptRowVisY.y = lRowsBottom - 1;
							else
								ptRowVisY.y = ptVisY.y;
						}

						// Text-Rechteck berechnen
						int iJustify = GetEffCellTextJustify(m_lCellTipRow, m_hWndCellTipCol);
						int iVAlign = GetEffCellTextVAlign(m_lCellTipRow, m_hWndCellTipCol);
						RECT rText = { 0, 0, 0, 0 };
						CalcCellTextRect(ptRowY.x, ptRowY.y, ptTAX.x, ptTAX.y, 0, tm.m_CellLeadingY, iJustify, iVAlign, siText, &rText);

						// Irgendwas nicht sichtbar?
						BOOL bLeftOut = rText.left < ptTAVisX.x;
						BOOL bRightOut = rText.right > ptTAVisX.y;
						BOOL bTopOut = rText.top < ptRowVisY.x;
						BOOL bBottomOut = rText.bottom > ptRowVisY.y;

						if ((bLeftOut || bRightOut || bTopOut || bBottomOut))
						{
							CMTblTip *pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, MTBL_TIP_CELL_COMPLETETEXT, TRUE);
							if (pTip)
							{
								// Tip modifizieren
								DWORD dwFlags;

								if (MustUseEndEllipsis(m_hWndCellTipCol) && siText.cx > (ptTAX.y - ptTAX.x))
								{
									SIZE siTextEE;
									GetCellTextExtent(m_hWndCellTipCol, m_lCellTipRow, &siTextEE, NULL, bWordWrap, FALSE, TRUE, FALSE);
									CalcCellTextRect(ptRowY.x, ptRowY.y, ptTAX.x, ptTAX.y, 0, tm.m_CellLeadingY, iJustify, iVAlign, siTextEE, &rText);
								}

								// Über Hyperlink?
								BOOL bHyperlink = IsOverHyperlink(m_hWndCellTipCol, m_lCellTipRow, m_dwTimerMFlags);

								// Position
								pTip->m_iPosX = rText.left - 4;
								pTip->m_iPosY = rText.top - 2;

								// Breite/Höhe
								pTip->m_iWid = siText.cx + 8;
								pTip->m_iHt = siText.cy + 4;

								// Permeable-Flag
								pTip->SetFlags(MTBL_TIP_FLAG_PERMEABLE, TRUE);

								// Singleline-Flag
								if (bWordWrap || QueryFlags(MTBL_FLAG_EXPAND_CRLF))
									pTip->SetFlags(MTBL_TIP_FLAG_SINGLELINE, FALSE);
								else
									pTip->SetFlags(MTBL_TIP_FLAG_SINGLELINE, TRUE);

								// Ausrichtungs-Flags
								pTip->SetFlags(MTBL_TIP_FLAG_CENTER, FALSE);
								pTip->SetFlags(MTBL_TIP_FLAG_RIGHT, FALSE);
								if (iJustify == DT_CENTER)
									pTip->SetFlags(MTBL_TIP_FLAG_CENTER, TRUE);
								else if (iJustify == DT_RIGHT)
									pTip->SetFlags(MTBL_TIP_FLAG_RIGHT, TRUE);

								// Text
								GetEffCellText(m_lCellTipRow, m_hWndCellTipCol, pTip->m_sText);

								// Font
								HDC hDC = GetDC(m_hWndTbl);
								if (hDC)
								{
									BOOL bMustDelete = FALSE;

									dwFlags = 0;
									if (bHyperlink)
										dwFlags |= GEF_HYPERLINK;
									if (AnyHighlightedItems())
										dwFlags |= GEF_HILI;

									HFONT hFont = GetEffCellFont(hDC, m_lCellTipRow, m_hWndCellTipCol, &bMustDelete, dwFlags);
									pTip->m_Font.Set(hDC, hFont);

									if (bMustDelete)
										DeleteObject(hFont);

									ReleaseDC(m_hWndTbl, hDC);
								}

								lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellCompleteTextTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
								bShowComplTextTip = (lMsgRet != MTBL_NODEFAULTACTION) && pTip->MustShow();
								if (bShowComplTextTip)
								{
									m_iTipType = MTBL_TIP_CELL_COMPLETETEXT;
									m_pTip = pTip;
								}
							}
						}
					}

					if (!bShowComplTextTip && IsTipTypeEnabled(MTBL_TIP_CELL_TEXT))
					{
						lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellTextTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
						if (lMsgRet != MTBL_NODEFAULTACTION)
						{
							m_iTipType = MTBL_TIP_CELL_TEXT;
							m_pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, m_iTipType);
						}
					}

					if ((m_pTip && !m_pTip->MustShow()) || !m_pTip)
					{
						if (bOverImage && IsTipTypeEnabled(MTBL_TIP_CELL_IMAGE))
						{
							lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellImageTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
							if (lMsgRet != MTBL_NODEFAULTACTION)
							{
								m_iTipType = MTBL_TIP_CELL_IMAGE;
								m_pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, m_iTipType);
							}
						}

						if (((m_pTip && !m_pTip->MustShow()) || !m_pTip) && IsTipTypeEnabled(MTBL_TIP_CELL))
						{
							lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
							if (lMsgRet != MTBL_NODEFAULTACTION)
							{
								m_iTipType = MTBL_TIP_CELL;
								m_pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, m_iTipType);
							}
						}
					}
				}
				else if (bOverImage && IsTipTypeEnabled(MTBL_TIP_CELL_IMAGE))
				{
					lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellImageTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
					if (lMsgRet != MTBL_NODEFAULTACTION)
					{
						m_iTipType = MTBL_TIP_CELL_IMAGE;
						m_pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, m_iTipType);
					}

					if (((m_pTip && !m_pTip->MustShow()) || !m_pTip) && IsTipTypeEnabled(MTBL_TIP_CELL))
					{
						lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
						if (lMsgRet != MTBL_NODEFAULTACTION)
						{
							m_iTipType = MTBL_TIP_CELL;
							m_pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, m_iTipType);
						}
					}
				}
				else if (IsTipTypeEnabled(MTBL_TIP_CELL))
				{
					lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCellTip, (WPARAM)m_hWndCellTipCol, (LPARAM)m_lCellTipRow);
					if (lMsgRet != MTBL_NODEFAULTACTION)
					{
						m_iTipType = MTBL_TIP_CELL;
						m_pTip = GetTip(m_lCellTipRow, m_hWndCellTipCol, m_iTipType);
					}
				}
			}
			else if (bOverColHdr && IsTipTypeEnabled(MTBL_TIP_COLHDR))
			{
				lMsgRet = SendMessage(m_hWndTbl, MTM_ShowColHdrTip, (WPARAM)m_hWndColHdrTipCol, 0);
				if (lMsgRet != MTBL_NODEFAULTACTION)
				{
					m_iTipType = MTBL_TIP_COLHDR;
					m_pTip = GetTip(TBL_Error, m_hWndColHdrTipCol, m_iTipType);
				}
			}
			else if (bOverColHdrGrp && IsTipTypeEnabled(MTBL_TIP_COLHDRGRP))
			{
				lMsgRet = SendMessage(m_hWndTbl, MTM_ShowColHdrGrpTip, (WPARAM)m_hWndTbl, (LPARAM)m_pColHdrGrpTipObj);
				if (lMsgRet != MTBL_NODEFAULTACTION)
				{
					m_iTipType = MTBL_TIP_COLHDRGRP;
					m_pTip = GetTip((long)m_pColHdrGrpTipObj, NULL, m_iTipType);
				}
			}
			else if (bOverCorner && IsTipTypeEnabled(MTBL_TIP_CORNER))
			{
				lMsgRet = SendMessage(m_hWndTbl, MTM_ShowCornerTip, (WPARAM)m_hWndTbl, 0);
				if (lMsgRet != MTBL_NODEFAULTACTION)
				{
					m_iTipType = MTBL_TIP_CORNER;
					m_pTip = GetTip(0, NULL, m_iTipType);
				}
			}
			else if (bOverRowHdr && IsTipTypeEnabled(MTBL_TIP_ROWHDR))
			{
				lMsgRet = SendMessage(m_hWndTbl, MTM_ShowRowHdrTip, (WPARAM)m_hWndTbl, m_lRowHdrTipRow);
				if (lMsgRet != MTBL_NODEFAULTACTION)
				{
					m_iTipType = MTBL_TIP_ROWHDR;
					m_pTip = GetTip(m_lRowHdrTipRow, NULL, m_iTipType);
				}
			}
		}

		if (m_pTip && m_pTip->MustShow())
		{
			bOtherTip = TRUE;
			if (bOverCell && !bOverBtn)
			{
				if (!bOtherCell && iCellTipTypeOld == m_iTipType)
					bOtherTip = FALSE;
			}
			else
			{
				if (!bOtherItem && iCellTipTypeOld == m_iTipType)
					bOtherTip = FALSE;
			}

			if (bOtherTip)
			{
				if (m_bTipShowing)
					bDestroyTip = TRUE;
				m_iiTipTime = _I64_MAX;
			}
			else
			{
				if (!m_bTipShowing && bSameCursorPos && m_iiTipTime == _I64_MAX)
					m_iiTipTime = m_iiTimer + m_pTip->GetDelayI64();
				else if (!m_bTipShowing && !bSameCursorPos)
					m_iiTipTime = _I64_MAX;
			}
		}
		else
		{
			if (m_bTipShowing)
				bDestroyTip = TRUE;
			m_iTipType = 0;
			m_pTip = NULL;
			m_iiTipTime = _I64_MAX;
		}
	}
	else if (!bCursorOverTip)
	{
		if (m_bTipShowing)
			bDestroyTip = TRUE;
		m_iTipType = 0;
		m_pTip = NULL;
		m_iiTipTime = _I64_MAX;
	}

	if (m_iiTimer >= m_iiTipTime)
	{
		m_iiTipTime = _I64_MAX;

		if (!IsWindow(m_hWndTip))
			CreateTipWnd();

		HDC hDC = GetDC(m_hWndTip);
		if (hDC)
		{
			HFONT hFont = m_pTip->CreateFont(hDC);
			if (hFont)
			{
				HGDIOBJ hOldFont = SelectObject(hDC, hFont);
				RECT r = { 0, 0, 0, 0 };
				m_dwTipDrawFlags = DT_NOPREFIX;
				if (QueryFlags(MTBL_FLAG_EXPAND_TABS))
					m_dwTipDrawFlags |= DT_EXPANDTABS;
				if (m_pTip->QueryFlags(MTBL_TIP_FLAG_SINGLELINE))
					m_dwTipDrawFlags |= DT_SINGLELINE;
				if (m_pTip->QueryFlags(MTBL_TIP_FLAG_CENTER))
					m_dwTipDrawFlags |= DT_CENTER;
				if (m_pTip->QueryFlags(MTBL_TIP_FLAG_RIGHT))
					m_dwTipDrawFlags |= DT_RIGHT;

				int iTipWid = m_pTip->GetWid();
				if (iTipWid != INT_MIN)
				{
					r.right = iTipWid - 8;
					m_dwTipDrawFlags |= DT_WORDBREAK;
				}
				DrawText(hDC, m_pTip->GetText().c_str(), -1, &r, m_dwTipDrawFlags | DT_CALCRECT);
				int iWid = r.right - r.left + 8;
				int iHt = r.bottom - r.top + 4;
				SelectObject(hDC, hOldFont);
				DeleteObject(hFont);

				/*int iScrHt, iScrLeft, iScrTop, iScrWid;
				GetScreenMetrics(iScrLeft, iScrTop, iScrWid, iScrHt);

				POINT ptTip = m_ptTimerPos;

				int iOffsY = m_pTip->GetOffsY();
				if (iOffsY != INT_MIN)
				{
					if (ptTip.y + iOffsY + iHt > iScrHt)
						ptTip.y = (iScrHt - iHt);
					else if (ptTip.y + iOffsY < iScrTop)
						ptTip.y = iScrTop;
					else
						ptTip.y += iOffsY;
				}

				int iOffsX = m_pTip->GetOffsX();
				if (iOffsX != INT_MIN)
				{
					if (ptTip.x + iOffsX + iWid > iScrWid)
						ptTip.x = iScrWid - iWid;
					else if (ptTip.x + iOffsX < iScrLeft)
						ptTip.x = iScrLeft;
					else
						ptTip.x += iOffsX;
				}*/

				POINT ptTip = m_ptTimerPos;
				int iOffsX = m_pTip->GetOffsX();
				int iOffsY = m_pTip->GetOffsY();

				RECT rScreen;

				if (GetScreenRectFromPoint(ptTip, rScreen))
				{
					if (iOffsX != INT_MIN)
					{
						if (ptTip.x + iOffsX + iWid > rScreen.right)
							ptTip.x = rScreen.right - iWid;
						else if (ptTip.x + iOffsX < rScreen.left)
							ptTip.x = rScreen.left;
						else
							ptTip.x += iOffsX;
					}

					if (iOffsY != INT_MIN)
					{
						if (ptTip.y + iOffsY + iHt > rScreen.bottom)
							ptTip.y = (rScreen.bottom - iHt);
						else if (ptTip.y + iOffsY < rScreen.top)
							ptTip.y = rScreen.top;
						else
							ptTip.y += iOffsY;
					}
				}


				BOOL bFixX = FALSE, bFixY = FALSE;
				POINT ptTipFix = { 0, 0 };
				int iTipPosX = m_pTip->GetPosX();
				if (iTipPosX != INT_MIN)
				{
					ptTipFix.x = iTipPosX;
					bFixX = TRUE;
				}

				int iTipPosY = m_pTip->GetPosY();
				if (iTipPosY != INT_MIN)
				{
					ptTipFix.y = iTipPosY;
					bFixY = TRUE;
				}

				if (bFixX || bFixY)
				{
					ClientToScreen(m_hWndTbl, &ptTipFix);
					if (bFixX)
						ptTip.x = ptTipFix.x;
					if (bFixY)
						ptTip.y = ptTipFix.y;
				}

				if (iTipWid != INT_MIN)
				{
					iWid = iTipWid;
				}

				int iTipHt = m_pTip->GetHt();
				if (iTipHt != INT_MIN)
				{
					iHt = iTipHt;
				}

				if (IsWindowVisible(m_hWndTip))
					ShowWindow(m_hWndTip, SW_HIDE);
				MoveWindow(m_hWndTip, ptTip.x, ptTip.y, iWid, iHt, FALSE);

				HRGN hRgnOld = m_hRgnTip;
				m_hRgnTip = CreateRoundRectRgn(0, 0, iWid + 1, iHt + 1, m_pTip->GetRCEWid(), m_pTip->GetRCEHt());
				SetWindowRgn(m_hWndTip, m_hRgnTip, FALSE);
				if (hRgnOld)
					DeleteObject(hRgnOld);

				CAlphaWnd aw = theAlphaWnd();
				int iFadeInTime = m_pTip->GetFadeInTime();
				if (aw.CanUse() && iFadeInTime >= 0)
				{
					int iOpacity = m_pTip->GetOpacity();
					aw.SetAlpha(m_hWndTip, 0);
					ShowWindow(m_hWndTip, SW_SHOWNOACTIVATE);
					if (iOpacity > 0)
					{
						__int64 iiStartTime = GetSystemTimeI64();
						__int64 iiFadeInTime = __int64(iFadeInTime * 10000);
						__int64 iiEndTime = iiStartTime + iiFadeInTime;
						__int64 iiCurTime;
						int iCurOpacity = 0, iNewOpacity, iPct;
						for (;;)
						{
							iiCurTime = GetSystemTimeI64();
							if (iiCurTime >= iiEndTime)
								iNewOpacity = iOpacity;
							else
							{
								iPct = (iiCurTime - iiStartTime) * 100 / iiFadeInTime;
								iNewOpacity = iOpacity * iPct / 100;
								if (iNewOpacity < 0)
									iNewOpacity = iOpacity;
							}

							if (iNewOpacity > iCurOpacity)
							{
								aw.SetAlpha(m_hWndTip, iNewOpacity);
								iCurOpacity = iNewOpacity;
								RedrawWindow(m_hWndTip, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
							}

							if (iNewOpacity >= iOpacity)
								break;
						}
					}
				}
				else
					ShowWindow(m_hWndTip, SW_SHOWNOACTIVATE);
			}

			ReleaseDC(m_hWndTip, hDC);
		}

		m_bTipShowing = TRUE;
	}
	else if (bDestroyTip)
	{
		/*if ( IsWindowVisible( m_hWndTip ) )
		ShowWindow( m_hWndTip, SW_HIDE );*/
		if (IsWindow(m_hWndTip))
		{
			DestroyWindow(m_hWndTip);
			m_hWndTip = NULL;
		}
		m_bTipShowing = FALSE;
	}


	// Ggf. Timer stoppen?
	if (!IsMyWindow(hWndFromPoint))
	{
		StopInternalTimer();
	}

	return 1;
}


// OnKeyDown
LRESULT MTBLSUBCLASS::OnKeyDown(WPARAM wParam, LPARAM lParam)
{
	BOOL bCheckFocusRowRedraw, bCtrlPressed, bEnterEdit = FALSE, bLockUpdate = FALSE, bModifyFocusPaint, bNextFocusCellFound = FALSE, bNextFocusRowFound = FALSE, bNoScroll, bOldProc, bRemoveNewRowSel = FALSE, bShiftPressed, bSingleSelection, bSplitRow, bTrapColSelChanges = FALSE, bTrapRowSelChanges = FALSE;
	CMTblRow * pRow;
	HWND hWndCol, hWndEffFocusCol = NULL, hWndFocusCol, hWndNextCol = NULL, hWndSalParent;
	int iCellType, iPos, iRedrawMode, iSelChangeMode;
	long lEffFocusRow = TBL_Error, lFocusRow = TBL_Error, lFocusRowBefore = TBL_Error, lMaxVisRow, lMinVisRow, lNextRow = TBL_Error, lRet = 0, lRow, lRowSelTrapRow = TBL_Error;
	LPMTBLMERGEROW pmr;
	RowPtrVector vSelRowsAfter, vSelRowsBefore, *pvSelRowsAfter = NULL, *pvSelRowsBefore = NULL;
	WORD wChar;

	if (SendMessage(m_hWndTbl, MTM_KeyDown, wParam, lParam) == MTBL_NODEFAULTACTION)
		return 0;

	if (!Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol))
		return CallOldWndProc(WM_KEYDOWN, wParam, lParam);

	if (ProcessHotkey(WM_KEYDOWN, wParam, lParam))
		return 0;

	bCtrlPressed = HIBYTE(GetKeyState(VK_CONTROL));
	bShiftPressed = HIBYTE(GetKeyState(VK_SHIFT));

	// Tastenkombination kopieren?
	if (bCtrlPressed && wParam == VK_INSERT)
		CopyShortcutPressed();

	// Automatisch in Edit-Mode wechseln?
	bEnterEdit = m_bCellMode && (m_dwCellModeFlags & MTBL_CM_FLAG_AUTO_EDIT_MODE);
	bEnterEdit = bEnterEdit && !bCtrlPressed && (wChar = LOWORD(MapVirtualKey(wParam, 2)) && wParam != VK_ESCAPE);
	bEnterEdit = bEnterEdit && GetEffFocusCell(lEffFocusRow, hWndEffFocusCol) && hWndEffFocusCol && lEffFocusRow != TBL_Error && !IsRectEmpty(&m_rFocus);
	bEnterEdit = bEnterEdit && CanSetFocusCell(hWndEffFocusCol, lEffFocusRow);

	//iCellType = Ctd.TblGetCellType( hWndEffFocusCol );
	iCellType = GetEffCellType(lEffFocusRow, hWndEffFocusCol);
	bEnterEdit = bEnterEdit && !(iCellType == COL_CellType_CheckBox && wParam != VK_SPACE);

	bEnterEdit = bEnterEdit && !(wParam == VK_ADD && m_Rows->QueryRowFlags(lEffFocusRow, MTBL_ROW_CANEXPAND) && !m_Rows->QueryRowFlags(lEffFocusRow, MTBL_ROW_ISEXPANDED));
	bEnterEdit = bEnterEdit && !(wParam == VK_SUBTRACT && m_Rows->QueryRowFlags(lEffFocusRow, MTBL_ROW_CANEXPAND) && m_Rows->QueryRowFlags(lEffFocusRow, MTBL_ROW_ISEXPANDED));
	bEnterEdit = bEnterEdit && !(wParam == VK_MULTIPLY && m_Rows->QueryRowFlags(lEffFocusRow, MTBL_ROW_CANEXPAND));

	if (bEnterEdit)
	{
		if (Ctd.TblSetFocusCell(m_hWndTbl, lEffFocusRow, hWndEffFocusCol, 0, -1))
		{
			HWND hWndFocus = GetFocus();
			if (hWndFocus != m_hWndTbl)
			{
				if (iCellType == COL_CellType_PopupEdit)
					SendMessage(hWndFocus, EM_SETSEL, 0, -1);

				keybd_event(wParam, 0, 0, GetMessageExtraInfo());

				return 0;
			}
		}
	}

	// Control simulieren?
	BOOL bCtrlSim = FALSE;
	/*if ( m_bCellMode && (m_dwCellModeFlags & MTBL_CM_FLAG_NAV_NO_SELCHANGE) )
	{
	BYTE ks[256];

	switch ( wParam )
	{
	case VK_DOWN:
	case VK_NEXT:
	case VK_UP:
	case VK_PRIOR:
	case VK_END:
	case VK_HOME:
	if ( GetKeyboardState( ks ) )
	{
	ks[VK_CONTROL] |= 0x80;
	SetKeyboardState( ks );
	bCtrlSim = TRUE;
	bCtrlPressed = TRUE;
	}
	break;
	}
	}*/

	bSingleSelection = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection);

	bCheckFocusRowRedraw = !bSingleSelection && MustRedrawSelections() && !gbCtdHooked;
	if (bCheckFocusRowRedraw)
	{
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRowBefore, &hWndFocusCol);
		pvSelRowsBefore = &vSelRowsBefore;
		pvSelRowsAfter = &vSelRowsAfter;
	}

	if (MustTrapColSelChanges())
	{
		switch (wParam)
		{
		case VK_DOWN:
		case VK_NEXT:
		case VK_UP:
		case VK_PRIOR:
		case VK_END:
		case VK_HOME:
			if (bSingleSelection)
				bTrapColSelChanges = TRUE;
			else if (!bCtrlPressed)
				bTrapColSelChanges = TRUE;
			break;
		case VK_INSERT:
			if (!bCtrlPressed)
				bTrapColSelChanges = TRUE;
			break;
		}
	}

	if (MustTrapRowSelChanges())
	{
		switch (wParam)
		{
		case VK_DOWN:
		case VK_NEXT:
		case VK_UP:
		case VK_PRIOR:
		case VK_END:
		case VK_HOME:
			if (bSingleSelection)
				bTrapRowSelChanges = TRUE;
			else if (!bCtrlPressed)
				bTrapRowSelChanges = TRUE;
			break;
		case VK_INSERT:
			if (!bCtrlPressed)
				bTrapRowSelChanges = TRUE;
			break;
		case VK_SPACE:
			if (!bSingleSelection)
			{
				bTrapRowSelChanges = TRUE;
				lRowSelTrapRow = lRow;
			}
			break;
		}
	}

	if (MustRedrawSelections() && !IsSplitted() && !gbCtdHooked)
	{
		switch (wParam)
		{
		case VK_DOWN:
		case VK_UP:
			if (Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol) && lFocusRow >= 0 && Ctd.TblQueryVisibleRange(m_hWndTbl, &lMinVisRow, &lMaxVisRow))
			{
				long lRow = lFocusRow;
				if (wParam == VK_DOWN && (lFocusRow < lMaxVisRow || !Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, ROW_Hidden)))
					bLockUpdate = TRUE;
				else if (wParam == VK_UP && (lFocusRow > lMinVisRow || !Ctd.TblFindPrevRow(m_hWndTbl, &lRow, 0, ROW_Hidden)))
					bLockUpdate = TRUE;
			}
			break;
		case VK_NEXT:
		case VK_PRIOR:
		case VK_END:
		case VK_HOME:
			if (!bCtrlPressed && bShiftPressed)
				bLockUpdate = TRUE;
			break;
		case VK_SPACE:
			if (!bSingleSelection)
				bLockUpdate = TRUE;
			break;
		}
	}

	if (bTrapColSelChanges)
		BeginColSelTrap();

	if (bTrapRowSelChanges)
		BeginRowSelTrap(lRowSelTrapRow, pvSelRowsBefore);

	if (bLockUpdate)
		bLockUpdate = LockWindowUpdate(m_hWndTbl);

	switch (wParam)
	{
	case VK_ADD:
		if (m_Rows->QueryRowFlags(lRow, MTBL_ROW_CANEXPAND) &&
			!m_Rows->QueryRowFlags(lRow, MTBL_ROW_ISEXPANDED))
		{
			pRow = m_Rows->EnsureValid(lRow);
			SendMessage(m_hWndTbl, MTM_ExpandRow, MTER_REDRAW | MTER_BY_USER, lRow);
			if (pRow)
				Ctd.TblSetFocusRow(m_hWndTbl, pRow->GetNr());
		}

		lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);
		break;

	case VK_INSERT:
		if (!bCtrlPressed)
		{
			if (m_bCellMode)
			{
				if (m_hWndColFocus && GetRowNrFocus() != TBL_Error && CanSetFocusCell(m_hWndColFocus, GetRowNrFocus()))
					Ctd.TblSetFocusCell(m_hWndTbl, GetRowNrFocus(), m_hWndColFocus, 0, -1);
			}
			else if (!m_Rows->IsRowDisabled(lRow))
			{
				iPos = 0;
				while (TRUE)
				{
					// Nächste Spalte suchen
					iPos++;
					hWndCol = Ctd.TblGetColumnWindow(m_hWndTbl, iPos, COL_GetPos);
					if (!hWndCol) break;
					//if ( Ctd.IsWindowVisible( hWndCol ) )
					if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible))
					{
						if (CanSetFocusCell(hWndCol, lRow))
						{
							Ctd.TblSetFocusCell(m_hWndTbl, lRow, hWndCol, 0, -1);
							break;
						}
					}
				}
			}
		}

		lRet = 0;
		break;

	case VK_MULTIPLY:
		if (m_Rows->QueryRowFlags(lRow, MTBL_ROW_CANEXPAND))
		{
			pRow = m_Rows->EnsureValid(lRow);
			SendMessage(m_hWndTbl, MTM_ExpandRow, MTER_REDRAW | MTER_RECURSIVE | MTER_BY_USER, lRow);
			if (pRow)
				Ctd.TblSetFocusRow(m_hWndTbl, pRow->GetNr());

			bEnterEdit = FALSE;
		}

		lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);
		break;

	case VK_SUBTRACT:
		if (m_Rows->QueryRowFlags(lRow, MTBL_ROW_CANEXPAND) &&
			m_Rows->QueryRowFlags(lRow, MTBL_ROW_ISEXPANDED))
		{
			SendMessage(m_hWndTbl, MTM_CollapseRow, MTCR_REDRAW | MTCR_BY_USER, lRow);
			bEnterEdit = FALSE;
		}

		lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);
		break;

	case VK_DOWN:
	case VK_NEXT:
	case VK_UP:
	case VK_PRIOR:
		bModifyFocusPaint = IsRow(lRow) && lRow >= 0;
		//if ( bModifyFocusPaint )
		//	m_bNoFocusPaint = TRUE;

		bOldProc = TRUE;
		bNextFocusCellFound = FALSE;
		bNextFocusRowFound = FALSE;
		if (m_bCellMode && (wParam == VK_UP || wParam == VK_DOWN))
		{
			hWndNextCol = m_hWndColFocus;
			lNextRow = GetRowNrFocus();

			bNextFocusCellFound = FindNextCell(&lNextRow, &hWndNextCol, wParam, TRUE);
			if (!bNextFocusCellFound)
			{
				hWndNextCol = NULL;
				lNextRow = TBL_Error;
			}

			bOldProc = FALSE;
		}
		else if (!m_bCellMode && (wParam == VK_UP || wParam == VK_DOWN) && m_pCounter->GetMergeRows() > 0)
		{
			pmr = FindMergeRow(lRow, FMR_ANY);

			if (wParam == VK_UP)
			{
				lNextRow = pmr ? pmr->lRowFrom : lRow;
				bNextFocusRowFound = FindPrevRow(&lNextRow, 0, ROW_Hidden, (lNextRow < 0 ? TBL_MinSplitRow : 0));
			}
			else if (wParam == VK_DOWN)
			{
				lNextRow = pmr ? pmr->lRowTo : lRow;
				bNextFocusRowFound = FindNextRow(&lNextRow, 0, ROW_Hidden, (lNextRow < 0 ? -1 : TBL_MaxRow));
			}

			if (bNextFocusRowFound)
			{
				pmr = FindMergeRow(lNextRow, FMR_ANY);
				if (pmr)
				{
					if (wParam == VK_UP)
					{
						lNextRow = pmr->lRowFrom;
					}
					else if (wParam == VK_DOWN)
					{
						lNextRow = pmr->lRowFrom;

						int iRet;
						POINT pt, ptVis;

						//iRet = MTblGetRowYCoord( m_hWndTbl, pmr->lRowTo, NULL, &pt, &ptVis );
						iRet = GetRowYCoord(pmr->lRowTo, NULL, &pt, &ptVis);
						if (iRet == MTGR_RET_INVISIBLE || iRet == MTGR_RET_PARTLY_VISIBLE)
						{
							m_pCounter->IncNoFocusUpdate();
							Ctd.TblScroll(m_hWndTbl, pmr->lRowTo, NULL, TBL_AutoScroll);
							m_pCounter->DecNoFocusUpdate();
						}
					}
				}

				m_pCounter->IncNoFocusUpdate();
				Ctd.TblSetFocusRow(m_hWndTbl, lNextRow);
				m_pCounter->DecNoFocusUpdate();
			}
			else
				lNextRow = TBL_Error;

			bOldProc = FALSE;
		}
		else if (m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
		{
			bSplitRow = (lRow < 0);

			if (wParam == VK_UP)
			{
				lNextRow = lRow;
				bNextFocusRowFound = FindPrevRow(&lNextRow, 0, ROW_Hidden, (bSplitRow ? TBL_MinSplitRow : 0));
			}
			else if (wParam == VK_DOWN)
			{
				lNextRow = lRow;
				bNextFocusRowFound = FindNextRow(&lNextRow, 0, ROW_Hidden, (bSplitRow ? -1 : TBL_MaxRow));
			}
			else if (wParam == VK_NEXT)
			{
				long lMinRow = TBL_Error, lMaxRow = TBL_Error;
				/*if ( bSplitRow )
				MTblQueryVisibleSplitRange( m_hWndTbl, &lMinRow, &lMaxRow );
				else
				Ctd.TblQueryVisibleRange( m_hWndTbl, &lMinRow, &lMaxRow );*/
				QueryVisRange(lMinRow, lMaxRow, bSplitRow);

				if (lMaxRow != TBL_Error)
				{
					if (IsRow(lMaxRow))
					{
						bNextFocusRowFound = TRUE;
						lNextRow = lMaxRow;
					}
					else
					{
						lNextRow = bSplitRow ? -1 : TBL_MaxRow;
						if (FindPrevRow(&lNextRow, 0, ROW_Hidden, (bSplitRow ? TBL_MinSplitRow : 0)))
							bNextFocusRowFound = (lNextRow != lRow);
					}

					if (bNextFocusRowFound)
					{
						m_pCounter->IncNoFocusUpdate();
						Ctd.TblScroll(m_hWndTbl, lNextRow, NULL, TBL_ScrollTop);
						m_pCounter->DecNoFocusUpdate();
					}
				}
			}
			else if (wParam == VK_PRIOR)
			{
				long lMinRow = TBL_Error, lMaxRow = TBL_Error;
				/*if ( bSplitRow )
				MTblQueryVisibleSplitRange( m_hWndTbl, &lMinRow, &lMaxRow );
				else
				Ctd.TblQueryVisibleRange( m_hWndTbl, &lMinRow, &lMaxRow );*/
				QueryVisRange(lMinRow, lMaxRow, bSplitRow);

				if (lMinRow != TBL_Error && IsRow(lMinRow))
				{
					m_pCounter->IncNoFocusUpdate();
					Ctd.TblScroll(m_hWndTbl, lMinRow, NULL, TBL_ScrollBottom);
					m_pCounter->DecNoFocusUpdate();

					lMinRow = TBL_Error, lMaxRow = TBL_Error;
					/*if ( bSplitRow )
					MTblQueryVisibleSplitRange( m_hWndTbl, &lMinRow, &lMaxRow );
					else
					Ctd.TblQueryVisibleRange( m_hWndTbl, &lMinRow, &lMaxRow );*/
					QueryVisRange(lMinRow, lMaxRow, bSplitRow);

					if (lMinRow != TBL_Error && IsRow(lMinRow) && lMinRow != lRow)
					{
						bNextFocusRowFound = TRUE;
						lNextRow = lMinRow;
					}
				}
			}

			if (bNextFocusRowFound)
			{
				m_pCounter->IncNoFocusUpdate();
				Ctd.TblSetFocusRow(m_hWndTbl, lNextRow);
				m_pCounter->DecNoFocusUpdate();
			}
			else
				lNextRow = TBL_Error;

			bOldProc = FALSE;
		}

		if (bOldProc)
		{
			m_pCounter->IncNoFocusUpdate();
			lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);
			m_pCounter->DecNoFocusUpdate();
		}

		if (m_bCellMode && (wParam == VK_PRIOR || wParam == VK_NEXT))
		{
			HWND hWndFocusCol = m_hWndColFocus;
			lFocusRow = Ctd.TblGetFocusRow(m_hWndTbl);
			if (hWndFocusCol && lFocusRow != TBL_Error && lFocusRow != GetRowNrFocus() && !CanSetFocusToCell(hWndFocusCol, lFocusRow))
			{
				// Zelle im sichtbaren Bereich suchen, auf die der Focus gesetzt werden kann
				long lMinRow, lMaxRow;
				if (Ctd.TblQueryVisibleRange(m_hWndTbl, &lMinRow, &lMaxRow))
				{
					BOOL bFound = FALSE;
					HWND hWndNextCol2 = hWndFocusCol;
					long lNextRow2 = lFocusRow;
					if (FindNextCell(&lNextRow2, &hWndNextCol2, VK_DOWN, TRUE) && lNextRow2 <= lMaxRow)
						bFound = TRUE;
					else if (FindNextCell(&lNextRow2, &hWndNextCol2, VK_UP, TRUE) && lNextRow2 >= lMinRow)
						bFound = TRUE;

					if (bFound)
					{
						lNextRow = lNextRow2;
						bNextFocusCellFound = TRUE;
					}
				}
			}
		}

		if (bModifyFocusPaint)
			m_bNoFocusPaint = FALSE;

		if (QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION))
			Ctd.TblClearSelection(m_hWndTbl);

		// Update Focus
		iRedrawMode = UF_REDRAW_AUTO;

		if (m_bCellMode)
		{
			if (bNextFocusCellFound && lNextRow != GetRowNrFocus())
			{
				m_pCounter->IncNoFocusUpdate();
				Ctd.TblSetFocusRow(m_hWndTbl, lNextRow);
				m_pCounter->DecNoFocusUpdate();
				if (Ctd.TblGetFocusRow(m_hWndTbl) != lNextRow)
					lNextRow = GetRowNrFocus();
			}

			if (m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION)
			{
				iSelChangeMode = UF_SELCHANGE_NONE;
			}
			else if (m_dwCellModeFlags & MTBL_CM_FLAG_SINGLE_SELECTION)
			{
				if (bCtrlPressed)
					iSelChangeMode = UF_SELCHANGE_FOCUS_CELL_ONLY;
				else
					iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
			}
			else
			{
				if (!bShiftPressed)
					m_CellSelRangeKey.SetEmpty();

				if (bCtrlPressed)
					iSelChangeMode = UF_SELCHANGE_NONE;
				else if (bShiftPressed)
					iSelChangeMode = UF_SELCHANGE_FOCUS_ADD_KEY_RANGE;
				else
					iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
			}

			bRemoveNewRowSel = TRUE;
		}
		else if (lNextRow != TBL_Error)
		{
			if (bSingleSelection)
				iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
			else
			{
				if (bCtrlPressed)
					iSelChangeMode = UF_SELCHANGE_NONE;
				else if (bShiftPressed)
				{
					if (wParam == VK_UP || wParam == VK_DOWN)
						iSelChangeMode = UF_SELCHANGE_FOCUS_ADD;
					else
					{
						iSelChangeMode = UF_SELCHANGE_FOCUS_ADD_RANGE;
						/*iSelChangeMode = UF_SELCHANGE_NONE;
						Ctd.TblSetFlagsRowRange( m_hWndTbl, lRow, lNextRow, 0, ROW_Hidden | ROW_Selected, ROW_Selected, TRUE );*/
					}
				}
				else
					iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
			}
		}
		else
			iSelChangeMode = UF_SELCHANGE_NONE;

		UpdateFocus(lNextRow, hWndNextCol, iRedrawMode, iSelChangeMode, UF_FLAG_SCROLL);

		break;

	case VK_END:
	case VK_HOME:
		bNoScroll = m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

		m_pCounter->IncNoFocusUpdate();
		if (bNoScroll)
			m_bNoVScrolling = TRUE;

		lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);

		if (bNoScroll)
			m_bNoVScrolling = FALSE;
		m_pCounter->DecNoFocusUpdate();

		if (bNoScroll)
		{
			lFocusRow = Ctd.TblGetFocusRow(m_hWndTbl);

			m_pCounter->IncNoFocusUpdate();
			Ctd.TblScroll(m_hWndTbl, lFocusRow, NULL, TBL_ScrollBottom);
			m_pCounter->DecNoFocusUpdate();
		}

		if (m_bCellMode)
		{
			HWND hWndFocusCol = m_hWndColFocus;
			lFocusRow = Ctd.TblGetFocusRow(m_hWndTbl);
			if (hWndFocusCol && lFocusRow != TBL_Error && lFocusRow != GetRowNrFocus() && !CanSetFocusToCell(hWndFocusCol, lFocusRow))
			{
				// Zelle im sichtbaren Bereich suchen, auf die der Focus gesetzt werden kann
				long lMinRow, lMaxRow;
				if (Ctd.TblQueryVisibleRange(m_hWndTbl, &lMinRow, &lMaxRow))
				{
					BOOL bFound = FALSE;
					HWND hWndNextCol2 = hWndFocusCol;
					long lNextRow2 = lFocusRow;
					if (FindNextCell(&lNextRow2, &hWndNextCol2, VK_DOWN, TRUE) && lNextRow2 <= lMaxRow)
						bFound = TRUE;
					else if (FindNextCell(&lNextRow2, &hWndNextCol2, VK_UP, TRUE) && lNextRow2 >= lMinRow)
						bFound = TRUE;

					if (bFound)
					{
						lNextRow = lNextRow2;
						bNextFocusCellFound = TRUE;
					}

				}
			}
		}

		if (QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION))
			Ctd.TblClearSelection(m_hWndTbl);

		iRedrawMode = UF_REDRAW_AUTO;

		if (m_bCellMode)
		{
			if (bNextFocusCellFound && lNextRow != GetRowNrFocus())
			{
				m_pCounter->IncNoFocusUpdate();
				Ctd.TblSetFocusRow(m_hWndTbl, lNextRow);
				m_pCounter->DecNoFocusUpdate();
				if (Ctd.TblGetFocusRow(m_hWndTbl) != lNextRow)
					lNextRow = GetRowNrFocus();
			}

			if (m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION)
			{
				iSelChangeMode = UF_SELCHANGE_NONE;
			}
			else if (m_dwCellModeFlags & MTBL_CM_FLAG_SINGLE_SELECTION)
				if (bCtrlPressed)
					iSelChangeMode = UF_SELCHANGE_FOCUS_CELL_ONLY;
				else
					iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
			else
			{
				if (!bShiftPressed)
					m_CellSelRangeKey.SetEmpty();

				if (bCtrlPressed)
					iSelChangeMode = UF_SELCHANGE_NONE;
				else if (bShiftPressed)
					iSelChangeMode = UF_SELCHANGE_FOCUS_ADD_KEY_RANGE;
				else
					iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
			}

			bRemoveNewRowSel = TRUE;
		}
		else
			iSelChangeMode = UF_SELCHANGE_NONE;

		UpdateFocus(TBL_Error, NULL, iRedrawMode, iSelChangeMode, UF_FLAG_SCROLL);

		break;

	case VK_SPACE:
		if (QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION) && !Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection))
			lRet = 0;
		else
			lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);

		break;

	case VK_LEFT:
	case VK_RIGHT:
		if (m_bCellMode)
		{
			HWND hWndNextCol = m_hWndColFocus;
			long lNextRow = GetRowNrFocus();

			bNextFocusCellFound = FindNextCell(&lNextRow, &hWndNextCol, wParam, TRUE);
			if (bNextFocusCellFound)
			{
				m_bNoFocusPaint = FALSE;

				if (lNextRow != GetRowNrFocus())
				{
					m_pCounter->IncNoFocusUpdate();
					Ctd.TblSetFocusRow(m_hWndTbl, lNextRow);
					m_pCounter->DecNoFocusUpdate();
					if (Ctd.TblGetFocusRow(m_hWndTbl) != lNextRow)
						lNextRow = GetRowNrFocus();
				}

				iRedrawMode = UF_REDRAW_AUTO;

				if (m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION)
				{
					iSelChangeMode = UF_SELCHANGE_NONE;
				}
				else if (m_dwCellModeFlags & MTBL_CM_FLAG_SINGLE_SELECTION)
				{
					if (bCtrlPressed)
						iSelChangeMode = UF_SELCHANGE_FOCUS_CELL_ONLY;
					else
						iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
				}
				else
				{
					if (!bShiftPressed)
						m_CellSelRangeKey.SetEmpty();

					if (bCtrlPressed)
						iSelChangeMode = UF_SELCHANGE_NONE;
					else if (bShiftPressed)
						iSelChangeMode = UF_SELCHANGE_FOCUS_ADD_KEY_RANGE;
					else
						iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
				}

				UpdateFocus(lNextRow, hWndNextCol, iRedrawMode, iSelChangeMode, UF_FLAG_SCROLL);
			}
		}
		else
			lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);
		break;

	case VK_TAB:
		// Defect #180: Bei Table Windows führt das Drücken von TAB unter Umständen dazu, dass eine Zelle in den Edit-Modus wechselt, obwohl das nicht der Fall sein darf
		bOldProc = TRUE;
		if (Ctd.GetType(m_hWndTbl) == TYPE_TableWindow && !bCtrlPressed && !bShiftPressed)
		{
			// Sal-Elternfenster ermitteln
			hWndSalParent = Ctd.ParentWindow(m_hWndTbl);

			// Kein Sal-Elternfenster oder Sal-Elternfenster ist ein MDI-Fenster?
			if (!hWndSalParent || Ctd.GetType(hWndSalParent) == TYPE_MDIWindow)
			{
				// Focus "manuell" setzen
				bOldProc = FALSE;
				hWndNextCol = NULL;
				while (FindNextCol(&hWndNextCol, COL_Visible, 0))
				{
					if (CanSetFocusCell(hWndNextCol, lRow))
					{
						SetFocusCell(lRow, hWndNextCol);
						break;
					}
				}
			}
		}

		if (bOldProc)
			lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);

		break;

	default:
		// Alte Fensterprozedur aufrufen
		lRet = CallOldWndProc(WM_KEYDOWN, wParam, lParam);
	}

	if (MustSuppressRowFocusCTD())
		//Ctd.TblKillFocus( m_hWndTbl );
		KillFocusFrameI();

	if (bTrapColSelChanges)
		EndColSelTrap();

	if (bTrapRowSelChanges)
		EndRowSelTrap(lRowSelTrapRow, pvSelRowsAfter, bRemoveNewRowSel);

	if (bLockUpdate)
		LockWindowUpdate(NULL);

	if (bCheckFocusRowRedraw && !bLockUpdate)
	{
		// Wenn sich der Zeilenfokus geändert hat und sich die Selektion der Fokuszeile
		// nicht geändert hat, die Fokuszeile neu zeichnen, da sonst u.U. Elemente, die
		// nicht invertiert sein sollen, trotzdem invertiert sind.
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);
		if (lFocusRow != lFocusRowBefore)
		{
			CMTblRow *pRow = m_Rows->GetRowPtr(lFocusRow);
			if (pRow)
			{
				BOOL bSelectedBefore = find(vSelRowsBefore.begin(), vSelRowsBefore.end(), pRow) != vSelRowsBefore.end();
				BOOL bSelectedAfter = find(vSelRowsAfter.begin(), vSelRowsAfter.end(), pRow) != vSelRowsAfter.end();
				if (bSelectedBefore && bSelectedAfter)
				{
					if (InvalidateRow(lFocusRow, FALSE))
					{
						UpdateWindow(m_hWndTbl);
					}
				}
			}
		}
	}

	// Control-Simulation beenden
	if (bCtrlSim)
	{
		BYTE ks[256];
		if (GetKeyboardState(ks))
		{
			ks[VK_CONTROL] &= 0x00;
			SetKeyboardState(ks);
		}
	}

	return lRet;
}

// OnKillEdit
/*LRESULT MTBLSUBCLASS::OnKillEdit( WPARAM wParam, LPARAM lParam )
{
BOOL bSuppressRowSelection = QueryFlags( MTBL_FLAG_SUPPRESS_ROW_SELECTION ), bTrapRowSelChanges;
HWND hWndCol = NULL;
long lRow = TBL_Error;

Ctd.TblQueryFocus( m_hWndTbl, &lRow, &hWndCol );
if ( !hWndCol )
return 1;

bTrapRowSelChanges = MustTrapRowSelChanges( );
if ( bTrapRowSelChanges )
BeginRowSelTrap( );

m_bNoFocusPaint = FALSE;

m_pCounter->IncNoFocusUpdate( );
long lRet = CallOldWndProc( TIM_KillEdit, wParam, lParam );
m_pCounter->DecNoFocusUpdate( );

if ( MustSuppressRowFocusCTD( ) )
{
Ctd.TblKillFocus( m_hWndTbl );
InvalidateRowFocus( lRow, FALSE );
//UpdateWindow( m_hWndTbl );
}

if ( ( m_bCellMode && (m_dwCellModeFlags & MTBL_CM_FLAG_KILL_EDIT_NO_ROWSEL) ) || bSuppressRowSelection )
Ctd.TblSetRowFlags( m_hWndTbl, lRow, ROW_Selected, FALSE );

if ( bTrapRowSelChanges )
EndRowSelTrap( );

UpdateFocus( );
//UpdateWindow( m_hWndTbl );

return lRet;
}*/
LRESULT MTBLSUBCLASS::OnKillEdit(WPARAM wParam, LPARAM lParam)
{
	HWND hWndCol = NULL;
	long lRow = TBL_Error;
	Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndCol);
	if (!hWndCol)
		return 1;

	BeforeKillEdit();
	long lRet = CallOldWndProc(TIM_KillEdit, wParam, lParam);
	AfterKillEdit();

	return lRet;
}

// OnKillFocus
LRESULT MTBLSUBCLASS::OnKillFocus(WPARAM wParam, LPARAM lParam)
{
	BOOL bUpdateFocus = FALSE;
	if (!m_bSettingFocus && !IsMyWindow(HWND(wParam)))
	{
		m_bNoFocusPaint = TRUE;
		bUpdateFocus = TRUE;
	}

	// Zeilenselektion per Maus beendet
	m_bRowSelByMouseClick = FALSE;
	m_iRowSelByMouse = 0;
	if (m_bTempSingleSelection)
	{
		Ctd.TblSetTableFlags(m_hWndTbl, TBL_Flag_SingleSelection, FALSE);
		m_bTempSingleSelection = FALSE;
	}

	long lRet = CallOldWndProc(WM_KILLFOCUS, wParam, lParam);

	if (bUpdateFocus)
		UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);

	return lRet;
}

// OnKillFocusFrame
LRESULT MTBLSUBCLASS::OnKillFocusFrame(WPARAM wParam, LPARAM lParam)
{
	if (!m_bKillFocusFrameI)
		m_bNoFocusPaint = TRUE;

	long lRet = CallOldWndProc(TIM_KillFocusFrame, wParam, lParam);

	if (!m_bKillFocusFrameI)
		UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);

	return lRet;
}

// OnLButtonDblClk
LRESULT MTBLSUBCLASS::OnLButtonDblClk(WPARAM wParam, LPARAM lParam)
{
	BOOL bDrill = FALSE, bOldProc = TRUE;
	DWORD dwFlags, dwMFlags;
	HWND hWndCol;
	long lRet, lRow;
	POINT pt;

	pt.x = LOWORD(lParam);
	pt.y = HIWORD(lParam);
	if (ObjFromPoint(pt, &lRow, &hWndCol, &dwFlags, &dwMFlags))
	{
		// Klick auf Zeile?
		if (lRow != TBL_Error && !(dwFlags & TBL_YOverColumnHeader || dwFlags & TBL_YOverSplitBar))
		{
			// Über Knoten?
			if (dwMFlags & MTOFP_OVERNODE)
			{
				if (Ctd.TblSetFocusRow(m_hWndTbl, lRow))
				{
					bDrill = TRUE;
					bOldProc = FALSE;

					if (dwFlags & TBL_XOverRowHeader)
					{
						PostMessage(m_hWndTbl, SAM_RowHeaderDoubleClick, 0, lRow);
					}
					else
					{
						PostMessage(hWndCol, SAM_DoubleClick, 0, lRow);
						PostMessage(m_hWndTbl, SAM_DoubleClick, 0, lRow);
					}
				}
			}
			else
			{
				bDrill = m_Rows->QueryRowFlags(lRow, MTBL_ROW_CANEXPAND) && QueryFlags(MTBL_FLAG_EXPAND_ROW_ON_DBLCLK);
				if (bDrill && (GetRowHdrFlags() & TBL_RowHdr_AllowRowSizing) && (dwMFlags & MTOFP_OVERROWHDRSEP))
					bDrill = FALSE;
			}
		}
	}

	// Alte Fensterprozedur aufrufen
	if (bOldProc)
	{
		lRet = CallOldWndProc(WM_LBUTTONDBLCLK, wParam, lParam);
		if (bDrill)
		{
			if (Ctd.TblGetFocusRow(m_hWndTbl) != lRow)
				bDrill = FALSE;
		}
	}
	else
		lRet = 0;

	// Auf/Zuklappen?
	if (bDrill)
	{
		if (m_Rows->QueryRowFlags(lRow, MTBL_ROW_ISEXPANDED))
			SendMessage(m_hWndTbl, MTM_CollapseRow, MTCR_REDRAW | MTCR_BY_USER, lRow);
		else
			SendMessage(m_hWndTbl, MTM_ExpandRow, MTER_REDRAW | MTER_BY_USER, lRow);
	}

	// Erweiterte Nachrichten
	if (m_bExtMsgs)
	{
		// XY-Position ermitteln
		iLDblClkX = LOWORD(lParam);
		iLDblClkY = HIWORD(lParam);
		// Nachricht generieren
		GenerateExtMsgs(WM_LBUTTONDBLCLK);
		// Letzte LDown-Nachricht merken
		uLastLDownMsg = WM_LBUTTONDBLCLK;
	}

	return lRet;
}


// OnLButtonDown
LRESULT MTBLSUBCLASS::OnLButtonDown(WPARAM wParam, LPARAM lParam)
{
	if (SendMessage(m_hWndTbl, MTM_LButtonDown, wParam, lParam) == MTBL_NODEFAULTACTION)
		return 0;

	BOOL bBtnClick, bCanSetFocusToCell = TRUE, bCellClick, bDDBtnClick, bCellSelectClick, bCenter, bColMoveClick = FALSE, bColSelectClick = FALSE, bColSizeClick = FALSE, bCtrlPressed, bEmptyRowsClick, bEnable = FALSE, bFixColJustify = FALSE, bHyperlinkClick = FALSE, bKillRowFocus = FALSE, bLockUpdate = FALSE, bNodeClick, bNoColHdr, bOverColHdr, bOverCorner, bOverMergedCell, bOverRow, bRight, bRowClick, bRowHeights, bRowHdrSizeClick = FALSE, bRowSelectClick, bShiftPressed, bRowSizeClick, /*bSetFocus = FALSE,*/ bOldProc = TRUE, bRemoveSelection = FALSE, bSelCellRange = FALSE, bSingleSelection, bSetFocusRet = TRUE, bSplitBarDragClick = FALSE, bSuppressFocus, bSuppressRowSelection, bTrapColSelChanges = FALSE, bTrapRowSelChanges = FALSE, bUpdateFocus = FALSE, bValidate;
	DWORD dwFlags, dwMFlags;
	HWND hWndCol, hWndFocusCol = NULL, hWndMergeCol = NULL;
	long lFocusRow = TBL_Error, lMergeRow = TBL_Error, lRet, lRow, lRowSelTrapRow = TBL_Error;
	MTBLCELLRANGE cr;
	POINT pt;

	pt.x = LOWORD(lParam);
	pt.y = HIWORD(lParam);

	if (!ObjFromPoint(pt, &lRow, &hWndCol, &dwFlags, &dwMFlags))
		return CallOldWndProc(WM_LBUTTONDOWN, wParam, lParam);

	bRowHeights = m_pCounter->GetRowHeights() > 0;
	bNoColHdr = QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	bCtrlPressed = HIBYTE(GetKeyState(VK_CONTROL));
	bShiftPressed = HIBYTE(GetKeyState(VK_SHIFT));

	if (m_bCellMode)
		bCanSetFocusToCell = CanSetFocusToCell(hWndCol, lRow);

	BOOL bRestoreRowSizingAllowed = FALSE;
	BOOL bRowSizingAllowed = (GetRowHdrFlags() & TBL_RowHdr_AllowRowSizing) != 0;
	long lSepRow = TBL_Error;
	if (bRowSizingAllowed)
	{
		if (dwMFlags & MTOFP_OVERROWHDRSEP)
		{
			lSepRow = GetSepRow(pt);
		}
		else if (dwMFlags & MTOFP_OVERMERGEDROWHDRSEP)
		{
			SetRowHdrFlags(TBL_RowHdr_AllowRowSizing, FALSE);
			bRestoreRowSizingAllowed = TRUE;
		}
	}

	bOverColHdr = dwFlags & TBL_YOverColumnHeader;
	bOverCorner = bOverColHdr && (dwFlags & TBL_XOverRowHeader);
	bOverMergedCell = dwMFlags & MTOFP_OVERMERGEDCELL;
	bOverRow = (lRow != TBL_Error) && !(dwFlags & TBL_YOverColumnHeader || dwFlags & TBL_YOverSplitBar);
	bRowSizeClick = lSepRow != TBL_Error;
	bRowClick = bOverRow && !bRowSizeClick;
	bCellClick = bRowClick && hWndCol;
	bBtnClick = /*bCellClick &&*/ (dwMFlags & MTOFP_OVERBTN);
	bDDBtnClick = bCellClick && (dwMFlags & MTOFP_OVERDDBTN);
	bNodeClick = bCellClick && (dwMFlags & MTOFP_OVERNODE);
	bHyperlinkClick = bCellClick && !bNodeClick && !bBtnClick && !bDDBtnClick && IsOverHyperlink(hWndCol, lRow, dwMFlags);
	bCellSelectClick = m_bCellMode && bCellClick && !bNodeClick && !bHyperlinkClick && !bBtnClick && !bDDBtnClick && bCanSetFocusToCell && !CanSetFocusCell(hWndCol, lRow);
	bRowSelectClick = ((bRowClick && !bCellClick) || (bCellClick && !m_bCellMode && !CanSetFocusCell(hWndCol, lRow))) && !bNodeClick && !bHyperlinkClick && !bBtnClick && !bDDBtnClick;
	bSplitBarDragClick = (dwFlags & TBL_YOverSplitBar) && !bRowSizeClick;
	bEmptyRowsClick = !(bOverColHdr || bOverRow || bRowSizeClick || bSplitBarDragClick);

	if (bOverColHdr && !bBtnClick)
	{
		HSTRING hs = 0;
		HWND hWndCol;
		int iWid;
		WORD wFlags;
		if (Ctd.TblQueryRowHeader(m_hWndTbl, &hs, 65536, &iWid, &wFlags, &hWndCol) && (wFlags & TBL_RowHdr_Visible) && (wFlags & TBL_RowHdr_Sizable))
		{
			if (pt.x >= iWid - 3 && pt.x <= iWid + 3)
				bRowHdrSizeClick = TRUE;
		}

		if (hs)
			Ctd.HStringUnRef(hs);
	}

	HWND hWndSepCol = GetSepCol(LOWORD(lParam), HIWORD(lParam));
	bColSizeClick = bOverColHdr && hWndSepCol && !bBtnClick;

	if (bOverColHdr && hWndCol && Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_MovableCols) && !bColSizeClick && !bRowHdrSizeClick && !IsLockedColumn(hWndCol) && !bBtnClick)
	{
		RECT r;
		MTblGetColHdrRect(hWndCol, LPINT(&r.left), LPINT(&r.top), LPINT(&r.right), LPINT(&r.bottom));

		if ((pt.y >= r.bottom - 5))
			bColMoveClick = TRUE;
	}

	// DragDrop
	if (!m_bDragCanAutoStart && !bSplitBarDragClick && !bRowSizeClick && !bRowHdrSizeClick && !bColMoveClick && !bColSizeClick && SendMessage(m_hWndTbl, MTM_DragCanAutoStart, pt.x, pt.y))
	{
		DragDropAutoStart(wParam, lParam);
		return 0;
	}

	// Validieren
	Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);
	bValidate = hWndFocusCol && !bOverCorner && (bOverColHdr || bCellClick || bRowClick || bRowSizeClick);
	if (bValidate && !Validate(hWndFocusCol, hWndCol))
		return 0;

	// Bereich Zellenselektion/Tastatur zurücksetzen
	m_CellSelRangeKey.SetEmpty();

	bTrapColSelChanges = MustTrapColSelChanges() && !bNodeClick && !bHyperlinkClick && !bBtnClick && !bDDBtnClick;
	if (bTrapColSelChanges)
		BeginColSelTrap();

	bSuppressFocus = MustSuppressRowFocusCTD();
	bSuppressRowSelection = QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION);
	bTrapRowSelChanges = MustTrapRowSelChanges() && !bNodeClick && !bHyperlinkClick && !bBtnClick && !bDDBtnClick;

	bSingleSelection = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection);

	if (bTrapRowSelChanges)
	{
		if (bRowSelectClick && bCtrlPressed && !bSingleSelection)
			lRowSelTrapRow = lRow;
		BeginRowSelTrap(lRowSelTrapRow);
	}

	// Klick auf Zeile?
	if (bRowClick)
	{
		m_bNoFocusPaint = FALSE;

		if (bSuppressRowSelection)
			bRemoveSelection = TRUE;

		if (bSuppressFocus)
			bKillRowFocus = TRUE;
	}

	// Select-Klick Zeile?
	if (bRowSelectClick)
	{
		if (bRowHeights || bNoColHdr)
		{
			bOldProc = FALSE;
			bUpdateFocus = TRUE;

			m_pCounter->IncNoFocusUpdate();
			bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lRow);
			m_pCounter->DecNoFocusUpdate();

			if (bSetFocusRet)
			{
				if (dwFlags & TBL_XOverRowHeader)
					SendMessage(m_hWndTbl, SAM_RowHeaderClick, 0, lRow);
				else
				{
					if (hWndCol)
						SendMessage(hWndCol, SAM_Click, 0, lRow);
					SendMessage(m_hWndTbl, SAM_Click, 0, lRow);
				}
			}
		}
		else
			bUpdateFocus = m_bCellMode;
	}

	// Klick Zeilengröße?
	if (bRowSizeClick)
	{
		m_bRowSizing = TRUE;
		BOOL bVarRowHt = QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT);
		if (bVarRowHt || bNoColHdr)
		{
			bOldProc = FALSE;

			m_lRowSizing = lSepRow;

			m_lRowSizingOffsY = 0;
			POINT ptRow, ptRowVis;
			int iRet = GetRowYCoord(m_lRowSizing, NULL, &ptRow, &ptRowVis, TRUE);
			if (iRet != MTGR_RET_ERROR)
				m_lRowSizingOffsY = pt.y - ptRow.y;

			Ctd.TblKillEdit(m_hWndTbl);

			if (bNoColHdr && !bVarRowHt)
			{
				Ctd.TblQueryLinesPerRow(m_hWndTbl, &m_iRowSizingLinesPerRow);
				DrawLinesPerRowIndicators(m_iRowSizingLinesPerRow);
			}

			SetCapture(m_hWndTbl);

		}
	}

	// Klick auf Zelle?
	if (bCellClick)
	{
		// Button-Click?
		if (bBtnClick)
		{
			bOldProc = FALSE;
			bUpdateFocus = TRUE;

			HWND hWndBtnCol = hWndCol;
			long lBtnRow = lRow;
			if (bOverMergedCell && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
			{
				hWndBtnCol = hWndMergeCol;
				lBtnRow = lMergeRow;
			}

			m_pCounter->IncNoFocusUpdate();
			bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lBtnRow);
			m_pCounter->DecNoFocusUpdate();

			if (bSetFocusRet)
			{
				SendMessage(hWndCol, SAM_Click, 0, lRow);
				SendMessage(m_hWndTbl, SAM_Click, 0, lRow);

				CMTblBtn *pBtn = GetEffCellBtn(hWndBtnCol, lBtnRow);

				if (!(pBtn->m_dwFlags & MTBL_BTN_DISABLED))
				{
					CMTblRow *pRow = m_Rows->GetRowPtr(lBtnRow);
					if (pRow)
					{
						pRow->m_Cols->SetBtnPushed(hWndBtnCol, TRUE);
						m_hWndColBtnPushed = hWndBtnCol;
						m_lRowBtnPushed = lBtnRow;
						InvalidateBtn(hWndBtnCol, lBtnRow, NULL, FALSE);
						SetCapture(m_hWndTbl);
					}
				}
			}
		}
		// Node-Click?
		else if (bNodeClick)
		{
			bOldProc = FALSE;
			bUpdateFocus = TRUE;

			long lRowMsg = lRow;
			if (bOverMergedCell && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
				lRowMsg = lMergeRow;

			m_pCounter->IncNoFocusUpdate();
			bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lRowMsg);
			m_pCounter->DecNoFocusUpdate();

			if (bSetFocusRet)
			{
				SendMessage(hWndCol, SAM_Click, 0, lRow);
				SendMessage(m_hWndTbl, SAM_Click, 0, lRow);

				if (m_Rows->QueryRowFlags(lRowMsg, MTBL_ROW_ISEXPANDED))
					SendMessage(m_hWndTbl, MTM_CollapseRow, MTCR_REDRAW | MTCR_BY_USER, lRowMsg);
				else
					SendMessage(m_hWndTbl, MTM_ExpandRow, MTER_REDRAW | MTER_BY_USER, lRowMsg);
			}
		}
		// Hyperlink-Click?
		else if (bHyperlinkClick)
		{
			bOldProc = FALSE;
			bUpdateFocus = TRUE;

			m_pCounter->IncNoFocusUpdate();
			bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lRow);
			m_pCounter->DecNoFocusUpdate();

			if (bSetFocusRet)
			{
				SendMessage(hWndCol, SAM_Click, 0, lRow);
				SendMessage(m_hWndTbl, SAM_Click, 0, lRow);

				HWND hWndMsg = hWndCol;
				long lRowMsg = lRow;
				if (bOverMergedCell && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
				{
					hWndMsg = hWndMergeCol;
					lRowMsg = lMergeRow;
				}
				SendMessage(m_hWndTbl, MTM_HyperlinkClick, WPARAM(hWndMsg), lRowMsg);
			}
		}
		else
		{
			// Zelle gesperrt?			
			if (!CanSetFocusCell(hWndCol, lRow))
			{
				bUpdateFocus = TRUE;

				Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol);
				if (hWndFocusCol && lFocusRow != lRow)
				{
					KillEditNeutral();
				}
				if (Ctd.IsWindowEnabled(hWndCol))
				{
					Ctd.DisableWindow(hWndCol);
					bEnable = TRUE;
				}

				if (m_bCellMode)
				{
					bOldProc = FALSE;

					m_pCounter->IncNoFocusUpdate();
					bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lRow);
					m_pCounter->DecNoFocusUpdate();

					if (bSetFocusRet)
					{
						SendMessage(hWndCol, SAM_Click, 0, lRow);
						SendMessage(m_hWndTbl, SAM_Click, 0, lRow);

						if (bShiftPressed && bCanSetFocusToCell && !(m_dwCellModeFlags & MTBL_CM_FLAG_SINGLE_SELECTION) && !(m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION))
						{
							cr.hWndColFrom = m_hWndColFocus;
							cr.lRowFrom = GetRowNrFocus();

							cr.hWndColTo = hWndCol;
							cr.lRowTo = lRow;

							bSelCellRange = cr.IsValid();
						}
					}
				}
			}
			else
			{
				bOldProc = FALSE;

				HWND hWndFocusCol = hWndCol;
				long lFocusRow = lRow;

				if (bOverMergedCell && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
				{
					hWndFocusCol = hWndMergeCol;
					lFocusRow = lMergeRow;
				}

				m_bMouseClickSim = TRUE;
				bSetFocusRet = Ctd.TblSetFocusCell(m_hWndTbl, lFocusRow, hWndFocusCol, 0, -1);
			}
		}
	}

	// Size-Click Row-Header?
	if (bRowHdrSizeClick)
		m_bRowHdrSizing = TRUE;

	// Klick auf Column-Header?
	if (bOverColHdr && hWndCol && !(Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SelectableCols)) && !hWndSepCol && !bRowHdrSizeClick && !bBtnClick)
	{
		RECT r;
		MTblGetColHdrRect(hWndCol, LPINT(&r.left), LPINT(&r.top), LPINT(&r.right), LPINT(&r.bottom));

		if ((pt.y < r.bottom - 5) || IsLockedColumn(hWndCol) || !Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_MovableCols))
		{
			HWND hWndTmp;
			long lRow = TBL_Error;
			Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndTmp);
			if (IsRow(lRow))
			{
				BOOL bEnabled = Ctd.IsWindowEnabled(hWndCol);
				BOOL bCanSetFocusCell = CanSetFocusCell(hWndCol, lRow);
				if (bEnabled)
				{
					if (!bCanSetFocusCell)
					{
						Ctd.DisableWindow(hWndCol);
						bEnable = TRUE;
					}
					else
					{
						bOldProc = FALSE;

						HWND hWndFocusCol = hWndCol;
						long lFocusRow = lRow;

						if (GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
						{
							hWndFocusCol = hWndMergeCol;
							lFocusRow = lMergeRow;
						}

						bSetFocusRet = Ctd.TblSetFocusCell(m_hWndTbl, lRow, hWndCol, 0, -1);

						if (bSetFocusRet)
						{
							SendMessage(hWndFocusCol, SAM_Click, 0, lFocusRow);
							SendMessage(m_hWndTbl, SAM_Click, 0, lFocusRow);
						}

						/*if ( GetMergeCellEx( hWndCol, lRow, hWndMergeCol, lMergeRow ) )
						{
						bSetFocusRet = Ctd.TblSetFocusCell( m_hWndTbl, lMergeRow, hWndMergeCol, 0, -1 );

						if ( bSetFocusRet )
						{
						SendMessage( hWndMergeCol, SAM_Click, 0, lMergeRow );
						SendMessage( m_hWndTbl, SAM_Click, 0, lRow );
						}
						}
						else
						{
						bSetFocusRet = Ctd.TblSetFocusCell( m_hWndTbl, lRow, hWndCol, 0, -1 );

						if ( bSetFocusRet )
						{
						SendMessage( hWndCol, SAM_Click, 0, lRow );
						SendMessage( m_hWndTbl, SAM_Click, 0, lRow );
						}
						}*/
					}
				}

				if ((!bEnabled || !bCanSetFocusCell) && bSuppressFocus)
					bKillRowFocus = TRUE;
			}
		}
	}

	// Select-Klick Column-Header?
	if (bOverColHdr && hWndCol && Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SelectableCols) && !hWndSepCol && !bRowHdrSizeClick && !bShiftPressed && !bBtnClick)
	{
		RECT r;
		MTblGetColHdrRect(hWndCol, LPINT(&r.left), LPINT(&r.top), LPINT(&r.right), LPINT(&r.bottom));

		if ((pt.y < r.bottom - 5) || IsLockedColumn(hWndCol) || !Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_MovableCols))
		{
			bColSelectClick = TRUE;

			// Evtl. Edit-Modus beenden
			if (Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol) && hWndFocusCol)
			{
				//Ctd.TblKillEdit( m_hWndTbl );
				bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lFocusRow);
				if (bSetFocusRet)
					KillEditNeutral();
				else
					bOldProc = FALSE;
				//Ctd.TblSetRowFlags( m_hWndTbl, lFocusRow, ROW_Selected, FALSE );
			}

			// Fokus nicht zeichnen
			m_bNoFocusPaint = TRUE;

			// Fokus aktualisieren?
			if (m_bCellMode)
				bUpdateFocus = TRUE;

			// Fix Spaltenausrichtung
			if (m_bFixColJustifyBug)
			{
				bCenter = Ctd.TblQueryColumnFlags(hWndCol, COL_CenterJustify);
				bRight = Ctd.TblQueryColumnFlags(hWndCol, COL_RightJustify);
				bFixColJustify = (bCenter || bRight);
			}
		}
	}

	// Move-Click auf Column-Header?
	if (bColMoveClick)
	{
		m_bColMoving = TRUE;

		if (Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol) && hWndFocusCol)
		{
			bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lFocusRow);
			if (bSetFocusRet)
			{
				KillEditNeutral();
				//Ctd.TblSetRowFlags( m_hWndTbl, lFocusRow, ROW_Selected, TRUE );
			}
			else
				bOldProc = FALSE;
		}
	}

	// Size-Click auf Column-Header?
	if (bColSizeClick)
	{
		m_bColSizing = TRUE;

		/*if ( Ctd.TblQueryFocus( m_hWndTbl, &lFocusRow, &hWndFocusCol ) && hWndFocusCol )
		{
		bSetFocusRet = Ctd.TblSetFocusRow( m_hWndTbl, lFocusRow );
		if ( bSetFocusRet )
		{
		KillEditNeutral( );
		}
		else
		bOldProc = FALSE;

		}*/

		Ctd.TblKillEdit(m_hWndTbl);
		if (Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol) && hWndFocusCol)
			bOldProc = FALSE;
	}

	// Button-Click in Column-Header?
	if (bBtnClick && bOverColHdr)
	{
		bOldProc = FALSE;

		SendMessage(hWndCol, SAM_Click, 0, TBL_Error);
		SendMessage(m_hWndTbl, SAM_Click, 0, TBL_Error);

		CMTblBtn *pBtn = m_Cols->GetHdrBtn(hWndCol);

		if (!(pBtn->m_dwFlags & MTBL_BTN_DISABLED))
		{
			m_Cols->SetHdrBtnPushed(hWndCol, TRUE);
			m_hWndColBtnPushed = hWndCol;
			m_lRowBtnPushed = TBL_Error;
			InvalidateBtn(hWndCol, TBL_Error, NULL, FALSE);
			SetCapture(m_hWndTbl);
		}
	}


	// Node-Click auf Row-Header?
	if ((dwMFlags & MTOFP_OVERNODE) && (dwFlags & TBL_XOverRowHeader))
	{
		bOldProc = FALSE;

		bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, lRow);
		if (bSetFocusRet)
		{
			bNodeClick = TRUE;

			SendMessage(m_hWndTbl, SAM_RowHeaderClick, 0, lRow);

			if (m_Rows->QueryRowFlags(lRow, MTBL_ROW_ISEXPANDED))
				SendMessage(m_hWndTbl, MTM_CollapseRow, MTCR_REDRAW | MTCR_BY_USER, lRow);
			else
				SendMessage(m_hWndTbl, MTM_ExpandRow, MTER_REDRAW | MTER_BY_USER, lRow);
		}
	}

	// Split-Bar-Drag-Click?
	if (bSplitBarDragClick)
		m_bSplitBarDragging = TRUE;

	// Klick auf den Bereich der leeren Zeilen
	if (bEmptyRowsClick)
	{
		if (bRowHeights || bNoColHdr)
		{
			bOldProc = FALSE;
		}
	}

	// Alte Fensterprozedur aufrufen?
	if (bOldProc)
	{
		int iValidate = VALIDATE_Ok;

		/*if ( bSetFocus )
		iValidate = BeforeSetFocusCell( hWndCol, lRow, TRUE );*/

		bLockUpdate = (bRowSelectClick || bColSelectClick) && MustRedrawSelections() && !gbCtdHooked;
		if (bLockUpdate)
			bLockUpdate = LockWindowUpdate(m_hWndTbl);

		if ( /*bSetFocus ||*/ bUpdateFocus)
			m_pCounter->IncNoFocusUpdate();

		if (iValidate != VALIDATE_Cancel)
		{
			// Verhindern, dass Zeilenhöhe verändert wird wenn Maus über Split-Bar und Row-Separator im Row-Header
			if (dwFlags & TBL_YOverSplitBar && dwFlags & TBL_XOverRowHeader)
			{
				CMTblMetrics tm(m_hWndTbl);
				if (pt.x <= tm.m_RowHeaderRight)
				{
					lParam = MAKELONG((tm.m_RowHeaderRight + 1), pt.y);
				}
			}
			lRet = CallOldWndProc(WM_LBUTTONDOWN, wParam, lParam);
		}
		else
			lRet = 0;

		long lFocusRow = Ctd.TblGetFocusRow(m_hWndTbl);
		if (lFocusRow != lRow)
			bSetFocusRet = FALSE;

		if ( /*bSetFocus ||*/ bUpdateFocus)
			m_pCounter->DecNoFocusUpdate();

		if (bLockUpdate)
			LockWindowUpdate(NULL);

		/*if ( bSetFocus )
		AfterSetFocusCell( hWndCol, lRow, FALSE );*/
	}
	else
		lRet = 0;

	if (bFixColJustify)
	{
		DWORD dwFlags = 0;
		if (bCenter)
			dwFlags |= COL_CenterJustify;
		if (bRight)
			dwFlags |= COL_RightJustify;
		Ctd.TblSetColumnFlags(hWndCol, dwFlags, TRUE);
	}

	if (bEnable)
		Ctd.EnableWindow(hWndCol);

	if (bRestoreRowSizingAllowed)
		SetRowHdrFlags(TBL_RowHdr_AllowRowSizing, TRUE);

	/*if ( bRemoveSelection )
	Ctd.TblClearSelection( m_hWndTbl );*/

	// Zeilenselektion per Maus?
	/*if ( bRowSelectClick )
	{
	m_bRowSelByMouseClick = TRUE;
	m_ptRowSelByMouseClick = pt;

	if ( m_bCellMode )
	bUpdateFocus = TRUE;
	}
	else
	{
	m_bRowSelByMouseClick = FALSE;
	}*/

	/*if ( bKillRowFocus )
	KillFocusFrameI( );

	if ( bTrapColSelChanges )
	EndColSelTrap( );

	if ( bTrapRowSelChanges )
	EndRowSelTrap( lRowSelTrapRow );*/

	if (bUpdateFocus)
	{
		int iRedrawMode = UF_REDRAW_AUTO;
		int iSelChangeMode = UF_SELCHANGE_NONE;
		DWORD dwFlags = 0;

		if (bRowSelectClick && (bRowHeights || bNoColHdr))
		{
			if (!bRemoveSelection)
			{
				if (bSingleSelection)
				{
					iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
				}
				else
				{
					if (bCtrlPressed)
						iSelChangeMode = UF_SELCHANGE_FOCUS_SWITCH;
					else if (bShiftPressed)
						iSelChangeMode = UF_SELCHANGE_FOCUS_ADD_RANGE;
					else
						iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
				}
			}

			dwFlags = UF_FLAG_SELCHANGE_ROWS;

			if (m_bCellMode && !bCtrlPressed && !bShiftPressed)
				ClearCellSelection(TRUE);
		}
		else if (m_bCellMode)
		{
			if (bCellClick && !bNodeClick && !bHyperlinkClick && !bBtnClick && !bDDBtnClick)
			{
				if ((m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION) || !bCanSetFocusToCell)
				{
					iSelChangeMode = UF_SELCHANGE_NONE;
				}
				else if (m_dwCellModeFlags & MTBL_CM_FLAG_SINGLE_SELECTION)
				{
					if (bCtrlPressed)
						iSelChangeMode = UF_SELCHANGE_FOCUS_CELL_ONLY;
					else
						iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
				}
				else
				{
					if (bShiftPressed || bCtrlPressed)
					{
						if (bCtrlPressed)
							iSelChangeMode = UF_SELCHANGE_FOCUS_SWITCH;
						else
							iSelChangeMode = UF_SELCHANGE_FOCUS_ADD;
					}
					else
						iSelChangeMode = UF_SELCHANGE_FOCUS_ONLY;
				}
			}
			else if (bRowSelectClick && !bCtrlPressed && !bShiftPressed)
			{
				iSelChangeMode = UF_SELCHANGE_CLEAR_CELLS;
			}
			else if (bColSelectClick && !bCtrlPressed)
			{
				iSelChangeMode = UF_SELCHANGE_CLEAR_CELLS;
			}
		}

		if (bSetFocusRet)
			UpdateFocus(lRow, hWndCol, iRedrawMode, iSelChangeMode, dwFlags);
		else
			UpdateFocus(TBL_Error, NULL, iRedrawMode, iSelChangeMode, dwFlags);
	}

	if (bKillRowFocus)
		KillFocusFrameI();

	if (bRemoveSelection)
		Ctd.TblClearSelection(m_hWndTbl);

	if (bTrapColSelChanges)
		EndColSelTrap();

	if (bTrapRowSelChanges)
		EndRowSelTrap(lRowSelTrapRow);

	if (bSelCellRange)
	{
		if (!bCtrlPressed)
			ClearCellSelection(TRUE);
		AdaptCellRange(cr);
		SetCellRangeFlags(&cr, MTBL_CELL_FLAG_SELECTED, TRUE, TRUE);
	}

	// Zeilenselektion per Maus?
	if (bRowSelectClick)
	{
		m_bRowSelByMouseClick = TRUE;
		m_ptRowSelByMouseClick = pt;

		if (bRowHeights || bNoColHdr)
		{
			long lMergeRow = GetMergeRow(lRow);
			m_RowSelRange.lRowFrom = m_RowSelRange.lRowTo = (lMergeRow != TBL_Error ? lMergeRow : lRow);
			SetCapture(m_hWndTbl);
		}
	}
	else
	{
		m_bRowSelByMouseClick = FALSE;
	}

	// Zellenselektion per Maus?
	if (bCellSelectClick && !(m_dwCellModeFlags & MTBL_CM_FLAG_SINGLE_SELECTION) && !(m_dwCellModeFlags & MTBL_CM_FLAG_NO_SELECTION))
	{
		m_bCellSelByMouseClick = TRUE;

		m_CellSelRange.hWndColFrom = m_CellSelRange.hWndColTo = hWndCol;
		m_CellSelRange.lRowFrom = m_CellSelRange.lRowTo = lRow;

		m_CellSelRange.CopyTo(m_CellSelRangeA);
		AdaptCellRange(m_CellSelRangeA);

		SetCapture(m_hWndTbl);
	}
	else
	{
		m_bCellSelByMouseClick = FALSE;
	}

	// Erweiterte Nachrichten
	if (m_bExtMsgs)
	{
		// XY-Position ermitteln
		iLDownX = LOWORD(lParam);
		iLDownY = HIWORD(lParam);
		// Nachricht generieren
		GenerateExtMsgs(WM_LBUTTONDOWN);

		// Letzte LDown-Nachricht merken
		uLastLDownMsg = WM_LBUTTONDOWN;
	}

	return lRet;
}


// OnLButtonUp
LRESULT MTBLSUBCLASS::OnLButtonUp(WPARAM wParam, LPARAM lParam)
{
	// DragDrop
	if (m_bDragCanAutoStart)
		DragDropAutoStartCancel();

	// Aktuelle Spalten ungültig?
	if (m_bColMoving || m_bColSizing)
		m_CurCols->SetInvalid();

	BOOL bRemoveSelection = FALSE, bUpdateFocus = FALSE;

	if (QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION) && Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection))
		bRemoveSelection = TRUE;

	// Alte Fensterprozedur aufrufen
	long lRet = CallOldWndProc(WM_LBUTTONUP, wParam, lParam);

	BOOL bRowHeights = m_pCounter->GetRowHeights() > 0;
	BOOL bNoColHdr = QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	// Spalte evtl. bewegt?
	if (m_bColMoving)
	{
		m_bColMoving = FALSE;
		m_bMergeCellsInvalid = TRUE;
		if (m_pCounter->GetMergeCells() > 0)
			InvalidateRect(m_hWndTbl, NULL, FALSE);
		if (m_bCellMode)
			m_bNoFocusPaint = FALSE;
		bUpdateFocus = TRUE;
	}

	// Spaltengröße evtl. geändert?
	if (m_bColSizing)
	{
		m_bColSizing = FALSE;
		if (m_bCellMode)
			m_bNoFocusPaint = FALSE;
		bUpdateFocus = TRUE;
	}

	// Row-Header-Größe evtl. geändert?
	if (m_bRowHdrSizing)
	{
		m_bRowHdrSizing = FALSE;
		/*if ( m_bCellMode )
		m_bNoFocusPaint = FALSE;*/
		bUpdateFocus = TRUE;
	}

	// Zeilengröße evtl. geändert
	if (m_bRowSizing)
	{
		BOOL bVarRowHt = QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT);
		BOOL bNoColHdr = QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);
		if (bVarRowHt && m_lRowSizing != TBL_Error)
			AdaptScrollRangeV(m_lRowSizing < 0);
		else if (bNoColHdr && m_iRowSizingLinesPerRow > 0)
		{
			DrawLinesPerRowIndicators(m_iRowSizingLinesPerRow);

			int iLinesPerRowCurrent;
			Ctd.TblQueryLinesPerRow(m_hWndTbl, &iLinesPerRowCurrent);
			if (m_iRowSizingLinesPerRow != iLinesPerRowCurrent)
				Ctd.TblSetLinesPerRow(m_hWndTbl, m_iRowSizingLinesPerRow);
		}

		m_bRowSizing = FALSE;
		m_lRowSizing = TBL_Error;
		m_lRowSizingOffsY = 0;
		m_iRowSizingLinesPerRow = 0;

		bUpdateFocus = TRUE;
		FixSplitWindow();
	}

	// Split-Bar bewegt?
	if (m_bSplitBarDragging)
		m_bSplitBarDragging = FALSE;

	// Selektion entfernen?
	if (bRemoveSelection)
	{
		if (Ctd.TblAnyRows(m_hWndTbl, ROW_Selected, ROW_Hidden))
			Ctd.TblClearSelection(m_hWndTbl);
	}

	// Zeilenselektion per Maus?
	if (m_iRowSelByMouse)
	{
		if (bRowHeights || bNoColHdr)
		{
			m_pCounter->IncNoFocusUpdate();
			BOOL bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, m_RowSelRange.lRowTo);
			m_pCounter->DecNoFocusUpdate();
		}

		// Focus entfernen?
		if (MustSuppressRowFocusCTD())
		{
			//Ctd.TblKillFocus( m_hWndTbl );
			KillFocusFrameI();
			InvalidateRowFocus(Ctd.TblGetFocusRow(m_hWndTbl), FALSE);
			UpdateWindow(m_hWndTbl);
		}

		bUpdateFocus = m_bCellMode || bRowHeights || bNoColHdr;
	}

	// Zeilenselektion per Maus beendet
	m_bRowSelByMouseClick = FALSE;
	m_iRowSelByMouse = 0;
	m_RowSelRange.SetEmpty();
	if (m_bTempSingleSelection)
	{
		Ctd.TblSetTableFlags(m_hWndTbl, TBL_Flag_SingleSelection, FALSE);
		m_bTempSingleSelection = FALSE;
	}

	// Zellenselektion per Maus?
	if (m_bCellSelByMouse)
	{
		// Zeilen/Zellenfokus setzen
		m_pCounter->IncNoFocusUpdate();
		BOOL bSetFocusRet = Ctd.TblSetFocusRow(m_hWndTbl, m_CellSelRange.lRowTo);
		m_pCounter->DecNoFocusUpdate();

		HWND hWndUpdateFocusCol = bSetFocusRet ? m_CellSelRange.hWndColTo : NULL;
		long lUpdateFocusRow = bSetFocusRet ? m_CellSelRange.lRowTo : TBL_Error;

		UpdateFocus(lUpdateFocusRow, hWndUpdateFocusCol, UF_REDRAW_AUTO, UF_SELCHANGE_NONE);
	}

	// Zellenselektion per Maus beendet
	m_bCellSelByMouseClick = m_bCellSelByMouse = FALSE;

	// Button gedrückt?
	if (m_hWndColBtnPushed && m_lRowBtnPushed != TBL_Error)
	{
		BOOL bClicked = FALSE;
		DWORD dwFlags = 0, dwMFlags = 0;
		HWND hWndCol = NULL;
		long lRow = TBL_Error;
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };

		if (ObjFromPoint(pt, &lRow, &hWndCol, &dwFlags, &dwMFlags))
		{
			HWND hWndMergeCol;
			long lMergeRow;
			if ((dwMFlags & MTOFP_OVERMERGEDCELL) && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
			{
				hWndCol = hWndMergeCol;
				lRow = lMergeRow;
			}

			bClicked = (dwMFlags & MTOFP_OVERBTN) && hWndCol == m_hWndColBtnPushed && lRow == m_lRowBtnPushed;
		}

		CMTblRow *pRow = m_Rows->GetRowPtr(m_lRowBtnPushed);
		if (pRow)
		{
			pRow->m_Cols->SetBtnPushed(m_hWndColBtnPushed, FALSE);
			InvalidateBtn(m_hWndColBtnPushed, m_lRowBtnPushed, NULL, FALSE);
			if (bClicked)
				PostMessage(m_hWndTbl, MTM_BtnClick, WPARAM(m_hWndColBtnPushed), m_lRowBtnPushed);

			m_hWndColBtnPushed = NULL;
			m_lRowBtnPushed = TBL_Error;
		}


	}
	else if (m_hWndColBtnPushed && m_lRowBtnPushed == TBL_Error)
	{
		BOOL bClicked = FALSE;
		DWORD dwFlags = 0, dwMFlags = 0;
		HWND hWndCol = NULL;
		long lRow = TBL_Error;
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };

		if (ObjFromPoint(pt, &lRow, &hWndCol, &dwFlags, &dwMFlags))
		{
			bClicked = (dwMFlags & MTOFP_OVERBTN) && hWndCol == m_hWndColBtnPushed;
		}

		m_Cols->SetHdrBtnPushed(m_hWndColBtnPushed, FALSE);
		InvalidateBtn(m_hWndColBtnPushed, TBL_Error, NULL, FALSE);
		if (bClicked)
			PostMessage(m_hWndTbl, MTM_BtnClick, WPARAM(m_hWndColBtnPushed), TBL_Error);

		m_hWndColBtnPushed = NULL;
	}

	// Capture beenden
	SetCapture(NULL);

	// Fokus-Zelle aktualisieren?
	if (bUpdateFocus)
		UpdateFocus();

	// Erweiterte Nachrichten
	if (m_bExtMsgs)
	{
		// XY-Position ermitteln
		iLUpX = LOWORD(lParam);
		iLUpY = HIWORD(lParam);
		// Nachricht generieren
		GenerateExtMsgs(WM_LBUTTONUP);
	}

	return lRet;
}


// OnListClipParentNotify
LRESULT MTBLSUBCLASS::OnListClipParentNotify(HWND hWndTbl, HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	int iEvent = int(LOWORD(wParam));

	if (iEvent == WM_CREATE)
	{
		TCHAR szBuf[255];
		HWND hWndChild = HWND(lParam);

		if (GetClassName(hWndChild, szBuf, (sizeof(szBuf) / sizeof(TCHAR)) - 1))
		{
			BOOL bEdit = !lstrcmp(szBuf, _T("Edit"));

			if (bEdit)
			{
				// Edit-Stil-Hook entfernen
				if (m_hHookEditStyle && m_pscHookEditStyle == (void*)this)
				{
					UnhookWindowsHookEx(m_hHookEditStyle);
					m_hHookEditStyle = NULL;
					m_pscHookEditStyle = NULL;
					m_iHookEditJustify = -1;
					m_bHookEditMultiline = FALSE;
				}

				long lStyle = GetWindowLongPtr(hWndChild, GWL_STYLE);
				lStyle = lStyle | WS_CLIPSIBLINGS | WS_CLIPCHILDREN;
				SetWindowLongPtr(hWndChild, GWL_STYLE, lStyle);
				m_hWndEdit = hWndChild;
				PostMessage(m_hWndTbl, MTM_Internal, IA_SUBCLASS_LISTCLIP_CTRL, (LPARAM)hWndChild);
			}
			else
			{
				if (!lstrcmp(szBuf, CLSNAME_CONTAINER) || !lstrcmp(szBuf, CLSNAME_CONTAINER_TBW))
				{
					SubClassContainer(hWndChild);
					HWND hWndCB = FindFirstChildClass(hWndChild, _T("Button"));
					if (hWndCB)
						OnContainerParentNotify(m_hWndTbl, hWndChild, MAKELONG(WM_CREATE, 0), (LPARAM)hWndCB);
				}
				else if (!lstrcmp(szBuf, CLSNAME_DROPDOWNBTN))
				{
					//PostMessage( m_hWndTbl, MTM_Internal, IA_SUBCLASS_LISTCLIP_CTRL, (LPARAM)hWndChild );
					SubClassListClipCtrl(hWndChild);
				}
			}
		}
	}

	return 0;
}


// OnMouseMove
LRESULT MTBLSUBCLASS::OnMouseMove(WPARAM wParam, LPARAM lParam)
{
	POINT pt;

	GetCursorPos(&pt);
	ScreenToClient(m_hWndTbl, &pt);

	if (m_bDragCanAutoStart)
		DragDropAutoStartCancel();

	BOOL bVarRowHt = QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT);
	BOOL bRowHeights = m_pCounter->GetRowHeights() > 0;
	BOOL bNoColHdr = QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);
	BOOL bSuppressRowSelection = QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION), bTrapRowSelChanges;

	if (m_bRowSelByMouseClick && !m_iRowSelByMouse)
	{
		if (pt.y > m_ptRowSelByMouseClick.y)
			m_iRowSelByMouse = 1; // Top-Down-Selektion
		else if (pt.y < m_ptRowSelByMouseClick.y)
			m_iRowSelByMouse = 2; // Bottom-Up-Selektion
	}

	if (m_bCellSelByMouseClick)
		m_bCellSelByMouse = TRUE;

	if (m_iRowSelByMouse && bSuppressRowSelection && !Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection))
	{
		Ctd.TblSetTableFlags(m_hWndTbl, TBL_Flag_SingleSelection, TRUE);
		m_bTempSingleSelection = TRUE;
	}

	HWND hWndFocusCol;
	long lFocusRowAfter = TBL_Error, lFocusRowBefore = TBL_Error;
	if (m_iRowSelByMouse)
	{
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRowBefore, &hWndFocusCol);
	}

	bTrapRowSelChanges = m_iRowSelByMouse && MustTrapRowSelChanges() && !(bRowHeights || bNoColHdr);
	if (bTrapRowSelChanges)
		BeginRowSelTrap();

	long lRet = CallOldWndProc(WM_MOUSEMOVE, wParam, lParam);

	if ((bRowHeights || bNoColHdr) && (m_iRowSelByMouse || m_bCellSelByMouse))
	{
		CMTblMetrics tm(m_hWndTbl);

		long lRowsTop = m_RowSelRange.lRowFrom >= 0 ? tm.m_RowsTop : tm.m_SplitRowsTop;
		long lRowsBottom = m_RowSelRange.lRowFrom >= 0 ? tm.m_RowsBottom : tm.m_SplitRowsBottom;

		// Vertikale Scrollbar ermitteln
		HWND hWndScr = NULL;
		int iScrCtl;
		GetScrollBarV(&tm, m_RowSelRange.lRowFrom < 0, &hWndScr, &iScrCtl);

		// Ggf. Scrollen
		if (hWndScr)
		{
			BOOL bLBtnDown = HIBYTE(GetAsyncKeyState(GetSystemMetrics(SM_SWAPBUTTON) ? VK_RBUTTON : VK_LBUTTON));

			int iMaxRange, iMinRange;
			FlatSB_GetScrollRange(hWndScr, iScrCtl, &iMinRange, &iMaxRange);

			int iScrPos = FlatSB_GetScrollPos(hWndScr, iScrCtl);

			if (pt.y < lRowsTop && bLBtnDown)
			{
				if (iScrPos > iMinRange)
				{
					OnVScroll(SB_LINEUP, LPARAM(iScrCtl == SB_CTL ? hWndScr : 0));
					UpdateWindow(m_hWndTbl);

					if (FlatSB_GetScrollPos(hWndScr, iScrCtl) > iMinRange)
						PostMessage(m_hWndTbl, WM_MOUSEMOVE, wParam, lParam);
				}
			}
			else if (pt.y > lRowsBottom && bLBtnDown)
			{
				if (iScrPos < iMaxRange)
				{
					OnVScroll(SB_LINEDOWN, LPARAM(iScrCtl == SB_CTL ? hWndScr : 0));
					UpdateWindow(m_hWndTbl);

					FlatSB_GetScrollRange(hWndScr, iScrCtl, &iMinRange, &iMaxRange);
					if (FlatSB_GetScrollPos(hWndScr, iScrCtl) < iMaxRange)
						PostMessage(m_hWndTbl, WM_MOUSEMOVE, wParam, lParam);
				}
			}
		}

		// Ggf. Zeilenselektion/Focus anpassen
		if (m_iRowSelByMouse)
		{
			long lRowTo = TBL_Error;

			POINT ptObj = pt;
			ptObj.x = 0;
			if (pt.y < lRowsTop)
				ptObj.y = lRowsTop;
			else if (pt.y > lRowsBottom)
				ptObj.y = lRowsBottom - 1;

			DWORD dwFlags;
			HWND hWndCol;
			Ctd.TblObjectsFromPoint(m_hWndTbl, ptObj.x, ptObj.y, &lRowTo, &hWndCol, &dwFlags);

			long lMergeRowTo = lRowTo != TBL_Error ? GetMergeRow(lRowTo) : TBL_Error;
			lRowTo = lMergeRowTo != TBL_Error ? lMergeRowTo : lRowTo;

			if (lRowTo != TBL_Error && lRowTo != m_RowSelRange.lRowTo)
			{
				long lOldRowTo = m_RowSelRange.lRowTo;
				m_RowSelRange.lRowTo = lRowTo;

				if (!Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection))
				{
					long lRowFrom = m_RowSelRange.lRowFrom;

					BOOL bSelect = Ctd.TblQueryRowFlags(m_hWndTbl, m_RowSelRange.lRowFrom, ROW_Selected);

					if (lOldRowTo != TBL_Error)
					{
						if (m_iRowSelByMouse == 1 && lRowTo < m_RowSelRange.lRowFrom)
							m_iRowSelByMouse = 2;
						else if (m_iRowSelByMouse == 2 && lRowTo > m_RowSelRange.lRowFrom)
							m_iRowSelByMouse = 1;

						if (m_iRowSelByMouse == 1 && lRowTo < lOldRowTo)
						{
							if (bSelect)
								bSelect = FALSE;

							lRowFrom = lRowTo + 1;
							if (lMergeRowTo != TBL_Error)
							{
								while (GetMergeRow(lRowFrom) == lMergeRowTo)
								{
									lRowFrom++;
								}
							}
							lRowTo = lOldRowTo;
						}
						else if (m_iRowSelByMouse == 2 && lRowTo > lOldRowTo)
						{
							if (bSelect)
								bSelect = FALSE;

							lRowFrom = lOldRowTo;
							lRowTo = lRowTo - 1;

							long lMergedOldRowTo = GetMergeRow(lOldRowTo);
							if (lMergedOldRowTo != TBL_Error)
							{
								while (GetMergeRow(lRowTo) == lMergedOldRowTo)
								{
									lRowTo--;
								}
							}
						}
					}

					WORD wFlagsOn = bSelect ? 0 : ROW_Selected;
					WORD wFlagsOff = bSelect ? ROW_Selected | ROW_Hidden : ROW_Hidden;
					Ctd.TblSetFlagsRowRange(m_hWndTbl, lRowFrom, lRowTo, wFlagsOn, wFlagsOff, ROW_Selected, bSelect);
				}
				else
				{
					Ctd.TblSetFocusRow(m_hWndTbl, lRowTo);
					Ctd.TblClearSelection(m_hWndTbl);
					Ctd.TblSetRowFlags(m_hWndTbl, lRowTo, ROW_Selected, TRUE);
				}
			}
		}
	}


	if (m_iRowSelByMouse)
	{
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRowAfter, &hWndFocusCol);
		if (lFocusRowAfter != lFocusRowBefore)
		{
			BOOL bUpdate = FALSE;
			if (InvalidateRowFocus(lFocusRowBefore, FALSE))
				bUpdate = TRUE;
			if (InvalidateRowFocus(lFocusRowAfter, FALSE))
				bUpdate = TRUE;
			if (bUpdate)
			{
				UpdateWindow(m_hWndTbl);
			}
		}
	}

	if (bTrapRowSelChanges)
		EndRowSelTrap();

	if (m_bCellSelByMouse)
	{
		DWORD dwFlags = 0;
		HWND hWndCol = NULL;
		long lRow = TBL_Error;
		Ctd.TblObjectsFromPoint(m_hWndTbl, pt.x, pt.y, &lRow, &hWndCol, &dwFlags);
		if (hWndCol != m_CellSelRange.hWndColTo || lRow != m_CellSelRange.lRowTo)
		{
			CMTblMetrics tm(m_hWndTbl);
			HWND hWndColScrollTo = NULL;
			long lRowScrollTo = TBL_Error;

			long lRowsTop = m_CellSelRange.lRowFrom >= 0 ? tm.m_RowsTop : tm.m_SplitRowsTop;
			long lRowsBottom = m_CellSelRange.lRowFrom >= 0 ? tm.m_RowsBottom : tm.m_SplitRowsBottom;

			if (pt.y < lRowsTop)
			{
				DWORD dwFlags2 = 0;
				HWND hWndCol2 = NULL;
				long lRow2 = TBL_Error;
				Ctd.TblObjectsFromPoint(m_hWndTbl, 0, lRowsTop, &lRow2, &hWndCol2, &dwFlags2);

				if (lRow2 != TBL_Error)
				{
					if (Ctd.TblFindPrevRow(m_hWndTbl, &lRow2, 0, ROW_Hidden) && AreRowsOfSameArea(lRow2, m_CellSelRange.lRowFrom))
					{
						lRowScrollTo = lRow2;
						if (lRow2 < m_CellSelRange.lRowTo)
							lRow = lRow2;
					}
				}
			}
			else if (pt.y > lRowsBottom)
			{
				DWORD dwFlags2 = 0;
				HWND hWndCol2 = NULL;
				long lRow2 = TBL_Error;
				Ctd.TblObjectsFromPoint(m_hWndTbl, 0, lRowsBottom - 1, &lRow2, &hWndCol2, &dwFlags2);

				if (lRow2 != TBL_Error)
				{
					long lNextRow = lRow2;
					if (Ctd.TblFindNextRow(m_hWndTbl, &lNextRow, 0, ROW_Hidden) && AreRowsOfSameArea(lNextRow, m_CellSelRange.lRowFrom))
						lRow2 = lNextRow;

					lRowScrollTo = lRow2;
					if (lRow2 > m_CellSelRange.lRowTo)
						lRow = lRow2;
				}
			}
			else if (pt.x < tm.m_ClientRect.left || pt.x > tm.m_ClientRect.right)
			{
				DWORD dwFlags2 = 0;
				HWND hWndCol2 = NULL;
				long lRow2 = TBL_Error;
				Ctd.TblObjectsFromPoint(m_hWndTbl, 0, pt.y, &lRow2, &hWndCol2, &dwFlags2);

				if (lRow2 != TBL_Error && AreRowsOfSameArea(lRow2, m_CellSelRange.lRowFrom))
				{
					if (lRow2 != m_CellSelRange.lRowTo)
						lRow = lRow2;
				}
			}

			if (lRow != TBL_Error)
			{
				if (!AreRowsOfSameArea(lRow, m_CellSelRange.lRowFrom))
					lRow = TBL_Error;
			}

			if (!hWndCol)
			{
				DWORD dwFlags2 = 0;
				HWND hWndCol2 = NULL;
				long lRow2 = TBL_Error;
				Ctd.TblObjectsFromPoint(m_hWndTbl, pt.x, 0, &lRow2, &hWndCol2, &dwFlags2);

				if (hWndCol2)
				{
					if (hWndCol2 != m_CellSelRange.hWndColTo)
						hWndCol = hWndCol2;
				}
				else
				{
					POINT pt2, ptVis2;
					//int iRet = MTblGetColXCoord( m_hWndTbl, m_CellSelRange.hWndColTo, &tm, &pt2, &ptVis2 );
					int iRet = GetColXCoord(m_CellSelRange.hWndColTo, &tm, &pt2, &ptVis2);
					if (iRet > MTGR_RET_INVISIBLE)
					{
						if (pt.x < ptVis2.x)
						{
							HWND hWndCol2 = m_CellSelRange.hWndColTo;
							//MTblFindPrevCol( m_hWndTbl, &hWndCol2, COL_Visible, 0 );
							FindPrevCol(&hWndCol2, COL_Visible, 0);
							hWndColScrollTo = hWndCol2;
							hWndCol = hWndCol2;
						}
						else if (pt.x > ptVis2.x)
						{
							HWND hWndCol2 = m_CellSelRange.hWndColTo;
							//MTblFindNextCol( m_hWndTbl, &hWndCol2, COL_Visible, 0 );
							FindNextCol(&hWndCol2, COL_Visible, 0);
							hWndColScrollTo = hWndCol2;
							hWndCol = hWndCol2;
						}
					}
				}
			}

			if ((lRow != TBL_Error || hWndCol != NULL) && (hWndCol != m_CellSelRange.hWndColTo || lRow != m_CellSelRange.lRowTo))
			{
				MTBLCELLRANGE crOld, crNew;
				m_CellSelRangeA.CopyTo(crOld);

				m_CellSelRange.CopyTo(crNew);
				if (hWndCol)
					crNew.hWndColTo = hWndCol;
				if (lRow != TBL_Error)
					crNew.lRowTo = lRow;

				AdaptCellRange(crNew);

				CellSelRangeChanged(crOld, crNew, TRUE);

				crNew.CopyTo(m_CellSelRangeA);

				if (hWndCol)
					m_CellSelRange.hWndColTo = hWndCol;
				if (lRow != TBL_Error)
					m_CellSelRange.lRowTo = lRow;

			}

			if (lRowScrollTo != TBL_Error || hWndColScrollTo != NULL)
				Ctd.TblScroll(m_hWndTbl, lRowScrollTo, hWndColScrollTo, TBL_AutoScroll);
		}
	}

	// Button gedrückt?
	if (m_hWndColBtnPushed && m_lRowBtnPushed != TBL_Error)
	{
		BOOL bPushed = FALSE;
		DWORD dwFlags = 0, dwMFlags = 0;
		HWND hWndCol = NULL;
		long lRow = TBL_Error;
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };

		if (ObjFromPoint(pt, &lRow, &hWndCol, &dwFlags, &dwMFlags))
		{
			HWND hWndMergeCol;
			long lMergeRow;
			if ((dwMFlags & MTOFP_OVERMERGEDCELL) && GetMergeCellEx(hWndCol, lRow, hWndMergeCol, lMergeRow))
			{
				hWndCol = hWndMergeCol;
				lRow = lMergeRow;
			}

			bPushed = (dwMFlags & MTOFP_OVERBTN) && hWndCol == m_hWndColBtnPushed && lRow == m_lRowBtnPushed;
		}

		CMTblRow *pRow = m_Rows->GetRowPtr(m_lRowBtnPushed);
		if (pRow && pRow->m_Cols->GetBtnPushed(m_hWndColBtnPushed) != bPushed)
		{
			pRow->m_Cols->SetBtnPushed(m_hWndColBtnPushed, bPushed);
			InvalidateBtn(m_hWndColBtnPushed, m_lRowBtnPushed, NULL, FALSE);
		}
	}
	else if (m_hWndColBtnPushed && m_lRowBtnPushed == TBL_Error)
	{
		BOOL bPushed = FALSE;
		DWORD dwFlags = 0, dwMFlags = 0;
		HWND hWndCol = NULL;
		long lRow = TBL_Error;
		POINT pt = { LOWORD(lParam), HIWORD(lParam) };

		if (ObjFromPoint(pt, &lRow, &hWndCol, &dwFlags, &dwMFlags))
		{
			bPushed = (dwMFlags & MTOFP_OVERBTN) && hWndCol == m_hWndColBtnPushed;
		}

		if (m_Cols->GetHdrBtnPushed(m_hWndColBtnPushed) != bPushed)
		{
			m_Cols->SetHdrBtnPushed(m_hWndColBtnPushed, bPushed);
			InvalidateBtn(m_hWndColBtnPushed, TBL_Error, NULL, FALSE);
		}
	}

	// Wird gerade die Zeilenhöhe geändert?
	if (bVarRowHt && m_lRowSizing != TBL_Error)
	{
		CMTblMetrics tm(m_hWndTbl);

		long lRowsTop = m_lRowSizing >= 0 ? tm.m_RowsTop : tm.m_SplitRowsTop;
		long lRowsBottom = m_lRowSizing >= 0 ? tm.m_RowsBottom : tm.m_SplitRowsBottom;

		POINT ptRow, ptRowVis;
		int iRet = GetRowYCoord(m_lRowSizing, &tm, &ptRow, &ptRowVis, FALSE);
		if (iRet != MTGR_RET_ERROR)
		{
			typedef struct T_ROWHEIGHT
			{
				long lRow;
				long lHt;
				long lHtNew;
			} ROWHEIGHT;

			long lY = pt.y - m_lRowSizingOffsY, lRowX = ptRow.x, lRowY = ptRow.y;

			long lRows = 1;
			ROWHEIGHT *pRows = NULL;
			long lHtOld = 0;

			LPMTBLMERGEROW pmr = FindMergeRow(m_lRowSizing, FMR_ANY);
			if (pmr)
			{
				lRows = pmr->iMergedRows + 1;
				pRows = new ROWHEIGHT[lRows];

				long l;
				long lHt = GetEffRowHeight(m_lRowSizing, &tm);
				pRows[0].lRow = m_lRowSizing;
				pRows[0].lHt = pRows[0].lHtNew = lHt;
				l = 1;

				long lRow = m_lRowSizing;
				lHtOld = lHt;
				while (FindPrevRow(&lRow, 0, ROW_Hidden, pmr->lRowFrom))
				{
					lHt = GetEffRowHeight(lRow, &tm);

					pRows[l].lRow = lRow;
					pRows[l].lHt = pRows[l].lHtNew = lHt;
					++l;

					lRowX -= lHt;
					lHtOld += lHt;
				}
			}


			BOOL bSized = FALSE;
			long lHt = max((lY - lRowX + 1), lRows);

			if (pmr)
			{
				long lDiff = lHt - lHtOld;
				if (abs(lDiff) >= lRows)
				{
					// Sortieren
					int i, j;
					for (i = 0; i < lRows - 1; ++i)
					{
						for (j = 0; j < lRows - i - 1; ++j)
						{
							if (lDiff > 0 ? pRows[j].lHt > pRows[j + 1].lHt : pRows[j].lHt < pRows[j + 1].lHt)
							{
								ROWHEIGHT tmp = pRows[j];
								pRows[j] = pRows[j + 1];
								pRows[j + 1] = tmp;
							}
						}
					}

					// Differenz verteilen
					i = 0;
					long lAdd = (lDiff > 0 ? 1 : -1);
					long lHtNew;
					long lMinHtCount = 0;
					while (lDiff)
					{
						lHtNew = pRows[i].lHtNew + lAdd;
						if (lHtNew > 0)
						{
							pRows[i].lHtNew = lHtNew;
							lDiff -= lAdd;
						}
						else
							++lMinHtCount;

						++i;
						if (i == lRows)
						{
							i = 0;
							if (lMinHtCount == lRows)
								break;
							lMinHtCount = 0;
						}
					}

					// Zeilenhöhen setzen
					for (i = 0; i < lRows; ++i)
					{
						if (pRows[i].lHtNew != pRows[i].lHt)
						{
							m_Rows->SetRowHeight(pRows[i].lRow, pRows[i].lHtNew == tm.m_RowHeight ? 0 : pRows[i].lHtNew);
							bSized = TRUE;
						}
					}

				}

				delete[]pRows;
			}
			else
			{
				if (GetEffRowHeight(m_lRowSizing, &tm) != lHt)
				{
					m_Rows->SetRowHeight(m_lRowSizing, lHt == tm.m_RowHeight ? 0 : lHt);
					bSized = TRUE;
				}
			}

			if (bSized)
			{
				UpdateFocus();

				RECT r = tm.m_ClientRect;
				r.top = max(lRowsTop, lRowX);
				r.bottom = lRowsBottom;
				InvalidateRect(m_hWndTbl, &r, FALSE);
				UpdateWindow(m_hWndTbl);
			}
		}
	}
	else if (bNoColHdr && m_lRowSizing != TBL_Error && m_iRowSizingLinesPerRow > 0)
	{
		CMTblMetrics tm(m_hWndTbl);

		long lRowsTop = m_lRowSizing >= 0 ? tm.m_RowsTop : tm.m_SplitRowsTop;
		long lRowsBottom = m_lRowSizing >= 0 ? tm.m_RowsBottom : tm.m_SplitRowsBottom;

		POINT ptRow, ptRowVis;
		int iRet = GetRowYCoord(m_lRowSizing, &tm, &ptRow, &ptRowVis, FALSE);
		if (iRet != MTGR_RET_ERROR)
		{
			long lY = pt.y - m_lRowSizingOffsY, lRowX = ptRow.x, lRowY = ptRow.y;
			long lHt = max((lY - lRowX), 0);

			int iLinesPerRowNew = 1;
			while (lHt >= tm.CalcRowHeight(iLinesPerRowNew + 1))
			{
				++iLinesPerRowNew;
			}

			if (iLinesPerRowNew != m_iRowSizingLinesPerRow)
			{
				DrawLinesPerRowIndicators(m_iRowSizingLinesPerRow);
				DrawLinesPerRowIndicators(iLinesPerRowNew);
				m_iRowSizingLinesPerRow = iLinesPerRowNew;
			}
		}
	}

	// Ggf. internen Timer starten
	if (!m_uiTimer)
		StartInternalTimer(TRUE);

	return lRet;
}

// OnMouseWheel
LRESULT MTBLSUBCLASS::OnMouseWheel(WPARAM wParam, LPARAM lParam)
{
	if (m_bMWheelScroll)
	{
		BOOL bScrollSplit;
		TCHAR szClsName[1024] = _T("");
		DWORD dwFlags;
		HWND hWnd, hWndCol, hWndFocusCol = NULL;
		int iDelta, iRowsPerNotch;
		long /*l,*/ lFocusRow = TBL_Error, lKeyCode, lMin, lMax, lRow, lRows, lRowsFound;
		POINT p;

		p.x = (short)LOWORD(lParam);
		p.y = (short)HIWORD(lParam);

		iDelta = (short)HIWORD(wParam);

		hWnd = WindowFromPoint(p);
		if (hWnd && hWnd != m_hWndTbl)
		{
			if (GetClassName(hWnd, szClsName, sizeof(szClsName)) && lstrcmp(szClsName, _T("ListBox")) == 0)
			{
				return SendMessage(hWnd, WM_MOUSEWHEEL, wParam, lParam);
			}
		}

		ScreenToClient(m_hWndTbl, &p);
		if (iDelta && Ctd.TblObjectsFromPoint(m_hWndTbl, p.x, p.y, &lRow, &hWndCol, &dwFlags))
		{
			if (Ctd.TblQueryFocus(m_hWndTbl, &lFocusRow, &hWndFocusCol) && hWndFocusCol && hWndFocusCol == hWndCol && lFocusRow == lRow)
			{
				if (MTblIsListOpen(m_hWndTbl, hWndCol))
				{
					if (iDelta < 0)
						lKeyCode = VK_DOWN;
					else
						lKeyCode = VK_UP;
					if (IsWindow(m_hWndEdit))
						SendMessage(m_hWndEdit, WM_KEYDOWN, lKeyCode, 0);
					return 0;
				}
			}

			bScrollSplit = dwFlags & TBL_YOverSplitRows;
			if (bScrollSplit)
			{
				//if ( MTblQueryVisibleSplitRange( m_hWndTbl, &lMin, &lMax ) )
				if (QueryVisRange(lMin, lMax, TRUE))
				{
					lRow = (iDelta < 0) ? lMax : lMin;
					if (m_iMWPageKey && GetKeyState(m_iMWPageKey) < 0)
					{
						CMTblMetrics tm(m_hWndTbl);
						iRowsPerNotch = (tm.m_SplitRowsBottom - tm.m_SplitRowsTop) / tm.m_RowHeight;
					}
					else
					{
						iRowsPerNotch = m_iMWSplitRowsPerNotch;
					}

					lRows = (abs(iDelta) / WHEEL_DELTA) * iRowsPerNotch;

					for (lRowsFound = 0; lRowsFound < lRows; lRowsFound++)
					{
						if (!((iDelta < 0) ? FindNextRow(&lRow, 0, ROW_Hidden, -1) : FindPrevRow(&lRow, 0, ROW_Hidden)))
							break;
					}
					if (lRowsFound)
					{
						(iDelta < 0) ? Ctd.TblScroll(m_hWndTbl, lRow, NULL, TBL_ScrollBottom) : Ctd.TblScroll(m_hWndTbl, lRow, NULL, TBL_ScrollTop);
					}
				}
			}
			else
			{
				if (Ctd.TblAnyRows(m_hWndTbl, 0, ROW_Hidden))
				{
					if (Ctd.TblQueryVisibleRange(m_hWndTbl, &lMin, &lMax))
					{
						lRow = (iDelta < 0) ? lMax : lMin;

						if (m_iMWPageKey && GetKeyState(m_iMWPageKey) < 0)
						{
							CMTblMetrics tm(m_hWndTbl);
							iRowsPerNotch = (tm.m_RowsBottom - tm.m_RowsTop) / tm.m_RowHeight;
						}
						else
						{
							iRowsPerNotch = m_iMWRowsPerNotch;
						}

						lRows = (abs(iDelta) / WHEEL_DELTA) * iRowsPerNotch;

						for (lRowsFound = 0; lRowsFound < lRows; lRowsFound++)
						{
							if (!((iDelta < 0) ? FindNextRow(&lRow, 0, ROW_Hidden) : FindPrevRow(&lRow, 0, ROW_Hidden, 0)))
								break;
						}
						if (lRowsFound)
						{
							(iDelta < 0) ? Ctd.TblScroll(m_hWndTbl, lRow, NULL, TBL_ScrollBottom) : Ctd.TblScroll(m_hWndTbl, lRow, NULL, TBL_ScrollTop);
						}
					}
				}

				/*long lMsgs = abs( iDelta ) / WHEEL_DELTA;
				WORD wScrollCode;
				if ( m_iMWPageKey && GetKeyState( m_iMWPageKey ) < 0 )
				{
				wScrollCode = ( iDelta < 0 ) ? SB_PAGEDOWN : SB_PAGEUP;
				}
				else
				{
				wScrollCode = ( iDelta < 0 ) ? SB_LINEDOWN : SB_LINEUP;
				lMsgs *= m_iMWRowsPerNotch;
				}
				WPARAM wp = MAKELONG( wScrollCode, 0 );
				for ( l = 0; l < lMsgs; l++ )
				{
				SendMessage( m_hWndTbl, WM_VSCROLL, wp, 0 );
				}*/
			}
		}

		return 0;
	}
	else
	{
		// Alte Fensterprozedur aufrufen
		BeforeScroll();
		long lRet = CallOldWndProc(WM_MOUSEWHEEL, wParam, lParam);
		AfterScroll();
		return lRet;
	}
}


// OnNCDestroy
LRESULT MTBLSUBCLASS::OnNCDestroy(WPARAM wParam, LPARAM lParam)
{
	int iRet;
	// Tip-Fenster zerstören
	if (IsWindow(m_hWndTip))
		DestroyWindow(m_hWndTip);
	// Timer stoppen
	/*if ( m_uiTimer )
	KillTimer( m_hWndTbl, m_uiTimer );*/
	StopInternalTimer();
	// Liste der mit TBL_Adjust gelöschten Zeilen freigeben
	delete m_DelAdjRows;
	// Liste der letzten Sortierspalten löschen
	delete m_LastSortCols;
	// Liste der letzten Sortierflags löschen
	delete m_LastSortFlags;
	// Zeilen löschen
	delete m_Rows;
	// Spalten löschen
	delete m_Cols;
	// Aktuelle Spalten löschen
	delete m_CurCols;
	// Corner löschen
	delete m_Corner;
	// Spalten-Header-Gruppen löschen
	delete m_ColHdrGrps;
	// Default-Tip löschen
	delete m_DefTip;
	// EntireText-Tip löschen
	delete m_ETTip;
	// Subclass-Map löschen
	delete m_scm;
	// Export Ersatzfarben-Map löschen
	delete m_escm;
	// Highlighting-Definitions-Map löschen
	delete m_pHLDM;
	// Highlighted-Items-Liste löschen
	delete m_pHLIL;
	// PaintCols löschen
	for (int i = 0; i < MTBL_MAX_COLS; ++i)
	{
		delete m_ppd->PaintCols[i];
	}
	// PaintMergeCells löschen
	FreePaintMergeCells();
	// MergeCells löschen
	FreeMergeCells();
	// PaintMergeRowHdrs löschen
	FreePaintMergeRowHdrs();
	// MergeRows löschen
	FreeMergeRows();
	// Paint-Daten löschen
	delete m_ppd;
	// Row-Flag-Images löschen
	DeleteRowFlagImages();
	// Counter löschen
	delete m_pCounter;
	// Tip-Region löschen
	if (m_hRgnTip)
		DeleteObject(m_hRgnTip);
	// Zeilen, für die Kindzeilen geladen werden, löschen
	delete m_RowsLoadingChilds;
	// Liniendefinitionen löschen
	delete m_pHdrLineDef;
	delete m_pColsVertLineDef;
	delete m_pLastLockedColVertLineDef;
	delete m_pRowsHorzLineDef;
	// GDI-Objekte freigeben
	if (m_hEditBrush)
		iRet = DeleteObject(m_hEditBrush);
	if (m_bDeleteEditFont && m_hEditFont)
		iRet = DeleteObject(m_hEditFont);
	// Alte Fensterprozedur setzen
	SetWindowLongPtr(m_hWndTbl, GWLP_WNDPROC, (LONG_PTR)m_wpOld);
	// Eintrag aus Subclass-Map entfernen
	gscm.erase(m_hWndTbl);

	return 0;
}


// OnObjectsFromPoint
LRESULT MTBLSUBCLASS::OnObjectsFromPoint(WPARAM wParam, LPARAM lParam)
{
	typedef struct
	{
		long lX;
		long lY;
		long lRow;
		HWND hWndCol;
		DWORD dwFlags;
	} OBJFROMPOINT, *POBJFROMPOINT;

	long lRet = CallOldWndProc(TIM_ObjectsFromPoint, wParam, lParam);

	if (lRet && (m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER)))
	{
		CMTblMetrics tm(m_hWndTbl);
		POBJFROMPOINT pofp = (POBJFROMPOINT)lParam;

		if (pofp->dwFlags & TBL_YOverColumnHeader && QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
		{
			pofp->dwFlags = (pofp->dwFlags & TBL_YOverColumnHeader) ^ pofp->dwFlags;
			pofp->dwFlags |= TBL_YOverNormalRows;
		}

		pofp->lRow = TBL_Error;

		BOOL bOverNormalRows = (pofp->dwFlags & TBL_YOverNormalRows) != 0;
		BOOL bOverSplitRows = (pofp->dwFlags & TBL_YOverSplitRows) != 0;
		BOOL bOverSplitBar = (pofp->dwFlags & TBL_YOverSplitBar) != 0;

		if ((bOverNormalRows || bOverSplitRows) && !bOverSplitBar)
		{
			int iScrPosV = GetScrollPosV(&tm, bOverSplitRows);
			long lRow = m_Rows->GetVisPosRow(iScrPosV, bOverSplitRows);
			long lRowTop = bOverSplitRows ? tm.m_SplitRowsTop : tm.m_RowsTop;
			long lRowsBottom = bOverSplitRows ? tm.m_SplitRowsBottom : tm.m_RowsBottom;

			// Zeilenarray ermitteln
			long lLastValidPos, lUpperBound;
			CMTblRow **RowArr = m_Rows->GetRowArray(lRow, &lUpperBound, &lLastValidPos);

			// Zeile ermitteln
			if (RowArr)
			{				
				CMTblRow *pRow;
				long lHt, lRowBottom;
				long lPos = bOverSplitRows ? m_Rows->GetSplitRowPos(lRow) : lRow;
				for (; lRowTop <= lRowsBottom; ++lPos)
				{
					pRow = lRow <= lLastValidPos ? RowArr[lPos] : NULL;
					if (pRow && !pRow->IsVisible())
						continue;

					if (!pRow)
						lHt = tm.m_RowHeight;
					else
					{
						lHt = pRow->GetHeight();
						if (lHt == 0)
							lHt = tm.m_RowHeight;
					}

					lRowBottom = lRowTop + (lHt - 1);
					if (pofp->lY >= lRowTop && pofp->lY <= lRowBottom)
					{
						lRow = bOverSplitRows ? m_Rows->GetSplitPosRow(lPos) : lPos;
						if (IsRow(lRow))
							pofp->lRow = lRow;
						break;
					}

					lRowTop = lRowBottom + 1;
				}
			}			
		}
	}

	return lRet;
}


// OnParentNotify
LRESULT MTBLSUBCLASS::OnParentNotify(WPARAM wParam, LPARAM lParam)
{
	TCHAR szBuf[255];
	HWND hWnd = HWND(lParam);
	int iEvent = int(LOWORD(wParam));
	if (GetClassName(hWnd, szBuf, (sizeof(szBuf) / sizeof(TCHAR)) - 1))
	{
		// Klasse ermitteln
		LPTSTR lpsClass;
		long lTblType = Ctd.GetType(m_hWndTbl);

		// List-Clip?
		if (lTblType == TYPE_TableWindow)
			lpsClass = CLSNAME_LISTCLIP_TBW;
		else
			lpsClass = CLSNAME_LISTCLIP;
		BOOL bListClip = !lstrcmp(szBuf, lpsClass);

		// Column?
		if (lTblType == TYPE_TableWindow)
			lpsClass = CLSNAME_COLUMN_TBW;
		else
			lpsClass = CLSNAME_COLUMN;
		BOOL bColumn = !lstrcmp(szBuf, lpsClass);

		if (iEvent == WM_CREATE)
		{
			if (bListClip)
			{
				SubClassListClip(hWnd);
				m_hWndListClip = hWnd;
				// Edit-Stil-Hook setzen
				if (!m_hHookEditStyle && m_pscHookEditStyle == (void*)this)
					m_hHookEditStyle = SetWindowsHookEx(WH_CALLWNDPROC, HookProcEditStyle, (HINSTANCE)ghInstance, GetCurrentThreadId());
			}

			if (bColumn)
			{
				m_bMergeCellsInvalid = TRUE;
				//SubClassCol( hWnd );
				AttachCol(hWnd);
			}
		}
		else if (iEvent == WM_DESTROY)
		{
			if (bListClip)
			{
				if (m_hRgnListClip)
				{
					SetWindowRgn(m_hWndListClip, NULL, FALSE);
					DeleteObject(m_hRgnListClip);
					m_hRgnListClip = NULL;
				}
				m_hWndListClip = NULL;
			}
		}
	}

	return 0;
}


// OnPaste
LRESULT MTBLSUBCLASS::OnPaste(WPARAM wParam, LPARAM lParam)
{
	return CallOldWndProc(WM_PASTE, wParam, lParam);
}


// OnPosChanged
LRESULT MTBLSUBCLASS::OnPosChanged(WPARAM wParam, LPARAM lParam)
{
	long lRet = CallOldWndProc(TIM_PosChanged, wParam, lParam);
	return lRet;
}



// OnQueryVisRange
LRESULT MTBLSUBCLASS::OnQueryVisRange(WPARAM wParam, LPARAM lParam)
{
	typedef struct TAG_QUERYVISRANGE
	{
		long lDummy;
		long lRangeMin;
		long lRangeMax;
	} QUERYVISRANGE, *PQUERYVISRANGE;

	long lRet = 0;

	if (m_pCounter->GetRowHeights() <= 0 && !QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
	{
		PQUERYVISRANGE pqvr = (PQUERYVISRANGE)lParam;
		lRet = CallOldWndProc(TIM_QueryVisRange, wParam, lParam);
	}
	else
	{
		PQUERYVISRANGE pqvr = (PQUERYVISRANGE)lParam;
		if (QueryVisRange(pqvr->lRangeMin, pqvr->lRangeMax))
			return 1;
		else
			return 0;

		CMTblMetrics tm;
		tm.Get(m_hWndTbl, TRUE);

		//BOOL bDelAdj, bHidden, bIsRow;
		int iScrPosV = GetScrollPosV(&tm, FALSE);
		long lRow = m_Rows->GetVisPosRow(iScrPosV, FALSE);
		long lRowTop = tm.m_RowsTop;
		long lRowsBottom = tm.m_RowsBottom;

		pqvr->lRangeMin = pqvr->lRangeMax = lRow;

		// Zeilenarray ermitteln
		long lLastValidPos, lUpperBound;
		CMTblRow **RowArr = m_Rows->GetRowArray(0, &lUpperBound, &lLastValidPos);

		CMTblRow *pRow;
		long lHt, lRowBottom;
		for (;; ++lRow)
		{
			pRow = lRow <= lLastValidPos ? RowArr[lRow] : NULL;
			if (pRow && !pRow->IsVisible())
				continue;

			if (!pRow)
				lHt = tm.m_RowHeight;
			else
			{
				lHt = pRow->GetHeight();
				if (lHt == 0)
					lHt = tm.m_RowHeight;
			}
			lRowBottom = lRowTop + (lHt - 1);

			if (lRowBottom > lRowsBottom)
			{
				lRet = 1;
				break;
			}
			else
				pqvr->lRangeMax = lRow;

			lRowTop = lRowBottom + 1;
		}
	}

	return lRet;
}


// OnRButtonDblClk
LRESULT MTBLSUBCLASS::OnRButtonDblClk(WPARAM wParam, LPARAM lParam)
{
	// Alte Fensterprozedur aufrufen
	long lRet = CallOldWndProc(WM_RBUTTONDBLCLK, wParam, lParam);
	if (m_bExtMsgs)
	{
		// XY-Position ermitteln
		iRDblClkX = LOWORD(lParam);
		iRDblClkY = HIWORD(lParam);
		// Nachricht generieren
		GenerateExtMsgs(WM_RBUTTONDBLCLK);
		// Letzte RDown-Nachricht merken
		uLastRDownMsg = WM_RBUTTONDBLCLK;
	}

	return lRet;
}


// OnRButtonDown
LRESULT MTBLSUBCLASS::OnRButtonDown(WPARAM wParam, LPARAM lParam)
{
	// Alte Fensterprozedur aufrufen
	long lRet = CallOldWndProc(WM_RBUTTONDOWN, wParam, lParam);
	if (m_bExtMsgs)
	{
		// XY-Position ermitteln
		iRDownX = LOWORD(lParam);
		iRDownY = HIWORD(lParam);
		// Nachricht generieren
		GenerateExtMsgs(WM_RBUTTONDOWN);
		// Letzte RDown-Nachricht merken
		uLastRDownMsg = WM_RBUTTONDOWN;
	}

	return lRet;
}


// OnRButtonUp
LRESULT MTBLSUBCLASS::OnRButtonUp(WPARAM wParam, LPARAM lParam)
{
	// Alte Fensterprozedur aufrufen
	long lRet = CallOldWndProc(WM_RBUTTONUP, wParam, lParam);

	// Erweiterte Nachrichten
	if (m_bExtMsgs)
	{
		// XY-Position ermitteln
		iRUpX = LOWORD(lParam);
		iRUpY = HIWORD(lParam);
		// Nachricht generieren
		GenerateExtMsgs(WM_RBUTTONUP);
	}

	return lRet;
}


// OnReset
LRESULT MTBLSUBCLASS::OnReset(WPARAM wParam, LPARAM lParam)
{
	// Zeilenhöhen zurücksetzen + Scrollbereiche anpassen, da sonst ggf. GPF in tablixx.dll
	m_Rows->ResetRowHeights(FALSE);
	m_Rows->ResetRowHeights(TRUE);
	m_pCounter->InitRowHeights();

	AdaptScrollRangeV(FALSE);
	AdaptScrollRangeV(TRUE);

	// Alte Fensterprozedur aufrufen
	m_pCounter->IncNoFocusUpdate();
	long lRet = CallOldWndProc(TIM_Reset, wParam, lParam);
	m_pCounter->DecNoFocusUpdate();

	if (lRet)
	{
		// Bereich Zellenselektion/Tastatur zurücksetzen
		m_CellSelRangeKey.SetEmpty();

		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		// Merge-Zeilen ungültig
		m_bMergeRowsInvalid = TRUE;

		// Tooltip zurücksetzen
		if ( m_pTip )
			ResetTip();

		// Liste der mit TBL_Adjust gelöschten Zeilen zurücksetzen
		m_DelAdjRows->clear();

		// Zeilen zurücksetzen
		m_Rows->Reset();

		// Enter/Leave-Zeilen auf NULL setzen
		m_pELRow = m_pELRow_Cell = m_pELRow_Btn = m_pELRow_DDBtn = NULL;

		// Daten Fokus zurücksetzen
		ResetFocusData();
		UpdateFocus();

		if (MustSuppressRowFocusCTD())
			KillFocusFrameI();

		// Neu zeichnen
		if (QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
			InvalidateRect(m_hWndTbl, NULL, FALSE);
		else
			InvalidateColHdr(FALSE);

		// Erweiterte Nachricht senden
		if (m_bExtMsgs && lRet)
			PostMessage(m_hWndTbl, MTM_Reset, wParam, lParam);
	}

	return lRet;
}


// OnRowSetContext
LRESULT MTBLSUBCLASS::OnRowSetContext(WPARAM wParam, LPARAM lParam)
{
	if (QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION))
		Ctd.TblClearSelection(m_hWndTbl);

	long lRet = CallOldWndProc(TIM_RowSetContext, wParam, lParam);

	UpdateFocus();

	return lRet;
}


// OnScroll
LRESULT MTBLSUBCLASS::OnScroll(WPARAM wParam, LPARAM lParam)
{
	typedef struct
	{
		long lRow;
		HWND hWndCol;
	} SCROLL, *PSCROLL;

	long lRet = 0;

	PSCROLL ps = (PSCROLL)lParam;

	BeforeScroll();

	if (m_pCounter->GetRowHeights() <= 0 && !QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
		lRet = CallOldWndProc(TIM_Scroll, wParam, lParam);
	else
	{
		// Spalte scrollen lassen
		if (ps->hWndCol)
		{
			long lTmp = ps->lRow;
			ps->lRow = -1;
			long lRetTmp = CallOldWndProc(TIM_Scroll, wParam, lParam);
			ps->lRow = lTmp;
			if (!lRetTmp)
				goto end;
		}

		// Zeile scrollen
		if (ps->lRow != -1)
		{
			// Ende, wenn keine gültige Zeile
			if (!IsRow(ps->lRow))
				goto end;

			// Ende, wenn Zeile nicht sichtbar
			if (Ctd.TblQueryRowFlags(m_hWndTbl, ps->lRow, ROW_Hidden))
				goto end;

			// Metriken...
			CMTblMetrics tm(m_hWndTbl);

			// Split-Zeile?
			BOOL bSplitRow = (ps->lRow < 0);

			// Aktuell sichtbaren Bereich ermitteln
			long lRowsTop, lRowsBottom;
			long lVisMin, lVisMax;
			if (bSplitRow)
			{
				lRowsTop = tm.m_SplitRowsTop;
				lRowsBottom = tm.m_SplitRowsBottom;
				//if ( !MTblQueryVisibleSplitRange( m_hWndTbl, &lVisMin, &lVisMax ) )
				if (!QueryVisRange(lVisMin, lVisMax, TRUE))
					goto end;
			}
			else
			{
				lRowsTop = tm.m_RowsTop;
				lRowsBottom = tm.m_RowsBottom;
				if (!Ctd.TblQueryVisibleRange(m_hWndTbl, &lVisMin, &lVisMax))
					goto end;
			}
			long lRowsHt = (lRowsBottom - lRowsTop);

			// Auf geht's...
			long lRowScrTop = TBL_Error;
			long lScr = wParam;

			// Automatisch scrollen?
			if (lScr == TBL_AutoScroll)
			{
				if (ps->lRow < lVisMin)
					lScr = TBL_ScrollTop;
				else if (ps->lRow > lVisMax)
					lScr = TBL_ScrollBottom;
				else
				{
					// Schon im sichtbaren Bereich, nix zu tun...
					lRet = 1;
					goto end;
				}
			}


			// Nach oben scrollen?
			if (lScr == TBL_ScrollTop)
			{
				lRowScrTop = ps->lRow;
			}

			// Nach unten scrollen?
			if (lScr == TBL_ScrollBottom)
			{
				long lMinRow = bSplitRow ? TBL_MinSplitRow : 0;
				long lRow = ps->lRow;
				long lSumHt = 0;

				lRowScrTop = lRow;
				while (lRow != TBL_Error)
				{
					lSumHt += GetEffRowHeight(lRow, &tm);
					if (lSumHt > lRowsHt)
						break;

					lRowScrTop = lRow;

					if (!FindPrevRow(&lRow, 0, ROW_Hidden, lMinRow))
						lRow = TBL_Error;
				}
			}


			// Scrollen...
			if (lRowScrTop != TBL_Error)
			{
				long lPos = TBL_Error;

				// Scrollbar ermitteln...
				HWND hWndScr;
				int iScrCtl;
				if (GetScrollBarV(&tm, lRowScrTop < 0, &hWndScr, &iScrCtl))
				{
					// Sichtbare Position ermitteln
					lPos = m_Rows->GetRowVisPos(lRowScrTop);
					if (lPos != TBL_Error)
					{
						if (bSplitRow)
							lPos += TBL_MinSplitRow;

						// Aktuelle Scrollrange ermitteln
						int iMin = 0, iMax = 0;
						//GetScrollRange( hWndScr, iScrCtl, &iMin, &iMax );
						iMax = CalcMaxScrollRangeV(lRowScrTop < 0, lPos);

						// Scrollposition größer Maximum?
						if (iMax != TBL_Error && lPos > iMax)
						{
							// Position = Maximum
							lPos = iMax;

							// Zeile neu ermitteln
							long lVisPos = lPos;
							if (bSplitRow)
								lVisPos -= TBL_MinSplitRow;

							lRowScrTop = m_Rows->GetVisPosRow(lVisPos, bSplitRow);
						}
					}
				}

				if (lRowScrTop != TBL_Error && lPos != TBL_Error)
				{
					// Scrollen
					BOOL bScrolled = FALSE;
					if (lRowScrTop != lVisMin)
					{
						int iScrY;
						long lxRowScrLog, lxRowTopLog;
						RECT rClip, rScroll;

						SetRectEmpty(&rScroll);

						// Logische X-Position der obersten sichtbaren Zeile
						lxRowTopLog = m_Rows->GetRowLogX(lVisMin, m_ppd->tm.m_RowHeight);

						// Logische X-Position der nach oben zu scrollenden Zeile
						lxRowScrLog = m_Rows->GetRowLogX(lRowScrTop, m_ppd->tm.m_RowHeight);

						// Oberkante Scroll-Rechteck = Oberkante Zeilen
						rScroll.top = lRowsTop;

						// Unterkante Scroll-Rechteck = Unterkante Zeilen
						rScroll.bottom = lRowsBottom;

						// Zu scrollende Pixel
						iScrY = lxRowTopLog - lxRowScrLog;

						// Scroll-Rechteck komplettieren
						rScroll.left = tm.m_ClientRect.left;
						rScroll.right = tm.m_ClientRect.right;

						// Clip-Rechteck setzen
						rClip.top = lRowsTop;
						rClip.bottom = lRowsBottom;
						rClip.left = tm.m_ClientRect.left;
						rClip.right = tm.m_ClientRect.right;

						// Scrollen
						bScrolled = ScrollWindow_Hooked(m_hWndTbl, 0, iScrY, &rScroll, &rClip);
					}

					// Scrollrange anpassen + Scrollposition setzen
					AdaptScrollRangeV(lRowScrTop < 0, lPos);
					FlatSB_SetScrollPos_Hooked(hWndScr, iScrCtl, lPos, TRUE);

					// Listclip anpassen
					if (bScrolled)
						AdaptListClip();
				}
			}
		}
	}

end:
	AfterScroll();

	return lRet;
}


// OnSetContext
LRESULT MTBLSUBCLASS::OnSetContext(WPARAM wParam, LPARAM lParam)
{
	long lRet = CallOldWndProc(TIM_SetContext, wParam, lParam);
	if (lRet)
		// Gültige Zeile sicherstellen
		m_Rows->EnsureValid(lParam);

	return lRet;
}


// OnSetCursor
LRESULT MTBLSUBCLASS::OnSetCursor(WPARAM wParam, LPARAM lParam)
{
	BOOL bCallOldProc = TRUE;
	long lRet = 0;

	WORD wHitTest = LOWORD(lParam);
	WORD wMouseMsg = HIWORD(wParam);

	if ((HWND)wParam == m_hWndTbl)
	{
		if (SendMessage(m_hWndTbl, MTM_SetCursor, wParam, lParam) == MTBL_NODEFAULTACTION)
			return 1;
	}

	if ((HWND)wParam == m_hWndTbl && wHitTest == HTCLIENT)
	{

		if (QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT) && m_bRowSizing)
		{
			SetCursor(LoadCursor((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDC_SIZE_ROW)));
			return 1;
		}

		DWORD dwFlags, dwMFlags;
		HWND hWndCol;
		long lRow;
		POINT p;

		GetCursorPos(&p);
		ScreenToClient(m_hWndTbl, &p);
		if (ObjFromPoint(p, &lRow, &hWndCol, &dwFlags, &dwMFlags))
		{
			if (lRow != TBL_Error && hWndCol && !(dwFlags & TBL_YOverColumnHeader || dwFlags & TBL_YOverSplitBar))
			{
				if (IsOverHyperlink(hWndCol, lRow, dwMFlags))
				{
					SetCursor(LoadCursor((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDC_HYPERLINK)));
					bCallOldProc = FALSE;
				}
				else if (dwMFlags & MTOFP_OVERNODE)
				{
					SetCursor(LoadCursor((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDC_REVARROW)));
					bCallOldProc = FALSE;
				}
				else if ((dwMFlags & MTOFP_OVERBTN) || (dwMFlags & MTOFP_OVERDDBTN))
				{
					SetCursor(LoadCursor(NULL, IDC_ARROW));
					bCallOldProc = FALSE;
				}
				else if (!CanSetFocusCell(hWndCol, lRow))
				{
					SetCursor(LoadCursor((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDC_REVARROW)));
					bCallOldProc = FALSE;
				}
				else
				{
					HWND hWndEffCol;
					long lEffRow;
					if (!GetMergeCellEx(hWndCol, lRow, hWndEffCol, lEffRow))
					{
						hWndEffCol = hWndCol;
						lEffRow = lRow;
					}
					//int iCellType = Ctd.TblGetCellType( hWndMergeCol );
					int iCellType = GetEffCellType(lEffRow, hWndEffCol);
					if (iCellType == COL_CellType_CheckBox)
						SetCursor(LoadCursor(NULL, IDC_ARROW));
					else
						SetCursor(LoadCursor(NULL, IDC_IBEAM));

					bCallOldProc = FALSE;
				}
			}
			else if (hWndCol && (dwFlags & TBL_YOverColumnHeader) && (dwMFlags & MTOFP_OVERBTN))
			{
				SetCursor(LoadCursor(NULL, IDC_ARROW));
				bCallOldProc = FALSE;
			}
			else if ((dwFlags & TBL_XOverRowHeader) && (GetRowHdrFlags() & TBL_RowHdr_AllowRowSizing))
			{
				if (dwMFlags & MTOFP_OVERROWHDRSEP)
				{
					SetCursor(LoadCursor((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDC_SIZE_ROW)));
					bCallOldProc = FALSE;
				}
				else if (!(dwFlags & TBL_YOverColumnHeader) && !(dwFlags & TBL_YOverSplitBar))
				{
					SetCursor(LoadCursor(NULL, IDC_ARROW));
					bCallOldProc = FALSE;
				}
			}
			else if (QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER) && lRow == TBL_Error && !(dwFlags & TBL_XOverRowHeader) && !(dwFlags & TBL_YOverColumnHeader) && !(dwFlags & TBL_YOverSplitBar))
			{
				if (hWndCol && SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Editable))
					SetCursor(LoadCursor(NULL, IDC_IBEAM));
				else
					SetCursor(LoadCursor((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDC_REVARROW)));

				bCallOldProc = FALSE;
			}
		}
	}

	if (bCallOldProc)
		// Alte Fensterprozedur aufrufen
		lRet = CallOldWndProc(WM_SETCURSOR, wParam, lParam);
	else
		lRet = 1;

	return lRet;
}


// OnSetFocus
LRESULT MTBLSUBCLASS::OnSetFocus(WPARAM wParam, LPARAM lParam)
{
	if (!m_bSettingFocus)
		m_bNoFocusPaint = FALSE;

	long lRet = CallOldWndProc(WM_SETFOCUS, wParam, lParam);

	if (!m_bSettingFocus && MustSuppressRowFocusCTD())
	{
		KillFocusFrameI();
	}

	UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);

	return lRet;
}


// OnSetFocusCell
LRESULT MTBLSUBCLASS::OnSetFocusCell(WPARAM wParam, LPARAM lParam)
{
	typedef struct
	{
		long lRow;
		HWND hWndCol;
		long lEditMin;
		long lEditMax;
	} SETFOCUSCELL, *PSETFOCUSCELL;

	long lRet = 0;

	BOOL bMouseClickSim = m_bMouseClickSim;
	m_bMouseClickSim = FALSE;

	PSETFOCUSCELL psfc = (PSETFOCUSCELL)lParam;
	HWND hWndMergeCol;
	long lMergeRow, lOrgRow = psfc->lRow;
	if (GetMergeCellEx(psfc->hWndCol, psfc->lRow, hWndMergeCol, lMergeRow))
	{
		psfc->hWndCol = hWndMergeCol;
		psfc->lRow = lMergeRow;
	}

	static BOOL s_bSendingBeforeEnterEdit = FALSE;
	BOOL bExit = FALSE;

	if (s_bSendingBeforeEnterEdit)
		bExit = TRUE;
	else
	{
		s_bSendingBeforeEnterEdit = TRUE;
		if (SendMessage(m_hWndTbl, MTM_BeforeEnterEditMode, WPARAM(psfc->hWndCol), LPARAM(psfc->lRow)) == MTBL_NODEFAULTACTION)
			bExit = TRUE;
		s_bSendingBeforeEnterEdit = FALSE;
	}

	if (bExit)
		return lRet;

	BOOL bTrapColSelChanges = MustTrapColSelChanges();
	if (bTrapColSelChanges)
		BeginColSelTrap();

	BOOL bTrapRowSelChanges = MustTrapRowSelChanges();
	if (bTrapRowSelChanges)
		BeginRowSelTrap();

	// Kann Focus gesetzt werden?
	if (Ctd.IsWindowEnabled(psfc->hWndCol) && !CanSetFocusCell(psfc->hWndCol, psfc->lRow))
	{
		if (Ctd.TblSetFocusRow(m_hWndTbl, lOrgRow))
		{
			Ctd.TblSetRowFlags(m_hWndTbl, lOrgRow, ROW_Selected, TRUE);
			lRet = 1;
		}

		if (bTrapColSelChanges)
			EndColSelTrap();

		if (bTrapRowSelChanges)
			EndRowSelTrap();

		return lRet;
	}

	int iValidate = BeforeSetFocusCell(psfc->hWndCol, psfc->lRow, TRUE);

	// Alte Fensterprozedur aufrufen
	m_pCounter->IncNoFocusUpdate();

	if (iValidate != VALIDATE_Cancel)
	{
		BOOL bScrollBefore = m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

		if (bScrollBefore)
			Ctd.TblScroll(m_hWndTbl, psfc->lRow, psfc->hWndCol, TBL_AutoScroll);

		m_lNoSetFocusCellProcessing++;
		if (bScrollBefore)
			m_bNoVScrolling = TRUE;

		lRet = CallOldWndProc(TIM_SetFocusCell, wParam, lParam);

		if (bScrollBefore)
			m_bNoVScrolling = FALSE;
		m_lNoSetFocusCellProcessing--;
	}
	else
		lRet = 0;

	m_pCounter->DecNoFocusUpdate();

	// Evtl. SAM_Click senden
	if (bMouseClickSim && lRet)
	{
		SendMessage(psfc->hWndCol, SAM_Click, 0, psfc->lRow);
		SendMessage(m_hWndTbl, SAM_Click, 0, psfc->lRow);
	}

	AfterSetFocusCell(psfc->hWndCol, psfc->lRow, bMouseClickSim);

	if (bTrapColSelChanges)
		EndColSelTrap();

	if (bTrapRowSelChanges)
		EndRowSelTrap();

	return lRet;
}


// OnSetFocusRow
LRESULT MTBLSUBCLASS::OnSetFocusRow(WPARAM wParam, LPARAM lParam)
{
	BOOL bOldVal = m_bNoFocusPaint;
	//m_bNoFocusPaint = TRUE;

	BOOL bScrollBefore = m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	// Scrollen
	if (bScrollBefore)
		Ctd.TblScroll(m_hWndTbl, lParam, NULL, TBL_AutoScroll);

	m_pCounter->IncNoFocusUpdate();
	if (bScrollBefore)
		m_bNoVScrolling = TRUE;

	long lRet = CallOldWndProc(TIM_SetFocusRow, wParam, lParam);

	if (bScrollBefore)
		m_bNoVScrolling = FALSE;
	m_pCounter->DecNoFocusUpdate();

	if (lRet)
	{
		// Focus wieder zeichnen
		m_bNoFocusPaint = FALSE;

		if (MustSuppressRowFocusCTD())
			KillFocusFrameI();

		// Zellen-Fokus
		UpdateFocus(lParam);
	}
	else
		m_bNoFocusPaint = bOldVal;

	return lRet;
}


// OnSetFont
LRESULT MTBLSUBCLASS::OnSetFont(WPARAM wParam, LPARAM lParam)
{
	m_hFont = (HFONT)wParam;
	m_CurCols->SetInvalid();
	long lRet = CallOldWndProc(WM_SETFONT, wParam, lParam);
	UpdateFocus(TBL_Error, NULL, UF_REDRAW_INVALIDATE);
	return lRet;
}


// OnSetLinesPerRow
LRESULT MTBLSUBCLASS::OnSetLinesPerRow(WPARAM wParam, LPARAM lParam)
{
	long lRet = CallOldWndProc(TIM_SetLinesPerRow, wParam, lParam);
	if (lRet)
	{
		if (m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
		{
			AdaptScrollRangeV(FALSE);
			AdaptScrollRangeV(TRUE);

		}

		FixSplitWindow();
		UpdateFocus();
	}

	return lRet;
}


// OnSetLockedColumns
LRESULT MTBLSUBCLASS::OnSetLockedColumns(WPARAM wParam, LPARAM lParam)
{
	long lRet = CallOldWndProc(TIM_SetLockedColumns, wParam, lParam);
	if (lRet)
	{
		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		UpdateFocus();
	}

	return lRet;
}


// OnSetRowFlags
LRESULT MTBLSUBCLASS::OnSetRowFlags(WPARAM wParam, LPARAM lParam)
{
	// Parameter ermitteln
	LPIS_SETROWFLAGS psrf = (LPIS_SETROWFLAGS)lParam;

	// Evtl. bestimmte Flags entfernen
	if (wParam & ROW_Selected)
	{
		BOOL bRemove = FALSE;

		BOOL bSuppressRowSelection = QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION);
		if ((bSuppressRowSelection && psrf->bSet) || (psrf->bSet == Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Selected)))
			bRemove = TRUE;

		if (bRemove)
			wParam = (wParam & ROW_Selected) ^ wParam;
	}

	if (wParam & ROW_Hidden)
	{
		BOOL bRemove = FALSE;

		if ((psrf->bSet == Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Hidden)) || (!psrf->bSet && QueryRowFlags(psrf->lRow, MTBL_ROW_HIDDEN)))
			bRemove = TRUE;

		if (bRemove)
			wParam = (wParam & ROW_Hidden) ^ wParam;
	}

	// Überhaupt noch was zum Setzen da?
	if (!wParam)
		return 1;

	// Vorsorglich KillEdit durchführen?
	if (IsInEditMode() && ((wParam & ROW_Selected) || (wParam & ROW_Hidden)))
		Ctd.TblKillEdit(m_hWndTbl);

	// Wird sich die Selektion ändern?
	BOOL bSelWillChange = (wParam & ROW_Selected) && Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Selected) != psrf->bSet && !Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Hidden);

	// Evtl. Verfolgung von Selektionsänderungen beginnen
	BOOL bTrapRowSelChanges = bSelWillChange && MustTrapRowSelChanges();
	if (bTrapRowSelChanges)
		BeginRowSelTrap(psrf->lRow);

	// Evtl. LockUpdate durchführen
	BOOL bLockUpdate = bSelWillChange && MustRedrawSelections() && !gbCtdHooked;
	if (bLockUpdate)
		bLockUpdate = LockWindowUpdate(m_hWndTbl);

	// Zeile vorher versteckt?
	BOOL bHiddenBefore = Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Hidden);

	// Zeilenhöhen vorhanden?
	BOOL bRowHeights = m_pCounter->GetRowHeights() > 0;

	// Kein Column-Header?
	BOOL bNoColHdr = QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);

	// Ggf. Zeile vorsorglich auf "hidden" setzen
	BOOL bSetHiddenBefore = FALSE;
	if ((wParam & ROW_Hidden) && (bRowHeights || bNoColHdr))
		bSetHiddenBefore = m_Rows->SetRowHidden(psrf->lRow, psrf->bSet);

	// Alte Fensterprozedur aufrufen
	m_pCounter->IncNoFocusUpdate();
	if ((wParam & ROW_Hidden) && (bRowHeights || bNoColHdr))
		m_bNoVScrolling = TRUE;

	long lRet = CallOldWndProc(TIM_SetRowFlags, wParam, lParam);

	if ((wParam & ROW_Hidden) && (bRowHeights || bNoColHdr))
		m_bNoVScrolling = FALSE;
	m_pCounter->DecNoFocusUpdate();

	// Fehlgeschlagen?
	if (!lRet)
	{
		if (bSetHiddenBefore)
			m_Rows->SetRowHidden(psrf->lRow, !psrf->bSet);
	}

	// Hat sich Sichtbarkeit der Zeile geändert?
	if (lRet && Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Hidden) != bHiddenBefore)
	{
		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		// Merge-Zeilen ungültig
		m_bMergeRowsInvalid = TRUE;

		// Zeile auf "hidden" setzen
		m_Rows->SetRowHidden(psrf->lRow, psrf->bSet);

		// Vertikalen Scrollbereich anpassen
		if (bRowHeights || bNoColHdr)
			AdaptScrollRangeV(psrf->lRow < 1);

		// Ggf. neu zeichnen
		if (bRowHeights || bNoColHdr)
			InvalidateRect(m_hWndTbl, NULL, FALSE);

		// Fokus aktualisieren
		UpdateFocus();
	}

	// Evtl. Header neu zeichnen
	if (!Ctd.TblQueryRowFlags(m_hWndTbl, psrf->lRow, ROW_Hidden))
	{
		DWORD dwFlags = ROW_HideMarks | ROW_UserFlag1 | ROW_UserFlag2 | ROW_UserFlag3 | ROW_UserFlag4 | ROW_UserFlag5;
		if (wParam & dwFlags || bRowHeights || bNoColHdr)
		{
			InvalidateRowHdr(psrf->lRow, FALSE, TRUE);
		}
	}

	// Evtl. Ende LockUpdate
	if (bLockUpdate)
		LockWindowUpdate(NULL);

	// Evtl. Verfolgung von Selektionsänderungen beenden
	if (bTrapRowSelChanges)
		EndRowSelTrap(psrf->lRow);

	// Evtl. innerhalb des Zeilenverbundes replizieren
	if (!m_bReplMergeRowsFlags)
	{
		LPMTBLMERGEROW pmr = FindMergeRow(psrf->lRow, FMR_ANY);
		if (pmr)
		{
			DWORD dwFlags = wParam;
			DWORD dwDelFlags = 0;

			// ROW_Selected entfernen
			dwDelFlags = ROW_Selected;

			// Ggf. User-Flags entfernen
			if (QueryFlags(MTBL_FLAG_NO_USER_ROW_FLAG_REPL))
				dwDelFlags |= ROW_UserFlag1 | ROW_UserFlag2 | ROW_UserFlag3 | ROW_UserFlag4 | ROW_UserFlag5;

			dwFlags = (dwFlags & dwDelFlags) ^ dwFlags;

			if (dwFlags)
			{
				m_bReplMergeRowsFlags = TRUE;
				//SetMergeRowFlags(pmr, wParam, psrf->bSet, FALSE, psrf->lRow);
				SetMergeRowFlags(pmr, dwFlags, psrf->bSet, FALSE, psrf->lRow);
				m_bReplMergeRowsFlags = FALSE;
			}
		}
	}

	return lRet;
}


// OnSetScrollRange
BOOL MTBLSUBCLASS::OnSetScrollRange(HWND hWnd, int iBar, int iMinPos, int iMaxPos, BOOL bRedraw, int iScrPos /*= TBL_Error*/, BOOL bForce /*=FALSE*/)
{
	// Vertikalen Scrollbereich anpassen?	
	BOOL bAdaptV = FALSE;
	BOOL bSplitBar = FALSE;

	if (m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER) || bForce)
	{
		int iRet = IsScrollBarV(hWnd, iBar);
		if (iRet)
		{
			bAdaptV = TRUE;
			if (iRet == 2)
				bSplitBar = TRUE;
		}
	}

	// Wenn Anpassung der vertikalen Scrollbereichs stattfinden soll...
	if (bAdaptV)
	{
		// Max. Scrollbereich ermitteln
		long lMaxPos = CalcMaxScrollRangeV(bSplitBar, iScrPos);

		// Max. Scrollbereich festlegen
		if (lMaxPos != TBL_Error)
			iMaxPos = lMaxPos;
	}

	// Max. Scrollbereich setzen
	m_bSettingScrollRange = TRUE;
	BOOL bOk = FlatSB_SetScrollRange_Hooked(hWnd, iBar, iMinPos, iMaxPos, bRedraw);
	// Defect #181/185/186: Vertikale Scrollbar nicht sichtbar
	/*if (bOk && bAdaptV)
	{
		BOOL bShowVScroll = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_ShowVScroll);
		bOk = FlatSB_ShowScrollBar(hWnd, iBar, bShowVScroll || bSplitBar ? TRUE : iMaxPos > iMinPos);
	}*/
	//
	m_bSettingScrollRange = FALSE;

	return bOk;
}


// OnSetTableFlags
LRESULT MTBLSUBCLASS::OnSetTableFlags(WPARAM wParam, LPARAM lParam)
{
	BOOL bSetSingleSelection = (wParam & TBL_Flag_SingleSelection), bSuppressRowSelection = QueryFlags(MTBL_FLAG_SUPPRESS_ROW_SELECTION), bTrapColSelChanges, bTrapRowSelChanges;
	long lRet;

	bTrapColSelChanges = bSetSingleSelection && MustTrapColSelChanges();
	bTrapRowSelChanges = bSetSingleSelection && MustTrapRowSelChanges();

	if (bTrapColSelChanges)
		BeginColSelTrap();

	if (bTrapRowSelChanges)
		BeginRowSelTrap();

	BOOL bLockUpdate = bSetSingleSelection && MustRedrawSelections() && !gbCtdHooked;
	if (bLockUpdate)
		bLockUpdate = LockWindowUpdate(m_hWndTbl);

	lRet = CallOldWndProc(TIM_SetTableFlags, wParam, lParam);

	if (bLockUpdate)
		LockWindowUpdate(NULL);

	if (bSuppressRowSelection || m_bCellMode)
		if (bSetSingleSelection && lParam)
			Ctd.TblClearSelection(m_hWndTbl);

	if (bTrapColSelChanges)
		EndColSelTrap();

	if (bTrapRowSelChanges)
		EndRowSelTrap();

	return lRet;
}


// OnSize
LRESULT MTBLSUBCLASS::OnSize(WPARAM wParam, LPARAM lParam)
{
	BOOL bAdaptScrollRangeV = m_pCounter->GetRowHeights() > 0 || QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER);
	BOOL bLockPaint = FALSE;

	if (bAdaptScrollRangeV)
	{
		m_bNoVScrolling = TRUE;

		if (wParam == SIZE_MAXIMIZED || wParam == SIZE_RESTORED)
		{
			if (!m_bSettingScrollRange)
			{
				bLockPaint = LockPaint() > 0;
				AdaptScrollRangeV(FALSE);
				AdaptScrollRangeV(TRUE);
			}
		}
	}

	long lRet = 0;
	lRet = CallOldWndProc(WM_SIZE, wParam, lParam);

	if (bAdaptScrollRangeV)
	{
		m_bNoVScrolling = FALSE;
	}

	if (bLockPaint)
		UnlockPaint();

	// Bestimmte Bereiche neu zeichnen
	if (wParam == SIZE_MAXIMIZED || wParam == SIZE_RESTORED)
	{
		if (bAdaptScrollRangeV && !m_bSettingScrollRange)
		{
			InvalidateRect(m_hWndTbl, NULL, FALSE);
		}
		else
			InvalidateCol(NULL, FALSE);
	}

	AdaptListClip();
	UpdateFocus(TBL_Error, NULL, m_bCellMode ? UF_REDRAW_AUTO : UF_REDRAW_NO_INVALIDATE);

	GetClientRect(m_hWndTbl, &m_LastClientRectOnSize);

	return lRet;
}


// OnSwapRows
LRESULT MTBLSUBCLASS::OnSwapRows(WPARAM wParam, LPARAM lParam)
{
	MTBLSWAPROWS sr, srorg;
	srorg.lRow1 = ((LPMTBLSWAPROWS)lParam)->lRow1;
	srorg.lRow2 = ((LPMTBLSWAPROWS)lParam)->lRow2;
	if (srorg.lRow1 < 0 || srorg.lRow2 < 0) return 0;
	sr = srorg;

	// Zeilennummern anpassen, wenn mit TBL_Adjust gelöschte Zeilen vorhanden sind
	if (m_DelAdjRows->size() > 0)
	{
		for (LongList::iterator l = m_DelAdjRows->begin(); l != m_DelAdjRows->end(); ++l)
		{
			if (*l < srorg.lRow1)
				sr.lRow1--;
			if (*l < srorg.lRow2)
				sr.lRow2--;
		}
	}

	// Alte Fensterprozedur aufrufen
	long lRet = CallOldWndProc(TIM_SwapRows, wParam, (LPARAM)&sr);

	if (lRet)
	{
		// Merge-Zellen ungültig
		m_bMergeCellsInvalid = TRUE;

		// Merge-Zeilen ungültig
		m_bMergeRowsInvalid = TRUE;

		// Zeilen vertauschen
		m_Rows->SwapRows(srorg.lRow1, srorg.lRow2);

		// Evtl. Bereich Zellenselektion/Tastatur zurücksetzen
		if (srorg.lRow1 == m_CellSelRangeKey.lRowFrom || srorg.lRow2 == m_CellSelRangeKey.lRowFrom)
			m_CellSelRangeKey.SetEmpty();

		// Zellen-Fokus aktualisieren
		if (m_bCellMode)
		{
			long lRowFocus = GetRowNrFocus();
			if (lRowFocus == srorg.lRow1 || lRowFocus == srorg.lRow1)
				UpdateFocus();
		}

		// Erweiterte Nachricht senden
		if (m_bExtMsgs)
			PostMessage(m_hWndTbl, MTM_RowsSwapped, srorg.lRow1, srorg.lRow2);
	}

	return lRet;
}


// OnSysKeyDown
LRESULT MTBLSUBCLASS::OnSysKeyDown(WPARAM wParam, LPARAM lParam)
{
	if (ProcessHotkey(WM_SYSKEYDOWN, wParam, lParam))
		return 0;
	return CallOldWndProc(WM_SYSKEYDOWN, wParam, lParam);
}


// OnTipPaint
LRESULT MTBLSUBCLASS::OnTipPaint(HWND hWndTip, WPARAM wParam, LPARAM lParam)
{
	RECT rUpd;
	if (GetUpdateRect(hWndTip, &rUpd, FALSE))
	{
		// Zeichnen beginnen
		PAINTSTRUCT ps;
		BeginPaint(hWndTip, &ps);

		if (m_pTip)
		{
			DrawTip(hWndTip, ps.hdc);
		}
		else
			ValidateRect(hWndTip, &rUpd);

		// Zeichnen beenden
		EndPaint(hWndTip, &ps);
	}

	return 0;
}


// OnTipPrintClient
LRESULT MTBLSUBCLASS::OnTipPrintClient(HWND hWndTip, WPARAM wParam, LPARAM lParam)
{
	if (lParam & PRF_CHECKVISIBLE && !IsWindowVisible(hWndTip))
		return 0;

	DrawTip(hWndTip, (HDC)wParam);

	return 0;
}


// OnTipShow
LRESULT MTBLSUBCLASS::OnTipShow(HWND hWndTip, WPARAM wParam, LPARAM lParam)
{
	long lRet = DefWindowProc(hWndTip, WM_SHOWWINDOW, wParam, lParam);
	return lRet;
}


// OnVScroll
LRESULT MTBLSUBCLASS::OnVScroll(WPARAM wParam, LPARAM lParam)
{
	long lRet = 0;

	BOOL bRemWaitCursor, bThumbtrackScroll;
	bRemWaitCursor = FALSE;
	bThumbtrackScroll = QueryFlags(MTBL_FLAG_THUMBTRACK_VSCROLL) && LOWORD(wParam) == SB_THUMBTRACK /*&& ( lParam ? TRUE : !Ctd.TblAnyRows( hWndTbl, 0, ROW_InCache ) )*/;

	if (bThumbtrackScroll)
	{
		wParam = MAKELONG(SB_THUMBPOSITION, HIWORD(wParam));
		if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_ShowWaitCursor))
			bRemWaitCursor = TRUE;
	}

	WORD wScrollCode = LOWORD(wParam);
	BOOL bScrollFuncs = wScrollCode != SB_THUMBTRACK && wScrollCode != SB_ENDSCROLL;

	if (bScrollFuncs)
		BeforeScroll();

	if (bRemWaitCursor)
		Ctd.TblSetTableFlags(m_hWndTbl, TBL_Flag_ShowWaitCursor, FALSE);

	if ((m_pCounter->GetRowHeights() <= 0 && !QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER)) || !bScrollFuncs)
		lRet = CallOldWndProc(WM_VSCROLL, wParam, lParam);
	else
	{
		// Metriken...
		CMTblMetrics tm(m_hWndTbl);

		// Split-Bereich gescrollt?
		BOOL bSplitScroll = FALSE;
		if (lParam && tm.m_SplitBarTop > 0)
		{
			HWND hWndScr = NULL;
			int iScrCtl;
			if (GetScrollBarV(&tm, TRUE, &hWndScr, &iScrCtl) && hWndScr == (HWND)lParam)
				bSplitScroll = TRUE;
		}

		// Scrollbar ermitteln
		HWND hWndScr = NULL;
		int iScrCtl;
		if (!GetScrollBarV(&tm, bSplitScroll, &hWndScr, &iScrCtl))
			goto end;

		// Scrollinfo ermitteln
		SCROLLINFO si;
		memset(&si, 0, sizeof(SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);
		si.fMask = SIF_ALL;
		if (!FlatSB_GetScrollInfo(hWndScr, iScrCtl, &si))
			goto end;

		// Sichtbaren Zeilenbereich ermitteln
		long lVisMax, lVisMin;
		if (bSplitScroll)
		{
			lVisMin = TBL_MinSplitRow;
			lVisMax = TBL_MinSplitRow;
			//MTblQueryVisibleSplitRange( m_hWndTbl, &lVisMin, &lVisMax );
			QueryVisRange(lVisMin, lVisMax, TRUE);
		}
		else
		{
			lVisMin = 0;
			lVisMax = 0;
			Ctd.TblQueryVisibleRange(m_hWndTbl, &lVisMin, &lVisMax);
		}

		// Auf geht's...
		long lRowMax, lRowMin, lRowScr;
		switch (wScrollCode)
		{
		case SB_LINEDOWN:
			// Ende, wenn überhaupt nicht gescrollt werden kann
			if (si.nPos >= si.nMax)
				goto end;

			// Standard-Scrolling, wenn Höhe der obersten Zeile Standard ist...
			if ( /*GetEffRowHeight( lVisMin, &tm ) == tm.m_RowHeight*/ FALSE)
			{
				lRet = CallOldWndProc(WM_VSCROLL, wParam, lParam);
			}
			// ... ansonsten selbst scrollen
			else
			{
				// Zeile ermitteln, welche nach dem Scrollen unten sein soll
				lRowScr = lVisMin;
				lRowMax = bSplitScroll ? -1 : TBL_MaxRow;
				if (!FindNextRow(&lRowScr, 0, ROW_Hidden, lRowMax))
					goto end;

				// Scrollen
				Ctd.TblScroll(m_hWndTbl, lRowScr, NULL, TBL_ScrollTop);
			}

			break;

		case SB_LINEUP:
			// Ende, wenn überhaupt nicht gescrollt werden kann
			if (si.nPos <= si.nMin)
				goto end;

			// Zeile ermitteln, welche nach dem Scrollen oben sein soll
			lRowScr = lVisMin;
			lRowMin = bSplitScroll ? TBL_MinSplitRow : 0;
			if (!FindPrevRow(&lRowScr, 0, ROW_Hidden, lRowMin))
				goto end;

			// Standard-Scrolling, wenn Höhe der nach oben zu scrollenden Zeile Standard ist...
			if ( /*GetEffRowHeight( lRowScr, &tm ) == tm.m_RowHeight*/ FALSE)
			{
				lRet = CallOldWndProc(WM_VSCROLL, wParam, lParam);
			}
			// ... ansonsten selbst scrollen
			else
			{
				Ctd.TblScroll(m_hWndTbl, lRowScr, NULL, TBL_ScrollTop);
			}

			break;

		case SB_PAGEDOWN:
			// Ende, wenn überhaupt nicht gescrollt werden kann
			if (si.nPos >= si.nMax)
				goto end;

			// Zeile ermitteln, welche nach dem Scrollen oben sein soll			
			lRowScr = lVisMax;
			lRowMax = bSplitScroll ? -1 : TBL_MaxRow;
			if (!FindNextRow(&lRowScr, 0, ROW_Hidden, lRowMax))
				goto end;

			// Scrollen
			Ctd.TblScroll(m_hWndTbl, lRowScr, NULL, TBL_ScrollTop);

			break;

		case SB_PAGEUP:
			// Ende, wenn überhaupt nicht gescrollt werden kann
			if (si.nPos <= si.nMin)
				goto end;

			// Zeile ermitteln, welche nach dem Scrollen unten sein soll			
			lRowScr = lVisMin;
			lRowMin = bSplitScroll ? TBL_MinSplitRow : 0;
			if (!FindPrevRow(&lRowScr, 0, ROW_Hidden, lRowMin))
				goto end;

			// Scrollen
			Ctd.TblScroll(m_hWndTbl, lRowScr, NULL, TBL_ScrollBottom);

			break;

		case SB_THUMBPOSITION:
			// Ende, wenn ungültige Position
			if (si.nTrackPos < si.nMin || si.nTrackPos > si.nMax)
				goto end;

			// Zeile ermitteln, welche nach dem Scrollen oben sein soll
			lRowScr = m_Rows->GetVisPosRow(si.nTrackPos);

			// Scrollen
			Ctd.TblScroll(m_hWndTbl, lRowScr, NULL, TBL_ScrollTop);

			break;
		}
	}

end:
	if (bRemWaitCursor)
		Ctd.TblSetTableFlags(m_hWndTbl, TBL_Flag_ShowWaitCursor, TRUE);

	if (bScrollFuncs)
		AfterScroll();

	return lRet;
}


// PaintBtn
BOOL MTBLSUBCLASS::PaintBtn(LPMTBLPAINTBTN pb)
{
	// Knopf
	HTHEME hTheme = NULL;
	CTheme &th = theTheme();
	if (th.CanUse() && th.IsAppThemed() && !QueryFlags(MTBL_FLAG_IGNORE_WIN_THEME))
	{
		hTheme = th.OpenThemeData(m_hWndTbl, L"Button");
	}

	// Hintergrund
	if (!(pb->dwFlags & MTBL_BTN_NO_BKGND))
	{
		if (hTheme)
		{
			int iPart = BP_PUSHBUTTON;
			int iState = 0;
			if (pb->dwFlags & MTBL_BTN_DISABLED)
				iState = PBS_DISABLED;
			else
				iState = pb->bPushed ? PBS_PRESSED : PBS_NORMAL;
			th.DrawThemeBackground(hTheme, pb->hDC, iPart, iState, &pb->rBtn, NULL);
		}
		else
		{
			int iStyle = DFCS_BUTTONPUSH;
			if ((pb->dwFlags & MTBL_BTN_FLAT) || QueryFlags(MTBL_FLAG_BUTTONS_FLAT))
				iStyle |= DFCS_FLAT;
			if (pb->bPushed)
				iStyle |= DFCS_PUSHED;
			DrawFrameControl(pb->hDC, &pb->rBtn, DFC_BUTTON, iStyle);
		}
	}

	// Bilddaten ermitteln
	POINT ptImg = { 0, 0 };
	SIZE siImg = { 0, 0 };
	BOOL bImage = MImgIsValidHandle(pb->hImage);
	if (bImage)
		MImgGetSize(pb->hImage, &siImg);

	// Clipping
	BOOL bRestoreClipRgn = FALSE;
	HRGN hRgnCur = CreateRectRgn(0, 0, 0, 0);
	if (hRgnCur && pb->bRestoreClipRgn)
	{
		if (GetClipRgn(pb->hDC, hRgnCur) > 0)
			bRestoreClipRgn = TRUE;
	}
	HRGN hRgn = CreateRectRgnIndirect(&pb->rBtn);
	if (hRgn)
	{
		if (bRestoreClipRgn)
			CombineRgn(hRgn, hRgn, hRgnCur, RGN_AND);
		SelectClipRgn(pb->hDC, hRgn);
		DeleteObject(hRgn);
	}

	// Text zeichnen
	int iTxtLen = 0;
	POINT ptText = { 0, 0 };
	SIZE siTxt = { 0, 0 };
	BOOL bText = pb->lpsText && (iTxtLen = lstrlen(pb->lpsText));
	if (bText)
	{
		HGDIOBJ hOldFont = SelectObject(pb->hDC, pb->hFont);
		int iOldBkMode = SetBkMode(pb->hDC, TRANSPARENT);

		COLORREF clrTextOld = GetTextColor(pb->hDC);

		RECT rd;
		SetRectEmpty(&rd);
		DrawText(pb->hDC, pb->lpsText, iTxtLen, &rd, DT_SINGLELINE | DT_CALCRECT);

		siTxt.cx = (rd.right - rd.left);
		siTxt.cy = (rd.bottom - rd.top);
		if (siImg.cx > 0)
		{
			ptText.x = max(pb->rBtn.left, pb->rBtn.left + (((pb->rBtn.right - pb->rBtn.left) - (siImg.cx + 2 + siTxt.cx)) / 2)) + siImg.cx + 2;
		}
		else
			ptText.x = max(pb->rBtn.left, pb->rBtn.left + (((pb->rBtn.right - pb->rBtn.left) - siTxt.cx) / 2));
		ptText.y = max(pb->rBtn.top, pb->rBtn.top + (((pb->rBtn.bottom - pb->rBtn.top) - siTxt.cy) / 2));

		if (pb->bPushed && (!hTheme || (pb->dwFlags & MTBL_BTN_NO_BKGND)))
		{
			ptText.x += 1;
			ptText.y += 1;
		}

		UINT uiFlags = DST_TEXT;
		if (pb->dwFlags & MTBL_BTN_DISABLED)
		{
			if (hTheme)
			{
				COLORREF clrText = RGB(0, 0, 0);

				int iPart = BP_PUSHBUTTON;
				int iState = PBS_DISABLED;
				HRESULT hr = th.GetThemeColor(hTheme, iPart, iState, TMT_TEXTCOLOR, &clrText);

				SetTextColor(pb->hDC, clrText);
			}
			else
				uiFlags |= DSS_DISABLED;
		}

		DrawState(pb->hDC, NULL, NULL, (LPARAM)pb->lpsText, (WPARAM)iTxtLen, ptText.x, ptText.y, siTxt.cx, siTxt.cy, uiFlags);


		SetBkMode(pb->hDC, iOldBkMode);
		SetTextColor(pb->hDC, clrTextOld);
		SelectObject(pb->hDC, hOldFont);
	}

	// Nur Bild ohne Hintergrund?
	BOOL bImageOnly = bImage && !bText && (pb->dwFlags & MTBL_BTN_NO_BKGND);

	// Bild zeichnen
	if (bImage)
	{
		if (siTxt.cx > 0)
			ptImg.x = max(pb->rBtn.left, pb->rBtn.left + (((pb->rBtn.right - pb->rBtn.left) - (siImg.cx + 2 + siTxt.cx)) / 2));
		else
			ptImg.x = max(pb->rBtn.left, pb->rBtn.left + (((pb->rBtn.right - pb->rBtn.left) - siImg.cx) / 2));
		ptImg.y = max(pb->rBtn.top, pb->rBtn.top + (((pb->rBtn.bottom - pb->rBtn.top) - siImg.cy) / 2));
		if (pb->bPushed && (!hTheme || (pb->dwFlags & MTBL_BTN_NO_BKGND)))
		{
			ptImg.x += 1;
			ptImg.y += 1;
		}

		RECT rImg = { ptImg.x, ptImg.y, ptImg.x + siImg.cx, ptImg.y + siImg.cy };
		DWORD dwFlags = (pb->dwFlags & MTBL_BTN_DISABLED) ? MIMG_DRAW_DISABLED : 0;
		if (bImageOnly && pb->bDrawInverted)
			dwFlags |= MIMG_DRAW_INVERTED;
		MImgDraw(pb->hImage, pb->hDC, &rImg, dwFlags);
	}

	// Ggf. Theme schließen
	if (hTheme)
		th.CloseThemeData(hTheme);

	// Ggf. invertieren
	if (pb->bDrawInverted && !bImageOnly)
		InvertRect_Hooked(pb->hDC, &pb->rBtn);

	// Clipping
	if (bRestoreClipRgn)
		SelectClipRgn(pb->hDC, hRgnCur);
	if (hRgnCur)
		DeleteObject(hRgnCur);

	return TRUE;
}

// PaintCell
BOOL MTBLSUBCLASS::PaintCell(MTBLPAINTCELL &pc)
{
	HRGN rgnClip = NULL;
	CMTblLineDef ld;

	int iCellLeadingY = pc.iCellLeadingY;

	// Font selektieren
	HGDIOBJ hOldFont = SelectObject(m_ppd->hDC, pc.hTextFont);

	// Fontmaße ermittteln
	TEXTMETRIC txtm;
	GetTextMetrics(m_ppd->hDC, &txtm);


	// Bild-Rechteck ermitteln
	if (pc.Img.GetHandle())
	{
		BOOL bBtn = FALSE, bDDBtn = FALSE;
		HWND hWndCol = NULL;
		long lRow = TBL_Error;
		if (pc.bGroupColHdr)
		{
			hWndCol = NULL;
			lRow = TBL_Error;
		}
		else if (m_ppd->bRowHdrMirror)
		{
			hWndCol = m_ppd->hWndRowHdrCol;
			lRow = m_ppd->lRow;
		}
		else
		{
			hWndCol = m_ppd->pEffPaintCol->hWnd;
			if (m_ppd->bColHdr || m_ppd->bEmptyRow)
				lRow = TBL_Error;
			else
				lRow = m_ppd->lEffPaintRow;
		}

		// *** Check Image ***
		BOOL bCheckImage = FALSE;

		if (bCheckImage)
		{
			HANDLE hImg = pc.Img.GetHandle();
			if (hImg)
			{
				if (pc.Img.IsValidImageHandle(hImg) && !MImgIsValidHandle(hImg))
				{
					TCHAR szBuf[255];
					tstring s = _T("");

					wsprintf(szBuf, _T("Invalid image handle: %ld\n"), hImg);
					s += szBuf;

					wsprintf(szBuf, _T("Column handle: %ld\n"), hWndCol);
					s += szBuf;

					wsprintf(szBuf, _T("Row: %d"), lRow);
					s += szBuf;

					MessageBox(m_hWndTbl, s.c_str(), _T("M!Table image check"), MB_OK);

					pc.Img.SetHandle(NULL, NULL);
				}
			}
		}

		if (hWndCol && lRow != TBL_Error && !m_ppd->bColHdr)
		{
			BOOL bPermanent = TRUE;
			if (m_ppd->hWndFocusCol && hWndCol == m_ppd->hWndFocusCol && lRow == m_ppd->lFocusRow)
				bPermanent = FALSE;
			bBtn = MustShowBtn(hWndCol, lRow, TRUE, bPermanent);
			bDDBtn = MustShowDDBtn(hWndCol, lRow, TRUE, bPermanent);
		}
		else if (pc.Item.IsRealColHdr())
		{
			CMTblBtn* pBtn = GetEffItemBtn(pc.Item);
			bBtn = pBtn && pBtn->m_bShow;
		}
		long lCellWid = pc.rCell.right - pc.rCell.left;
		long lCellHt = pc.rCell.bottom - pc.rCell.top;
		if (!GetImgRect(pc.Img, lRow, hWndCol, pc.rCell.left, pc.rCell.top, lCellWid, lCellHt, pc.rText.top, txtm.tmHeight, m_ppd->tm.m_CellLeadingX, iCellLeadingY, bBtn, bDDBtn, &pc.rImg))
			pc.Img.SetHandle(NULL, NULL);
	}

	// Knoten-Rechteck ermitteln
	if (pc.bNode)
		GetNodeRect(pc.rCell.left, pc.rText.top, txtm.tmHeight, m_ppd->iRowLevel, m_ppd->tm.m_CellLeadingX, iCellLeadingY, &pc.rNode);

	// Clipping-Rechteck der kompletten Zelle setzen 
	RECT rClipCell;
	rClipCell.left = pc.rCellVis.left;
	rClipCell.top = pc.rCellVis.top;
	rClipCell.right = min(pc.rCellVis.right + 1, m_ppd->tm.m_ClientRect.right);
	rClipCell.bottom = pc.rCellVis.bottom + 1;

	RECT rClipCellDP;
	CopyRect(&rClipCellDP, &rClipCell);
	LPtoDP(m_ppd->hDC, (LPPOINT)&rClipCellDP, 2);

	// Clipping für Hintergrund + Linien
	rgnClip = CreateRectRgnIndirect(&rClipCellDP);
	if (rgnClip)
	{
		SelectClipRgn(m_ppd->hDC, rgnClip);
		DeleteObject(rgnClip);
	}

	// Hintergrund füllen
	BOOL bDeleteBrush = FALSE;
	BOOL bGradient = (m_ppd->bColHdr || m_ppd->pPaintCol->bRowHdr) && QueryFlags(MTBL_FLAG_GRADIENT_HEADERS);
	BOOL bInvBkgnd = (pc.Bits.Selected && pc.Bits.NoSelInvBkgnd && !m_ppd->bAnySelColors && !m_bySelDarkening);
	COLORREF crBackColor = (pc.Bits.Selected && m_dwSelBackColor != MTBL_COLOR_UNDEF && pc.Hili.dwBackColor == MTBL_COLOR_UNDEF ? GetDrawColor(m_dwSelBackColor) : pc.crBackColor);
	crBackColor = (bInvBkgnd ? GetDrawColor(0xffffff - GetNormalColor(crBackColor)) : pc.Bits.Selected ? GetDrawColor(DarkenColor(crBackColor, m_bySelDarkening)) : crBackColor);
	HBRUSH hBrush = NULL;
	RECT r;

	CopyRect(&r, &pc.rCell);
	r.right += 1;
	r.bottom += 1;

	if (bGradient)
		bGradient = GradientRect(m_ppd->hDC, /*pc.rCell*/r, RGB(255, 255, 255), GetNormalColor(crBackColor), GRFM_TB);
	else
	{
		if (crBackColor != m_ppd->clrTblBackColor)
		{
			hBrush = CreateSolidBrush(crBackColor);
			bDeleteBrush = TRUE;
		}
		else
			hBrush = m_ppd->hBrTblBack;

		if (hBrush)
		{
			FillRect(m_ppd->hDC, &r, hBrush);

			if (bDeleteBrush)
				DeleteObject(hBrush);
		}
	}

	if (m_ppd->bClipFocusArea && !m_ppd->pPaintCol->bRowHdr)
	{
		// Focus-Bereich mit Hintergrund der Tabelle malen

		// Oben
		if (!m_ppd->bColHdr)
		{
			r.bottom = r.top + 1;
			if (bInvBkgnd)
			{
				hBrush = CreateSolidBrush(0xffffff - m_ppd->clrTblBackColor);
				bDeleteBrush = TRUE;
			}
			else
			{
				hBrush = m_ppd->hBrTblBack;
				bDeleteBrush = FALSE;
			}

			FillRect(m_ppd->hDC, &r, hBrush);

			if (bDeleteBrush)
				DeleteObject(hBrush);
		}

		// Unten
		if (!pc.bGroupColHdrTop)
		{
			r.bottom = pc.rCell.bottom + 1;
			r.top = r.bottom - 2;
			if (m_ppd->bColHdr)
			{
				hBrush = CreateSolidBrush(bInvBkgnd ? GetDrawColor(0xffffff - GetNormalColor(m_ppd->clrHdrBack)) : m_ppd->clrHdrBack);
				bDeleteBrush = TRUE;
			}
			else
			{
				if (bInvBkgnd)
				{
					hBrush = CreateSolidBrush(GetDrawColor(0xffffff - GetNormalColor(m_ppd->clrTblBackColor)));
					bDeleteBrush = TRUE;
				}
				else
				{
					hBrush = m_ppd->hBrTblBack;
					bDeleteBrush = FALSE;
				}
			}

			FillRect(m_ppd->hDC, &r, hBrush);

			if (bDeleteBrush)
				DeleteObject(hBrush);
		}
	}

	// 3D zeichnen
	if (m_ppd->bGrayedHeaders && (m_ppd->bColHdr || (m_ppd->pPaintCol->bRowHdr && !QueryFlags(MTBL_FLAG_SOLID_ROWHDR))) && !(pc.Bits.Selected && m_dwSelBackColor != MTBL_COLOR_UNDEF) && !bGradient)
	{
		COLORREF clr = (bInvBkgnd ? GetDrawColor(0xffffff - GetNormalColor(m_ppd->clr3DHiLight)) : m_ppd->clr3DHiLight);
		HPEN hPen = CreatePen(PS_SOLID, 1, clr);
		if (hPen)
		{
			HGDIOBJ hOldPen = SelectObject(m_ppd->hDC, hPen);

			// Links
			if (!(m_ppd->pPaintCol->bFirstVisible || (m_ppd->pPaintCol->bEmpty && m_ppd->bSuppressLastColLine) || (pc.bGroupColHdr && !pc.bFirstGroupColHdr && pc.pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOINNERVERTLINES))))
			{
				MoveToEx(m_ppd->hDC, pc.rCell.left, pc.rCell.top, NULL);
				LineTo(m_ppd->hDC, pc.rCell.left, pc.rCell.bottom + 1);
			}

			// Oben
			if (pc.bGroupColHdrTop || pc.bGroupColHdrNCH || !pc.bGroupColHdr)
			{
				MoveToEx(m_ppd->hDC, pc.rCell.left, pc.rCell.top, NULL);
				LineTo(m_ppd->hDC, pc.rCellVis.right + 1, pc.rCell.top);
			}

			SelectObject(m_ppd->hDC, hOldPen);
			DeleteObject(hPen);
		}
	}

	BOOL bOdd;
	COLORREF clr;
	HDC hDCSrc;
	HPEN hPen;
	int iMod, iOffs;
	if (pc.bRowLine)
	{
		ld.Init();
		GetEffItemLine(pc.HorzLineItem, MTLT_HORZ, ld);

		if (ld.Style != MTLS_INVISIBLE)
		{
			clr = GetDrawColor(ld.Color);

			if (ld.Style == MTLS_SOLID)
			{
				// Durchgezogene Linie
				if (!(m_ppd->pPaintCol->bRowHdr && !m_ppd->bColHdr && QueryFlags(MTBL_FLAG_SOLID_ROWHDR)))
				{					
					iOffs = 0;
					hPen = CreatePen(PS_SOLID, ld.Thickness, clr);
				}
			}
			else if (ld.Style == MTLS_DOT)
			{
				// Gepunktete Linie
				iMod = (m_ppd->pEffPaintCol->ptLogX.x - m_ppd->tm.m_RowHeaderRight) % 2;
				bOdd = (pc.dwBottomLog % 2 != 0) ? (iMod == 0) : (iMod != 0);
				iOffs = bOdd ? 1 : 0;

				LOGBRUSH lb;
				lb.lbColor = clr;
				lb.lbStyle = PS_SOLID;

				hPen = ExtCreatePen(PS_COSMETIC | PS_ALTERNATE, 1, &lb, 0, NULL);
			}
			else
				hPen = NULL;

			if (hPen)
			{
				HGDIOBJ hOldPen = SelectObject(m_ppd->hDC, hPen);

				MoveToEx(m_ppd->hDC, pc.rCellVis.left - iOffs, pc.rCell.bottom, NULL);
				LineTo(m_ppd->hDC, pc.rCellVis.right + 1, pc.rCell.bottom);

				SelectObject(m_ppd->hDC, hOldPen);
				DeleteObject(hPen);
			}
		}
	}

	if (pc.bColLine)
	{
		ld.Init();
		GetEffItemLine(pc.VertLineItem, MTLT_VERT, ld);

		if (ld.Style != MTLS_INVISIBLE)
		{
			clr = GetDrawColor(ld.Color);
			LPMTBLPAINTCOL pPaintCol = (pc.bMerge ? m_ppd->pLastMergeCol : m_ppd->pPaintCol);

			if (ld.Style == MTLS_SOLID)
			{
				// Durchgezogene Linie
				iOffs = 0;
				hPen = CreatePen(PS_SOLID, ld.Thickness, clr);
			}
			else if (ld.Style == MTLS_DOT)
			{
				iMod = pc.dwTopLog % 2;
				bOdd = pPaintCol->bOddLine ? (iMod == 0) : (iMod != 0);
				iOffs = bOdd ? 1 : 0;

				LOGBRUSH lb;
				lb.lbColor = clr;
				lb.lbStyle = PS_SOLID;

				hPen = ExtCreatePen(PS_COSMETIC | PS_ALTERNATE, 1, &lb, 0, NULL);
			}
			else
				hPen = NULL;

			if (hPen)
			{
				HGDIOBJ hOldPen = SelectObject(m_ppd->hDC, hPen);

				MoveToEx(m_ppd->hDC, pc.rCell.right, pc.rCellVis.top - iOffs, NULL);
				LineTo(m_ppd->hDC, pc.rCell.right, pc.rCellVis.bottom + 1);

				SelectObject(m_ppd->hDC, hOldPen);
				DeleteObject(hPen);
			}
		}		
	}

	// Tree-Linien
	if (m_TreeLineDef.Style != MTLS_INVISIBLE)
	{
		BOOL bTextRectDet = FALSE;
		clr = (m_TreeLineDef.Color == MTBL_COLOR_UNDEF ? 0 : m_TreeLineDef.Color);
		if ((pc.Bits.Selected && pc.Bits.NoSelInvTreeLines && !m_ppd->bAnySelColors && !m_bySelDarkening))
			clr = 0xffffff - clr;
		else if (pc.Bits.Selected && m_bySelDarkening)
			clr = DarkenColor(clr, m_bySelDarkening);
		clr = GetDrawColor(clr);
		/*int iRowMod = pc.dwTopLog % 2;*/
		int iStyle = (m_TreeLineDef.Style == MTLS_SOLID ? PS_SOLID : PS_DOT);
		RECT rTxt;

		hPen = CreatePen(iStyle, 1, clr);
		if (hPen)
		{
			HGDIOBJ hOldPen = SelectObject(m_ppd->hDC, hPen);

			MTBLDOTLINE dl;
			dl.hDC = m_ppd->hDC;
			dl.clr1 = clr;
			dl.clr2 = crBackColor;

			RECT rl;
			if (pc.bHorzTreeLine)
			{
				int iDrawLevel = QueryTreeFlags(MTBL_TREE_FLAG_FLAT_STRUCT) ? 0 : m_ppd->iRowLevel;
				if (iDrawLevel > 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE))
					iDrawLevel--;
				rl.left = pc.rCell.left + (m_ppd->tm.m_CellLeadingX - 1) + (iDrawLevel * m_iIndent) + (m_iIndent / 2) + 1;
				rl.top = pc.rText.top + (txtm.tmHeight / 2);
				if (pc.Img.GetHandle() && pc.rImg.left <= pc.rText.left)
					rl.right = pc.rImg.left;
				else
					rl.right = pc.rText.left - 1;
				rl.bottom = rl.top;
				if (iStyle == PS_SOLID)
				{
					MoveToEx(m_ppd->hDC, rl.left, rl.top, NULL);
					LineTo(m_ppd->hDC, rl.right, rl.bottom);
				}
				else if (iStyle == PS_DOT)
				{
					dl.rCoord = rl;
					dl.bVert = FALSE;
					dl.bOdd = TRUE;
					LineDDA(rl.left, rl.top, rl.right, rl.bottom, DotlineProc, (LPARAM)&dl);
				}
			}
			if (pc.bVertTreeLineToParent)
			{
				CMTblRow *pRow, *pParentRow;
				BOOL bDraw, bLastChild;

				for (int iRowLevel = m_ppd->iRowLevel; iRowLevel > 0; --iRowLevel)
				{
					if (iRowLevel == m_ppd->iRowLevel)
					{
						if (m_ppd->ppmc)
							pRow = pc.bTreeBottomUp ? m_ppd->ppmc->pRowFrom : m_ppd->ppmc->pRowTo;
						else
							pRow = m_ppd->pEffPaintRow;
					}
					else
						pRow = pParentRow;
					pParentRow = pRow->GetParentRow();

					bDraw = TRUE;
					bLastChild = TRUE;
					while (pRow = m_Rows->GetNextChildRow(pRow, pc.bTreeBottomUp))
					{
						if (!Ctd.TblQueryRowFlags(m_hWndTbl, pRow->GetNr(), ROW_Hidden))
						{
							bLastChild = FALSE;
							break;
						}
					}

					if (bLastChild && iRowLevel < m_ppd->iRowLevel)
						bDraw = FALSE;

					if (bDraw)
					{
						int iDrawLevel = QueryTreeFlags(MTBL_TREE_FLAG_FLAT_STRUCT) ? 0 : iRowLevel;
						if (iDrawLevel > 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE))
							iDrawLevel--;
						//rl.left = pc.rCell.left + m_ppd->tm.m_CellLeadingX + ( iDrawLevel * m_iIndent ) + ( m_iIndent / 2 );
						rl.left = pc.rCell.left + (m_ppd->tm.m_CellLeadingX - 1) + (iDrawLevel * m_iIndent) + (m_iIndent / 2);
						rl.right = rl.left;

						if (bLastChild)
						{
							if (pc.bTreeBottomUp)
							{
								rl.top = pc.rText.top + (txtm.tmHeight / 2);
								rl.bottom = pc.rCell.bottom + 1;
							}
							else
							{
								rl.top = pc.rCell.top;
								rl.bottom = pc.rText.top + (txtm.tmHeight / 2) + 1;
							}
						}
						else
						{
							rl.top = pc.rCell.top;
							rl.bottom = pc.rCell.bottom + 1;
						}

						if (iStyle == PS_SOLID)
						{
							MoveToEx(m_ppd->hDC, rl.left, rl.top, NULL);
							LineTo(m_ppd->hDC, rl.right, rl.bottom);
						}
						else if (iStyle == PS_DOT)
						{
							dl.rCoord = rl;
							dl.bVert = TRUE;
							dl.bOdd = /*iRowMod != 0*/ (pc.dwTopLog + (pc.rCell.top - rl.top)) % 2 != 0;
							LineDDA(rl.left, rl.top, rl.right, rl.bottom, DotlineProc, (LPARAM)&dl);
						}
					}
				}
			}

			if (pc.bVertTreeLineToChild)
			{
				int iDrawLevel = QueryTreeFlags(MTBL_TREE_FLAG_FLAT_STRUCT) ? 0 : m_ppd->iRowLevel + 1;
				if (iDrawLevel > 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE))
					iDrawLevel--;
				//rl.left = pc.rCell.left + m_ppd->tm.m_CellLeadingX + ( iDrawLevel * m_iIndent ) + ( m_iIndent / 2 );
				rl.left = pc.rCell.left + (m_ppd->tm.m_CellLeadingX - 1) + (iDrawLevel * m_iIndent) + (m_iIndent / 2);
				rl.right = rl.left;

				if (pc.bTreeBottomUp)
				{
					rl.top = pc.rCell.top;

					if (pc.bNode && pc.rNode.left <= rl.left)
						rl.bottom = pc.rNode.top;
					else if (pc.Img.GetHandle() && pc.rImg.left <= rl.left)
						rl.bottom = pc.rImg.top;
					else if (pc.rText.left <= rl.left)
						rl.bottom = pc.rText.top;
					else
						rl.bottom = pc.rCell.top + ((pc.rCell.bottom - pc.rCell.top) / 2);
				}
				else
				{
					if (pc.bNode && pc.rNode.left <= rl.left)
						rl.top = pc.rNode.bottom;
					else if (pc.Img.GetHandle() && pc.rImg.left <= rl.left)
						rl.top = pc.rImg.bottom;
					else if (pc.rText.left <= rl.left)
					{
						if (!bTextRectDet)
						{
							rTxt = pc.rText;
							DrawText(m_ppd->hDC, pc.lpsText, pc.lTextLen, &rTxt, pc.dwTextDrawFlags | DT_CALCRECT);
							SIZE si = { rTxt.right - rTxt.left, rTxt.bottom - rTxt.top };
							CalcCellTextRect(pc.rCell.top, pc.rCell.bottom, pc.lTextAreaLeft, pc.lTextAreaRight, m_ppd->tm.m_CellLeadingX, m_ppd->tm.m_CellLeadingY, GetEffCellTextJustify(m_ppd->lRow, m_ppd->pEffPaintCol->hWnd), GetEffCellTextVAlign(m_ppd->lRow, m_ppd->pEffPaintCol->hWnd), si, &rTxt);
							bTextRectDet = TRUE;
						}
						rl.top = rTxt.bottom;
					}
					else
						rl.top = pc.rCell.top + ((pc.rCell.bottom - pc.rCell.top) / 2);

					rl.bottom = pc.rCell.bottom + 1;
				}

				if (iStyle == PS_SOLID)
				{
					MoveToEx(m_ppd->hDC, rl.left, rl.top, NULL);
					LineTo(m_ppd->hDC, rl.right, rl.bottom);
				}
				else if (iStyle == PS_DOT)
				{
					dl.rCoord = rl;
					dl.bVert = TRUE;
					dl.bOdd = /*iRowMod != 0;*/ (pc.dwTopLog + (pc.rCell.top - rl.top)) % 2 != 0;
					LineDDA(rl.left, rl.top, rl.right, rl.bottom, DotlineProc, (LPARAM)&dl);
				}
			}

			SelectObject(m_ppd->hDC, hOldPen);
			DeleteObject(hPen);
		}
	}

	// Text, Bild, Knoten, Buttons

	// Clipping setzen
	RECT rClip;
	rClip.left = pc.rCellVis.left;
	rClip.top = pc.rCellVis.top;
	if (m_ppd->bClipFocusArea && !m_ppd->bColHdr)
	{
		if (pc.rCellVis.top == pc.rCell.top)
			++rClip.top;
		else
			rClip.top = max(pc.rCell.top + 1, pc.rCell.top);
	}
	rClip.right = pc.rCellVis.right;
	rClip.bottom = pc.rCellVis.bottom;
	if (m_ppd->bClipFocusArea && !pc.bGroupColHdrTop)
	{
		if (pc.rCellVis.bottom == pc.rCell.bottom)
			--rClip.bottom;
		else
			rClip.bottom = min(pc.rCell.bottom - 1, pc.rCell.bottom);
	}

	RECT rClipDP;
	CopyRect(&rClipDP, &rClip);
	LPtoDP(m_ppd->hDC, (LPPOINT)&rClipDP, 2);

	rgnClip = CreateRectRgn(rClipDP.left, rClipDP.top, rClipDP.right, rClipDP.bottom);
	if (rgnClip)
	{
		SelectClipRgn(m_ppd->hDC, rgnClip);
		DeleteObject(rgnClip);
	}

	// Bild zeichnen
	if (pc.Img.GetHandle())
	{

		DWORD dwDrawFlags = pc.dwImgDrawFlags;
		if (pc.Img.HasOpacity())
		{
			pc.Img.SetDrawClipRect(&rClip);
			dwDrawFlags |= MIMG_DRAW_USECLIPRECT;
		}
		if ((pc.Bits.Selected && (pc.Bits.NoSelInvImage || pc.Img.QueryFlags(MTSI_NOSELINV)) && !m_ppd->bAnySelColors && !m_bySelDarkening))
			dwDrawFlags |= MIMG_DRAW_INVERTED;
		pc.Img.Draw(m_ppd->hDC, &pc.rImg, dwDrawFlags);
	}

	// Text ausgeben

	// Hintergrundmodus setzen
	int iOldBkMode = SetBkMode(m_ppd->hDC, TRANSPARENT);

	// Textfarbe setzen
	BOOL bInvText = pc.Bits.Selected && pc.Bits.NoSelInvText && !m_ppd->bAnySelColors && !m_bySelDarkening;
	COLORREF crTextColor = (pc.Bits.Selected && m_dwSelTextColor != MTBL_COLOR_UNDEF && pc.Hili.dwTextColor == MTBL_COLOR_UNDEF ? GetDrawColor(m_dwSelTextColor) : pc.crTextColor);
	crTextColor = bInvText ? GetDrawColor(0xffffff - GetNormalColor(crTextColor)) : pc.Bits.Selected ? GetDrawColor(DarkenColor(crTextColor, m_bySelDarkening)) : crTextColor;
	COLORREF crOldTextColor = SetTextColor(m_ppd->hDC, crTextColor);

	// Text
	if (m_ppd->bColHdr && pc.lpsText)
	{
		// Text Spaltenheader
		TCHAR c;
		int iDiff, iDrawPos = 0, iImgWid = 0, iMinLeft, iPos = 0, iTxtHt, iWid;
		RECT r;
		SIZE s;

		CopyRect(&r, &pc.rText);
		//r.top += iCellLeadingY;

		//iWid = pc.rText.right - pc.rText.left + 1;
		iWid = pc.rCell.right - pc.rCell.left + 1;

		BOOL bImgLeft = FALSE, bImgRight = FALSE;
		if (pc.Img.GetHandle())
		{
			BOOL bTile = pc.Img.QueryFlags(MTSI_ALIGN_TILE);
			if (!bTile)
			{
				BOOL bHCenter = pc.Img.QueryFlags(MTSI_ALIGN_HCENTER);
				bImgRight = pc.Img.QueryFlags(MTSI_ALIGN_RIGHT);
				bImgLeft = !(bHCenter || bImgRight);
			}
			iImgWid = pc.rImg.right - pc.rImg.left + 1;
		}

		if (pc.dwTextDrawFlags & DT_CENTER)
		{
			LPTSTR lpsDraw = new TCHAR[pc.lTextLen + 1];

			/*if ( bImgLeft )
			iMinLeft = pc.rImg.right + m_ppd->tm.m_CellLeadingX;
			else
			iMinLeft = pc.rText.left + m_ppd->tm.m_CellLeadingX;

			if ( bImgRight )
			r.right = pc.rImg.left - 1;*/

			iMinLeft = pc.rText.left;

			GetTextExtentPoint32(m_ppd->hDC, _T("A"), 1, &s);
			iTxtHt = s.cy;

			for (;;)
			{
				c = pc.lpsText[iPos];
				lpsDraw[iDrawPos] = c;

				if (c == 13 || c == 0)
				{
					lpsDraw[iDrawPos] = 0;
					// Zeile ausgeben
					s.cx = s.cy = 0;
					GetTextExtentPoint32(m_ppd->hDC, lpsDraw, iDrawPos, &s);
					//r.left = max( iMinLeft, pc.rText.left + ( ( iWid - s.cx ) / 2 ) );
					r.left = max(iMinLeft, pc.rCell.left + ((iWid - s.cx) / 2));
					if ( /*bImgRight &&*/ r.left + s.cx > r.right)
					{
						if ((iDiff = r.left + s.cx - r.right) > 0)
							r.left = max(iMinLeft, r.left - iDiff);
					}
					if (pc.bTextDisabled)
						DrawState(m_ppd->hDC, NULL, NULL, (LPARAM)lpsDraw, 0, r.left, r.top, 0, 0, DST_TEXT | DSS_DISABLED);
					else
						DrawText(m_ppd->hDC, lpsDraw, -1, &r, DT_LEFT | DT_NOPREFIX | DT_SINGLELINE);

					if (c == 0)
					{
						break;
					}
					else
					{
						iPos += 2;
						iDrawPos = 0;
						r.top += iTxtHt;
					}
				}
				else
				{
					++iPos;
					++iDrawPos;
				}
			}
			delete[]lpsDraw;
		}
		else
		{
			/*if ( pc.dwTextDrawFlags & DT_RIGHT )
			{
			if ( bImgLeft )
			r.left = pc.rImg.right + 1;

			if ( bImgRight )
			r.right = pc.rImg.left - m_ppd->tm.m_CellLeadingX;
			else
			r.right -= m_ppd->tm.m_CellLeadingX;
			}
			else
			{
			if ( bImgLeft )
			r.left = pc.rImg.right + m_ppd->tm.m_CellLeadingX;
			else
			r.left += m_ppd->tm.m_CellLeadingX;

			if ( bImgRight )
			r.right = pc.rImg.left - 1;
			}*/

			DrawText(m_ppd->hDC, pc.lpsText, pc.lTextLen, &r, pc.dwTextDrawFlags);
		}
	}
	else if ( /*m_ppd->pEffPaintCol->iCellType*/ pc.GetCellType() == COL_CellType_CheckBox && pc.lpsText && !pc.bTextHidden)
	{
		// Check-Box
		int iJustify = GetEffCellTextJustifyMTbl(pc.pCell, m_ppd->pEffPaintCol->pCol, m_ppd->pEffPaintRow);
		if (iJustify < 0)
			iJustify = m_ppd->pEffPaintCol->lJustify;
		int iVAlign = GetEffCellTextVAlignMTbl(pc.pCell, m_ppd->pEffPaintCol->pCol, m_ppd->pEffPaintRow);
		if (iVAlign < 0)
			iVAlign = DT_TOP;

		RECT rCbArea = { pc.lTextAreaLeft, pc.rCell.top, pc.lTextAreaRight, pc.rCell.bottom };
		RECT r;
		CalcCheckBoxRect(iJustify, iVAlign, rCbArea, m_ppd->tm.m_CellLeadingX, iCellLeadingY, m_ppd->tm.m_CharBoxY, r);

		//BOOL bChecked = m_ppd->pEffPaintCol->dwCBFlags & COL_CheckBox_IgnoreCase ? !lstrcmpi( pc.lpsText, m_ppd->pEffPaintCol->sCBCheckedVal.c_str( ) ) : !lstrcmp( pc.lpsText, m_ppd->pEffPaintCol->sCBCheckedVal.c_str( ) );		
		BOOL bChecked = pc.IsCellTypeCheckedVal(pc.lpsText);

		if (m_ppd->hTheme && !QueryFlags(MTBL_FLAG_IGNORE_WIN_THEME))
		{
			CTheme &th = theTheme();

			int iPart = BP_CHECKBOX;
			int iState = bChecked ? CBS_CHECKEDNORMAL : CBS_UNCHECKEDNORMAL;
			th.DrawThemeBackground(m_ppd->hTheme, m_ppd->hDC, iPart, iState, &r, NULL);
		}
		else
		{
			UINT uState = DFCS_BUTTONCHECK;
			if (bChecked)
				uState |= DFCS_CHECKED;
			DrawFrameControl(m_ppd->hDC, &r, DFC_BUTTON, uState);
		}

		if (bInvText)
			InvertRect_Hooked(m_ppd->hDC, &r);
	}
	else if (pc.lpsText && !pc.bTextHidden)
	{
		// Normaler Text
		DrawText(m_ppd->hDC, pc.lpsText, pc.lTextLen, &pc.rText, pc.dwTextDrawFlags);
	}

	// Alte Textfarbe setzen
	SetTextColor(m_ppd->hDC, crOldTextColor);

	// Alten Hintergrundmodus setzen
	SetBkMode(m_ppd->hDC, iOldBkMode);

	// Alten Font selektieren
	SelectObject(m_ppd->hDC, hOldFont);

	// Knoten zeichnen
	if (pc.bNode)
	{
		PaintNode(pc, &pc.rNode, (pc.Bits.Selected && pc.Bits.NoSelInvNodes && !m_ppd->bAnySelColors && !m_bySelDarkening));
	}

	// Zell-Buttons zeichnen
	if (m_ppd->pEffPaintRow && m_ppd->pEffPaintCol)
	{
		HWND hWndCol = m_ppd->pEffPaintCol->hWnd;
		long lRow = m_ppd->lEffPaintRow;
		// Drop-Down-Button
		int iDDBtnWid = 0;
		if (MustShowDDBtn(hWndCol, lRow, TRUE, TRUE))
		{
			HFONT hBtnFont = pc.hTextFont;

			MTBLPAINTDDBTN b;

			b.hDC = m_ppd->hDC;
			b.bPushed = FALSE;

			RECT rBounds;
			CopyRect(&rBounds, &pc.rCell);
			rBounds.bottom -= 1;

			CalcDDBtnRect(hWndCol, lRow, &m_ppd->tm, &rBounds, m_ppd->hDC, hBtnFont, &b.rBtn);
			iDDBtnWid = b.rBtn.right - b.rBtn.left;

			PaintDDBtn(&b);
			if (pc.Bits.Selected && !m_ppd->bAnySelColors && !m_bySelDarkening)
				InvertRect_Hooked(m_ppd->hDC, &b.rBtn);
		}

		// M!Table-Button
		if (MustShowBtn(hWndCol, lRow, TRUE, TRUE))
		{
			CMTblBtn *pBtn = GetEffCellBtn(hWndCol, lRow);
			if (pBtn)
			{
				BOOL bMustDeleteFont = FALSE;
				HFONT hBtnFont = (pBtn->m_dwFlags & MTBL_BTN_SHARE_FONT) ? GetEffCellFont(m_ppd->hDC, lRow, hWndCol, &bMustDeleteFont) : (HFONT)SendMessage(m_hWndTbl, WM_GETFONT, 0, 0);

				MTBLPAINTBTN b;

				b.hDC = m_ppd->hDC;
				b.hFont = hBtnFont;
				b.hImage = pBtn->m_hImage;
				b.lpsText = pBtn->m_sText.c_str();
				b.dwFlags = pBtn->m_dwFlags;
				b.bPushed = m_ppd->pEffPaintRow->m_Cols->GetBtnPushed(hWndCol);
				b.bDrawInverted = (pc.Bits.Selected && !m_ppd->bAnySelColors && !m_bySelDarkening);
				b.bRestoreClipRgn = TRUE;

				RECT rBounds;
				CopyRect(&rBounds, &pc.rCell);
				rBounds.right -= iDDBtnWid;
				rBounds.bottom -= 1;

				CalcBtnRect(pBtn, hWndCol, lRow, &m_ppd->tm, &rBounds, m_ppd->hDC, hBtnFont, &b.rBtn);

				PaintBtn(&b);

				if (bMustDeleteFont)
					DeleteObject(hBtnFont);
			}
		}
	}

	// Column-Header-Button zeichnen
	if (pc.bRealColHdr)
	{
		CMTblBtn *pBtn = m_Cols->GetHdrBtn(pc.Item.WndHandle);
		if (pBtn && pBtn->m_bShow)
		{
			BOOL bMustDeleteFont = FALSE;
			HFONT hBtnFont = (pBtn->m_dwFlags & MTBL_BTN_SHARE_FONT) ? GetEffColHdrFont(m_ppd->hDC, pc.Item.WndHandle, &bMustDeleteFont) : (HFONT)SendMessage(m_hWndTbl, WM_GETFONT, 0, 0);

			MTBLPAINTBTN b;

			b.hDC = m_ppd->hDC;
			b.hFont = hBtnFont;
			b.hImage = pBtn->m_hImage;
			b.lpsText = pBtn->m_sText.c_str();
			b.dwFlags = pBtn->m_dwFlags;
			b.bPushed = m_Cols->GetHdrBtnPushed(pc.Item.WndHandle);
			b.bDrawInverted = (pc.Bits.Selected && !m_ppd->bAnySelColors && !m_bySelDarkening);
			b.bRestoreClipRgn = TRUE;

			RECT rBounds;
			CopyRect(&rBounds, &pc.rCell);
			rBounds.bottom -= 1;

			CalcBtnRect(pBtn, pc.Item.WndHandle, TBL_Error, &m_ppd->tm, &rBounds, m_ppd->hDC, hBtnFont, &b.rBtn);

			PaintBtn(&b);

			if (bMustDeleteFont)
				DeleteObject(hBtnFont);
		}
	}

	// Row-Flag Bild zeichnen
	if (!pc.Img.GetHandle() &&
		m_ppd->pPaintCol->bRowHdr && !m_ppd->bColHdr && !m_ppd->bEmptyRow &&
		m_ppd->wRowHdrFlags & TBL_RowHdr_MarkEdits &&
		!m_ppd->hWndRowHdrCol &&
		!Ctd.TblQueryRowFlags(m_hWndTbl, m_ppd->lEffPaintRow, ROW_HideMarks)
		)
	{
		DWORD dwFlag;
		DWORD dwFlags[] = { 0, 0, 0, 0, 0, 0, 0, 0 };
		DWORD dwFlagsCheck[] = { ROW_UserFlag1, ROW_UserFlag2, ROW_UserFlag3, ROW_UserFlag4, ROW_UserFlag5, ROW_MarkDeleted, ROW_New, ROW_Edited };
		int i, iFlags = 0, iMax;

		vector<long> vRows;
		vector<long>::iterator it;
		vRows.push_back(m_ppd->lEffPaintRow);

		iMax = sizeof(dwFlagsCheck) / sizeof(DWORD);
		for (i = 0; i < iMax; i++)
		{
			dwFlag = dwFlagsCheck[i];
			for (it = vRows.begin(); it != vRows.end(); it++)
			{
				if (Ctd.TblQueryRowFlags(m_hWndTbl, *it, dwFlag))
				{
					dwFlags[iFlags] = dwFlag;
					iFlags++;
					break;
				}
			}
		}

		for (i = 0; i < iFlags; i++)
		{
			dwFlag = dwFlags[i];
			// Benutzerdefiniertes Bild da?
			HANDLE hImg = GetRowFlagImg(dwFlag, TRUE);
			BOOL bUserImage = (hImg != NULL);
			if (!hImg)
				// OK, dann nehmen wir den Standard
				hImg = GetRowFlagImg(dwFlag, FALSE);

			if (hImg && DWORD(hImg) != MTBL_HIMG_NODEFAULT)
			{
				SIZE s;
				if (MImgGetSize(hImg, &s))
				{
					RECT rSrc = { 0, 0, s.cx, s.cy };
					RECT rDest;
					int iWid = pc.rCell.right - pc.rCell.left;
					int iMaxWid = iWid - 2;
					if (iMaxWid > s.cx)
					{
						rDest.left = pc.rCell.left + 1 + ((iWid - s.cx) / 2);
						rDest.right = rDest.left + s.cx;
					}
					else
					{
						rDest.left = pc.rCell.left + 1;
						rDest.right = rDest.left + iMaxWid;
					}

					int iHt = pc.rCell.bottom - pc.rCell.top;
					int iMaxHt = iHt - 2;
					if (iMaxHt > s.cy)
					{
						rDest.top = pc.rCell.top + 1 + ((iHt - s.cy) / 2);
						rDest.bottom = rDest.top + s.cy;
					}
					else
					{
						rDest.top = pc.rCell.top + 1;
						rDest.bottom = rDest.top + iMaxHt;
					}

					DWORD dwDrawFlags = 0;
					if ((pc.Bits.Selected && QueryRowFlagImgFlags(dwFlag, MTSI_NOSELINV) && !m_ppd->bAnySelColors && !m_bySelDarkening))
						dwDrawFlags |= MIMG_DRAW_INVERTED;
					MImgDraw(hImg, m_ppd->hDC, &rDest, dwDrawFlags);

					break;
				}
			}
		}
	}

	// Benutzerdefiniert gezeichnet?
	if (pc.iODIType)
	{
		// Clipping
		HRGN rgnClip = CreateRectRgnIndirect(&rClipCellDP);
		if (rgnClip)
		{
			SelectClipRgn(m_ppd->hDC, rgnClip);
			DeleteObject(rgnClip);
		}

		// Struktur erzeugen
		m_pODI = new (nothrow)MTBLODI;
		if (m_pODI)
		{
			// Struktur füllen
			m_pODI->lType = pc.iODIType;
			switch (pc.iODIType)
			{
			case MTBL_ODI_CELL:
				m_pODI->hWndParam = m_ppd->pEffPaintCol->hWnd;
				m_pODI->hWndPart = m_ppd->pPaintCol->hWnd;
				break;
			case MTBL_ODI_COLHDR:
				m_pODI->hWndParam = m_ppd->pPaintCol->hWnd;
				m_pODI->hWndPart = m_ppd->pPaintCol->hWnd;
				break;
			case MTBL_ODI_COLHDRGRP:
				m_pODI->hWndParam = NULL;
				m_pODI->hWndPart = m_ppd->pPaintCol->hWnd;
				break;
			default:
				m_pODI->hWndParam = NULL;
				m_pODI->hWndPart = NULL;
			}

			switch (pc.iODIType)
			{
			case MTBL_ODI_CELL:
				m_pODI->lParam = m_ppd->lEffPaintRow;
				m_pODI->lPart = m_ppd->lRow;
				break;
			case MTBL_ODI_ROWHDR:
				m_pODI->lParam = m_ppd->lEffPaintRow;
				m_pODI->lPart = m_ppd->lRow;
				break;
			case MTBL_ODI_COLHDRGRP:
				m_pODI->lParam = (long)pc.pColHdrGrp;
				m_pODI->lPart = (long)pc.pColHdrGrp;
				break;
			default:
				m_pODI->lParam = TBL_Error;
				m_pODI->lPart = TBL_Error;
			}
			m_pODI->hDC = m_ppd->hDC;
			m_pODI->r = pc.rCell;
			m_pODI->lSelected = pc.Bits.Selected;
			m_pODI->lXLeading = m_ppd->tm.m_CellLeadingX;
			m_pODI->lYLeading = m_ppd->tm.m_CellLeadingY;
			m_pODI->lPrinting = 0;
			m_pODI->dYResFactor = 1.0;
			m_pODI->dXResFactor = 1.0;

			// Nachricht senden
			SendMessage(m_hWndTbl, MTM_DrawItem, pc.iODIType, (LPARAM)m_pODI);
		}
	}
	if (m_pODI)
	{
		delete m_pODI;
		m_pODI = NULL;
	}

	// Clipping entfernen
	SelectClipRgn(m_ppd->hDC, NULL);

	return TRUE;
}

// PaintDDBtn
BOOL MTBLSUBCLASS::PaintDDBtn(LPMTBLPAINTDDBTN pb)
{
	HTHEME hTheme = NULL;
	CTheme &th = theTheme();
	if (th.CanUse() && th.IsAppThemed() && !QueryFlags(MTBL_FLAG_IGNORE_WIN_THEME))
	{
		hTheme = th.OpenThemeData(m_hWndTbl, L"Combobox");
		int iPart = CP_DROPDOWNBUTTON;
		int iState = pb->bPushed ? CBXS_PRESSED : CBXS_NORMAL;

		th.DrawThemeBackground(hTheme, pb->hDC, iPart, iState, &pb->rBtn, NULL);
		th.CloseThemeData(hTheme);
	}
	else
	{
		int iStyle = DFCS_SCROLLCOMBOBOX;
		if (QueryFlags(MTBL_FLAG_BUTTONS_FLAT))
			iStyle |= DFCS_FLAT;
		if (pb->bPushed)
			iStyle |= DFCS_PUSHED;
		DrawFrameControl(pb->hDC, &pb->rBtn, DFC_SCROLL, iStyle);
		//DrawFrameControl( pb->hDC, &pb->rBtn, DFC_BUTTON, iStyle );
		//FrameRect( pb->hDC, &pb->rBtn, (HBRUSH)GetStockObject( BLACK_BRUSH ) );
		//DrawEdge( pb->hDC, &pb->rBtn, EDGE_RAISED, BF_RECT | BF_SOFT );
	}

	return TRUE;
}


// PaintFocus
BOOL MTBLSUBCLASS::PaintFocus(LPCRECT pRect)
{
	if (!pRect)
		return FALSE;

	if (m_ppd->bPaintFocus)
	{
		RECT rClip;
		rClip.left = max(pRect->left, m_ppd->rFocusClip.left);
		rClip.top = max(pRect->top, m_ppd->rFocusClip.top);
		rClip.right = min(pRect->right, m_ppd->rFocusClip.right);
		rClip.bottom = min(pRect->bottom, m_ppd->rFocusClip.bottom);
		if (rClip.left <= rClip.right && rClip.top <= rClip.bottom)
		{
			int iRet;

			rClip.right++;
			rClip.bottom++;

			LPtoDP(m_ppd->hDC, (LPPOINT)&rClip, 2);
			HRGN rgnClip = CreateRectRgnIndirect(&rClip);
			if (rgnClip)
			{
				iRet = SelectClipRgn(m_ppd->hDC, rgnClip);
				DeleteObject(rgnClip);
			}

			HRGN rgnFrame = GetFocusRgn(&m_rFocus, GFR_FRAME, m_bCellMode);
			if (rgnFrame)
			{
				if (m_dwFocusFrameColor != MTBL_COLOR_UNDEF)
				{
					HBRUSH hBrush = CreateSolidBrush(GetDrawColor(m_dwFocusFrameColor));
					FillRgn(m_ppd->hDC, rgnFrame, hBrush);
					DeleteObject(hBrush);
				}
				else
					InvertRgn(m_ppd->hDC, rgnFrame);

				DeleteObject(rgnFrame);
			}

			iRet = SelectClipRgn(m_ppd->hDC, NULL);
		}
	}

	return TRUE;
}


// PaintNode
BOOL MTBLSUBCLASS::PaintNode(MTBLPAINTCELL &pc, LPRECT pr, BOOL bInverted /*= FALSE*/)
{

	if (MImgIsValidHandle(pc.hNodeImg))
	{
		HRGN hRgnCurr = CreateRectRgn(0, 0, 0, 0);
		int iRgnCurr = GetClipRgn(m_ppd->hDC, hRgnCurr);
		if (iRgnCurr == 1)
		{
			IntersectClipRect(m_ppd->hDC, pr->left, pr->top, pr->right + 1, pr->bottom + 1);
		}

		DWORD dwFlags = 0;
		if (bInverted)
			dwFlags |= MIMG_DRAW_INVERTED;
		MImgDrawAligned(pc.hNodeImg, m_ppd->hDC, pr, -1, -1, MIMG_ALIGN_HCENTER | MIMG_ALIGN_VCENTER, dwFlags);

		if (iRgnCurr == 1)
			SelectClipRgn(m_ppd->hDC, hRgnCurr);

		if (hRgnCurr)
			DeleteObject(hRgnCurr);
	}
	else
	{
		// Hintergrund
		BOOL bDeleteBrush = FALSE;
		HBRUSH hBrush;
		COLORREF clr = (bInverted ? GetDrawColor(0xffffff - GetNormalColor(pc.crNodeBackColor)) : pc.crNodeBackColor);
		if (clr != m_ppd->clrTblBackColor)
		{
			hBrush = CreateSolidBrush(clr);
			bDeleteBrush = TRUE;
		}
		else
			hBrush = m_ppd->hBrTblBack;

		if (hBrush)
		{
			RECT r;
			CopyRect(&r, pr);
			r.right += 1;
			r.bottom += 1;
			FillRect(m_ppd->hDC, &r, hBrush);

			// Pinsel löschen
			if (bDeleteBrush)
				DeleteObject(hBrush);
		}

		// Umrandung
		// Stift erzeugen
		HGDIOBJ hOldPen;
		clr = (bInverted ? GetDrawColor(0xffffff - GetNormalColor(pc.crNodeColor)) : pc.crNodeColor);
		HPEN hPen = CreatePen(PS_SOLID, 1, clr);
		if (hPen)
		{
			// Stift setzen
			hOldPen = SelectObject(m_ppd->hDC, hPen);

			MoveToEx(m_ppd->hDC, pr->left, pr->top, NULL);
			LineTo(m_ppd->hDC, pr->right, pr->top);
			LineTo(m_ppd->hDC, pr->right, pr->bottom);
			LineTo(m_ppd->hDC, pr->left, pr->bottom);
			LineTo(m_ppd->hDC, pr->left, pr->top);

			// Alten Stift selektieren
			SelectObject(m_ppd->hDC, hOldPen);

			// Stift löschen
			DeleteObject(hPen);
		}

		// "Innenleben"
		// Stift erzeugen
		clr = (bInverted ? GetDrawColor(0xffffff - GetNormalColor(pc.crNodeInnerColor)) : pc.crNodeInnerColor);
		hPen = CreatePen(PS_SOLID, 1, clr);
		if (hPen)
		{
			// Stift setzen
			hOldPen = SelectObject(m_ppd->hDC, hPen);

			// Horizontaler Strich
			MoveToEx(m_ppd->hDC, pr->left + 2, pr->top + (m_iNodeSize / 2), NULL);
			LineTo(m_ppd->hDC, pr->right - 1, pr->top + (m_iNodeSize / 2));

			// Vertikaler Strich
			if (pc.iNodeType == MTBL_NODE_PLUS)
			{
				MoveToEx(m_ppd->hDC, pr->left + (m_iNodeSize / 2), pr->top + 2, NULL);
				LineTo(m_ppd->hDC, pr->left + (m_iNodeSize / 2), pr->bottom - 1);
			}

			// Alten Stift selektieren
			SelectObject(m_ppd->hDC, hOldPen);

			// Stift löschen
			DeleteObject(hPen);
		}
	}

	return TRUE;
}

// PaintRow
BOOL MTBLSUBCLASS::PaintRow()
{
	// Zeile zeichnen?
	BOOL bPaint = !(m_ppd->dwRowTop >= m_ppd->rUpd.bottom || m_ppd->dwRowBottom < m_ppd->rUpd.top);

	if (bPaint)
	{
		BOOL bColHdrGrpLockBreak = FALSE, bDeleteFont, bFirstGroupColHdr, bNotLastGroupColHdr, bPaintCell, bPseudoRow;
		CMTblColHdrGrp *pColHdrGrp;
		DWORD dwFlags;
		int iFirstPaintCol = -1, iGroupColHdrHt, iJustify, iVAlign;
		MTBLHILI hili;
		MTBLPAINTCELL &pc = m_ppd->pc;
		LPMTBLPAINTCOL pLastCol;
		RECT rCellOrig, rCellVisOrig;
		tstring sColHdrText, sGroupColHdrText;

		// Pseudo-Zeile?
		bPseudoRow = (m_ppd->bColHdr || m_ppd->bEmptyRow);

		// Kontext setzen ( und bei Bedarf fetchen )
		if (!bPseudoRow)
			Ctd.TblSetContextEx(m_hWndTbl, m_ppd->lRow);

		// Zeile setzen
		if (bPseudoRow)
			m_ppd->pRow = NULL;
		else
			m_ppd->pRow = m_Rows->EnsureValid(m_ppd->lRow);

		// Merge-Zeile setzen
		if (bPseudoRow)
			m_ppd->pmr = NULL;
		else
			m_ppd->pmr = FindMergeRow(m_ppd->lRow, FMR_ANY);

		// Spalten
		CMTblCol *pCell = NULL;
		m_ppd->ppmc = NULL;

		m_ppd->pFirstGroupColHdr = m_ppd->pLastGroupColHdr = NULL;
		pColHdrGrp = NULL;

		for (int i = 0; i < m_ppd->iCols; ++i)
		{
			m_ppd->pPaintCol = m_ppd->PaintCols[i];
			pCell = m_ppd->pRow ? m_ppd->pRow->m_Cols->Find(m_ppd->pPaintCol->hWnd) : NULL;
			bPaintCell = m_ppd->pPaintCol->bPaint;

			ZeroMemory(&pc.Bits, sizeof(pc.Bits));

			// Merge-Daten ermitteln
			m_ppd->ppmc = NULL;
			m_ppd->ppmrh = NULL;
			m_ppd->pFirstMergeCol = m_ppd->pLastMergeCol = NULL;
			m_ppd->pFirstMergeRow = m_ppd->pLastMergeRow = NULL;

			// Ist das möglicherweise Teil einer Merge-Zelle?
			if (bPaintCell && !m_ppd->pPaintCol->bPseudo && m_ppd->pRow && m_pCounter->GetMergeCells() > 0)
			{
				if (m_ppd->ppmc = FindPaintMergeCell(m_ppd->pPaintCol->hWnd, m_ppd->lRow, FMC_ANY))
				{
					m_ppd->pFirstMergeCol = m_ppd->ppmc->pColFrom;
					m_ppd->pLastMergeCol = m_ppd->ppmc->pColTo;
					m_ppd->pFirstMergeRow = m_ppd->ppmc->pRowFrom;
					m_ppd->pLastMergeRow = m_ppd->ppmc->pRowTo;
				}
			}

			// Ist das möglicherweise Teil eines Merge-Rowheaders?
			if (bPaintCell && m_ppd->pPaintCol->bRowHdr && m_ppd->pRow && m_pCounter->GetMergeRows() > 0)
			{
				if (m_ppd->ppmrh = FindPaintMergeRowHdr(m_ppd->lRow, FMC_ANY))
				{
					m_ppd->pFirstMergeRow = m_ppd->ppmrh->pRowFrom;
					m_ppd->pLastMergeRow = m_ppd->ppmrh->pRowTo;
				}
			}

			// "Effektive" PaintCol setzen
			m_ppd->pEffPaintCol = (m_ppd->pFirstMergeCol ? m_ppd->pFirstMergeCol : m_ppd->pPaintCol);

			// "Effektive" PaintRow setzen
			m_ppd->pEffPaintRow = (m_ppd->pFirstMergeRow ? m_ppd->pFirstMergeRow : m_ppd->pRow);
			m_ppd->lEffPaintRow = m_ppd->pEffPaintRow ? m_ppd->pEffPaintRow->GetNr() : m_ppd->lRow;

			// "Effektive" Zelle setzen
			CMTblCol *pEffCell = m_ppd->pEffPaintRow ? m_ppd->pEffPaintRow->m_Cols->Find(m_ppd->pEffPaintCol->hWnd) : NULL;
			pc.pCell = pEffCell;

			// Zellentyp setzen
			if (!m_ppd->pEffPaintCol->bPseudo && !bPseudoRow)
				pc.pCellType = GetEffCellTypePtr(m_ppd->pEffPaintRow->GetNr(), m_ppd->pEffPaintCol->hWnd);
			else
				pc.pCellType = NULL;

			// Zeile markiert?
			if (bPseudoRow)
				m_ppd->bRowSelected = FALSE;
			else
			{
				m_ppd->bRowSelected = Ctd.TblQueryRowFlags(m_hWndTbl, m_ppd->lRow, ROW_Selected);

				if (Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SingleSelection))
				{
					LPMTBLMERGEROW pmr = FindMergeRow(m_ppd->lRow, FMR_ANY);
					if (pmr)
						m_ppd->bRowSelected = AnyMergeRowSelected(pmr);
				}
			}

			// Daten Gruppen-Spaltenheader ermitteln
			pc.pColHdrGrp = NULL;
			bFirstGroupColHdr = FALSE;
			bNotLastGroupColHdr = FALSE;

			if (m_ppd->bColHdr)
			{
				// Spaltenheader-Gruppe splitten, wenn das die erste nicht verankerte Spalte ist
				if (i > 0 && m_ppd->PaintCols[i - 1]->bLastLocked)
					pColHdrGrp = NULL;
				// Ist das ein gruppierter Spalten-Header?
				if (m_ppd->pPaintCol->pHdrGrp && m_ppd->pPaintCol->pHdrGrp != pColHdrGrp)
				{
					// Aktuelle Spaltengruppe setzen
					pColHdrGrp = m_ppd->pPaintCol->pHdrGrp;
					// Erste Spalte der Gruppe festlegen
					m_ppd->pFirstGroupColHdr = m_ppd->pPaintCol;
					// Letzte Spalte der Gruppe festlegen
					m_ppd->pLastGroupColHdr = m_ppd->pPaintCol;
					for (int j = i + 1; j < m_ppd->iCols; ++j)
					{
						if (m_ppd->pPaintCol->bLocked && !m_ppd->PaintCols[j]->bLocked)
							break;
						if (m_ppd->PaintCols[j]->pHdrGrp != pColHdrGrp)
							break;
						m_ppd->pLastGroupColHdr = m_ppd->PaintCols[j];
					}
				}
				else if (!m_ppd->pPaintCol->pHdrGrp)
				{
					pColHdrGrp = NULL;
					m_ppd->pFirstGroupColHdr = m_ppd->pLastGroupColHdr = NULL;
				}

				pc.pColHdrGrp = pColHdrGrp;
			}

			// Gruppierter Spalten-Header, und zwar der erste?
			if (pColHdrGrp && m_ppd->pPaintCol == m_ppd->pFirstGroupColHdr)
				bFirstGroupColHdr = TRUE;

			// Gruppierter Spalten-Header, aber nicht der letzte?
			if (pColHdrGrp && m_ppd->pPaintCol != m_ppd->pLastGroupColHdr)
				bNotLastGroupColHdr = TRUE;

			if (!bPaintCell)
			{
				if (iFirstPaintCol >= 0)
					break;
				else
					continue;
			}

			// Y-Leading
			if (pColHdrGrp)
				pc.iCellLeadingY = m_ppd->tm.m_CellLeadingY / 3;
			else
				pc.iCellLeadingY = m_ppd->tm.m_CellLeadingY;

			// Items definieren
			if (m_ppd->pPaintCol->bRowHdr && m_ppd->bColHdr)
			{
				pc.Item.DefineAsCorner(m_hWndTbl);
				pc.HorzLineItem = pc.VertLineItem = pc.Item;
			}
				
			else if (m_ppd->pPaintCol->bRowHdr)
			{
				pc.Item.DefineAsRowHdr(m_hWndTbl, m_ppd->pEffPaintRow);
				pc.HorzLineItem.DefineAsRowHdr(m_hWndTbl, m_ppd->pRow);
				pc.VertLineItem = pc.HorzLineItem;
			}
			else if (m_ppd->bColHdr)
			{
				pc.Item.DefineAsColHdr(m_ppd->pEffPaintCol->pCol ? m_ppd->pEffPaintCol->pCol->m_hWnd : NULL);
				pc.HorzLineItem.DefineAsColHdr(m_ppd->pPaintCol->pCol ? m_ppd->pPaintCol->pCol->m_hWnd : NULL);
				pc.VertLineItem = pc.HorzLineItem;
			}				
			else
			{
				pc.Item.DefineAsCell(m_ppd->pEffPaintCol->pCol ? m_ppd->pEffPaintCol->pCol->m_hWnd : NULL, m_ppd->pEffPaintRow);
				pc.VertLineItem = pc.HorzLineItem = pc.Item;
			}				

			// Identifikations-Booleans setzen
			pc.bCell = pc.Item.IsCell();
			pc.bRealCell = pc.Item.IsRealCell();
			pc.bColHdr = pc.Item.IsColHdr();
			pc.bRealColHdr = pc.Item.IsRealColHdr();
			pc.bRowHdr = pc.Item.IsRowHdr();
			pc.bRealRowHdr = pc.Item.IsRealRowHdr();
			pc.bCorner = pc.Item.IsCorner();

			pc.bColHdrGrp = FALSE;

			// Effektives Highlighting ermitteln
			pc.Hili.Init();
			BOOL bAnyHighlightedItems = AnyHighlightedItems();
			if (bAnyHighlightedItems)
			{
				// Hier wird bewusst NICHT die effektive Zeile/Spalte verwendet
				HWND hWndHili = m_ppd->pPaintCol->bPseudo ? m_hWndTbl : m_ppd->pPaintCol->pCol->m_hWnd;
				CMTblRow *pHiliRow = m_ppd->pRow;

				CMTblItem item = pc.Item;
				item.WndHandle = hWndHili;
				item.Id = pHiliRow;
				GetEffItemHighlighting(item, pc.Hili);
			}

			// Kontext setzen
			if (!bPseudoRow)
				Ctd.TblSetContextOnly(m_hWndTbl, m_ppd->lEffPaintRow);

			// Zeilenebene setzen
			if (bPseudoRow)
				m_ppd->iRowLevel = 0;
			else
				m_ppd->iRowLevel = m_ppd->pEffPaintRow->GetLevel();

			// Spiegeln einer Spalte im Row-Header?
			m_ppd->bRowHdrMirror = (m_ppd->pPaintCol->bRowHdr && !bPseudoRow && m_ppd->hWndRowHdrCol);

			// Benutzerdefiniert gezeichnet?
			pc.iODIType = 0;
			if (pc.bRealCell /*!( m_ppd->bColHdr || m_ppd->bEmptyRow || m_ppd->pPaintCol->bPseudo )*/)
			{
				if ((pEffCell && pEffCell->IsOwnerdrawn()) ||
					(m_ppd->pEffPaintCol->pCol && m_ppd->pEffPaintCol->pCol->IsOwnerdrawn()) ||
					(m_ppd->pEffPaintRow && m_ppd->pEffPaintRow->IsOwnerdrawn())
					)
					pc.iODIType = MTBL_ODI_CELL;
			}
			else if (pc.bRealColHdr /*m_ppd->bColHdr && !m_ppd->pPaintCol->bPseudo*/)
			{
				CMTblCol *pColHdr = m_Cols->FindHdr(m_ppd->pPaintCol->hWnd);
				if ((pColHdr && pColHdr->IsOwnerdrawn()) ||
					(m_ppd->pPaintCol->pCol && m_ppd->pPaintCol->pCol->IsOwnerdrawn())
					)
					pc.iODIType = MTBL_ODI_COLHDR;
			}
			else if (pc.bRealRowHdr /*m_ppd->pPaintCol->bRowHdr && !m_ppd->bColHdr && !bPseudoRow*/)
			{
				if (m_ppd->pEffPaintRow && (m_ppd->pEffPaintRow->IsHdrOwnerdrawn() || m_ppd->pEffPaintRow->IsOwnerdrawn()))
				{
					pc.iODIType = MTBL_ODI_ROWHDR;
				}
			}
			else if (pc.bCorner /*m_ppd->pPaintCol->bRowHdr && m_ppd->bColHdr*/)
			{
				if (m_bCornerOwnerdrawn)
					pc.iODIType = MTBL_ODI_CORNER;
			}

			// X-Koordinaten setzen
			if (m_ppd->ppmc)
			{
				pc.rCell.left = m_ppd->pFirstMergeCol->ptX.x;
				pc.rCell.right = m_ppd->pLastMergeCol->ptX.y;
			}
			else
			{
				pc.rCell.left = m_ppd->pPaintCol->ptX.x;
				pc.rCell.right = m_ppd->pPaintCol->ptX.y;
			}

			// Sichtbare X-Koordinaten setzen
			if (m_ppd->ppmc && m_ppd->pPaintCol != m_ppd->ppmc->pColTo)
			{
				pc.rCellVis.left = m_ppd->pPaintCol->ptVisX.x;
				pc.rCellVis.right = m_ppd->pPaintCol->ptVisX.y + 1;
			}
			else
			{
				pc.rCellVis.left = m_ppd->pPaintCol->ptVisX.x;
				pc.rCellVis.right = m_ppd->pPaintCol->ptVisX.y;
			}

			// Y-Koordinaten setzen
			if (m_ppd->ppmc)
			{
				pc.rCell.top = m_ppd->ppmc->r.top;
				pc.rCell.bottom = m_ppd->ppmc->r.bottom;
			}
			else if (m_ppd->ppmrh)
			{
				pc.rCell.top = m_ppd->ppmrh->r.top;
				pc.rCell.bottom = m_ppd->ppmrh->r.bottom;
			}
			else
			{
				pc.rCell.top = m_ppd->dwRowTop;
				pc.rCell.bottom = m_ppd->dwRowBottom;
			}

			// Sichtbare Y-Koordinaten setzen
			if (m_ppd->ppmc && m_ppd->pRow != m_ppd->ppmc->pRowTo)
			{
				pc.rCellVis.top = m_ppd->dwRowTopVis;
				pc.rCellVis.bottom = m_ppd->dwRowBottomVis + 1;
			}
			else if (m_ppd->ppmrh && m_ppd->pRow != m_ppd->ppmrh->pRowTo)
			{
				pc.rCellVis.top = m_ppd->dwRowTopVis;
				pc.rCellVis.bottom = m_ppd->dwRowBottomVis + 1;
			}
			else
			{
				pc.rCellVis.top = m_ppd->dwRowTopVis;
				pc.rCellVis.bottom = m_ppd->dwRowBottomVis;
			}

			// Evtl. Y-Koordinaten anpassen
			if (pColHdrGrp)
			{
				CopyRect(&rCellOrig, &pc.rCell);
				CopyRect(&rCellVisOrig, &pc.rCellVis);

				iGroupColHdrHt = GetColHdrGrpHeight(pColHdrGrp);
				if (iGroupColHdrHt > 0)
				{
					if (pc.rCell.bottom > iGroupColHdrHt)
					{
						pc.rCell.top += iGroupColHdrHt + 1;
						pc.rCellVis.top = max(pc.rCell.top, pc.rCellVis.top);
					}
				}
			}

			// Linien
			// Spaltenlinie
			if (pc.bRowHdr)
				pc.bColLine = TRUE;
			else if (pc.bColHdr)
				if (pColHdrGrp && bNotLastGroupColHdr)
					pc.bColLine = !pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOINNERVERTLINES);
				else
					pc.bColLine = !(m_ppd->pPaintCol->bLastVisible && m_ppd->bSuppressLastColLine);
			else
			{
				pLastCol = (m_ppd->ppmc ? m_ppd->pLastMergeCol : m_ppd->pPaintCol);
				pc.bColLine = !pLastCol->bEmpty && !(m_ppd->bSuppressLastColLine && pLastCol->bLastVisible) /*&& !((pLastCol->bLastLocked ? m_pLastLockedColVertLineDef->Style : m_pColsVertLineDef->Style) == MTLS_INVISIBLE)*/;
				if (pc.bColLine)
				{
					if ((m_ppd->bEmptyRow) && QueryFlags(MTBL_FLAG_NO_FREE_ROW_AREA_LINES))
						pc.bColLine = FALSE;
					else if (m_ppd->pPaintCol->pCol && (m_ppd->pPaintCol->pCol->QueryFlags(MTBL_COL_FLAG_NO_COLLINE) || (m_ppd->pPaintCol->pCol->m_pVertLineDef ? m_ppd->pPaintCol->pCol->m_pVertLineDef->Style == MTLS_INVISIBLE : FALSE)))
						pc.bColLine = FALSE;
					else if (m_ppd->pRow && m_ppd->pRow->QueryFlags(MTBL_ROW_NO_COLLINES))
						pc.bColLine = FALSE;
					else if (pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_NO_COLLINE))
						pc.bColLine = FALSE;
				}
			}

			// Zeilenlinie
			if (pc.bColHdr)
				pc.bRowLine = TRUE;
			else if (pc.bRowHdr)
				pc.bRowLine = !m_ppd->bNoRowLines;
			else if ((m_dwTreeFlags & MTBL_TREE_FLAG_NO_ROWLINES) && (m_hWndTreeCol && m_ppd->pPaintCol->hWnd == m_hWndTreeCol))
				pc.bRowLine = FALSE;
			else
			{
				pc.bRowLine = !m_ppd->bNoRowLines && !(m_pRowsHorzLineDef->Style == MTLS_INVISIBLE);
				if (pc.bRowLine)
				{
					if ((m_ppd->bEmptyRow) && QueryFlags(MTBL_FLAG_NO_FREE_ROW_AREA_LINES))
						pc.bRowLine = FALSE;
					else if ((m_ppd->pPaintCol->bEmpty) && QueryFlags(MTBL_FLAG_NO_FREE_COL_AREA_LINES))
						pc.bRowLine = FALSE;
					else if (m_ppd->pPaintCol->pCol && m_ppd->pPaintCol->pCol->QueryFlags(MTBL_COL_FLAG_NO_ROWLINES))
						pc.bRowLine = FALSE;
					else if (m_ppd->pRow && m_ppd->pRow->QueryFlags(MTBL_ROW_NO_ROWLINE))
						pc.bRowLine = FALSE;
					else if (pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_NO_ROWLINE))
						pc.bRowLine = FALSE;
				}
			}

			// Hintergrundfarbe ermitteln
			pc.crBackColor = m_ppd->clrTblBackColor;

			if (pc.bCell)
			{
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				if (!pc.bRealCell)
				{
					if (m_ppd->bEmptyRow)
						dwFlags |= GEC_EMPTY_ROW;
					if (m_ppd->pPaintCol->bEmpty)
						dwFlags |= GEC_EMPTY_COL;
				}

				pc.crBackColor = GetEffCellBackColorMTbl(pEffCell, m_ppd->pEffPaintCol->pCol, m_ppd->pEffPaintRow, dwFlags);
				if (pc.crBackColor == MTBL_COLOR_UNDEF)
					pc.crBackColor = m_ppd->clrTblBackColor;
				else
					pc.crBackColor = GetDrawColor(pc.crBackColor);

				// Hintergrund bei Selektion nicht invertieren?
				if ((pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_NOSELINV_BKGND)) ||
					(m_ppd->pEffPaintCol->pCol && m_ppd->pEffPaintCol->pCol->QueryFlags(MTBL_COL_FLAG_NOSELINV_BKGND)) ||
					(m_ppd->pEffPaintRow && m_ppd->pEffPaintRow->QueryFlags(MTBL_ROW_NOSELINV_BKGND))
					)
				{
					pc.Bits.NoSelInvBkgnd = 1;
				}
			}
			else if (pc.bRowHdr)
			{
				if (pc.bRealRowHdr)
				{
					dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
					pc.crBackColor = GetDrawColor(GetEffRowHdrBackColor(m_ppd->pEffPaintRow->GetNr(), dwFlags));
				}
				else
					pc.crBackColor = GetDrawColor(GetEffRowHdrBackColor(TBL_Error));
			}
			else if (pc.bColHdr)
			{
				if (pc.bRealColHdr)
				{
					dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
					pc.crBackColor = GetEffColHdrBackColor(m_ppd->pEffPaintCol->hWnd, dwFlags);
				}
				else
					pc.crBackColor = GetEffColHdrFreeBackColor();
			}
			else if (pc.bCorner)
			{
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				pc.crBackColor = GetDrawColor(GetEffCornerBackColor(dwFlags));
			}


			// Textfarbe ermitteln
			pc.crTextColor = m_ppd->clrTblTextColor;
			if (pc.bRealCell)
			{
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				if (m_hWndHLCol == m_ppd->pEffPaintCol->hWnd && m_lHLRow == m_ppd->lEffPaintRow)
					dwFlags |= GEC_HYPERLINK;

				pc.crTextColor = GetEffCellTextColorMTbl(pEffCell, m_ppd->pEffPaintCol->pCol, m_ppd->pEffPaintRow, dwFlags);
				if (pc.crTextColor == MTBL_COLOR_UNDEF)
				{
					if (pEffCell)
						pc.crTextColor = pEffCell->m_TextColorSal.Get();

					if (pc.crTextColor == MTBL_COLOR_UNDEF)
						pc.crTextColor = m_ppd->clrTblTextColor;
					else
						pc.crTextColor = GetDrawColor(pc.crTextColor);
				}
				else
					pc.crTextColor = GetDrawColor(pc.crTextColor);

				// Text bei Selektion nicht invertieren?
				if ((pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_NOSELINV_TEXT)) ||
					(m_ppd->pEffPaintCol->pCol && m_ppd->pEffPaintCol->pCol->QueryFlags(MTBL_COL_FLAG_NOSELINV_TEXT)) ||
					(m_ppd->pEffPaintRow && m_ppd->pEffPaintRow->QueryFlags(MTBL_ROW_NOSELINV_TEXT))
					)
				{
					pc.Bits.NoSelInvText = 1;
				}
			}
			else if (pc.bRealRowHdr)
			{
				if (bAnyHighlightedItems && pc.Hili.dwTextColor != MTBL_COLOR_UNDEF)
					pc.crTextColor = GetDrawColor(pc.Hili.dwTextColor);
				else
				{
					if (m_ppd->hWndRowHdrCol && m_ppd->wRowHdrFlags & TBL_RowHdr_ShareColor)
					{
						dwFlags = bAnyHighlightedItems ? GEC_HILI : 0;
						pc.crTextColor = GetDrawColor(GetEffCellTextColor(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol, dwFlags));
					}
					else
						pc.crTextColor = m_ppd->clrTblTextColor;
				}

				// Text bei Selektion nicht invertieren?
				if ((m_ppd->hWndRowHdrCol && QueryCellFlags(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol, MTBL_CELL_FLAG_NOSELINV_TEXT)) ||
					(m_ppd->hWndRowHdrCol && QueryColumnFlags(m_ppd->hWndRowHdrCol, MTBL_COL_FLAG_NOSELINV_TEXT)) ||
					(m_ppd->pEffPaintRow && m_ppd->pEffPaintRow->QueryFlags(MTBL_ROW_NOSELINV_TEXT))
					)
				{
					pc.Bits.NoSelInvText = 1;
				}
			}
			else if (pc.bRealColHdr)
			{
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				pc.crTextColor = GetDrawColor(GetEffColHdrTextColor(m_ppd->pPaintCol->hWnd, dwFlags));
			}
			else if (pc.bCorner)
			{
				if (bAnyHighlightedItems && pc.Hili.dwTextColor != MTBL_COLOR_UNDEF)
					pc.crTextColor = GetDrawColor(pc.Hili.dwTextColor);
				else
					pc.crTextColor = m_ppd->clrTblTextColor;
			}


			// Knoten ermitteln
			pc.bNode = FALSE;

			BOOL bNode = FALSE;
			if (pc.bRealCell)
				bNode = m_ppd->pEffPaintCol->hWnd == m_hWndTreeCol;
			else if (pc.bRealRowHdr && m_ppd->bRowHdrMirror)
				bNode = m_ppd->hWndRowHdrCol == m_hWndTreeCol;

			if (bNode)
			{
				if (m_iNodeSize > 0 && m_Rows->QueryRowFlags(m_ppd->lEffPaintRow, MTBL_ROW_CANEXPAND) && !(m_Rows->GetRowLevel(m_ppd->lEffPaintRow) == 0 && QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE)))
				{
					// Knoten zeichnen
					pc.bNode = TRUE;

					// Hervorhebung ermitteln
					CMTblItem item;
					if (m_ppd->bRowHdrMirror)
						item.DefineAsRowHdrNode(m_hWndTbl, m_ppd->pEffPaintRow);
					else
						item.DefineAsCellNode(m_hWndTreeCol, m_ppd->pEffPaintRow);
					GetEffItemHighlighting(item, hili);

					// Knoten-Typ setzen
					pc.iNodeType = m_Rows->QueryRowFlags(m_ppd->lEffPaintRow, MTBL_ROW_ISEXPANDED) ? MTBL_NODE_MINUS : MTBL_NODE_PLUS;

					// Knotenfarben setzen
					pc.crNodeColor = (m_dwNodeColor == MTBL_COLOR_UNDEF ? GetDrawColor(RGB(0, 0, 0)) : GetDrawColor(m_dwNodeColor));
					pc.crNodeInnerColor = (m_dwNodeInnerColor == MTBL_COLOR_UNDEF ? pc.crNodeColor : GetDrawColor(m_dwNodeInnerColor));
					if (hili.dwBackColor != MTBL_COLOR_UNDEF)
						pc.crNodeBackColor = GetDrawColor(hili.dwBackColor);
					else
						pc.crNodeBackColor = (m_dwNodeBackColor == MTBL_COLOR_UNDEF ? pc.crBackColor : GetDrawColor(m_dwNodeBackColor));

					// Knotenbild setzen
					pc.hNodeImg = NULL;
					if (pc.iNodeType == MTBL_NODE_PLUS)
					{
						if (HANDLE hImg = hili.Img.GetHandle())
							pc.hNodeImg = hImg;
						else
							pc.hNodeImg = m_hImgNode;
					}
					else
					{
						if (HANDLE hImg = hili.ImgExp.GetHandle())
							pc.hNodeImg = hImg;
						else
							pc.hNodeImg = m_hImgNodeExp;
					}

					// Knoten bei Selektion nicht invertieren?
					if (QueryTreeFlags(MTBL_TREE_FLAG_NOSELINV_NODES))
						pc.Bits.NoSelInvNodes = 1;
				}

				// Tree-Linien bei Selektion nicht invertieren?
				if (QueryTreeFlags(MTBL_TREE_FLAG_NOSELINV_LINES))
					pc.Bits.NoSelInvTreeLines = 1;
			}

			// Bild ermitteln
			pc.Img.SetHandle(NULL, NULL);
			pc.dwImgDrawFlags = 0;

			if (pc.bRealCell)
			{
				// Bild ermitteln
				dwFlags = bAnyHighlightedItems ? GEI_HILI_PD : 0;
				pc.Img = GetEffCellImage(m_ppd->lEffPaintRow, m_ppd->pEffPaintCol->hWnd, dwFlags);

				// Bild bei Selektion nicht invertieren?
				if ((pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_NOSELINV_IMAGE)) ||
					(m_ppd->pEffPaintCol->pCol && m_ppd->pEffPaintCol->pCol->QueryFlags(MTBL_COL_FLAG_NOSELINV_IMAGE)) ||
					(m_ppd->pEffPaintRow && m_ppd->pEffPaintRow->QueryFlags(MTBL_ROW_NOSELINV_IMAGE))
					)
				{
					pc.Bits.NoSelInvImage = 1;
				}
			}
			else if (pc.bRealRowHdr)
			{
				// Gespiegelt?
				if (m_ppd->hWndRowHdrCol)
				{
					// Bild der gespiegelten Zelle ermitteln
					dwFlags = bAnyHighlightedItems ? GEI_HILI : 0;
					pc.Img = GetEffCellImage(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol, dwFlags);

					// Bild bei Selektion nicht invertieren?
					if ((QueryCellFlags(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol, MTBL_CELL_FLAG_NOSELINV_IMAGE)) ||
						(QueryColumnFlags(m_ppd->hWndRowHdrCol, MTBL_COL_FLAG_NOSELINV_IMAGE)) ||
						(m_ppd->pEffPaintRow && m_ppd->pEffPaintRow->QueryFlags(MTBL_ROW_NOSELINV_IMAGE))
						)
					{
						pc.Bits.NoSelInvImage = 1;
					}
				}

				// Highlighting-Bild ermitteln
				if (bAnyHighlightedItems)
				{
					BOOL bExpanded = m_ppd->pEffPaintRow->QueryFlags(MTBL_ROW_ISEXPANDED);
					CMTblImg &ImgHili = bExpanded ? pc.Hili.ImgExp : pc.Hili.Img;

					if (HANDLE hImg = ImgHili.GetHandle())
					{
						if (!pc.Img.GetHandle())
							pc.Img.SetFlags(MTSI_ALIGN_HCENTER | MTSI_ALIGN_VCENTER, TRUE);

						pc.Img.SetHandle(hImg, NULL);
					}
				}
			}
			else if (pc.bRealColHdr)
			{
				// Bild ermitteln
				dwFlags = bAnyHighlightedItems ? GEI_HILI_PD : 0;
				pc.Img = GetEffColHdrImage(m_ppd->pPaintCol->hWnd, dwFlags);
			}
			else if (pc.bCorner)
			{
				// Bild ermitteln
				dwFlags = bAnyHighlightedItems ? GEI_HILI_PD : 0;
				pc.Img = GetEffCornerImage(dwFlags);
			}


			// Font ermitteln
			bDeleteFont = FALSE;
			pc.hTextFont = m_hFont;

			if (pc.bRealCell)
			{
				dwFlags = bAnyHighlightedItems ? GEI_HILI_PD : 0;
				if (m_hWndHLCol == m_ppd->pEffPaintCol->hWnd && m_lHLRow == m_ppd->lEffPaintRow)
					dwFlags |= GEF_HYPERLINK;

				pc.hTextFont = GetEffCellFont(m_ppd->hDC, m_ppd->lEffPaintRow, m_ppd->pEffPaintCol->hWnd, &bDeleteFont, dwFlags);
			}
			else if (pc.bRealRowHdr)
			{
				if (bAnyHighlightedItems && pc.Hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
				{
					CMTblFont f;
					f.Set(m_ppd->hDC, m_hFont);
					f.AddEnh(pc.Hili.lFontEnh);
					pc.hTextFont = f.Create(m_ppd->hDC);
					bDeleteFont = TRUE;
				}
			}
			else if (pc.bRealColHdr)
			{
				dwFlags = bAnyHighlightedItems ? GEI_HILI_PD : 0;
				pc.hTextFont = GetEffColHdrFont(m_ppd->hDC, m_ppd->pEffPaintCol->hWnd, &bDeleteFont, dwFlags);
			}
			else if (pc.bCorner)
			{
				if (bAnyHighlightedItems && pc.Hili.lFontEnh != MTBL_FONT_UNDEF_ENH)
				{
					CMTblFont f;
					f.Set(m_ppd->hDC, m_hFont);
					f.AddEnh(pc.Hili.lFontEnh);
					pc.hTextFont = f.Create(m_ppd->hDC);
					bDeleteFont = TRUE;
				}
			}

			// Text, Textausgabe-Flags und Textausgabe-Rechteck ermitteln
			BOOL bVarRowHtEnabled = QueryFlags(MTBL_FLAG_VARIABLE_ROW_HEIGHT) > 0;

			pc.hsText = 0;
			pc.lpsText = NULL;
			pc.lTextLen = 0;
			pc.bTextHidden = FALSE;
			pc.bTextDisabled = FALSE;

			pc.dwTextDrawFlags = 0;

			if (pc.bRealCell)
			{
				// Text ermitteln
				BOOL bUnformatted = pc.GetCellType() == COL_CellType_CheckBox;
				if (bUnformatted)
					Ctd.FmtFieldToStr(m_ppd->pEffPaintCol->hWnd, &pc.hsText, FALSE);
				else
					Ctd.TblGetColumnText(m_hWndTbl, m_ppd->pEffPaintCol->iID, &pc.hsText);

				if (pc.hsText)
				{
					pc.lpsText = Ctd.StringGetBuffer(pc.hsText, &pc.lTextLen);
					pc.lTextLen = Ctd.StrLenFromBufLen(pc.lTextLen);
				}
				

				if (pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_HIDDEN_TEXT))
					pc.bTextHidden = TRUE;

				if (pc.GetCellType() != COL_CellType_CheckBox && m_ppd->pEffPaintCol->bInvisible && pc.lpsText)
				{
					for (int i = 0; i < pc.lTextLen; ++i)
					{
						pc.lpsText[i] = m_cPassword;
					}
				}

				// Textausgabe-Flags ermitteln
				iJustify = GetEffCellTextJustifyMTbl(pEffCell, m_ppd->pEffPaintCol->pCol, m_ppd->pEffPaintRow);
				if (iJustify < 0)
					iJustify = m_ppd->pEffPaintCol->lJustify;
				pc.dwTextDrawFlags = DT_NOPREFIX | iJustify;

				if (m_ppd->pEffPaintCol->bWordWrap && (m_ppd->tm.m_LinesPerRow > 1 || bVarRowHtEnabled || (m_ppd->ppmc && m_ppd->ppmc->iMergedRows > 0)))
					pc.dwTextDrawFlags |= DT_WORDBREAK;
				else if (!QueryFlags(MTBL_FLAG_EXPAND_CRLF))
					pc.dwTextDrawFlags |= DT_SINGLELINE;

				if (MustUseEndEllipsis(m_ppd->pEffPaintCol->hWnd))
					pc.dwTextDrawFlags |= DT_END_ELLIPSIS;

				// Textausgabe-Rechteck ermitteln
				BOOL bBtn = FALSE;
				BOOL bPermanent = TRUE;
				if (m_ppd->hWndFocusCol && m_ppd->pEffPaintCol->hWnd == m_ppd->hWndFocusCol && m_ppd->lEffPaintRow == m_ppd->lFocusRow)
					bPermanent = FALSE;
				if (!(pc.pCellType->IsCheckBox() && iJustify == DT_CENTER))
					bBtn = MustShowBtn(m_ppd->pEffPaintCol->hWnd, m_ppd->lEffPaintRow, TRUE, bPermanent);
				BOOL bDDBtn = MustShowDDBtn(m_ppd->pEffPaintCol->hWnd, m_ppd->lEffPaintRow, TRUE, bPermanent);

				pc.lTextAreaLeft = pc.rText.left = GetCellTextAreaLeft(m_ppd->lEffPaintRow, m_ppd->pEffPaintCol->hWnd, pc.rCell.left, m_ppd->tm.m_CellLeadingX, pc.Img, bBtn);
				if (iJustify == DT_LEFT || !(pc.dwTextDrawFlags & DT_SINGLELINE))
					pc.rText.left += m_ppd->tm.m_CellLeadingX;

				pc.rText.top = pc.rCell.top + pc.iCellLeadingY;

				pc.lTextAreaRight = pc.rText.right = GetCellTextAreaRight(m_ppd->lEffPaintRow, m_ppd->pEffPaintCol->hWnd, pc.rCell.right, m_ppd->tm.m_CellLeadingX, pc.Img, bBtn, bDDBtn);
				if (iJustify == DT_RIGHT || !(pc.dwTextDrawFlags & DT_SINGLELINE))
					pc.rText.right -= m_ppd->tm.m_CellLeadingX;
				if (pc.rText.right < pc.rText.left)
					pc.rText.right = pc.rText.left;

				pc.rText.bottom = pc.rCellVis.bottom;

				if (m_ppd->pPaintCol->bIndicateOverflow && pc.lTextLen > 0 && !m_ppd->pPaintCol->bInvisible && !(pc.dwTextDrawFlags & DT_WORDBREAK))
				{
					UINT uiFlags = pc.dwTextDrawFlags | DT_CALCRECT;
					RECT r = { 0, 0, 0, 0 };
					HGDIOBJ hOldFont = SelectObject(m_ppd->hDC, pc.hTextFont);
					DrawText(m_ppd->hDC, pc.lpsText, pc.lTextLen, &r, uiFlags);
					SelectObject(m_ppd->hDC, hOldFont);
					if ((r.right - r.left) > (pc.rText.right - pc.rText.left))
					{
						for (int i = 0; i < pc.lTextLen; ++i)
						{
							pc.lpsText[i] = '#';
						}
					}
				}

				iVAlign = GetEffCellTextVAlignMTbl(pEffCell, m_ppd->pEffPaintCol->pCol, m_ppd->pEffPaintRow);
				if (iVAlign < 0)
					iVAlign = DT_TOP;

				if (iVAlign != DT_TOP)
				{
					UINT uiFlags = pc.dwTextDrawFlags | DT_CALCRECT;
					RECT r = { pc.rText.left, 0, pc.rText.right, 0 };
					HGDIOBJ hOldFont = SelectObject(m_ppd->hDC, pc.hTextFont);

					if (pc.lTextLen < 1)
					{
						TCHAR szText[] = _T("A");
						DrawText(m_ppd->hDC, szText, 1, &r, uiFlags);
					}
					else
						DrawText(m_ppd->hDC, pc.lpsText, pc.lTextLen, &r, uiFlags);

					SelectObject(m_ppd->hDC, hOldFont);
					int iTxtHt = r.bottom - r.top;
					int iCellHt = pc.rCell.bottom - pc.rCell.top;

					if (iCellHt > iTxtHt)
					{
						if (iVAlign == DT_VCENTER)
							pc.rText.top = max(pc.rText.top, pc.rCell.top + ((iCellHt - iTxtHt) / 2));
						else if (iVAlign == DT_BOTTOM)
							pc.rText.top = pc.rCell.bottom - pc.iCellLeadingY - iTxtHt;
					}
				}
			}
			else if (pc.bRealRowHdr && m_ppd->hWndRowHdrCol)
			{
				// Text ermitteln
				BOOL bUnformatted = GetEffCellType(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol) == COL_CellType_CheckBox;
				if (bUnformatted)
					Ctd.FmtFieldToStr(m_ppd->hWndRowHdrCol, &pc.hsText, FALSE);
				else
					Ctd.TblGetColumnText(m_hWndTbl, Ctd.TblQueryColumnID(m_ppd->hWndRowHdrCol), &pc.hsText);

				if (pc.hsText)
				{
					pc.lpsText = Ctd.StringGetBuffer(pc.hsText, &pc.lTextLen);
					pc.lTextLen = Ctd.StrLenFromBufLen(pc.lTextLen);
				}				

				// Textausgabe-Flags ermitteln
				int iJustify = GetEffCellTextJustify(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol);
				pc.dwTextDrawFlags = DT_NOPREFIX | iJustify;
				if (m_ppd->bRowHdrColWordWrap && (m_ppd->tm.m_LinesPerRow > 1 || bVarRowHtEnabled))
					pc.dwTextDrawFlags |= DT_WORDBREAK;
				else if (!QueryFlags(MTBL_FLAG_EXPAND_CRLF))
					pc.dwTextDrawFlags |= DT_SINGLELINE;

				// Textausgabe-Rechteck ermitteln
				pc.lTextAreaLeft = pc.rText.left = GetCellTextAreaLeft(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol, pc.rCell.left, m_ppd->tm.m_CellLeadingX, pc.Img);
				if (iJustify == DT_LEFT || !(pc.dwTextDrawFlags & DT_SINGLELINE))
					pc.rText.left += m_ppd->tm.m_CellLeadingX;

				pc.rText.top = pc.rCell.top + pc.iCellLeadingY;

				pc.lTextAreaRight = pc.rText.right = GetCellTextAreaRight(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol, pc.rCell.right, m_ppd->tm.m_CellLeadingX, pc.Img);
				if (iJustify == DT_RIGHT || !(pc.dwTextDrawFlags & DT_SINGLELINE))
					pc.rText.right -= m_ppd->tm.m_CellLeadingX;
				if (pc.rText.right < pc.rText.left)
					pc.rText.right = pc.rText.left;

				pc.rText.bottom = pc.rCellVis.bottom;

				iVAlign = GetEffCellTextVAlign(m_ppd->lEffPaintRow, m_ppd->hWndRowHdrCol);
				if (iVAlign != DT_TOP)
				{
					UINT uiFlags = pc.dwTextDrawFlags | DT_CALCRECT;
					RECT r = { pc.rText.left, 0, pc.rText.right, 0 };
					HGDIOBJ hOldFont = SelectObject(m_ppd->hDC, pc.hTextFont);

					if (pc.lTextLen < 1)
					{
						TCHAR szText[] = _T("A");
						DrawText(m_ppd->hDC, szText, 1, &r, uiFlags);
					}
					else
						DrawText(m_ppd->hDC, pc.lpsText, pc.lTextLen, &r, uiFlags);

					SelectObject(m_ppd->hDC, hOldFont);
					int iTxtHt = r.bottom - r.top;
					int iCellHt = pc.rCell.bottom - pc.rCell.top;

					if (iCellHt > iTxtHt)
					{
						if (iVAlign == DT_VCENTER)
							pc.rText.top = max(pc.rText.top, pc.rCell.top + ((iCellHt - iTxtHt) / 2));
						else if (iVAlign == DT_BOTTOM)
							pc.rText.top = pc.rCell.bottom - pc.iCellLeadingY - iTxtHt;
					}
				}
			}
			else if (pc.bRealColHdr)
			{
				// Text ermitteln
				sColHdrText = _T("");
				GetColHdrText(m_ppd->pEffPaintCol->hWnd, sColHdrText);
				pc.lpsText = (LPTSTR)sColHdrText.c_str();
				pc.lTextLen = sColHdrText.size();

				if (Ctd.TblQueryColumnFlags(m_ppd->pEffPaintCol->hWnd, COL_DimTitle))
					pc.bTextDisabled = TRUE;

				// Textausgabe-Flags ermitteln
				if (QueryColumnHdrFlags(m_ppd->pPaintCol->hWnd, MTBL_COLHDR_FLAG_TXTALIGN_LEFT))
					pc.dwTextDrawFlags = DT_LEFT;
				else if (QueryColumnHdrFlags(m_ppd->pPaintCol->hWnd, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT))
					pc.dwTextDrawFlags = DT_RIGHT;
				else
					pc.dwTextDrawFlags = DT_CENTER;

				// Textausgabe-Rechteck ermitteln
				BOOL bBtn = FALSE;
				CMTblBtn* pBtn = GetEffItemBtn(pc.Item);
				bBtn = pBtn && pBtn->m_bShow;

				CopyRect(&pc.rText, &pc.rCell);
				pc.lTextAreaLeft = pc.rText.left = GetItemTextAreaLeft(pc.Item, pc.rCell.left, m_ppd->tm.m_CellLeadingX, pc.Img, bBtn);
				pc.lTextAreaRight = pc.rText.right = GetItemTextAreaRight(pc.Item, pc.rCell.right, m_ppd->tm.m_CellLeadingX, pc.Img, bBtn);

				pc.rText.top += pc.iCellLeadingY;
				if (pc.dwTextDrawFlags & DT_RIGHT)
					pc.rText.right -= m_ppd->tm.m_CellLeadingX;
				else
					pc.rText.left += m_ppd->tm.m_CellLeadingX;

				/*CopyRect( &pc.rText, &pc.rCell );
				pc.lTextAreaLeft = pc.rText.left;
				pc.lTextAreaRight = pc.rText.right;*/
			}
			else if (pc.bCorner)
			{
				// Text ermitteln
				pc.lpsText = (LPTSTR)m_ppd->sRowHdrTitle.c_str();
				pc.lTextLen = m_ppd->sRowHdrTitle.size();

				// Textausgabe-Flags ermitteln
				pc.dwTextDrawFlags = DT_CENTER;

				// Textausgabe-Rechteck ermitteln
				pc.lTextAreaLeft = pc.rText.left = GetCellTextAreaLeft(m_ppd->lRow, m_ppd->pPaintCol->hWnd, pc.rCell.left, m_ppd->tm.m_CellLeadingX, pc.Img);
				pc.rText.top = pc.rCell.top;
				pc.lTextAreaRight = pc.rText.right = GetCellTextAreaRight(m_ppd->lRow, m_ppd->pPaintCol->hWnd, pc.rCell.right, m_ppd->tm.m_CellLeadingX, pc.Img);
				pc.rText.bottom = pc.rCell.bottom;
			}

			if (QueryFlags(MTBL_FLAG_EXPAND_TABS))
				pc.dwTextDrawFlags |= DT_EXPANDTABS;

			// Zu zeichnende Tree-Linien ermitteln
			pc.bTreeBottomUp = (m_ppd->lRow >= 0 ? m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP : m_dwTreeFlags & MTBL_TREE_FLAG_SPLIT_BOTTOM_UP);
			pc.bHorzTreeLine = FALSE;
			pc.bVertTreeLineToParent = FALSE;
			pc.bVertTreeLineToChild = FALSE;
			if (!bPseudoRow && (m_ppd->bRowHdrMirror ? m_ppd->hWndRowHdrCol == m_hWndTreeCol : !m_ppd->pPaintCol->bPseudo && m_ppd->pEffPaintCol->hWnd == m_hWndTreeCol))
			{
				// Horizontale Linie
				if (m_ppd->iRowLevel == 0)
					pc.bHorzTreeLine = m_Rows->QueryRowFlags(m_ppd->lEffPaintRow, MTBL_ROW_CANEXPAND) && m_iNodeSize > 0 && !QueryTreeFlags(MTBL_TREE_FLAG_INDENT_NONE);
				else if (m_ppd->iRowLevel > 0)
				{
					pc.bHorzTreeLine = TRUE;
				}

				// Vertikale Linie zur Elternzeile?
				if (m_Rows->GetParentRow(m_ppd->lEffPaintRow) != TBL_Error)
					pc.bVertTreeLineToParent = TRUE;

				// Vertikale Linie zur Kindzeile?
				BOOL bChildRowMerged;
				CMTblRow *pRow = m_ppd->pEffPaintRow;
				pRow = m_Rows->GetFirstChildRow(pRow, pc.bTreeBottomUp);
				while (pRow)
				{
					if (m_ppd->ppmc)
					{
						bChildRowMerged = pRow->GetNr() >= m_ppd->ppmc->pRowFrom->GetNr() && pRow->GetNr() <= m_ppd->ppmc->pRowTo->GetNr();
					}
					else
						bChildRowMerged = FALSE;

					if (!Ctd.TblQueryRowFlags(m_hWndTbl, pRow->GetNr(), ROW_Hidden) && !bChildRowMerged)
					{
						pc.bVertTreeLineToChild = TRUE;
						break;
					}
					pRow = m_Rows->GetNextChildRow(pRow, pc.bTreeBottomUp);
				}
			}

			// Zelle zeichnen
			pc.bMerge = m_ppd->ppmc != NULL;

			pc.bGroupColHdr = (pColHdrGrp != NULL);
			pc.bGroupColHdrTop = FALSE;
			pc.bGroupColHdrNCH = FALSE;
			pc.bLastGroupColHdr = pColHdrGrp != NULL && !bNotLastGroupColHdr;
			pc.bFirstGroupColHdr = pColHdrGrp != NULL && bFirstGroupColHdr;

			pc.dwTopLog = m_ppd->ppmc ? m_ppd->ppmc->dwTopLog : m_ppd->dwRowTopLog;
			pc.dwBottomLog = m_ppd->ppmc ? m_ppd->ppmc->dwBottomLog : m_ppd->dwRowBottomLog;

			// Selektiert zeichnen?
			BOOL bSelected = FALSE;
			BOOL bSelectedCell = FALSE;

			if (m_bCellMode && pEffCell && pEffCell->QueryFlags(MTBL_CELL_FLAG_SELECTED))
			{
				bSelected = TRUE;
				bSelectedCell = TRUE;
			}

			if (!bSelected && (m_ppd->pPaintCol->bSelected || m_ppd->bRowSelected))
			{
				if (QueryFlags(MTBL_FLAG_FULL_OVERLAP_SELECTION))
					bSelected = TRUE;
				else
					bSelected = !(m_ppd->pPaintCol->bSelected && m_ppd->bRowSelected);
			}

			pc.Bits.Selected = (bSelected ? 1 : 0);

			// Zeichnen
			if (!(pColHdrGrp && pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS)))
				PaintCell(pc);

			if (pc.hsText)
			{
				Ctd.HStrUnlock(pc.hsText);
				Ctd.HStringUnRef(pc.hsText);
			}				

			// Font löschen?
			if (bDeleteFont)
			{
				DeleteObject(pc.hTextFont);
				pc.hTextFont = NULL;
			}

			// Gruppen-Spaltenheader zeichnen?
			if (pColHdrGrp)
			{
				// Items setzen
				pc.Item.DefineAsColHdrGrp(m_hWndTbl, pColHdrGrp);
				pc.HorzLineItem = pc.VertLineItem = pc.Item;

				// Effektives Highlighting ermitteln
				if (bAnyHighlightedItems)
				{
					CMTblItem item = pc.Item;
					item.WndHandle = m_ppd->pPaintCol->pCol->m_hWnd;
					GetEffItemHighlighting(item, pc.Hili);
				}

				// Identifikations-Booleans setzen
				pc.bCell = pc.bRealCell = pc.bColHdr = pc.bRealColHdr = pc.bRowHdr = pc.bRealRowHdr = pc.bCorner = FALSE;
				pc.bColHdrGrp = TRUE;

				// Benutzerdefiniert gezeichnet?
				pc.iODIType = 0;
				if (pColHdrGrp->IsOwnerdrawn())
					pc.iODIType = MTBL_ODI_COLHDRGRP;

				// Kennzeichen setzen
				pc.bGroupColHdrNCH = pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS);
				pc.bGroupColHdrTop = !pc.bGroupColHdrNCH;

				// Y-Leading
				if (pc.bGroupColHdrTop)
					pc.iCellLeadingY = m_ppd->tm.m_CellLeadingY / 3;
				else
					pc.iCellLeadingY = m_ppd->tm.m_CellLeadingY;

				// X-Koordinaten setzen
				pc.rCell.left = m_ppd->pFirstGroupColHdr->ptX.x;
				pc.rCell.right = m_ppd->pLastGroupColHdr->ptX.y;
				if (bNotLastGroupColHdr)
					pc.rCellVis.right += 1;

				// Y-Koordinaten setzen
				pc.rCell.top = rCellOrig.top;
				pc.rCell.bottom = pc.rCell.top + iGroupColHdrHt;
				pc.rCellVis.top = rCellVisOrig.top;
				pc.rCellVis.bottom = min(pc.rCell.bottom, rCellVisOrig.bottom);

				// Linien
				if (bNotLastGroupColHdr)
					pc.bColLine = FALSE;
				if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOINNERHORZLINE) && !pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS))
					pc.bRowLine = FALSE;

				// Bild
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				pc.Img = GetEffColHdrGrpImage(pColHdrGrp, dwFlags);

				// Text
				pc.lpsText = (LPTSTR)pColHdrGrp->m_Text.c_str();
				pc.lTextLen = pColHdrGrp->m_Text.size();
				pc.bTextDisabled = FALSE;

				// Textausgabe-Rechteck
				CopyRect(&pc.rText, &pc.rCell);
				pc.rText.top += pc.iCellLeadingY;

				// Textausgabe-Flags
				if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_TXTALIGN_LEFT))
					pc.dwTextDrawFlags = DT_LEFT;
				else if (pColHdrGrp->QueryFlags(MTBL_CHG_FLAG_TXTALIGN_RIGHT))
					pc.dwTextDrawFlags = DT_RIGHT;
				else
					pc.dwTextDrawFlags = DT_CENTER;

				// Hintergrundfarbe
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				pc.crBackColor = GetDrawColor(GetEffColHdrGrpBackColor(pColHdrGrp, dwFlags));

				// Textfarbe
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				pc.crTextColor = GetDrawColor(GetEffColHdrGrpTextColor(pColHdrGrp, dwFlags));

				// Font
				bDeleteFont = FALSE;
				dwFlags = bAnyHighlightedItems ? GEC_HILI_PD : 0;
				pc.hTextFont = GetEffColHdrGrpFont(m_ppd->hDC, pColHdrGrp, &bDeleteFont, dwFlags);

				// Zeichnen
				PaintCell(pc);

				// Font löschen?
				if (bDeleteFont)
				{
					DeleteObject(pc.hTextFont);
					pc.hTextFont = NULL;
				}

				// Original-Koordinaten wiederherstellen
				CopyRect(&pc.rCell, &rCellOrig);
				CopyRect(&pc.rCellVis, &rCellVisOrig);
			}

			if (!m_ppd->pPaintCol->bPseudo && !m_ppd->pPaintCol->bLocked && iFirstPaintCol < 0)
				iFirstPaintCol = i;

			// Selektion invertieren oder abdunkeln
			if (!m_ppd->bAnySelColors && !m_bySelDarkening)
			{
				LPMTBLPAINTCOL pCol;
				RECT r;

				pCol = m_ppd->pPaintCol;
				if (bSelectedCell)
				{
					r.left = pCol->ptVisX.x;
					if (pCol->bFirstUnlocked && r.left == pCol->ptX.x)
						++r.left;
					r.top = pc.rCellVis.top;
					r.right = pCol->ptVisX.y + 1;
					r.bottom = pc.rCellVis.bottom + 1;
					InvertRect_Hooked(m_ppd->hDC, &r);
				}
				else
				{
					if (QueryFlags(MTBL_FLAG_FULL_OVERLAP_SELECTION) && pCol->bSelected && m_ppd->bRowSelected)
					{
						r.left = pCol->ptVisX.x;
						if (pCol->bFirstUnlocked && r.left == pCol->ptX.x)
							++r.left;
						r.top = pc.rCellVis.top;
						r.right = pCol->ptVisX.y + 1;
						r.bottom = pc.rCellVis.bottom + 1;
						InvertRect_Hooked(m_ppd->hDC, &r);
					}
					else
					{
						if (pCol->bSelected)
						{
							r.left = pCol->ptVisX.x;
							if (pCol->bFirstUnlocked && r.left == pCol->ptX.x)
								++r.left;
							r.top = pc.rCellVis.top;
							r.right = pCol->ptVisX.y;
							r.bottom = pc.rCellVis.bottom + 1;
							InvertRect_Hooked(m_ppd->hDC, &r);
						}

						if (m_ppd->bRowSelected)
						{
							r.left = pCol->ptVisX.x;
							r.top = pc.rCellVis.top;
							r.right = pCol->ptVisX.y + 1;
							r.bottom = pc.rCellVis.bottom;

							BOOL bMergedCellNotLast = FALSE;
							if (m_ppd->ppmc && m_ppd->lRow != m_ppd->ppmc->pRowTo->GetNr())
								//--r.bottom;
								bMergedCellNotLast = TRUE;

							BOOL bMergedRowNotLast = FALSE;
							if (m_ppd->pmr && m_ppd->lRow != m_ppd->pmr->lRowTo)
								bMergedRowNotLast = TRUE;

							if (bMergedRowNotLast && !bMergedCellNotLast)
								++r.bottom;

							InvertRect_Hooked(m_ppd->hDC, &r);
						}
					}
				}
			}

			// Focus
			PaintFocus(&pc.rCellVis);
		}
	}

	return TRUE;
}

// PaintTable
BOOL MTBLSUBCLASS::PaintTable()
{
	// Double-Buffer?
	BOOL bDrawToBuffer = FALSE;
	//m_bDoubleBuffer = TRUE;
	if (m_bDoubleBuffer)
	{
		int iWid = m_ppd->rUpd.right - m_ppd->rUpd.left;
		int iHt = m_ppd->rUpd.bottom - m_ppd->rUpd.top;
		m_ppd->hBitmDoubleBuffer = CreateCompatibleBitmap(m_ppd->hDCTbl, iWid, iHt);
		if (m_ppd->hBitmDoubleBuffer)
		{
			m_ppd->hDCDoubleBuffer = CreateCompatibleDC(m_ppd->hDCTbl);
			if (m_ppd->hDCDoubleBuffer)
			{
				m_ppd->hOldBitmDoubleBuffer = (HBITMAP)SelectObject(m_ppd->hDCDoubleBuffer, m_ppd->hBitmDoubleBuffer);

				RECT r;
				GetClientRect(m_hWndTbl, &r);

				int iOffsX = r.left - m_ppd->rUpd.left;
				int iOffsY = r.top - m_ppd->rUpd.top;
				if (SetViewportOrgEx(m_ppd->hDCDoubleBuffer, iOffsX, iOffsY, NULL))
				{
					m_ppd->hDC = m_ppd->hDCDoubleBuffer;
					bDrawToBuffer = TRUE;
				}
			}
		}
	}

	// Allgemeine Initialisierungen
	m_ppd->pc.hsText = 0;

	// Irgendeine Selektionsfarbe gesetzt?
	m_ppd->bAnySelColors = m_dwSelBackColor != MTBL_COLOR_UNDEF || m_dwSelTextColor != MTBL_COLOR_UNDEF;

	// Theme
	CTheme &th = theTheme();
	if (th.CanUse() && th.IsAppThemed())
		m_ppd->hTheme = th.OpenThemeData(m_hWndTbl, L"Button");
	else
		m_ppd->hTheme = NULL;


	// Tabellen-Maße holen
	m_ppd->tm.Get(m_hWndTbl);

	// Vert. Scroll-Positionen ermitteln
	m_ppd->iScrPosV = GetScrollPosV(&m_ppd->tm, FALSE);
	if (m_ppd->tm.m_SplitBarTop > 0)
		m_ppd->iScrPosVSplit = GetScrollPosV(&m_ppd->tm, TRUE);
	else
		m_ppd->iScrPosVSplit = 0;

	// Row-Header ermitteln
	m_ppd->wRowHdrFlags = 0;
	m_ppd->sRowHdrTitle = _T("");
	m_ppd->hWndRowHdrCol = NULL;
	m_ppd->bRowHdrColWordWrap = FALSE;
	m_ppd->lRowHdrColJustify = 0;
	int iRowHdrWid = 0;
	if (Ctd.TblQueryRowHeader(m_hWndTbl, &m_ppd->pc.hsText, 2048, &iRowHdrWid, &m_ppd->wRowHdrFlags, &m_ppd->hWndRowHdrCol))
	{
		//Ctd.HStrToTStr(m_ppd->pc.hsText, m_ppd->sRowHdrTitle);
		Ctd.HStrToTStr(m_ppd->pc.hsText, m_ppd->sRowHdrTitle, TRUE);

		if (m_ppd->hWndRowHdrCol)
		{
			m_ppd->bRowHdrColWordWrap = SendMessage(m_ppd->hWndRowHdrCol, TIM_QueryColumnFlags, 0, COL_MultilineCell);
			m_ppd->lRowHdrColJustify = GetColJustify(m_ppd->hWndRowHdrCol);
		}
	}

	// Split-Bereich ermitteln
	m_ppd->bSplit = FALSE;
	m_ppd->lMinSplitRow = m_ppd->lMaxSplitRow = TBL_Error;
	if (m_ppd->tm.m_SplitBarTop > 0)
	{
		m_ppd->bSplit = TRUE;
		//if ( MTblQueryVisibleSplitRange( m_hWndTbl, &m_ppd->lMinSplitRow, &m_ppd->lMaxSplitRow ) )
		if (QueryVisRange(m_ppd->lMinSplitRow, m_ppd->lMaxSplitRow, TRUE))
			m_ppd->lMaxSplitRow += 1;
	}

	// Was ist an sichtbaren Zeilen da?
	long lRow = -1;
	m_ppd->bAnyRow = FindNextRow(&lRow, 0, ROW_Hidden);
	if (m_ppd->bSplit)
	{
		lRow = TBL_MinRow;
		m_ppd->bAnySplitRow = FindNextRow(&lRow, 0, ROW_Hidden, -1);
	}
	else
		m_ppd->bAnySplitRow = FALSE;

	// Focus ermitteln
	Ctd.TblQueryFocus(m_hWndTbl, &m_ppd->lFocusRow, &m_ppd->hWndFocusCol);

	// Kein Zeichnen im Focus-Bereich?
	m_ppd->bClipFocusArea = QueryFlags(MTBL_FLAG_CLIP_FOCUS_AREA);

	// Keine Zeilenlinien?
	m_ppd->bNoRowLines = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SuppressRowLines);

	// Letzte Spaltenlinie unterdrücken?
	m_ppd->bSuppressLastColLine = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_SuppressLastColLine);

	// Spalten ermitteln
	BOOL bFirstUnlocked = FALSE, bFirstVis, bFirstVisFound = FALSE, bLocked, bLastLocked, bPaint;
	HWND hWndCol = NULL;
	int iColPos;
	LPMTBLPAINTCOL pPaintCol;
	POINT pt, ptPhys;

	m_ppd->iNormalCols = m_ppd->iCols = 0;
	m_ppd->iPaintColFocusCell = -1;

	// Spalte Row-Header
	if (m_ppd->tm.m_RowHeaderRight > 0)
	{
		pt.x = ptPhys.x = 0;
		pt.y = ptPhys.y = m_ppd->tm.m_RowHeaderRight - 1;

		// Spalte
		//pPaintCol = m_ppd->PaintCols[m_ppd->iCols];
		pPaintCol = m_ppd->EnsureValidPaintCol(m_ppd->iCols);

		// Daten setzen
		SetPaintColData(NULL, 0, TRUE, FALSE, FALSE, FALSE, FALSE, FALSE, TRUE, &pt, &ptPhys, pPaintCol);

		++m_ppd->iCols;
	}

	// Spalten
	pt.y = m_ppd->tm.m_RowHeaderRight;
	ptPhys.y = pt.y;
	iColPos = 1;

	CURCOLS_HANDLE_VECTOR & CurCols = m_CurCols->GetVectorHandle();
	CURCOLS_WIDTH_VECTOR & CurColsWid = m_CurCols->GetVectorWidth();
	int iCol, iCols = CurCols.size();

	for (iCol = 0; iCol < iCols; ++iCol)
	{
		hWndCol = CurCols[iCol];

		if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible))
		{
			bPaint = TRUE;

			pt.x = pt.y;
			//pt.y += SendMessage( hWndCol, TIM_QueryColumnWidth, 0, 0 );
			pt.y += CurColsWid[iCol];
			pt.y -= 1;

			// Erste sichtbare Spalte?
			if (!bFirstVisFound)
			{
				bFirstVis = TRUE;
				bFirstVisFound = TRUE;
			}
			else
				bFirstVis = FALSE;

			// Verankerte Spalte?
			bLocked = (iColPos <= m_ppd->tm.m_LockedColumns && pt.y <= m_ppd->tm.m_LockedColumnsRight);

			// Letzte verankerte Spalte
			bLastLocked = bLocked;
			if (m_ppd->iNormalCols > 0 && bLocked && m_ppd->PaintCols[m_ppd->iCols - 1]->bLocked)
				m_ppd->PaintCols[m_ppd->iCols - 1]->bLastLocked = FALSE;

			// Erste nicht verankerte Spalte?
			if (!bLocked && m_ppd->iNormalCols > 0 && m_ppd->PaintCols[m_ppd->iCols - 1]->bLastLocked)
				bFirstUnlocked = TRUE;
			else
				bFirstUnlocked = FALSE;

			// 1 Pixel verbreitern, wenn das die erste nicht verankerte Spalte ist
			if (bFirstUnlocked)
				pt.y += 1;

			// Physikalische Koordinaten ermitteln
			ptPhys.x = (bLocked ? pt.x : pt.x - m_ppd->tm.m_ColumnOrigin);
			ptPhys.y = (bLocked ? pt.y : pt.y - m_ppd->tm.m_ColumnOrigin);

			// Überhaupt ein Teil sichtbar?
			if (ptPhys.x >= (bLocked ? m_ppd->tm.m_LockedColumnsRight : m_ppd->tm.m_ColumnsRight))
				bPaint = FALSE;
			if (ptPhys.y < (bLocked ? m_ppd->tm.m_LockedColumnsLeft : m_ppd->tm.m_ColumnsLeft))
				bPaint = FALSE;

			// Spalte im Update-Rechteck?
			if (ptPhys.x >= m_ppd->rUpd.right || ptPhys.y < m_ppd->rUpd.left)
				bPaint = FALSE;

			// Spalte
			//pPaintCol = m_ppd->PaintCols[m_ppd->iCols];
			pPaintCol = m_ppd->EnsureValidPaintCol(m_ppd->iCols);

			// Daten setzen
			SetPaintColData(hWndCol, iColPos, bPaint, bLocked, bLastLocked, bFirstUnlocked, bFirstVis, FALSE, FALSE, &pt, &ptPhys, pPaintCol);

			// Index Paint-Spalte der Fokus-Zelle
			if (m_bCellMode)
			{
				if (m_hWndColFocus && hWndCol == m_hWndColFocus)
					m_ppd->iPaintColFocusCell = m_ppd->iCols;
				else if (!m_hWndColFocus && !m_pRowFocus && !m_ppd->bAnyRow && bFirstVis)
					m_ppd->iPaintColFocusCell = m_ppd->iCols;
			}

			++m_ppd->iNormalCols;
			++m_ppd->iCols;

			pt.y += 1;
		}
		++iColPos;
	}

	if (m_ppd->iNormalCols > 0)
	{
		// Letzte sichtbare Spalte setzen
		pPaintCol->bLastVisible = TRUE;

		// Leere Spalte nötig?
		if (pPaintCol->ptX.y < m_ppd->tm.m_ColumnsRight)
		{
			ptPhys.x = pPaintCol->ptX.y + 1;
			ptPhys.y = m_ppd->tm.m_ColumnsRight;
			if (!(ptPhys.x >= m_ppd->rUpd.right || ptPhys.y < m_ppd->rUpd.left))
			{
				// Logische Koordinaten setzen
				pt.x = pPaintCol->ptLogX.y + 1;
				pt.y = pPaintCol->ptLogX.y + 1 + (ptPhys.y - ptPhys.x + 1);

				// Spalte erzeugen
				//pPaintCol = m_ppd->PaintCols[m_ppd->iCols];
				pPaintCol = m_ppd->EnsureValidPaintCol(m_ppd->iCols);

				// Daten setzen
				SetPaintColData(NULL, 0, TRUE, FALSE, FALSE, FALSE, FALSE, TRUE, FALSE, &pt, &ptPhys, pPaintCol);

				++m_ppd->iCols;
			}
		}
	}
	else
	{
		// Leere Spalte nötig?
		if (m_ppd->tm.m_ColumnsLeft < m_ppd->tm.m_ColumnsRight)
		{
			ptPhys.x = m_ppd->tm.m_ColumnsLeft;
			ptPhys.y = m_ppd->tm.m_ColumnsRight;
			if (!(ptPhys.x >= m_ppd->rUpd.right || ptPhys.y < m_ppd->rUpd.left))
			{
				// Logische Koordinaten setzen
				pt.x = m_ppd->tm.m_ColumnsLeft;
				pt.y = m_ppd->tm.m_ColumnsRight;

				// Spalte
				//pPaintCol = m_ppd->PaintCols[m_ppd->iCols];
				pPaintCol = m_ppd->EnsureValidPaintCol(m_ppd->iCols);

				// Daten setzen
				SetPaintColData(NULL, 0, TRUE, FALSE, FALSE, FALSE, TRUE, TRUE, FALSE, &pt, &ptPhys, pPaintCol);

				++m_ppd->iCols;
			}
		}
	}

	// Zu zeichnende Merge-Rowheader ermitteln
	GetPaintMergeRowHdrs();

	// Zu zeichnende Merge-Zellen ermitteln
	GetPaintMergeCells();

	// Zeichnen
	long lContextRow;

	// Hintergrundfarbe der Tabelle ermitteln
	COLORREF clrTblBackColor = GetDrawColor((COLORREF)Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindow));
	m_ppd->clrTblBackColor = clrTblBackColor;

	// Textfarbe der Tabelle ermitteln
	m_ppd->clrTblTextColor = GetDrawColor((COLORREF)Ctd.ColorGet(m_hWndTbl, COLOR_IndexWindowText));

	// Graue Header?
	m_ppd->bGrayedHeaders = Ctd.TblQueryTableFlags(m_hWndTbl, TBL_Flag_GrayedHeaders);

	// Hintergrundfarbe Header
	if (m_ppd->bGrayedHeaders)
		//m_ppd->clrHdrBack = GetDrawColor( GetSysColor( COLOR_3DFACE ) );
		m_ppd->clrHdrBack = GetDrawColor(GetDefaultColor(GDC_HDR_BKGND));
	else
		m_ppd->clrHdrBack = m_ppd->clrTblBackColor;

	// Linienfarbe Header
	if (m_ppd->bGrayedHeaders)
		m_ppd->clrHdrLine = GetDrawColor(GetDefaultColor(GDC_HDR_LINE));
	else
		m_ppd->clrHdrLine = GetDrawColor(0);

	// 3D HiLight-Farbe
	//m_ppd->clr3DHiLight = GetDrawColor( GetSysColor( COLOR_3DHILIGHT ) );
	m_ppd->clr3DHiLight = GetDrawColor(GetDefaultColor(GDC_HDR_3DHL));

	// Kontext der Tabelle ermitteln
	lContextRow = Ctd.TblQueryContext(m_hWndTbl);

	// Fokus...
	m_ppd->bPaintFocus = (!m_bNoFocusPaint && !QueryFlags(MTBL_FLAG_SUPPRESS_FOCUS) && !IsRectEmpty(&m_rFocus) && (m_bCellMode ? m_ppd->iPaintColFocusCell >= 0 : TRUE));
	if (m_ppd->bPaintFocus)
	{
		CopyRect(&m_ppd->rFocusClip, &m_rFocus);
		if (m_bCellMode)
		{
			if (m_ppd->tm.m_LockedColumnsRight > 0 && !m_ppd->PaintCols[m_ppd->iPaintColFocusCell]->bLocked)
			{
				if (m_rFocus.left + m_lFocusFrameThick + m_lFocusFrameOffsX < m_ppd->tm.m_LockedColumnsRight + 1)
					m_ppd->rFocusClip.left = max(m_ppd->tm.m_LockedColumnsRight + 1, m_ppd->rFocusClip.left);
			}
			else
			{
				if (m_rFocus.left + m_lFocusFrameThick + m_lFocusFrameOffsX < m_ppd->tm.m_RowHeaderRight)
					m_ppd->rFocusClip.left = max(m_ppd->tm.m_RowHeaderRight, m_ppd->rFocusClip.left);
			}
		}

		long lRowFocus = GetRowNrFocus();
		if (lRowFocus >= 0)
		{
			if (m_rFocus.top + m_lFocusFrameThick + m_lFocusFrameOffsY < m_ppd->tm.m_RowsTop)
				m_ppd->rFocusClip.top = max(m_ppd->tm.m_RowsTop, m_ppd->rFocusClip.top);

			if (m_ppd->tm.m_SplitRowsTop > 0)
				m_ppd->rFocusClip.bottom = min(m_ppd->tm.m_RowsBottom, m_ppd->rFocusClip.bottom);
		}
		else
		{
			if (m_ppd->tm.m_SplitRowsTop > 0 && m_rFocus.top + m_lFocusFrameThick + m_lFocusFrameOffsY < m_ppd->tm.m_SplitRowsTop)
				m_ppd->rFocusClip.top = max(m_ppd->tm.m_SplitRowsTop, m_ppd->rFocusClip.top);
			else if (m_ppd->tm.m_SplitBarTop > 0)
				m_ppd->rFocusClip.top = max(m_ppd->tm.m_SplitBarTop, m_ppd->rFocusClip.top);
		}
	}
	else
	{
		SetRectEmpty(&m_ppd->rFocusClip);
	}

	// *** Allgemeine GDI-Objekte erzeugen ***
	m_ppd->hBrTblBack = CreateSolidBrush(m_ppd->clrTblBackColor);

	// *** Spaltenheader zeichnen ***
	m_ppd->bColHdr = TRUE;
	m_ppd->bEmptyRow = FALSE;
	m_ppd->lRow = -1;
	m_ppd->bSplitRow = FALSE;

	m_ppd->dwRowTop = 0;
	m_ppd->dwRowTopVis = m_ppd->dwRowTop;

	m_ppd->dwRowBottom = m_ppd->tm.m_HeaderHeight - 1;
	m_ppd->dwRowBottomVis = max(min(m_ppd->tm.m_HeaderHeight - 1, m_ppd->tm.m_ClientRect.bottom), 0);

	m_ppd->dwRowTopLog = 0;
	m_ppd->dwRowBottomLog = m_ppd->tm.m_HeaderHeight - 1;

	PaintRow();

	// *** Normale Zeilen zeichnen ***
	m_ppd->bColHdr = FALSE;
	m_ppd->bEmptyRow = FALSE;
	m_ppd->bSplitRow = FALSE;

	m_ppd->dwRowTop = m_ppd->tm.m_RowsTop;
	m_ppd->dwRowTopVis = m_ppd->dwRowTop;

	lRow = m_ppd->lRow = m_ppd->tm.m_MinVisibleRow - 1;

	m_ppd->dwRowTopLog = m_ppd->tm.m_RowsTop - 1;
	if (m_pCounter->GetRowHeights() <= 0)
		m_ppd->dwRowTopLog += ((m_ppd->iScrPosV * m_ppd->tm.m_RowHeight) + 1);
	else
		m_ppd->dwRowTopLog += (m_Rows->GetRowLogX(m_Rows->GetVisPosRow(m_ppd->iScrPosV), m_ppd->tm.m_RowHeight) + 1);

	long lRowHt;
	while (m_ppd->dwRowTop < m_ppd->tm.m_RowsBottom)
	{
		// Nächste Zeile ermitteln
		if (!m_ppd->bEmptyRow)
		{
			if (FindNextRow(&lRow, 0, ROW_Hidden))
			{
				m_ppd->lRow = lRow;
			}
			else
			{
				m_ppd->bEmptyRow = TRUE;
				++m_ppd->lRow;
			}
		}
		else
			++m_ppd->lRow;

		// Zeilenhöhe ermitteln
		if (m_pCounter->GetRowHeights() <= 0 || m_ppd->bEmptyRow)
			lRowHt = m_ppd->tm.m_RowHeight;
		else
			lRowHt = GetEffRowHeight(m_ppd->lRow, &m_ppd->tm);

		// Unten setzen
		m_ppd->dwRowBottomVis = min(m_ppd->dwRowTopVis + lRowHt - 1, m_ppd->tm.m_RowsBottom);
		m_ppd->dwRowBottom = m_ppd->dwRowTop + lRowHt - 1;
		m_ppd->dwRowBottomLog = m_ppd->dwRowTopLog + lRowHt - 1;

		// Zeile zeichnen
		PaintRow();

		// Oben um Zeilenhöhe nach unten schieben
		m_ppd->dwRowTop += lRowHt;
		m_ppd->dwRowTopVis += lRowHt;
		m_ppd->dwRowTopLog += lRowHt;
	}

	// *** Split-Bar zeichnen ***
	if (m_ppd->bSplit)
	{
		RECT r, rCl;
		GetClientRect(m_hWndTbl, &rCl);

		// Hintergrund
		COLORREF clr = GetDrawColor(GetSysColor(COLOR_3DFACE));
		HBRUSH hBr = CreateSolidBrush(clr);
		if (hBr)
		{
			SetRect(&r, rCl.left, m_ppd->tm.m_SplitBarTop, rCl.right, m_ppd->tm.m_SplitBarTop + m_ppd->tm.m_SplitBarHeight - 1);
			FillRect(m_ppd->hDC, &r, hBr);
			DeleteObject(hBr);
			PaintFocus(&r);
		}

		// Linien
		int iY;
		HPEN hPen = (HPEN)GetStockObject(BLACK_PEN);
		HGDIOBJ hOldPen = SelectObject(m_ppd->hDC, hPen);

		// Obere Linie
		iY = m_ppd->tm.m_SplitBarTop;
		MoveToEx(m_ppd->hDC, rCl.left, iY, NULL);
		LineTo(m_ppd->hDC, rCl.right, iY);

		// Untere Linie
		iY = m_ppd->tm.m_SplitBarTop + m_ppd->tm.m_SplitBarHeight - 1;
		MoveToEx(m_ppd->hDC, rCl.left, iY, NULL);
		LineTo(m_ppd->hDC, rCl.right, iY);

		SelectObject(m_ppd->hDC, hOldPen);

		// Invertierungen
		MTBLPAINTCOL *pPaintCol;
		r.top = m_ppd->tm.m_SplitBarTop;
		r.bottom = m_ppd->tm.m_SplitBarTop + m_ppd->tm.m_SplitBarHeight;
		for (int i = 0; i < m_ppd->iCols; ++i)
		{
			pPaintCol = m_ppd->PaintCols[i];
			if (pPaintCol->bPaint && pPaintCol->bSelected && !m_ppd->bAnySelColors && !m_bySelDarkening)
			{
				r.left = pPaintCol->ptVisX.x;
				if (pPaintCol->bFirstUnlocked && r.left == pPaintCol->ptX.x)
					++r.left;
				r.right = pPaintCol->ptVisX.y;
				InvertRect_Hooked(m_ppd->hDC, &r);
			}
		}
	}

	// *** Split-Zeilen zeichnen ***
	if (m_ppd->bSplit)
	{
		m_ppd->bColHdr = FALSE;
		m_ppd->bEmptyRow = FALSE;
		m_ppd->bSplitRow = TRUE;

		m_ppd->dwRowTop = m_ppd->tm.m_SplitRowsTop;
		m_ppd->dwRowTopVis = m_ppd->dwRowTop;

		if (m_ppd->lMinSplitRow != TBL_Error)
		{
			m_ppd->lRow = m_ppd->lMinSplitRow - 1;
			if (m_ppd->lRow < TBL_MinSplitRow)
				m_ppd->lRow = TBL_MinRow;
		}
		else
			m_ppd->lRow = TBL_MinRow;
		lRow = m_ppd->lRow;

		m_ppd->dwRowTopLog = m_ppd->tm.m_SplitRowsTop - 1;
		if (m_pCounter->GetRowHeights() <= 0)
			m_ppd->dwRowTopLog += ((m_ppd->iScrPosVSplit * m_ppd->tm.m_RowHeight) + 1);
		else
			m_ppd->dwRowTopLog += (m_Rows->GetRowLogX(m_Rows->GetVisPosRow(m_ppd->iScrPosVSplit), m_ppd->tm.m_RowHeight) + 1);

		while (m_ppd->dwRowTop < m_ppd->tm.m_SplitRowsBottom)
		{
			// Nächste Zeile ermitteln
			if (!m_ppd->bEmptyRow)
			{
				if (FindNextRow(&lRow, 0, ROW_Hidden, -1))
					m_ppd->lRow = lRow;
				else
				{
					m_ppd->bEmptyRow = TRUE;
					++m_ppd->lRow;
				}
			}
			else
				++m_ppd->lRow;

			// Zeilenhöhe ermitteln
			if (m_pCounter->GetRowHeights() <= 0 || m_ppd->bEmptyRow)
				lRowHt = m_ppd->tm.m_RowHeight;
			else
				lRowHt = GetEffRowHeight(m_ppd->lRow, &m_ppd->tm);

			// Unten setzen
			m_ppd->dwRowBottomVis = min(m_ppd->dwRowTopVis + lRowHt - 1, m_ppd->tm.m_SplitRowsBottom);
			m_ppd->dwRowBottom = m_ppd->dwRowTop + lRowHt - 1;
			m_ppd->dwRowBottomLog = m_ppd->dwRowTopLog + lRowHt - 1;

			// Zeile zeichnen
			PaintRow();

			// Oben um Zeilenhöhe nach unten schieben
			m_ppd->dwRowTop += lRowHt;
			m_ppd->dwRowTopVis += lRowHt;
			m_ppd->dwRowTopLog += lRowHt;
		}
	}

	// Kontext wiederherstellen
	Ctd.TblSetContextEx(m_hWndTbl, lContextRow);

	// *** Aufräumen ***
	cleanup:

	// HSTRING
	/*if (m_ppd->pc.hsText)
		Ctd.HStringUnRef(m_ppd->pc.hsText);*/

	// GDI-Objekte
	int iRet;
	if (m_ppd->hBrTblBack)
		iRet = DeleteObject(m_ppd->hBrTblBack);

	// Theme
	if (m_ppd->hTheme)
		th.CloseThemeData(m_ppd->hTheme);

	// Double Buffer
	if (m_bDoubleBuffer)
	{
		if (m_ppd->hDCDoubleBuffer)
			SetViewportOrgEx(m_ppd->hDCDoubleBuffer, 0, 0, NULL);

		if (bDrawToBuffer)
		{
			int iX = m_ppd->rUpd.left;
			int iY = m_ppd->rUpd.top;
			int iWid = m_ppd->rUpd.right - m_ppd->rUpd.left;
			int iHt = m_ppd->rUpd.bottom - m_ppd->rUpd.top;
			BitBlt(m_ppd->hDCTbl, iX, iY, iWid, iHt, m_ppd->hDCDoubleBuffer, 0, 0, SRCCOPY);
		}

		if (m_ppd->hOldBitmDoubleBuffer)
			SelectObject(m_ppd->hDCDoubleBuffer, m_ppd->hOldBitmDoubleBuffer);
		if (m_ppd->hDCDoubleBuffer)
			DeleteDC(m_ppd->hDCDoubleBuffer);
		m_ppd->hOldBitmDoubleBuffer = NULL;

		if (m_ppd->hBitmDoubleBuffer)
			DeleteObject(m_ppd->hBitmDoubleBuffer);
		m_ppd->hBitmDoubleBuffer = NULL;
	}

	// Nachricht schicken
	SendMessage(m_hWndTbl, MTM_PaintDone, 0, 0);

	return 0;
}

// ProcessHotkey
BOOL MTBLSUBCLASS::ProcessHotkey(UINT uKeyMsg, WORD wParam, LONG lParam)
{
	// Cell-Mode und Hotkey eines permanenten Drop-Down-Buttons gedrückt?
	if (m_bCellMode && wParam == VK_DOWN)
	{
		HWND hWndFocusCol = NULL;
		long lFocusRow = TBL_Error;
		if (GetEffFocusCell(lFocusRow, hWndFocusCol))
		{
			if (MustShowDDBtn(hWndFocusCol, lFocusRow, TRUE, TRUE) && (GetAsyncKeyState(VK_MENU) & 0x8000))
			{
				Ctd.TblSetFocusCell(m_hWndTbl, lFocusRow, hWndFocusCol, 0, -1);
				if (!MTblIsListOpen(m_hWndTbl, hWndFocusCol))
				{
					HWND hWnd = GetFocus();
					if (hWnd != m_hWndTbl)
						SendMessage(hWnd, uKeyMsg, wParam, lParam);
				}
				return TRUE;
			}
		}
	}

	// Cell-Mode und Hotkey eines permanenten Buttons gedrückt?
	if (m_bCellMode && m_iBtnKey && m_iBtnKey == wParam)
	{
		HWND hWndFocusCol = NULL;
		long lFocusRow = TBL_Error;
		if (GetEffFocusCell(lFocusRow, hWndFocusCol))
		{
			if (MustShowBtn(hWndFocusCol, lFocusRow, TRUE, TRUE) && IsButtonHotkeyPressed())
			{
				CMTblBtn *pbtn = GetEffCellBtn(hWndFocusCol, lFocusRow);
				if (!(pbtn && pbtn->IsDisabled()))
				{
					PostMessage(m_hWndTbl, MTM_BtnClick, WPARAM(hWndFocusCol), LPARAM(lFocusRow));
					return TRUE;
				}
			}
		}
	}

	return FALSE;
}


// QueryCellFlags
BOOL MTBLSUBCLASS::QueryCellFlags(long lRow, HWND hWndCol, DWORD dwFlags)
{
	CMTblRow * pRow = m_Rows->GetRowPtr(lRow);
	if (pRow)
	{
		CMTblCol * pCol = pRow->m_Cols->Find(hWndCol);
		if (pCol)
		{
			return pCol->QueryFlags(dwFlags);
		}
	}

	return FALSE;
}


// QueryColumnFlags
BOOL MTBLSUBCLASS::QueryColumnFlags(HWND hWndCol, DWORD dwFlags)
{
	CMTblCol * pCol = m_Cols->Find(hWndCol);
	if (pCol)
		return pCol->QueryFlags(dwFlags);

	return FALSE;
}


// QueryColumnHdrFlags
BOOL MTBLSUBCLASS::QueryColumnHdrFlags(HWND hWndCol, DWORD dwFlags)
{
	CMTblCol * pCol = m_Cols->FindHdr(hWndCol);
	if (pCol)
		return pCol->QueryFlags(dwFlags);

	return FALSE;
}


// QueryRowFlagImgFlags
BOOL MTBLSUBCLASS::QueryRowFlagImgFlags(DWORD dwRowFlag, DWORD dwFlags)
{
	if (dwRowFlag)
	{
		RowFlagImgMap::iterator it;
		it = m_RowFlagImages->find(dwRowFlag);
		if (it != m_RowFlagImages->end())
		{
			BOOL bUserDefined = (*it).second.second->GetHandle() != NULL;
			if (bUserDefined)
				return (*it).second.second->QueryFlags(dwFlags);
			else
				return (*it).second.first->QueryFlags(dwFlags);
		}
	}

	return FALSE;
}


// QueryRowFlags
BOOL MTBLSUBCLASS::QueryRowFlags(long lRow, DWORD dwFlags)
{
	return m_Rows->QueryRowFlags(lRow, dwFlags);
}


// QueryVisRange
BOOL MTBLSUBCLASS::QueryVisRange(long &lRowFrom, long &lRowTo, BOOL bSplitRows /*= FALSE*/)
{
	lRowFrom = lRowTo = TBL_Error;

	if (bSplitRows && !m_bSplitted)
		return FALSE;

	if (!bSplitRows && m_pCounter->GetRowHeights() <= 0 && !QueryFlags(MTBL_FLAG_NO_COLUMN_HEADER))
	{
		return Ctd.TblQueryVisibleRange(m_hWndTbl, &lRowFrom, &lRowTo);
	}

	CMTblMetrics tm;
	tm.Get(m_hWndTbl, TRUE);

	int iScrPosV = GetScrollPosV(&tm, bSplitRows);
	if (iScrPosV < 0)
		return FALSE;

	long lRow = m_Rows->GetVisPosRow(iScrPosV, bSplitRows);
	if (lRow == TBL_Error)
		return FALSE;

	long lRowTop = bSplitRows ? tm.m_SplitRowsTop : tm.m_RowsTop;
	long lRowsBottom = bSplitRows ? tm.m_SplitRowsBottom : tm.m_RowsBottom;

	lRowFrom = lRowTo = lRow;

	long lLastValidPos, lUpperBound;
	CMTblRow **RowArr = m_Rows->GetRowArray(bSplitRows ? -1 : 0, &lUpperBound, &lLastValidPos);

	CMTblRow *pRow;
	long lHt, lRowBottom;
	long lPos = bSplitRows ? m_Rows->GetSplitRowPos(lRow) : lRow;
	for (;; ++lPos)
	{
		pRow = lPos <= lLastValidPos ? RowArr[lPos] : NULL;
		if (pRow && !pRow->IsVisible())
			continue;

		if (!pRow)
			lHt = tm.m_RowHeight;
		else
		{
			lHt = pRow->GetHeight();
			if (lHt == 0)
				lHt = tm.m_RowHeight;
		}
		lRowBottom = lRowTop + (lHt - 1);

		if (lRowBottom > lRowsBottom)
			break;
		else
			lRowTo = bSplitRows ? m_Rows->GetSplitPosRow(lPos) : lPos;

		lRowTop = lRowBottom + 1;
	}

	return TRUE;
}


// RegisterBtnClass
BOOL MTBLSUBCLASS::RegisterBtnClass()
{
	WNDCLASSEX wc;

	wc.cbSize = sizeof(WNDCLASSEX);

	if (!GetClassInfoEx(HINSTANCE(ghInstance), MTBL_BTN_CLASS, &wc))
	{
		wc.style = CS_OWNDC;
		wc.lpfnWndProc = BtnWndProc;
		wc.cbClsExtra = 0;
		wc.cbWndExtra = 4;
		wc.hInstance = HINSTANCE(ghInstance);
		wc.hIcon = NULL;
		wc.hCursor = LoadCursor(NULL, IDC_ARROW);
		wc.hbrBackground = NULL;
		wc.lpszMenuName = NULL;
		wc.lpszClassName = MTBL_BTN_CLASS;
		wc.hIconSm = NULL;

		return RegisterClassEx(&wc);
	}
	else
		return TRUE;
}


// ResetFocusData
BOOL MTBLSUBCLASS::ResetFocusData()
{
	m_hWndColFocus = NULL;
	m_pRowFocus = NULL;

	if (!IsRectEmpty(&m_rFocus))
		InvalidateRect(m_hWndTbl, &m_rFocus, FALSE);

	SetRectEmpty(&m_rFocus);

	return TRUE;
}


// ResetTip
void MTBLSUBCLASS::ResetTip()
{
	if (IsWindow(m_hWndTip))
		DestroyWindow(m_hWndTip);

	m_pTip = NULL;
	m_pColHdrGrpTipObj = NULL;

	m_bTipShowing = FALSE;
	m_bCornerTip = FALSE;

	m_iTipType = 0;
	m_iiTipTime = _I64_MAX;

	m_hWndTip = m_hWndBtnTipCol = m_hWndCellTipCol = m_hWndColHdrTipCol = m_hWndTipBtn = NULL;
	m_lRowHdrTipRow = m_lBtnTipRow = m_lCellTipRow = TBL_Error;
}


// RowFromPoint
long MTBLSUBCLASS::RowFromPoint(POINT pt)
{
	DWORD dwFlags;
	HWND hWndCol;
	long lRow;

	if (Ctd.TblObjectsFromPoint(m_hWndTbl, pt.x, pt.y, &lRow, &hWndCol, &dwFlags))
	{
		return lRow;
	}

	return TBL_Error;
}


// SetBtnWnd
BOOL MTBLSUBCLASS::SetBtnWndData(CMTblBtn *pBtn, int iHeight)
{
	if (!(pBtn && IsWindow(m_hWndBtn)))
		return FALSE;

	LPMTBLBTN pbOld = (LPMTBLBTN)GetWindowLongPtr(m_hWndBtn, GWLP_USERDATA);
	if (pbOld)
		delete pbOld;

	LPMTBLBTN pb = new (nothrow)MTBLBTN;
	if (pb)
	{
		memset(pb, 0, sizeof(MTBLBTN));
		pb->hWndTbl = m_hWndTbl;
		pb->hImage = pBtn->m_hImage;
		pb->dwFlags = pBtn->m_dwFlags;
		pb->iHeight = iHeight;
	}
	SetWindowLongPtr(m_hWndBtn, GWLP_USERDATA, (LONG_PTR)pb);

	return TRUE;
}


// SetCellFlags
BOOL MTBLSUBCLASS::SetCellFlags(long lRow, HWND hWndCol, CMTblMetrics *ptm, DWORD dwFlags, BOOL bSet, BOOL bInvalidate, DWORD *pdwDiff /*=NULL*/)
{
	// Zeile ermitteln
	CMTblRow * pRow = m_Rows->EnsureValid(lRow);
	if (!pRow)
		return FALSE;

	// Zelle ermitteln
	CMTblCol * pCol = pRow->m_Cols->EnsureValid(hWndCol);
	if (!pCol)
		return FALSE;

	// Behandlung MTBL_CELL_FLAG_SELECTED
	if ((dwFlags & MTBL_CELL_FLAG_SELECTED) && bSet && !(pCol->m_dwFlags & MTBL_CELL_FLAG_SELECTED))
	{
		if (SendMessage(m_hWndTbl, MTM_QueryNoCellSelection, (WPARAM)hWndCol, lRow) != 0)
			dwFlags ^= MTBL_CELL_FLAG_SELECTED;
	}

	// Flags setzen
	DWORD dwDiff = pCol->SetFlags(dwFlags, bSet);
	if (pdwDiff)
		*pdwDiff = dwDiff;

	// Read-Only-Status sicherstellen
	if (dwDiff & MTBL_CELL_FLAG_READ_ONLY)
		EnsureFocusCellReadOnly();

	// Neu zeichnen?
	if (dwDiff && bInvalidate)
		InvalidateCell(hWndCol, lRow, ptm, FALSE);

	return TRUE;
}

// SetCellIndent
// Setzt Einrückungs-Ebene einer Zelle
BOOL MTBLSUBCLASS::SetCellIndentLevel(long lRow, HWND hWndCol, long lIndentLevel, CMTblMetrics *ptm, BOOL bInvalidate)
{
	// Zeile ermitteln
	CMTblRow * pRow = m_Rows->EnsureValid(lRow);
	if (!pRow)
		return FALSE;

	// Zelle ermitteln
	CMTblCol * pCell = pRow->m_Cols->EnsureValid(hWndCol);
	if (!pCell)
		return FALSE;

	// Einrückungs-Ebene setzen
	pCell->m_IndentLevel = max(lIndentLevel, 0);

	// Neu zeichnen?
	if (bInvalidate)
		InvalidateCell(hWndCol, lRow, ptm, FALSE);

	return TRUE;
}


// SetCellRangeFlags
BOOL MTBLSUBCLASS::SetCellRangeFlags(LPMTBLCELLRANGE pcr, DWORD dwFlags, BOOL bSet, BOOL bInvalidate)
{
	if (!pcr)
		return FALSE;

	MTBLCELLRANGE &cr = *pcr;

	// Spaltenpositionen
	int iColPosFrom = Ctd.TblQueryColumnPos(cr.hWndColFrom);
	int iColPosTo = Ctd.TblQueryColumnPos(cr.hWndColTo);

	// Durchlaufen...
	CMTblMetrics tm(m_hWndTbl);
	HWND hWndCol;
	int iColPos;
	long lRow = cr.lRowFrom;

	// Sichtbare Spalten ermitteln
	HWND hWndaVisCols[MTBL_MAX_COLS];
	for (iColPos = 1; hWndCol = Ctd.TblGetColumnWindow(m_hWndTbl, iColPos, COL_GetPos); iColPos++)
	{
		if (SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Visible))
			hWndaVisCols[iColPos] = hWndCol;
		else
			hWndaVisCols[iColPos] = FALSE;
	}

	while (lRow <= cr.lRowTo)
	{
		if (IsRow(lRow) && !Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Hidden))
		{
			iColPos = iColPosFrom;
			while (iColPos <= iColPosTo)
			{
				hWndCol = hWndaVisCols[iColPos];
				if (hWndCol)
					SetCellFlags(lRow, hWndCol, &tm, dwFlags, bSet, bInvalidate);

				iColPos++;
			}
		}

		if (!Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, ROW_Hidden))
			break;
	}

	return TRUE;
}


// SetColumnTitle
BOOL MTBLSUBCLASS::SetColumnTitle(HWND hWndCol, LPCTSTR lpsTitle)
{
	tstring sTitle(_T(""));
	if (lpsTitle)
		sTitle = lpsTitle;

	CMTblColHdrGrp *pGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
	if (pGrp)
	{
		int iLines = pGrp->GetNrOfTextLines();
		if (iLines > 0)
		{
			// Anhängende CRLFs entfernen
			int iTrCRLF = GetTrailingCRLFCount(sTitle.c_str());
			if (iTrCRLF > 0)
				sTitle = sTitle.substr(0, sTitle.size() - (iTrCRLF * 2));

			// CRLFs anhängen
			int iNeeded;
			if (pGrp->QueryFlags(MTBL_CHG_FLAG_NOCOLHEADERS))
				iNeeded = iLines - GetCRLFCount(sTitle.c_str()) - 1;
			else
				iNeeded = iLines;

			TCHAR szCRLF[3] = { 13, 10, 0 };
			for (int i = 0; i < iNeeded; ++i)
			{
				sTitle += szCRLF;
			}
		}
	}

	return Ctd.TblSetColumnTitle(hWndCol, (LPTSTR)sTitle.c_str());
}


// SetCurrentCellType
BOOL MTBLSUBCLASS::SetCurrentCellType(HWND hWndCol, CMTblCellType *pct)
{
	if (!pct)
		return FALSE;

	CMTblCol *pCol = m_Cols->Find(hWndCol);
	if (!pCol)
		return FALSE;

	if (!pct->SetColumn(hWndCol))
		return FALSE;

	pCol->m_pCellTypeCurr = pct;

	if (pct->IsDropDownList())
		SubClassExtEdit(hWndCol, NULL);

	return TRUE;
}


// SetDropDownListContext
BOOL MTBLSUBCLASS::SetDropDownListContext(HWND hWndCol, long lRow)
{
	CMTblCol *pCol = m_Cols->EnsureValid(hWndCol);
	if (!pCol)
		return FALSE;

	CMTblCellType *pct = NULL;

	if (lRow != TBL_Error)
	{
		if (!IsRow(lRow))
			return FALSE;

		if (m_Rows->IsRowDelAdj(lRow))
			return FALSE;

		CMTblRow *pRow = m_Rows->EnsureValid(lRow);
		if (!pRow)
			return FALSE;

		CMTblCol *pCell = pRow->m_Cols->EnsureValid(hWndCol);
		if (!pCell)
			return FALSE;

		pct = pCell->m_pCellType;
		if (!pct || !pct->IsDropDownList())
		{
			pct = pCol->m_pCellType;
			lRow = TBL_Error;
		}
	}
	else
		pct = pCol->m_pCellType;


	if (pct && pct->IsDropDownList())
	{
		pCol->SetDropDownListContext(lRow);
		//pct->SetColumn( hWndCol );
		//SubClassExtEdit( hWndCol, NULL );
		if (!SetCurrentCellType(hWndCol, pct))
			return FALSE;

		// Listeintrag selektieren
		long lCurrContext = Ctd.TblQueryContext(m_hWndTbl);
		Ctd.TblSetContext(m_hWndTbl, lRow);

		HSTRING hsText = 0;
		if (Ctd.FmtFieldToStr(hWndCol, &hsText, FALSE))
		{
			HWND hWndExtEdit = (HWND)SendMessage(hWndCol, CIM_GetExtEdit, 0, 0);
			if (hWndExtEdit)
			{
				tstring tsText;
				if (Ctd.HStrToTStr(hsText, tsText) && !tsText.empty())
				{
					long lIndex = SendMessage(hWndExtEdit, LB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)tsText.c_str());
					if (lIndex != LB_ERR)
						SendMessage(hWndExtEdit, LB_SETCURSEL, (WPARAM)lIndex, 0);
				}
			}
		}

		if (hsText != 0)
			Ctd.HStringUnRef(hsText);

		Ctd.TblSetContext(m_hWndTbl, lCurrContext);

		return TRUE;
	}
	else
		return FALSE;
}


// SetEffCellPropOrder
BOOL MTBLSUBCLASS::SetEffCellPropDetOrder(DWORD dwOrder)
{
	BOOL bSet = FALSE;
	switch (dwOrder)
	{
	case MTECPD_ORDER_CELL_COL_ROW:
		m_byEffCellPropCellPos = 0;
		m_byEffCellPropColPos = 1;
		m_byEffCellPropRowPos = 2;
		bSet = TRUE;
		break;
	case MTECPD_ORDER_CELL_ROW_COL:
		m_byEffCellPropCellPos = 0;
		m_byEffCellPropColPos = 2;
		m_byEffCellPropRowPos = 1;
		bSet = TRUE;
		break;
	case MTECPD_ORDER_COL_CELL_ROW:
		m_byEffCellPropCellPos = 1;
		m_byEffCellPropColPos = 0;
		m_byEffCellPropRowPos = 2;
		bSet = TRUE;
		break;
	case MTECPD_ORDER_COL_ROW_CELL:
		m_byEffCellPropCellPos = 2;
		m_byEffCellPropColPos = 0;
		m_byEffCellPropRowPos = 1;
		bSet = TRUE;
		break;
	case MTECPD_ORDER_ROW_CELL_COL:
		m_byEffCellPropCellPos = 1;
		m_byEffCellPropColPos = 2;
		m_byEffCellPropRowPos = 0;
		bSet = TRUE;
		break;
	case MTECPD_ORDER_ROW_COL_CELL:
		m_byEffCellPropCellPos = 2;
		m_byEffCellPropColPos = 1;
		m_byEffCellPropRowPos = 0;
		bSet = TRUE;
		break;
	}

	if (bSet)
		m_dwEffCellPropDetOrder = dwOrder;

	return bSet;
}


// SetFlags
BOOL MTBLSUBCLASS::SetFlags(DWORD dwFlags, BOOL bSet)
{
	m_dwFlags = (bSet ? (m_dwFlags | dwFlags) : ((m_dwFlags & dwFlags) ^ m_dwFlags));
	return TRUE;
}


// SetFocusCell
BOOL MTBLSUBCLASS::SetFocusCell(long lRow, HWND hWndCol)
{
	if (m_pCounter->GetNoFocusUpdate() > 0)
		return FALSE;

	if (!IsWindow(hWndCol) || !Ctd.TblQueryColumnFlags(hWndCol, COL_Visible))
		return FALSE;

	if (!IsRow(lRow))
		return FALSE;

	if (!m_bCellMode)
		return Ctd.TblSetFocusCell(m_hWndTbl, lRow, hWndCol, 0, -1);

	if (!CanSetFocusToCell(hWndCol, lRow))
		return FALSE;

	HWND hWndColFocus = NULL;
	long lRowFocus = TBL_Error;
	if (Ctd.TblQueryFocus(m_hWndTbl, &lRowFocus, &hWndColFocus))
	{
		if (lRowFocus != lRow)
		{
			m_pCounter->IncNoFocusUpdate();
			BOOL bOk = Ctd.TblSetFocusRow(m_hWndTbl, lRow);
			m_pCounter->DecNoFocusUpdate();

			if (!bOk)
				return FALSE;
		}
		else if (hWndColFocus && hWndColFocus != hWndCol)
		{
			if (!Ctd.TblKillEdit(m_hWndTbl))
				return FALSE;
		}
	}

	return UpdateFocus(lRow, hWndCol, UF_REDRAW_AUTO, UF_SELCHANGE_NONE, UF_FLAG_SCROLL);
}


// SetItemHighlighted
BOOL MTBLSUBCLASS::SetItemHighlighted(CMTblItem &item, BOOL bSet)
{
	ItemList::iterator it = FindHighlightedItem(item);

	if (bSet && it == m_pHLIL->end())
	{
		// Hinzufügen...
		m_pHLIL->push_back(item);
	}
	else if (!bSet && it != m_pHLIL->end())
	{
		// Löschen
		m_pHLIL->erase(it);
	}

	return TRUE;
}


// SetMergeRowFlags
BOOL MTBLSUBCLASS::SetMergeRowFlags(LPMTBLMERGEROW pmr, DWORD dwFlags, BOOL bSet, BOOL bInternal /*= FALSE*/, long lRowExcept /*= TBL_Error*/)
{
	if (!pmr)
		return FALSE;

	CMTblRow *pRow;
	long lRow = pmr->lRowFrom;

	if (lRow != lRowExcept)
	{
		if (bInternal)
		{
			pRow = m_Rows->EnsureValid(lRow);
			pRow->SetInternalFlags(dwFlags, bSet);
		}
		else
		{
			Ctd.TblSetRowFlags(m_hWndTbl, lRow, dwFlags, bSet);
			//InvalidateRow( lRow, FALSE );
		}
	}

	while (FindNextRow(&lRow, 0, ROW_Hidden, pmr->lRowTo))
	{
		if (lRow != lRowExcept)
		{
			if (bInternal)
			{
				pRow = m_Rows->EnsureValid(lRow);
				pRow->SetInternalFlags(dwFlags, bSet);
			}
			else
			{
				Ctd.TblSetRowFlags(m_hWndTbl, lRow, dwFlags, bSet);
				//InvalidateRow( lRow, FALSE );
			}
		}
	}

	return TRUE;
}


// SetNodeImages
BOOL MTBLSUBCLASS::SetNodeImages(HANDLE hImg, HANDLE hImgExp)
{
	if (hImg && !MImgIsValidHandle(hImg))
		return FALSE;
	if (hImgExp && !MImgIsValidHandle(hImgExp))
		return FALSE;

	m_hImgNode = hImg;
	m_hImgNodeExp = hImgExp;

	return TRUE;
}


// SetOwnerDrawnItem
BOOL MTBLSUBCLASS::SetOwnerDrawnItem(HWND hWndParam, long lParam, int iType, BOOL bSet)
{
	CMTblCol *pCol;
	CMTblColHdrGrp *pColHdrGrp;
	CMTblRow *pRow;

	switch (iType)
	{
	case MTBL_ODI_CELL:
		pRow = m_Rows->EnsureValid(lParam);
		if (!pRow) return FALSE;
		pCol = pRow->m_Cols->EnsureValid(hWndParam);
		if (!pCol) return FALSE;
		if (pCol->IsOwnerdrawn() != bSet)
		{
			pCol->SetOwnerdrawn(bSet);
			InvalidateCell(hWndParam, lParam, NULL, FALSE);
		}
		break;
	case MTBL_ODI_COLUMN:
		pCol = m_Cols->EnsureValid(hWndParam);
		if (!pCol) return FALSE;
		if (pCol->IsOwnerdrawn() != bSet)
		{
			pCol->SetOwnerdrawn(bSet);
			InvalidateCol(hWndParam, FALSE);
		}
		break;
	case MTBL_ODI_COLHDR:
		pCol = m_Cols->EnsureValidHdr(hWndParam);
		if (!pCol) return FALSE;
		if (pCol->IsOwnerdrawn() != bSet)
		{
			pCol->SetOwnerdrawn(bSet);
			InvalidateColHdr(FALSE);
		}
		break;
	case MTBL_ODI_COLHDRGRP:
		pColHdrGrp = (CMTblColHdrGrp*)lParam;
		if (!m_ColHdrGrps->IsValidGrp(pColHdrGrp)) return FALSE;
		if (pColHdrGrp->IsOwnerdrawn() != bSet)
		{
			pColHdrGrp->SetOwnerdrawn(bSet);
			InvalidateColHdr(FALSE);
		}
		break;
	case MTBL_ODI_ROW:
		pRow = m_Rows->EnsureValid(lParam);
		if (!pRow) return FALSE;
		if (pRow->IsOwnerdrawn() != bSet)
		{
			pRow->SetOwnerdrawn(bSet);
			InvalidateRow(lParam, FALSE, TRUE);
		}
		break;
	case MTBL_ODI_ROWHDR:
		pRow = m_Rows->EnsureValid(lParam);
		if (!pRow) return FALSE;
		if (pRow->IsHdrOwnerdrawn() != bSet)
		{
			pRow->SetHdrOwnerdrawn(bSet);
			InvalidateRowHdr(lParam, FALSE, TRUE);
		}
		break;
	case MTBL_ODI_CORNER:
		if (m_bCornerOwnerdrawn != bSet)
		{
			m_bCornerOwnerdrawn = bSet;
			InvalidateCorner(FALSE);
		}
		break;
	default:
		return FALSE;
	}

	return TRUE;
}

// SetPaintColData
void MTBLSUBCLASS::SetPaintColData(HWND hWndCol, int iColPos, BOOL bPaint, BOOL bLocked, BOOL bLastLocked, BOOL bFirstUnlocked, BOOL bFirstVis, BOOL bEmpty, BOOL bRowHdr, LPPOINT pptLog, LPPOINT pptPhys, LPMTBLPAINTCOL pPaintCol)
{
	BOOL bPseudo = (bEmpty || bRowHdr);

	// Pseudo setzen
	pPaintCol->bPseudo = bPseudo;

	// Handle setzen
	pPaintCol->hWnd = hWndCol;

	// Position setzen
	pPaintCol->iPos = iColPos;

	// Spalten-ID setzen
	if (bPseudo)
		pPaintCol->iID = 0;
	else
		pPaintCol->iID = Ctd.TblQueryColumnID(hWndCol);

	// Zeichnen?
	pPaintCol->bPaint = bPaint;

	// Verankert setzen
	pPaintCol->bLocked = bLocked;

	// Letzte Verankerte setzen
	pPaintCol->bLastLocked = bLastLocked;

	// Erste nicht Verankerte setzen
	pPaintCol->bFirstUnlocked = bFirstUnlocked;

	// Erste Sichtbare?
	pPaintCol->bFirstVisible = bFirstVis;

	// Letzte Sichtbare?
	pPaintCol->bLastVisible = FALSE;

	// Leere?
	pPaintCol->bEmpty = bEmpty;

	// Row-Header?
	pPaintCol->bRowHdr = bRowHdr;

	// "Ungerade" Linie
	pPaintCol->bOddLine = (pptPhys->y - m_ppd->tm.m_RowHeaderRight) % 2 != 0;

	// Markiert setzen
	if (bPseudo)
		pPaintCol->bSelected = FALSE;
	else
		pPaintCol->bSelected = SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Selected);

	// Word Wrap setzen
	if (bPseudo)
		pPaintCol->bWordWrap = FALSE;
	else
		pPaintCol->bWordWrap = SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_MultilineCell);

	// Invisible setzen
	if (bPseudo)
		pPaintCol->bInvisible = FALSE;
	else
		pPaintCol->bInvisible = SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_Invisible);

	// IndicateOverflow setzen
	if (bPseudo)
		pPaintCol->bIndicateOverflow = FALSE;
	else
		pPaintCol->bIndicateOverflow = SendMessage(hWndCol, TIM_QueryColumnFlags, 0, COL_IndicateOverflow);

	// Ausrichtung setzen
	if (bPseudo)
		pPaintCol->lJustify = DT_LEFT;
	else
		pPaintCol->lJustify = GetColJustify(hWndCol);

	// Zellentyp setzen
	if (bPseudo)
		pPaintCol->iCellType = COL_CellType_Standard;
	else
		//pPaintCol->iCellType = Ctd.TblGetCellType( hWndCol );
		pPaintCol->iCellType = GetEffCellType(TBL_Error, hWndCol);

	// Check-Box?
	if (pPaintCol->iCellType == COL_CellType_CheckBox)
	{
		HSTRING hsCheckedVal = 0, hsUncheckedVal = 0;
		if (Ctd.TblQueryCheckBoxColumn(pPaintCol->hWnd, &pPaintCol->dwCBFlags, &hsCheckedVal, &hsUncheckedVal))
		{
			Ctd.HStrToTStr(hsCheckedVal, pPaintCol->sCBCheckedVal, TRUE);
			Ctd.HStrToTStr(hsUncheckedVal, pPaintCol->sCBUncheckedVal, TRUE);
		}
	}

	// Koordinaten setzen
	pPaintCol->ptX.x = pptPhys->x;
	pPaintCol->ptX.y = pptPhys->y;

	// Logische Koordinaten setzen
	pPaintCol->ptLogX.x = pptLog->x;
	pPaintCol->ptLogX.y = pptLog->y;

	// Sichtbare Koordinaten setzen
	if (bRowHdr)
	{
		pPaintCol->ptVisX.x = pPaintCol->ptX.x;
		pPaintCol->ptVisX.y = min(pPaintCol->ptX.y, m_ppd->tm.m_ClientRect.right);
	}
	else
	{
		if (bLocked)
		{
			pPaintCol->ptVisX.x = max(pPaintCol->ptX.x, m_ppd->tm.m_LockedColumnsLeft);
			pPaintCol->ptVisX.y = min(pPaintCol->ptX.y, m_ppd->tm.m_LockedColumnsRight);
		}
		else
		{
			pPaintCol->ptVisX.x = max(pPaintCol->ptX.x, m_ppd->tm.m_ColumnsLeft);
			pPaintCol->ptVisX.y = min(pPaintCol->ptX.y, m_ppd->tm.m_ColumnsRight);
		}
	}

	// Spalte setzen
	if (!hWndCol)
		pPaintCol->pCol = NULL;
	else
		pPaintCol->pCol = m_Cols->Find(hWndCol);

	// Header-Gruppe setzen
	if (!hWndCol)
		pPaintCol->pHdrGrp = NULL;
	else
		pPaintCol->pHdrGrp = m_ColHdrGrps->GetGrpOfCol(hWndCol);
}

// SetParentRow
BOOL MTBLSUBCLASS::SetParentRow(long lRow, long lParentRow, DWORD dwFlags)
{
	// Gültige Zeile?
	if (!IsRow(lRow)) return FALSE;

	// Aktuellen Zustand merken
	long lOldParentRow = m_Rows->GetParentRow(lRow);
	BOOL bNewParentIsExpanded = FALSE;
	if (lParentRow != TBL_Error)
		bNewParentIsExpanded = m_Rows->QueryRowFlags(lParentRow, MTBL_ROW_ISEXPANDED);

	// Setzen
	BOOL bOk = m_Rows->SetParentRow(lRow, lParentRow, dwFlags & MTSPR_AS_ORIG);

	// Nacharbeiten
	if (bOk && lParentRow != lOldParentRow)
	{
		LockPaint();

		if (bNewParentIsExpanded || lParentRow == TBL_Error)
		{
			if (!m_Rows->QueryRowFlags(lRow, MTBL_ROW_HIDDEN))
				Ctd.TblSetRowFlags(m_hWndTbl, lRow, ROW_Hidden, FALSE);
		}
		else if (lParentRow != TBL_Error)
		{
			CollapseRow(m_Rows->GetRowPtr(lRow), 0);
			Ctd.TblSetRowFlags(m_hWndTbl, lRow, ROW_Hidden, TRUE);
		}

		UnlockPaint();

		if (dwFlags & MTSPR_REDRAW)
		{
			InvalidateRect(m_hWndTbl, NULL, FALSE);
		}
	}

	return bOk;
}


// SetRowFlagBitmap
BOOL MTBLSUBCLASS::SetRowFlagBitmap(DWORD dwFlag, int iBitmRes, DWORD dwTranspClr)
{
	if (!dwFlag)
		return FALSE;

	HRSRC hRes = FindResource(HMODULE(ghInstance), MAKEINTRESOURCE(iBitmRes), RT_BITMAP);
	if (!hRes)
		return FALSE;

	HANDLE hImg = MImgLoadFromCResource(HMODULE(ghInstance), hRes, MIMG_TYPE_BMP, 0);
	if (!hImg)
		return FALSE;

	MImgSetFlags(hImg, MIMG_FLAG_PRIVATE, TRUE);

	if (dwTranspClr == MIMG_COLOR_UNDEF)
		dwTranspClr = MImgGetPixelColor(hImg, 0, 0);
	MImgSetTranspColor(hImg, dwTranspClr);

	return SetRowFlagImg(dwFlag, hImg, FALSE);
}

// SetRowFlagImg
BOOL MTBLSUBCLASS::SetRowFlagImg(DWORD dwFlag, HANDLE hImage, BOOL bUserDefined, DWORD dwFlags /*=0*/)
{
	if (!(dwFlag)) return FALSE;

	RowFlagImgMap::iterator it;
	it = m_RowFlagImages->find(dwFlag);
	if (it != m_RowFlagImages->end())
	{
		CMTblImg *pImg = bUserDefined ? (*it).second.second : (*it).second.first;
		pImg->SetHandle(hImage, m_pCounter);
		pImg->ClearFlags();
		if (CMTblImg::IsValidImageHandle(hImage))
		{
			pImg->SetFlags(dwFlags, TRUE);
			pImg->SetFlags(MTSI_REDRAW, FALSE);
		}
	}
	else
	{
		pair<CMTblImg*, CMTblImg*> val;
		val.first = new CMTblImg;
		val.second = new CMTblImg;

		val.first->SetHandle(bUserDefined ? NULL : hImage, m_pCounter);
		val.second->SetHandle(bUserDefined ? hImage : NULL, m_pCounter);
		CMTblImg *pImg = bUserDefined ? val.second : val.first;
		pImg->ClearFlags();
		if (CMTblImg::IsValidImageHandle(hImage))
		{
			pImg->SetFlags(dwFlags, TRUE);
			pImg->SetFlags(MTSI_REDRAW, FALSE);
		}

		pair<RowFlagImgMap::iterator, bool> ret;
		ret = m_RowFlagImages->insert(RowFlagImgMap::value_type(dwFlag, val));
		if (!ret.second)
			return FALSE;
	}

	return TRUE;
}


// SetRowHdrFlags
BOOL MTBLSUBCLASS::SetRowHdrFlags(WORD wFlagsSet, BOOL bSet)
{
	BOOL bOk = FALSE;
	HSTRING hsText = 0;
	HWND hWndCol;
	int iWid;
	WORD wFlags = 0;

	if (Ctd.TblQueryRowHeader(m_hWndTbl, &hsText, 254, &iWid, &wFlags, &hWndCol))
	{
		wFlags = (bSet ? (wFlags | wFlagsSet) : ((wFlags & wFlagsSet) ^ wFlags));
		bOk = Ctd.TblDefineRowHeader(m_hWndTbl, Ctd.HStrToLPTSTR(hsText), iWid, wFlags, hWndCol);
	}

	if (hsText)
		Ctd.HStringUnRef(hsText);

	return bOk;
}


// SetTreeFlags
BOOL MTBLSUBCLASS::SetTreeFlags(DWORD dwFlags, BOOL bSet)
{
	BOOL bNoSelInvNodesBefore = QueryTreeFlags(MTBL_TREE_FLAG_NOSELINV_NODES);
	BOOL bNoSelInvLinesBefore = QueryTreeFlags(MTBL_TREE_FLAG_NOSELINV_LINES);

	m_dwTreeFlags = (bSet ? (m_dwTreeFlags | dwFlags) : ((m_dwTreeFlags & dwFlags) ^ m_dwTreeFlags));

	BOOL bNoSelInvNodes = QueryTreeFlags(MTBL_TREE_FLAG_NOSELINV_NODES);
	if (bNoSelInvNodesBefore != bNoSelInvNodes)
	{
		long lInc = (bNoSelInvNodes && !bNoSelInvNodesBefore) ? 1 : -1;
		m_pCounter->IncNoSelInv(lInc);
	}

	BOOL bNoSelInvLines = QueryTreeFlags(MTBL_TREE_FLAG_NOSELINV_LINES);
	if (bNoSelInvLinesBefore != bNoSelInvLines)
	{
		long lInc = (bNoSelInvLines && !bNoSelInvLinesBefore) ? 1 : -1;
		m_pCounter->IncNoSelInv(lInc);
	}

	return TRUE;
}


// Sort
BOOL MTBLSUBCLASS::Sort(HARRAY hWndaCols, HARRAY naFlags, long lMinRow, long lMaxRow, long lFromLevel /*= -1*/, long lToLevel /*= -1 */, long lParentRow /*= TBL_Error */)
{
	const TCHAR cFill = 1;

	BOOL bAnyMergeCells = FALSE, bAnyMergeRows = FALSE, bError = FALSE, bKeepRowMerge = FALSE, bRestoreTree = FALSE, bMergeRowRange = FALSE, bTreeInvalid = FALSE, bTreeSorted = FALSE;
	CMTblRow *pParentRow = NULL, *pValRow = NULL;
	DATETIME dt;
	HANDLE hdl;
	HWND hWndCol = NULL;
	HSTRING hsColText = 0, hsQuery = 0;
	int i, iCols = 0, iFlags, iSortrows, iSortStrLen = 0, k;
	long lContext, lLen, lLevel, lLower, lLowerS, lRow, lRows = 0, lUpper, lUpperS, lValue, lValueRow;
	long *pl = NULL;
	LPTSTR lpColText = NULL, lpDate = NULL;
	NUMBER n;
	RowRange rmr;
	SORTCOLINFO *psci = NULL;
	SORTROW *psr = NULL;

	// Sind überhaupt Zeilen zum Sortieren da?
	if (!Ctd.TblAnyRows(m_hWndTbl, 0, 0)) return FALSE;

	// Kontext ermitteln
	lContext = Ctd.TblQueryContext(m_hWndTbl);

	// Parameter checken bzw. anpassen
	if (lMinRow < 0 || lMinRow == TBL_Error || lMinRow == TBL_MinRow)
		lMinRow = 0;

	if (lMaxRow < 0 || lMaxRow == TBL_Error || lMaxRow == TBL_MaxRow)
	{
		lRow = -1;
		while (Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, 0))
		{
			if (!Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_InCache))
				SendMessage(m_hWndTbl, TIM_FetchRow, 0, lRow);
		}
		lMaxRow = lRow;
		Ctd.TblSetContextEx(m_hWndTbl, lContext);
	}

	if (lMaxRow <= lMinRow)
		return FALSE;

	if (lToLevel >= 0 && lToLevel < lFromLevel)
		lToLevel = lFromLevel;

	if (lParentRow < 0)
		return FALSE;
	if (lParentRow != TBL_Error)
	{
		pParentRow = m_Rows->GetRowPtr(lParentRow);
		if (!pParentRow)
			return FALSE;
	}

	if (Ctd.ArrayIsEmpty(hWndaCols)) return FALSE;

	// Row-Merge zusammenhalten?
	bKeepRowMerge = QueryFlags(MTBL_FLAG_SORT_KEEP_ROWMERGE);

	// Tree wiederherstellen?
	bRestoreTree = QueryFlags(MTBL_FLAG_SORT_RESTORE_TREE);

	// Evtl. Merge-Zellen/Zeilen neu ermitteln, Min./Max. anpassen
	if (bKeepRowMerge)
	{
		bAnyMergeRows = (m_pCounter->GetMergeRows() > 0);
		bAnyMergeCells = (m_pCounter->GetMergeCells() > 0);
		if (bAnyMergeRows || bAnyMergeCells)
		{
			if (bAnyMergeRows)
			{
				m_bMergeRowsInvalid = TRUE;
				GetMergeRows();
			}

			if (bAnyMergeCells)
			{
				m_bMergeCellsInvalid = TRUE;
				GetMergeCells();
			}

			if (GetMergeRowRange(lMinRow, GCMRR_UP, rmr) && rmr.first < lMinRow)
				lMinRow = rmr.first;

			if (GetMergeRowRange(lMaxRow, GCMRR_DOWN, rmr) && rmr.second > lMaxRow)
				lMaxRow = rmr.second;
		}
	}

	// Grenzen des Spaltenarrays ermitteln
	if (!Ctd.ArrayGetLowerBound(hWndaCols, 1, &lLower)) return FALSE;
	if (!Ctd.ArrayGetUpperBound(hWndaCols, 1, &lUpper)) return FALSE;

	// Grenzen des Sortierflagarrays ermitteln
	if (!Ctd.ArrayGetLowerBound(naFlags, 1, &lLowerS)) return FALSE;
	if (!Ctd.ArrayGetUpperBound(naFlags, 1, &lUpperS)) return FALSE;

	// Sortierzeilen-Array erzeugen
	iSortrows = lMaxRow - lMinRow + 1;
	psr = (SORTROW*)new SORTROW[iSortrows];
	if (!psr) return FALSE;
	memset(psr, 0, sizeof(SORTROW) * iSortrows);

	// Zeilenarray erzeugen
	pl = (long*)new long[iSortrows];
	if (!pl) return FALSE;
	memset(pl, 0, sizeof(long) * iSortrows);

	// Daten über letzte Sortierung zurücksetzen
	m_LastSortCols->clear();
	m_LastSortFlags->clear();

	// Spalteninformationen sammeln
	psci = (SORTCOLINFO*)new SORTCOLINFO[lUpper - lLower + 2]; // Eine Pseudospalte für Zeilen-Nr.
	if (!psci) return FALSE;
	memset(psci, 0, sizeof(SORTCOLINFO) * (lUpper - lLower + 1));

	for (i = lLower; i <= lUpper; i++)
	{
		// Handle der Spalte ermitteln
		if (!Ctd.MDArrayGetHandle(hWndaCols, &hdl, i))
		{
			bError = TRUE;
			break;
		}
		hWndCol = (HWND)hdl;

		if (IsWindow(hWndCol) && Ctd.ParentWindow(hWndCol) == m_hWndTbl)
		{
			// Zur Liste der letzten Sortierspalten hinzufügen
			m_LastSortCols->push_back(hWndCol);

			// Flags ermitteln
			if (!Ctd.MDArrayGetNumber(naFlags, &n, i + (lLowerS - lLower)))
			{
				bError = TRUE;
				break;
			}
			if (!Ctd.CvtNumberToInt(&n, &iFlags))
			{
				bError = TRUE;
				break;
			}

			// Zur Liste der letzten Sortierflags hinzufügen
			m_LastSortFlags->push_back(iFlags);

			// Handle setzen
			psci[iCols].hWnd = hWndCol;

			// ID setzen
			psci[iCols].iID = Ctd.TblQueryColumnID(hWndCol);

			// Datentyp setzen
			if (iFlags & MTS_DT_STRING || iFlags & MTS_DT_CUSTOM_STRING)
			{
				psci[iCols].iDataType = DT_String;
				if (iFlags & MTS_DT_CUSTOM_STRING)
					psci[iCols].bQuery = TRUE;
			}
			else if (iFlags & MTS_DT_NUMBER || iFlags & MTS_DT_CUSTOM)
			{
				psci[iCols].iDataType = DT_Number;
				if (iFlags & MTS_DT_CUSTOM)
					psci[iCols].bQuery = TRUE;
			}
			else if (iFlags & MTS_DT_DATETIME)
				psci[iCols].iDataType = DT_DateTime;
			else
				psci[iCols].iDataType = Ctd.GetDataType(hWndCol);

			// Datentyp der Spalte setzen
			psci[iCols].iColDataType = Ctd.GetDataType(hWndCol);

			// Sortierfolge setzen
			psci[i].iSortOrder = iFlags & MTS_DESC ? TBL_SortDecreasing : TBL_SortIncreasing;

			// Groß/Kleinschreibung ignorieren?
			psci[i].bIgnoreCase = iFlags & MTS_IGNORECASE ? TRUE : FALSE;

			// Stringsort verwenden?
			psci[i].bStringSort = iFlags & MTS_STRINGSORT ? TRUE : FALSE;

			// Spaltenzähler erhöhen
			iCols++;
		}
	}

	// Pseudospalte für Zeilennummer setzen
	i = lUpper + 1;

	// Handle setzen
	psci[iCols].hWnd = NULL;

	// ID setzen
	psci[iCols].iID = 0;

	// Datentyp setzen
	psci[iCols].iDataType = DT_Number;

	// Datentyp der Spalte setzen
	psci[iCols].iColDataType = DT_Number;

	// Sortierfolge setzen
	psci[i].iSortOrder = TBL_SortIncreasing;

	// Groß/Kleinschreibung ignorieren?
	psci[i].bIgnoreCase = TRUE;

	// Stringsort verwenden?
	psci[i].bStringSort = FALSE;

	// Spaltenzähler erhöhen
	iCols++;

	if (!bError)
	{
		// Tabelle durchlaufen
		lRow = lMinRow - 1;
		while (Ctd.TblFindNextRow(m_hWndTbl, &lRow, 0, 0))
		{
			if (lRow > lMaxRow) break;

			lValueRow = lRow;

			// Merge-Zeilen/Zellen
			if (bKeepRowMerge && (bAnyMergeRows || bAnyMergeCells))
			{
				// Merge-Row-Bereich ermitteln
				if (!bMergeRowRange || (bMergeRowRange && lRow > rmr.second))
					bMergeRowRange = GetMergeRowRange(lRow, GCMRR_DOWN, rmr);

				if (bMergeRowRange)
					lValueRow = rmr.first;
			}

			// Zeile nicht sortieren?
			pValRow = m_Rows->GetRowPtr(lValueRow);
			if (pValRow)
			{
				if (pValRow->QueryFlags(MTBL_ROW_NO_SORT))
					continue;

				if (pParentRow && !pValRow->IsDescendantOf(pParentRow))
					continue;

				lLevel = pValRow->GetLevel();
				if (lFromLevel >= 0 && lLevel < lFromLevel)
					continue;
				if (lToLevel >= 0 && lLevel > lToLevel)
					continue;

				if (pValRow->IsChild() || pValRow->IsParent())
					bTreeInvalid = TRUE;
			}

			// Kontext setzen
			if (!Ctd.TblSetContextEx(m_hWndTbl, lValueRow))
			{
				bError = TRUE;
				break;
			}

			// Sortierwerte-Array erzeugen
			psr[lRows].pSortVals = new CMTblSortVal[iCols];
			psr[lRows].pSortColInfos = psci;
			psr[lRows].lCols = iCols;
			psr[lRows].lRow = lRow;

			// Sortierwerte setzen
			for (i = 0; i < iCols; i++)
			{
				if (psci[i].hWnd)
				{
					if (psci[i].bQuery)
					{
						// Wert ermitteln + setzen
						if (psci[i].iDataType == DT_Number)
						{
							lValue = SendMessage(m_hWndTbl, MTM_QuerySortValue, WPARAM(psci[i].hWnd), lValueRow);
							Ctd.CvtLongToNumber(lValue, &n);

							bError = !psr[lRows].pSortVals[i].SetNumber(n);
						}
						else if (psci[i].iDataType == DT_String)
						{
							lpColText = NULL;
							lValue = SendMessage(m_hWndTbl, MTM_QuerySortString, WPARAM(psci[i].hWnd), lValueRow);
							if (lValue > 0)
							{
								hsQuery = Ctd.NumberToHString(lValue);
								lpColText = Ctd.StringGetBuffer(hsQuery, &lLen);
							}								

							bError = !psr[lRows].pSortVals[i].SetString(lpColText ? lpColText : _T(""));

							if (lValue > 0)
								Ctd.HStrUnlock(hsQuery);
						}

						if (bError)
							break;
					}
					else
					{
						// Spaltentext ermitteln
						if ((psci[i].iColDataType == DT_DateTime && psci[i].iDataType == DT_DateTime) ||
							(psci[i].iColDataType == DT_Number && psci[i].iDataType == DT_Number)
							)
						{
							if (!Ctd.FmtFieldToStr(psci[i].hWnd, &hsColText, FALSE))
							{
								bError = TRUE;
								break;
							}
						}
						else
						{
							if (!Ctd.TblGetColumnText(m_hWndTbl, psci[i].iID, &hsColText))
							{
								bError = TRUE;
								break;
							}
						}

						lpColText = Ctd.StringGetBuffer(hsColText, &lLen);
						if (lpColText)
						{
							// Werte setzen
							switch (psci[i].iDataType)
							{
							case DT_String:
							case DT_LongString:
								bError = !psr[lRows].pSortVals[i].SetString(lpColText);
								break;

							case DT_Number:
								n = Ctd.StrToNumber(lpColText);
								bError = !psr[lRows].pSortVals[i].SetNumber(n);
								break;

							case DT_DateTime:
								dt = Ctd.StrToDate(lpColText);
								bError = !psr[lRows].pSortVals[i].SetDate(dt);
								break;
							}
						}
						else
							bError = TRUE;

						Ctd.HStrUnlock(hsColText);
						
						if (bError)
							break;
					}
				}
				else // Pseudospalte Zeilennummer
				{
					// Zeilennummer setzen
					if (!Ctd.CvtLongToNumber(lRow, &n))
					{
						bError = TRUE;
						break;
					}
					if (!psr[lRows].pSortVals[i].SetNumber(n))
					{
						bError = TRUE;
						break;
					}
				}
			}

			if (bError)
				break;

			// Zeile merken
			pl[lRows] = lRow;

			// Anzahl Zeilen erhöhen
			lRows++;
		}
	}

	// Sortieren
	if (!bError && lRows > 1)
	{
		qsort(psr, lRows, sizeof(SORTROW), MTblSortCompare);

		Ctd.TblKillEdit(m_hWndTbl);

		m_pCounter->IncNoFocusUpdate();
		for (i = 0; i < lRows; ++i)
		{
			if (psr[i].lRow != pl[i])
			{
				if (!SwapRows(psr[i].lRow, pl[i]))
					bError = TRUE;
				for (k = i; k < lRows; ++k)
				{
					if (psr[k].lRow == pl[i])
					{
						psr[k].lRow = psr[i].lRow;
						break;
					}
				}
			}
		}
		m_pCounter->DecNoFocusUpdate();
		UpdateFocus();
	}

	// Evtl. Tree wiederherstellen
	if (bRestoreTree && bTreeInvalid)
	{
		DWORD dwMode;
		if (m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP)
			dwMode = MTST_BOTTOMUP;
		else
			dwMode = MTST_TOPDOWN;
		if (!(bTreeSorted = SortTree(dwMode, TRUE)))
			bError = TRUE;
	}

	// Kontext wiederherstellen
	Ctd.TblSetContextEx(m_hWndTbl, lContext);

	// Tabelle neu zeichnen
	InvalidateRect(m_hWndTbl, NULL, FALSE);
	UpdateWindow(m_hWndTbl);

	// Aufräumen
	if (psr)
	{
		for (i = 0; i < lRows; i++)
		{
			if (psr[i].pSortVals)
				delete[]psr[i].pSortVals;
		}
		delete[]psr;
	}

	if (pl)
		delete[]pl;
	if (psci)
		delete[]psci;
	if (hsColText)
		Ctd.HStringUnRef(hsColText);

	// Fättich!
	return !bError;
}

/*{
const char cFill = 1;

BOOL bError = FALSE;
CMTblRow *pParentRow = NULL, *pRow = NULL;
DATETIME dt;
HANDLE hdl;
HWND hWndCol = NULL;
HSTRING hsColText = 0;
int i, iCols = 0, iFlags, iSortrows, iSortStrLen = 0, k;
long lColTextLen, lContext, lLevel, lLower, lLowerS, lRow, lRows = 0, lUpper, lUpperS, lValue;
long *pl = NULL;
LPTSTR lpColText = NULL, lpDate = NULL;
NUMBER n;
SORTCOLINFO *psci = NULL;
SORTROW *psr = NULL;

// Sind überhaupt Zeilen zum Sortieren da?
if ( !Ctd.TblAnyRows( m_hWndTbl, 0, 0 ) ) return FALSE;

// Kontext ermitteln
lContext = Ctd.TblQueryContext( m_hWndTbl );

// Parameter checken bzw. anpassen
if ( lMinRow < 0 || lMinRow == TBL_Error || lMinRow == TBL_MinRow )
lMinRow = 0;

if ( lMaxRow < 0 || lMaxRow == TBL_Error || lMaxRow == TBL_MaxRow )
{
lRow = -1;
while ( Ctd.TblFindNextRow( m_hWndTbl, &lRow, 0, 0 ) )
{
if ( !Ctd.TblQueryRowFlags( m_hWndTbl, lRow, ROW_InCache ) )
SendMessage( m_hWndTbl, TIM_FetchRow, 0, lRow );
}
lMaxRow = lRow;
Ctd.TblSetContextEx( m_hWndTbl, lContext );
}

if ( lMaxRow < lMinRow )
return FALSE;

if ( lToLevel >= 0 && lToLevel < lFromLevel )
lToLevel = lFromLevel;

if ( lParentRow < 0 )
return FALSE;
if ( lParentRow != TBL_Error )
{
pParentRow = m_Rows->GetRowPtr( lParentRow );
if ( !pParentRow )
return FALSE;
}

if ( Ctd.ArrayIsEmpty( hWndaCols ) ) return FALSE;

// Grenzen des Spaltenarrays ermitteln
if ( !Ctd.ArrayGetLowerBound( hWndaCols, 1, &lLower ) ) return FALSE;
if ( !Ctd.ArrayGetUpperBound( hWndaCols, 1, &lUpper ) ) return FALSE;

// Grenzen des Sortierflagarrays ermitteln
if ( !Ctd.ArrayGetLowerBound( naFlags, 1, &lLowerS ) ) return FALSE;
if ( !Ctd.ArrayGetUpperBound( naFlags, 1, &lUpperS ) ) return FALSE;

// Sortierzeilen-Array erzeugen
iSortrows = lMaxRow - lMinRow + 1;
psr = (SORTROW*)new SORTROW[iSortrows];
if ( !psr ) return FALSE;
memset( psr, 0, sizeof(SORTROW) * iSortrows );

// Zeilenarray erzeugen
pl = (long*)new long[iSortrows];
if ( !pl ) return FALSE;
memset( pl, 0, sizeof(long) * iSortrows );

// Daten über letzte Sortierung zurücksetzen
m_LastSortCols->clear( );
m_LastSortFlags->clear( );

// Spalteninformationen sammeln
psci = (SORTCOLINFO*)new SORTCOLINFO[lUpper - lLower + 1];
if ( !psci ) return FALSE;
memset( psci, 0, sizeof(SORTCOLINFO) * ( lUpper - lLower + 1 ) );

for ( i = lLower; i <= lUpper; i++ )
{
// Handle der Spalte ermitteln
if ( !Ctd.MDArrayGetHandle( hWndaCols, &hdl, i ) )
{
bError = TRUE;
break;
}
hWndCol = (HWND)hdl;

if ( IsWindow( hWndCol ) && Ctd.ParentWindow( hWndCol ) == m_hWndTbl )
{
// Zur Liste der letzten Sortierspalten hinzufügen
m_LastSortCols->push_back( hWndCol );

// Flags ermitteln
if ( !Ctd.MDArrayGetNumber( naFlags, &n, i + ( lLowerS - lLower ) ) )
{
bError = TRUE;
break;
}
if ( !Ctd.CvtNumberToInt( &n, &iFlags ) )
{
bError = TRUE;
break;
}

// Zur Liste der letzten Sortierflags hinzufügen
m_LastSortFlags->push_back( iFlags );

// Handle setzen
psci[iCols].hWnd = hWndCol;

// ID setzen
psci[iCols].iID = Ctd.TblQueryColumnID( hWndCol );

// Datentyp setzen
if ( iFlags & MTS_DT_STRING || iFlags & MTS_DT_CUSTOM_STRING )
{
psci[iCols].iDataType = DT_String;
if ( iFlags & MTS_DT_CUSTOM_STRING )
psci[iCols].bQuery = TRUE;
}
else if ( iFlags & MTS_DT_NUMBER || iFlags & MTS_DT_CUSTOM )
{
psci[iCols].iDataType = DT_Number;
if ( iFlags & MTS_DT_CUSTOM )
psci[iCols].bQuery = TRUE;
}
else if ( iFlags & MTS_DT_DATETIME )
psci[iCols].iDataType = DT_DateTime;
else
psci[iCols].iDataType = Ctd.GetDataType( hWndCol );

// Datentyp der Spalte setzen
psci[iCols].iColDataType = Ctd.GetDataType( hWndCol );

// Sortierfolge setzen
psci[i].iSortOrder = iFlags & MTS_DESC ? TBL_SortDecreasing : TBL_SortIncreasing;

// Groß/Kleinschreibung ignorieren?
psci[i].bIgnoreCase = iFlags & MTS_IGNORECASE ? TRUE : FALSE;

// Stringsort verwenden?
psci[i].bStringSort = iFlags & MTS_STRINGSORT ? TRUE : FALSE;

// Spaltenzähler erhöhen
iCols++;
}
}

if ( !bError )
{
// Tabelle durchlaufen
lRow = lMinRow - 1;
while( Ctd.TblFindNextRow( m_hWndTbl, &lRow, 0, 0 ) )
{
if ( lRow > lMaxRow ) break;

// Zeile nicht sortieren?
pRow = m_Rows->GetRowPtr( lRow );
if ( pRow )
{
if ( pRow->QueryFlags( MTBL_ROW_NO_SORT ) )
continue;

if ( pParentRow && !pRow->IsDescendantOf( pParentRow ) )
continue;

lLevel = pRow->GetLevel( );
if ( lFromLevel >= 0 && lLevel < lFromLevel )
continue;
if ( lToLevel >= 0 && lLevel > lToLevel )
continue;
}

// Kontext setzen
if ( !Ctd.TblSetContextEx( m_hWndTbl, lRow ) )
{
bError = TRUE;
break;
}

// Sortierwerte-Array erzeugen
psr[lRows].pSortVals = new CMTblSortVal[iCols];
psr[lRows].pSortColInfos = psci;
psr[lRows].lCols = iCols;
psr[lRows].lRow = lRow;

// Sortierstring bauen
for ( i = 0; i < iCols; i++ )
{
if ( psci[i].bQuery )
{
// Wert ermitteln + setzen
if ( psci[i].iDataType == DT_Number )
{
lValue = SendMessage( m_hWndTbl, MTM_QuerySortValue, WPARAM( psci[i].hWnd ), lRow );
Ctd.CvtLongToNumber( lValue, &n );

bError = !psr[lRows].pSortVals[i].SetNumber( n );
}
else if ( psci[i].iDataType == DT_String )
{
lpColText = NULL;
lValue = SendMessage( m_hWndTbl, MTM_QuerySortString, WPARAM( psci[i].hWnd ), lRow );
if ( lValue > 0 )
lpColText = Ctd.StringGetBuffer( Ctd.NumberToHString( lValue ), &lColTextLen );

bError = !psr[lRows].pSortVals[i].SetString( lpColText ? lpColText : "" );
}

if ( bError )
break;
}
else
{
// Spaltentext ermitteln
if ( ( psci[i].iColDataType == DT_DateTime && psci[i].iDataType == DT_DateTime ) ||
( psci[i].iColDataType == DT_Number && psci[i].iDataType == DT_Number )
)
{
if ( !Ctd.FmtFieldToStr( psci[i].hWnd, &hsColText, FALSE ) )
{
bError = TRUE;
break;
}
}
else
{
if ( !Ctd.TblGetColumnText( m_hWndTbl, psci[i].iID, &hsColText ) )
{
bError = TRUE;
break;
}
}

lpColText = Ctd.StringGetBuffer( hsColText, &lColTextLen );
if ( !lpColText )
{
bError = TRUE;
break;
}
lColTextLen--;

// Werte setzen
switch ( psci[i].iDataType )
{
case DT_String:
case DT_LongString:
bError = !psr[lRows].pSortVals[i].SetString( lpColText );
break;

case DT_Number:
n = Ctd.StrToNumber( lpColText );
bError = !psr[lRows].pSortVals[i].SetNumber( n );
break;

case DT_DateTime:
dt = Ctd.StrToDate( lpColText );
bError = !psr[lRows].pSortVals[i].SetDate( dt );
break;

if ( bError )
break;
}
}
}

if ( bError )
break;

// Zeile merken
pl[lRows] = lRow;

// Anzahl Zeilen erhöhen
lRows++;
}
}

// Sortieren
if ( !bError && lRows > 1 )
{
qsort( psr, lRows, sizeof(SORTROW), MTblSortCompare );
for ( i = 0; i < lRows; ++i )
{
if ( psr[i].lRow != pl[i] )
{
if ( !MTblSwapRows( m_hWndTbl, psr[i].lRow, pl[i], 0 ) )
bError = TRUE;
for ( k = i; k < lRows; ++k )
{
if ( psr[k].lRow == pl[i] )
{
psr[k].lRow = psr[i].lRow;
break;
}
}
}
}
}

// Kontext wiederherstellen
Ctd.TblSetContextEx( m_hWndTbl, lContext );

// Tabelle neu zeichnen
Ctd.InvalidateWindow( m_hWndTbl );
Ctd.UpdateWindow( m_hWndTbl );

// Aufräumen
if ( psr )
{
for ( i = 0; i < lRows; i++ )
{
if ( psr[i].pSortVals )
delete []psr[i].pSortVals;
}
delete []psr;
}

if ( pl )
delete []pl;
if ( psci )
delete []psci;
if ( hsColText )
Ctd.HStringUnRef( hsColText );

// Fättich!
return !bError;
}*/




// SortTree
BOOL MTBLSUBCLASS::SortTree(DWORD dwFlags, BOOL bNoRedraw /*= FALSE*/)
{
	// Sind überhaupt Zeilen da?
	if (!Ctd.TblAnyRows(m_hWndTbl, 0, 0)) return TRUE;

	LPMTBLMERGEROW pmr;

	// Bereich ermitteln
	long lUpperBound, lLastValidPos;
	CMTblRow **pRows = m_Rows->GetRowArray(0, &lUpperBound, &lLastValidPos);

	BOOL bFound = FALSE;
	CMTblRow *pRow;
	// Minimum suchen
	long lMin;
	for (lMin = 0; lMin <= lLastValidPos; lMin++)
	{
		if (pRow = pRows[lMin])
		{
			if (!pRow->IsDelAdj() && (pRow->IsParent() || pRow->IsChild()))
			{
				pmr = FindMergeRow(lMin, FMR_ANY);
				if (pmr && pmr->lRowFrom < lMin)
					lMin = pmr->lRowFrom;

				bFound = TRUE;
				break;
			}
		}
	}
	if (!bFound)
		return TRUE;

	// Maximum suchen
	bFound = FALSE;
	long lMax;
	for (lMax = lLastValidPos; lMax >= 0; lMax--)
	{
		if (pRow = pRows[lMax])
		{
			if (!pRow->IsDelAdj() && (pRow->IsParent() || pRow->IsChild()))
			{
				pmr = FindMergeRow(lMax, FMR_ANY);
				if (pmr && pmr->lRowTo > lMax)
					lMax = pmr->lRowTo;

				bFound = TRUE;
				break;
			}
		}
	}
	if (!bFound)
		return FALSE;

	// Nochmal checken
	if (lMin >= lMax)
		return FALSE;

	// Anzahl der zu sortierenden Zeilen ermitteln
	int iSortrows = lMax - lMin + 1;

	// Sortierzeilen-Array erzeugen
	typedef struct
	{
		long lSort;
		long lRow;
	}
	SORTROW;

	SORTROW * psr = (SORTROW*)new (nothrow)SORTROW[iSortrows];
	if (!psr) return FALSE;
	memset(psr, 0, sizeof(SORTROW) * iSortrows);

	// Besucht-Flags aller normalen Zeilen zurücksetzen
	if (!m_Rows->SetRowsInternalFlags(FALSE, ROW_IFLAG_SORTVISIT, FALSE)) return FALSE;

	// Noch nicht gefetchte Zeilen im Bereich fetchen
	long lRow = lMin;
	while (lRow <= lMax)
	{
		if (!Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_InCache))
			SendMessage(m_hWndTbl, TIM_FetchRow, 0, lRow);
		lRow++;
	}

	// Tabelle durchlaufen
	BOOL bBottomUp = (dwFlags == MTST_BOTTOMUP);
	BOOL bBottomUpCur = (m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP) != 0;
	BOOL bError = FALSE, bVisited;
	long lInc, lMergeRow, lMergeRowBegin, lMergeRowEnd, lRows = 0, lSort = 0;
	stack<CMTblRow*> st;

	lRow = bBottomUp ? lMax + 1 : lMin - 1;
	while (bBottomUp ? FindPrevRow(&lRow, 0, 0, lMin) : FindNextRow(&lRow, 0, 0, lMax))
	{
		// Ist das eine Zeile auf oberstem Level?
		if (m_Rows->GetRowLevel(lRow) == 0)
		{
			// Row-Pointer ermitteln
			pRow = m_Rows->EnsureValid(lRow);
			if (!pRow)
			{
				bError = TRUE;
				break;
			}

			// Auf geht's
			while (pRow)
			{
				// Schon besucht?
				bVisited = pRow->QueryInternalFlags(ROW_IFLAG_SORTVISIT);

				if (bVisited)
				{
					// Nächste Kindzeile ermitteln
					pRow = m_Rows->GetNextChildRow(pRow, bBottomUp);

					if (!pRow && !st.empty())
					{
						// Mit vorheriger Zeile weitermachen
						pRow = st.top();
						st.pop();
					}
				}
				else
				{
					// Besucht-Flag setzen
					pRow->SetInternalFlags(ROW_IFLAG_SORTVISIT, TRUE);
					// Auf Stack legen
					st.push(pRow);
					// Teil eines Zeilenverbundes?
					pmr = FindMergeRow(pRow->GetNr(), FMR_ANY);
					// Sortierwerte setzen
					if (!pmr)
					{
						psr[lRows].lRow = pRow->GetNr();
						psr[lRows].lSort = ++lSort;
						++lRows;
					}
					else
					{
						// Wenn erste Zeile ( je nach Sortierrichtung ) des Zellenverbundes...
						if ((!bBottomUp && pRow->GetNr() == pmr->lRowFrom) || (bBottomUp && pRow->GetNr() == pmr->lRowTo))
						{
							// Zeilen des Verbundes durchlaufen und Sortierwerte setzen
							if (!bBottomUp)
							{
								lMergeRowBegin = pmr->lRowFrom;
								lMergeRowEnd = pmr->lRowTo + 1;
								lInc = 1;
							}
							else
							{
								lMergeRowBegin = pmr->lRowTo;
								lMergeRowEnd = pmr->lRowFrom - 1;
								lInc = -1;
							}

							for (lMergeRow = lMergeRowBegin; lMergeRow != lMergeRowEnd; lMergeRow += lInc)
							{
								if (!m_Rows->IsRowDelAdj(lMergeRow) && m_Rows->GetRowLevel(lMergeRow) == m_Rows->GetRowLevel(lMergeRowBegin))
								{
									psr[lRows].lRow = lMergeRow;
									psr[lRows].lSort = ++lSort;
									++lRows;
								}
							}
						}
					}
					// Erste Kindzeile ermitteln
					pRow = m_Rows->GetFirstChildRow(pRow, bBottomUp);

					if (!pRow && !st.empty())
					{
						// Mit vorheriger Zeile weitermachen
						pRow = st.top();
						st.pop();
					}
				}
			}
		}
	}

	if (!bError)
	{
		// Sortierarray nach Zeilennummern sortieren
		SORTROW srTmp;
		long i, j, k;

		k = lRows / 2;
		while (k > 0)
		{
			for (i = 0; i < lRows - k; i++)
			{
				j = i;
				while (j >= 0 && psr[j].lRow > psr[j + k].lRow)
				{
					srTmp.lRow = psr[j].lRow;
					srTmp.lSort = psr[j].lSort;
					psr[j].lRow = psr[j + k].lRow;
					psr[j].lSort = psr[j + k].lSort;
					psr[j + k].lRow = srTmp.lRow;
					psr[j + k].lSort = srTmp.lSort;

					if (j > k)
						j -= k;
					else
						j = 0;
				}
			}
			k = k / 2;
		}

		// Tabelle sortieren
		long lTmp;
		k = lRows / 2;

		Ctd.TblKillEdit(m_hWndTbl);

		m_pCounter->IncNoFocusUpdate();
		while (k > 0)
		{
			for (i = 0; i < lRows - k; i++)
			{
				j = i;
				while (j >= 0 && (bBottomUp ? psr[j].lSort < psr[j + k].lSort : psr[j].lSort > psr[j + k].lSort))
				{
					lTmp = psr[j].lSort;
					psr[j].lSort = psr[j + k].lSort;
					psr[j + k].lSort = lTmp;

					if (!SwapRows(psr[j].lRow, psr[j + k].lRow))
						bError = TRUE;

					if (j > k)
						j -= k;
					else
						j = 0;
				}
			}
			k = k / 2;
		}
		m_pCounter->DecNoFocusUpdate();
		UpdateFocus();

		// Tree-Flags anpassen
		if (bBottomUp)
			m_dwTreeFlags |= MTBL_TREE_FLAG_BOTTOM_UP;
		else
			m_dwTreeFlags = (m_dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP) ^ m_dwTreeFlags;

		// Tabelle neu zeichnen
		if (!bNoRedraw)
		{
			InvalidateRect(m_hWndTbl, NULL, FALSE);
			UpdateWindow(m_hWndTbl);
		}
	}

	// Aufräumen
	delete[]psr;

	return !bError;
}


// StartInternalTimer
void MTBLSUBCLASS::StartInternalTimer(BOOL bProcessImmediate /*= FALSE*/)
{
	StopInternalTimer();
	m_uiTimer = SetTimer(m_hWndTbl, 999, 100, (TIMERPROC)TimerProc);
	if (m_uiTimer && bProcessImmediate)
		OnInternalTimer(0, 0);
}


// StopInternalTimer
void MTBLSUBCLASS::StopInternalTimer()
{
	if (m_uiTimer && KillTimer(m_hWndTbl, m_uiTimer))
		m_uiTimer = 0;
}


// SubClassCol
BOOL MTBLSUBCLASS::SubClassCol(HWND hWndCol)
{
	if (!IsWindow(hWndCol)) return FALSE;

	if (IsColSubClassed(hWndCol)) return TRUE;

	LPMTBLCOLSUBCLASS psc = new MTBLCOLSUBCLASS;
	if (!psc) return FALSE;
	ZeroMemory(psc, sizeof(MTBLCOLSUBCLASS));

	psc->hWndTbl = m_hWndTbl;
	psc->wpOld = (WNDPROC)SetWindowLongPtr(hWndCol, GWLP_WNDPROC, (LONG_PTR)ColWndProc);

	SubClassMap::value_type val(hWndCol, (LPVOID)psc);
	m_scm->insert(val);

	return TRUE;
}


// SubClassCols
BOOL MTBLSUBCLASS::SubClassCols()
{
	HWND hWndCol = NULL;
	//while( MTblFindNextCol( m_hWndTbl, &hWndCol, 0, 0 ) )
	while (FindNextCol(&hWndCol, 0, 0))
		SubClassCol(hWndCol);

	return TRUE;
}


// SubClassContainer
BOOL MTBLSUBCLASS::SubClassContainer(HWND hWnd)
{
	LPMTBLCONTAINERSUBCLASS psc = NULL;

	// Prüfung Handle
	if (!IsWindow(hWnd)) return FALSE;

	// Property schon gesetzt?
	if (IsContainerSubClassed(hWnd)) return TRUE;

	// Subclass-Struktur erzeugen
	psc = (LPMTBLCONTAINERSUBCLASS)new MTBLCONTAINERSUBCLASS;
	if (!psc) return FALSE;

	// Property setzen
	SetProp(hWnd, MTBL_CONTAINER_SUBCLASS_PROP, (HANDLE)psc);

	// Subclass-Struktur initialisieren
	memset(LPVOID(psc), 0, sizeof(MTBLCONTAINERSUBCLASS));

	// Fensterprozedur setzen
	psc->wpOld = (WNDPROC)SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)&ContainerWndProc);

	// Stil setzen
	SetWindowLongPtr(hWnd, GWL_STYLE, GetWindowLongPtr(hWnd, GWL_STYLE) | WS_CLIPCHILDREN);

	// Handle der Tabelle setzen
	psc->hWndTbl = m_hWndTbl;

	return TRUE;
}


BOOL MTBLSUBCLASS::SubClassExtEdit(HWND hWndCol, HWND hWndExtEdit)
{
	if (!hWndExtEdit)
		hWndExtEdit = (HWND)SendMessage(hWndCol, CIM_GetExtEdit, 0, 0);

	if (!IsWindow(hWndExtEdit))
		return FALSE;

	if (IsExtEditSubClassed(hWndExtEdit)) return TRUE;

	LPMTBLEXTEDITSUBCLASS psc = new MTBLEXTEDITSUBCLASS;
	if (!psc) return FALSE;
	ZeroMemory(psc, sizeof(MTBLEXTEDITSUBCLASS));

	psc->hWndCol = hWndCol;
	psc->wpOld = (WNDPROC)SetWindowLongPtr(hWndExtEdit, GWLP_WNDPROC, (LONG_PTR)ExtEditWndProc);
	SetProp(hWndExtEdit, _T("m_hWndTbl"), m_hWndTbl);

	SubClassMap::value_type val(hWndExtEdit, (LPVOID)psc);
	m_scm->insert(val);

	return TRUE;
}


// SubClassListClip
BOOL MTBLSUBCLASS::SubClassListClip(HWND hWnd)
{
	LPMTBLLISTCLIPSUBCLASS psc = NULL;

	// Prüfung Handle
	if (!IsWindow(hWnd)) return FALSE;

	// Property schon gesetzt?
	if (IsListClipSubClassed(hWnd)) return TRUE;

	// Subclass-Struktur erzeugen
	psc = (LPMTBLLISTCLIPSUBCLASS)new MTBLLISTCLIPSUBCLASS;
	if (!psc) return FALSE;

	// Property setzen
	SetProp(hWnd, MTBL_LISTCLIP_SUBCLASS_PROP, (HANDLE)psc);

	// Subclass-Struktur initialisieren
	memset(LPVOID(psc), 0, sizeof(MTBLLISTCLIPSUBCLASS));

	// Fensterprozedur setzen
	psc->wpOld = (WNDPROC)SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)ListClipWndProc);

	// Stil setzen
	SetWindowLongPtr(hWnd, GWL_STYLE, GetWindowLongPtr(hWnd, GWL_STYLE) | WS_CLIPCHILDREN);

	// Handle der Tabelle setzen
	psc->hWndTbl = m_hWndTbl;

	return TRUE;
}


// SubClassListClipCtrl
BOOL MTBLSUBCLASS::SubClassListClipCtrl(HWND hWnd)
{
	LPMTBLLISTCLIPCTRLSUBCLASS psc = NULL;

	// Prüfung Handle
	if (!IsWindow(hWnd)) return FALSE;

	// Property schon gesetzt?
	if (IsListClipCtrlSubClassed(hWnd)) return TRUE;

	// Subclass-Struktur erzeugen
	psc = (LPMTBLLISTCLIPCTRLSUBCLASS)new MTBLLISTCLIPCTRLSUBCLASS;
	if (!psc) return FALSE;

	// Property setzen
	SetProp(hWnd, MTBL_LISTCLIP_CTRL_SUBCLASS_PROP, (HANDLE)psc);

	// Subclass-Struktur initialisieren
	memset(LPVOID(psc), 0, sizeof(MTBLLISTCLIPCTRLSUBCLASS));

	// Fensterprozedur setzen
	psc->wpOld = (WNDPROC)SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)ListClipCtrlWndProc);

	// Handle der Tabelle setzen
	psc->hWndTbl = m_hWndTbl;

	return TRUE;
}


// SwapRows
BOOL MTBLSUBCLASS::SwapRows(long lRow1, long lRow2)
{
	MTBLSWAPROWS sr = { lRow1, lRow2 };
	return (BOOL)SendMessage(m_hWndTbl, TIM_SwapRows, 0, (LPARAM)&sr);
}


// TimerProc
void CALLBACK MTBLSUBCLASS::TimerProc(HWND hWnd, UINT uMsg, UINT uEvent, DWORD dwTime)
{
	// Subclass ermitteln
	LPMTBLSUBCLASS psc = MTblGetSubClass(hWnd);

	if (psc)
	{
		if (uEvent == psc->m_uiTimer)
		{
			psc->OnInternalTimer(0, 0);
		}
		else if (uEvent == psc->m_uiDragDropTimer)
		{
			psc->DragDropBegin();
		}
	}
}


// TimerWndProc
LRESULT CALLBACK MTBLSUBCLASS::TimerWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	// Subclass ermitteln
	LPMTBLSUBCLASS psc = (LPMTBLSUBCLASS)GetWindowLongPtr(hWnd, GWLP_USERDATA);

	long lRet;

	switch (uMsg)
	{
	case WM_CREATE:
		psc = (LPMTBLSUBCLASS)LPCREATESTRUCT(lParam)->lpCreateParams;
		SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)psc);
		if (psc)
			psc->m_uiTimer = SetTimer(hWnd, 0, 100, NULL);
		lRet = 0;
		break;
	case WM_DESTROY:
		if (psc && psc->m_uiTimer)
			KillTimer(hWnd, psc->m_uiTimer);
		lRet = 0;
		break;
	default:
		lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
	}

	return lRet;
}


// TipWndProc
LRESULT CALLBACK MTBLSUBCLASS::TipWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	// Subclass ermitteln
	LPMTBLSUBCLASS psc = (LPMTBLSUBCLASS)GetWindowLongPtr(hWnd, GWLP_USERDATA);

	long lRet;

	switch (uMsg)
	{
	case WM_CREATE:
		psc = (LPMTBLSUBCLASS)LPCREATESTRUCT(lParam)->lpCreateParams;
		SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)psc);
		lRet = 0;
		break;
	case WM_NCHITTEST:
		if (psc && psc->m_pTip && psc->m_pTip->QueryFlags(MTBL_TIP_FLAG_PERMEABLE))
			lRet = HTTRANSPARENT;
		else
			lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
		break;
	case WM_PAINT:
		if (psc)
			lRet = psc->OnTipPaint(hWnd, wParam, lParam);
		else
			lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
		break;
	case WM_PRINTCLIENT:
		if (psc)
			lRet = psc->OnTipPrintClient(hWnd, wParam, lParam);
		else
			lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
		break;
	case WM_SHOWWINDOW:
		if (psc)
			lRet = psc->OnTipShow(hWnd, wParam, lParam);
		else
			lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
	default:
		lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
	}

	return lRet;
}


// UnlockPaint
long MTBLSUBCLASS::UnlockPaint(void)
{
	--m_lLockPaint;
	return m_lLockPaint;
}

// UpdateFocus
BOOL MTBLSUBCLASS::UpdateFocus(long lRow /*=TBL_Error*/, HWND hWndCol /*=NULL*/, int iRedrawMode /*=UF_REDRAW_AUTO*/, int iSelChangeMode /*=UF_SELCHANGE_AUTO*/, DWORD dwFlags /*=0*/)
{
	static BOOL s_bNoFocusPaint = m_bNoFocusPaint;
	static BOOL s_bCellMode = m_bCellMode;

	if (m_pCounter->GetNoFocusUpdate() < 1)
	{
		m_pCounter->IncNoFocusUpdate();

		// "Alte" Werte merken
		HWND hWndColFocusOld = m_hWndColFocus;
		long lRowFocusOld = GetRowNrFocus();

		// Focus TD ermitteln
		HWND hWndFocusColTD = NULL;
		long lFocusRowTD = TBL_Error;
		Ctd.TblQueryFocus(m_hWndTbl, &lFocusRowTD, &hWndFocusColTD);

		// Wenn Zeile nicht definiert ist oder im Edit-Mode, Zellenfokus-Zeile = Fokus-Zeile TD
		if (lRow == TBL_Error || hWndFocusColTD)
		{
			lRow = lFocusRowTD;
		}

		if (!IsRow(lRow) || Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Hidden))
			lRow = TBL_Error;

		if (m_bCellMode && lRow != TBL_Error)
		{
			if (!hWndCol)
			{
				// Wenn im Edit-Mode, Fokus-Spalte TD verwenden
				if (hWndFocusColTD)
					hWndCol = hWndFocusColTD;
				else
				{
					// ... ansonsten aktuelle Spalte, wenn diese gültig und sichtbar ist...
					BOOL bCurColValid = m_hWndColFocus && IsWindow(m_hWndColFocus);
					if (bCurColValid && Ctd.TblQueryColumnFlags(m_hWndColFocus, COL_Visible))
						hWndCol = m_hWndColFocus;
					else
					{
						// ... ansonsten Spalte ermitteln
						HWND hWndFocusCol;
						hWndFocusCol = bCurColValid ? m_hWndColFocus : NULL;
						//if ( MTblFindNextCol( m_hWndTbl, &hWndFocusCol, COL_Visible, 0 ) )
						if (FindNextCol(&hWndFocusCol, COL_Visible, 0))
							hWndCol = hWndFocusCol;
						else if (bCurColValid)
						{
							hWndFocusCol = m_hWndColFocus;
							//if ( MTblFindPrevCol( m_hWndTbl, &hWndFocusCol, COL_Visible, 0 ) )
							if (FindPrevCol(&hWndFocusCol, COL_Visible, 0))
								hWndCol = hWndFocusCol;
						}
					}
				}
			}

			if (!IsWindow(hWndCol) || !Ctd.TblQueryColumnFlags(hWndCol, COL_Visible))
			{
				hWndCol = NULL;
				lRow = TBL_Error;
			}
		}
		else
			hWndCol = NULL;

		// Wird sich der Fokus ändern?
		BOOL bWillChange = lRow != lRowFocusOld || hWndCol != hWndColFocusOld;

		// Muss Nachricht gesendet werden?
		if (m_bCellMode && bWillChange && lRow != TBL_Error && hWndCol && !hWndFocusColTD)
		{
			HWND hWndColEffNew = hWndCol, hWndColEffOld = hWndColFocusOld;
			long lRowEffNew = lRow, lRowEffOld = lRowFocusOld;

			GetMasterCell(hWndColEffOld, lRowEffOld);
			GetMasterCell(hWndColEffNew, lRowEffNew);

			if (hWndColEffOld != hWndColEffNew || lRowEffOld != lRowEffNew)
			{
				if (!CanSetFocusToCell(hWndColEffNew, lRowEffNew))
				{
					hWndCol = hWndColFocusOld;
					lRow = lRowFocusOld;
				}
			}
		}

		// Evtl. scrollen
		BOOL bScrolled = FALSE;
		if (dwFlags & UF_FLAG_SCROLL)
		{
			BOOL bDoScroll = FALSE;
			HWND hWndScrollTo = hWndCol;
			long lRowScrollTo = lRow;

			if (m_bCellMode && hWndCol && hWndCol != m_hWndColFocus)
			{
				LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_SLAVE);
				if (pmc)
				{
					hWndScrollTo = pmc->hWndColFrom;
					lRowScrollTo = pmc->lRowFrom;
				}

				int iLeft, iTop, iRight, iBottom;
				int iRet = MTblGetCellRectEx(hWndScrollTo, lRowScrollTo, &iLeft, &iTop, &iRight, &iBottom, 0);

				if (iRet == MTGR_RET_INVISIBLE || iRet == MTGR_RET_PARTLY_VISIBLE)
					bDoScroll = TRUE;
			}
			else if (!m_bCellMode && lRow != TBL_Error && lRow != lRowFocusOld)
			{
				RECT r;
				int iRet = GetRowRect(lRow, &r, 0);

				if (iRet == MTGR_RET_INVISIBLE || iRet == MTGR_RET_PARTLY_VISIBLE)
					bDoScroll = TRUE;
			}
			if (bDoScroll)
			{
				LRESULT lRet = SendMessage(m_hWndTbl, TIM_QueryPixelOrigin, 0, 0);
				int iColOriginBefore = (int)LOWORD(lRet);
				int iRowOriginBefore = (int)HIWORD(lRet);
				if (Ctd.TblScroll(m_hWndTbl, lRowScrollTo, hWndScrollTo, TBL_AutoScroll))
				{
					lRet = SendMessage(m_hWndTbl, TIM_QueryPixelOrigin, 0, 0);
					int iColOriginAfter = (int)LOWORD(lRet);
					int iRowOriginAfter = (int)HIWORD(lRet);
					if (iColOriginBefore != iColOriginAfter || iRowOriginBefore != iRowOriginAfter)
						bScrolled = TRUE;
				}
			}
		}

		// Letztes Rechteck merken
		RECT rLast;
		CopyRect(&rLast, &m_rFocus);

		// Rechteck des aktuellen Fokus-Objekts ermitteln
		RECT rCur = rLast;
		if ((m_bScrolling || bScrolled) && !IsRectEmpty(&rLast))
		{
			LRESULT lRet = SendMessage(m_hWndTbl, TIM_QueryPixelOrigin, 0, 0);
			int iColOrigin = (int)LOWORD(lRet);
			int iRowOrigin = (int)HIWORD(lRet);

			if (iColOrigin != m_iColOriginBeforeScroll || iRowOrigin != m_iRowOriginBeforeScroll)
			{
				int iColDiff = (iColOrigin - m_iColOriginBeforeScroll);
				int iRowDiff = (iRowOrigin - m_iRowOriginBeforeScroll);

				rCur.left = rLast.left - iColDiff;
				rCur.top = rLast.top - iRowDiff;
				rCur.right = rLast.right - iColDiff;
				rCur.bottom = rLast.bottom - iRowDiff;
			}
		}
		else if (m_bCellMode ? (hWndCol != m_hWndColFocus || lRow != lRowFocusOld) : (lRow != lRowFocusOld))
		{
			if (m_bCellMode)
				GetFocusCellRect(&rCur);
			else
				GetRowFocusRect(lRowFocusOld, &rCur);
		}

		// Hat sich der Fokus geändert?
		BOOL bChanged = lRow != lRowFocusOld || hWndCol != m_hWndColFocus;

		// Muss Nachricht gesendet werden?
		BOOL bSendMsg = FALSE;
		if (m_bCellMode && bChanged && lRow != TBL_Error && hWndCol)
		{
			HWND hWndColEffNew = hWndCol, hWndColEffOld = m_hWndColFocus;
			long lRowEffNew = lRow, lRowEffOld = lRowFocusOld;

			GetMasterCell(hWndColEffOld, lRowEffOld);
			GetMasterCell(hWndColEffNew, lRowEffNew);

			bSendMsg = hWndColEffOld != hWndColEffNew || lRowEffOld != lRowEffNew;
		}

		// Neue Daten Fokus setzen
		//m_pRowFocus = m_Rows->GetRowPtr( lRow );
		m_pRowFocus = lRow != TBL_Error ? m_Rows->EnsureValid(lRow) : NULL;
		m_hWndColFocus = hWndCol;

		// Nachricht senden
		if (bSendMsg)
		{
			HWND hWndColSend = NULL;
			long lRowSend = TBL_Error;
			GetEffFocusCell(lRowSend, hWndColSend);
			SendMessage(m_hWndTbl, MTM_FocusCellChanged, (WPARAM)hWndColSend, (LPARAM)lRowSend);
		}

		// Neues Rechteck ermitteln
		if (m_bCellMode)
			GetFocusCellRect(&m_rFocus);
		else
			GetRowFocusRect(lRow, &m_rFocus);

		// Ermitteln, ob neu gezeichnet werden soll
		BOOL bForceInvalidate = FALSE;
		BOOL bNoInvalidate = FALSE;
		switch (iRedrawMode)
		{
		case UF_REDRAW_AUTO:
			bForceInvalidate = bChanged || m_bNoFocusPaint != s_bNoFocusPaint;
			break;
		case UF_REDRAW_INVALIDATE:
			bForceInvalidate = TRUE;
			break;
		case UF_REDRAW_NO_INVALIDATE:
			bNoInvalidate = TRUE;
			break;
		}

		// Selektion anpassen
		BOOL bSelChangeCells = m_bCellMode && !(dwFlags & UF_FLAG_SELCHANGE_ROWS);

		BOOL bKeyRange = FALSE;
		BOOL bSelSet = FALSE;
		BOOL bSelSetRange = FALSE;
		BOOL bSelSwitch = FALSE;
		BOOL bSelClearRowsAndCols = FALSE;
		BOOL bSelClearCells = FALSE;

		switch (iSelChangeMode)
		{
		case UF_SELCHANGE_FOCUS_ONLY:
			bSelClearRowsAndCols = TRUE;
			bSelClearCells = TRUE;
			bSelSet = TRUE;
			break;
		case UF_SELCHANGE_FOCUS_CELL_ONLY:
			bSelClearCells = TRUE;
			bSelSet = TRUE;
			break;
		case UF_SELCHANGE_FOCUS_ADD:
			bSelSet = TRUE;
			break;
		case UF_SELCHANGE_FOCUS_ADD_KEY_RANGE:
			bSelSetRange = TRUE;
			bKeyRange = TRUE;
			break;
		case UF_SELCHANGE_FOCUS_ADD_RANGE:
			bSelSetRange = TRUE;
			break;
		case UF_SELCHANGE_FOCUS_SWITCH:
			bSelSet = TRUE;
			bSelSwitch = TRUE;
			break;
		case UF_SELCHANGE_CLEAR_ALL:
			bSelClearRowsAndCols = TRUE;
			bSelClearCells = TRUE;
			break;
		case UF_SELCHANGE_CLEAR_ROWS_AND_COLS:
			bSelClearRowsAndCols = TRUE;
			break;
		case UF_SELCHANGE_CLEAR_CELLS:
			bSelClearCells = TRUE;
			break;
		}

		if (bSelClearRowsAndCols)
		{
			Ctd.TblClearSelection(m_hWndTbl);
			ClearColumnSelection();
		}
		if (bSelClearCells)
			ClearCellSelection();
		if (bSelSet)
		{
			BOOL bSet = TRUE;
			if (bSelChangeCells)
			{
				HWND hWndColSet = hWndCol;
				long lRowSet = lRow;

				LPMTBLMERGECELL pmc = FindMergeCell(hWndCol, lRow, FMC_SLAVE);
				if (pmc)
				{
					hWndColSet = pmc->hWndColFrom;
					lRowSet = pmc->lRowFrom;
				}

				if (bSelSwitch)
					bSet = !QueryCellFlags(lRowSet, hWndColSet, MTBL_CELL_FLAG_SELECTED);

				SetCellFlags(lRowSet, hWndColSet, NULL, MTBL_CELL_FLAG_SELECTED, bSet, TRUE);
			}
			else
			{
				if (bSelSwitch)
					bSet = !Ctd.TblQueryRowFlags(m_hWndTbl, lRow, ROW_Selected);

				Ctd.TblSetRowFlags(m_hWndTbl, lRow, ROW_Selected, bSet);
			}
		}
		if (bSelSetRange)
		{
			if (bSelChangeCells)
			{
				MTBLCELLRANGE cr, crA, crAOld;

				crAOld.SetEmpty();

				if (bKeyRange && m_CellSelRangeKey.IsValid())
				{
					cr.lRowFrom = m_CellSelRangeKey.lRowFrom;
					cr.hWndColFrom = m_CellSelRangeKey.hWndColFrom;
				}
				else
				{
					cr.lRowFrom = lRowFocusOld;
					cr.hWndColFrom = hWndColFocusOld;
				}

				cr.lRowTo = lRow;
				cr.hWndColTo = hWndCol;

				if (cr.IsValid())
				{
					if (bKeyRange)
					{
						if (m_CellSelRangeKey.IsValid())
						{
							m_CellSelRangeKey.CopyTo(crAOld);
							crAOld.Normalize();
							AdaptCellRange(crAOld);
						}

						cr.CopyTo(m_CellSelRangeKey);
					}

					cr.CopyTo(crA);
					crA.Normalize();
					AdaptCellRange(crA);

					if (bKeyRange && crAOld.IsValid())
						CellSelRangeChanged(crAOld, crA, TRUE);
					else
						SetCellRangeFlags(&crA, MTBL_CELL_FLAG_SELECTED, TRUE, TRUE);
				}
			}
			else
			{
				Ctd.TblSetFlagsRowRange(m_hWndTbl, lRowFocusOld, lRow, 0, ROW_Hidden | ROW_Selected, ROW_Selected, TRUE);
			}
		}

		// Focus zeichnen
		if (!bNoInvalidate)
		{
			HRGN rgn;
			HRGN rgnInv = CreateRectRgn(0, 0, 0, 0);
			if (rgnInv)
			{
				int iRgnType = NULLREGION;

				if (!IsRectEmpty(&rLast) && (bForceInvalidate || !EqualRect(&rLast, &rCur) || !EqualRect(&rLast, &m_rFocus)))
				{
					rgn = GetFocusRgn(&rLast, GFR_FRAME, s_bCellMode);
					if (rgn)
					{
						iRgnType = CombineRgn(rgnInv, rgnInv, rgn, RGN_OR);
						DeleteObject(rgn);
					}
				}
				if (!IsRectEmpty(&rCur) && (bForceInvalidate || !EqualRect(&rCur, &rLast) || !EqualRect(&rCur, &m_rFocus)))
				{
					rgn = GetFocusRgn(&rCur, GFR_FRAME, m_bCellMode);
					if (rgn)
					{
						iRgnType = CombineRgn(rgnInv, rgnInv, rgn, RGN_OR);
						DeleteObject(rgn);
					}
				}
				if (!IsRectEmpty(&m_rFocus) && (bForceInvalidate || !EqualRect(&m_rFocus, &rLast) || !EqualRect(&m_rFocus, &rCur)))
				{
					rgn = GetFocusRgn(&m_rFocus, GFR_FRAME, m_bCellMode);
					if (rgn)
					{
						iRgnType = CombineRgn(rgnInv, rgnInv, rgn, RGN_OR);
						DeleteObject(rgn);
					}
				}

				if (iRgnType != ERROR && iRgnType != NULLREGION)
				{
					InvalidateRgn(m_hWndTbl, rgnInv, FALSE);
				}

				DeleteObject(rgnInv);
			}
		}

		s_bNoFocusPaint = m_bNoFocusPaint;
		s_bCellMode = m_bCellMode;

		m_pCounter->DecNoFocusUpdate();
	}

	return TRUE;
}

BOOL MTBLSUBCLASS::Validate(HWND hWndColHasFocus, HWND hWndColGetFocus)
{
	if (!hWndColHasFocus)
	{
		long lRow = TBL_Error;
		Ctd.TblQueryFocus(m_hWndTbl, &lRow, &hWndColHasFocus);
	}

	if (hWndColHasFocus && Ctd.QueryFieldEdit(hWndColHasFocus))
	{
		if (!hWndColGetFocus)
			hWndColGetFocus = hWndColHasFocus;

		if (!Ctd.ValidateSet(hWndColGetFocus, FALSE, 0))
			return FALSE;
	}

	return TRUE;
}

//************************************************************
//Arbeitsbereich Alexander Hakert
//************************************************************

long mtblPrint_lCurrentPage;	//Aktuelle Seite
long mtblPrint_PageCount;		//Seitenanzahl
std::vector<MTBLPRINT_PRINTERDEVMODE> haMTblPrinterDevModes;	//Device Modes für Drucker

//******************** Print Part ********************

/*
Thread, der ein Statusfenster erstellt und Nachrichten des Fensters abfängt
Parameter:
lpParam:	Pointer auf ein PRINTSTATUSFORM Objekt, das als Statusform verwendet werden soll
Return:
-1, wenn ein Fehler aufgetreten ist
0, wenn das Statusfenster geschlossen wird
*/
DWORD WINAPI StatusFormThread(LPVOID lpParam)
{
	LPPRINTSTATUSFORM lpStatusForm = (LPPRINTSTATUSFORM)lpParam;
	//Erstelle Statusform
	lpStatusForm->CreateStatusForm();

	BOOL ret;
	MSG msg;
	while ((ret = GetMessage(&msg, 0, 0, 0)) != 0)
	{
		if (ret == -1)
			return -1;

		if (msg.message == WM_KEYDOWN)
			TranslateMessage(&msg);

		if (msg.message == WM_CHAR && msg.wParam == VK_ESCAPE)
			lpStatusForm->OnCancelClick();

		if (!IsDialogMessage(lpStatusForm->hWndStatusForm, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return 0;
}

/*
Callback-Funktion für die Aufzählung der auf dem Drucker installierten Fonts.
Speichert den Type Face Namen des LogFonts ab.
Parameter:
lpLogFont:		Pointer auf den Logfont mit den Attributen des Fonts
lpTextMetric:	Physikalische Attribute des Fonts
dwType:			Typ des Fonts (DEVICE_FONTTYPE, RASTER_FONTTYPE, TRUETYPE_FONTTYPE)
lpData:			tstring Vektor, in den der Font gespeichert werden soll
Return 1.
*/
int CALLBACK MTblPrintEnumFontsProc(const LOGFONT *lpLogFont, const TEXTMETRIC *lpTextMetric, DWORD dwType, MTBLSUBCLASS *lpSubclass)
{
	tstring tsFaceName = lpLogFont->lfFaceName;
	tstring tsEnclosed = _T("'");
	tstring tsPrintFont = tsEnclosed + tsFaceName + tsEnclosed;
	lpSubclass->m_tsPrintFonts += tsPrintFont;
	return 1;
}

/*
Interne Druckfunktion
Sammelt alle nötigen Daten aus der übergebenen Tabelle und druckt diese mit den übergebenen Druckoptionen
Parameter:
hWndTbl:		Handle für die Tabelle, die gedruckt werden soll
printOptions:	Druckoptionen (Druckername, Anzahl Kopien, ...)
Return:
error, wenn ein Fehler aufgetreten ist.
1, wenn erfolreich.
*/
int MTBLSUBCLASS::Print(HWND hWndTbl, PRINTOPTIONS printOptions)
{
	VecColInfo vecColsInfo;
	TABLEINFO tableInfo;
	PRINTMETRICS printMetrics;
	PRINTPAGEPARAMS printPageParams;
	int iPage;
	BOOL bAbort = FALSE, bCancel, bOk;
	long lColNext, lColFrom, lColTo, lColLow, lColMax, lStartRow, lNextStartRow, lLastRow, lLastTotal;
	BOOL bFound;
	LPTSTR lpsPrinter;
	HSTRING hsLanguageFile;
	HANDLE hPrinter;
	HDC hDeviceContext = NULL;
	int iResult = 0;
	int iPrintJob;
	HSTRING defPrinterName = 0;
	tstring sPrinter = _T("");

	DOCINFO docInfo;
	BOOL bStatus = FALSE;

	//Prüfen, ob überhaupt Drucker installiert sind
	if (!PrinterExists(_T("")))
	{
		//Keine Drucker installiert
		return ERR_NO_PRINTER;
	}
	//Sprache prüfen
	hsLanguageFile = 0;
	if (!MTblGetLanguageFile(printOptions.iLanguage, &hsLanguageFile))
	{
		//Invalid Language
		return ERR_INVALID_LANG;
	}
	if (hsLanguageFile != 0)
	{
		Ctd.HStringUnRef(hsLanguageFile);
	}

	//Drucker setzen
	//Drucker in den Druckeroptionen angegeben?
	if (!printOptions.tsPrinterName.empty())
	{
		//Drucker der Druckoptionen installiert?
		if (!PrinterExists(printOptions.tsPrinterName))
		{
			return ERR_INVALID_PRINTER;
		}
		sPrinter = printOptions.tsPrinterName;
	}
	else
	{
		//Kein Drucker angegeben, nehme Standarddrucker
		defPrinterName = MTblPrintGetDefPrinterName();
		Ctd.HStrToTStr(defPrinterName, sPrinter, TRUE);
	}

	//Kein Drucker gewählt und kein Default Drucker?
	if (sPrinter.empty())
		return ERR_NO_PRINTER;

	LPDEVMODE pDevMode = NULL;

	//Handle für Drucker ermitteln
	bStatus = OpenPrinter((LPTSTR)sPrinter.c_str(), &hPrinter, NULL);
	if (bStatus)
	{
		DWORD dwRet;
		pDevMode = GetPrinterDevModeCopy(sPrinter);
		//Kein Device Mode vorhanden?
		if (pDevMode == NULL)
		{
			//Kein Device Mode vorhanden
			DWORD dwNeeded = DocumentProperties(NULL, hPrinter, (LPTSTR)sPrinter.c_str(), NULL, NULL, 0);
			pDevMode = (LPDEVMODE)malloc(dwNeeded);

			//Device Mode vom Drucker abfragen
			DWORD dwRet = DocumentProperties(NULL, hPrinter, (LPTSTR)sPrinter.c_str(), pDevMode, NULL, DM_OUT_BUFFER);
			if (dwRet != IDOK)
			{
				free(pDevMode);
				ClosePrinter(hPrinter);
				return 0;
			}
		}

		if (pDevMode->dmFields & DM_ORIENTATION)
			pDevMode->dmOrientation = printOptions.orientation;

		if (pDevMode->dmFields & DM_COPIES)
			pDevMode->dmCopies = printOptions.iCopies;

		//Device Mode aktualisieren
		dwRet = DocumentProperties(NULL, hPrinter, (LPTSTR)sPrinter.c_str(), pDevMode, pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);

		ClosePrinter(hPrinter);
	}

	//Device Context erstellen
	hDeviceContext = CreateDC(_T(""), (LPTSTR)sPrinter.c_str(), NULL, pDevMode);

	//Fülle Vektor mit den vom Drucker unterstützten Fonts
	EnumFonts(hDeviceContext, NULL, (FONTENUMPROC)MTblPrintEnumFontsProc, (LPARAM)(this));

	//Tabelleninformationen ermitteln
	if (!GetTableInfo(hWndTbl, printOptions, tableInfo))
		return iResult;

	//Zu druckende Spalten ermitteln
	if (GetColsToPrint(tableInfo, vecColsInfo, printOptions) < 1)
		return ERR_NO_COLS_TO_PRINT;

	//Druckermetriken ermitteln
	if (!GetPrintMetrics(hDeviceContext, printOptions, tableInfo, vecColsInfo, &printMetrics))
		return iResult;

	//Ränder zu groß?
	if (printMetrics.clipRect.right <= printMetrics.clipRect.left || printMetrics.clipRect.bottom <= printMetrics.clipRect.top)
		return ERR_MARGINS_TOO_BIG;

	//Anzahl Seiten ermitteln
	if (IsPageCountNeeded(printOptions))
		mtblPrint_PageCount = GetPageCount(printOptions, printMetrics, tableInfo, vecColsInfo, hDeviceContext);
	else
		mtblPrint_PageCount = -1;

	//Eigentlicher Druck
	if (!(printOptions.bPreview))
	{
		//Fill in the structure with info about this document
		memset(&docInfo, 0, sizeof(DOCINFO));

		//Dokumentname setzen
		if (printOptions.tsDocName.empty())
			docInfo.lpszDocName = (LPTSTR)_T("MTBLPRINT");
		else
			docInfo.lpszDocName = printOptions.tsDocName.c_str();

		//Druck starten
		iPrintJob = StartDoc(hDeviceContext, &docInfo);
		m_iPrintJob = iPrintJob;
		if (iPrintJob > 0)
		{
			//Mapping-Modus setzen
			SetMapMode(hDeviceContext, MM_ANISOTROPIC);

			//1:1 Ausgabe
			SetWindowExtEx(hDeviceContext, 1, 1, NULL);
			SetViewportExtEx(hDeviceContext, 1, 1, NULL);

			//Parameter zum Drucken der Seite setzen
			printPageParams.hdcCanvas = hDeviceContext;
			printPageParams.printOptions = printOptions;
			printPageParams.printMetrics = printMetrics;
			printPageParams.tableInfo = tableInfo;
			printPageParams.vecColsInfo = vecColsInfo;

			//Statusfenster erzeugen
			LPPRINTSTATUSFORM lpStatusForm = new PRINTSTATUSFORM(printOptions, this, &bCancel);
			DWORD threadID;
			lpStatusForm->hStatusThread = CreateThread(NULL, 0, StatusFormThread, lpStatusForm, 0, &threadID);
			tstring tsStatusFormText;
			bCancel = FALSE;

			//Startwerte definieren
			lStartRow = -1;
			lNextStartRow = 0;
			lLastTotal = -1;
			iPage = 1;
			lColLow = 0;
			lColMax = vecColsInfo.size() - 1;
			lColFrom = lColLow;
			lColTo = lColLow;

			//Seite drucken
			while (TRUE)
			{
				//Neue Seite drucken
				bOk = StartPage(hDeviceContext);

				if (bOk)
				{
					//Format string mit Seitenzahl und setze in Statusform
					this->GetUIStr(printOptions, _T("StatDlg.Text"), tsStatusFormText);
					//Position des zu ersetzenden Teilstrings ermitteln
					tstring::size_type pos = tsStatusFormText.find(_T("%d"));

					//Teilstring gefunden?
					if (pos != tstring::npos)
					{
						//Teilstring ersetzen
						tsStatusFormText.replace(pos, 2, to_tstring(iPage));
					}
					SetDlgItemText(lpStatusForm->hWndStatusForm, IDC_STATUSFORM_TEXT, tsStatusFormText.c_str());

					//Bis-Spalte ermitteln
					lColNext = lColFrom + 1;
					bFound = FALSE;
					while (lColNext <= lColMax)
					{
						if (vecColsInfo.at(lColNext).bNewPage == TRUE)
						{
							lColTo = lColNext - 1;
							bFound = TRUE;
							break;
						}
						lColNext++;
					}

					if (!bFound)
						lColTo = lColMax;

					//Paramter zum Seitendruck setzen
					printPageParams.iPageNr = iPage;
					printPageParams.lStartRow = lStartRow;
					printPageParams.iFromCol = lColFrom;
					printPageParams.iToCol = lColTo;
					printPageParams.bColHeaders = (printOptions.bColHeaders && (lNextStartRow != -3));
					printPageParams.bCells = (lNextStartRow != -3);

					//Seite drucken
					lNextStartRow = PrintPage(printPageParams, lLastRow, lLastTotal);
					bOk = EndPage(hDeviceContext);

					//Fehler oder Abbruch?
					if (lNextStartRow == -2 || bCancel)
					{
						bAbort = TRUE;
						break;
					}

					//Druck beenden?
					if (lNextStartRow == -1 && lColTo == lColMax)
						break;

					//Nächste Seite
					iPage++;

					//Spalte von/bis und Startzeile setzen
					if (lColTo == lColMax)
					{
						lColFrom = lColLow;
						lColTo = lColLow;
						lStartRow = lNextStartRow;
					}
					else
						lColFrom = lColTo + 1;
				}
			}

			delete lpStatusForm;

			//Druck beenden
			if (bAbort)
				bOk = AbortDoc(hDeviceContext);
			else
				bOk = EndDoc(hDeviceContext);

			iResult = iPage;
		}
	}
	else
	{
		//Preview erstellen und öffnen
		LPPRINTPREVIEW lpPrintPreview;
		lpPrintPreview = new PRINTPREVIEW(this, printOptions, printMetrics, tableInfo, vecColsInfo, hDeviceContext, pDevMode);
		delete lpPrintPreview;
		iResult = 1;
	}

	// Aufräumen...
	if (pDevMode)
		free(pDevMode);

	if (hDeviceContext)
		DeleteDC(hDeviceContext);

	FreeDevModes();

	return iResult;
}

/*
Druckt eine Seite mit den übergebenen Parametern.
Schreibt den Index der letzten gedruckten Zeile in lLastPrintedRow und die letzte gedruckte Summe in lLastPrintedTotal
Parameter:
printPageParams:	Parameter der Seite, die gedruckt werden soll
lLastPrintedRow:	Letzte gedruckte Zeile, wird neu gesetzt
lLastPrintedTotal:	Letzte gedruckte Summenzeile, wird neu gesetzt
Return:
Letzte gedruckte Zeile, wenn der Druck noch nicht fertig ist.
-1, wenn keine Zeile mehr gedruckt werden muss, -3 wenn Summen nicht zu ende gedruckt wurden.
-2, wenn ein Fehler aufgetreten ist.
*/
int MTBLSUBCLASS::PrintPage(PRINTPAGEPARAMS printPageParams, long &lLastPrintedRow, long &lLastPrintedTotal)
{
	int i, iResult;
	PRINTCOLHEADERSPARAMS colHeaderParams;
	PRINTCELLSPARAMS cellsParams;
	PRINTLINEPARAMS lineParams;
	int iHeight, iY, iYmax;
	RECT rect;

	//Aktuelle Seite setzen
	mtblPrint_lCurrentPage = printPageParams.iPageNr;

	HFONT hFont = PrintCreateFont(printPageParams.hdcCanvas, printPageParams.printMetrics.tsFontName, printPageParams.printMetrics.dFontHeight, printPageParams.printMetrics.lFontStyle);
	HFONT hOldFont = (HFONT)SelectObject(printPageParams.hdcCanvas, hFont);

	rect.left = 0;
	rect.top = 0;
	DrawText(printPageParams.hdcCanvas, _T("A"), -1, &rect, DT_CALCRECT | DT_NOPREFIX);

	SelectObject(printPageParams.hdcCanvas, hOldFont);
	DeleteObject(hFont);

	//Parameter zum Drucken der Spaltenköpfe setzen
	colHeaderParams.hdcCanvas = printPageParams.hdcCanvas;
	colHeaderParams.printOptions = printPageParams.printOptions;
	colHeaderParams.printMetrics = printPageParams.printMetrics;
	colHeaderParams.tableInfo = printPageParams.tableInfo;
	colHeaderParams.vecColsInfo = printPageParams.vecColsInfo;
	colHeaderParams.iFromCol = printPageParams.iFromCol;
	colHeaderParams.iToCol = printPageParams.iToCol;

	//Parameter zum Drucken der Zellen setzen
	cellsParams.hdcCanvas = printPageParams.hdcCanvas;
	cellsParams.lStartRow = printPageParams.lStartRow;
	cellsParams.printOptions = printPageParams.printOptions;
	cellsParams.printMetrics = printPageParams.printMetrics;
	cellsParams.tableInfo = printPageParams.tableInfo;
	cellsParams.vecColsInfo = printPageParams.vecColsInfo;
	cellsParams.iFromCol = printPageParams.iFromCol;
	cellsParams.iToCol = printPageParams.iToCol;

	//Parameter zum Drucken der Zeilen setzen
	lineParams.hdcCanvas = printPageParams.hdcCanvas;
	lineParams.clipRect = printPageParams.printMetrics.clipRect;
	lineParams.tsFontName = printPageParams.printMetrics.tsFontName;
	lineParams.lpPrintMetrics = &(printPageParams.printMetrics);
	lineParams.lpTableInfo = &(printPageParams.tableInfo);
	lineParams.bSimulation = FALSE;

	//Drucken starten
	iY = printPageParams.printMetrics.clipRect.top;
	iYmax = printPageParams.printMetrics.clipRect.bottom;

	//Titel drucken, wenn erste Seite
	if (printPageParams.lStartRow == -1)
	{
		lineParams.hdcCanvas = printPageParams.hdcCanvas;
		lineParams.clipRect = printPageParams.printMetrics.clipRect;
		lineParams.tsFontName = printPageParams.printMetrics.tsFontName;
		lineParams.lpPrintMetrics = &(printPageParams.printMetrics);
		lineParams.lpTableInfo = &(printPageParams.tableInfo);
		lineParams.bYReversed = FALSE;

		if (printPageParams.iPageNr > 1)
			lineParams.bSimulation = TRUE;

		for (i = 0; i < printPageParams.printOptions.iTitleCount; i++)
		{
			lineParams.line = printPageParams.printOptions.vecTitles.at(i);
			lineParams.iYStart = iY;
			lineParams.clipRect.top = iY;
			if ((lineParams.line.dwFlags & MTPL_NO_TEXT_SCALE) == 0)
				lineParams.dFontHeight = printPageParams.printMetrics.dFontHeight * lineParams.line.dFontSizeMulti;
			else
				lineParams.dFontHeight = printPageParams.printMetrics.dFontHeightNS * lineParams.line.dFontSizeMulti;

			lineParams.lFontStyle = printPageParams.printOptions.vecTitles.at(i).lFontEnh;
			lineParams.crFontColor = printPageParams.printOptions.vecTitles.at(i).crFontColor;
			rect = PrintLine(lineParams);

			if (rect.bottom > rect.top)
				iY = rect.bottom;
		}
	}

	//Seitenkopfzeilen drucken
	lineParams.hdcCanvas = printPageParams.hdcCanvas;
	lineParams.clipRect = printPageParams.printMetrics.clipRect;
	lineParams.tsFontName = printPageParams.printMetrics.tsFontName;
	lineParams.lpPrintMetrics = &(printPageParams.printMetrics);
	lineParams.lpTableInfo = &(printPageParams.tableInfo);
	lineParams.bYReversed = FALSE;
	lineParams.bSimulation = FALSE;

	for (i = 0; i < printPageParams.printOptions.iPageHeaderCount; i++)
	{
		lineParams.line = printPageParams.printOptions.vecPageHeaders.at(i);
		lineParams.iYStart = iY;
		lineParams.clipRect.top = iY;

		if ((lineParams.line.dwFlags & MTPL_NO_TEXT_SCALE) == 0)
			lineParams.dFontHeight = printPageParams.printMetrics.dFontHeight * lineParams.line.dFontSizeMulti;
		else
			lineParams.dFontHeight = printPageParams.printMetrics.dFontHeightNS * lineParams.line.dFontSizeMulti;

		lineParams.lFontStyle = printPageParams.printOptions.vecPageHeaders.at(i).lFontEnh;
		lineParams.crFontColor = printPageParams.printOptions.vecPageHeaders.at(i).crFontColor;
		rect = PrintLine(lineParams);

		if (rect.bottom > rect.top)
			iY = rect.bottom;
	}

	//Seitenzahlen
	if (printPageParams.printOptions.bPageNrs)
	{
		lineParams.hdcCanvas = printPageParams.hdcCanvas;
		GetUIStr(printPageParams.printOptions, _T("DefPageNr"), lineParams.line.tsCenterText);
		lineParams.line.tsLeftText = tstring(_T(""));
		lineParams.line.tsRightText = tstring(_T(""));
		lineParams.line.iPosTextCount = 0;
		lineParams.line.hLeftImage = 0;
		lineParams.line.hCenterImage = 0;
		lineParams.line.hRightImage = 0;
		lineParams.iYStart = iYmax;
		lineParams.clipRect = printPageParams.printMetrics.clipRect;
		lineParams.clipRect.top = iY;
		lineParams.tsFontName = printPageParams.printMetrics.tsFontName;

		if (!(printPageParams.printOptions.bNoPageNrsScale))
			lineParams.dFontHeight = printPageParams.printMetrics.dFontHeight;
		else
			lineParams.dFontHeight = printPageParams.printMetrics.dFontHeightNS;

		lineParams.lFontStyle = 0;
		lineParams.crFontColor = RGB(0, 0, 0);
		lineParams.lpPrintMetrics = &(printPageParams.printMetrics);
		lineParams.lpTableInfo = &(printPageParams.tableInfo);
		lineParams.bYReversed = TRUE;
		lineParams.bSimulation = FALSE;
		rect = PrintLine(lineParams);

		if (rect.bottom > rect.top)
			iYmax = rect.top;
	}

	//Seitenfußzeilen
	lineParams.hdcCanvas = printPageParams.hdcCanvas;
	lineParams.clipRect = printPageParams.printMetrics.clipRect;
	lineParams.tsFontName = printPageParams.printMetrics.tsFontName;
	lineParams.lpPrintMetrics = &(printPageParams.printMetrics);
	lineParams.lpTableInfo = &(printPageParams.tableInfo);
	lineParams.bYReversed = TRUE;
	lineParams.bSimulation = FALSE;

	for (i = 0; i < printPageParams.printOptions.iPageFooterCount; i++)
	{
		lineParams.line = printPageParams.printOptions.vecPageFooters.at(i);
		lineParams.iYStart = iYmax;
		lineParams.clipRect.top = iY;

		if ((lineParams.line.dwFlags & MTPL_NO_TEXT_SCALE) == 0)
			lineParams.dFontHeight = printPageParams.printMetrics.dFontHeight * lineParams.line.dFontSizeMulti;
		else
			lineParams.dFontHeight = printPageParams.printMetrics.dFontHeightNS * lineParams.line.dFontSizeMulti;

		lineParams.lFontStyle = printPageParams.printOptions.vecPageFooters.at(i).lFontEnh;
		lineParams.crFontColor = printPageParams.printOptions.vecPageFooters.at(i).crFontColor;
		rect = PrintLine(lineParams);

		if (rect.bottom > rect.top)
			iYmax = rect.top;
	}

	//Spaltenköpfe drucken
	if (printPageParams.bColHeaders)
	{
		colHeaderParams.iYStart = iY;
		colHeaderParams.iYMax = iYmax;
		iHeight = PrintColHeaders(printPageParams.hdcCanvas, colHeaderParams);

		//Fehler?
		if (iHeight < 0)
		{
			return -2;
		}

		iY += iHeight;
	}

	//Zellen drucken
	if (printPageParams.bCells)
	{
		cellsParams.iYStart = iY;
		cellsParams.iYMax = iYmax;
		iResult = PrintCells(printPageParams.hdcCanvas, cellsParams, lLastPrintedRow, iY);
	}
	else
		iResult = -1;

	//Totals drucken
	if (iResult == -1 && cellsParams.iToCol == cellsParams.vecColsInfo.size() - 1)
	{
		lineParams.hdcCanvas = printPageParams.hdcCanvas;
		lineParams.clipRect = printPageParams.printMetrics.clipRect;
		lineParams.clipRect.bottom = iYmax;
		lineParams.tsFontName = printPageParams.printMetrics.tsFontName;
		lineParams.lpPrintMetrics = &(printPageParams.printMetrics);
		lineParams.bYReversed = FALSE;
		lineParams.bSimulation = FALSE;

		for (i = lLastPrintedTotal + 1; i < printPageParams.printOptions.iTotalCount; i++)
		{
			lineParams.line = printPageParams.printOptions.vecTotals.at(i);
			lineParams.iYStart = iY;
			lineParams.clipRect.top = iY;
			if ((lineParams.line.dwFlags & MTPL_NO_TEXT_SCALE) == 0)
				lineParams.dFontHeight = printPageParams.printMetrics.dFontHeight * lineParams.line.dFontSizeMulti;
			else
				lineParams.dFontHeight = printPageParams.printMetrics.dFontHeightNS * lineParams.line.dFontSizeMulti;

			lineParams.lFontStyle = printPageParams.printOptions.vecTotals.at(i).lFontEnh;
			lineParams.crFontColor = printPageParams.printOptions.vecTotals.at(i).crFontColor;
			rect = PrintLine(lineParams);

			if (rect.bottom > rect.top)
			{
				iY = rect.bottom;
				lLastPrintedTotal = i;
			}
			else
			{
				iResult = -3;
				lLastPrintedTotal = i - 1;
				break;
			}
		}
	}

	return iResult;
}

/*
Druckt eine Zeile mit den übergebenen Parametern.
Parameter:
printLineParams:	Parameter der Zeile, die gedruckt werden soll
Return:
RECT Struktur, die den Platz, den die Zeile in Anspruch genommen hat, beinhaltet.
*/
RECT MTBLSUBCLASS::PrintLine(PRINTLINEPARAMS printLineParams)
{
	BOOL bAnyPosTextToPrint;
	DWORD dwDrawFlags;
	int i, iHeight;
	tstring tsLeftText, tsCenterText, tsRightText;
	std::vector<tstring> vecPosTexts;
	RECT rect, rectPosText;
	long lJfyFlags;
	double dWidth;
	RECT rectResult;

	rectResult.left = printLineParams.clipRect.left;
	rectResult.right = printLineParams.clipRect.left;
	rectResult.top = printLineParams.iYStart;
	rectResult.bottom = printLineParams.iYStart;

	//Positionstexte kopieren
	bAnyPosTextToPrint = FALSE;
	if (printLineParams.line.iPosTextCount > 0)
	{
		for (i = 0; i < printLineParams.line.iPosTextCount; i++)
		{
			vecPosTexts.push_back(printLineParams.line.vecplPosText.at(i).tsText);
			if (!vecPosTexts.at(i).empty())
				bAnyPosTextToPrint = TRUE;
		}
	}

	//Überhaupt was zu drucken?
	if (printLineParams.line.tsLeftText.empty() &&
		printLineParams.line.tsCenterText.empty() &&
		printLineParams.line.tsRightText.empty() &&
		!bAnyPosTextToPrint &&
		printLineParams.line.hLeftImage == 0 && printLineParams.line.hCenterImage == 0 && printLineParams.line.hRightImage == 0)
	{
		return rectResult;
	}

	tsLeftText = ExpandPlaceholders(printLineParams.line.tsLeftText);
	tsCenterText = ExpandPlaceholders(printLineParams.line.tsCenterText);
	tsRightText = ExpandPlaceholders(printLineParams.line.tsRightText);

	for (i = 0; i < printLineParams.line.iPosTextCount; i++)
	{
		vecPosTexts.at(i) = ExpandPlaceholders(vecPosTexts.at(i));
	}

	//Font
	HFONT hFont = PrintCreateFont(printLineParams.hdcCanvas, printLineParams.tsFontName, printLineParams.dFontHeight, printLineParams.lFontStyle);
	HFONT hOldFont = (HFONT)SelectObject(printLineParams.hdcCanvas, hFont);

	//Ausgabeflags setzen
	dwDrawFlags = DT_NOPREFIX;

	//Benötigte Zeilenhöhe ermitteln
	iHeight = 0;
	//Texte
	iHeight = max(iHeight, PrintLine_GetTextHeight(tsLeftText, printLineParams, dwDrawFlags));
	iHeight = max(iHeight, PrintLine_GetTextHeight(tsCenterText, printLineParams, dwDrawFlags));
	iHeight = max(iHeight, PrintLine_GetTextHeight(tsRightText, printLineParams, dwDrawFlags));

	for (i = 0; i < printLineParams.line.iPosTextCount; i++)
	{
		iHeight = max(iHeight, PrintLine_GetTextHeight(vecPosTexts.at(i), printLineParams, dwDrawFlags));
	}

	//Bilderhöhe
	iHeight = max(iHeight, PrintLine_GetImageHeight(printLineParams.line.hLeftImage, printLineParams));
	iHeight = max(iHeight, PrintLine_GetImageHeight(printLineParams.line.hCenterImage, printLineParams));
	iHeight = max(iHeight, PrintLine_GetImageHeight(printLineParams.line.hRightImage, printLineParams));

	rectResult.bottom = rectResult.top + iHeight;

	if (printLineParams.bYReversed)
	{
		iHeight = rectResult.bottom - rectResult.top;
		rectResult.bottom = rectResult.top;
		rectResult.top = rectResult.bottom - iHeight;
	}

	rectResult.left = max(rectResult.left, printLineParams.clipRect.left);
	rectResult.top = max(rectResult.top, printLineParams.clipRect.top);
	rectResult.right = printLineParams.clipRect.right;
	rectResult.bottom = min(rectResult.bottom, printLineParams.clipRect.bottom);

	if (printLineParams.bSimulation || rectResult.right <= rectResult.left || rectResult.bottom <= rectResult.top || rectResult.bottom > printLineParams.clipRect.bottom)
		return rectResult;

	//Zeile ausgeben
	//Bilder
	PrintLine_DrawImage(printLineParams.line.hLeftImage, DT_LEFT, rectResult, printLineParams);
	PrintLine_DrawImage(printLineParams.line.hCenterImage, DT_CENTER, rectResult, printLineParams);
	PrintLine_DrawImage(printLineParams.line.hRightImage, DT_RIGHT, rectResult, printLineParams);

	//Texte
	SetBkMode(printLineParams.hdcCanvas, TRANSPARENT);

	int h;

	if (!tsLeftText.empty())
		h = DrawText(printLineParams.hdcCanvas, tsLeftText.c_str(), -1, &rectResult, dwDrawFlags | DT_LEFT);
	if (!tsCenterText.empty())
		h = DrawText(printLineParams.hdcCanvas, tsCenterText.c_str(), -1, &rectResult, dwDrawFlags | DT_CENTER);
	if (!tsRightText.empty())
		h = DrawText(printLineParams.hdcCanvas, tsRightText.c_str(), -1, &rectResult, dwDrawFlags | DT_RIGHT);

	//Positionierte Texte
	if (bAnyPosTextToPrint)
	{
		//Fixe Rechteck-Werte
		rectPosText.top = rectResult.top;
		rectPosText.bottom = rectResult.bottom;

		for (i = 0; i < printLineParams.line.iPosTextCount; i++)
		{
			if (!vecPosTexts.at(i).empty())
			{
				//Dynamische Breite?
				if (printLineParams.line.vecplPosText.at(i).dWidth == 0.0)
				{
					SetRectEmpty(&rect);
					DrawText(printLineParams.hdcCanvas, vecPosTexts.at(0).c_str(), -1, &rect, DT_CALCRECT | DT_EXTERNALLEADING);
					dWidth = rect.right - rect.left;
				}
				else
					dWidth = printLineParams.line.vecplPosText.at(i).dWidth;

				//Rechteck setzen
				rectPosText.left = max(printLineParams.lpPrintMetrics->iLeftMargin + round(printLineParams.line.vecplPosText.at(i).dLeft * printLineParams.lpPrintMetrics->dPixelPerMMX), rectResult.left);
				rectPosText.right = min(rectPosText.left + round(dWidth * printLineParams.lpPrintMetrics->dPixelPerMMX), rectResult.right);

				//Ausrichtung ermitteln
				if (printLineParams.line.vecplPosText.at(i).justify == jfyCenter)
					lJfyFlags = DT_CENTER;
				else if (printLineParams.line.vecplPosText.at(i).justify == jfyRight)
					lJfyFlags = DT_RIGHT;
				else
					lJfyFlags = DT_LEFT;

				//Ausgeben
				DrawText(printLineParams.hdcCanvas, vecPosTexts.at(i).c_str(), -1, &rectPosText, dwDrawFlags | lJfyFlags);
			}
		}
	}

	SelectObject(printLineParams.hdcCanvas, hOldFont);
	DeleteObject(hFont);

	return rectResult;
}

/*
Druckt eine Tabellenzelle mit den übergebenen Parametern. Dabei wird der folgende Ablauf durchlaufen:
-Hintergrund zeichnen
-Umrandung zeichnen
-Baumlinien zeichnen
-Knoten zeichnen
-Bild zeichnen
-Text schreiben
Parameter:
printCellParams:	Parameter der Zelle, die gedruckt werden soll
Return:
1, wenn die Zelle erfolgreich gedruckt wurde.
0, wenn sich das Rechteck der Zelle außerhalb des Clippingbereichs befindet.
*/
int MTBLSUBCLASS::PrintCell(PRINTCELLPARAMS printCellParams)
{
	BOOL bLeftBorder, bTopBorder, bRightBorder, bBottomBorder, bHigher, bWider;
	RECT rect, rectClip, rectImage, rectText;
	int iNodeSize, iStartOffset, iEndOffset, iOldBackgroundMode, iCurrentRange;
	HRGN hRange;
	int iResult = 0;

	rect = printCellParams.cellData.rect;
	//Ausgaberechteck mit Clipping-Rechteck vergleichen
	//Ausßerhalb des Clippingbereichs?
	if (rect.left >= printCellParams.cellData.clipRect.right)
		return iResult;
	if (rect.top >= printCellParams.cellData.clipRect.bottom)
		return iResult;
	if (printCellParams.cellData.clipRect.left > printCellParams.cellData.clipRect.right)
		return iResult;

	//Ausgabe auf Clippingbereich beschränken
	rect.left = max(rect.left, printCellParams.cellData.clipRect.left);
	rect.top = max(rect.top, printCellParams.cellData.clipRect.top);
	rect.right = min(rect.right, printCellParams.cellData.clipRect.right);
	rect.bottom = min(rect.bottom, printCellParams.cellData.clipRect.bottom);

	//Hintergrund zeichnen
	HBRUSH hBrush = CreateSolidBrush(printCellParams.cellData.crBackgroundColor);
	FillRect(printCellParams.hdcCanvas, &rect, hBrush);
	DeleteObject(hBrush);

	//Umrandung ausgeben
	bLeftBorder = printCellParams.cellData.rBorders.left;
	bTopBorder = printCellParams.cellData.rBorders.top;
	bRightBorder = printCellParams.cellData.rBorders.right;
	bBottomBorder = printCellParams.cellData.rBorders.bottom;

	HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
	HPEN hOldPen = (HPEN)SelectObject(printCellParams.hdcCanvas, hPen);

	if (bLeftBorder)
	{
		MoveToEx(printCellParams.hdcCanvas, rect.left, rect.top, NULL);
		LineTo(printCellParams.hdcCanvas, rect.left, rect.bottom);
	}
	if (bTopBorder)
	{
		MoveToEx(printCellParams.hdcCanvas, rect.left, rect.top, NULL);
		LineTo(printCellParams.hdcCanvas, rect.right, rect.top);
	}
	if (bRightBorder)
	{
		MoveToEx(printCellParams.hdcCanvas, rect.right, rect.bottom, NULL);
		LineTo(printCellParams.hdcCanvas, rect.right, rect.top);
	}
	if (bBottomBorder)
	{
		MoveToEx(printCellParams.hdcCanvas, rect.left, rect.bottom, NULL);
		LineTo(printCellParams.hdcCanvas, rect.right, rect.bottom);
	}

	SelectObject(printCellParams.hdcCanvas, hOldPen);
	DeleteObject(hPen);

	//Clipping für Tree-Linien setzen
	rectClip = printCellParams.cellData.clipRect;

	POINT points[2];
	points[0].x = rectClip.left;
	points[0].y = rectClip.top;
	points[1].x = rectClip.right;
	points[1].y = rectClip.bottom;

	//Logische zu Device Koordinaten
	LPtoDP(printCellParams.hdcCanvas, points, 2);

	rectClip.left = points[0].x;
	rectClip.top = points[0].y;
	rectClip.right = points[1].x;
	rectClip.bottom = points[1].y;

	hRange = CreateRectRgn(rectClip.left, rectClip.top, rectClip.right, rectClip.bottom);

	if (hRange != 0)
	{
		SelectClipRgn(printCellParams.hdcCanvas, hRange);
		DeleteObject(hRange);
	}

	//Tree-Linien zeichnen
	if (printCellParams.cellData.bHorzTreeLine || printCellParams.cellData.bVertTreeLineToParent || printCellParams.cellData.bVertTreeLineToChild)
	{
		HPEN hPen = CreatePen(PS_SOLID, 1, printCellParams.cellData.crTreeLineColor);
		HPEN hOldPen = (HPEN)SelectObject(printCellParams.hdcCanvas, hPen);

		//Horizontale Linie
		if (printCellParams.cellData.bHorzTreeLine)
		{
			MoveToEx(printCellParams.hdcCanvas, printCellParams.cellData.rHTL_Pos.left, printCellParams.cellData.rHTL_Pos.top, NULL);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.rHTL_Pos.right, printCellParams.cellData.rHTL_Pos.bottom);
		}

		//Vertikale Linie Kind zu Elternzelle
		if (printCellParams.cellData.bVertTreeLineToParent)
		{
			for (int l = 0; l < printCellParams.cellData.vecVTLTP_Pos.size(); l++)
			{
				MoveToEx(printCellParams.hdcCanvas, printCellParams.cellData.vecVTLTP_Pos.at(l).left, printCellParams.cellData.vecVTLTP_Pos.at(l).top, NULL);
				LineTo(printCellParams.hdcCanvas, printCellParams.cellData.vecVTLTP_Pos.at(l).right, printCellParams.cellData.vecVTLTP_Pos.at(l).bottom);
			}
		}

		//Vertikale Linie Elternzelle zu Kind
		if (printCellParams.cellData.bVertTreeLineToChild)
		{
			MoveToEx(printCellParams.hdcCanvas, printCellParams.cellData.rVTLTC_Pos.left, printCellParams.cellData.rVTLTC_Pos.top, NULL);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.rVTLTC_Pos.right, printCellParams.cellData.rVTLTC_Pos.bottom);
		}

		SelectObject(printCellParams.hdcCanvas, hOldPen);
		DeleteObject(hPen);
	}

	//Clipping für Knoten, Bild und Text setzen
	rectClip = printCellParams.cellData.clipRect;
	rectClip.left++;
	rectClip.top++;
	rectClip.right--;
	rectClip.bottom--;

	points[0].x = rectClip.left;
	points[0].y = rectClip.top;
	points[1].x = rectClip.right;
	points[1].y = rectClip.bottom;

	LPtoDP(printCellParams.hdcCanvas, points, 2);

	rectClip.left = points[0].x;
	rectClip.top = points[0].y;
	rectClip.right = points[1].x;
	rectClip.bottom = points[1].y;

	hRange = CreateRectRgn(rectClip.left, rectClip.top, rectClip.right, rectClip.bottom);

	if (hRange != 0)
	{
		SelectClipRgn(printCellParams.hdcCanvas, hRange);
		DeleteObject(hRange);
	}

	//Knoten zeichnen
	if (printCellParams.cellData.bPaintNode)
	{
		iNodeSize = printCellParams.cellData.nodeRect.right - printCellParams.cellData.nodeRect.left;

		if (printCellParams.cellData.hNodeImageHandle == 0)
		{
			//Hintergrund
			HBRUSH hBrush = CreateSolidBrush(printCellParams.cellData.crNodeBackColor);
			FillRect(printCellParams.hdcCanvas, &printCellParams.cellData.nodeRect, hBrush);
			DeleteObject(hBrush);

			//Rechteck zeichnen
			HPEN hPen = CreatePen(PS_SOLID, 1, printCellParams.cellData.crNodeOuterColor);
			HPEN hOldPen = (HPEN)SelectObject(printCellParams.hdcCanvas, hPen);

			MoveToEx(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left, printCellParams.cellData.nodeRect.top, NULL);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.right, printCellParams.cellData.nodeRect.top);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.right, printCellParams.cellData.nodeRect.bottom);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left, printCellParams.cellData.nodeRect.bottom);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left, printCellParams.cellData.nodeRect.top);

			SelectObject(printCellParams.hdcCanvas, hOldPen);
			DeleteObject(hPen);

			hPen = CreatePen(PS_SOLID, 1, printCellParams.cellData.crNodeInnerColor);
			hOldPen = (HPEN)SelectObject(printCellParams.hdcCanvas, hPen);

			iStartOffset = trunc(2 * printCellParams.cellData.dTblPixSize);
			iEndOffset = trunc(2 * printCellParams.cellData.dTblPixSize);

			//Horizontaler Strich
			MoveToEx(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left + iStartOffset, printCellParams.cellData.nodeRect.top + iNodeSize / 2.0, NULL);
			LineTo(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.right - iEndOffset, printCellParams.cellData.nodeRect.top + iNodeSize / 2.0);

			//Vertikaler Strich
			if (printCellParams.cellData.iNodeType == MTBL_NODE_PLUS)
			{
				MoveToEx(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left + iNodeSize / 2.0, printCellParams.cellData.nodeRect.top + iStartOffset, NULL);
				LineTo(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left + iNodeSize / 2.0, printCellParams.cellData.nodeRect.bottom - iEndOffset);
			}

			SelectObject(printCellParams.hdcCanvas, hOldPen);
			DeleteObject(hPen);
		}
		else
		{
			bWider = (printCellParams.cellData.siNodeImageSize.cx > iNodeSize);
			bHigher = (printCellParams.cellData.siNodeImageSize.cy > iNodeSize);
			//Rechteck für Knotenimage anpassen
			if (!bWider)
				rectImage.left = printCellParams.cellData.nodeRect.left + ((iNodeSize - printCellParams.cellData.siNodeImageSize.cx) / 2.0);
			else
				rectImage.left = printCellParams.cellData.nodeRect.left;

			rectImage.right = rectImage.left + printCellParams.cellData.siNodeImageSize.cx;

			if (!bHigher)
				rectImage.top = printCellParams.cellData.nodeRect.top + ((iNodeSize - printCellParams.cellData.siNodeImageSize.cy) / 2.0);
			else
				rectImage.top = printCellParams.cellData.nodeRect.top;

			rectImage.bottom = rectImage.top + printCellParams.cellData.siNodeImageSize.cy;

			hRange = CreateRectRgn(0, 0, 0, 0);
			iCurrentRange = GetClipRgn(printCellParams.hdcCanvas, hRange);

			if (iCurrentRange == 1)
				IntersectClipRect(printCellParams.hdcCanvas, printCellParams.cellData.nodeRect.left, printCellParams.cellData.nodeRect.top, printCellParams.cellData.nodeRect.right, printCellParams.cellData.nodeRect.bottom);

			//Knoten zeichnen
			MImgPrint(printCellParams.cellData.hNodeImageHandle, printCellParams.hdcCanvas, &rectImage, printCellParams.cellData.crBackgroundColor, 0);

			if (iCurrentRange == 1)
				SelectClipRgn(printCellParams.hdcCanvas, hRange);
			if (hRange != 0)
				DeleteObject(hRange);
		}
	}

	//Bild zeichnen
	if (printCellParams.cellData.hImage != 0)
	{
		rectImage = printCellParams.cellData.rImageRect;
		MImgPrint(printCellParams.cellData.hImage, printCellParams.hdcCanvas, &rectImage, printCellParams.cellData.crBackgroundColor, 0);
	}

	//Font setzen
	HFONT hFont = PrintCreateFont(printCellParams.hdcCanvas, printCellParams.cellData.tsFontName, printCellParams.cellData.dFontHeight, printCellParams.cellData.lFontStyle);
	HFONT hOldFont = (HFONT)SelectObject(printCellParams.hdcCanvas, hFont);
	COLORREF crOldColor = SetTextColor(printCellParams.hdcCanvas, printCellParams.cellData.crFontColor);

	//Text ausgeben
	iOldBackgroundMode = SetBkMode(printCellParams.hdcCanvas, TRANSPARENT);
	rectText = printCellParams.cellData.rTextRect;
	DrawText(printCellParams.hdcCanvas, printCellParams.cellData.tsText.c_str(), -1, &rectText, printCellParams.cellData.dwTextDrawFlags);
	SetBkMode(printCellParams.hdcCanvas, iOldBackgroundMode);

	SetTextColor(printCellParams.hdcCanvas, crOldColor);
	SelectObject(printCellParams.hdcCanvas, hOldFont);
	DeleteObject(hFont);

	//Benutzerdefiniert gezeichnet?
	if (printCellParams.cellData.odi.lType > 0)
	{
		//Canvas zurücksetzen
		SelectObject(printCellParams.hdcCanvas, GetStockObject(BLACK_PEN));
		SelectObject(printCellParams.hdcCanvas, GetStockObject(HOLLOW_BRUSH));
		SelectObject(printCellParams.hdcCanvas, GetStockObject(SYSTEM_FONT));

		//Zeichnen lassen
		MTblODIPrint(printCellParams.hwTable, &printCellParams.cellData.odi);
	}

	//Clipping entfernen
	if (hRange != 0)
		SelectClipRgn(printCellParams.hdcCanvas, 0);

	iResult = 1;

	return iResult;
}

/*
Druckt alle Zellen der Tabelle. Dabei werden folgende Schritte durchlaufen:
Für jede Zeile und Zelle:
-Daten für die Zelle ermitteln (Mergezelle oder normale Zelle, Ausgaberechteck, Text, Hintegrund, Bild, Knoten, Baumlinien, Ränder, ...)
-Zelle drucken
Ist das Ende einer Seite erreicht, wird der Vorgang abgebrochen.
Parameter:
hDeviceContext:		Device Kontext, auf den gedruckt werden soll
printCellsParams:	Parameter der Zellen, die gedruckt werden sollen
lLastPrintedRow:	Letzte gedruckte Zeile, wird neu gesetzt
iYBottom:			Wird auf die Höhe des gedruckten Zellenbereichs gesetzt
Return:
Letzte gedruckte Zeile, wenn noch Zeilen gedruckt werden müssen.
-1, wenn keine Zeile mehr gedruckt werden muss.
-2, wenn ein Fehler aufgetreten ist.
*/
int MTBLSUBCLASS::PrintCells(HDC hDeviceContext, PRINTCELLSPARAMS printCellsParams, long &lLastPrintedRow, int &iYBottom)
{
	int i, iDrawLevel, iEff, iEffRowLevel, iFontHeight, iTreeIndentPerLevel, iNodeWid, iNodeHt, iNodeLevel, iX, iY, iColws, iLastRow, iPrinted, iTxtOffsX, iTxtOffsY;
	HWND hWndCol, hWndEffCol;
	DWORD dwColor;
	HSTRING hsText;
	long lLen, lParentRow, lRow, lRow2, lRowHeight, lRowHeight2, lFirstRow, lEffRow;
	PRINTCELLPARAMS printCellParams;
	BOOL bChildRowMerged, bMergeCell, bPageFull, bPageBreak, bInvLastLockedColLine, bVarRowHeight;
	std::vector<int> vecColWidth;
	RECT rCellBorders, rect;
	TEXTMETRIC textMetric;
	int j, k;
	std::vector<BOOL> vecBNoLeftBorder, vecBNoTopBorder, vecBNoRightBorder, vecBNoBottomBorder;
	LPMTBLMERGECELL lpMergeCell;
	LPCELLTOPRINT lpCellToPrint;
	HANDLE hImageNodeCol, hImageNodeExp;
	int iResult = -2;
	LPTSTR lps;

	lFirstRow = TBL_Error;
	iColws = 0;

	//Spaltenbreiten in Druckerpixeln ermitteln
	for (i = 0; i < printCellsParams.vecColsInfo.size(); i++)
	{
		vecColWidth.push_back(PixTblToPrint(printCellsParams.vecColsInfo.at(i).iWidthPix, FALSE, printCellsParams.tableInfo, printCellsParams.printMetrics));
		if (i >= printCellsParams.iFromCol && i <= printCellsParams.iToCol)
			iColws += vecColWidth.at(i);
	}

	//Texthöhe und Textoffsets in Druckerpixeln ermitteln
	iTxtOffsX = round(printCellsParams.tableInfo.metrics.ptCellLeading.x / (double)(printCellsParams.tableInfo.iPixelPerInchX) * printCellsParams.printMetrics.iPixelPerInchX * printCellsParams.printMetrics.dScale);
	iTxtOffsY = round(printCellsParams.tableInfo.metrics.ptCellLeading.y / (double)(printCellsParams.tableInfo.iPixelPerInchY) * printCellsParams.printMetrics.iPixelPerInchY * printCellsParams.printMetrics.dScale);

	//Fixe Parameter zum Druck der Zellen setzen
	//Device Context
	printCellParams.hdcCanvas = printCellsParams.hdcCanvas;
	//Tabellen Handle
	printCellParams.hwTable = printCellsParams.tableInfo.handle;
	//Ränder allgemein
	if (printCellsParams.printOptions.gridType == gtStandard3 || printCellsParams.printOptions.gridType == gtStandard4)
		BorderFlagsFromGridType(gtNo, rCellBorders);
	else
		BorderFlagsFromGridType(printCellsParams.printOptions.gridType, rCellBorders);

	//Text-Offsets
	printCellParams.cellData.iTextOffsetX = iTxtOffsX;
	printCellParams.cellData.iTextOffsetY = iTxtOffsY;
	//Größe Tabellenpixel
	printCellParams.cellData.dTblPixSize = 1.0 / (double)(printCellsParams.tableInfo.iPixelPerInchX) * printCellsParams.printMetrics.iPixelPerInchX * printCellsParams.printMetrics.dScale;
	//Kein Spaltenkopf?
	printCellParams.cellData.bColHdr = FALSE;
	//Baum-Richtung
	printCellParams.cellData.bTreeBottomUp = (printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_BOTTOM_UP) != 0;

	hsText = 0;
	iY = printCellsParams.iYStart;
	lRow = printCellsParams.lStartRow;
	iLastRow = -1;
	iPrinted = 0;
	bPageFull = FALSE;
	bPageBreak = FALSE;
	for (i = 0; i <= printCellsParams.iToCol; i++)
	{
		vecBNoLeftBorder.push_back(FALSE);
		vecBNoTopBorder.push_back(FALSE);
		vecBNoRightBorder.push_back(FALSE);
		vecBNoBottomBorder.push_back(FALSE);
	}

	//Indent-Breite Tree pro Level ermitteln
	iTreeIndentPerLevel = round(printCellsParams.tableInfo.iTreeIndentPerLevel / (double)(printCellsParams.tableInfo.iPixelPerInchX) * printCellsParams.printMetrics.iPixelPerInchX * printCellsParams.printMetrics.dScale);

	//Variable Zeilenhöhe?
	bVarRowHeight = MTblQueryFlags(printCellsParams.tableInfo.handle, MTBL_FLAG_VARIABLE_ROW_HEIGHT);

	while (Ctd.TblFindNextRow(printCellsParams.tableInfo.handle, &lRow, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff))
	{
		//Erste Zeile?
		if (lFirstRow == TBL_Error)
			lFirstRow = lRow;

		//Fetchen
		Ctd.TblFetchRow(printCellsParams.tableInfo.handle, lRow);

		//Context setzen
		Ctd.TblSetContext(printCellsParams.tableInfo.handle, lRow);

		//Zeilenhöhe ermitteln
		lRowHeight = PrintCellsGetRowHeight(lRow, printCellsParams);

		// Beginn Block?
		if (this->QueryRowFlags(lRow, MTBL_ROW_BLOCK_BEGIN))
		{
			BOOL bEndFound = FALSE;
			long lRow2 = lRow;
			long lRowsHeight = lRowHeight;

			// Zeilenhöhen bis Ende Block aufaddieren
			while (Ctd.TblFindNextRow(printCellsParams.tableInfo.handle, &lRow2, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff))
			{
				if (this->QueryRowFlags(lRow2, MTBL_ROW_BLOCK_BEGIN))
				{
					bEndFound = TRUE;
					break;
				}

				lRowsHeight += PrintCellsGetRowHeight(lRow2, printCellsParams);

				if (this->QueryRowFlags(lRow2, MTBL_ROW_BLOCK_END))
				{
					bEndFound = TRUE;
					break;
				}
			}

			if (bEndFound)
			{
				// Geht der Block über das Seitenende hinaus?
				if (iY + lRowsHeight > printCellsParams.iYMax)
				{
					// Seitenwechsel, wenn schon Zeilen gedruckt wurden
					if (iPrinted > 0)
					{
						bPageFull = TRUE;
						break;
					}
				}
			}
		}

		//Seite voll?
		if (iY + lRowHeight > printCellsParams.iYMax)
		{
			//Erste Zeile und Zellenhöhe > Höhe Zellenbereich?
			if (iPrinted == 0 && lRowHeight > (printCellsParams.iYMax - printCellsParams.iYStart + 1))
				lRowHeight = printCellsParams.iYMax - printCellsParams.iYStart + 1;
			else
			{
				bPageFull = TRUE;
				break;
			}
		}

		//Eventuell obere Linie ausgeben
		if (iPrinted == 0 && printCellsParams.printOptions.gridType == gtStandard && !printCellsParams.printOptions.bColHeaders)
		{
			HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
			HPEN hOldPen = (HPEN)SelectObject(printCellsParams.hdcCanvas, hPen);
			MoveToEx(printCellsParams.hdcCanvas, printCellsParams.printMetrics.clipRect.left, iY, NULL);
			LineTo(printCellsParams.hdcCanvas, printCellsParams.printMetrics.clipRect.left + iColws, iY);

			SelectObject(printCellsParams.hdcCanvas, hOldPen);
			DeleteObject(hPen);
		}

		//Daten für zu druckende Zellen ermitteln
		iX = printCellsParams.printMetrics.clipRect.left;

		for (i = printCellsParams.iFromCol; i <= printCellsParams.iToCol; i++)
		{
			lpCellToPrint = new CELLTOPRINT;
			hWndCol = printCellsParams.vecColsInfo.at(i).handle;

			//Merge-Zelle?
			bMergeCell = FALSE;
			lpMergeCell = (LPMTBLMERGECELL)MTblFindMergeCell(hWndCol, lRow, FMC_MASTER, printCellsParams.tableInfo.pMergeCells);
			if (lpMergeCell != NULL)
			{
				bMergeCell = TRUE;

				lpCellToPrint->lFromRow = lpMergeCell->lRowFrom;
				lpCellToPrint->lToRow = lpMergeCell->lRowTo;
				lpCellToPrint->lRows = 1;
				lpCellToPrint->hwFromCol = lpMergeCell->hWndColFrom;
				lpCellToPrint->hwToCol = lpMergeCell->hWndColTo;
				lpCellToPrint->lFromColPos = i;
				lpCellToPrint->lToColPos = i + lpMergeCell->iMergedCols;
				lpCellToPrint->lX = iX;
				lpCellToPrint->lY = iY;

				lpCellToPrint->lWidth = vecColWidth.at(i);
				if (lpMergeCell->iMergedCols > 0)
				{
					for (j = i + 1; j <= min(i + lpMergeCell->iMergedCols, printCellsParams.vecColsInfo.size() - 1); j++)
					{
						lpCellToPrint->lWidth += vecColWidth.at(j);
					}
				}

				lpCellToPrint->lHeight = lRowHeight;
				if (lpMergeCell->iMergedRows > 0)
				{
					lRow2 = lRow;
					while (Ctd.TblFindNextRow(printCellsParams.tableInfo.handle, &lRow2, printCellsParams.printOptions.dwColFlagsOn, printCellsParams.printOptions.dwColFlagsOff) && lRow2 <= lpMergeCell->lRowTo)
					{
						Ctd.TblFetchRow(printCellsParams.tableInfo.handle, lRow2);
						Ctd.TblSetContext(printCellsParams.tableInfo.handle, lRow2);

						lRowHeight2 = PrintCellsGetRowHeight(lRow2, printCellsParams);

						lpCellToPrint->lHeight += lRowHeight2;
						lpCellToPrint->lRows++;
					}
				}
			}

			//Normale oder Kind-Merge-Zelle?
			if (!bMergeCell)
			{
				lpMergeCell = (LPMTBLMERGECELL)MTblFindMergeCell(hWndCol, lRow, FMC_SLAVE, printCellsParams.tableInfo.pMergeCells);

				if (lpMergeCell != NULL)
				{
					lpCellToPrint->lFromRow = lpMergeCell->lRowFrom;
					lpCellToPrint->lToRow = lpMergeCell->lRowTo;
					lpCellToPrint->lRows = 1;
					lpCellToPrint->hwFromCol = lpMergeCell->hWndColFrom;
					lpCellToPrint->hwToCol = lpMergeCell->hWndColTo;
					lpCellToPrint->lFromColPos = i;

					if (lpMergeCell->iMergedCols > 0)
					{
						for (j = i - 1; j >= 0; j--)
						{
							if (printCellsParams.vecColsInfo.at(j).handle == lpMergeCell->hWndColFrom)
							{
								lpCellToPrint->lFromColPos = j;
								break;
							}
						}
					}

					lpCellToPrint->lToColPos = lpCellToPrint->lFromColPos + lpMergeCell->iMergedCols;
					lpCellToPrint->lX = iX;
					lpCellToPrint->lY = iY;

					lpCellToPrint->lWidth = vecColWidth.at(i);

					if (lpMergeCell->iMergedCols > 0)
					{
						/*
						Änderung auf max, wegen Fehler, dass unterschiedliche Ränder bei Mergezellen ausgegeben wurden.
						*/
						for (j = i - 1; j >= max(lpCellToPrint->lFromColPos, printCellsParams.iFromCol); j--)
						{
							lpCellToPrint->lWidth += vecColWidth.at(j);
							lpCellToPrint->lX -= vecColWidth.at(j);
						}

						for (j = i + 1; j <= min(lpCellToPrint->lToColPos, printCellsParams.vecColsInfo.size() - 1); j++)			//???
						{
							lpCellToPrint->lWidth += vecColWidth.at(j);
						}
					}

					lpCellToPrint->lHeight = lRowHeight;

					lRow2 = lRow;
					while (Ctd.TblFindPrevRow(printCellsParams.tableInfo.handle, &lRow2, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff) && lRow2 >= lpMergeCell->lRowFrom)
					{
						Ctd.TblFetchRow(printCellsParams.tableInfo.handle, lRow2);
						Ctd.TblSetContext(printCellsParams.tableInfo.handle, lRow2);

						lRowHeight2 = PrintCellsGetRowHeight(lRow2, printCellsParams);

						lpCellToPrint->lHeight += lRowHeight2;
						lpCellToPrint->lRows++;
						lpCellToPrint->lY -= lRowHeight2;
					}

					lRow2 = lRow;
					while (Ctd.TblFindNextRow(printCellsParams.tableInfo.handle, &lRow2, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff) && lRow2 <= lpMergeCell->lRowTo)
					{
						Ctd.TblFetchRow(printCellsParams.tableInfo.handle, lRow2);
						Ctd.TblSetContext(printCellsParams.tableInfo.handle, lRow2);

						lRowHeight2 = PrintCellsGetRowHeight(lRow2, printCellsParams);

						lpCellToPrint->lHeight += lRowHeight2;
						lpCellToPrint->lRows++;
					}
				}
				else if (i >= printCellsParams.iFromCol && i <= printCellsParams.iToCol)
				{
					lpCellToPrint->lFromRow = lRow;
					lpCellToPrint->lToRow = lRow;
					lpCellToPrint->lRows = 1;
					lpCellToPrint->hwFromCol = hWndCol;
					lpCellToPrint->hwToCol = hWndCol;
					lpCellToPrint->lFromColPos = i;
					lpCellToPrint->lToColPos = i;
					lpCellToPrint->lX = iX;
					lpCellToPrint->lY = iY;
					lpCellToPrint->lWidth = vecColWidth.at(i);
					lpCellToPrint->lHeight = lRowHeight;
				}
			}

			//Überspringen?
			if (lpCellToPrint == NULL)
			{
				iX += vecColWidth.at(i);
				continue;
			}

			//Handle effektive Spalte
			hWndEffCol = lpCellToPrint->hwFromCol;

			//Nummer effektive Zeile
			lEffRow = lpCellToPrint->lFromRow;

			//Level effektive Zeile
			iEffRowLevel = MTblGetRowLevel(printCellsParams.tableInfo.handle, lpCellToPrint->lFromRow);

			//Effektives i
			iEff = lpCellToPrint->lFromColPos;

			//Effektive Zeile fetchen
			Ctd.TblFetchRow(printCellsParams.tableInfo.handle, lEffRow);

			//Context auf effektive Zeile setzen
			Ctd.TblSetContext(printCellsParams.tableInfo.handle, lEffRow);

			//Ausgaberechteck
			printCellParams.cellData.rect.left = max(lpCellToPrint->lX, printCellsParams.printMetrics.clipRect.left);		//Änderung wegen unterschiedlichen Rändern
			printCellParams.cellData.rect.top = lpCellToPrint->lY;
			printCellParams.cellData.rect.right = printCellParams.cellData.rect.left + lpCellToPrint->lWidth;
			printCellParams.cellData.rect.bottom = printCellParams.cellData.rect.top + lpCellToPrint->lHeight;

			//Clipping-Rechteck
			printCellParams.cellData.clipRect.left = min(iX, printCellsParams.printMetrics.clipRect.right);

			printCellParams.cellData.clipRect.top = iY;
			printCellParams.cellData.clipRect.right = min(printCellParams.cellData.clipRect.left + vecColWidth.at(i)
				, min(printCellsParams.printMetrics.clipRect.left + iColws, printCellsParams.printMetrics.clipRect.right));
			printCellParams.cellData.clipRect.bottom = printCellParams.cellData.clipRect.top + lRowHeight;

			//Horizontale Textausrichtung
			printCellParams.cellData.iTextJustify = MTblGetEffCellTextJustify(hWndEffCol, lEffRow);

			//Vertikale Textausrichtung
			printCellParams.cellData.iTextVAlign = MTblGetEffCellTextVAlign(hWndEffCol, lEffRow);

			//Bild
			printCellParams.cellData.hImage = MTblGetEffCellImage(hWndEffCol, lEffRow);

			//Text
			Ctd.TblGetColumnText(printCellsParams.tableInfo.handle, printCellsParams.vecColsInfo.at(iEff).iID, &hsText);
			Ctd.HStrToTStr(hsText, printCellParams.cellData.tsText, TRUE);

			if (MTblQueryCellFlags(hWndEffCol, lEffRow, MTBL_CELL_FLAG_HIDDEN_TEXT))
			{
				printCellParams.cellData.tsText = tstring(_T(""));
			}
			else if (printCellsParams.vecColsInfo.at(iEff).bInvisible && printCellsParams.vecColsInfo.at(iEff).iCellType != COL_CellType_CheckBox)
			{
				tstring tsTmp(printCellParams.cellData.tsText.length(), printCellsParams.tableInfo.tcPasswordChar);
				printCellParams.cellData.tsText = tsTmp;
			}

			tstring tsFontName;

			//Font
			if (!PrintGetCellFont(hDeviceContext, printCellsParams.vecColsInfo.at(iEff), lEffRow,
				printCellsParams.printMetrics, printCellsParams.tableInfo,
				tsFontName, printCellParams.cellData.dFontHeight, printCellParams.cellData.lFontStyle))
				return iResult;

			printCellParams.cellData.tsFontName = tsFontName;

			//Fontfarbe
			MTblGetEffCellTextColor(hWndEffCol, lEffRow, &dwColor);
			printCellParams.cellData.crFontColor = dwColor;

			//Hintergrundfarbe
			if (!MTblGetEffCellBackColor(printCellsParams.vecColsInfo.at(iEff).handle, lEffRow, &printCellParams.cellData.crBackgroundColor))
				return iResult;

			//Texthöhe ermitteln
			HFONT hFont = PrintCreateFont(hDeviceContext, printCellParams.cellData.tsFontName, printCellParams.cellData.dFontHeight, printCellParams.cellData.lFontStyle);

			HFONT hOldFont = (HFONT)SelectObject(hDeviceContext, hFont);

			GetTextMetrics(hDeviceContext, &textMetric);
			iFontHeight = textMetric.tmAscent;

			SelectObject(hDeviceContext, hOldFont);
			DeleteObject(hFont);

			//Positionen und Größen
			if (printCellsParams.vecColsInfo.at(iEff).bTreeCol || printCellParams.cellData.hImage != 0)						//??? Warum treeCol abfragen?
			{
				if (printCellParams.cellData.hImage != 0)
				{
					MImgGetSize(printCellParams.cellData.hImage, &printCellParams.cellData.siImageSize);
					printCellParams.cellData.siImageSize.cx = round(printCellParams.cellData.siImageSize.cx / (double)(printCellsParams.tableInfo.iPixelPerInchX) * printCellsParams.printMetrics.iPixelPerInchX * printCellsParams.printMetrics.dScale) + 1;
					printCellParams.cellData.siImageSize.cy = round(printCellParams.cellData.siImageSize.cy / (double)(printCellsParams.tableInfo.iPixelPerInchY) * printCellsParams.printMetrics.iPixelPerInchY * printCellsParams.printMetrics.dScale);
				}
			}

			//Automatischer Zeilenumbruch?
			if (printCellsParams.printOptions.bFitCellSize)
				printCellParams.cellData.bWordBreak = printCellsParams.vecColsInfo.at(iEff).bWordWrap;
			else
				printCellParams.cellData.bWordBreak = (((printCellsParams.tableInfo.iLinesPerRow > 1) || (lpCellToPrint->lRows > 1) || bVarRowHeight) && printCellsParams.vecColsInfo.at(iEff).bWordWrap)
				|| MTblQueryFlags(printCellsParams.tableInfo.handle, MTBL_FLAG_EXPAND_CRLF);

			//Einzeilige Ausgabe?
			if (printCellsParams.printOptions.bFitCellSize)
				printCellParams.cellData.bSingleLine = !(printCellsParams.vecColsInfo.at(iEff).bWordWrap) && !(MTblQueryFlags(printCellsParams.tableInfo.handle, MTBL_FLAG_EXPAND_CRLF));
			else
				printCellParams.cellData.bSingleLine = (((printCellsParams.tableInfo.iLinesPerRow < 2) && (lpCellToPrint->lRows < 2) && !bVarRowHeight) || (!printCellsParams.vecColsInfo.at(iEff).bWordWrap)
				&& !MTblQueryFlags(printCellsParams.tableInfo.handle, MTBL_FLAG_EXPAND_CRLF));

			//Textausgabe-Flags setzen
			printCellParams.cellData.dwTextDrawFlags = DT_NOPREFIX | printCellParams.cellData.iTextJustify;
			if (printCellParams.cellData.bWordBreak)
				printCellParams.cellData.dwTextDrawFlags = printCellParams.cellData.dwTextDrawFlags | DT_WORDBREAK;
			if (printCellParams.cellData.bSingleLine)
				printCellParams.cellData.dwTextDrawFlags = printCellParams.cellData.dwTextDrawFlags | DT_SINGLELINE;
			if (MTblQueryFlags(printCellsParams.tableInfo.handle, MTBL_FLAG_EXPAND_TABS))
				printCellParams.cellData.dwTextDrawFlags = printCellParams.cellData.dwTextDrawFlags | DT_EXPANDTABS;

			//Textausgabe-Bereich setzen
			printCellParams.cellData.rTextArea.left = PrintCellsGetCellTextAreaLeft(printCellsParams, iEff, lEffRow, printCellParams, lpCellToPrint, iTxtOffsX);
			printCellParams.cellData.rTextArea.top = printCellParams.cellData.rect.top;
			printCellParams.cellData.rTextArea.right = PrintCellsGetCellTextAreaRight(printCellsParams, printCellParams, iEff, lEffRow);
			printCellParams.cellData.rTextArea.bottom = printCellParams.cellData.rect.bottom;

			//Textausgabe-Rechteck setzen
			PrintCellsGetTextRect(printCellsParams, printCellParams);

			//Bildausgabe-Rechteck setzen
			PrintCellsGetImageRect(printCellsParams, printCellParams, iEff, lEffRow, lpCellToPrint, iFontHeight);

			//Knoten
			printCellParams.cellData.bPaintNode = (printCellsParams.vecColsInfo.at(iEff).bTreeCol)
				&& (printCellsParams.tableInfo.iNodeSize > 0)
				&& (MTblQueryRowFlags(printCellsParams.tableInfo.handle, lEffRow, MTBL_ROW_CANEXPAND))
				&& !((iEffRowLevel == 0)
				&& ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0));

			if (printCellParams.cellData.bPaintNode)
			{
				//Knoten-Rechteck ermitteln
				iNodeWid = round(printCellsParams.tableInfo.iNodeSize / (double)(printCellsParams.tableInfo.iPixelPerInchX) * printCellsParams.printMetrics.iPixelPerInchX * printCellsParams.printMetrics.dScale);
				iNodeHt = round(printCellsParams.tableInfo.iNodeSize / (double)(printCellsParams.tableInfo.iPixelPerInchY) * printCellsParams.printMetrics.iPixelPerInchY * printCellsParams.printMetrics.dScale);

				if ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_FLAT_STRUCT) != 0)
					iNodeLevel = 0;
				else
					iNodeLevel = iEffRowLevel;

				if (iNodeLevel > 0 && ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0))
					iNodeLevel -= 1;

				PrintGetNodeRect(printCellParams.cellData.rect.left, printCellParams.cellData.rTextRect.top, iFontHeight, iNodeLevel, iTxtOffsX, iTxtOffsY, iNodeWid, iNodeHt, iTreeIndentPerLevel, printCellParams.cellData.nodeRect);

				//Knoten-Bilder ermitteln
				MTblGetNodeImages(printCellsParams.tableInfo.handle, &hImageNodeCol, &hImageNodeExp);

				//Knoten-Typ ermitteln, Knoten-Bild setzen
				if (MTblQueryRowFlags(printCellsParams.tableInfo.handle, lEffRow, MTBL_ROW_ISEXPANDED))
				{
					printCellParams.cellData.iNodeType = MTBL_NODE_MINUS;
					printCellParams.cellData.hNodeImageHandle = hImageNodeExp;
				}
				else
				{
					printCellParams.cellData.iNodeType = MTBL_NODE_PLUS;
					printCellParams.cellData.hNodeImageHandle = hImageNodeCol;
				}

				//Größe Knoten-Bild ermitteln
				if (printCellParams.cellData.hNodeImageHandle != 0)
				{
					MImgGetSize(printCellParams.cellData.hNodeImageHandle, &printCellParams.cellData.siNodeImageSize);
					printCellParams.cellData.siNodeImageSize.cx = PixTblToPrint(printCellParams.cellData.siNodeImageSize.cx, FALSE, printCellsParams.tableInfo, printCellsParams.printMetrics);
					printCellParams.cellData.siNodeImageSize.cy = PixTblToPrint(printCellParams.cellData.siNodeImageSize.cy, TRUE, printCellsParams.tableInfo, printCellsParams.printMetrics);
				}

				//Knoten-Farben
				if (printCellsParams.tableInfo.dwNodeOuterColor == MTBL_COLOR_UNDEF)
					dwColor = 0;
				else
					dwColor = printCellsParams.tableInfo.dwNodeOuterColor;
				printCellParams.cellData.crNodeOuterColor = dwColor;

				if (printCellsParams.tableInfo.dwNodeInnerColor != MTBL_COLOR_UNDEF)
					dwColor = printCellsParams.tableInfo.dwNodeInnerColor;
				printCellParams.cellData.crNodeInnerColor = dwColor;

				if (printCellsParams.tableInfo.dwNodeBackColor == MTBL_COLOR_UNDEF)
					printCellParams.cellData.crNodeBackColor = printCellParams.cellData.crBackgroundColor;
				else
					printCellParams.cellData.crNodeBackColor = printCellsParams.tableInfo.dwNodeBackColor;
			}

			//Tree-Linien
			printCellParams.cellData.bHorzTreeLine = FALSE;
			printCellParams.cellData.bVertTreeLineToParent = FALSE;
			printCellParams.cellData.bVertTreeLineToChild = FALSE;

			if (MTblQueryTreeLines(printCellsParams.tableInfo.handle, &printCellParams.cellData.iTreeLineStyle, &dwColor))
			{
				if (printCellsParams.vecColsInfo.at(iEff).bTreeCol && printCellParams.cellData.iTreeLineStyle != MTLS_INVISIBLE)
				{
					//Farbe
					if (dwColor == MTBL_COLOR_UNDEF)
						dwColor = RGB(0, 0, 0);
					printCellParams.cellData.crTreeLineColor = dwColor;

					//Zu zeichnende Tree-Linien ermitteln
					lParentRow = MTblGetParentRow(printCellsParams.tableInfo.handle, lEffRow);

					//Horizontal?
					if (iEffRowLevel == 0)
						printCellParams.cellData.bHorzTreeLine = (MTblQueryRowFlags(printCellsParams.tableInfo.handle, lEffRow, MTBL_ROW_CANEXPAND)
						&& (printCellsParams.tableInfo.iNodeSize > 0) && !((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0));
					else if (iEffRowLevel > 0)
						printCellParams.cellData.bHorzTreeLine = (lParentRow != TBL_Error)
						&& PrintCellsCheckRowFlags(printCellsParams.tableInfo, lParentRow, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff);

					if (printCellParams.cellData.bHorzTreeLine)
					{
						if ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_FLAT_STRUCT) != 0)
							iDrawLevel = 0;
						else
							iDrawLevel = iEffRowLevel;

						if (iDrawLevel > 0 && (printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0)
							iDrawLevel -= 1;

						printCellParams.cellData.rHTL_Pos.left = printCellParams.cellData.rect.left + (iTxtOffsX - 1) + (iDrawLevel * iTreeIndentPerLevel) + iTreeIndentPerLevel / 2.0;
						printCellParams.cellData.rHTL_Pos.top = printCellParams.cellData.rTextRect.top + iTxtOffsY + (iFontHeight / 2.0);

						if (printCellParams.cellData.hImage != 0 && printCellParams.cellData.rImageRect.left <= printCellParams.cellData.rTextRect.left)
							printCellParams.cellData.rHTL_Pos.right = printCellParams.cellData.rImageRect.left;
						else
							printCellParams.cellData.rHTL_Pos.right = printCellParams.cellData.rTextRect.left - 1;

						printCellParams.cellData.rHTL_Pos.bottom = printCellParams.cellData.rHTL_Pos.top;
					}

					//Vertikal zur Elternzeile?
					if (lParentRow != TBL_Error
						&& PrintCellsCheckRowFlags(printCellsParams.tableInfo, lParentRow, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff))
						printCellParams.cellData.bVertTreeLineToParent = TRUE;

					if (printCellParams.cellData.bVertTreeLineToParent)
						PrintCellsGetVertTreeLinesToParent(printCellsParams, printCellParams, iEffRowLevel, iTreeIndentPerLevel, iFontHeight, lpCellToPrint);

					//Verikal zur Kindezeile?
					lRow2 = MTblGetFirstChildRow(printCellsParams.tableInfo.handle, lEffRow);
					while (lRow2 != TBL_Error)
					{
						bChildRowMerged = (lRow2 >= lpCellToPrint->lFromRow) && (lRow2 <= lpCellToPrint->lToRow);
						if (PrintCellsCheckRowFlags(printCellsParams.tableInfo, lRow2, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff) && !bChildRowMerged)
						{
							printCellParams.cellData.bVertTreeLineToChild = TRUE;
							break;
						}

						lRow2 = MTblGetNextChildRow(printCellsParams.tableInfo.handle, lRow2);
					}

					if (printCellParams.cellData.bVertTreeLineToChild)
					{
						if ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0)
							iDrawLevel = 0;
						else
							iDrawLevel = iEffRowLevel + 1;

						if (iDrawLevel > 0 && ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0))
							iDrawLevel -= 1;

						printCellParams.cellData.rVTLTC_Pos.left = printCellParams.cellData.rect.left + (iTxtOffsX - 1) + (iDrawLevel * iTreeIndentPerLevel) + iTreeIndentPerLevel / 2.0;
						printCellParams.cellData.rVTLTC_Pos.right = printCellParams.cellData.rVTLTC_Pos.left;

						if (printCellParams.cellData.bTreeBottomUp)
						{
							printCellParams.cellData.rVTLTC_Pos.top = printCellParams.cellData.rect.top;

							if (printCellParams.cellData.bPaintNode && printCellParams.cellData.nodeRect.left <= printCellParams.cellData.rVTLTC_Pos.left)
								printCellParams.cellData.rVTLTC_Pos.bottom = printCellParams.cellData.nodeRect.top;
							else if (printCellParams.cellData.hImage != 0 && printCellParams.cellData.rImageRect.left <= printCellParams.cellData.rVTLTC_Pos.left)
								printCellParams.cellData.rVTLTC_Pos.bottom = printCellParams.cellData.rImageRect.top;
							else if (printCellParams.cellData.rTextRect.left <= printCellParams.cellData.rVTLTC_Pos.left)
								printCellParams.cellData.rVTLTC_Pos.bottom = printCellParams.cellData.rTextRect.top;
							else
								printCellParams.cellData.rVTLTC_Pos.bottom = printCellParams.cellData.rect.top + ((printCellParams.cellData.rect.bottom - printCellParams.cellData.rect.top) / 2.0);
						}
						else
						{
							if (printCellParams.cellData.bPaintNode && printCellParams.cellData.nodeRect.left <= printCellParams.cellData.rVTLTC_Pos.left)
								printCellParams.cellData.rVTLTC_Pos.top = printCellParams.cellData.nodeRect.bottom;
							else if (printCellParams.cellData.hImage != 0 && printCellParams.cellData.rImageRect.left <= printCellParams.cellData.rVTLTC_Pos.left)
								printCellParams.cellData.rVTLTC_Pos.top = printCellParams.cellData.rImageRect.bottom;
							else if (printCellParams.cellData.rTextRect.left <= printCellParams.cellData.rVTLTC_Pos.left)
							{
								rect = printCellParams.cellData.rTextRect;
								DrawText(hDeviceContext, printCellParams.cellData.tsText.c_str(), -1, &rect, printCellParams.cellData.dwTextDrawFlags | DT_CALCRECT);
								printCellParams.cellData.rVTLTC_Pos.top = rect.bottom;
							}
							else
								printCellParams.cellData.rVTLTC_Pos.top = printCellParams.cellData.rect.top + ((printCellParams.cellData.rect.bottom - printCellParams.cellData.rect.top) / 2.0);

							printCellParams.cellData.rVTLTC_Pos.bottom = printCellParams.cellData.rect.bottom + 1;
						}
					}
				}
			}
			//Ränder
			printCellParams.cellData.rBorders = rCellBorders;

			if (i < lpCellToPrint->lToColPos && i < printCellsParams.iToCol)
				printCellParams.cellData.rBorders.right = 0;

			if (i > lpCellToPrint->lFromColPos && i > printCellsParams.iFromCol)
				printCellParams.cellData.rBorders.left = 0;

			if (lRow < lpCellToPrint->lToRow)
				printCellParams.cellData.rBorders.bottom = 0;

			if (lRow > lpCellToPrint->lFromRow)
				printCellParams.cellData.rBorders.top = 0;

			if (vecBNoLeftBorder.at(i))
			{
				printCellParams.cellData.rBorders.left = 0;
				vecBNoLeftBorder.at(i) = FALSE;
			}

			if (vecBNoTopBorder.at(i))
			{
				printCellParams.cellData.rBorders.top = 0;
				vecBNoTopBorder.at(i) = FALSE;
			}

			if (vecBNoRightBorder.at(i))
			{
				printCellParams.cellData.rBorders.right = 0;
				vecBNoRightBorder.at(i) = FALSE;
			}

			if (vecBNoBottomBorder.at(i))
			{
				printCellParams.cellData.rBorders.bottom = 0;
				vecBNoBottomBorder.at(i) = FALSE;
			}

			if ((printCellsParams.vecColsInfo.at(i).bTreeCol && ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_NO_ROWLINES) != 0))
				|| printCellsParams.tableInfo.bNoRowLines
				|| printCellsParams.vecColsInfo.at(i).bNoRowLines
				|| MTblQueryRowFlags(printCellsParams.tableInfo.handle, lRow, MTBL_ROW_NO_ROWLINE)
				|| MTblQueryCellFlags(hWndEffCol, lEffRow, MTBL_CELL_FLAG_NO_ROWLINE))
			{
				printCellParams.cellData.rBorders.bottom = 0;
				vecBNoTopBorder.at(i) = TRUE;
			}

			//Unsichtbare Spaltenlinie letzte verankterte Spalte?
			bInvLastLockedColLine = FALSE;
			if (iEff < printCellsParams.iToCol && printCellsParams.vecColsInfo.at(iEff).bLocked && !printCellsParams.vecColsInfo.at(iEff + 1).bLocked)
			{
				if (MTblQueryLastLockedColLine(printCellsParams.tableInfo.handle, &k, &dwColor) && k == MTLS_INVISIBLE)
					bInvLastLockedColLine = TRUE;
			}

			if (i < printCellsParams.iToCol
				&& (printCellsParams.tableInfo.bNoColLines
				|| printCellsParams.vecColsInfo.at(i).bNoColLine
				|| bInvLastLockedColLine
				|| MTblQueryRowFlags(printCellsParams.tableInfo.handle, lRow, MTBL_ROW_NO_COLLINES)
				|| MTblQueryCellFlags(hWndEffCol, lEffRow, MTBL_CELL_FLAG_NO_COLLINE)))
			{
				printCellParams.cellData.rBorders.right = 0;
				vecBNoLeftBorder.at(i + 1) = TRUE;
			}

			//Benutzerdefiniert gezeichnet?
			if (MTblIsOwnerDrawnItem(hWndEffCol, lEffRow, MTBL_ODI_CELL))
			{
				printCellParams.cellData.odi.lType = MTBL_ODI_CELL;
				printCellParams.cellData.odi.hWndParam = hWndEffCol;
				printCellParams.cellData.odi.lParam = lEffRow;
				printCellParams.cellData.odi.hWndPart = hWndCol;
				printCellParams.cellData.odi.lPart = lRow;
				printCellParams.cellData.odi.hDC = printCellParams.hdcCanvas;
				printCellParams.cellData.odi.rectOdi.left = printCellParams.cellData.rect.left + 1;
				printCellParams.cellData.odi.rectOdi.top = printCellParams.cellData.rect.top + 1;
				printCellParams.cellData.odi.rectOdi.right = printCellParams.cellData.rect.right;
				printCellParams.cellData.odi.rectOdi.bottom = printCellParams.cellData.rect.bottom;
				printCellParams.cellData.odi.lSelected = 0;
				printCellParams.cellData.odi.lXLeading = printCellParams.cellData.iTextOffsetX;
				printCellParams.cellData.odi.lYLeading = printCellParams.cellData.iTextOffsetY;
				printCellParams.cellData.odi.lPrinting = 1;
				printCellParams.cellData.odi.dXResFactor = printCellsParams.printMetrics.iPixelPerInchX / (double)(printCellsParams.tableInfo.iPixelPerInchX) * printCellsParams.printMetrics.dScale;
				printCellParams.cellData.odi.dYResFactor = printCellsParams.printMetrics.iPixelPerInchY / (double)(printCellsParams.tableInfo.iPixelPerInchY) * printCellsParams.printMetrics.dScale;
			}
			else
				printCellParams.cellData.odi.lType = 0;

			//Änderung wegen der ungleichen Zellenränder bei Mergezellen
			if (i != lpCellToPrint->lFromColPos && printCellsParams.iFromCol > lpCellToPrint->lFromColPos)
			{
				//Mergezellen-Startspalte ist auf einer anderen Seite als die Mergezellenkinder, Zelle soll keinen Inhalt haben
				printCellParams.cellData.bHorzTreeLine = FALSE;
				printCellParams.cellData.bVertTreeLineToParent = FALSE;
				printCellParams.cellData.bVertTreeLineToChild = FALSE;
				printCellParams.cellData.bPaintNode = FALSE;
				printCellParams.cellData.hImage = 0;
				printCellParams.cellData.tsText = tstring(_T(""));
			}
			//Daten setzen
			lpCellToPrint->pcp = printCellParams;
			PrintCell(lpCellToPrint->pcp);

			delete lpCellToPrint;

			//X erhöhen
			iX += vecColWidth.at(i);
		}

		iY += lRowHeight;
		iLastRow = lRow;
		iPrinted++;

		//Seitenumbruch
		if (MTblQueryRowFlags(printCellsParams.tableInfo.handle, lRow, MTBL_ROW_PAGEBREAK))
		{
			bPageBreak = TRUE;
			break;
		}
	}

	//Eventuell abschließende Linien ausgeben
	if (printCellsParams.printOptions.gridType != gtNo)
	{
		HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		HPEN hOldPen = (HPEN)SelectObject(hDeviceContext, hPen);
		//Obere
		if (printCellsParams.printOptions.bColHeaders || (((printCellsParams.printOptions.gridType == gtStandard) || (printCellsParams.printOptions.gridType == gtStandard3)) && iPrinted > 0))
		{
			MoveToEx(hDeviceContext, printCellsParams.printMetrics.clipRect.left, printCellsParams.iYStart, NULL);
			LineTo(hDeviceContext, min(printCellsParams.printMetrics.clipRect.left + iColws, printCellsParams.printMetrics.clipRect.right), printCellsParams.iYStart);
		}
		//Untere
		if (iPrinted > 0 && printCellsParams.printOptions.gridType != gtStandard4)
		{
			MoveToEx(hDeviceContext, printCellsParams.printMetrics.clipRect.left, iY, NULL);
			LineTo(hDeviceContext, min(printCellsParams.printMetrics.clipRect.left + iColws, printCellsParams.printMetrics.clipRect.right), iY);
		}

		SelectObject(hDeviceContext, hOldPen);

		DeleteObject(hPen);
	}

	if (iPrinted > 0 && ((printCellsParams.printOptions.gridType == gtStandard2) || (printCellsParams.printOptions.gridType == gtStandard3)))
	{
		HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		HPEN hOldPen = (HPEN)SelectObject(hDeviceContext, hPen);

		//Linke
		MoveToEx(hDeviceContext, printCellsParams.printMetrics.clipRect.left, printCellsParams.iYStart, NULL);
		LineTo(hDeviceContext, printCellsParams.printMetrics.clipRect.left, iY);

		//Rechte
		MoveToEx(hDeviceContext, min(printCellsParams.printMetrics.clipRect.left + iColws, printCellsParams.printMetrics.clipRect.right), printCellsParams.iYStart, NULL);
		LineTo(hDeviceContext, min(printCellsParams.printMetrics.clipRect.left + iColws, printCellsParams.printMetrics.clipRect.right), iY);

		SelectObject(hDeviceContext, hOldPen);
		DeleteObject(hPen);
	}

	//Fertig
	iYBottom = iY;
	if (iPrinted > 0)
		lLastPrintedRow = iLastRow;
	else
		lLastPrintedRow = -1;

	if (bPageFull && iPrinted > 0)
		iResult = iLastRow;
	else
		iResult = -1; //Keine weiteren Zellen zu drucken

	if (bPageBreak)
	{
		lRow2 = lRow;
		if (Ctd.TblFindNextRow(printCellsParams.tableInfo.handle, &lRow2, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff))
		{
			Ctd.TblFetchRow(printCellsParams.tableInfo.handle, lRow2);
			iResult = iLastRow;
		}
		else
			iResult = -1;
	}

	return iResult;
}

/*
Druckt die Spaltenköpfe der Tabelle. Dabei werden die folgenden Schritt durchlaufen:
-Textoffsets ermitteln
-Höhe des Spaltenkopfbereichs ermitteln
-Parameter für den Druck der Zellen setzen(Bild, Text, Umrandung, Font, Hintergrund, ...)
-Gegebenenfalls Daten für Headergruppe ermitteln
-Zeichenbereiche ermitteln
-Gegebenenfalls benutzerdefiniert gezeichnetes setzen
-Zelle drucken
Parameter:
hDeviceContext:		Device Kontext, auf den gedruckt werden soll
colHeaderParams:	Parameter der Spaltenköpfe, die gedruckt werden sollen
Return:
Höhe des Spaltenkopfbereichs bei erfolgreichem Zeichnen.
-1, wenn ein Fehler aufgetreten ist.
*/
int MTBLSUBCLASS::PrintColHeaders(HDC hDeviceContext, PRINTCOLHEADERSPARAMS colHeaderParams)
{
	int i, j, iX, iY, iColWidth, iTitleHeight, iTextOffsetX, iTextOffsetY, iLines, iHeaderGrpHeight;
	RECT rect;
	int iHeaderGrpWidth, iTitleWidth;
	PRINTCELLPARAMS printCellParams;
	HSTRING hsFontName;
	tstring tsFontName, tsTmpFontName, tsHeaderGrpFontName, tsHeaderGrpText;

	long lFontSize, lLen, lHeaderGrpFontStyle;
	double dHeaderGrpFontHeight;

	COLORREF crColor;
	LPVOID pHeaderGroup;
	BOOL bFirstHeaderGrpCol, bHeaderGrpNoColHeaders, bHeaderGrpNoHorzLine, bHeaderGrpNoVertLines, bHeaderGrpTextAlignLeft, bHeaderGrpTextAlignRight, bLastHeaderGrpCol, bPrintHeaderGrp;
	int iResult = -1;

	LPTSTR lps;

	//Textoffsets ermitteln
	iTextOffsetX = round(colHeaderParams.tableInfo.metrics.ptCellLeading.x / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale);
	iTextOffsetY = round(colHeaderParams.tableInfo.metrics.ptCellLeading.y / (double)(colHeaderParams.tableInfo.iPixelPerInchY) * colHeaderParams.printMetrics.iPixelPerInchY * colHeaderParams.printMetrics.dScale);

	//Höhe des höchsten Titels ermitteln
	iTitleHeight = 0;
	colHeaderParams.hdcCanvas = hDeviceContext;

	HFONT hFont = PrintCreateFont(hDeviceContext, colHeaderParams.printMetrics.tsFontName, colHeaderParams.printMetrics.dFontHeight, colHeaderParams.printMetrics.lFontStyle);

	HFONT hOldFont = (HFONT)SelectObject(hDeviceContext, hFont);

	for (i = 0; i < colHeaderParams.vecColsInfo.size(); i++)
	{
		rect.left = 0;
		rect.top = 0;
		DrawText(colHeaderParams.hdcCanvas, colHeaderParams.vecColsInfo.at(i).tsTitle.c_str(), -1, &rect, DT_CALCRECT | DT_CENTER | DT_NOPREFIX);
		iTitleHeight = max(iTitleHeight, rect.bottom - rect.top);
	}

	SelectObject(hDeviceContext, hOldFont);
	DeleteObject(hFont);

	iTitleHeight += 2 * iTextOffsetY;

	//Fixe Parameter zum Druck der Zellen setzen
	//Canvas setzen
	printCellParams.hdcCanvas = colHeaderParams.hdcCanvas;

	//Tabellen-Handle setzen
	printCellParams.hwTable = colHeaderParams.tableInfo.handle;

	//Text Offset X setzen
	printCellParams.cellData.iTextOffsetX = iTextOffsetX;

	//Größe in Tabellenpixeln
	printCellParams.cellData.dTblPixSize = 1.0 / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale;

	//Keine Knoten im Spaltenkopf
	printCellParams.cellData.bPaintNode = FALSE;

	//Bild
	printCellParams.cellData.hImage = 0;

	//Clipping Rechteck setzen
	printCellParams.cellData.clipRect = colHeaderParams.printMetrics.clipRect;
	printCellParams.cellData.clipRect.bottom = colHeaderParams.iYMax;

	//Automatischer Zeilenumbruch?
	printCellParams.cellData.bWordBreak = FALSE;

	//Einzeilige Ausgabe?
	printCellParams.cellData.bSingleLine = FALSE;

	//Keine Baumlinien im Spaltenkopf
	printCellParams.cellData.bHorzTreeLine = FALSE;
	printCellParams.cellData.bVertTreeLineToChild = FALSE;
	printCellParams.cellData.bVertTreeLineToParent = FALSE;

	//Spaltenkopf?
	printCellParams.cellData.bColHdr = TRUE;

	//Titel drucken
	pHeaderGroup = NULL;
	iHeaderGrpHeight = 0;
	hsFontName = 0;

	iX = colHeaderParams.printMetrics.clipRect.left;
	iY = colHeaderParams.iYStart;

	for (i = colHeaderParams.iFromCol; i <= colHeaderParams.iToCol; i++)
	{
		//Header-Gruppe?
		if (colHeaderParams.vecColsInfo.at(i).HdrGrp != NULL)
		{
			if (i > 0 && colHeaderParams.vecColsInfo.at(i).bLocked != colHeaderParams.vecColsInfo.at(i - 1).bLocked)
				bPrintHeaderGrp = TRUE;
			else
				bPrintHeaderGrp = (pHeaderGroup != colHeaderParams.vecColsInfo.at(i).HdrGrp);

			pHeaderGroup = colHeaderParams.vecColsInfo.at(i).HdrGrp;

			//Erste/Letzte Spalte der Gruppe?
			bFirstHeaderGrpCol = bPrintHeaderGrp;
			bLastHeaderGrpCol = ((pHeaderGroup != NULL) && (i == colHeaderParams.iToCol || (i < colHeaderParams.iToCol && colHeaderParams.vecColsInfo.at(i + 1).HdrGrp != pHeaderGroup)));

			//Flags ermitteln
			bHeaderGrpNoVertLines = MTblQueryColHdrGrpFlags(colHeaderParams.tableInfo.handle, pHeaderGroup, MTBL_CHG_FLAG_NOINNERVERTLINES);
			bHeaderGrpNoHorzLine = MTblQueryColHdrGrpFlags(colHeaderParams.tableInfo.handle, pHeaderGroup, MTBL_CHG_FLAG_NOINNERHORZLINE);
			bHeaderGrpNoColHeaders = MTblQueryColHdrGrpFlags(colHeaderParams.tableInfo.handle, pHeaderGroup, MTBL_CHG_FLAG_NOCOLHEADERS);
			bHeaderGrpTextAlignLeft = MTblQueryColHdrGrpFlags(colHeaderParams.tableInfo.handle, pHeaderGroup, MTBL_CHG_FLAG_TXTALIGN_LEFT);
			bHeaderGrpTextAlignRight = MTblQueryColHdrGrpFlags(colHeaderParams.tableInfo.handle, pHeaderGroup, MTBL_CHG_FLAG_TXTALIGN_RIGHT);
		}
		else
		{
			pHeaderGroup = NULL;
			iHeaderGrpHeight = 0;
			bPrintHeaderGrp = FALSE;
			bHeaderGrpNoVertLines = FALSE;
			bHeaderGrpNoHorzLine = FALSE;
			bHeaderGrpNoColHeaders = FALSE;
			bHeaderGrpTextAlignLeft = FALSE;
			bHeaderGrpTextAlignRight = FALSE;
			bFirstHeaderGrpCol = FALSE;
			bLastHeaderGrpCol = FALSE;
		}

		//Daten der Header-Gruppe ermitteln
		if (bPrintHeaderGrp)
		{
			//Text
			GetColHeaderGrpText(colHeaderParams.tableInfo.handle, pHeaderGroup, tsHeaderGrpText);

			//Font
			if (!(MTblGetEffColHdrGrpFont(colHeaderParams.tableInfo.handle, pHeaderGroup, &hsFontName, &lFontSize, &lHeaderGrpFontStyle)))
				return iResult;

			Ctd.HStrToTStr(hsFontName, tsTmpFontName, TRUE);

			BOOL bFound = PrintFontExists(tsTmpFontName);
			if (bFound)
				tsFontName = tsTmpFontName;
			else
				tsFontName = colHeaderParams.printMetrics.tsFontName;

			tsHeaderGrpFontName = tsFontName;

			dHeaderGrpFontHeight = lFontSize * colHeaderParams.printMetrics.dScale;

			//Höhe
			//iLines sind die Anzahl der Reihen einer Spaltenkopfgruppe
			iLines = MTblGetColHdrGrpTextLineCount(colHeaderParams.tableInfo.handle, pHeaderGroup);

			if (iLines > 0)
			{
				HFONT hFont = PrintCreateFont(hDeviceContext, tsHeaderGrpFontName, dHeaderGrpFontHeight, lHeaderGrpFontStyle);
				HFONT hOldFont = (HFONT)SelectObject(hDeviceContext, hFont);

				DrawText(hDeviceContext, _T("A"), -1, &rect, DT_CALCRECT | DT_CENTER | DT_NOPREFIX);

				iHeaderGrpHeight = (iLines * (rect.bottom - rect.top)) + round(printCellParams.cellData.dTblPixSize * 2);

				SelectObject(hDeviceContext, hOldFont);
				DeleteObject(hFont);
			}
		}

		//Spaltenbreite in Druckerpixeln ermitteln
		iColWidth = round(colHeaderParams.vecColsInfo.at(i).iWidthPix / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale);

		//Parameter zum Druck der Zelle setzen
		//TextOffset Y
		if (pHeaderGroup != NULL && !bHeaderGrpNoColHeaders)
			printCellParams.cellData.iTextOffsetY = round(printCellParams.cellData.dTblPixSize);
		else
			printCellParams.cellData.iTextOffsetY = iTextOffsetY;

		//Ausgaberechteck setzen
		printCellParams.cellData.rect.left = iX;
		printCellParams.cellData.rect.top = iY + iHeaderGrpHeight;
		printCellParams.cellData.rect.right = iX + iColWidth;
		printCellParams.cellData.rect.bottom = iY + iTitleHeight;

		//Ränder setzen
		if (colHeaderParams.printOptions.gridType == gtNo)
			BorderFlagsFromGridType(gtNo, printCellParams.cellData.rBorders);
		else
			BorderFlagsFromGridType(gtTable, printCellParams.cellData.rBorders);

		if (bHeaderGrpNoHorzLine)
			printCellParams.cellData.rBorders.top = 0;

		if (bHeaderGrpNoVertLines)
		{
			if (!(bFirstHeaderGrpCol && bLastHeaderGrpCol))
			{
				if (!bLastHeaderGrpCol)
					printCellParams.cellData.rBorders.right = 0;

				if (!bFirstHeaderGrpCol)
					printCellParams.cellData.rBorders.left = 0;
			}
		}

		//Bild ermitteln
		MTblGetColumnHdrImage(colHeaderParams.vecColsInfo.at(i).handle, &printCellParams.cellData.hImage);

		//Text setzen
		printCellParams.cellData.tsText = colHeaderParams.vecColsInfo.at(i).tsTitle;

		//Font ermitteln
		hsFontName = 0;
		if (!MTblGetEffColumnHdrFont(colHeaderParams.vecColsInfo.at(i).handle, &hsFontName, &lFontSize, &(printCellParams.cellData.lFontStyle)))
			return iResult;

		Ctd.HStrToTStr(hsFontName, tsTmpFontName, TRUE);

		//Prüfe, ob Font vorhanden ist
		BOOL bFound = PrintFontExists(tsTmpFontName);
		if (bFound)
			tsFontName = tsTmpFontName;
		else
			tsFontName = colHeaderParams.printMetrics.tsFontName;

		printCellParams.cellData.tsFontName = tsFontName;

		printCellParams.cellData.dFontHeight = lFontSize * colHeaderParams.printMetrics.dScale;

		//Textfarbe ermitteln
		if (!MTblGetEffColumnHdrTextColor(colHeaderParams.vecColsInfo.at(i).handle, &crColor))
			return iResult;

		printCellParams.cellData.crFontColor = crColor;

		//Hintergrundfarbe ermitteln
		if (!MTblGetEffColumnHdrBackColor(colHeaderParams.vecColsInfo.at(i).handle, &crColor))
			return iResult;

		printCellParams.cellData.crBackgroundColor = crColor;

		//Titelbreite ermitteln
		HFONT hFont = PrintCreateFont(hDeviceContext, printCellParams.cellData.tsFontName, printCellParams.cellData.dFontHeight, printCellParams.cellData.lFontStyle);

		HFONT hOldFont = (HFONT)SelectObject(hDeviceContext, hFont);

		rect.top = 0;
		rect.left = 0;
		DrawText(hDeviceContext, printCellParams.cellData.tsText.c_str(), -1, &rect, DT_CALCRECT | DT_CENTER | DT_NOPREFIX);

		SelectObject(hDeviceContext, hOldFont);
		DeleteObject(hFont);

		iTitleWidth = rect.right - rect.left;

		//Bilddaten ermitteln
		if (printCellParams.cellData.hImage != 0)
		{
			MImgGetSize(printCellParams.cellData.hImage, &printCellParams.cellData.siImageSize);
			printCellParams.cellData.siImageSize.cx = round(printCellParams.cellData.siImageSize.cx / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale) + 1;
			printCellParams.cellData.siImageSize.cy = round(printCellParams.cellData.siImageSize.cy / (double)(colHeaderParams.tableInfo.iPixelPerInchY) * colHeaderParams.printMetrics.iPixelPerInchY * colHeaderParams.printMetrics.dScale);
			PrintColHeadersGetImageRect(FALSE, colHeaderParams, printCellParams, i, pHeaderGroup);
		}

		//Testausgabe-Bereich ermitteln
		printCellParams.cellData.rTextArea.left = PrintColHeadersGetTextAreaLeft(FALSE, colHeaderParams, printCellParams, i, pHeaderGroup);
		printCellParams.cellData.rTextArea.top = printCellParams.cellData.rect.top;
		printCellParams.cellData.rTextArea.right = PrintColHeadersGetTextAreaRight(FALSE, colHeaderParams, printCellParams, i, pHeaderGroup);
		printCellParams.cellData.rTextArea.bottom = printCellParams.cellData.rect.bottom;

		//Textausrichtung setzen
		if (MTblQueryColumnHdrFlags(colHeaderParams.vecColsInfo.at(i).handle, MTBL_COLHDR_FLAG_TXTALIGN_LEFT))
			printCellParams.cellData.iTextJustify = DT_LEFT;
		else if (MTblQueryColumnHdrFlags(colHeaderParams.vecColsInfo.at(i).handle, MTBL_COLHDR_FLAG_TXTALIGN_RIGHT))
			printCellParams.cellData.iTextJustify = DT_RIGHT;
		else if ((iTitleWidth + iTextOffsetX * 2) > (printCellParams.cellData.rTextArea.right - printCellParams.cellData.rTextArea.left))
			//Ausgabeverhalten der Tabelle simulieren: Wenn Textbreite + Ränder > Spaltenbreite, dann linksbündig ausgeben
			printCellParams.cellData.iTextJustify = DT_LEFT;
		else
			printCellParams.cellData.iTextJustify = DT_CENTER;

		//Textausgabe-Rechteck setzen
		printCellParams.cellData.rTextRect.left = printCellParams.cellData.rTextArea.left;
		if (printCellParams.cellData.iTextJustify == DT_LEFT)
			printCellParams.cellData.rTextRect.left += printCellParams.cellData.iTextOffsetX;
		printCellParams.cellData.rTextRect.top = printCellParams.cellData.rTextArea.top + printCellParams.cellData.iTextOffsetY;
		printCellParams.cellData.rTextRect.right = printCellParams.cellData.rTextArea.right;
		if (printCellParams.cellData.iTextJustify == DT_RIGHT)
			printCellParams.cellData.rTextRect.right -= printCellParams.cellData.iTextOffsetX;

		printCellParams.cellData.rTextRect.bottom = printCellParams.cellData.rTextArea.bottom;

		//Textausgabe-Flags setzen
		printCellParams.cellData.dwTextDrawFlags = DT_NOPREFIX | printCellParams.cellData.iTextJustify;
		if (MTblQueryFlags(colHeaderParams.tableInfo.handle, MTBL_FLAG_EXPAND_TABS))
			printCellParams.cellData.dwTextDrawFlags = printCellParams.cellData.dwTextDrawFlags | DT_EXPANDTABS;

		//Benutzerdefiniert gezeichnet?
		if ((MTblIsOwnerDrawnItem(colHeaderParams.vecColsInfo.at(i).handle, TBL_Error, MTBL_ODI_COLHDR)) ||
			(MTblIsOwnerDrawnItem(colHeaderParams.vecColsInfo.at(i).handle, TBL_Error, MTBL_ODI_COLUMN)))
		{
			printCellParams.cellData.odi.lType = MTBL_ODI_COLHDR;
			printCellParams.cellData.odi.hWndParam = colHeaderParams.vecColsInfo.at(i).handle;
			printCellParams.cellData.odi.lParam = TBL_Error;
			printCellParams.cellData.odi.hWndPart = colHeaderParams.vecColsInfo.at(i).handle;
			printCellParams.cellData.odi.lPart = TBL_Error;
			printCellParams.cellData.odi.hDC = printCellParams.hdcCanvas;
			printCellParams.cellData.odi.rectOdi.left = printCellParams.cellData.rect.left + 1;
			printCellParams.cellData.odi.rectOdi.top = printCellParams.cellData.rect.top + 1;
			printCellParams.cellData.odi.rectOdi.right = printCellParams.cellData.rect.right;
			printCellParams.cellData.odi.rectOdi.bottom = printCellParams.cellData.rect.bottom;
			printCellParams.cellData.odi.lSelected = 0;
			printCellParams.cellData.odi.lXLeading = printCellParams.cellData.iTextOffsetX;
			printCellParams.cellData.odi.lYLeading = printCellParams.cellData.iTextOffsetY;
			printCellParams.cellData.odi.lPrinting = 1;
			printCellParams.cellData.odi.dYResFactor = colHeaderParams.printMetrics.iPixelPerInchX / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.dScale;
			printCellParams.cellData.odi.dXResFactor = colHeaderParams.printMetrics.iPixelPerInchY / (double)(colHeaderParams.tableInfo.iPixelPerInchY) * colHeaderParams.printMetrics.dScale;
		}
		else
			printCellParams.cellData.odi.lType = 0;

		//Zelle drucken
		if (!bHeaderGrpNoColHeaders)
			PrintCell(printCellParams);

		//Header-Gruppe zeichnen
		if (bPrintHeaderGrp)
		{
			//Breite ermitteln
			iHeaderGrpWidth = round(colHeaderParams.vecColsInfo.at(i).iWidthPix / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale);
			for (j = i + 1; j <= colHeaderParams.iToCol; j++)
			{
				if (colHeaderParams.vecColsInfo.at(j).HdrGrp != colHeaderParams.vecColsInfo.at(i).HdrGrp)
					break;
				if (colHeaderParams.vecColsInfo.at(j).bLocked != colHeaderParams.vecColsInfo.at(i).bLocked)
					break;
				iHeaderGrpWidth += round(colHeaderParams.vecColsInfo.at(j).iWidthPix / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale);
			}

			//Ausgaberechteck setzen
			printCellParams.cellData.rect.left = iX;
			printCellParams.cellData.rect.top = iY;
			printCellParams.cellData.rect.right = iX + iHeaderGrpWidth;
			if (bHeaderGrpNoColHeaders)
				printCellParams.cellData.rect.bottom = iY + iTitleHeight;
			else
				printCellParams.cellData.rect.bottom = iY + iHeaderGrpHeight;

			//Ränder setzen
			if (colHeaderParams.printOptions.gridType == gtNo)
				BorderFlagsFromGridType(gtNo, printCellParams.cellData.rBorders);
			else
				BorderFlagsFromGridType(gtTable, printCellParams.cellData.rBorders);

			if (bHeaderGrpNoHorzLine && !bHeaderGrpNoColHeaders)
				printCellParams.cellData.rBorders.bottom = 0;

			//Text setzen
			printCellParams.cellData.tsText = tsHeaderGrpText;
			//Bild ermitteln
			if (!MTblGetColHdrGrpImage(colHeaderParams.tableInfo.handle, pHeaderGroup, &printCellParams.cellData.hImage))
				return iResult;

			//Bilddaten ermitteln
			if (printCellParams.cellData.hImage != 0)
			{
				MImgGetSize(printCellParams.cellData.hImage, &printCellParams.cellData.siImageSize);
				printCellParams.cellData.siImageSize.cx = round(printCellParams.cellData.siImageSize.cx / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.iPixelPerInchX * colHeaderParams.printMetrics.dScale) + 1;
				printCellParams.cellData.siImageSize.cy = round(printCellParams.cellData.siImageSize.cy / (double)(colHeaderParams.tableInfo.iPixelPerInchY) * colHeaderParams.printMetrics.iPixelPerInchY * colHeaderParams.printMetrics.dScale);
				PrintColHeadersGetImageRect(TRUE, colHeaderParams, printCellParams, i, pHeaderGroup);
			}

			//Titelbreite ermitteln
			rect.left = 0;
			rect.top = 0;
			DrawText(hDeviceContext, tsHeaderGrpText.c_str(), -1, &rect, DT_CALCRECT | DT_CENTER | DT_NOPREFIX);
			iTitleWidth = rect.right - rect.left;

			//Textausgabe-Bereich setzen
			printCellParams.cellData.rTextArea.left = PrintColHeadersGetTextAreaLeft(TRUE, colHeaderParams, printCellParams, i, pHeaderGroup);
			printCellParams.cellData.rTextArea.top = printCellParams.cellData.rect.top;
			printCellParams.cellData.rTextArea.right = PrintColHeadersGetTextAreaRight(TRUE, colHeaderParams, printCellParams, i, pHeaderGroup);
			printCellParams.cellData.rTextArea.bottom = printCellParams.cellData.rect.bottom;

			//Textausrichtung setzen
			if (bHeaderGrpTextAlignLeft)
				printCellParams.cellData.iTextJustify = DT_LEFT;
			else if (bHeaderGrpTextAlignRight)
				printCellParams.cellData.iTextJustify = DT_RIGHT;
			else if (iTitleWidth + (iTextOffsetX * 2) > (printCellParams.cellData.rTextArea.right - printCellParams.cellData.rTextArea.left))
				//Ausgabeverhalten der Tabelle simulieren: Wenn Textbreite + Ränder > Spaltenbreite, dann linksbündig ausgeben
				printCellParams.cellData.iTextJustify = DT_LEFT;
			else
				printCellParams.cellData.iTextJustify = DT_CENTER;

			//Textausgabe-Rechteck setzen
			printCellParams.cellData.rTextRect.left = printCellParams.cellData.rTextArea.left;
			if (printCellParams.cellData.iTextJustify == DT_LEFT)
				printCellParams.cellData.rTextRect.left += printCellParams.cellData.iTextOffsetX;
			printCellParams.cellData.rTextRect.top = printCellParams.cellData.rTextArea.top + printCellParams.cellData.iTextOffsetY;
			printCellParams.cellData.rTextRect.right = printCellParams.cellData.rTextArea.right;
			if (printCellParams.cellData.iTextJustify == DT_RIGHT)
				printCellParams.cellData.rTextRect.right -= printCellParams.cellData.iTextOffsetX;
			printCellParams.cellData.rTextRect.bottom = printCellParams.cellData.rTextArea.bottom;

			//Font
			printCellParams.cellData.tsFontName = tsHeaderGrpFontName;
			printCellParams.cellData.dFontHeight = dHeaderGrpFontHeight;
			printCellParams.cellData.lFontStyle = lHeaderGrpFontStyle;

			//Textfarbe
			if (!MTblGetEffColHdrGrpTextColor(colHeaderParams.tableInfo.handle, pHeaderGroup, &crColor))
				return iResult;
			printCellParams.cellData.crFontColor = crColor;

			//Hintergrundfarbe
			if (!MTblGetEffColHdrGrpBackColor(colHeaderParams.tableInfo.handle, pHeaderGroup, &crColor))
				return iResult;
			printCellParams.cellData.crBackgroundColor = crColor;

			//Textausgabe-Flags setzen
			printCellParams.cellData.dwTextDrawFlags = DT_NOPREFIX | printCellParams.cellData.iTextJustify;
			if (MTblQueryFlags(colHeaderParams.tableInfo.handle, MTBL_FLAG_EXPAND_TABS))
				printCellParams.cellData.dwTextDrawFlags = printCellParams.cellData.dwTextDrawFlags | DT_EXPANDTABS;

			//Benutzerdefiniert gezeichnet?
			if (MTblIsOwnerDrawnItem(colHeaderParams.tableInfo.handle, (long)(pHeaderGroup), MTBL_ODI_COLHDRGRP))
			{
				printCellParams.cellData.odi.lType = MTBL_ODI_COLHDRGRP;
				printCellParams.cellData.odi.hWndParam = 0;
				printCellParams.cellData.odi.lParam = (long)pHeaderGroup;
				printCellParams.cellData.odi.lPart = (long)pHeaderGroup;
				printCellParams.cellData.odi.hDC = printCellParams.hdcCanvas;
				printCellParams.cellData.odi.rectOdi.left = printCellParams.cellData.rect.left + 1;
				printCellParams.cellData.odi.rectOdi.top = printCellParams.cellData.rect.top + 1;
				printCellParams.cellData.odi.rectOdi.right = printCellParams.cellData.rect.right;
				printCellParams.cellData.odi.rectOdi.bottom = printCellParams.cellData.rect.bottom;
				printCellParams.cellData.odi.lSelected = 0;
				printCellParams.cellData.odi.lXLeading = printCellParams.cellData.iTextOffsetX;
				printCellParams.cellData.odi.lYLeading = printCellParams.cellData.iTextOffsetY;
				printCellParams.cellData.odi.lPrinting = 1;
				printCellParams.cellData.odi.dYResFactor = colHeaderParams.printMetrics.iPixelPerInchY / (double)(colHeaderParams.tableInfo.iPixelPerInchY) * colHeaderParams.printMetrics.dScale;
				printCellParams.cellData.odi.dXResFactor = colHeaderParams.printMetrics.iPixelPerInchX / (double)(colHeaderParams.tableInfo.iPixelPerInchX) * colHeaderParams.printMetrics.dScale;
			}
			else
				printCellParams.cellData.odi.lType = 0;

			//Header-Gruppe drucken
			PrintCell(printCellParams);
		}

		//X verschieben
		iX += iColWidth;
	}

	iResult = iTitleHeight;

	return iResult;
}

/*
Erstellt einen Font mit einer Gleitkommazahl als Höhe.
Parameter:
hDeviceContext:	Device Kontext, für den der Font erstellt werden soll
tsFontName:		Name des Fonts, der erstellt werden soll
dFontHeight:	Höhe des Fonts
lFontStyle:		Style des Fonts
Return:
Handle auf den erstellten Font, wenn er erfolgreich erstellt wird.
NULL, wenn der Device Kontext ungültig ist.
*/
HFONT MTBLSUBCLASS::PrintCreateFont(HDC hDeviceContext, tstring tsFontName, double dFontHeight, long lFontStyle)
{
	if (!hDeviceContext) return NULL;

	LOGFONT lf;

	// Logfont-Struktur initialisieren
	lf.lfHeight = 0;
	lf.lfWidth = 0;
	lf.lfEscapement = 0;
	lf.lfOrientation = 0;
	lf.lfWeight = FW_NORMAL;
	lf.lfItalic = FALSE;
	lf.lfUnderline = FALSE;
	lf.lfStrikeOut = FALSE;
	lf.lfCharSet = DEFAULT_CHARSET;
	lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
	lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	lf.lfQuality = DEFAULT_QUALITY;
	lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
	lf.lfFaceName[0] = 0;

	LPCTSTR lpsName = tsFontName.c_str();
	if (lpsName)
		lstrcpy(lf.lfFaceName, lpsName);

	lf.lfHeight = -round(dFontHeight / 72.0 * GetDeviceCaps(hDeviceContext, LOGPIXELSY));

	if (lFontStyle & FONT_EnhBold)
		lf.lfWeight = FW_BOLD;
	if (lFontStyle & FONT_EnhItalic)
		lf.lfItalic = TRUE;
	if (lFontStyle & FONT_EnhUnderline)
		lf.lfUnderline = TRUE;
	if (lFontStyle & FONT_EnhStrikeOut)
		lf.lfStrikeOut = TRUE;

	return CreateFontIndirect(&lf);
}

/*
Sucht den Index eines Fonts in der Liste der vom Drucker unterstützten Fonts.
Parameter:
tsFontName:	Name des Fonts, dessen Index gesucht werden soll
Return:
Index des Fonts, wenn er in der Liste vorhanden ist.
-1, wenn der Font nicht in der Liste ist.
*/
BOOL MTBLSUBCLASS::PrintFontExists(tstring tsFontName)
{
	tstring tsEnclose = _T("'");
	tstring tsSearch = tsEnclose + tsFontName + tsEnclose;

	size_t pos = m_tsPrintFonts.find(tsSearch);
	return pos != std::string::npos;
}

/*
Ermittelt den Text einer Headergruppe.
Parameter:
hWndTbl:	Handle für das Fenster der Tabelle
pHeaderGrp:	Pointer auf die Headergruppe, dessen Text ermittelt werden soll
tString:	String, in den der ermittelte Text geschrieben wird
Return:
TRUE, wenn der Text erfolgreich ermittelt wurde.
FALSE, wenn ein Fehler aufgetreten ist.
*/
BOOL MTBLSUBCLASS::GetColHeaderGrpText(HWND hWndTbl, LPVOID pHeaderGrp, tstring &tString)
{
	HSTRING hString = 0;
	BOOL bResult = FALSE;

	if (MTblGetColHdrGrpText(hWndTbl, pHeaderGrp, &hString))
		bResult = Ctd.HStrToTStr(hString, tString, TRUE);

	return bResult;
}

/*
Ermittelt den Zeichenbereich eines Bildes, das sich in einem Spaltenkopf befindet.
Parameter:
bHeaderGrp:			Headergruppe ja/nein
colHeaderParams:	Parameter des Spaltenkopfs, in dem sich das Bild befindet
printCellParams:	Parameter der Zelle, in die der Bereich des Bildes geschrieben wird
iColInfoIndex:		Spaltenindex der betrachteten Spalte
pHeaderGrp:			Pointer auf die Headergruppe
*/
void MTBLSUBCLASS::PrintColHeadersGetImageRect(BOOL bHeaderGrp, PRINTCOLHEADERSPARAMS colHeaderParams, PRINTCELLPARAMS &printCellParams, int iColInfoIndex, LPVOID pHeaderGrp)
{
	BOOL bAlignBottom, bAlignHorzCenter, bAlignLeft, bAlignRight, bAlignTile, bAlignTop, bAlignVertCenter;
	BOOL bNoLeadLeft, bNoLeadTop;
	long lBottom, lCellBottom, lCellLeft, lCellRight, lCellTop, lLeft, lRight, lTop;
	TEXTMETRIC textMetric;

	//Image Flags ermitteln
	if (!bHeaderGrp)
	{
		bAlignTile = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_TILE);
		bAlignHorzCenter = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_HCENTER);
		bAlignRight = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_RIGHT);
		bAlignLeft = (!bAlignHorzCenter && !bAlignRight);
		bAlignTop = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_TOP);
		bAlignVertCenter = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_VCENTER);
		bAlignBottom = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_BOTTOM);
		bNoLeadLeft = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_NOLEAD_LEFT);
		bNoLeadTop = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_NOLEAD_TOP);
	}
	else
	{
		bAlignTile = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_TILE);
		bAlignHorzCenter = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_HCENTER);
		bAlignRight = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_RIGHT);
		bAlignLeft = (!bAlignHorzCenter && !bAlignRight);
		bAlignTop = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_TOP);
		bAlignVertCenter = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_VCENTER);
		bAlignBottom = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_BOTTOM);
		bNoLeadLeft = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_NOLEAD_LEFT);
		bNoLeadTop = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_NOLEAD_TOP);
	}

	lCellTop = printCellParams.cellData.rect.top;
	lCellLeft = printCellParams.cellData.rect.left;
	lCellRight = printCellParams.cellData.rect.right;
	lCellBottom = printCellParams.cellData.rect.bottom;

	lLeft = lCellLeft;
	if (!bNoLeadLeft)
		lLeft += printCellParams.cellData.iTextOffsetX;
	else
		lLeft++;

	lTop = printCellParams.cellData.rect.top;
	if (!bNoLeadTop)
		lTop += printCellParams.cellData.iTextOffsetY;

	lRight = printCellParams.cellData.rect.right;
	lBottom = printCellParams.cellData.rect.bottom;

	//Image Zeichenbereich setzen
	if (bAlignTile)
	{
		printCellParams.cellData.rImageRect.left = lLeft;
		printCellParams.cellData.rImageRect.top = lTop;
		printCellParams.cellData.rImageRect.right = lRight;
		printCellParams.cellData.rImageRect.bottom = lBottom;
	}
	else
	{
		if (bAlignLeft)
			printCellParams.cellData.rImageRect.left = lLeft;
		else if (bAlignHorzCenter)
			printCellParams.cellData.rImageRect.left = max(lLeft, lCellLeft + ((lCellRight - lCellLeft - printCellParams.cellData.siImageSize.cx) / 2.0));
		else
			printCellParams.cellData.rImageRect.left = max(lLeft, lCellRight - printCellParams.cellData.siImageSize.cx);

		if (bAlignTop)
			printCellParams.cellData.rImageRect.top = lTop;
		else if (bAlignVertCenter)
			printCellParams.cellData.rImageRect.top = max(lTop, lCellTop + ((lCellBottom - lCellTop - printCellParams.cellData.siImageSize.cy) / 2.0));
		else if (bAlignBottom)
			printCellParams.cellData.rImageRect.top = max(lTop, lCellBottom - printCellParams.cellData.siImageSize.cy);
		else
		{
			GetTextMetrics(colHeaderParams.hdcCanvas, &textMetric);
			printCellParams.cellData.rImageRect.top = max(lTop, lTop + ((textMetric.tmHeight - printCellParams.cellData.siImageSize.cy) / 2.0));
		}

		printCellParams.cellData.rImageRect.right = printCellParams.cellData.rImageRect.left + printCellParams.cellData.siImageSize.cx;
		printCellParams.cellData.rImageRect.bottom = printCellParams.cellData.rImageRect.top + printCellParams.cellData.siImageSize.cy;
	}
}

/*
Ermittelt den linken Abstand des Zeichenbereichs für die Ausgabe von Text in einem Spaltenkopf. Dabei wird ein Textoffset und die Breite des Bildes einbezogen.
Parameter:
bHeaderGrp:			Headergruppe ja/nein
colHeaderParams:	Parameter des betrachteten Spaltenkopfs
printCellParams:	Parameter der betrachteten Zelle
iColInfoIndex:		Index der betrachteten Spalte
pHeaderGrp:			Pointer auf die zugehörige Headergruppe
Return:
Linken Abstand des Zeichenbereichs für die Ausgabe von Text.
*/
long MTBLSUBCLASS::PrintColHeadersGetTextAreaLeft(BOOL bHeaderGrp, PRINTCOLHEADERSPARAMS colHeaderParams, PRINTCELLPARAMS printCellParams, int iColInfoIndex, LPVOID pHeaderGrp)
{
	BOOL bAlignHorzCenter, bAlignRight, bAlignTile, bNoLeadLeft;
	long lLeft, lOffset;

	lLeft = printCellParams.cellData.rect.left + 1;
	if (printCellParams.cellData.hImage != 0)
	{
		//Bild vorhanden, Image Flags ermitteln
		if (!bHeaderGrp)
		{
			bAlignTile = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_TILE);
			bAlignHorzCenter = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_HCENTER);
			bAlignRight = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_RIGHT);
			bNoLeadLeft = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_NOLEAD_LEFT);
		}
		else
		{
			bAlignTile = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_TILE);
			bAlignHorzCenter = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_HCENTER);
			bAlignRight = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_RIGHT);
			bNoLeadLeft = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_NOLEAD_LEFT);
		}

		if (!bAlignHorzCenter && !bAlignRight && !bAlignTile)
		{
			if (bNoLeadLeft)
				lOffset = 0;
			else
				lOffset = printCellParams.cellData.iTextOffsetX;

			lLeft += lOffset + printCellParams.cellData.siImageSize.cx + 1;
		}
	}

	return lLeft;
}

/*
Ermittelt den rechten Abstand des Zeichenbereichs für die Ausgabe von Text in einem Spaltenkopf. Dabei wird die Breite des Bildes einbezogen.
Parameter:
bHeaderGrp:			Headergruppe ja/nein
colHeaderParams:	Parameter des betrachteten Spaltenkopfs
printCellParams:	Parameter der betrachteten Zelle
iColInfoIndex:		Index der betrachteten Spalte
pHeaderGrp:			Pointer auf die zugehörige Headergruppe
Return:
Rechten Abstand des Zeichenbereichs für die Ausgabe von Text.
*/
long MTBLSUBCLASS::PrintColHeadersGetTextAreaRight(BOOL bHeaderGrp, PRINTCOLHEADERSPARAMS colHeaderParams, PRINTCELLPARAMS printCellParams, int iColInfoIndex, LPVOID pHeaderGrp)
{
	BOOL bAlignRight, bAlignTile;
	long lRight;

	lRight = printCellParams.cellData.rect.right;

	if (printCellParams.cellData.hImage != 0)
	{
		//Bild vorhanden, Image Flags ermitteln
		if (!bHeaderGrp)
		{
			bAlignTile = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_TILE);
			bAlignRight = MTblQueryColumnHdrImageFlags(colHeaderParams.vecColsInfo.at(iColInfoIndex).handle, MTSI_ALIGN_RIGHT);
		}
		else
		{
			bAlignTile = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_TILE);
			bAlignRight = MTblQueryColHdrGrpImageFlags(colHeaderParams.tableInfo.handle, pHeaderGrp, MTSI_ALIGN_RIGHT);
		}

		if (bAlignRight && !bAlignTile)
			lRight -= (printCellParams.cellData.siImageSize.cx + 1);
	}

	return lRight;
}

/*
Ermittelt die beste Zellenhöhe für eine Zelle.
Parameter:
hWndCol:	Handle für die verwendete Spalte
lRow:		Zeile, in der sich die Zelle befindet
printCellsParams:	Parameter der Zellen
Return:
Die beste Höhe für die übergebene Zelle.
-1, wenn ein Fehler auftritt.
*/
int MTBLSUBCLASS::PrintCellsGetBestCellHeight(HWND hWndCol, long lRow, PRINTCELLSPARAMS printCellsParams)
{
	HSTRING hsFontName = 0;
	long lFontStyle, lSize;
	tstring tsFontName;
	int iResult;

	MTblGetEffCellFont(hWndCol, lRow, &hsFontName, &lSize, &lFontStyle);
	Ctd.HStrToTStr(hsFontName, tsFontName, TRUE);

	BOOL bFound = PrintFontExists(tsFontName);
	if (!bFound)
		tsFontName = printCellsParams.printMetrics.tsFontName;

	HDC hTableContext = GetDC(printCellsParams.tableInfo.handle);
	HFONT hPrintFont = PrintCreateFont(hTableContext, tsFontName, lSize, lFontStyle);

	//Beste Zellenhöhe ermitteln
	iResult = MTblGetBestCellHeight(printCellsParams.tableInfo.handle, hWndCol, lRow,
		printCellsParams.tableInfo.metrics.ptCellLeading.y, printCellsParams.tableInfo.metrics.ptCharBox.y,
		hPrintFont, printCellsParams.tableInfo.pMergeCells, TRUE, TRUE);

	if (iResult > 0)
		iResult = PixTblToPrint(iResult, TRUE, printCellsParams.tableInfo, printCellsParams.printMetrics);

	DeleteObject(hPrintFont);
	ReleaseDC(printCellsParams.tableInfo.handle, hTableContext);

	return iResult;
}

/*
Ermittelt die Zeilenhöhe einer Zeile.
Ist FitCellSize nicht gesetzt, wird die in der Tabelle verwendete Zeilenhöhe für die Zeile verwendet.
Parameter:
lRow:	Zeile, für die die Höhe ermittelt werden soll
printCellsParams:	Parameter der Zellen, wird verwendet, um die Höhe der Zellen zu bestimmen
Return:
Die Höhe der übergebenen Zeile.
*/
long MTBLSUBCLASS::PrintCellsGetRowHeight(long lRow, PRINTCELLSPARAMS printCellsParams)
{
	int i, iMaxHeight, iMinHeight;
	long lResult = 0;

	if (printCellsParams.printOptions.bFitCellSize)
	{
		//Fit Cell Size gesetzt, ermittle beste Zeilenhöhe
		for (i = 0; i < printCellsParams.vecColsInfo.size(); i++)
		{
			lResult = max(lResult, PrintCellsGetBestCellHeight(printCellsParams.vecColsInfo.at(i).handle, lRow, printCellsParams));
		}

		iMinHeight = SendMessage(printCellsParams.tableInfo.handle, MTM_QueryMinRowHeight, lRow, lResult);
		iMaxHeight = SendMessage(printCellsParams.tableInfo.handle, MTM_QueryMaxRowHeight, lRow, lResult);

		if (iMinHeight > 0)
		{
			iMinHeight = PixTblToPrint(iMinHeight, TRUE, printCellsParams.tableInfo, printCellsParams.printMetrics);
			if (iMinHeight > lResult)
				lResult = iMinHeight;
		}

		if (iMaxHeight > 0)
		{
			iMaxHeight = PixTblToPrint(iMaxHeight, TRUE, printCellsParams.tableInfo, printCellsParams.printMetrics);
			if (iMaxHeight < lResult)
				lResult = iMaxHeight;
		}
	}

	if (lResult == 0)
	{
		//Übernehme Zeilenhöhe aus der Tabelle
		MTblGetEffRowHeight(printCellsParams.tableInfo.handle, lRow, &lResult);
		lResult = PixTblToPrint(lResult, TRUE, printCellsParams.tableInfo, printCellsParams.printMetrics);
	}

	return lResult;
}

/*
Ermittelt den Font einer Zelle.
Der Name des Fonts wird in den übergebenen String geschrieben, die Höhe und der Style des Fonts in die entsprechenden übergebenen Parameter.
Parameter:
hDeviceContext:	Device Kontext des Druckers
colInfo:		Informationen über die Spalte der Zelle
lRow:			Zeile der Zelle
printMetrics:	Druckmetriken, um den Font zu definieren
Return:
TRUE, wenn der Font erfolgreich ermittelt wurde.
FALSE, wenn ein Fehler aufgetreten ist.
*/
BOOL MTBLSUBCLASS::PrintGetCellFont(HDC hDeviceContext, COLINFO colInfo, long lRow, PRINTMETRICS printMetrics, TABLEINFO tableInfo, tstring &tsOutFontName, double &dFontHeight, long &lFontStyle)
{
	HSTRING hsFontName = 0;
	tstring tsTmpFontName, tsFontName;
	long lSize, lStyle;
	BOOL bResult = FALSE;

	if (!MTblGetEffCellFont(colInfo.handle, lRow, &hsFontName, &lSize, &lStyle))
		return bResult;

	Ctd.HStrToTStr(hsFontName, tsTmpFontName, TRUE);

	BOOL bFound = PrintFontExists(tsTmpFontName);
	if (!bFound)
		tsFontName = printMetrics.tsFontName;
	else
		tsFontName = tsTmpFontName;

	tsOutFontName = tsFontName;
	dFontHeight = lSize * printMetrics.dScale;
	lFontStyle = lStyle;

	bResult = TRUE;

	return bResult;
}

/*
Ermittelt die Einrückung einer Zelle.
Parameter:
printCellsParams: Parameter der Zellen
iColIndex:	Spaltenindex der betrachteten Zelle
lRow:		Zeile der betrachteten Zelle
Return:
Die Einrückung der Zelle.
*/
long MTBLSUBCLASS::PrintCellsGetEffCellIndent(PRINTCELLSPARAMS printCellsParams, int iColIndex, long lRow)
{
	long lIndent;

	lIndent = 0;
	if (MTblGetEffCellIndent(printCellsParams.vecColsInfo.at(iColIndex).handle, lRow, &lIndent))
		lIndent = PixTblToPrint(lIndent, FALSE, printCellsParams.tableInfo, printCellsParams.printMetrics);

	return lIndent;
}

/*
Ermittelt den linken Abstand des Zeichenbereichs einer Zelle.
Parameter:
printCellsParams:	Parameter der Zellen
iEffIndex:		Spaltenindex der betrachteten Zelle
lEffRow:		Zeile der betrachteten Zelle
lpCellToPrint:	Zu druckende Zelle, verwendet für die Koordinate der Zelle
Return:
Linken Abstand des Zeichenbereichs einer Zelle.
*/
long MTBLSUBCLASS::PrintCellsGetCellAreaLeft(PRINTCELLSPARAMS printCellsParams, int iEffIndex, long lEffRow, LPCELLTOPRINT lpCellToPrint)
{
	long lEffIndent;

	lEffIndent = PrintCellsGetEffCellIndent(printCellsParams, iEffIndex, lEffRow);
	if (lEffIndent > 0)
		return lpCellToPrint->lX + lEffIndent;

	return lpCellToPrint->lX + 1;
}

/*
Ermittelt den linken Abstand des Zeichenbereichs für den Text einer Zelle. Dabei wird das Bild in einer Zelle berücksichtigt.
Parameter:
printCellsParams:	Parameter der Zellen
iEffIndex:			Spaltenindex der betrachteten Zelle
lEffRow:			Zeile der betrachteten Zelle
printCellParams:	Parameter der Zelle, verwendet für Einbeziehung eines Zellenbildes
lpCellToPrint:		Zu druckende Zelle
iTextOffsetX:		Offset der Zelle
Return:
Linken Abstand des Zeichenbereichs für den Text einer Zelle.
*/
long MTBLSUBCLASS::PrintCellsGetCellTextAreaLeft(PRINTCELLSPARAMS printCellsParams, int iEffIndex, long lEffRow, PRINTCELLPARAMS printCellParams, LPCELLTOPRINT lpCellToPrint, int iTextOffsetX)
{
	BOOL bAlignHCenter, bAlignRight, bAlignTile, bIndented, bNoLeadLeft;
	long lLeft, lOffset;

	lLeft = PrintCellsGetCellAreaLeft(printCellsParams, iEffIndex, lEffRow, lpCellToPrint);

	//Image Flags ermitteln
	bAlignHCenter = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_HCENTER);
	bAlignRight = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_RIGHT);
	bAlignTile = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_TILE);

	if (printCellParams.cellData.hImage != 0 && !bAlignHCenter && !bAlignRight && !bAlignTile)
	{
		bNoLeadLeft = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_NOLEAD_LEFT);
		bIndented = PrintCellsGetEffCellIndent(printCellsParams, iEffIndex, lEffRow) > 0;

		if (bNoLeadLeft && !bIndented)
			lOffset = 0;
		else
			lOffset = iTextOffsetX;

		lLeft += lOffset + printCellParams.cellData.siImageSize.cx + 1;
	}

	return lLeft;
}

/*
Ermittelt den rechten Abstand des Zeichenbereichs für den Text einer Zelle. Dabei wird das Bild in einer Zelle berücksichtigt.
Parameter:
printCellsParams:	Parameter der Zellen
printCellParams:	Parameter der Zelle
iEffIndex:			Spaltenindex der betrachteten Zelle
lEffRow:			Zeile der betrachteten Zelle
Return:
Rechten Abstand des Zeichenbereichs für den Text einer Zelle.
*/
long MTBLSUBCLASS::PrintCellsGetCellTextAreaRight(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS printCellParams, int iEffIndex, long lEffRow)
{
	BOOL bAlignRight, bAlignTile;
	long lRight;
	lRight = printCellParams.cellData.rect.right;
	bAlignRight = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_RIGHT);
	bAlignTile = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_TILE);
	if (printCellParams.cellData.hImage != 0 && bAlignRight && !bAlignTile)
		lRight -= printCellParams.cellData.siImageSize.cx + 1;

	return lRight;
}

/*
Ermittelt den Zeichenbereich für den Text einer übergebenen Zelle. Der resultierende Zeichenbereich wird direkt in die übergebene Zelle an die Stelle des TextRects geschrieben.
Parameter:
printCellsParams:	Parameter der Zellen
printCellParams:	Parameter der Zelle, wird mit dem Zeichenbereich des Textes gefüllt
*/
void MTBLSUBCLASS::PrintCellsGetTextRect(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS &printCellParams)
{
	int iCellHeight, iTextHeight;
	RECT rect;
	DWORD dwDrawFlags;

	printCellParams.cellData.rTextRect.left = printCellParams.cellData.rTextArea.left;
	if (printCellParams.cellData.iTextJustify == DT_LEFT || !printCellParams.cellData.bSingleLine)
		printCellParams.cellData.rTextRect.left += printCellParams.cellData.iTextOffsetX;
	else
		printCellParams.cellData.rTextRect.left++;

	printCellParams.cellData.rTextRect.top = printCellParams.cellData.rTextArea.top + printCellParams.cellData.iTextOffsetY;
	printCellParams.cellData.rTextRect.right = printCellParams.cellData.rTextArea.right;

	if (printCellParams.cellData.iTextJustify == DT_RIGHT || !printCellParams.cellData.bSingleLine)
		printCellParams.cellData.rTextRect.right -= printCellParams.cellData.iTextOffsetX;

	printCellParams.cellData.rTextRect.bottom = printCellParams.cellData.rTextArea.bottom;

	if (printCellParams.cellData.iTextVAlign != DT_TOP)
	{
		dwDrawFlags = printCellParams.cellData.dwTextDrawFlags | DT_CALCRECT;

		rect.left = printCellParams.cellData.rTextRect.left;
		rect.top = 0;
		rect.right = printCellParams.cellData.rTextRect.right;
		rect.bottom = 0;

		HFONT hFont = PrintCreateFont(printCellsParams.hdcCanvas, printCellParams.cellData.tsFontName, printCellParams.cellData.dFontHeight, printCellParams.cellData.lFontStyle);
		HFONT hOldFont = (HFONT)SelectObject(printCellsParams.hdcCanvas, hFont);
		COLORREF crOldColor = SetTextColor(printCellsParams.hdcCanvas, printCellParams.cellData.crFontColor);

		//Text Rect ermitteln
		if (printCellParams.cellData.tsText.empty())
			DrawText(printCellsParams.hdcCanvas, _T("A"), -1, &rect, dwDrawFlags);
		else
			DrawText(printCellsParams.hdcCanvas, printCellParams.cellData.tsText.c_str(), -1, &rect, dwDrawFlags);

		iTextHeight = rect.bottom - rect.top;
		iCellHeight = printCellParams.cellData.rect.bottom - printCellParams.cellData.rect.top;

		if (iCellHeight > iTextHeight)
		{
			if (printCellParams.cellData.iTextVAlign == DT_VCENTER)
				printCellParams.cellData.rTextRect.top = max(printCellParams.cellData.rTextRect.top, printCellParams.cellData.rect.top + (iCellHeight - iTextHeight) / 2.0);
			else
				printCellParams.cellData.rTextRect.top = printCellParams.cellData.rect.bottom - printCellParams.cellData.iTextOffsetY - iTextHeight;
		}

		SelectObject(printCellsParams.hdcCanvas, hOldFont);
		DeleteObject(hFont);
	}
}

/*
Ermittelt den Zeichenbereich für das Bild einer Zelle.
Parameter:
printCellsParams:	Parameter der Zellen
printCellParams:	Parameter der Zelle, wird mit dem Zeichenbereich des Bildes gefüllt
iEffIndex:			Spaltenindex der betrachteten Zelle
lEffRow:			Zeile der betrachteten Zelle
lpCellToPrint:		Zu druckende Zelle
iFontHeight:		Font Höhe
Der resultierende Zeichenbereich wird direkt in die übergebenen Zellenparameter an die Stelle des ImageRects geschrieben.
*/
void MTBLSUBCLASS::PrintCellsGetImageRect(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS &printCellParams, int iEffIndex, long lEffRow, LPCELLTOPRINT lpCellToPrint, int iFontHeight)
{
	BOOL bAlignBottom, bAlignHorzCenter, bAlignLeft, bAlignRight, bAlignTile, bAlignTop, bAlignVertCenter;
	BOOL bNoLeadLeft, bNoLeadTop;
	long lBottom, lCellBottom, lCellLeft, lCellRight, lCellTop, lLeft, lRight, lTop;

	//Image Flags ermitteln
	bAlignTile = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_TILE);
	bAlignHorzCenter = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_HCENTER);
	bAlignRight = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_RIGHT);
	bAlignLeft = !bAlignHorzCenter && !bAlignRight;

	bAlignTop = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_TOP);
	bAlignVertCenter = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_VCENTER);
	bAlignBottom = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_ALIGN_BOTTOM);
	bNoLeadLeft = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_NOLEAD_LEFT);
	bNoLeadTop = MTblQueryEffCellImageFlags(printCellsParams.vecColsInfo.at(iEffIndex).handle, lEffRow, MTSI_NOLEAD_TOP);

	lCellTop = printCellParams.cellData.rect.top;
	lCellLeft = PrintCellsGetCellAreaLeft(printCellsParams, iEffIndex, lEffRow, lpCellToPrint);
	lCellRight = printCellParams.cellData.rect.right - 1;
	lCellBottom = printCellParams.cellData.rect.bottom - 1;

	lLeft = lCellLeft;

	if (!bNoLeadLeft || PrintCellsGetEffCellIndent(printCellsParams, iEffIndex, lEffRow) > 0)
		lLeft += printCellParams.cellData.iTextOffsetX;
	else
		lLeft++;

	lTop = printCellParams.cellData.rect.top;
	if (!bNoLeadTop)
		lTop += printCellParams.cellData.iTextOffsetY;
	else
		lTop++;

	lRight = printCellParams.cellData.rect.right;
	lBottom = printCellParams.cellData.rect.bottom;

	if (bAlignTile)
	{
		printCellParams.cellData.rImageRect.left = lLeft;
		printCellParams.cellData.rImageRect.top = lTop;
		printCellParams.cellData.rImageRect.right = lRight;
		printCellParams.cellData.rImageRect.bottom = lBottom;
	}
	else
	{
		if (bAlignLeft)
			printCellParams.cellData.rImageRect.left = lLeft;
		else if (bAlignHorzCenter)
			printCellParams.cellData.rImageRect.left = max(lLeft, lCellLeft + (((lCellRight - lCellLeft) - printCellParams.cellData.siImageSize.cx) / 2.0));
		else
			printCellParams.cellData.rImageRect.left = max(lLeft, lCellRight - printCellParams.cellData.siImageSize.cx);

		if (bAlignTop)
			printCellParams.cellData.rImageRect.top = lTop;
		else if (bAlignVertCenter)
			printCellParams.cellData.rImageRect.top = max(lTop, lCellTop + (((lCellBottom - lCellTop) - printCellParams.cellData.siImageSize.cy) / 2.0));
		else if (bAlignBottom)
			printCellParams.cellData.rImageRect.top = max(lTop, lCellBottom - printCellParams.cellData.siImageSize.cy);
		else
			printCellParams.cellData.rImageRect.top = max(lTop, printCellParams.cellData.rTextRect.top + ((iFontHeight - printCellParams.cellData.siImageSize.cy) / 2.0));

		printCellParams.cellData.rImageRect.right = printCellParams.cellData.rImageRect.left + printCellParams.cellData.siImageSize.cx;
		printCellParams.cellData.rImageRect.bottom = printCellParams.cellData.rImageRect.top + printCellParams.cellData.siImageSize.cy;
	}
}

/*
Überprüft, ob für die übergebene Zeile die übergebenen Flags, die in einer Zeile gesetzt sein müssen, gesetzt sind und die übergebenen Flags, die nicht gesetzt sein dürfen, nicht gesetzt sind.
Parameter:
tableInfo:	Tabelleninformationen
lRow:		Betrachtete Zeile
dwFlagsOn:	Flags, die gesetzt sein müssen, damit die Zeile gedruckt wird
dwFlagsOff:	Flags, die nicht gesetzt sein dürfen, damit die Zeile gedruckt wird
Return:
TRUE, wenn die benötigten Flags gesetzt und die unzulässigen Flags nicht gesetzt sind.
FALSE, wenn die Flags nicht korrekt sind.
*/
BOOL MTBLSUBCLASS::PrintCellsCheckRowFlags(TABLEINFO tableInfo, long lRow, WORD dwFlagsOn, WORD dwFlagsOff)
{
	BOOL bOnOk, bOffOk;

	if (dwFlagsOn != 0)
		bOnOk = Ctd.TblQueryRowFlags(tableInfo.handle, lRow, dwFlagsOn);
	else
		bOnOk = TRUE;

	if (dwFlagsOff != 0)
		bOffOk = !Ctd.TblQueryRowFlags(tableInfo.handle, lRow, dwFlagsOff);
	else
		bOffOk = TRUE;

	return bOnOk && bOffOk;
}

/*
Ermittelt die vertikalen Linien der Baumstruktur von allen Kindern zu den Eltern.
Parameter:
printCellsParams:		Parameter der Zellen
printCellParams:		Parameter der betrachteten Zelle, wird durch die vertikalen Linien ergänzt
iEffRowLevel:			Zeilenlevel, gibt an, in welcher Einrückungsebene sich die betrachtete Zelle befindet
iTreeIndentPerLevel:	Einrückung pro Ebene
iFontHeight:			Font Höhe
lpCellToPrint:			Zu druckende Zelle, bestimmt, von welcher Zeile aus die Baumlinien ermittelt werden sollen
*/
void MTBLSUBCLASS::PrintCellsGetVertTreeLinesToParent(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS &printCellParams, int iEffRowLevel, int iTreeIndentPerLevel, int iFontHeight, LPCELLTOPRINT lpCellToPrint)
{
	BOOL bDraw, bLastChild;
	int iCurrentRowLevel, iDrawLevel;
	long lCurrentRow, lParentRow;

	printCellParams.cellData.vecVTLTP_Pos.clear();

	lParentRow = TBL_Error;
	for (iCurrentRowLevel = iEffRowLevel; iCurrentRowLevel >= 0; iCurrentRowLevel--)
	{
		if (iCurrentRowLevel == iEffRowLevel)
		{
			if (printCellParams.cellData.bTreeBottomUp)
				lCurrentRow = lpCellToPrint->lFromRow;
			else
				lCurrentRow = lpCellToPrint->lToRow;
		}
		else
			lCurrentRow = lParentRow;

		//Parent Zeile der betrachteten Zeile ermitteln
		lParentRow = MTblGetParentRow(printCellsParams.tableInfo.handle, lCurrentRow);

		bDraw = TRUE;
		bLastChild = TRUE;

		//Letzte Kindzeile der betrachteten Zeile ermitteln
		lCurrentRow = MTblGetNextChildRow(printCellsParams.tableInfo.handle, lCurrentRow);
		while (lCurrentRow != TBL_Error)
		{
			if (PrintCellsCheckRowFlags(printCellsParams.tableInfo, lCurrentRow, printCellsParams.printOptions.dwRowFlagsOn, printCellsParams.printOptions.dwRowFlagsOff))
			{
				bLastChild = FALSE;
				break;
			}
			lCurrentRow = MTblGetNextChildRow(printCellsParams.tableInfo.handle, lCurrentRow);
		}

		if (bLastChild && (iCurrentRowLevel < iEffRowLevel))
			bDraw = FALSE;

		if (bDraw)
		{
			if ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_FLAT_STRUCT) != 0)
				iDrawLevel = 0;
			else
				iDrawLevel = iCurrentRowLevel;

			if (iDrawLevel > 0 && ((printCellsParams.tableInfo.dwTreeFlags & MTBL_TREE_FLAG_INDENT_NONE) != 0))
				iDrawLevel -= 1;

			RECT rect;
			rect.left = printCellParams.cellData.rect.left + (printCellParams.cellData.iTextOffsetX - 1) + (iDrawLevel * iTreeIndentPerLevel) + (iTreeIndentPerLevel / 2.0);
			rect.right = rect.left;

			if (bLastChild)
			{
				if (printCellParams.cellData.bTreeBottomUp)
				{
					rect.top = printCellParams.cellData.rTextRect.top + printCellParams.cellData.iTextOffsetY + (iFontHeight / 2.0);
					rect.bottom = printCellParams.cellData.rect.bottom + 1;
				}
				else
				{
					rect.top = printCellParams.cellData.rect.top;
					rect.bottom = printCellParams.cellData.rTextRect.top + printCellParams.cellData.iTextOffsetY + (iFontHeight / 2.0) + 1;
				}
			}
			else
			{
				rect.top = printCellParams.cellData.rect.top;
				rect.bottom = printCellParams.cellData.rect.bottom + 1;
			}

			printCellParams.cellData.vecVTLTP_Pos.push_back(rect);
		}
	}
}

/*
Ermittelt den Zeichenbereich eines Knotens.
Parameter:
lCellLeft:				Linker Zellenabstand
lTextTop:				Oberer Textabstand
iFontHeight:			Font Höhe
iNodeLevel:				Einrückungsebene des Knotens
iCellLeadingX:			Zellenoffset x
iCellLeadingY:			Zellenoffset y
iNodeWid:				Knotenbreite
iNodeHt:				Knotenhöhe
iTreeIndentPerLevel:	Einrückung pro Ebene
nodeRect:				Knotenrechteck, wird mit dem Zeichenbereich des Knotens gefüllt
Return:
TRUE.
*/
BOOL MTBLSUBCLASS::PrintGetNodeRect(long lCellLeft, long lTextTop, int iFontHeight, int iNodeLevel, int iCellLeadingX, int iCellLeadingY, int iNodeWid, int iNodeHt, int iTreeIndentPerLevel, RECT &nodeRect)
{
	nodeRect.left = lCellLeft + (iCellLeadingX - 1) + (iNodeLevel * iTreeIndentPerLevel) + ((iTreeIndentPerLevel - iNodeWid) / 2.0);
	nodeRect.top = lTextTop + iCellLeadingY + (max(1, (iFontHeight - iNodeHt) / 2.0));
	nodeRect.right = nodeRect.left + iNodeWid;
	nodeRect.bottom = nodeRect.top + iNodeHt;

	return TRUE;
}

/*
Überführt einen Gridtyp in ein RECT struct.
Parameter:
gridType:	Gridtyp
rect:		Rechteck, wird mit den entsprechenden Kanten des Gridtyps gefüllt
Die im Gridtyp definierten Ränder werden im RECT struct auf 1 gesetzt.
*/
void MTBLSUBCLASS::BorderFlagsFromGridType(GridType gridType, RECT &rect)
{
	switch (gridType)
	{
	case gtNo:
		rect.left = 0;
		rect.top = 0;
		rect.right = 0;
		rect.bottom = 0;
		break;
	case gtStandard:
		rect.left = 1;
		rect.top = 0;
		rect.right = 1;
		rect.bottom = 0;
		break;
	case gtStandard2:
		rect.left = 0;
		rect.top = 1;
		rect.right = 0;
		rect.bottom = 1;
		break;
	case gtTable:
		rect.left = 1;
		rect.top = 1;
		rect.right = 1;
		rect.bottom = 1;
		break;

	}
}

/*
Ermittelt die Anzahl Druckerpixel für eine übergebene Anzahl Pixel.
Parameter:
iPix:			Anzahl Pixel, die umgerechnet werden sollen
bVertical:		Berechnung für vertikale Pixel ja/nein
tableInfo:		Informationen über die Tabelle
printMetrics:	Druckmetriken für die Umrechnung
Return:
Anzahl Druckerpixel
*/
int MTBLSUBCLASS::PixTblToPrint(int iPix, BOOL bVertical, TABLEINFO tableInfo, PRINTMETRICS printMetrics)
{
	int iResult;
	if (bVertical)
		iResult = round(iPix / (double)(tableInfo.iPixelPerInchY) * printMetrics.iPixelPerInchY * printMetrics.dScale);
	else
		iResult = round(iPix / (double)(tableInfo.iPixelPerInchX) * printMetrics.iPixelPerInchX * printMetrics.dScale);

	return iResult;
}

/*
Ermittelt den aktuellen Druckauftragsstatus.
Parameter:
lpsPrinter:		Name des Druckers, auf dem der Druckauftrag durchgeführt wird
iPrintJob:		Identifier des Druckauftrags
Return:
Status des Druckauftrags (Ref: JOB_INFO Struktur)
*/
DWORD MTBLSUBCLASS::GetPrintJobStatus(tstring tsPrinter, int iPrintJob)
{
	DWORD dwResult = 0;
	HANDLE hPrinter;
	DWORD dwBufferSizeNeeded = 0;
	LPBYTE lpbPrintJobInfoStruct;

	//Drucker Handle ermitteln
	if (OpenPrinter((LPTSTR)tsPrinter.c_str(), &hPrinter, NULL))
	{
		GetJob(hPrinter, iPrintJob, 1, 0, 0, &dwBufferSizeNeeded);
		if (dwBufferSizeNeeded > 0)
		{
			lpbPrintJobInfoStruct = new BYTE[dwBufferSizeNeeded];
			if (GetJob(hPrinter, iPrintJob, 1, lpbPrintJobInfoStruct, dwBufferSizeNeeded, &dwBufferSizeNeeded))
			{
				JOB_INFO_1* jobInfo = (JOB_INFO_1*)(lpbPrintJobInfoStruct);
				dwResult = jobInfo->Status;
			}
			delete[] lpbPrintJobInfoStruct;
		}		
		ClosePrinter(hPrinter);
	}
	
	return dwResult;
}

/*
Ermittelt die Anzahl der zu druckenden Seiten. Zur Ermittlung der Seitenanzahl wird der Druckvorgang mithilfe
einer Kopie des Device Contexts des Druckers simuliert.
Parameter:
printOptions:	Verwendete Druckoptionen
printMetrics:	Druckermetriken
tableInfo:		Informationen über die zu druckende Tabelle
vecColInfo:		Informationen über die zu druckenden Spalten
hDeviceContext:	Device Kontext, der als Referenz für die Simulation dient
Return:
Anzahl zu druckender Seiten.
-1, wenn ein Fehler aufgetreten ist
*/
long MTBLSUBCLASS::GetPageCount(PRINTOPTIONS printOptions, PRINTMETRICS printMetrics, TABLEINFO tableInfo, VecColInfo vecColInfo, HDC hDeviceContext)
{
	PRINTPAGEPARAMS printPageParams;
	long lPage, lCurpage;
	BOOL bError;
	long lColNext, lColFrom, lColTo, lColLow, lColMax, lStartRow, lNextStartRow, lLastRow, lLastTotal;
	BOOL bFound;
	HDC hDummyDeviceContext;
	long lResult;

	bError = FALSE;
	//Dummy Device Context erzeugen
	hDummyDeviceContext = CreateCompatibleDC(hDeviceContext);

	//Fixe Parameter zur Ausgabe der Seite setzen
	printPageParams.hdcCanvas = hDummyDeviceContext;
	printPageParams.printOptions = printOptions;
	printPageParams.printMetrics = printMetrics;
	printPageParams.tableInfo = tableInfo;
	printPageParams.vecColsInfo = vecColInfo;

	lCurpage = mtblPrint_lCurrentPage;
	lPage = 1;
	lStartRow = -1;
	lNextStartRow = 0;
	lLastTotal = -1;
	lColLow = 0;
	lColMax = vecColInfo.size() - 1;
	lColFrom = lColLow;
	lColTo = lColLow;

	//Seite "drucken"
	while (TRUE)
	{
		//Bis-Spalte ermitteln
		lColNext = lColFrom + 1;
		bFound = FALSE;
		while (lColNext <= lColMax)
		{
			if (vecColInfo.at(lColNext).bNewPage)
			{
				lColTo = lColNext - 1;
				bFound = TRUE;
				break;
			}
			lColNext++;
		}

		if (!bFound)
			lColTo = lColMax;

		//Parameter zum Seitendruck setzen
		printPageParams.iPageNr = lPage;
		printPageParams.lStartRow = lStartRow;
		printPageParams.iFromCol = lColFrom;
		printPageParams.iToCol = lColTo;
		printPageParams.bColHeaders = (printOptions.bColHeaders && lNextStartRow != -3);
		printPageParams.bCells = (lNextStartRow != -3);

		lNextStartRow = PrintPage(printPageParams, lLastRow, lLastTotal);

		//Fehler?
		if (lNextStartRow == -2)
		{
			bError = TRUE;
			break;
		}

		//Letzte Seite?
		if (lNextStartRow == -1 && lColTo == lColMax)
			break;

		//Seite hochzählen
		lPage++;

		//Spalte von/bis und Startzeile setzen
		if (lColTo == lColMax)
		{
			lColFrom = lColLow;
			lColTo = lColLow;
			lStartRow = lNextStartRow;
		}
		else
			lColFrom = lColTo + 1;
	}

	//Device Context freigeben
	DeleteDC(hDummyDeviceContext);

	//Fertig
	mtblPrint_lCurrentPage = lCurpage;

	//Fehler?
	if (bError)
		lResult = -1;
	else
		lResult = lPage;

	return lResult;
}

/*
Ermittelt je nach Sprache mit einem übergebenen Schlüssel den dazu gehörigen Text.
Parameter:
pritnOptions:	Druckeroptionen, enthält die verwendete Sprache
lpsKey:			Schlüsselwort für die Abfrage des Textes
tsResult:		String, in den das Ergebnis geschrieben wird
*/
void MTBLSUBCLASS::GetUIStr(PRINTOPTIONS printOptions, LPTSTR lpsKey, tstring &tsResult)
{
	HSTRING hString = 0;
	LPTSTR lpsSection;

	//Ermitteln
	lpsSection = _T("Printing");
	//Text abfragen
	if (MTblGetLanguageString(printOptions.iLanguage, (LPCTSTR)lpsSection, (LPCTSTR)lpsKey, &hString) == TRUE)
	{
		Ctd.HStrToTStr(hString, tsResult, TRUE);
	}
}

/*
Ersetzt Platzhalter für Seitenzahlen oder Gesamtseitenzahl in einem übergebenen String.
Parameter:
tString:	String, der Platzhalter enthält, die ersetzt werden sollen
Return:
Übergebenen String mit ersetzten Platzhaltern.
*/
tstring MTBLSUBCLASS::ExpandPlaceholders(tstring tString)
{
	int i;

	for (i = 0; i < mtblPrint_PlaceholderCount; i++)
	{

		//Position des Placeholders ermitteln
		tstring::size_type pos = tString.find(mtblPrint_Placeholders[i]);

		//Placeholder vorhanden?
		if (pos != tstring::npos)
		{
			//Placeholder ersetzen
			if (mtblPrint_Placeholders[i].compare(_T("%PAGENR%")) == 0)
				tString.replace(pos, mtblPrint_Placeholders[i].length(), to_tstring(mtblPrint_lCurrentPage));
			else if (mtblPrint_Placeholders[i].compare(_T("%PAGECOUNT%")) == 0)
				tString.replace(pos, mtblPrint_Placeholders[i].length(), to_tstring(mtblPrint_PageCount));
			else
				tString.replace(pos, mtblPrint_Placeholders[i].length(), _T(""));
		}
	}

	return tString;
}

/*
Ermittelt die Höhe in Pixeln eines übergebenen Strings mit den für die Zeile definierten Parametern.
Parameter:
tString:			String, dessen Höhe bestimmt werden soll
printLineParams:	Parameter der betrachteten Zeile
dwDrawFlags:		Falgs zum Zeichnen des Textes
Return:
Benötigte Höhe des Strings in Pixeln.
*/
int MTBLSUBCLASS::PrintLine_GetTextHeight(tstring tString, PRINTLINEPARAMS printLineParams, DWORD dwDrawFlags)
{
	RECT rect;
	SetRectEmpty(&rect);
	DrawText(printLineParams.hdcCanvas, tString.c_str(), -1, &rect, dwDrawFlags | DT_CALCRECT | DT_EXTERNALLEADING);

	return rect.bottom - rect.top;
}

/*
Ermittelt die Größe eines Bildes.
Parameter:
hImage:				Handle für das betrachtete Bild
printLineParams:	Parameter der betrachteten Zeile
Return:
Größe des Bildes in Pixeln.
*/
SIZE MTBLSUBCLASS::PrintLine_GetImageSize(HANDLE hImage, PRINTLINEPARAMS printLineParams)
{
	double dScale;
	SIZE siResult;
	siResult.cx = 0;
	siResult.cy = 0;

	if (hImage != 0)
	{
		dScale = printLineParams.lpPrintMetrics->dScale;
		MImgGetSize(hImage, &siResult);
		if ((printLineParams.line.dwFlags & MTPL_NO_IMAGE_SCALE) != 0)
			dScale = 1.0;

		siResult.cx = round(siResult.cx / (double)(printLineParams.lpTableInfo->iPixelPerInchX) * printLineParams.lpPrintMetrics->iPixelPerInchX * dScale);
		siResult.cy = round(siResult.cy / (double)(printLineParams.lpTableInfo->iPixelPerInchY) * printLineParams.lpPrintMetrics->iPixelPerInchY * dScale);
	}

	return siResult;
}

/*
Ermittelt die Höhe eines Bildes.
Parameter:
hImage:				Handle für das betrachtete Bild
printLineParams:	Parameter der betrachteten Zeile
Return:
Höhe des Bilder in Pixeln.
*/
int MTBLSUBCLASS::PrintLine_GetImageHeight(HANDLE hImage, PRINTLINEPARAMS printLineParams)
{
	SIZE siResult;
	siResult = PrintLine_GetImageSize(hImage, printLineParams);

	return siResult.cy;
}

/*
Zeichnet ein Bild in ein bestimmtes Feld mit einer übergebenen Ausrichtung.
Parameter:
hImage:				Handle für das betrachtete Bild
iJustify:			Ausrichtung des Bildes
rect:				Feld, in das das Bild gezeichnet werden soll
printLineParams:	Parameter der betrachteten Zeile
*/
void MTBLSUBCLASS::PrintLine_DrawImage(HANDLE hImage, int iJustify, RECT rect, PRINTLINEPARAMS printLineParams)
{
	int iWidth;
	SIZE size;
	RECT rectImage;

	if (hImage != 0)
	{
		iWidth = rect.right - rect.left;
		size = PrintLine_GetImageSize(hImage, printLineParams);
		rectImage.top = rect.top;
		rectImage.bottom = rectImage.top + size.cy;

		//Ausrichtung
		if (iJustify == DT_LEFT)
			rectImage.left = rect.left;
		else if (iJustify == DT_CENTER)
			rectImage.left = rect.left + ((iWidth - size.cx) / 2);
		else if (iJustify == DT_RIGHT)
			rectImage.left = rect.right - size.cx;

		rectImage.right = rectImage.left + size.cx + 1;
		//Bild zeichnen
		MImgPrint(hImage, printLineParams.hdcCanvas, &rectImage, RGB(255, 255, 255), 0);
	}
}

/*
Prüft alle Titel, Totals, Seitenkopfzeilen und Seitenfußzeilen, ob die Gesamtseitenzahl in einem Text gefordert ist.
Parameter:
printOptions:	Druckeroptionen, enthält die Titel, Totals, Seitenkopfzeilen und Seitenfußzeilen
Return:
TRUE, wenn die Gesamtseitenzahl gefordert ist
FALSE, wenn die Gesamtseitenzahl nicht gefordert ist.
*/
BOOL MTBLSUBCLASS::IsPageCountNeeded(PRINTOPTIONS printOptions)
{
	BOOL bResult = TRUE;
	//Nach Seitenzahl-Platzhaltern suchen
	if (IsPageCountPhIncl(printOptions.vecTitles, printOptions.iTitleCount))
		return bResult;
	if (IsPageCountPhIncl(printOptions.vecTotals, printOptions.iTotalCount))
		return bResult;
	if (IsPageCountPhIncl(printOptions.vecPageHeaders, printOptions.iPageHeaderCount))
		return bResult;
	if (IsPageCountPhIncl(printOptions.vecPageFooters, printOptions.iPageFooterCount))
		return bResult;

	bResult = FALSE;
	return bResult;
}

/*
Prüft einen Vektor, ob in einem der Texte in dem übergebenen Vektor der Platzhalter %PAGECOUNT% vorkommt.
Dies bedeutet, dass die Gesamtanzahl an Seiten in einem der Texte aufegführt werden soll.
Parameter:
printLines:	Texte, die überprüft werden sollen
lItemCount:	Anzahl der Texte
Return:
TRUE, wenn der Platzhalter vorkommt
FALSE, wenn der Platzhalter nicht vorkommt
*/
BOOL MTBLSUBCLASS::IsPageCountPhIncl(std::vector<PRINTLINE> printLines, long lItemCount)
{
	long l;
	BOOL bResult = FALSE;

	for (l = 0; l < lItemCount; l++)
	{
		//PAGECOUNT gefunden?
		if (printLines.at(l).tsLeftText.find(_T("%PAGECOUNT%")) != tstring::npos ||
			printLines.at(l).tsCenterText.find(_T("%PAGECOUNT%")) != tstring::npos ||
			printLines.at(l).tsRightText.find(_T("%PAGECOUNT%")) != tstring::npos)
		{
			bResult = TRUE;
			break;
		}
	}

	return bResult;
}

/*
Ermittelt die Druckmetriken für den gegebenen Device Context.
Parameter:
hDeviceContext:	Device Kontext, für den die Druckmetriken bestimmt werden sollen
printOptions:	Verwendete Druckoptionen
tableInfo:		Informationen über die zu druckende Tabelle
vecColInfo:		Spalteninformationen, wird modifiziert (Spaltenbreite, neue Seite ja/nein)
lpPrintMetrics:	Druckmetriken, wird gefüllt
Return:
TRUE.
*/
BOOL MTBLSUBCLASS::GetPrintMetrics(HDC hDeviceContext, PRINTOPTIONS printOptions, TABLEINFO tableInfo, VecColInfo &vecColInfo, LPPRINTMETRICS lpPrintMetrics)
{
	int i;
	//CMTblFont tblFont, prnFont;
	double dWidth, dColWidth;

	//Pixel per Inch
	lpPrintMetrics->iPixelPerInchX = GetDeviceCaps(hDeviceContext, LOGPIXELSX);
	lpPrintMetrics->iPixelPerInchY = GetDeviceCaps(hDeviceContext, LOGPIXELSY);

	//Pixel per mm
	lpPrintMetrics->dPixelPerMMX = lpPrintMetrics->iPixelPerInchX / 25.4;
	lpPrintMetrics->dPixelPerMMY = lpPrintMetrics->iPixelPerInchY / 25.4;

	//Physikalische Höhe und Breite ermitteln in Pixel
	lpPrintMetrics->iPhysWidth = GetDeviceCaps(hDeviceContext, PHYSICALWIDTH);
	lpPrintMetrics->iPhysHeight = GetDeviceCaps(hDeviceContext, PHYSICALHEIGHT);

	//Offsets des bedruckbaren Bereichs ermitteln in Pixel
	lpPrintMetrics->iLeftOffset = GetDeviceCaps(hDeviceContext, PHYSICALOFFSETX);
	lpPrintMetrics->iTopOffset = GetDeviceCaps(hDeviceContext, PHYSICALOFFSETY);
	lpPrintMetrics->iRightOffset = lpPrintMetrics->iPhysWidth - GetDeviceCaps(hDeviceContext, HORZRES) - lpPrintMetrics->iLeftOffset;
	lpPrintMetrics->iBottomOffset = lpPrintMetrics->iPhysHeight - GetDeviceCaps(hDeviceContext, VERTRES) - lpPrintMetrics->iTopOffset;

	//Ränder bezogen auf den bedruckbaren Bereich festlegen in Pixel
	lpPrintMetrics->iLeftMargin = round(printOptions.iLeftMargin * lpPrintMetrics->dPixelPerMMX) - lpPrintMetrics->iLeftOffset;
	lpPrintMetrics->iTopMargin = round(printOptions.iTopMargin * lpPrintMetrics->dPixelPerMMY) - lpPrintMetrics->iTopOffset;
	lpPrintMetrics->iRightMargin = round(printOptions.iRightMargin * lpPrintMetrics->dPixelPerMMX) - lpPrintMetrics->iRightOffset;
	lpPrintMetrics->iBottomMargin = round(printOptions.iBottomMargin * lpPrintMetrics->dPixelPerMMY) - lpPrintMetrics->iBottomOffset;

	//Keine negativen Ränder zulassen
	if (lpPrintMetrics->iLeftMargin < 0)
		lpPrintMetrics->iLeftMargin = 0;
	if (lpPrintMetrics->iRightMargin < 0)
		lpPrintMetrics->iRightMargin = 0;
	if (lpPrintMetrics->iTopMargin < 0)
		lpPrintMetrics->iTopMargin = 0;
	if (lpPrintMetrics->iBottomMargin < 0)
		lpPrintMetrics->iBottomMargin = 0;

	//Minimal- und Maximalwerte setzen
	lpPrintMetrics->clipRect.left = lpPrintMetrics->iLeftMargin;
	lpPrintMetrics->clipRect.right = GetDeviceCaps(hDeviceContext, HORZRES) - lpPrintMetrics->iRightMargin - 1;
	lpPrintMetrics->clipRect.top = lpPrintMetrics->iTopMargin;
	lpPrintMetrics->clipRect.bottom = GetDeviceCaps(hDeviceContext, VERTRES) - lpPrintMetrics->iBottomMargin - 1;

	if (printOptions.bSpreadCols)
	{
		//Überprüfe, ob eine Spalte breiter als eine Seitenbreite ist, wenn ja, setze die Breite auf die Seitenbreite
		for (int l = 0; l < vecColInfo.size(); l++)
		{
			if (round(vecColInfo.at(l).iWidthPix / (double)(tableInfo.iPixelPerInchX) * lpPrintMetrics->iPixelPerInchX) >(lpPrintMetrics->clipRect.right - lpPrintMetrics->clipRect.left))
			{
				vecColInfo.at(l).iWidthPix = round((lpPrintMetrics->clipRect.right - lpPrintMetrics->clipRect.left) / (double)(lpPrintMetrics->iPixelPerInchX) * tableInfo.iPixelPerInchX);
			}
		}
	}

	//Gesamtbreite der zu druckenden Spalten ermitteln
	lpPrintMetrics->iTotalColWidth = 0;
	for (i = 0; i < vecColInfo.size(); i++)
	{
		lpPrintMetrics->iTotalColWidth += round(vecColInfo.at(i).iWidthPix / (double)(tableInfo.iPixelPerInchX) * lpPrintMetrics->iPixelPerInchX);
	}

	//Skalierungsfaktor ermitteln
	if (printOptions.dScale == 0)
	{
		lpPrintMetrics->dScale = 1.0;
		if (printOptions.bSpreadCols == FALSE)
		{
			if (lpPrintMetrics->iTotalColWidth > (lpPrintMetrics->clipRect.right - lpPrintMetrics->clipRect.left))
				lpPrintMetrics->dScale = (lpPrintMetrics->clipRect.right - lpPrintMetrics->clipRect.left) / (double)(lpPrintMetrics->iTotalColWidth);
		}
	}
	else
		lpPrintMetrics->dScale = printOptions.dScale;

	//Spalten ermitteln, ab denen eine neue Seite beginnt
	dWidth = 0.0;
	for (i = 0; i < vecColInfo.size(); i++)
	{
		if (printOptions.bSpreadCols)
		{
			dColWidth = round(vecColInfo.at(i).iWidthPix / (double)(tableInfo.iPixelPerInchX) * lpPrintMetrics->iPixelPerInchX) * lpPrintMetrics->dScale;
			dWidth += dColWidth;
			if (dWidth >(lpPrintMetrics->clipRect.right - lpPrintMetrics->clipRect.left))
			{
				dWidth = dColWidth;
				vecColInfo.at(i).bNewPage = TRUE;
			}
			else
				vecColInfo.at(i).bNewPage = FALSE;
		}
		else
			vecColInfo.at(i).bNewPage = FALSE;
	}

	//Wenn der Drucker den Font nicht unterstützt, Standard-Font verwenden, ansonsten Font der Tabelle
	BOOL bExists = PrintFontExists(tableInfo.tsFontName);

	if (bExists)
		lpPrintMetrics->tsFontName = tableInfo.tsFontName;
	else
		lpPrintMetrics->tsFontName = _T("Arial");

	//Höhe des Fonts in Druckerpixeln ermitteln
	lpPrintMetrics->dFontHeightNS = tableInfo.iFontSize;
	lpPrintMetrics->dFontHeight = tableInfo.iFontSize * lpPrintMetrics->dScale;

	lpPrintMetrics->lFontStyle = tableInfo.dwFontEnh;

	return TRUE;
}

/*
Ermittelt Informationen über die zu druckenden Spalten der Tabelle.
Parameter:
tableInfo:		Informationen über die zu druckende Tabelle
vecColInfo:		Vektor von Spalteninformationen der zu druckenden Spalten, diese Struktur wird gefüllt
printOptions:	Verwendete Druckoptionen
Return:
Anzahl zu druckender Spalten
-1, wenn ein Fehler aufgetreten ist
*/
int MTBLSUBCLASS::GetColsToPrint(TABLEINFO tableInfo, VecColInfo &vecColInfo, PRINTOPTIONS printOptions)
{
	int iC, iPos, iWidth, iMaxWidth, iMinWidth;
	int iJfy;
	long lLen, lRow;
	NUMBER number;
	HSTRING hsTitle;
	LPTSTR lpsTitle;
	HWND hWndCol = 0;
	HFONT hFont;
	HDC deviceContext;
	std::vector<long> vecLCellTextWidth;
	int iResult;

	BOOL bMustDelete;

	iPos = 0;

	//Informationen über Spalten ermitteln
	while (TRUE)
	{
		//Nächste Spalte ermitteln
		iPos = FindNextCol(&hWndCol, printOptions.dwColFlagsOn, printOptions.dwColFlagsOff);
		if (iPos == 0)
			break;

		COLINFO colInfo;

		//Handle
		colInfo.handle = hWndCol;

		//ID
		colInfo.iID = Ctd.TblQueryColumnID(hWndCol);

		//Zelltyp
		colInfo.iCellType = (this->m_Cols->Find(hWndCol))->m_pCellType->m_iType;

		//Invisible ermitteln
		//colInfo.bInvisible = (this->m_Cols->Find(hWndCol))->QueryFlags(COL_Invisible);
		colInfo.bInvisible = Ctd.TblQueryColumnFlags(hWndCol, COL_Invisible);

		//Word Wrap ermitteln
		//colInfo.bWordWrap = ((this->m_Cols->Find(hWndCol))->QueryFlags(COL_MultilineCell) && !colInfo.bInvisible);
		colInfo.bWordWrap = (Ctd.TblQueryColumnFlags(hWndCol, COL_MultilineCell) && !colInfo.bInvisible);

		//Breite in Pixeln ermitteln
		if (printOptions.bFitCellSize && colInfo.bWordWrap == FALSE)
			colInfo.iWidthPix = 0;
		else
		{
			Ctd.TblQueryColumnWidth(hWndCol, &number);
			colInfo.iWidthPix = Ctd.FormUnitsToPixels(hWndCol, number, FALSE);
		}

		//Ausrichtung ermitteln
		if ((this->m_Cols->Find(hWndCol))->QueryFlags(COL_RightJustify))
			iJfy = DT_RIGHT;
		else if ((this->m_Cols->Find(hWndCol))->QueryFlags(COL_CenterJustify))
			iJfy = DT_CENTER;
		else
			iJfy = DT_LEFT;

		colInfo.iJustify = iJfy;

		//Titel ermitteln
		hsTitle = 0;
		Ctd.TblGetColumnTitle(hWndCol, &hsTitle, 2048);
		Ctd.HStrToTStr(hsTitle, colInfo.tsTitle, TRUE);		

		//Tree Spalte?
		colInfo.bTreeCol = (colInfo.handle == tableInfo.hwTreeCol);

		//Gelockt?
		colInfo.bLocked = this->IsLockedColumn(hWndCol);

		//Header-Gruppe
		colInfo.HdrGrp = this->m_ColHdrGrps->GetGrpOfCol(hWndCol);

		//Keine Zeilenlinien?
		colInfo.bNoRowLines = this->QueryColumnFlags(hWndCol, MTBL_COL_FLAG_NO_ROWLINES);

		//Keine Spaltenlinien?
		colInfo.bNoColLine = this->QueryColumnFlags(hWndCol, MTBL_COL_FLAG_NO_COLLINE);

		vecColInfo.push_back(colInfo);
	}

	if (printOptions.bFitCellSize)
	{
		iResult = -1;
		//Create Font

		deviceContext = GetDC(tableInfo.handle);
		if (deviceContext == 0)
			return iResult;

		for (int i = 0; i < vecColInfo.size(); i++){
			vecLCellTextWidth.push_back(0);
		}

		lRow = -1;
		while (Ctd.TblFindNextRow(tableInfo.handle, &lRow, printOptions.dwRowFlagsOn, printOptions.dwRowFlagsOff))
		{
			Ctd.TblFetchRow(tableInfo.handle, lRow);
			if (!Ctd.TblSetContext(tableInfo.handle, lRow))
				return iResult;

			for (iC = 0; iC < vecColInfo.size(); iC++)
			{
				if (vecColInfo.at(iC).iWidthPix == 0)
				{
					//Font erstellen
					if (lRow == TBL_Error)
						hFont = GetEffColHdrFont(deviceContext, vecColInfo.at(iC).handle, &bMustDelete, 0);

					else
						hFont = GetEffCellFont(deviceContext, lRow, vecColInfo.at(iC).handle, &bMustDelete, 0);

					if (hFont == 0)
						return iResult;

					//Beste Breite ermitteln
					iWidth = MTblGetBestCellWidth(tableInfo.handle, vecColInfo.at(iC).handle, lRow, tableInfo.metrics.ptCellLeading.x, hFont, tableInfo.pMergeCells, TRUE);
					//Breite setzen
					vecLCellTextWidth.at(iC) = max(vecLCellTextWidth.at(iC), iWidth);

					if (bMustDelete)
						DeleteObject(hFont);
				}
			}
		}

		for (iC = 0; iC < vecColInfo.size(); iC++)
		{
			if (vecColInfo.at(iC).iWidthPix == 0)
				vecColInfo.at(iC).iWidthPix = vecLCellTextWidth.at(iC);
		}

		for (iC = 0; iC < vecColInfo.size(); iC++)
		{
			if (vecColInfo.at(iC).bWordWrap == FALSE)
			{
				//Column-Header
				if (printOptions.bColHeaders)
				{
					//N_SetPrintFont
					hFont = GetEffColHdrFont(deviceContext, vecColInfo.at(iC).handle, &bMustDelete, 0);
					if (hFont == 0)
						return iResult;

					//Beste Breite ermitteln
					iWidth = MTblGetBestColHdrWidth(tableInfo.handle, vecColInfo.at(iC).handle, tableInfo.metrics.ptCellLeading.x, hFont);
					//Breite setzen
					vecColInfo.at(iC).iWidthPix = max(vecColInfo.at(iC).iWidthPix, iWidth);

					if (bMustDelete)
						DeleteObject(hFont);
				}

				//Min- und Max-Breite
				iMinWidth = SendMessage(tableInfo.handle, MTM_QueryMinColWidth, (WPARAM)(vecColInfo.at(iC).handle), vecColInfo.at(iC).iWidthPix);
				iMaxWidth = SendMessage(tableInfo.handle, MTM_QueryMaxColWidth, (WPARAM)(vecColInfo.at(iC).handle), vecColInfo.at(iC).iWidthPix);

				if (iMinWidth > 0 && iMinWidth > vecColInfo.at(iC).iWidthPix)
					vecColInfo.at(iC).iWidthPix = iMinWidth;
				if (iMaxWidth > 0 && iMaxWidth < vecColInfo.at(iC).iWidthPix)
					vecColInfo.at(iC).iWidthPix = iMaxWidth;
			}
		}
		ReleaseDC(tableInfo.handle, deviceContext);
	}
	iResult = vecColInfo.size();

	return iResult;
}

/*
Ermittelt alle nötige Tabelleninformationen aus der Tabelle.
Parameter:
hWndTbl:		Handle für die Tabelle, über die Informationen ermittelt werden sollen
printOptions:	Verwendete Druckoptionen
tableInfo:		Struktur für Tabelleninformationen, diese Struktur wird gefüllt
Return:
TRUE, wenn erfolgreich, 
FALSE, wenn das gegebene Handle kein Fenster oder nicht vom Tabellentyp ist.
*/
BOOL MTBLSUBCLASS::GetTableInfo(HWND hWndTbl, PRINTOPTIONS printOptions, TABLEINFO &tableInfo)
{
	long lWndType;
	HDC deviceContext;
	int iStyle;
	DWORD dwColor;
	BOOL bResult = FALSE;

	//Ist kein Window?
	if (!IsWindow(hWndTbl))
		return bResult;

	//Ist kein Tabellentyp?
	lWndType = Ctd.GetType(hWndTbl);
	if (!(lWndType == TYPE_ChildTable || lWndType == TYPE_TableWindow))
		return bResult;

	//Handle setzen
	tableInfo.handle = hWndTbl;

	//Pixel pro Inch
	deviceContext = GetDC(hWndTbl);
	tableInfo.iPixelPerInchX = GetDeviceCaps(deviceContext, LOGPIXELSX);
	tableInfo.iPixelPerInchY = GetDeviceCaps(deviceContext, LOGPIXELSY);

	//Fontinformationen
	CMTblFont font;
	font.Set(deviceContext, this->m_hFont);
	tableInfo.iFontSize = font.GetSize();
	tableInfo.dwFontEnh = font.GetEnh();
	tableInfo.tsFontName = font.GetName();
	ReleaseDC(hWndTbl, deviceContext);


	//Textfarbe
	tableInfo.crTextColor = Ctd.ColorGet(hWndTbl, COLOR_IndexWindowText);
	//Hintergrundfarbe
	tableInfo.crBackgroundColor = Ctd.ColorGet(hWndTbl, COLOR_IndexWindow);

	//Linien pro Zeile
	Ctd.TblQueryLinesPerRow(hWndTbl, &(tableInfo.iLinesPerRow));

	//Graue Spaltenköpfe?
	tableInfo.bGrayedHeaders = Ctd.TblQueryTableFlags(hWndTbl, TBL_Flag_GrayedHeaders);

	//Treedefenition
	tableInfo.hwTreeCol = 0;
	tableInfo.iNodeSize = 0;
	tableInfo.iTreeIndentPerLevel = 0;

	MTblQueryTree(hWndTbl, &(tableInfo.hwTreeCol), &(tableInfo.iNodeSize), &(tableInfo.iTreeIndentPerLevel));

	//Tree-Flags
	//From Subclass
	tableInfo.dwTreeFlags = this->m_dwTreeFlags;

	//Tree-Node-Farben
	//From Subclass
	tableInfo.dwNodeOuterColor = this->m_dwNodeColor;
	tableInfo.dwNodeInnerColor = this->m_dwNodeInnerColor;
	tableInfo.dwNodeBackColor = this->m_dwNodeBackColor;

	//Einrückung pro Einrüclungsebene
	//From Subclass
	tableInfo.iIndentPerLevel = this->m_lIndentPerLevel;

	//Passwortzeichen
	//From Subclass
	tableInfo.tcPasswordChar = this->m_cPassword;
	if (tableInfo.tcPasswordChar == TCHAR(0))
		tableInfo.tcPasswordChar = ' ';

	//Tabellenmaße
	SendMessage(hWndTbl, TIM_GetMetrics, 0, (LPARAM)(&(tableInfo.metrics)));

	//Keine Zeilenlinien?
	tableInfo.bNoRowLines = FALSE;
	if ((MTblQueryRowLines(hWndTbl, &iStyle, &dwColor) && (iStyle == MTLS_INVISIBLE)) || Ctd.TblQueryTableFlags(hWndTbl, TBL_Flag_SuppressRowLines))
		tableInfo.bNoRowLines = TRUE;

	//Keine Spaltenlinien?
	tableInfo.bNoColLines = FALSE;
	if (MTblQueryColLines(hWndTbl, &iStyle, &dwColor) && (iStyle == MTLS_INVISIBLE))
		tableInfo.bNoColLines = TRUE;

	//Zeiger auf Merge-Zellen
	tableInfo.pMergeCells = (LPMTBLMERGECELLS)MTblGetMergeCells(hWndTbl, printOptions.dwRowFlagsOn, printOptions.dwRowFlagsOff, printOptions.dwColFlagsOn, printOptions.dwColFlagsOff);

	//Fertig
	bResult = TRUE;

	return bResult;
}

//****************** Print Part End ******************


//******************* Preview Part *******************

/*
CALLBACK Funktion um alle Fenster eines Threads, die aktiviert sind, zu deaktivieren und zu einem Vektor hinzuzufügen
Parameter:
hWnd:	Fenster, das deaktiviert werden soll
lpdata:	Vektor mit allen Fenstern, die deaktiviert wurden
Return:
TRUE
*/
BOOL CALLBACK DisableThreadWndsProc(HWND hWnd, std::vector<HWND> &lpData)
{
	//Ist Fenster aktiviert?
	if (IsWindowEnabled(hWnd))
	{
		EnableWindow(hWnd, FALSE);
		lpData.push_back(hWnd);
	}
	return TRUE;
}

/*
Fensterprozedur für die benutzerdefinierte Skalierungseingabe der Druckvorschau.
Das Fenster wird geöffnet, wenn der Nutzer den entsprechenden Button im Druckvorschaufenter klickt.
Parameter:
hWndUserDefinedScaleDlg:	Handle für das Fenster
uMsg:		Empfangene Nachricht
wParam:		Daten aus der Nachricht
lParam:		Daten aus der Nachricht
*/
INT_PTR CALLBACK PreviewUserDefinedScaleProc(HWND hWndUserDefinedScaleDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	//Preview Objekt laden
	LPPRINTPREVIEW lpPrintPreview;
	lpPrintPreview = (LPPRINTPREVIEW)GetProp(hWndUserDefinedScaleDlg, _T("PROP_PREV_OBJECT"));

	switch (uMsg)
	{
	case WM_INITDIALOG:
		lpPrintPreview = (LPPRINTPREVIEW)lParam;
		lpPrintPreview->OnDlgUserDefinedScaleInit(hWndUserDefinedScaleDlg);
		return TRUE;

	case WM_CLOSE:
		EndDialog(hWndUserDefinedScaleDlg, 0);
		return TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDCANCEL:
			EndDialog(hWndUserDefinedScaleDlg, 0);
			return TRUE;

		case IDOK:
			lpPrintPreview->OnDlgUserDefinedScaleOkClick(hWndUserDefinedScaleDlg);
			return TRUE;
		}
		break;
	}

	return FALSE;
}

/*
Fensterprozedur des inneren Dialogs der Druckvorschau.
In dem Fenster befindet sich ein Panel, welcher wiederum eine Metafile mit der anzuzeigenen Seite befindet.
Der innere Dialog wird beim Erstellen der Druckvorschau erzeugt.
Parameter:
hWndDialogContent: Handle für das Fenster
uMsg:	Empfangene Nachricht
wParam:		Daten aus der Nachricht
lParam:		Daten aus der Nachricht
*/
INT_PTR CALLBACK DialogContentProc(HWND hWndDialogContent, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	//Preview Objekt laden
	LPPRINTPREVIEW lpPrintPreview;
	lpPrintPreview = (LPPRINTPREVIEW)GetProp(hWndDialogContent, _T("PROP_PREV_OBJECT"));

	switch (uMsg)
	{
	case WM_INITDIALOG:
		lpPrintPreview = (LPPRINTPREVIEW)lParam;
		SetProp(hWndDialogContent, _T("PROP_PREV_OBJECT"), (HANDLE)lpPrintPreview);
		return TRUE;

	case WM_SIZE:
		lpPrintPreview->OnContentSizeChanged(LOWORD(lParam), HIWORD(lParam));
		return TRUE;

	case WM_MOUSEWHEEL:
		lpPrintPreview->OnMouseWheelScroll(GET_WHEEL_DELTA_WPARAM(wParam));
		return TRUE;

	case WM_HSCROLL:
		lpPrintPreview->OnHorizontalScroll(LOWORD(wParam), HIWORD(wParam));
		return TRUE;

	case WM_VSCROLL:
		lpPrintPreview->OnVerticalScroll(LOWORD(wParam), HIWORD(wParam));
		return TRUE;
	}

	return FALSE;
}

/*
Fensterprozedur des Druckvorschaufensters.
Die Druckvorschau dient zur Ansicht der zu druckenden Seiten.
Parameter:
hWndDialog:	Handle für das Fenter
uMsg:		Empfangene Nachricht
wParam:		Daten aus der Nachricht
lParam:		Daten aus der Nachricht
*/
INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LPNMHDR lpnm;
	LPNMTOOLBAR lpnmtb;

	//Preview Objekt laden
	LPPRINTPREVIEW lpPrintPreview;
	lpPrintPreview = (LPPRINTPREVIEW)GetProp(hWndDialog, _T("PROP_PREV_OBJECT"));

	switch (uMsg)
	{
	case WM_INITDIALOG:
		lpPrintPreview = (LPPRINTPREVIEW)lParam;
		lpPrintPreview->OnDialogInit(hWndDialog);
		return TRUE;

	case WM_CLOSE:
		lpPrintPreview->CloseBtnClick();
		return TRUE;

	case WM_SIZE:
		SetWindowPos(lpPrintPreview->hWndDialogContent, 0, 0, lpPrintPreview->iMenuPanelHeight, LOWORD(lParam), HIWORD(lParam) - lpPrintPreview->iMenuPanelHeight - lpPrintPreview->iStatusbarHeight, SWP_NOZORDER);
		SendMessage(lpPrintPreview->hWndToolbar, TB_AUTOSIZE, 0, 0);

		SetWindowPos(lpPrintPreview->hWndStatusbar, NULL, 0, HIWORD(lParam) - 20, LOWORD(lParam), 20, SWP_NOZORDER);
		SendMessage(lpPrintPreview->hWndStatusbar, WM_SIZE, 0, 0);
		return TRUE;

	case WM_DESTROY:
		if (lpPrintPreview->hImageList)
		{
			ImageList_Destroy(lpPrintPreview->hImageList);
			lpPrintPreview->hImageList = NULL;
		}

		if (lpPrintPreview->hMenuPrint)
		{
			DestroyMenu(lpPrintPreview->hMenuPrint);
			lpPrintPreview->hMenuPrint = NULL;
		}

		if (lpPrintPreview->hMenuScale)
		{
			DestroyMenu(lpPrintPreview->hMenuScale);
			lpPrintPreview->hMenuScale = NULL;
		}
		return TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDCANCEL:
			SendMessage(hWndDialog, WM_CLOSE, 0, 0);
			return TRUE;

		case IDOK:
			lpPrintPreview->PrintBtnClick(LOWORD(wParam));
			return TRUE;

		case IDC_MENU_BTN_CLOSE:
			lpPrintPreview->CloseBtnClick();
			return TRUE;

		case IDC_MENU_BTN_PRINT:
		case IDC_MENU_ITEM_PRINT_ALL:
		case IDC_MENU_ITEM_PRINT_CURRENT_PAGE:
			lpPrintPreview->PrintBtnClick(LOWORD(wParam));
			return TRUE;

		case IDC_MENU_BTN_FULLPAGE:
			lpPrintPreview->FullPageBtnClick();
			return TRUE;

		case IDC_MENU_BTN_FIRSTPAGE:
			lpPrintPreview->FirstPageBtnClick();

		case IDC_MENU_BTN_PREVPAGE:
			lpPrintPreview->PrevPageBtnClick();
			return TRUE;

		case IDC_MENU_BTN_NEXTPAGE:
			lpPrintPreview->NextPageBtnClick();
			return TRUE;

		case IDC_MENU_BTN_LASTPAGE:
			lpPrintPreview->LastPageBtnClick();
			return TRUE;

		case IDC_MENU_BTN_SCALING:
			lpPrintPreview->UserDefinedScaleClick();
			return TRUE;

		case IDC_MENU_ITEM_SCALE_USER_DEFINED:
		case IDC_MENU_ITEM_SCALE_10:
		case IDC_MENU_ITEM_SCALE_25:
		case IDC_MENU_ITEM_SCALE_50:
		case IDC_MENU_ITEM_SCALE_75:
		case IDC_MENU_ITEM_SCALE_100:
		case IDC_MENU_ITEM_SCALE_150:
		case IDC_MENU_ITEM_SCALE_200:
		case IDC_MENU_ITEM_SCALE_TO_PAGE_WIDTH:
			lpPrintPreview->ScaleMenuItemClick(LOWORD(wParam));
			return TRUE;
		}
		break;

	case WM_NOTIFY:
		lpnm = (LPNMHDR)lParam;
		if (lpnm->code == TBN_DROPDOWN)
		{
			lpnmtb = (LPNMTOOLBAR)lParam;
			lpPrintPreview->DropDownClick(lpnmtb);
			return TBDDRET_DEFAULT;
		}
		break;

	case WM_SETCURSOR:
		if (lpPrintPreview->bWaitCursor)
		{
			SetCursor(LoadCursor(NULL, IDC_WAIT));
			SetWindowLongPtr(hWndDialog, DWLP_MSGRESULT, TRUE);
			return TRUE;
		}
		break;
	}

	return FALSE;
}

/*
PRINTPREVIEW Konstruktor
Erstellt ein neues PRINTPREVIEW-Objekt und öffnet den Vorschaudialog.
Parameter:
lpRefSubClass:	Referenz auf das MTBLSUBCLASS-Objekt, das den Vorschaudialog aufruft
printOptions:	Druckoptionen, die bei dem Vorschaudialog und der anschließenden Druckausgabe verwendet werden
printMetrics:	Druckmetriken für den gewählten Drucker
tableInfo:		Tabelleninformationen der anzuzeigenden Tabelle
vecColsInfo:	Informationen über anzuzeigende Spalten
hdcPrinter:		Device Context des Druckers
lpDevMode:		Device Mode des Druckers
Return erst, wenn der Vorschaudialog beendet wurde.
*/
PRINTPREVIEW::PRINTPREVIEW(LPMTBLSUBCLASS lpRefSubClass, PRINTOPTIONS printOptions, PRINTMETRICS printMetrics, TABLEINFO tableInfo,
	VecColInfo vecColsInfo, HDC hdcPrinter, LPDEVMODE lpDevMode)
{
	//Initialisierung
	this->lpSubclassRef = lpRefSubClass;
	this->printOptions = printOptions;
	this->printMetrics = printMetrics;
	this->tableInfo = tableInfo;
	this->vecColsInfo = vecColsInfo;
	this->hdcPrinter = hdcPrinter;
	this->pDevMode = lpDevMode;
	this->bFullPage = FALSE;

	//Deaktiviere alle Fenster des Threads um zu verhindern, dass die Tabelle geändert wird, während der Vorschaudialog aktiv ist
	DWORD dwThreadID = GetCurrentThreadId();
	//Aktuell aktives Fenster merken, um es später wieder zu aktivieren
	hWndLastActiveWindow = GetActiveWindow();
	EnumThreadWindows(dwThreadID, (WNDENUMPROC)DisableThreadWndsProc, LPARAM(&this->vecHwndDisabled));
	//Vorschaudialog erstellen
	this->CreatePreviewWindow();
}

/*
Initialisiert den Vorschaudialog.
Es werden alle nötigen Kindfenster erstellt, Starteigenschaften gesetzt und Größen angepasst und die erste Seite angezeigt.
Parameter:
hWndDialog:		Handle für das Vorschaufenster
*/
void PRINTPREVIEW::OnDialogInit(HWND hWndDialog)
{
	this->hWndDialog = hWndDialog;
	//Referenz auf das PrintPreview-Objekt als Eigenschaft des Fensters setzen, um von der Fensterprozedur auf das Objekt zugreifen zu können
	SetProp(this->hWndDialog, _T("PROP_PREV_OBJECT"), (HANDLE)this);

	// Toolbar erzeugen
	this->CreateToolbar(hWndDialog);

	// Toolbar anzeigen
	SendMessage(this->hWndToolbar, TB_AUTOSIZE, 0, 0);
	ShowWindow(this->hWndToolbar, TRUE);

	// Popup-Menüs erzeugen
	this->CreatePopupMenus();

	// Dialog Inhalt erzeugen
	this->hWndDialogContent = CreateDialogParam((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDD_DIALOG2), this->hWndDialog, DialogContentProc, LPARAM(this));
	SetWindowLongPtr(this->hWndDialogContent, GWL_STYLE, GetWindowLongPtr(this->hWndDialogContent, GWL_STYLE) | WS_VSCROLL | WS_HSCROLL | ES_AUTOVSCROLL | ES_AUTOHSCROLL);
	ShowWindow(this->hWndDialogContent, SW_SHOW);

	// Panel(Container) erstellen
	this->hWndPanel = CreateWindow(_T("STATIC"), NULL, WS_VISIBLE | WS_CHILD, 0, 0, 2 * this->iImgOffset + this->iMetafileViewWidth, 2 * this->iImgOffset + this->iMetafileViewHeight, this->hWndDialogContent, (HMENU)IDC_PANEL, (HINSTANCE)ghInstance, NULL);

	// Metafile-Container erstellen
	this->hWndMetafileView = CreateWindow(_T("STATIC"), NULL, WS_VISIBLE | WS_CHILD | SS_ENHMETAFILE, this->iImgOffset, this->iImgOffset, this->iMetafileViewWidth, this->iMetafileViewHeight, this->hWndPanel, (HMENU)IDC_METAFILEVIEW, (HINSTANCE)ghInstance, NULL);

	//Statusbar erstellen
	this->hWndStatusbar = CreateWindowEx(0, STATUSCLASSNAME, _T("I am a Status Bar"), SBARS_SIZEGRIP | WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, this->hWndDialog, (HMENU)IDC_PRINTPREVIEW_STATUSBAR, (HINSTANCE)ghInstance, NULL);
	ShowWindow(this->hWndStatusbar, TRUE);
	//Höhe der Statusbar ermitteln
	RECT statusBarRect;
	SendMessage(this->hWndStatusbar, SB_GETRECT, 0, (LPARAM)&statusBarRect);
	this->iStatusbarHeight = statusBarRect.bottom - statusBarRect.top;
	//Statusbar in zwei Teile teilen
	int iWidths[2];
	iWidths[0] = 150;
	iWidths[1] = -1;
	SendMessage(this->hWndStatusbar, SB_SETPARTS, 2, (LPARAM)iWidths);

	//Scrollbar Variablen initialisieren
	this->xMinScroll = 0;
	this->xMaxScroll = 0;
	this->xCurrentScroll = 0;
	this->yMinScroll = 0;
	this->yMaxScroll = 0;
	this->yCurrentScroll = 0;

	//Fenstereigenschaften initialisieren
	this->InitWindowSettings();
	//Starte den Vorschaudialog mit der ersten Seite
	this->StartPreview();
}

/*
Initialisiert den Dialog für die Eingabe einer Skalierung.
Es werden Titel und Texte in Abhängigkeit der ausgewählten Sprache gesetzt.
Parameter:
hWndUserDefinedScaleDlg:	Handle für den Dialog
*/
void PRINTPREVIEW::OnDlgUserDefinedScaleInit(HWND hWndUserDefinedScaleDlg)
{
	tstring tsTitle, tsText, tsTextOk, tsTextCancel;
	//Referenz auf das PrintPreview-Objekt als Eigenschaft des Fensters setzen, um von der Fensterprozedur auf das Objekt zugreifen zu können
	SetProp(hWndUserDefinedScaleDlg, _T("PROP_PREV_OBJECT"), (HANDLE)this);

	//Texte abfragen
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Title"), tsTitle);
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Input.Text"), tsText);
	tsText.append(_T(" ( ")).append(to_tstring(int(this->dMinScale))).append(_T(" - ")).append(to_tstring(int(this->dMaxScale))).append(_T(" ):"));
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Input.BtnOk"), tsTextOk);
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Input.BtnCancel"), tsTextCancel);

	//Texte setzen
	SetWindowText(hWndUserDefinedScaleDlg, tsTitle.c_str());
	SetWindowText(GetDlgItem(hWndUserDefinedScaleDlg, IDC_UDV_SCALE_TEXT), tsText.c_str());
	SetWindowText(GetDlgItem(hWndUserDefinedScaleDlg, IDOK), tsTextOk.c_str());
	SetWindowText(GetDlgItem(hWndUserDefinedScaleDlg, IDCANCEL), tsTextCancel.c_str());
}

/*
Überprüft, ob die eingegebene Skalierung ein valider Wert zwischen der minimalen und maximalen Skalierung ist.
Ist die Eingabe nicht korrekt, wird der Nutzer darüber informiert und kann eine neue Skalierung eingeben.
Ist die Eingabe korrekt, wird die neue Skalierung gesetzt und die neu skalierte Seite angezeigt.
Parameter:
hWndUserDefinedScaleDlg:	Handle für den Dialog
*/
void PRINTPREVIEW::OnDlgUserDefinedScaleOkClick(HWND hWndUserDefinedScaleDlg)
{
	tstring tsTitle, tsText;
	double dScalePercent;

	try
	{
		LPTSTR lpStr;
		lpStr = new TCHAR[100];
		//Eingabe des Nutzers ermitteln
		GetDlgItemText(hWndUserDefinedScaleDlg, IDC_UDV_SCALE_EDIT, lpStr, 100);
		tstring tString;
		tString = lpStr;
		delete[] lpStr;

		//Eingabe zu Skalierung
		dScalePercent = std::stoi(tString, NULL);

		//Skalierung im Bereich zwischen minimaler und maximaler Skalierung?
		if (dScalePercent < this->dMinScale || dScalePercent > this->dMaxScale)
			throw invalid_argument("");

		//Seite skalieren. Wenn skaliert wurde, Checkbox in der Toolbar setzen
		if (this->Scale(dScalePercent / 100.0))
			this->CheckScaleMenuItem(IDC_MENU_ITEM_SCALE_USER_DEFINED);

		//Dialog schließen
		SendMessage(hWndUserDefinedScaleDlg, WM_CLOSE, 0, 0);
	}
	catch (invalid_argument& e)
	{
		//Eingabe ist nicht korrekt, informiere Nutzer darüber
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Title"), tsTitle);
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.InvValue.Text"), tsText);
		MessageBox(hWndUserDefinedScaleDlg, tsText.c_str(), tsTitle.c_str(), MB_ICONASTERISK | MB_OK);
	}
	catch (out_of_range& e)
	{
		//Eingabe ist zu lang, informiere Nutzer darüber
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Title"), tsTitle);
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.InvValue.Text"), tsText);
		MessageBox(hWndUserDefinedScaleDlg, tsText.c_str(), tsTitle.c_str(), MB_ICONASTERISK | MB_OK);
	}
}

/*
Schließt den Vorschaudialog.
Alle Metafiles werden freigegeben, die deaktivierten Fenster werden wieder aktiviert und der Fokus wird auf das zuletzt aktive Fenster gesetzt.
Der Vorschaudialog und alle Kindfenster werden geschlossen.
*/
void PRINTPREVIEW::CloseBtnClick()
{
	this->FreeMetafiles();

	int size;
	size = this->vecHwndDisabled.size();
	for (int i = 0; i < size; i++)
		EnableWindow(this->vecHwndDisabled.at(i), TRUE);
	this->vecHwndDisabled.clear();
	SetForegroundWindow(this->hWndLastActiveWindow);

	EndDialog(this->hWndDialog, 0);
}

/*
Bestätigt die Druckvorschau und druckt die Seite/n.
Parameter:
uiId:	Control ID des geklickten Buttons
Abhängig von der Control ID werden entweder alle Seiten oder nur die aktuelle Seite gedruckt.
*/
void PRINTPREVIEW::PrintBtnClick(UINT uiId)
{
	//Sanduhr an
	this->SetWaitCursor(TRUE);
	//Alle Seiten drucken?
	if (uiId == IDC_MENU_BTN_PRINT || uiId == IDC_MENU_ITEM_PRINT_ALL)
	{
		PRINTOPTIONS prntOptions;
		prntOptions = this->printOptions;
		prntOptions.dScale = this->printMetrics.dScale;
		prntOptions.bPreview = FALSE;
		this->lpSubclassRef->Print(this->tableInfo.handle, prntOptions);
	}
	else if (uiId == IDC_MENU_ITEM_PRINT_CURRENT_PAGE)	//Aktuelle Seite drucken?
		this->PrintCurrentPage();

	//Sanduhr aus
	this->SetWaitCursor(FALSE);
}

/*
Öffnet den Dialog für die Eingabe einer benutzerdefinierten Skalierung.
Return erst, wenn der Dialog entweder geschlossen wurde oder eine korrekte Skalierung bestätigt wurde.
*/
void PRINTPREVIEW::UserDefinedScaleClick()
{
	DialogBoxParam((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDD_DIALOG3), this->hWndDialog, PreviewUserDefinedScaleProc, LPARAM(this));
}

/*
Führt eine Skalierungsoption durch.
Parameter:
uiId:	Control ID des geklickten Buttons
Bei IDC_MENU_ITEM_SCALE_USER_DEFINED wird der Dialog für die Eingabe einer benutzerdefinierten Skalierung geöffnet.
*/
void PRINTPREVIEW::ScaleMenuItemClick(UINT uiId)
{
	//Benutzerdefinierte Skalierung, Skalierung auf Seitenbreite oder vordefinierte Skalierung?
	if (uiId == IDC_MENU_ITEM_SCALE_USER_DEFINED)
		this->UserDefinedScaleClick();
	else if (uiId == IDC_MENU_ITEM_SCALE_TO_PAGE_WIDTH)
		this->ScaleToPageWidthClick();
	else
	{
		double dScale;
		MENUITEMINFO mii;

		// Daten ermitteln (= Skalierung in %)
		mii.cbSize = sizeof(MENUITEMINFO);
		mii.fMask = MIIM_DATA;
		if (GetMenuItemInfo(this->hMenuScale, uiId, FALSE, &mii))
		{
			dScale = double(mii.dwItemData) / 100;

			// Skalieren
			if (this->Scale(dScale))
			{
				// Eintrag "anhaken"
				this->CheckScaleMenuItem(uiId);
			}
		}
	}
}

/*
Zeigt die aktuell betrachtete Seite als ganze Seite an.
Alle Scrollbars werden so aktualisiert, das sie wegfallen.
*/
void PRINTPREVIEW::FullPageBtnClick()
{
	TBBUTTONINFO tbb;
	tbb.cbSize = sizeof(TBBUTTONINFO);
	tbb.dwMask = TBIF_STATE;
	SendMessage(this->hWndToolbar, TB_GETBUTTONINFO, IDC_MENU_BTN_FULLPAGE, LPARAM(&tbb));

	//Aktuell nicht im "Ganze Seite"-Modus? (Nach dem Button-Klick wird checked gesetzt, erst dann wird die Funktion aufgerufen)
	if (tbb.fsState & TBSTATE_CHECKED)
	{
		//"Ganze Seite"-Modus aktivieren
		this->bFullPage = TRUE;

		//Scroll Informationen zurücksetzen, sodass diese nicht angezegit werden
		this->xCurrentScroll = 0;
		this->yCurrentScroll = 0;

		SCROLLINFO scrollInfo;
		scrollInfo.cbSize = sizeof(SCROLLINFO);
		scrollInfo.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
		scrollInfo.nMin = this->xMinScroll;
		scrollInfo.nMax = this->xMaxScroll;
		scrollInfo.nPage = this->xMaxScroll - this->xMinScroll + 1;
		scrollInfo.nPos = this->xCurrentScroll;
		SetScrollInfo(this->hWndDialogContent, SB_HORZ, &scrollInfo, TRUE);
		scrollInfo.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
		scrollInfo.nMin = this->yMinScroll;
		scrollInfo.nMax = this->yMaxScroll;
		scrollInfo.nPage = this->yMaxScroll - this->yMinScroll + 1;
		scrollInfo.nPos = this->yCurrentScroll;
		SetScrollInfo(this->hWndDialogContent, SB_VERT, &scrollInfo, TRUE);

		//Aktuelle Seite anzeigen
		this->ShowPage(this->iCurrentPage);
	}
	else
	{
		//"Ganze Seite"-Modus deaktivieren
		//Fenstergrößen aktualisieren und Seite anzeigen
		RECT rect;
		GetWindowRect(this->hWndDialogContent, &rect);
		this->bFullPage = FALSE;
		this->SizeObjects();
		this->ShowPage(this->iCurrentPage);
		this->UpdateScrollbar((rect.right - rect.left), (rect.bottom - rect.top));
	}
}

/*
Zeigt die erste Seite an, wenn sie nicht aktuell schon angezeigt wird.
*/
void PRINTPREVIEW::FirstPageBtnClick()
{
	if (this->iCurrentPage != 1)
		this->ShowPage(1);
}

/*
Zeigt die vorherige Seite an, wenn aktuell nicht die erste Seite angezeigt wird.
*/
void PRINTPREVIEW::PrevPageBtnClick()
{
	if (this->iCurrentPage != 1)
		this->ShowPage(this->iCurrentPage - 1);
}

/*
Zeigt die nächste Seite an, wenn aktuell nicht die letzte Seite angezeigt wird.
*/
void PRINTPREVIEW::NextPageBtnClick()
{
	if (this->iCurrentPage != this->iLastPage)
		this->ShowPage(this->iCurrentPage + 1);
}

/*
Zeigt die letzte Seite an.
Wenn die Anzahl der Seiten noch nicht bekannt ist, also noch keine Metafiles dafür erstellt wurden,
erstelle die restlichen Metafiles. Zeige anschließend die letzte Seite an.
*/
void PRINTPREVIEW::LastPageBtnClick()
{
	int iPage, iReturn;
	tstring tString, tsTmpString;

	//Letzte Seite unbekannt?
	if (this->iLastPage == PRINTPREVIEW_UNKNOWN)
	{
		//Fehlende Metafiles erzeugen
		//Sanduhr an
		this->SetWaitCursor(TRUE);

		//Erste zu erstellende Seite ermitteln
		iPage = this->vecPageMetafiles.size() + 1;

		//Text für Statusbar ermitteln
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.StatBar.CreatePage"), tString);

		//Alle restlichen Metafiles erstellen
		while (TRUE)
		{
			//Platzhalter im Text für die Statusbar durch die momentan erstellte Seite ersetzen
			tsTmpString = tString;
			this->ReplaceSubstring(tsTmpString, tstring(_T("%d")), to_tstring(iPage));
			SendMessage(this->hWndStatusbar, SB_SETTEXT, 0, (LPARAM)tsTmpString.c_str());

			//Metafile erzeugen
			iReturn = this->PrintCreateEnhMetafile(iPage);
			if (iReturn == -2 || this->iLastPage != PRINTPREVIEW_UNKNOWN)
				break;
			iPage++;
		}
		//Sanduhr aus
		this->SetWaitCursor(FALSE);
	}
	//Letzte Seite anzeigen
	this->ShowPage(this->iLastPage);
}

/*
Öffnet ein Popup-Menü der Toolbar.
Parameter:
lpnmtb:		Toolbar Struktur, die Informationen über den geklickten Button enthält
*/
void PRINTPREVIEW::DropDownClick(LPNMTOOLBAR lpnmtb)
{
	HMENU hMenu;
	RECT rBtn;
	TPMPARAMS tpmp;

	// Koordinaten des Buttons ermitteln
	SendMessage(lpnmtb->hdr.hwndFrom, TB_GETRECT, (WPARAM)lpnmtb->iItem, (LPARAM)&rBtn);

	// In Bildschirmkoordinaten umwandeln
	MapWindowPoints(lpnmtb->hdr.hwndFrom, HWND_DESKTOP, (LPPOINT)&rBtn, 2);

	// Anzuzeigendes Popup-Menüs ermitteln
	if (lpnmtb->iItem == IDC_MENU_BTN_PRINT)
		hMenu = this->hMenuPrint;
	else if (lpnmtb->iItem == IDC_MENU_BTN_SCALING)
		hMenu = this->hMenuScale;

	// Popup-Menü anzeigen
	tpmp.cbSize = sizeof(TPMPARAMS);
	tpmp.rcExclude = rBtn;

	TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_VERTICAL, rBtn.left, rBtn.bottom, this->hWndDialog, &tpmp);
}

/*
Scrollt die aktuell angezeigte Seite in vertikaler Richtung.
Diese Funktion wird verwendet, um das Scrollen des Nutzers mit dem Mausrad zu verarbeiten
Parameter:
sDeltaScroll:	Distanz, um die gescrollt werden soll.
Der Wert ist negativ, wenn der Nutzer zu sich hin scrollt.
Der Wert ist positiv, wenn der Nutzer von sich weg scrollt.
*/
void PRINTPREVIEW::OnMouseWheelScroll(short sDeltaScroll)
{
	//Aktuelle Scrollposition ermitteln
	SCROLLINFO scrollInfo;
	scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_POS;
	GetScrollInfo(this->hWndDialogContent, SB_VERT, &scrollInfo);

	//Neue Scrollposition bestimmen
	short sNewScrollPos = scrollInfo.nPos - sDeltaScroll;
	//Scrollen mit neuer Thumbposition
	if (sNewScrollPos > 0)
		this->OnVerticalScroll(SB_THUMBPOSITION, (WORD)sNewScrollPos);
	else
		this->OnVerticalScroll(SB_THUMBPOSITION, 0);
}

/*
Scrollt die aktuelle Seite in horizontaler Richtung.
Überprüft, welche Art von Scrolling gefragt ist und scrollt dementsprechend (Seite hoch/herunter, Zeile hoch/herunter, Thumbpositionsänderung, Thumbdrag).
Paramter:
wScrollRequest:		Scrollanfrage (SB_PAGEUP, SB_PAGEDOWN, SB_LINEUP, SB_LINEDOWN, SB_THUMBPOSITION, SB_THUMBTRACK)
wScrollboxPosition:	Neue Position der Scrollbox (Thumbposition). Wird nur bei SB_THUMBPOSITION und SB_THUMBTRACK verwendet
*/
void PRINTPREVIEW::OnHorizontalScroll(WORD wScrollRequest, WORD wScrollboxPosition)
{
	SCROLLINFO scrollInfo;
	int xDelta;
	int yDelta;
	int xNewPos;
	yDelta = 0;

	switch (wScrollRequest)
	{
		//Nutzer hat in den Scrollbar Bereich geklickt
	case SB_PAGEUP:
		xNewPos = this->xCurrentScroll - 50;
		break;

	case SB_PAGEDOWN:
		xNewPos = this->xCurrentScroll + 50;
		break;

		//Nutzer hat den linken Pfeil geklickt
	case SB_LINEUP:
		xNewPos = this->xCurrentScroll - 5;
		break;

	case SB_LINEDOWN:
		xNewPos = this->xCurrentScroll + 5;
		break;

		//Nutzer bewegt den Scrollbar Handler
	case SB_THUMBPOSITION:
		xNewPos = wScrollboxPosition;
		break;

	case SB_THUMBTRACK:
		xNewPos = wScrollboxPosition;
		break;

	default:
		xNewPos = this->xCurrentScroll;
	}

	//Neue Position muss sich innerhalb der Bildschirmbreite befinden
	xNewPos = max(0, xNewPos);
	xNewPos = min(this->xMaxScroll, xNewPos);

	//Keine Bewegeung bei keienr Änderung
	if (xNewPos == this->xCurrentScroll)
		return;

	//Scrollbar aktualisieren
	scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_POS;
	scrollInfo.nPos = xNewPos;
	SetScrollInfo(this->hWndDialogContent, SB_HORZ, &scrollInfo, TRUE);

	//nPos wurde eventuell modifiziert
	GetScrollInfo(this->hWndDialogContent, SB_HORZ, &scrollInfo);

	//Anzahl Pixel gescrollt
	xDelta = scrollInfo.nPos - this->xCurrentScroll;

	//Window scrollen
	ScrollWindowEx(this->hWndDialogContent, -xDelta, -yDelta, NULL, NULL, NULL, NULL, SW_SCROLLCHILDREN | SW_INVALIDATE | SW_ERASE);
	this->xCurrentScroll = scrollInfo.nPos;
}

/*
Scrollt die aktuelle Seite in vertikaler Richtung.
Überprüft, welche Art von Scrolling gefragt ist und scrollt dementsprechend (Seite hoch/herunter, Zeile hoch/herunter, Thumbpositionsänderung, Thumbdrag).
Paramter:
wScrollRequest:		Scrollanfrage (SB_PAGEUP, SB_PAGEDOWN, SB_LINEUP, SB_LINEDOWN, SB_THUMBPOSITION, SB_THUMBTRACK)
wScrollboxPosition:	Neue Position der Scrollbox (Thumbposition). Wird nur bei SB_THUMBPOSITION und SB_THUMBTRACK verwendet
*/
void PRINTPREVIEW::OnVerticalScroll(WORD wScrollRequest, WORD wScrollboxPosition)
{
	SCROLLINFO scrollInfo;
	int xDelta;
	int yDelta;
	int yNewPos;
	xDelta = 0;

	switch (wScrollRequest)
	{
		//Nutzer hat in den Scrollbar Bereich geklickt
	case SB_PAGEUP:
		yNewPos = this->yCurrentScroll - 50;
		break;

	case SB_PAGEDOWN:
		yNewPos = this->yCurrentScroll + 50;
		break;

		//Nutzer hat den linken Pfeil geklickt
	case SB_LINEUP:
		yNewPos = this->yCurrentScroll - 5;
		break;

	case SB_LINEDOWN:
		yNewPos = this->yCurrentScroll + 5;
		break;

		//Nutzer bewegt den Scrollbar Handler
	case SB_THUMBPOSITION:
		yNewPos = wScrollboxPosition;
		break;

	case SB_THUMBTRACK:
		yNewPos = wScrollboxPosition;
		break;

	default:
		yNewPos = this->yCurrentScroll;
	}

	//Neue Position muss sich innerhalb der Bildschirmbreite befinden
	yNewPos = max(0, yNewPos);
	yNewPos = min(this->yMaxScroll, yNewPos);

	//Keine Bewegeung bei keienr Änderung
	if (yNewPos == this->yCurrentScroll)
		return;

	//Scrollbar aktualisieren
	scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_POS;
	scrollInfo.nPos = yNewPos;
	SetScrollInfo(this->hWndDialogContent, SB_VERT, &scrollInfo, TRUE);

	//nPos wurde eventuell modifiziert
	GetScrollInfo(this->hWndDialogContent, SB_VERT, &scrollInfo);

	//Anzahl Pixel gescrollt
	yDelta = scrollInfo.nPos - this->yCurrentScroll;

	//Window scrollen
	ScrollWindowEx(this->hWndDialogContent, -xDelta, -yDelta, NULL, NULL, NULL, NULL, SW_SCROLLCHILDREN | SW_INVALIDATE | SW_ERASE);
	this->yCurrentScroll = scrollInfo.nPos;
}

/*
Aktualisiert die Scrollbars, wenn "Ganze Seiet"-Modus nicht aktiv ist.
Ansonsten wird die aktuelle Seite neu angezeigt.
Parameter:
wNewWidth:	Neue Breite des Inhaltsfensters
wNewHeight:	Neue Höhe des Inhaltsfensters
*/
void PRINTPREVIEW::OnContentSizeChanged(WORD wNewWidth, WORD wNewHeight)
{
	//Nicht im "Ganze Seite"-Modus?
	if (!this->bFullPage)
		this->UpdateScrollbar(wNewWidth, wNewHeight);
	else
		this->ShowPage(this->iCurrentPage);
}

/*
Erstellt den Vorschaudialog.
Initialisiert Objekteigenschaften, bestimmt die Größe des Metafile-Panels und öffnet den Dialog.
Return erst, wenn der Dialog beendet wurde.
*/
void PRINTPREVIEW::CreatePreviewWindow()
{
	this->dMinScale = 1.0;
	this->dMaxScale = 500.0;
	this->iImgOffset = 10;

	this->iCurrentPage = 0;
	this->iLastPage = PRINTPREVIEW_UNKNOWN;
	this->vecPageStartRows.push_back(-1);
	this->vecPageFromCols.push_back(-1);
	this->vecPageToCols.push_back(-1);
	this->vecPageLastTotals.push_back(-1);
	this->iMenuPanelHeight = 40;
	this->iMenuButtonWidth = 85;

	this->hImageList = NULL;
	this->hWndToolbar = NULL;
	this->hMenuPrint = NULL;
	this->hMenuScale = NULL;
	this->bWaitCursor = FALSE;

	//Größe des Panels, in dem die Metafile angezeigt wird
	this->iMetafileViewWidth = round(printMetrics.iPhysWidth / (double)(printMetrics.iPixelPerInchX) * tableInfo.iPixelPerInchX);
	this->iMetafileViewHeight = round(printMetrics.iPhysHeight / (double)(printMetrics.iPixelPerInchY) * tableInfo.iPixelPerInchY);

	//Dialog erstellen
	DialogBoxParam((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDD_DIALOG1), 0, DialogProc, LPARAM(this));
}

/*
Erstellt die Toolbar des Vorschaudialogs.
Alle Buttons und Bilder der Toolbar werden erstellt und alle Texte der Buttons gesetzt.
Parameter:
hWndParent: Handle für den Vorschaudialog
*/
void PRINTPREVIEW::CreateToolbar(HWND hWndParent)
{
	BOOL bOk;
	HBITMAP hbm;
	INITCOMMONCONTROLSEX iccex;
	int iBmIdx;
	TBBUTTON tbb;
	tstring tsLangStr;

	iccex.dwSize = sizeof(INITCOMMONCONTROLSEX);
	iccex.dwICC = ICC_BAR_CLASSES;
	bOk = InitCommonControlsEx(&iccex);

	// *** Image-List für Toolbar ***
	this->hImageList = ImageList_Create(16, 16, ILC_MASK, 9, 0);
	hbm = LoadBitmap((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDB_PP_TOOLBAR));
	iBmIdx = ImageList_AddMasked(this->hImageList, hbm, RGB(255, 0, 255));
	DeleteObject(hbm);

	// *** Toolbar ***
	this->hWndToolbar = CreateWindowEx(0, TOOLBARCLASSNAME, NULL, WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT, 0, 0, 0, 0, hWndParent, (HMENU)IDC_TOOLBAR, (HINSTANCE)ghInstance, NULL);
	SendMessage(this->hWndToolbar, TB_BUTTONSTRUCTSIZE, (WPARAM)sizeof(TBBUTTON), 0);
	SendMessage(this->hWndToolbar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS);
	SendMessage(this->hWndToolbar, TB_SETIMAGELIST, (WPARAM)0, (LPARAM)this->hImageList);

	// *** Toolbar-Buttons ***

	// Button "Close"
	tbb.iBitmap = 0;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_BUTTON;
	tbb.idCommand = IDC_MENU_BTN_CLOSE;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnClose"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	//Seperator
	tbb.fsStyle = TBSTYLE_SEP;
	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "Print"
	tbb.iBitmap = 1;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_DROPDOWN;
	tbb.idCommand = IDC_MENU_BTN_PRINT;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnPrint"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	//Seperator
	tbb.fsStyle = TBSTYLE_SEP;
	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "Scaling"
	tbb.iBitmap = 7;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_DROPDOWN;
	tbb.idCommand = IDC_MENU_BTN_SCALING;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnScale"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	//Seperator
	tbb.fsStyle = TBSTYLE_SEP;
	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "Full Page"
	tbb.iBitmap = 9;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_CHECK;
	tbb.idCommand = IDC_MENU_BTN_FULLPAGE;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnFullPage"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	//Seperator
	tbb.fsStyle = TBSTYLE_SEP;
	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "First Page"
	tbb.iBitmap = 5;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_BUTTON;
	tbb.idCommand = IDC_MENU_BTN_FIRSTPAGE;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnFirstPage"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "Previous Page"
	tbb.iBitmap = 8;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_BUTTON;
	tbb.idCommand = IDC_MENU_BTN_PREVPAGE;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnPrevPage"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "Next Page"
	tbb.iBitmap = 3;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_BUTTON;
	tbb.idCommand = IDC_MENU_BTN_NEXTPAGE;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnNextPage"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button "Last Page"
	tbb.iBitmap = 4;
	tbb.fsState = TBSTATE_ENABLED;
	tbb.fsStyle = TBSTYLE_BUTTON;
	tbb.idCommand = IDC_MENU_BTN_LASTPAGE;
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnLastPage"), tsLangStr);
	tbb.iString = (INT_PTR)tsLangStr.c_str();

	SendMessage(this->hWndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

	// Button-Größe setzen
	SendMessage(this->hWndToolbar, TB_SETBUTTONSIZE, 0, MAKELPARAM(this->iMenuButtonWidth, this->iMenuPanelHeight - 2));
}

/*
Erstellt die Popup-Menüs für die Dropdowns des Druck- und Skalierungsbuttons der Toolbar.
Die Buttons werden erstellt und mit Information versehen.
*/
void PRINTPREVIEW::CreatePopupMenus()
{
	BOOL bOk;
	HMENU hMenu;
	MENUITEMINFO mii = {};
	tstring tsLangStr;
	UINT uiFlags, uiId;

	// *** Popup-Menü "Print" ***
	this->hMenuPrint = hMenu = CreatePopupMenu();

	// Eintrag "All"
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnPrint.MenuItem.All"), tsLangStr);
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_PRINT_ALL;
	bOk = AppendMenu(hMenu, uiFlags, uiId, tsLangStr.c_str());
	bOk = SetMenuDefaultItem(hMenu, uiId, FALSE);

	// Eintrag "Current Page"
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnPrint.MenuItem.CurrPage"), tsLangStr);
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_PRINT_CURRENT_PAGE;
	bOk = AppendMenu(hMenu, uiFlags, uiId, tsLangStr.c_str());


	// *** Popup-Menü "Scale" ***
	this->hMenuScale = hMenu = CreatePopupMenu();

	// MENUITEMINO zum Setzen von Typ und Daten
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_FTYPE | MIIM_DATA;

	// Eintrag "User Defined"
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnScale.MenuItem.UserDef"), tsLangStr);
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_USER_DEFINED;
	bOk = AppendMenu(hMenu, uiFlags, uiId, tsLangStr.c_str());
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)0;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);
	bOk = SetMenuDefaultItem(hMenu, uiId, FALSE);

	// Trenner
	uiFlags = MF_SEPARATOR;
	uiId = 0;
	bOk = AppendMenu(hMenu, uiFlags, 0, NULL);

	// Eintrag "10 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_10;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("10 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)10;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Eintrag "25 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_25;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("25 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)25;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Eintrag "50 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_50;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("50 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)50;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Eintrag "75 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_75;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("75 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)75;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Eintrag "100 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_100;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("100 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)100;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Eintrag "150 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_150;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("150 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)150;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Eintrag "200 %"
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_200;
	bOk = AppendMenu(hMenu, uiFlags, uiId, _T("200 %"));
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)200;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Trenner
	uiFlags = MF_SEPARATOR;
	uiId = 0;
	bOk = AppendMenu(hMenu, uiFlags, 0, NULL);

	// Eintrag "To Page Width"
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.BtnScale.MenuItem.ToPageWid"), tsLangStr);
	uiFlags = MF_STRING | MF_ENABLED;
	uiId = IDC_MENU_ITEM_SCALE_TO_PAGE_WIDTH;
	bOk = AppendMenu(hMenu, uiFlags, uiId, tsLangStr.c_str());
	mii.fType = MFT_RADIOCHECK;
	mii.dwItemData = (ULONG_PTR)0;
	bOk = SetMenuItemInfo(hMenu, uiId, FALSE, &mii);

	// Je nach Skalierung das entsprechende Item anhaken
	if (this->printMetrics.dScale == 0.1)
		uiId = IDC_MENU_ITEM_SCALE_10;
	else if (this->printMetrics.dScale == 0.25)
		uiId = IDC_MENU_ITEM_SCALE_25;
	else if (this->printMetrics.dScale == 0.5)
		uiId = IDC_MENU_ITEM_SCALE_50;
	else if (this->printMetrics.dScale == 0.75)
		uiId = IDC_MENU_ITEM_SCALE_75;
	else if (this->printMetrics.dScale == 1.0)
		uiId = IDC_MENU_ITEM_SCALE_100;
	else if (this->printMetrics.dScale == 1.5)
		uiId = IDC_MENU_ITEM_SCALE_150;
	else if (this->printMetrics.dScale == 2.0)
		uiId = IDC_MENU_ITEM_SCALE_200;
	else
		uiId = IDC_MENU_ITEM_SCALE_USER_DEFINED;

	this->CheckScaleMenuItem(uiId);
}

/*
Setzt Daten aus den Druckoptionen als Starteigenschaften des Vorschaudialogs.
Zu den Daten gehören Titel, Position und Größe des Vorschaudialogs und ob sich der Dialog zu Beginn im "Ganze Seite"-Modus befindet.
*/
void PRINTPREVIEW::InitWindowSettings()
{
	//Titel
	//Vorschaudialog-Titel in Druckoptionen angegeben?
	if (!this->printOptions.tsPrevWndTitle.empty())
		SendMessage(this->hWndDialog, WM_SETTEXT, 0, (LPARAM)this->printOptions.tsPrevWndTitle.c_str());
	else
	{
		//Kein Titel angegeben, verwende Standardtitel
		tstring tsTitle;
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.DefCaption"), tsTitle);

		//Druckername ermitteln
		tstring tsPrinterName;
		//Drucker angegeben?
		if (this->printOptions.tsPrinterName.empty())
			tsPrinterName = this->printOptions.tsPrinterName;
		else
		{
			//Kein Drucker angegeben, nehme Standarddrucker
			HSTRING hsPrinterName = MTblPrintGetDefPrinterName();
			Ctd.HStrToTStr(hsPrinterName, tsPrinterName, TRUE);
		}
		if (!tsPrinterName.empty())
		{
			tsTitle.append(_T(" (")).append(tsPrinterName).append(_T(")"));
		}
		SendMessage(this->hWndDialog, WM_SETTEXT, 0, (LPARAM)tsTitle.c_str());
	}

	//Position und Größe
	MINMAXINFO minMaxInfo = {};
	SendMessage(this->hWndDialog, WM_GETMINMAXINFO, 0, (LPARAM)&minMaxInfo);
	SetWindowPos(this->hWndDialog, 0, this->printOptions.iPrevWndLeft, this->printOptions.iPrevWndTop,
		max(this->printOptions.iPrevWndWidth, minMaxInfo.ptMinTrackSize.x),
		max(this->printOptions.iPrevWndHeight, minMaxInfo.ptMinTrackSize.y), SWP_NOZORDER);

	ShowWindow(this->hWndDialog, this->printOptions.iPrevWndState + 1);

	//Ganze Seite?
	if (this->printOptions.bPrevWndFullPage)
	{
		//"Ganze Seite"-Button als checked setzen
		TBBUTTONINFO buttonInfo;
		buttonInfo.cbSize = sizeof(TBBUTTONINFO);
		buttonInfo.dwMask = TBIF_STATE;
		SendMessage(this->hWndToolbar, TB_GETBUTTONINFO, IDC_MENU_BTN_FULLPAGE, LPARAM(&buttonInfo));
		buttonInfo.fsState = buttonInfo.fsState | TBSTATE_CHECKED;
		SendMessage(this->hWndToolbar, TB_SETBUTTONINFO, IDC_MENU_BTN_FULLPAGE, LPARAM(&buttonInfo));
		this->bFullPage = TRUE;
	}

	//Fenstergröße bestimmen
	this->SizeObjects();
}

/*
Startet den Vorschaudialog mit der ersten Seite.
*/
void PRINTPREVIEW::StartPreview()
{
	//"Ganze Seite"-Modus?
	if (this->bFullPage)
	{
		this->iCurrentPage = 1;
		this->FullPageBtnClick();
	}
	else
		this->NextPageBtnClick();

	//Seite hat sich geändert
	this->PageChanged();
}

/*
Setzt einen Cursor auf Normal oder Warten.
Parameter:
bWait:	Cursor soll auf Warten gesetzt werden ja/nein
*/
void PRINTPREVIEW::SetWaitCursor(BOOL bWait)
{
	POINT pt;
	this->bWaitCursor = bWait;
	GetCursorPos(&pt);
	SetCursorPos(pt.x, pt.y);
	SetCursor(LoadCursor(NULL, IDC_WAIT));
}

/*
Erstellt, wenn nötig, eine neue Metafile für eine anzuzeigende Seite.
Parameter:
iPageNr:	Seitennummer, für die eine Metafile erstellt werden soll
Ist für die Seitennummer schon eine Metafile vorhanden, wird keine neue erstellt.
Return:
Die erste Zeile der nächsten Seite.
-1 wenn die letzte Seite erstellt wurde.
-2 wenn ein Fehler aufgetreten ist.
*/
int PRINTPREVIEW::PrintCreateEnhMetafile(int iPageNr)
{
	long lColNext, lColFrom, lColTo, lColLow, lColHigh, lNextStartRow, lLastRow, lLastTotal, lReturn;
	BOOL bFound;
	HENHMETAFILE hEnhMetafile;
	int iResult = -2;
	HDC hdcMetafile;

	lColLow = 0;
	lColHigh = this->vecColsInfo.size() - 1;

	int iTmp = (this->vecPageMetafiles.size() - 1);
	//Metafile noch nicht vorhanden?
	if (iTmp < (iPageNr - 1))
	{
		//Metafile erzeugen
		hdcMetafile = CreateEnhMetaFile(this->hdcPrinter, NULL, NULL, NULL);

		//Von-/Bis-Spalten ermitteln
		lColFrom = lColLow;
		lColTo = lColLow;

		//Nicht die erste Seite?
		if (iPageNr - 2 >= 0)
		{
			if (this->vecPageToCols.at(iPageNr - 2) < lColHigh)
				lColFrom = this->vecPageToCols.at(iPageNr - 2) + 1;
		}

		lColNext = lColFrom + 1;
		bFound = FALSE;

		//Ermittle die Spalten der Seite
		while (lColNext <= lColHigh)
		{
			if (this->vecColsInfo.at(lColNext).bNewPage)
			{
				lColTo = lColNext - 1;
				bFound = TRUE;
				break;
			}
			lColNext++;
		}

		if (!bFound)
			lColTo = lColHigh;

		//Von-/Bis-Spalten setzen
		iTmp = this->vecPageFromCols.size() - 1;
		if (iTmp < iPageNr - 1)
		{
			while (this->vecPageFromCols.size() < iPageNr)
				this->vecPageFromCols.push_back(-1);
		}
		this->vecPageFromCols.at(iPageNr - 1) = lColFrom;

		iTmp = this->vecPageToCols.size() - 1;
		if (iTmp < iPageNr - 1)
		{
			while (this->vecPageToCols.size() < iPageNr)
				this->vecPageToCols.push_back(-1);
		}
		this->vecPageToCols.at(iPageNr - 1) = lColTo;

		//Seitenparameter setzen
		this->printPageParams.hdcCanvas = hdcMetafile;
		this->printPageParams.iPageNr = iPageNr;
		this->printPageParams.lStartRow = vecPageStartRows.at(iPageNr - 1);
		this->printPageParams.iFromCol = lColFrom;
		this->printPageParams.iToCol = lColTo;
		this->printPageParams.printOptions = this->printOptions;
		this->printPageParams.printMetrics = this->printMetrics;
		this->printPageParams.tableInfo = this->tableInfo;
		this->printPageParams.vecColsInfo = this->vecColsInfo;
		this->printPageParams.bColHeaders = this->printOptions.bColHeaders && (this->printPageParams.lStartRow != -3);
		this->printPageParams.bCells = this->printPageParams.lStartRow != -3;

		//Hintergrund weiß zeichnen
		HBRUSH hBrush = CreateSolidBrush(0xffffff);
		RECT brushRect;
		brushRect.left = iImgOffset;
		brushRect.top = iImgOffset;
		brushRect.right = GetDeviceCaps(this->hdcPrinter, HORZRES);
		brushRect.bottom = GetDeviceCaps(this->hdcPrinter, VERTRES);
		FillRect(hdcMetafile, &brushRect, hBrush);

		//Seite in Metafile drucken
		lLastTotal = this->vecPageLastTotals.at(iPageNr - 1);
		lReturn = lpSubclassRef->PrintPage(this->printPageParams, lLastRow, lLastTotal);

		//Handle für Metafile erzeugen
		hEnhMetafile = CloseEnhMetaFile(hdcMetafile);

		DeleteDC(hdcMetafile);

		//Metafile in Array speichern
		PRINTMETAFILE metafile;
		while (this->vecPageMetafiles.size() < iPageNr)
			this->vecPageMetafiles.push_back(metafile);
		this->vecPageMetafiles.at(iPageNr - 1).hMetafile = hEnhMetafile;
		this->vecPageMetafiles.at(iPageNr - 1).iRows = lReturn;
	}
	else
	{
		//Metafile vorhanden
		lReturn = this->vecPageMetafiles.at(iPageNr - 1).iRows;
		lColTo = this->vecPageToCols.at(iPageNr - 1);
		lLastTotal = this->vecPageLastTotals.at(iPageNr - 1);
	}

	//Metafile erzeugt und letzte Seite wurde erstellt?
	if (lReturn == -1 && lColTo == lColHigh)
	{
		//Letzte Seite setzen
		iLastPage = iPageNr;
		//Letzte Seite erreicht
		iResult = -1;
	}
	else
	{
		//Startzeile der nächsten Seite ermitteln
		if (lColTo == lColHigh)
			lNextStartRow = lReturn;
		else
			lNextStartRow = this->vecPageStartRows.at(iPageNr - 1);

		//Startzeile der nächsten Seite setzen
		iTmp = this->vecPageStartRows.size() - 1;
		if (iTmp < iPageNr)
		{
			while (this->vecPageStartRows.size() < iPageNr + 1)
				this->vecPageStartRows.push_back(-1);
		}
		this->vecPageStartRows.at(iPageNr) = lNextStartRow;

		//Letzte Total-Zeile der nächsten Seite setzen
		iTmp = this->vecPageLastTotals.size() - 1;
		if (iTmp < iPageNr)
		{
			while (this->vecPageLastTotals.size() < iPageNr + 1)
				this->vecPageLastTotals.push_back(-1);
		}
		this->vecPageLastTotals.at(iPageNr) = lLastTotal;

		iResult = lNextStartRow;
	}
	return iResult;
}

/*
Zeigt eine Seite im Vorschaudialog an.
Die Seite wird bei Bedarf erstellt und dann gesetzt.
Parameter:
iPageNr:	Seite, die angezeigt werden soll
Return:
0 wenn die Seite erfolgreich angezeigt wurde.
-2 wenn ein Fehler aufgetreten ist.
*/
int PRINTPREVIEW::ShowPage(int iPageNr)
{
	int iResult;
	//Sanduhr an
	this->SetWaitCursor(TRUE);

	//Metafile erzeugen
	iResult = this->PrintCreateEnhMetafile(iPageNr);

	//Fehler?
	if (iResult == -2)
		return iResult;

	//"Ganze Seite"-Modus?
	if (this->bFullPage)
		this->ScaleContentToSize();

	//Metafile in Container setzen
	SendMessage(this->hWndMetafileView, STM_SETIMAGE, IMAGE_ENHMETAFILE, (LPARAM)this->vecPageMetafiles.at(iPageNr - 1).hMetafile);

	if (iPageNr != this->iCurrentPage)
	{
		//Aktuelle Seite setzen
		this->iCurrentPage = iPageNr;
		//Seite hat sich geändert
		this->PageChanged();
	}

	iResult = 0;
	//Sanduhr aus
	this->SetWaitCursor(FALSE);
	return iResult;
}

/*
Setzt Position und Größe von dem Inhaltsteil des Vorschaudialogs.
Der Metafile-Container befindet sich mit einem Offset in einen Rahmen, der sich wiederum in dem Inhaltscontainer des Vorschaudialogs befindet.
*/
void PRINTPREVIEW::SizeObjects()
{
	int l, t, iWidth, iHeight;

	l = this->iImgOffset;
	t = this->iImgOffset;
	//Breite und Höhe des Metafile-Containers
	iWidth = round(this->printMetrics.iPhysWidth / (double)(this->printMetrics.iPixelPerInchX) * this->tableInfo.iPixelPerInchX);
	iHeight = round(this->printMetrics.iPhysHeight / (double)(this->printMetrics.iPixelPerInchY) * this->tableInfo.iPixelPerInchY);

	SetWindowPos(this->hWndMetafileView, 0, l, t, iWidth, iHeight, SWP_NOZORDER | SWP_NOACTIVATE);

	iWidth += this->iImgOffset * 2;
	iHeight += this->iImgOffset * 2;

	SetWindowPos(this->hWndPanel, 0, 0, 0, iWidth, iHeight, SWP_NOZORDER | SWP_NOACTIVATE);

	RECT clientRect;
	GetClientRect(this->hWndDialog, &clientRect);

	SetWindowPos(this->hWndDialogContent, 0, 0, this->iMenuPanelHeight, clientRect.right, clientRect.bottom - this->iMenuPanelHeight - this->iStatusbarHeight, SWP_NOZORDER);
}

/*
Skaliert den Inhalt der anzuzeigenden Seiten.
Wenn aktuell nicht die erste Seite betrachtet wird, wird der Nutzer dazu aufgefordert, zur ersten Seite zu gehen.
Eine Skalierung findet nur statt, wenn der Nutzer zur ersten Seite geht.
Parameter:
dScale:		Die neue Skalierung
Return:
TRUE, wenn eine Skalierung durchgeführt wurde.
FALSE, wenn der Nutzer nicht zur ersten Seite gehen möchte, oder keine Skalierung durchgeführt wurde.
*/
BOOL PRINTPREVIEW::Scale(double dScale)
{
	// Ende, wenn Skalierung sich nicht ändert
	if (dScale == this->printMetrics.dScale)
		return FALSE;

	// Aktuelle Seite nicht Seite 1?
	if (this->iCurrentPage != 1)
	{
		//Nutzer fragen, ob er zur ersten Seite gehen möchte. Nur dann wird die Skalierung durchgeführt
		tstring tsTitle, tsText;
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Title"), tsTitle);
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.GoToFirstPage.Text"), tsText);
		if (MessageBox(NULL, tsText.c_str(), tsTitle.c_str(), MB_ICONQUESTION | MB_YESNO) == IDNO)
			return FALSE;
	}

	// Sanduhr an
	this->SetWaitCursor(TRUE);

	// Neue Skalierung setzen
	this->printOptions.dScale = dScale;

	// Neue Druckabmessungen ermitteln
	this->lpSubclassRef->GetPrintMetrics(this->hdcPrinter, this->printOptions, this->tableInfo, this->vecColsInfo, &this->printMetrics);

	// Neue Anzahl Seiten ermitteln
	if (this->lpSubclassRef->IsPageCountNeeded(this->printOptions))
		mtblPrint_PageCount = this->lpSubclassRef->GetPageCount(this->printOptions, this->printMetrics, this->tableInfo, this->vecColsInfo, this->hdcPrinter);

	// Erste Seite anzeigen
	this->iCurrentPage = 0;
	this->iLastPage = PRINTPREVIEW_UNKNOWN;

	this->vecPageStartRows.clear();
	this->vecPageStartRows.push_back(-1);

	this->vecPageFromCols.clear();
	this->vecPageFromCols.push_back(-1);

	this->vecPageToCols.clear();
	this->vecPageToCols.push_back(-1);

	this->vecPageLastTotals.clear();
	this->vecPageLastTotals.push_back(-1);

	this->FreeMetafiles();

	// Sanduhr aus
	this->SetWaitCursor(FALSE);

	this->NextPageBtnClick();

	//Statusbar aktualisieren
	this->SetStatusbar();

	return TRUE;
}

/*
Skaliert die anzuzeigenden Seiten auf die Seitenbreite.
Ist die ermittelte Skalierung größer als die maximale Skalierung, wird der Nutzer gefragt, ob er auf die maximale Skalierung skalieren möchte.
Nur dann wird die Skalierung durchgeführt.
*/
void PRINTPREVIEW::ScaleToPageWidthClick()
{
	double dScale;
	BOOL bDoIt;

	bDoIt = TRUE;
	//Skalierung auf Seitenbreite
	dScale = (this->printMetrics.clipRect.right - this->printMetrics.clipRect.left) / (double)(this->printMetrics.iTotalColWidth);

	if (dScale * 100 > this->dMaxScale)
	{
		//Skalierung größer als maximale Skalierung. Frage, ob auf maximale Skalierung skaliert werden soll.
		tstring tsTitle, tsText;
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.Title"), tsTitle);
		this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.Msg.Scale.MaxScaleExceeded.Text"), tsText);
		this->ReplaceSubstring(tsText, tstring(_T("%s")), to_tstring((int)this->dMaxScale));
		this->ReplaceSubstring(tsText, tstring(_T("%s")), to_tstring((int)this->dMaxScale));
		this->ReplaceSubstring(tsText, tstring(_T("%%")), tstring(_T("%")));
		this->ReplaceSubstring(tsText, tstring(_T("%%")), tstring(_T("%")));

		if (MessageBox(NULL, tsText.c_str(), tsTitle.c_str(), MB_ICONQUESTION | MB_YESNO) == IDNO)
			bDoIt = FALSE;
		else
			dScale = this->dMaxScale / 100.0;
	}

	if (bDoIt)
	{
		//Wenn skaliert wurde, setze die entsprechende Skalierungscheckbox
		if (this->Scale(dScale))
			this->CheckScaleMenuItem(IDC_MENU_ITEM_SCALE_TO_PAGE_WIDTH);
	}
}

/*
Skaliert das Panel, welches die Metafile enthält, auf die Größe des Dialog-Contents.
Die Seite wird als Vollbild mittig im Dialog-Content angezeigt.
*/
void PRINTPREVIEW::ScaleContentToSize()
{
	//Bestimme Größe des Panels, das die Metafile beinhaltet
	RECT panelRect;
	panelRect.left = 0;
	panelRect.top = 0;
	panelRect.right = round(printMetrics.iPhysWidth / (double)(printMetrics.iPixelPerInchX) * tableInfo.iPixelPerInchX) + 2 * iImgOffset;
	panelRect.bottom = round(printMetrics.iPhysHeight / (double)(printMetrics.iPixelPerInchY) * tableInfo.iPixelPerInchY) + 2 * iImgOffset;

	int iNewPanelHeight, iNewPanelWidth, iNewPanelPosX, iNewPanelPosY;
	double dXScaleFactor, dYScaleFactor;
	RECT contentRect;
	GetClientRect(this->hWndDialogContent, &contentRect);

	//Horizontaler Skalierungsfaktor des Panels zu dem Dialog-Content
	dXScaleFactor = (contentRect.right - contentRect.left) / (double)(panelRect.right - panelRect.left);
	//Vertikaler Skalierungsfaktor des Panels zu dem Dialog-Content
	dYScaleFactor = (contentRect.bottom - contentRect.top) / (double)(panelRect.bottom - panelRect.top);

	//Welche Skalierung muss verwendet werden?
	if (dXScaleFactor < dYScaleFactor)
	{
		iNewPanelHeight = (panelRect.bottom - panelRect.top) * dXScaleFactor;
		iNewPanelWidth = (panelRect.right - panelRect.left) * dXScaleFactor;
	}
	else
	{
		iNewPanelHeight = (panelRect.bottom - panelRect.top) * dYScaleFactor;
		iNewPanelWidth = (panelRect.right - panelRect.left) * dYScaleFactor;
	}

	//Platziere Panel mittig
	iNewPanelPosX = (contentRect.right - contentRect.left) / 2 - (iNewPanelWidth / 2);
	iNewPanelPosY = 0;

	//Neue Position und Größe der Fenster setzen
	SetWindowPos(this->hWndPanel, 0, iNewPanelPosX, iNewPanelPosY, iNewPanelWidth, iNewPanelHeight, SWP_NOZORDER);
	SetWindowPos(this->hWndMetafileView, 0, iImgOffset, iImgOffset, iNewPanelWidth - 2 * iImgOffset, iNewPanelHeight - 2 * iImgOffset, SWP_NOZORDER);
}

/*
Checkt eine der Checkboxen aus dem Skalierungsmenü. Alle anderen Checkboxen werden unchecked.
Parameter:
uiId:		ID der Checkbox, die gecheckt werden soll
*/
void PRINTPREVIEW::CheckScaleMenuItem(UINT uiId)
{
	if (!this->hMenuScale)
		return;

	MENUITEMINFO mii;
	int iItem;
	int iItemCount = GetMenuItemCount(this->hMenuScale);
	UINT uiCheck;

	// MENUITEMINO zum Ermitteln des Typs
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_FTYPE | MIIM_ID;

	// Menüeinträge durchlaufen...
	for (iItem = 0; iItem < iItemCount; iItem++)
	{
		// Info ermitteln
		if (GetMenuItemInfo(this->hMenuScale, (UINT)iItem, TRUE, &mii))
		{
			// Typ Radiocheck?
			if (mii.fType & MFT_RADIOCHECK)
			{
				// Wenn es sich um den übergebenen Eintrag handelt, dann MF_CHECKED setzen, ansonsten MF_UNCHECKED setzen 
				if (mii.wID == uiId)
					uiCheck = MF_BYCOMMAND | MF_CHECKED;
				else
					uiCheck = MF_BYCOMMAND | MF_UNCHECKED;
				CheckMenuItem(this->hMenuScale, mii.wID, uiCheck);
			}
		}
	}
}

/*
Druckt die aktuell angezeigte Seite der Druckvorschau.
*/
void PRINTPREVIEW::PrintCurrentPage()
{
	DOCINFO docInfo;
	BOOL bStatus;
	LPDEVMODE pDevMode = NULL;
	HANDLE hPrinter;
	int iPrintJob;
	long lLastTotal, lLastRow;

	//Orientierung setzen
	//Handle für Drucker ermitteln
	bStatus = OpenPrinter((LPTSTR)this->printOptions.tsPrinterName.c_str(), &hPrinter, NULL);
	if (bStatus)
	{
		DWORD dwNeeded = DocumentProperties(NULL, hPrinter, (LPTSTR)this->printOptions.tsPrinterName.c_str(), NULL, NULL, 0);
		pDevMode = (LPDEVMODE)malloc(dwNeeded);

		//Device Mode vom Drucker abfragen
		DWORD dwRet = DocumentProperties(NULL, hPrinter, (LPTSTR)this->printOptions.tsPrinterName.c_str(), pDevMode, NULL, DM_OUT_BUFFER);
		if (dwRet != IDOK)
		{
			free(pDevMode);
			ClosePrinter(hPrinter);
			return;
		}

		if (pDevMode->dmFields & DM_ORIENTATION)
			pDevMode->dmOrientation = this->printOptions.orientation;

		//Device Mode aktualisieren
		dwRet = DocumentProperties(NULL, hPrinter, (LPTSTR)this->printOptions.tsPrinterName.c_str(), pDevMode, pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);

		ClosePrinter(hPrinter);
	}

	//Informationen über das Dokument setzen
	memset(&docInfo, 0, sizeof(DOCINFO));
	//Dokumentname setzen
	if (this->printOptions.tsDocName.empty())
		docInfo.lpszDocName = (LPTSTR)_T("MTBLPRINT");
	else
		docInfo.lpszDocName = this->printOptions.tsDocName.c_str();

	//Druck starten
	iPrintJob = StartDoc(this->hdcPrinter, &docInfo);

	if (iPrintJob > 0)
	{
		bStatus = StartPage(hdcPrinter);

		//Mapping-Modus setzen
		SetMapMode(this->hdcPrinter, MM_ANISOTROPIC);
		//1:1 Ausgabe
		SetWindowExtEx(this->hdcPrinter, 1, 1, NULL);
		SetViewportExtEx(this->hdcPrinter, 1, 1, NULL);

		//Parameter zum Drucken der Seite setzen
		this->printPageParams.hdcCanvas = this->hdcPrinter;
		this->printPageParams.iPageNr = this->iCurrentPage;
		this->printPageParams.lStartRow = this->vecPageStartRows.at(this->iCurrentPage - 1);
		this->printPageParams.printOptions = this->printOptions;
		this->printPageParams.printMetrics = this->printMetrics;
		this->printPageParams.tableInfo = this->tableInfo;
		this->printPageParams.vecColsInfo = this->vecColsInfo;
		this->printPageParams.iFromCol = this->vecPageFromCols.at(this->iCurrentPage - 1);
		this->printPageParams.iToCol = this->vecPageToCols.at(this->iCurrentPage - 1);
		this->printPageParams.bColHeaders = this->printOptions.bColHeaders & (this->vecPageStartRows.at(this->iCurrentPage - 1) != -3);
		this->printPageParams.bCells = this->vecPageStartRows.at(this->iCurrentPage - 1) != -3;

		//Seite drucken
		lLastTotal = this->vecPageLastTotals.at(this->iCurrentPage - 1);
		this->lpSubclassRef->PrintPage(this->printPageParams, lLastRow, lLastTotal);

		//Druck beenden
		bStatus = EndPage(this->hdcPrinter);
		bStatus = EndDoc(this->hdcPrinter);
	}
}

/*
Aktiviert/Deaktiviert die Buttons der Toolbar und aktualisiert die Statusbar.
*/
void PRINTPREVIEW::PageChanged()
{
	this->EnableButtons();
	this->SetStatusbar();
}

/*
Aktiviert bzw. deaktiviert die Buttons der Toolbar abhängig davon, welche Seite aktuell angezeigt wird.
*/
void PRINTPREVIEW::EnableButtons()
{
	TBBUTTONINFO buttonInfo;
	buttonInfo.cbSize = sizeof(TBBUTTONINFO);
	buttonInfo.dwMask = TBIF_STATE;

	buttonInfo.idCommand = IDC_MENU_BTN_FIRSTPAGE;
	buttonInfo.fsState = (this->iCurrentPage != 1) ? TBSTATE_ENABLED : TBSTATE_INDETERMINATE;
	SendMessage(this->hWndToolbar, TB_SETBUTTONINFO, IDC_MENU_BTN_FIRSTPAGE, LPARAM(&buttonInfo));

	buttonInfo.idCommand = IDC_MENU_BTN_PREVPAGE;
	buttonInfo.fsState = (this->iCurrentPage > 1) ? TBSTATE_ENABLED : TBSTATE_INDETERMINATE;
	SendMessage(this->hWndToolbar, TB_SETBUTTONINFO, IDC_MENU_BTN_PREVPAGE, LPARAM(&buttonInfo));

	buttonInfo.idCommand = IDC_MENU_BTN_NEXTPAGE;
	buttonInfo.fsState = (this->iCurrentPage < this->iLastPage) ? TBSTATE_ENABLED : TBSTATE_INDETERMINATE;
	SendMessage(this->hWndToolbar, TB_SETBUTTONINFO, IDC_MENU_BTN_NEXTPAGE, LPARAM(&buttonInfo));

	buttonInfo.idCommand = IDC_MENU_BTN_LASTPAGE;
	buttonInfo.fsState = (this->iCurrentPage != this->iLastPage) ? TBSTATE_ENABLED : TBSTATE_INDETERMINATE;
	SendMessage(this->hWndToolbar, TB_SETBUTTONINFO, IDC_MENU_BTN_LASTPAGE, LPARAM(&buttonInfo));
}

/*
Setzt die Statusbar.
In der Statusbar wird angezeigt, welche Seite aktuell angezeigt wird, wieviele Seiten es insgesamt gibt und welche Skalierung aktuell verwendet wird.
*/
void PRINTPREVIEW::SetStatusbar()
{
	tstring tString, tsLastPage;
	//Aktuelle Seite
	if (mtblPrint_PageCount == -1)
	{
		if (this->iLastPage != PRINTPREVIEW_UNKNOWN)
			tsLastPage = to_tstring(this->iLastPage);
		else
			tsLastPage = _T("?");
	}

	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.StatBar.CurrPage"), tString);
	this->ReplaceSubstring(tString, tstring(_T("%d")), to_tstring(this->iCurrentPage));
	this->ReplaceSubstring(tString, tstring(_T("%s")), tsLastPage);
	SendMessage(this->hWndStatusbar, SB_SETTEXT, 0, (LPARAM)tString.c_str());

	//Aktuelle Skalierung
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("PrevWnd.StatBar.CurrScale"), tString);
	this->ReplaceSubstring(tString, tstring(_T("%s")), to_tstring((int)(this->printMetrics.dScale * 100)));
	this->ReplaceSubstring(tString, tstring(_T("%%")), tstring(_T("%")));
	SendMessage(this->hWndStatusbar, SB_SETTEXT, 1, (LPARAM)tString.c_str());
}

/*
Aktualisiert die horizontale und vertikale Scrollbar des Druckvorschaufensters.
Parameter:
iWndWidth:		Breite des Fensters, auf das sich die Scrollbar bezieht
iWndHeight:		Höhe des Fensters, auf das sich die Scrollbar bezieht
*/
void PRINTPREVIEW::UpdateScrollbar(int iWndWidth, int iWndHeight)
{
	SCROLLINFO scrollInfo;
	RECT rect;
	GetClientRect(this->hWndPanel, &rect);

	this->xMaxScroll = rect.right;
	this->xCurrentScroll = min(this->xCurrentScroll, this->xMaxScroll);

	//Informationen zur horizontalen Scrollbar setzen
	scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
	scrollInfo.nMin = this->xMinScroll;
	scrollInfo.nMax = this->xMaxScroll;
	scrollInfo.nPage = iWndWidth;
	scrollInfo.nPos = this->xCurrentScroll;
	SetScrollInfo(hWndDialogContent, SB_HORZ, &scrollInfo, TRUE);

	/*
	Scrollinfo nPos wird beim Setzen der Information eventuell geändert. Deswegen wird das MetafilePanel hier aktualisiert
	und die aktuelle Scoll Position korrigiert.
	*/
	GetScrollInfo(hWndDialogContent, SB_HORZ, &scrollInfo);
	ScrollWindowEx(hWndDialogContent, (this->xCurrentScroll - scrollInfo.nPos), 0, NULL, NULL, NULL, NULL, SW_SCROLLCHILDREN | SW_INVALIDATE | SW_ERASE);
	this->xCurrentScroll = scrollInfo.nPos;

	/*
	Nutze für die Seitenhöhe ein aktualisiertes Client Rect, da die horizontale Scrollbar eventuell verschwunden ist.
	Dadurch stimmt die Seitenhöhe nicht mehr mehr mit dem in lParam übergebenen Wert überein
	*/
	RECT updatedClientRect;
	GetClientRect(hWndDialogContent, &updatedClientRect);

	this->yMaxScroll = rect.bottom;
	this->yCurrentScroll = min(this->yCurrentScroll, this->yMaxScroll);

	//Informationen zur vertikalen Scrollbar setzen
	scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
	scrollInfo.nMin = this->yMinScroll;
	scrollInfo.nMax = this->yMaxScroll;
	scrollInfo.nPage = updatedClientRect.bottom;	//Nicht yNewSize, da sich die Höhe eventuell verändert hat
	scrollInfo.nPos = this->yCurrentScroll;

	SetScrollInfo(hWndDialogContent, SB_VERT, &scrollInfo, TRUE);

	/*
	Scrollinfo nPos wird beim Setzen der Information eventuell geändert. Deswegen wird das MetafilePanel hier aktualisiert
	und die aktuelle Scoll Position korrigiert.
	*/
	GetScrollInfo(hWndDialogContent, SB_VERT, &scrollInfo);
	ScrollWindowEx(hWndDialogContent, 0, (this->yCurrentScroll - scrollInfo.nPos), NULL, NULL, NULL, NULL, SW_SCROLLCHILDREN | SW_INVALIDATE | SW_ERASE);
	this->yCurrentScroll = scrollInfo.nPos;
}

/*
Löscht alle vorhandenen Metafiles.
*/
void PRINTPREVIEW::FreeMetafiles()
{
	int iSize;

	iSize = this->vecPageMetafiles.size();
	for (int i = 0; i < iSize; i++)
	{
		DeleteEnhMetaFile(this->vecPageMetafiles.at(i).hMetafile);
	}
	this->vecPageMetafiles.clear();
}

/*
Ersetzt in einem String das erste Vorkommnis eines Teilstring mit einem anderen String.
Parameter:
tString:		String, in dem ein Teilstring ersetzt werden soll
tsToReplace:	String, der in tString gefunden und ersetzt werden soll
tsReplaceWith:	String, der den Teilstring tsToReplace ersetzen soll
*/
void PRINTPREVIEW::ReplaceSubstring(tstring &tString, tstring tsToReplace, tstring tsReplaceWith)
{
	//Position des zu ersetzenden Teilstrings ermitteln
	tstring::size_type pos = tString.find(tsToReplace);

	//Teilstring gefunden?
	if (pos != tstring::npos)
	{
		//Teilstring ersetzen
		tString.replace(pos, tsToReplace.length(), tsReplaceWith);
	}
}


//***************** Preview Part End *****************


//******************* Status Part ********************

/*
Fensterprozedur des Statusfenster.
Das Statusfesnter zeigt dem Nutzer den aktuellen Status des Druckvorgangs an.
Das Statusfenster wird geöffnet, wenn den Druckvorgang gestartet wird.
Parameter:
hWndStatusForm:	Handle für das Fenster
uMsg:		Empfangene Nachricht
wParam:		Daten aus der Nachricht
lParam:		Daten aus der Nachricht
*/
INT_PTR CALLBACK StatusFormProc(HWND hWndStatusForm, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	//Statusform Objekt laden
	LPPRINTSTATUSFORM lpStatusForm;
	lpStatusForm = (LPPRINTSTATUSFORM)GetProp(hWndStatusForm, _T("STATUSFORM_OBJECT"));

	switch (uMsg)
	{
	case WM_INITDIALOG:
		lpStatusForm = (LPPRINTSTATUSFORM)lParam;
		lpStatusForm->OnStatusFormInit(hWndStatusForm);
		return TRUE;

	case WM_CLOSE:
		DestroyWindow(hWndStatusForm);
		ExitThread(0);
		return TRUE;

	case WM_CHAR:
		if (wParam == VK_ESCAPE)
			lpStatusForm->OnCancelClick();
		return TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_STATUSFORM_BTN_CANCEL:
			lpStatusForm->OnCancelClick();
			return TRUE;
		}
	}
	return FALSE;
}

/*
Statusfesnter Konstruktor
Initialisiert Eigenschaften und ermittelt die Texte des Statusfensters
Parameter:
printOptions:	Verwendete Druckoptionen
lpSubclassRef:	Referenz auf die aufgerufende Subclass
lpBoolCancel:	Pointer auf ein BOOL, der TRUE gesetzt werden soll, wenn das Statusfenster abgebrochen wird
*/
PRINTSTATUSFORM::PRINTSTATUSFORM(PRINTOPTIONS printOptions, LPMTBLSUBCLASS lpSubclassRef, LPBOOL lpBoolCancel)
{
	this->lpSubclassRef = lpSubclassRef;
	this->printOptions = printOptions;
	this->lpBoolCancel = lpBoolCancel;

	//Ermittle Titel und Buttontext
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("StatDlg.Title"), this->tsTitle);
	this->lpSubclassRef->GetUIStr(this->printOptions, _T("StatDlg.BtnCancel"), this->tsButtonText);
}

/*
Statusfenster Destruktor
Sendet eine Nachricht an das Statusfenster, dass dieses geschlossen werden soll.
Schließt anschließend das Handle des Threads.
*/
PRINTSTATUSFORM::~PRINTSTATUSFORM()
{
	PostMessage(this->hWndStatusForm, WM_CLOSE, NULL, NULL);
	CloseHandle(this->hStatusThread);
}

/*
Erstellt das Statusfenster und zeigt es an.
*/
void PRINTSTATUSFORM::CreateStatusForm()
{
	this->hWndStatusForm = CreateDialogParam((HINSTANCE)ghInstance, MAKEINTRESOURCE(IDD_DIALOG4), NULL, StatusFormProc, LPARAM(this));

	ShowWindow(this->hWndStatusForm, SW_SHOW);
}

/*
Initialisiert das Statusfenster.
Button und Label werden den im Dialog definierten Controls zugewiesen und Texte werden gesetzt.
Parameter:
hWndStatusForm:		Window Handle für das Fenster, das initialisiert werden soll
*/
void PRINTSTATUSFORM::OnStatusFormInit(HWND hWndStatusForm)
{
	this->hWndStatusForm = hWndStatusForm;
	SetProp(hWndStatusForm, _T("STATUSFORM_OBJECT"), (HANDLE)this);
	this->hWndLabel = GetDlgItem(hWndStatusForm, IDC_STATUSFORM_TEXT);
	this->hWndBtnCancel = GetDlgItem(hWndStatusForm, IDC_STATUSFORM_BTN_CANCEL);

	//SetDlgItemText(this->hWndStatusForm, IDC_STATUSFORM_TEXT, _T(this->tsTitle.c_str()));
	SetWindowText(this->hWndStatusForm, this->tsTitle.c_str());
	SetDlgItemText(this->hWndStatusForm, IDC_STATUSFORM_BTN_CANCEL, this->tsButtonText.c_str());
}

/*
Setzt die Abbruchvariable auf TRUE.
Dies signalisiert, dass das Statusfenster abgebrochen wurde und die Druckausgabe abgebrochen werden soll.
*/
void PRINTSTATUSFORM::OnCancelClick()
{
	*(this->lpBoolCancel) = TRUE;
}

//***************** Status Part End *****************

/*
Ermittelt aus der User Defined Variable für einen positionierten Text die entsprechenden Variablen.
Parameter:
hUdv:	User Defined Variable, die umgewandelt werden soll
Return:
TRUE, wenn die Ermittlung erfolgreich war.
FALSE, wenn die Dereferenzierung fehlgeschlagen ist.
*/
BOOL PRINTLINEPOSTEXT::FromUdv(HUDV hUdv)
{
	// UDV dereferenzieren
	LPPRINTLINEPOSTEXTUDV pUdv = (LPPRINTLINEPOSTEXTUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;

	// Werte setzen
	Ctd.HStrToTStr(pUdv->hsText, this->tsText);
	this->dLeft = max(Ctd.NumToDouble(pUdv->nLeft), 0.0);
	this->dWidth = max(Ctd.NumToDouble(pUdv->nWidth), 0.0);
	this->justify = static_cast<Justify>(Ctd.NumToInt(pUdv->nJustify));

	// Fertig
	return TRUE;
}

/*
Ermittelt aus der User Defined Variable für eine Zeile die entsprechenden Variablen.
Parameter:
hUdv:	User Defined Variable, die umgewandelt werden soll
Return:
TRUE, wenn die Ermittlung erfolgreich war.
FALSE, wenn die Dereferenzierung fehlgeschlagen ist.
*/
BOOL PRINTLINE::FromUdv(HUDV hUdv)
{
	// UDV dereferenzieren
	LPPRINTLINEUDV pUdv = (LPPRINTLINEUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;

	// Werte setzen
	Ctd.HStrToTStr(pUdv->hsLeftText, this->tsLeftText);
	Ctd.HStrToTStr(pUdv->hsCenterText, this->tsCenterText);
	Ctd.HStrToTStr(pUdv->hsRightText, this->tsRightText);
	this->iPosTextCount = Ctd.NumToInt(pUdv->nPosTextCount);

	double dFontSizeMulti = Ctd.NumToDouble(pUdv->nFontSizeMulti);
	if (dFontSizeMulti <= 0)
		this->dFontSizeMulti = 1.0;
	else
		this->dFontSizeMulti = dFontSizeMulti;

	this->lFontEnh = Ctd.NumToLong(pUdv->nFontEnh);
	this->crFontColor = Ctd.NumToULong(pUdv->nFontColor);
#ifdef _WIN64
	this->hLeftImage = (HANDLE)Ctd.NumToULongLong(pUdv->nLeftImage);
	this->hCenterImage = (HANDLE)Ctd.NumToULongLong(pUdv->nCenterImage);
	this->hRightImage = (HANDLE)Ctd.NumToULongLong(pUdv->nRightImage);
#else
	this->hLeftImage = (HANDLE)Ctd.NumToULong(pUdv->nLeftImage);
	this->hCenterImage = (HANDLE)Ctd.NumToULong(pUdv->nCenterImage);
	this->hRightImage = (HANDLE)Ctd.NumToULong(pUdv->nRightImage);
#endif
	this->dwFlags = Ctd.NumToULong(pUdv->nFlags);

	//PosTextArray
	if (this->iPosTextCount > 0)
	{
		for (int i = 0; i < this->iPosTextCount; i++){
			HUDV hudvPosText;
			Ctd.MDArrayGetUdv(pUdv->haPosTexts, &hudvPosText, i);
			PRINTLINEPOSTEXT printLinePosText;
			printLinePosText.FromUdv(hudvPosText);
			this->vecplPosText.push_back(printLinePosText);
		}
	}
	else
		this->iPosTextCount = 0;

	return TRUE;
}

/*
Ermittelt aus der User Defined Variable für Druckoptionen die entsprechenden Variablen.
Parameter:
hUdv:	User Defined Variable, die umgewandelt werden soll
Return:
TRUE, wenn die Ermittlung erfolgreich war.
FALSE, wenn die Dereferenzierung fehlgeschlagen ist.
*/
BOOL PRINTOPTIONS::FromUdv(HUDV hUdv){

	// UDV dereferenzieren
	LPPRINTOPTIONSUDV pUdv = (LPPRINTOPTIONSUDV)Ctd.UdvDeref(hUdv);
	if (!pUdv)
		return FALSE;
	
	// Werte setzen
	Ctd.HStrToTStr(pUdv->hsPrinterName, this->tsPrinterName);
	Ctd.HStrToTStr(pUdv->hsPrevWndTitle, this->tsPrevWndTitle);
	Ctd.HStrToTStr(pUdv->hsDocName, this->tsDocName);
	this->iTitleCount = Ctd.NumToInt(pUdv->nTitleCount);
	this->iTotalCount = Ctd.NumToInt(pUdv->nTotalCount);
	this->iPageHeaderCount = Ctd.NumToInt(pUdv->nPageHeaderCount);
	this->iPageFooterCount = Ctd.NumToInt(pUdv->nPageFooterCount);
	this->bPageNrs = Ctd.NumToBool(pUdv->nPageNrs);
	this->bNoPageNrsScale = Ctd.NumToBool(pUdv->nNoPageNrsScale);
	this->iCopies = max(Ctd.NumToInt(pUdv->nCopies), 1);
	this->dScale = max(min(Ctd.NumToDouble(pUdv->nScale), 5.0), 0.0);
	this->orientation = static_cast<MTblPrintOrientation>(Ctd.NumToInt(pUdv->nOrientation) + 1);
	this->gridType = static_cast<GridType>(Ctd.NumToInt(pUdv->nGridType));
	this->iLeftMargin = max(Ctd.NumToInt(pUdv->nLeftMargin), 0.0);
	this->iTopMargin = max(Ctd.NumToInt(pUdv->nTopMargin), 0.0);
	this->iRightMargin = max(Ctd.NumToInt(pUdv->nRightMargin), 0.0);
	this->iBottomMargin = max(Ctd.NumToInt(pUdv->nBottomMargin), 0.0);
	Ctd.CvtNumberToWord(&(pUdv->nRowFlagsOn), &(this->dwRowFlagsOn));
	Ctd.CvtNumberToWord(&(pUdv->nRowFlagsOff), &(this->dwRowFlagsOff));
	Ctd.CvtNumberToWord(&(pUdv->nColFlagsOn), &(this->dwColFlagsOn));
	Ctd.CvtNumberToWord(&(pUdv->nColFlagsOff), &(this->dwColFlagsOff));
	this->iPrevWndState = Ctd.NumToInt(pUdv->nPrevWndState);
	this->iPrevWndLeft = Ctd.NumToInt(pUdv->nPrevWndLeft);
	this->iPrevWndTop = Ctd.NumToInt(pUdv->nPrevWndTop);

	int iScreenWidth, iScreenHeight;
	GetDesktopResolution(iScreenWidth, iScreenHeight);
	this->iPrevWndWidth = min(max(Ctd.NumToInt(pUdv->nPrevWndWidth), 300), iScreenWidth);
	this->iPrevWndHeight = min(max(Ctd.NumToInt(pUdv->nPrevWndHeight), 300), iScreenHeight);

	this->bPrevWndFullPage = Ctd.NumToBool(pUdv->nPrevWndFullPage);
	this->bFitCellSize = Ctd.NumToBool(pUdv->nFitCellSize);
	this->bColHeaders = Ctd.NumToBool(pUdv->nColHeaders);
	this->iLanguage = Ctd.NumToInt(pUdv->nLanguage);
	this->bSpreadCols = Ctd.NumToBool(pUdv->nSpreadCols);
	this->bPreview = Ctd.NumToBool(pUdv->nPreview);

	//Titles
	if (this->iTitleCount > 0)
	{
		for (int i = 0; i < this->iTitleCount; i++)
		{
			HUDV hudvTitle;
			Ctd.MDArrayGetUdv(pUdv->haTitles, &hudvTitle, i);
			PRINTLINE printLineTitle;
			printLineTitle.FromUdv(hudvTitle);
			this->vecTitles.push_back(printLineTitle);
		}
	}
	else
		this->iTitleCount = 0;

	//Totals
	if (this->iTotalCount > 0)
	{
		for (int i = 0; i < this->iTotalCount; i++)
		{
			HUDV hudvTotal;
			Ctd.MDArrayGetUdv(pUdv->haTotals, &hudvTotal, i);
			PRINTLINE printLineTotals;
			printLineTotals.FromUdv(hudvTotal);
			this->vecTotals.push_back(printLineTotals);
		}
	}
	else
		this->iTotalCount = 0;

	//Page Headers
	if (this->iPageHeaderCount > 0)
	{
		for (int i = 0; i < this->iPageHeaderCount; i++)
		{
			HUDV hudvPageHeader;
			Ctd.MDArrayGetUdv(pUdv->haPageHeaders, &hudvPageHeader, i);
			PRINTLINE printLinePageHeader;
			printLinePageHeader.FromUdv(hudvPageHeader);
			this->vecPageHeaders.push_back(printLinePageHeader);
		}
	}
	else
		this->iPageHeaderCount = 0;

	//Page Footers
	if (this->iPageFooterCount > 0)
	{
		for (int i = 0; i < this->iPageFooterCount; i++)
		{
			HUDV hudvPageFooter;
			Ctd.MDArrayGetUdv(pUdv->haPageFooters, &hudvPageFooter, i);
			PRINTLINE printLinePageFooter;
			printLinePageFooter.FromUdv(hudvPageFooter);
			this->vecPageFooters.push_back(printLinePageFooter);
		}
	}
	else
		this->iPageFooterCount = 0;

	return TRUE;
}

/*
Ermittelt die Breite und Höhe des Bildschirms in Pixel.
Parameter:
horizontal:	Wird mit der Breite des Bildschirms in Pixel gefüllt
vertical: Wird mit der Höhe des Bildschirms in Pixel gefüllt
*/
void GetDesktopResolution(int &horizontal, int &vertical)
{
	RECT desktop;
	const HWND hDesktop = GetDesktopWindow();
	GetWindowRect(hDesktop, &desktop);
	horizontal = desktop.right;
	vertical = desktop.bottom;
}

/*
Erstellt eine Kopie eines gespeicherten Device Modes für einen Drucker.
Parameter:
tsPrinter:	Name des Druckers, dessen Device Mode kopiert werden soll.
Return:
Pointer auf eine Device Mode Kopie, wenn ein Device Mode für den Drucker gespeichert ist.
NULL, wenn kein Device Mode für den gegebenen Drucker gespeichert ist.
*/
LPDEVMODE GetPrinterDevModeCopy(tstring tsPrinter)
{
	LPDEVMODE lpDevMode, lpDevModeCopy;
	DWORD dwSize;
	LPVOID lpvDest, lpvSource;
	LPDEVMODE lpResult = NULL;
	lpDevMode = GetPrinterDevMode(tsPrinter);
	if (lpDevMode != NULL)
	{
		dwSize = GlobalSize(lpDevMode);
		if (dwSize > 0)
		{
			lpDevModeCopy = (LPDEVMODE)malloc(dwSize);
			if (lpDevModeCopy != NULL)
			{
				lpvSource = GlobalLock(lpDevMode);
				lpvDest = GlobalLock(lpDevModeCopy);
				if (lpvSource != NULL && lpvDest != NULL)
				{
					CopyMemory(lpvDest, lpvSource, dwSize);
					lpResult = lpDevModeCopy;
				}
				GlobalUnlock(lpDevMode);
				GlobalUnlock(lpDevModeCopy);
			}
		}
	}
	return lpResult;
}

/*
Setzt einen neuen Device Mode für einen Drucker.
Parameter:
tsPrinter:	Name des Druckers, für den der Device Mode gespeichert werden soll
lpDevMode:	Device Mode, der gespeichert werden soll
Return:
TRUE.
*/
BOOL SetPrinterDevMode(tstring tsPrinter, LPDEVMODE lpDevMode)
{
	int iIndex = -1;
	//Gibt es für den Printer schon einen Device Mode?
	for (int i = 0; i < haMTblPrinterDevModes.size(); i++)
	{
		if (tsPrinter.compare(haMTblPrinterDevModes.at(i).tsPrinter) == 0)
		{
			GlobalFree(haMTblPrinterDevModes.at(i).lpDevMode);
			iIndex = i;
			break;
		}
	}
	if (iIndex < 0)
	{
		//Noch kein Eintrag für den übergebenen Drucker, neuen Eintrag erstellen
		MTBLPRINT_PRINTERDEVMODE printerDevMode;
		printerDevMode.tsPrinter = tsPrinter;

		haMTblPrinterDevModes.push_back(printerDevMode);
		iIndex++;
	}

	//Device Mode setzen
	haMTblPrinterDevModes.at(iIndex).lpDevMode = lpDevMode;

	return TRUE;
}

/*
Ermittelt den Device Mode eines Druckers, sofern für diesen bereits ein Device Mode erstellt wurde.
Parameter:
tsPrinter:	Name des Druckers, dessen Device Mode ermittelt werden soll
Return:
Pointer auf den Device Mode, wenn er gefunden wurde.
NULL, wenn kein Device Mode für den gegebenen Drucker gespeichert ist.
*/
LPDEVMODE GetPrinterDevMode(tstring tsPrinter)
{
	LPDEVMODE result = NULL;
	for (int i = 0; i < haMTblPrinterDevModes.size(); i++)
	{
		if (tsPrinter.compare(haMTblPrinterDevModes.at(i).tsPrinter) == 0)
		{
			//Drucker gefunden
			result = haMTblPrinterDevModes.at(i).lpDevMode;
			break;
		}
	}
	return result;
}

//Gibt Speicher für gespeicherte Device Modes frei
void MTBLSUBCLASS::FreeDevModes()
{
	int iSize = haMTblPrinterDevModes.size();

	for (int i = 0; i < iSize; i++)
		GlobalFree(haMTblPrinterDevModes.at(i).lpDevMode);

	haMTblPrinterDevModes.clear();
}

/*
Überprüft, ob es den übergebenen Drucker gibt. 
Parameter:
lpsPrinter:	Name des Druckers, für den geprüft werden soll, ob er auf dem Computer installiert ist.
			Wird ein leerer String übergeben, wird geprüft, ob überhaupt ein Drucker installiert ist.
Return:
TRUE, wenn der Drucker installiert ist.
FALSE, wenn der Drucker nicht installiert ist.
Return bei leerem Übergabeparameter:
TRUE, wenn ein Drucker installiert ist.
FALSE, wenn kein Drucker installiert ist.
*/
BOOL PrinterExists(tstring tsPrinter)
{
	BOOL result = FALSE;
	DWORD dwBufferSizeNeeded = 0;
	DWORD dwPrinterCount = 0;
	LPBYTE lpbPrinterStructs;
	PRINTER_INFO_4* ppi;

	//Benötigte Buffergröße ermitteln
	EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 4, NULL, 0, &dwBufferSizeNeeded, &dwPrinterCount);

	lpbPrinterStructs = new BYTE[dwBufferSizeNeeded];
	//Drucker Info Structs für alle verfügbaren Drucker ermitteln
	if (EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 4, lpbPrinterStructs, dwBufferSizeNeeded, &dwBufferSizeNeeded, &dwPrinterCount))
	{
		//Leerer String übergeben?
		if (tsPrinter.empty())
		{
			//Wurden Drucker gefunden?
			if (dwPrinterCount > 0)
				result = TRUE;
			else
				result = FALSE;
		}
		else
		{
			//Druckername wurde übergeben
			for (int i = 0; i < dwPrinterCount; i++)
			{
				ppi = (PRINTER_INFO_4*)(lpbPrinterStructs + (i * sizeof(PRINTER_INFO_4)));
				if (lstrcmp(tsPrinter.c_str(), ppi->pPrinterName) == 0)
				{
					result = TRUE;
					break;
				}
			}
		}
	}

	delete[] lpbPrinterStructs;
	return result;
}

//************************************************************