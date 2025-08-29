#ifndef _MTBLSUBCLASS_H
#define _MTBLSUBCLASS_H

#include <stdexcept>
#include "mtbl.h"
#include "mtblprint.h"

// MTBLSUBCLASS
typedef struct MTBLSUBCLASS
{
	HWND m_hWndTbl;							// Handle der Tabelle
	CCurCols * m_CurCols;					// Aktuelle Spalten
	CMTblRows * m_Rows;						// Zeilen
	CMTblCols * m_Cols;						// Spalten
	CMTblCol * m_Corner;					// Corner
	CMTblColHdrGrps * m_ColHdrGrps;			// Column-Header-Gruppen
	CMTblTip * m_DefTip;					// Default-Tip
	CMTblTip * m_ETTip;						// Tip Entire Text
	SubClassMap * m_scm;					// Subclass-Map
	SubstColorMap * m_escm;					// Export Ersatzfarben-Map
	WNDPROC m_wpOld;						// Alte Fensterprozedur
	BOOL m_bExtMsgs;						// Erweiterte Nachrichten?
	DWORD m_dwExtMsgOpt;					// Optionen erweiterte Nachrichten
	BOOL m_bMWheelScroll;					// Scrollen bei Mousewheel?
	BOOL m_bEditLeftJustify;				// EditLeftJustify?
	long m_lLockPaint;						// LockPaint-Zähler
	int m_iIndent;							// Einzug
	PLongList m_DelAdjRows;					// Mit TBL_Adjust gelöschte Zeilen
	PHwndList m_LastSortCols;				// Letzte Sortierspalten
	PLongList m_LastSortFlags;				// Letzte Sortierflags
	PRowFlagImgMap m_RowFlagImages;			// Row-Flag Images
	LPMTBLPAINTDATA m_ppd;					// Paint-Daten
	BOOL m_bNoFocusPaint;					// Keinen Focus zeichnen?
	BOOL m_bSettingFocus;					// Wird der Focus gerade gesetzt?
	HFONT m_hFont;							// Font der Tabelle
	DWORD m_dwFlags;						// Flags
	CMTblLineDef* m_pHdrLineDef;			// Definition Header-Linien
	CMTblLineDef* m_pRowsHorzLineDef;			// Definition Zeilenlinien
	CMTblLineDef* m_pColsVertLineDef;		// Definition vertikale Linien aller Spalten
	CMTblLineDef* m_pLastLockedColVertLineDef;	// Definition Linie letzte gelockte Spalte
	BOOL m_bPainting;						// Zeichnet gerade?
	BOOL m_bHLHover;						// Hyperlink HOVER aktiviert ja/nein
	long m_lHLHoverFontEnh;					// Zusätzliche Font-Stile für Hyperlink HOVER
	DWORD m_dwHLHoverTextColor;				// Textfarbe für Hyperlink HOVER
	TCHAR m_cPassword;						// Password-Zeichen
	DWORD m_dwEnabledTipTypes;				// Aktivierte Tip-Typen
	BOOL m_bCornerOwnerdrawn;				// Corner benutzerdefiniert gezeichnet ja/nein
	WORD m_wTDVer;							// TD-Version
	BOOL m_bFixColJustifyBug;				// Spaltenausrichtungs-Bug fixen ja/nein
	BOOL m_bTrappingRowSelChanges;			// Änderungen der Zeilenselektion werden gerade verfolgt
	BOOL m_bTrappingColSelChanges;			// Änderungen der Spaltenselektion werden gerade verfolgt
	DWORD m_dwSelBackColor;					// Hintergrundfarbe für Selektionen
	DWORD m_dwSelTextColor;					// Textfarbe für Selektionen
	UCHAR m_bySelDarkening;					// Abdunkelungsfaktor für Selektion
	DWORD m_dwHdrsBackColor;				// Hintergrundfarbe der Header( Row, Corner, Column )
	DWORD m_dwColHdrFreeBackColor;			// Hintergrundfarbe freier Bereich Column-Header
	DWORD m_dwRowHdrFreeBackColor;			// Hintergrundfarbe freier Bereich Row-Header
	CMTblCounter *m_pCounter;				// Zähler
	BOOL m_bTempSingleSelection;			// SingleSelection-Flag temporär gesetzt?
	int m_iMWRowsPerNotch;					// Zeilen pro Mausrad-Raster
	int m_iMWSplitRowsPerNotch;				// Split-Zeilen pro Mausrad-Raster
	int m_iMWPageKey;						// Key für das Mausrad-Scrollen einer ganzen Seite
	int m_iBtnKey;							// Button-Hotkey
	BOOL m_bBtnKeyShift;					// Muss <Shift> bei Button-Hotkey gedrückt sein?
	BOOL m_bBtnKeyCtrl;						// Muss <Control> bei Button-Hotkey gedrückt sein?
	BOOL m_bBtnKeyAlt;						// Muss <Alt> bei Button-Hotkey gedrückt sein?
	BOOL m_bMouseClickSim;					// Mausklick-Simulation?
	DWORD m_dwRowAlternateBackColor1;		// Abwechselnede Zeilen-Hintergrundfarbe 1
	DWORD m_dwRowAlternateBackColor2;		// Abwechselnde Zeilen-Hintergrundfarbe 2
	DWORD m_dwSplitRowAlternateBackColor1;	// Abwechselnde Split-Zeilen-Hintergrundfarbe 1
	DWORD m_dwSplitRowAlternateBackColor2;	// Abwechselnde Split-Zeilen-Hintergrundfarbe 2
	DWORD m_dwEffCellPropDetOrder;			// Ermittlungs-Reihenfolge effektive Zellen-Eigenschaften
	DWORD m_dwEffCellPropDetFlags;			// Flags effektive Zellen-Eigenschaften
	BYTE m_byEffCellPropCellPos;			// Zellen-Position in der Ermittlungs-Reihenfolge effektive Zellen-Eigenschaften
	BYTE m_byEffCellPropColPos;				// Spalten-Position in der Ermittlungs-Reihenfolge effektive Zellen-Eigenschaften
	BYTE m_byEffCellPropRowPos;				// Zeilen-Position in der Ermittlungs-Reihenfolge effektive Zellen-Eigenschaften
	BOOL m_bDefiningSplitWindow;			// SalTblDefineSplitWindow wird gerade durchgeführt
	RECT m_LastClientRectOnSize;			// Letztes Client-Rechteck bei WM_SIZE

	// Variablen erweiterte Nachrichten
	BOOL bLDownCornerSep, bLUpCornerSep, bRDownCornerSep, bRUpCornerSep, bLDblClkCornerSep, bRDblClkCornerSep;
	BOOL bLDownBeginRowSizing;
	double dLDownSepColWidth, dLUpSepColWidth;
	DWORD dwLDownFlags, dwLUpFlags, dwRDownFlags, dwRUpFlags, dwLDblClkFlags, dwRDblClkFlags;
	HWND hWndLDownCol, hWndLDownSepCol, hWndLUpCol, hWndLUpSepCol, hWndRDownCol, hWndRDownSepCol, hWndRUpCol, hWndRUpSepCol, hWndLDblClkCol, hWndLDblClkSepCol, hWndRDblClkCol, hWndRDblClkSepCol;
	int iLDownX, iLDownY, iLDownColPos, iLUpX, iLUpY, iLUpColPos, iRDownX, iRDownY, iRUpX, iRUpY, iLDblClkX, iLDblClkY, iRDblClkX, iRDblClkY;
	int iLDownLinesPerRow, iLUpLinesPerRow;
	long lLDownRowHeight, lLUpRowHeight;
	long lLDownRow, lLUpRow, lRDownRow, lRUpRow, lLDblClkRow, lRDblClkRow;
	long lLDownSepRow, lLUpSepRow, lRDownSepRow, lRUpSepRow, lLDblClkSepRow, lRDblClkSepRow;
	MTBLGETMETRICS mLDownMetrics, mLUpMetrics;
	UINT uLastLDownMsg, uLastRDownMsg;

	// Variablen Tree
	HWND m_hWndTreeCol;						// Tree-Spalte
	int m_iNodeSize;						// Node-Größe
	DWORD m_dwTreeFlags;					// Tree-Flags
	DWORD m_dwNodeColor;					// Knoten-Farbe außen
	DWORD m_dwNodeInnerColor;				// Knoten-Farbe innen
	DWORD m_dwNodeBackColor;				// Knoten-Farbe Hintergrund
	CMTblLineDef m_TreeLineDef;				// Definition Linien
	HANDLE m_hImgNode;						// Bild Node normal
	HANDLE m_hImgNodeExp;					// Bild Node expandiert
	long m_lLoadChildRows;					// Zähler Laden Kindzeilen
	RowPtrStack *m_RowsLoadingChilds;		// Stack der Zeilen, für die Kindzeilen geladen werden

	// Variablen List-Clip
	BOOL m_bAdaptingListClip;
	HWND m_hWndListClip;
	HRGN m_hRgnListClip;

	// Variablen Edit
	HWND m_hWndEdit;
	HWND m_hWndLastEdit;
	HBRUSH m_hEditBrush;
	HFONT m_hEditFont;
	COLORREF m_clrEditBack, m_clrEditText;
	BOOL m_bDeleteEditFont;
	int m_iEditTextLineHt;
	BOOL m_bForceColorEdit;
	BOOL m_bForceEditFont;

	// Variablen Check-Box
	HWND m_hWndCheckBox;

	// Variablen Button
	HWND m_hWndBtn;
	//int m_iBtnHt;
	//HFONT m_hBtnFont;
	HWND m_hWndColBtnPushed;
	long m_lRowBtnPushed;

	// Variablen Timer
	__int64 m_iiTimer;
	DWORD m_dwTimerFlags, m_dwTimerMFlags;
	HWND m_hWndTimer, m_hWndTimerCol;
	long m_lTimerRow;
	POINT m_ptTimerPos;
	UINT m_uiTimer;

	// Variablen Enter/Leave
	LPMTBLITEM m_pELItem;
	HWND m_hWndELCol, m_hWndELCol_Btn, m_hWndELCol_Cell, m_hWndELCol_DDBtn;
	CMTblRow *m_pELRow, *m_pELRow_Btn, *m_pELRow_Cell, *m_pELRow_DDBtn, *m_pELRow_Hdr;
	CMTblColHdrGrp *m_pELColHdrGrp;
	DWORD m_dwELFlags, m_dwELMFlags;

	// Variablen Highlighting
	HiLiDefMap *m_pHLDM;
	ItemList *m_pHLIL;

	// Variablen Hyperlink
	HWND m_hWndHLCol;
	long m_lHLRow;

	// Variablen Tip
	__int64 m_iiTipTime;
	BOOL m_bCornerTip, m_bTipShowing;
	CMTblColHdrGrp *m_pColHdrGrpTipObj;
	CMTblTip *m_pTip;
	DWORD m_dwTipDrawFlags, m_dwTipFlags;
	HRGN m_hRgnTip;
	HWND m_hWndTip, m_hWndBtnTipCol, m_hWndCellTipCol, m_hWndColHdrTipCol, m_hWndTipBtn;
	int m_iTipType;
	long m_lRowHdrTipRow, m_lBtnTipRow, m_lCellTipRow;

	// Variablen benutzerdef. Zeichnen
	LPMTBLODI m_pODI;
	LPMTBLODI m_pPrODI;

	// Variablen Hook
	static HHOOK m_hHookEditStyle;
	static int m_iHookEditJustify;
	static BOOL m_bHookEditMultiline;
	static void *m_pscHookEditStyle;

	// Variablen Zellen-Modus
	BOOL m_bCellMode;
	DWORD m_dwCellModeFlags;

	// Variablen Focus
	CMTblRow *m_pRowFocus;
	DWORD m_dwFocusFrameColor;
	HWND m_hWndColFocus;
	long m_lFocusFrameThick;
	long m_lFocusFrameOffsX;
	long m_lFocusFrameOffsY;
	RECT m_rFocus;

	// Variablen Zellen-Selektion
	BOOL m_bCellSelByMouse;
	BOOL m_bCellSelByMouseClick;
	MTBLCELLRANGE m_CellSelRange;
	MTBLCELLRANGE m_CellSelRangeA;
	MTBLCELLRANGE m_CellSelRangeKey;

	// Variablen Zeilen-Selektion
	BOOL m_bRowSelByMouseClick;				// Hat ein Klick stattgefunden, mit dem die Zeilenselektion mit der Maus beginnen könnte?
	int m_iRowSelByMouse;					// Werden gerade Zeilen mit der Maus markiert? 0=Nein, 1=Top-Down, 2=Bottom-Up
	MTBLROWRANGE m_RowSelRange;
	MTBLROWRANGE m_RowSelRangeA;
	POINT m_ptRowSelByMouseClick;			// Position bei "RowSelByMouse-Klick"

	// Variablen Merging
	LPMTBLMERGECELLS m_pMergeCells;
	LPMTBLMERGEROWS m_pMergeRows;
	BOOL m_bMergeCellsInvalid;
	BOOL m_bMergeRowsInvalid;
	BOOL m_bReplMergeRowsFlags;
	long m_lRowAddToMergeCells;
	HWND m_hWndColAddToMergeCells;

	// Variablen Spalten-Bewegung
	BOOL m_bColMoving;

	// Variablen Spalten-Größenänderung
	BOOL m_bColSizing;

	// Variablen Row-Header-Größenänderung
	BOOL m_bRowHdrSizing;

	// Variablen Zeilen-Größenänderung
	BOOL m_bRowSizing;
	long m_lRowSizing;
	long m_lRowSizingOffsY;
	int m_iRowSizingLinesPerRow;

	// Variablen Split
	BOOL m_bSplitted;
	BOOL m_bSplitBarDragging;

	// Variablen Scrollen
	BOOL m_bNoVScrolling;
	BOOL m_bScrolling;
	BOOL m_bSettingScrollRange;
	int m_iColOriginBeforeScroll;
	int m_iRowOriginBeforeScroll;
	int m_iVScrollMaxCtd;
	int m_iVScrollMaxSplitCtd;

	// Variablen Double-Buffer
	BOOL m_bDoubleBuffer;

	// Variablen KillFocusFrameI
	BOOL m_bKillFocusFrameI;

	// Variablen Indent
	long m_lIndentPerLevel;

	// Variablen KillEdit
	BOOL m_bKillEditStartedRowSelTrap;		// Verfolgung der Zeilenselektion wurde in BeforeKillEdit() gestartet
	/*BOOL m_bKillEditSuppressSelection;*/		// Zeilen/Zellenselktion nach KillEdit unterdrücken

	// Variablen DragDrop
	BOOL m_bDragCanAutoStart;
	HWND m_hWndDragDrop;
	LPARAM m_DragDropAutoStart_lParam;
	UINT m_uiDragDropTimer;
	WPARAM m_DragDropAutoStart_wParam;

	// Variablen SetFocusCell
	long m_lNoSetFocusCellProcessing;

	// Interne Ereignis-Handler
	void OnEnterItem();
	void OnLeaveItem();

	// Message/Hook-Handler
	LRESULT OnCaptureChanged(WPARAM wParam, LPARAM lParam);
	LRESULT OnChar(WPARAM wParam, LPARAM lParam);
	LRESULT OnClearSelection(WPARAM wParam, LPARAM lParam);
	LRESULT OnCommand(WPARAM wParam, LPARAM lParam);
	LRESULT OnCtlColorEdit(WPARAM wParam, LPARAM lParam);
	LRESULT OnDefineSplitWindow(WPARAM wParam, LPARAM lParam);
	LRESULT OnDeleteRow(WPARAM wParam, LPARAM lParam);
	LRESULT OnEditModeChange(WPARAM wParam, LPARAM lParam);
	LRESULT OnFetchRow(WPARAM wParam, LPARAM lParam);
	LRESULT OnGetDlgCode(WPARAM wParam, LPARAM lParam);
	LRESULT OnGetMetrics(WPARAM wParam, LPARAM lParam);
	LRESULT OnInsertRow(WPARAM wParam, LPARAM lParam);
	LRESULT OnInternal(WPARAM wParam, LPARAM lParam);
	LRESULT OnInternalTimer(WPARAM wParam, LPARAM lParam);
	LRESULT OnKeyDown(WPARAM wParam, LPARAM lParam);
	LRESULT OnKillEdit(WPARAM wParam, LPARAM lParam);
	LRESULT OnKillFocus(WPARAM wParam, LPARAM lParam);
	LRESULT OnKillFocusFrame(WPARAM wParam, LPARAM lParam);
	LRESULT OnLButtonDblClk(WPARAM wParam, LPARAM lParam);
	LRESULT OnLButtonDown(WPARAM wParam, LPARAM lParam);
	LRESULT OnLButtonUp(WPARAM wParam, LPARAM lParam);
	LRESULT OnMouseMove(WPARAM wParam, LPARAM lParam);
	LRESULT OnMouseWheel(WPARAM wParam, LPARAM lParam);
	LRESULT OnNCDestroy(WPARAM wParam, LPARAM lParam);
	LRESULT OnObjectsFromPoint(WPARAM wParam, LPARAM lParam);
	LRESULT OnParentNotify(WPARAM wParam, LPARAM lParam);
	LRESULT OnPaste(WPARAM wParam, LPARAM lParam);
	LRESULT OnPosChanged(WPARAM wParam, LPARAM lParam);
	LRESULT OnQueryVisRange(WPARAM wParam, LPARAM lParam);
	LRESULT OnRButtonDblClk(WPARAM wParam, LPARAM lParam);
	LRESULT OnRButtonDown(WPARAM wParam, LPARAM lParam);
	LRESULT OnRButtonUp(WPARAM wParam, LPARAM lParam);
	LRESULT OnReset(WPARAM wParam, LPARAM lParam);
	LRESULT OnRowSetContext(WPARAM wParam, LPARAM lParam);
	LRESULT OnScroll(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetContext(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetCursor(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetFocus(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetFocusCell(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetFocusRow(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetFont(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetLinesPerRow(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetLockedColumns(WPARAM wParam, LPARAM lParam);
	LRESULT OnSetRowFlags(WPARAM wParam, LPARAM lParam);
	BOOL OnSetScrollRange(HWND hWnd, int iBar, int iMinPos, int iMaxPos, BOOL bRedraw, int iScrPos = TBL_Error, BOOL bForce = FALSE);
	LRESULT OnSetTableFlags(WPARAM wParam, LPARAM lParam);
	LRESULT OnSize(WPARAM wParam, LPARAM lParam);
	LRESULT OnSwapRows(WPARAM wParam, LPARAM lParam);
	LRESULT OnSysKeyDown(WPARAM wParam, LPARAM lParam);
	LRESULT OnVScroll(WPARAM wParam, LPARAM lParam);

	// Fensterprozedur Spalte
	static LRESULT CALLBACK ColWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur externes Edit
	static LRESULT CALLBACK ExtEditWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur Container
	static LRESULT CALLBACK ContainerWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur ListClip
	static LRESULT CALLBACK ListClipWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Message-Handler Container
	LRESULT OnContainerParentNotify(HWND hWndTbl, HWND hWnd, WPARAM wParam, LPARAM lParam);

	// Message-Handler ListClip
	LRESULT OnListClipParentNotify(HWND hWndTbl, HWND hWnd, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur ListClip-Control
	static LRESULT CALLBACK ListClipCtrlWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur Button
	static LRESULT CALLBACK BtnWndProc(HWND hWndBtn, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur Leere Spalte
	static LRESULT CALLBACK EmptyColWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Fensterprozedur Timer
	static LRESULT CALLBACK TimerWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Prozedur Timer
	static void CALLBACK TimerProc(HWND hWnd, UINT uMsg, UINT uEvent, DWORD dwTime);

	// Fensterprozedur Tip
	static LRESULT CALLBACK TipWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	// Message-Handler Button
	static LRESULT OnBtnCaptureChanged(HWND hWndBtn, WPARAM wParam, LPARAM lParam);
	static LRESULT OnBtnDestroy(HWND hWndBtn, WPARAM wParam, LPARAM lParam);
	static LRESULT OnBtnLButtonDown(HWND hWndBtn, WPARAM wParam, LPARAM lParam);
	static LRESULT OnBtnLButtonUp(HWND hWndBtn, WPARAM wParam, LPARAM lParam);
	static LRESULT OnBtnMouseMove(HWND hWndBtn, WPARAM wParam, LPARAM lParam);
	static LRESULT OnBtnPaint(HWND hWndBtn, WPARAM wParam, LPARAM lParam);

	// Message-Handler Tip
	LRESULT OnTipPaint(HWND hWndTip, WPARAM wParam, LPARAM lParam);
	LRESULT OnTipPrintClient(HWND hWndTip, WPARAM wParam, LPARAM lParam);
	LRESULT OnTipShow(HWND hWndTip, WPARAM wParam, LPARAM lParam);

	// Suchprozedur Kindklasse
	static BOOL CALLBACK FindChildClassProc(HWND hWnd, LPARAM lParam);

	// Dotline-Prozedur
	static void CALLBACK DotlineProc(int iX, int iY, LPARAM lParam);

	// Hook-Prozedur Edit-Stil
	static LRESULT CALLBACK HookProcEditStyle(int iCode, WPARAM wParam, LPARAM lParam);

	// Funktionen
	BOOL AdaptCellRange(MTBLCELLRANGE &cr);
	BOOL AdaptColHdrText(HWND hWndCol);
	void AdaptListClip();
	void AdaptScrollRangeV(BOOL bSplitBar = FALSE, int iScrPos = TBL_Error);
	void AfterKillEdit();
	void AfterScroll();
	void AfterSetFocusCell(HWND hWndCol, long lRow, BOOL bMouseClickSim = FALSE);
	BOOL AnyCellSelected();
	BOOL AnyHighlightedItems();
	BOOL AnyMergeRowSelected(LPMTBLMERGEROW pmr);
	BOOL AreRowsOfSameArea(long lRow1, long lRow2);
	BOOL AttachCol(HWND hWndCol);
	BOOL AttachCols();
	BOOL AutoMergeCells(long lRow, LPMTBLMERGEROW pmrOld, DWORD dwFlags);
	void BeforeKillEdit();
	void BeforeScroll();
	int BeforeSetFocusCell(HWND hWndCol, long lRow, BOOL bSendValidateMsg = FALSE);
	void BeginColSelTrap(HWND hWndCol = NULL);
	void BeginRowSelTrap(long lRow = TBL_Error, RowPtrVector *pvRows = NULL);
	BOOL CalcBtnRect(CMTblBtn *pb, HWND hWndCol, long lRow, CMTblMetrics *pTblMetrics, LPRECT prBounds, HDC hDC, HFONT hFont, LPRECT prBtn);
	int CalcBtnWidth(CMTblBtn * pb, HWND hWndCol, long lRow, HFONT hFont = NULL);
	void CalcCellTextAreaXCoord(int iTextAreaLeft, int iTextAreaRight, int iCellLeadingX, int iJustify, BOOL bWordWrap, POINT &pt);
	BOOL CalcCellTextRect(long lCellTop, long lCellBottom, long lTextAreaLeft, long lTextAreaRight, long lCellLeadingX, long lCellLeadingY, int iJustify, int iVAlign, SIZE &siText, LPRECT prTxt);
	void CalcCheckBoxRect(int iJustify, int iVAlign, const RECT &rCbArea, long lCellLeadingX, long lCellLeadingY, long lCharBoxY, RECT &r);
	BOOL CalcDDBtnRect(HWND hWndCol, long lRow, CMTblMetrics *pTblMetrics, LPRECT prBounds, HDC hDC, HFONT hFont, LPRECT prBtn);
	long CalcMaxScrollRangeV(BOOL bSplitBar = FALSE, int iScrPos = TBL_Error);
	LRESULT CallOldWndProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
	BOOL CanSetFocusCell(HWND hWndCol, long lRow, BOOL bAllowMerged = TRUE);
	BOOL CanSetFocusToCell(HWND hWndCol, long lRow);
	BOOL CellSelRangeChanged(MTBLCELLRANGE &crOld, MTBLCELLRANGE &crNew, BOOL bInvalidate);
	BOOL ClearCellSelection(BOOL bInvalidate = TRUE);
	BOOL ClearColumnSelection();
	BOOL CollapseRow(CMTblRow * pRow, DWORD dwFlags);
	BOOL CombineHighlighting(int iItem, MTBLHILI &hili, int iItemAdd, MTBLHILI &hiliAdd);
	BOOL CopyCurrentSelection();
	BOOL CopyRows(WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff, BOOL bHeaders = FALSE);
	int CopyShortcutPressed();
	BOOL CopyToClipboard(tstring &s);
	void CreateTipWnd();
	COLORREF DarkenColor(COLORREF cr, UCHAR ucDarkening);
	int DeleteDescRows(CMTblRow *pRow, WORD wMethod);
	void DeleteRowFlagImages();
	BOOL DragDropAutoStart(WPARAM wParam, LPARAM lParam, HWND hWndDragDrop = NULL);
	BOOL DragDropAutoStartCancel();
	BOOL DragDropBegin();
	void DrawLinesPerRowIndicators(int iLinesPerRow);
	void DrawTip(HWND hWnd, HDC hDC);
	BOOL EnableTipType(DWORD dwType, BOOL bEnable);
	void EndColSelTrap(HWND hWndCol = NULL);
	void EndRowSelTrap(long lRow = TBL_Error, RowPtrVector *pvRows = NULL, BOOL bRemoveNew = FALSE);
	BOOL EnsureFocusCellReadOnly();
	BOOL ExpandRow(CMTblRow * pRow, DWORD dwFlags);
	int ExportRowToHTML(LPMTBLERTHPARAMS p);
	int ExportToHTML(LPCTSTR lpsFile, DWORD dwHTMLFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff);
	int FetchMissingRows();
	HWND FindFirstChildClass(HWND hWndParent, LPTSTR lpsClassName);
	ItemList::iterator FindHighlightedItem(CMTblItem &item);
	LPMTBLMERGECELL FindMergeCell(HWND hWndCol, long lRow, int iMode, LPMTBLMERGECELLS pMergeCells = NULL);
	long FindMergeCells(HWND hWndCol, long lRow, int iMode, MTBLMERGECELL_VECTOR &v, int iFindMax = 0, LPMTBLMERGECELLS pMergeCells = NULL);
	LPMTBLMERGEROW FindMergeRow(long lRow, int iMode, LPMTBLMERGEROWS pMergeRows = NULL);
	long FindMergeRows(long lRow, int iMode, MTBLMERGEROW_VECTOR &v, int iFindMax = 0, LPMTBLMERGEROWS pMergeRows = NULL);
	BOOL FindNextCell(LPLONG plRow, LPHWND phWndCol, int iDirection, BOOL bCheckFocusCell = FALSE);
	int FindNextCol(LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff);
	HWND FindNextListClip(HWND hWndChildAfter);
	BOOL FindNextRow(LPLONG plRow, WORD wFlagsOn, WORD wFlagsOff, long lMaxRow = TBL_MaxRow);
	LPMTBLPAINTCOL FindPaintCol(HWND hWndCol);
	LPMTBLPAINTMERGECELL FindPaintMergeCell(HWND hWndCol, long lRow, int iMode);
	LPMTBLPAINTMERGEROWHDR FindPaintMergeRowHdr(long lRow, int iMode);
	int FindPrevCol(LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff);
	BOOL FindPrevRow(LPLONG plRow, WORD wFlagsOn, WORD wFlagsOff, long lMinRow = TBL_MinRow);
	BOOL FixSplitWindow();
	void FreeMergeCells(LPMTBLMERGECELLS pMergeCells = NULL);
	void FreeMergeRows(LPMTBLMERGEROWS pMergeRows = NULL);
	void FreePaintMergeCells();
	void FreePaintMergeRowHdrs();
	void GenerateColSelChangeMsg();
	void GenerateExtMsgs(UINT uMsg);
	void GenerateRowSelChangeMsg();
	int GetBestCellHeight(HWND hWndCol, long lRow, long lCellLeadingY, long lCharBoxY, HFONT hUseFont = NULL, LPMTBLMERGECELLS pMergeCells = NULL, BOOL bCheckboxText = FALSE, BOOL bIgnoreButtons = FALSE);
	int GetBestCellWidth(HWND hWndCol, long lRow, long lCellLeadingX, HFONT hUseFont = NULL, LPMTBLMERGECELLS pMergeCells = NULL, BOOL bCheckboxText = FALSE);
	int GetBestColHdrWidth(HWND hWndCol, long lCellLeadingX, HFONT hUseFont = NULL);
	int GetBtnBestHeight(CMTblBtn * pb, long lRow, HWND hWndCol);
	HFONT GetBtnFont(CMTblBtn * pb, long lRow, HWND hWndCol, LPBOOL pbMustDelete);
	//static int GetBtnHeight( HWND hWndTbl );
	BOOL GetBtnRect(HWND hWndCol, long lRow, CMTblMetrics * ptm, LPRECT pRect);
	int GetCellAreaLeft(long lRow, HWND hWndCol, long lCellLeft);
	BOOL GetCellFocusRect(long lRow, HWND hWndCol, LPRECT pRect);
	BOOL GetCellFontMetrics(long lRow, HWND hWndCol, LPTEXTMETRIC pMetrics);
	long GetCellIndentLevel(long lRow, HWND hWndCol);
	int GetCellRect(HWND hWndCol, long lRow, CMTblMetrics * ptm, LPRECT pRect, DWORD dwFlags);
	BOOL GetCellTextExtent(HWND hWndCol, long lRow, LPSIZE pSize, HFONT hUseFont = NULL, BOOL bWordbreak = FALSE, BOOL bForceMultiline = FALSE, BOOL bEndEllipsis = FALSE, BOOL bIgnoreButtons = FALSE, LPCTSTR lpsUseText = NULL, long lTextAreaWid = -1, LPMTBLMERGECELLS pMergeCells = NULL);
	int GetCellTextAreaLeft(long lRow, HWND hWndCol, long lCellLeft, long lCellLeadingX, CMTblImg & Image, BOOL bButton = FALSE);
	int GetCellTextAreaRight(long lRow, HWND hWndCol, long lCellRight, long lCellLeadingX, CMTblImg & Image, BOOL bButton = FALSE, BOOL bDDButton = FALSE);
	BOOL GetCellTextAreaXCoord(long lRow, HWND hWndCol, CMTblMetrics *ptm, LPMTBLMERGECELLS pMergeCells, LPPOINT ppt, LPPOINT pptVis, BOOL bIgnoreButtons = FALSE);
	BOOL GetCellWordWrap(HWND hWndCol, long lRow, CMTblMetrics *ptm = NULL);
	BOOL GetColHdrFontMetrics(HWND hWndCol, LPTEXTMETRIC pMetrics);
	int GetColHdrGrpHeight(CMTblColHdrGrp *pGrp);
	BOOL GetColHdrGrpTextExtent(CMTblColHdrGrp *pGrp, LPSIZE pSize, HFONT hUseFont = NULL);
	int GetColHdrRect(HWND hWndCol, RECT &r, DWORD dwFlags);
	BOOL GetColHdrText(HWND hWndCol, tstring &sText, int iMaxLen = 2048);
	BOOL GetColHdrTextExtent(HWND hWndCol, LPSIZE pSize, HFONT hUseFont = NULL);
	int GetColJustify(HWND hWndCol);
	int GetColRect(HWND hWndCol, LPRECT pr);
	long GetColSel(int iMode, HWND hWndCol = NULL);
	long GetColSelChanges(HARRAY haCols, HARRAY haChanges);
	HRGN GetColSelChangesRgn(HWND hWndCol = NULL);
	int GetColXCoord(HWND hWndCol, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis);
	inline LPMTBLCOLSUBCLASS GetColSubClass(HWND hWndCol);
	int GetCRLFCount(LPCTSTR lpsStr);
	BOOL GetCSSColors(DWORD dwColor, DWORD dwBackColor, tstring &sCSS);
	BOOL GetCSSFont(CMTblFont &f, tstring &sCSS);
	BOOL GetColumnTitle(HWND hWndCol, LPHSTRING phsTitle);
	BOOL GetCornerRect(LPRECT pr);
	BOOL GetDDBtnRect(HWND hWndCol, long lRow, CMTblMetrics * ptm, LPRECT pRect);
	COLORREF GetDefaultColor(int iColorIndex);
	COLORREF GetDefaultColorNoTdTheme(int iColorIndex);
	long GetDescendantMergeRows(long lRow, vector<long> &v, LPMTBLMERGECELLS pMergeCells = NULL);
	COLORREF GetDrawColor(COLORREF clr);
	BOOL GetDropDownListContext(HWND hWndCol, LPLONG plRow);
	COLORREF GetEffCellBackColor(long lRow, HWND hWndCol, DWORD dwFlags = 0);
	COLORREF GetEffCellBackColorMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow, DWORD dwFlags = 0);
	CMTblBtn * GetEffCellBtn(HWND hWndCol, long lRow);
	int GetEffCellEditJustify(long lRow, HWND hWndCol);
	int GetEffCellEditJustifyMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow);
	HFONT GetEffCellFont(HDC hDC, long lRow, HWND hWndCol, LPBOOL pbMustDelete, DWORD dwFlags = 0);
	HFONT GetEffCellFont(HDC hDC, CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow, LPBOOL pbMustDelete, DWORD dwFlags = 0);
	CMTblImg GetEffCellImage(long lRow, HWND hWndCol, DWORD dwFlags = 0);
	int GetEffCellIndent(long lRow, HWND hWndCol);
	BOOL GetEffCellText(long lRow, HWND hWndCol, tstring &sText);
	COLORREF GetEffCellTextColor(long lRow, HWND hWndCol, DWORD dwFlags = 0);
	COLORREF GetEffCellTextColorMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow, DWORD dwFlags = 0);
	int GetEffCellTextJustify(long lRow, HWND hWndCol);
	int GetEffCellTextJustifyMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow);
	int GetEffCellTextVAlign(long lRow, HWND hWndCol);
	int GetEffCellTextVAlignMTbl(CMTblCol *pCell, CMTblCol *pCol, CMTblRow *pRow);
	int GetEffCellType(long lRow, HWND hWndCol);
	CMTblCellType * GetEffCellTypePtr(long lRow, HWND hWndCol);
	COLORREF GetEffColHdrBackColor(HWND hWndCol, DWORD dwFlags = 0);
	COLORREF GetEffColHdrFreeBackColor();
	HFONT GetEffColHdrFont(HDC hDC, HWND hWndCol, LPBOOL pbMustDelete, DWORD dwFlags = 0);
	CMTblImg GetEffColHdrImage(HWND hWndCol, DWORD dwFlags = 0);
	COLORREF GetEffColHdrTextColor(HWND hWndCol, DWORD dwFlags = 0);
	COLORREF GetEffColHdrGrpBackColor(CMTblColHdrGrp *pGrp, DWORD dwFlags = 0);
	HFONT GetEffColHdrGrpFont(HDC hDC, CMTblColHdrGrp *pGrp, LPBOOL pbMustDelete, DWORD dwFlags = 0);
	CMTblImg GetEffColHdrGrpImage(CMTblColHdrGrp *pGrp, DWORD dwFlags = 0);
	COLORREF GetEffColHdrGrpTextColor(CMTblColHdrGrp *pGrp, DWORD dwFlags = 0);
	int GetEffColTextJustify(HWND hWndCol);
	COLORREF GetEffCornerBackColor(DWORD dwFlags = 0);
	CMTblImg GetEffCornerImage(DWORD dwFlags = 0);
	BOOL GetEffFocusCell(long &lRow, HWND &hWndCol);
	CMTblBtn * GetEffItemBtn(CMTblItem &item);
	BOOL GetEffItemHighlighting(CMTblItem &item, MTBLHILI &hili);
	BOOL GetEffItemLine(CMTblItem &item, int iLineType, CMTblLineDef &ld);
	COLORREF GetEffRowHdrBackColor(long lRow, DWORD dwFlags = 0);
	long GetEffRowHeight(long lRow, CMTblMetrics *ptm = NULL);
	DWORD GetExportSubstColor(DWORD dwColor);
	inline LPMTBLEXTEDITSUBCLASS GetExtEditSubClass(HWND hWndExtEdit);
	BOOL GetFocusCellRect(LPRECT pRect);
	HRGN GetFocusRgn(LPRECT pRect, int iType, BOOL bCellMode);
	BOOL GetHTMLText(tstring &sText);
	BOOL GetImgRect(CMTblImg &Img, long lRow, HWND hWndCol, long lCellLeft, long lCellTop, long lCellWid, long lCellHt, long lTextTop, int iFontHeight, long lCellLeadingX, long lCellLeadingY, BOOL bButton, BOOL bDDButton, LPRECT pr);
	int GetItemAreaLeft(CMTblItem &item, long lItemLeft);
	HFONT GetItemBtnFont(LPMTBLITEM pItem, CMTblBtn * pBtn, LPBOOL pbMustDelete);
	int GetItemTextAreaLeft(CMTblItem& item, long lItemLeft, long lItemLeadingX, CMTblImg & Image, BOOL bButton = FALSE);
	int GetItemTextAreaRight(CMTblItem& item, long lItemRight, long lItemLeadingX, CMTblImg & Image, BOOL bButton = FALSE, BOOL bDDButton = FALSE);
	LPMTBLHILI GetItemHighlighting(CMTblItem &item);
	LPMTBLLINEDEF GetItemLine(CMTblItem &item, int iLineType);
	HWND GetLastLockedColumn();
	HWND GetLastMergedCell(HWND hWndCol, long lRow, LPMTBLMERGECELLS pMergeCells = NULL);
	BOOL GetLastMergedCellEx(HWND hWndCol, long lRow, HWND &hWndLastMergedCol, long &lLastMergedRow, LPINT piCols = NULL, LPINT piRows = NULL, LPMTBLMERGECELLS pMergeCells = NULL);
	long GetLastMergedRow(long lRow, LPMTBLMERGEROWS pMergeRows = NULL);
	BOOL GetMasterCell(HWND &hWndCol, long &lRow);
	HWND GetMergeCell(HWND hWndCol, long lRow);
	BOOL GetMergeCellEx(HWND hWndCol, long lRow, HWND &hWndMergeCol, long &lMergeRow);
	LPMTBLMERGECELLS GetMergeCells(BOOL bCreateExternal = FALSE, WORD wRowFlagsOn = 0, WORD wRowFlagsOff = ROW_Hidden, DWORD dwColFlagsOn = COL_Visible, DWORD dwColFlagsOff = 0);
	int GetMergeCellType(HWND hWndCol, long lRow);
	long GetMergeRow(long lRow);
	LPMTBLMERGEROWS GetMergeRows(BOOL bCreateExternal = FALSE, WORD wRowFlagsOn = 0, WORD wRowFlagsOff = ROW_Hidden);
	BOOL GetMergeRowRange(long lRow, DWORD dwFlags, RowRange & rr, LPMTBLMERGECELLS pMergeCells = NULL);
	BOOL GetNodeRect(long lCellLeft, long lTextTop, int iFontHeight, int iLevel, int iCellLeadingX, int iCellLeadingY, LPRECT pr);
	COLORREF GetNormalColor(COLORREF clr);
	void GetPaintMergeCells();
	void GetPaintMergeRowHdrs();
	DWORD GetRowAlternateBackColor(long lRow);
	HANDLE GetRowFlagImg(DWORD dwFlag, BOOL bUserDefined);
	WORD GetRowHdrFlags();
	int GetRowHdrRect(long lRow, LPRECT pr, DWORD dwFlags = 0);
	BOOL GetRowFocusRect(long lRow, LPRECT pRect);
	long GetRowNrFocus();
	int GetRowRect(long lRow, LPRECT pr, DWORD dwFlags = 0);
	long GetRowSel(int iMode, long lRow = TBL_Error, RowPtrVector *pvRows = NULL);
	long GetRowSelChanges(HARRAY haRows, HARRAY haChanges);
	HRGN GetRowSelChangesRgn(long lRow = TBL_Error);
	int GetRowYCoord(long lRow, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis, BOOL bInVisRangeOnly = TRUE);
	BOOL GetScreenMetrics(int &iLeft, int &iTop, int &iWid, int &iHt);
	BOOL GetScreenRectFromPoint(POINT &pt, RECT &rMetrics);
	BOOL GetScrollBarV(CMTblMetrics * ptm, BOOL bSplit, LPHWND phWndScr, LPINT piScrCtl);
	int GetScrollPosV(CMTblMetrics * ptm, BOOL bSplit);
	HWND GetSepCol(int iX, int iY);
	long GetSepRow(POINT pt);
	__int64 GetSystemTimeI64();
	CMTblTip * GetTip(long lRow, HWND hWndCol, int iType, BOOL bEnsureValid = FALSE);
	int GetTrailingCRLFCount(LPCTSTR lpsStr);
	LRESULT HandlePaint(HWND hWnd, WPARAM wParam, LPARAM lParam);
	BOOL InvalidateBtn(HWND hWndCol, long lRow, CMTblMetrics *ptm, BOOL bErase);
	BOOL InvalidateCell(HWND hWndCol, long lRow, CMTblMetrics *ptm, BOOL bErase);
	BOOL InvalidateCol(HWND hWndCol, BOOL bErase);
	BOOL InvalidateColHdr(BOOL bErase);
	BOOL InvalidateColHdr(HWND hWndCol, BOOL bErase);
	BOOL InvalidateColHdr(HWND hWndCol, BOOL bInclColHdrGrp, BOOL bErase);
	BOOL InvalidateColSelChanges(HWND hWndCol = NULL, BOOL bErase = FALSE);
	BOOL InvalidateCorner(BOOL bErase);
	BOOL InvalidateEdit(BOOL bErase, BOOL bForceAdaptListClip = FALSE, BOOL bForceColorEdit = FALSE, BOOL bForceEditFont = FALSE);
	BOOL InvalidateRow(long lRow, BOOL bErase, BOOL bInclMerged = FALSE, BOOL bToEndOfArea = FALSE);
	BOOL InvalidateRowFocus(long lRow, BOOL bErase);
	BOOL InvalidateRowHdr(long lRow, BOOL bErase, BOOL bInclMerged = FALSE);
	BOOL InvalidateRowSelChanges(long lRow = TBL_Error, BOOL bErase = FALSE);
	BOOL InvalidateTreeCol(BOOL bErase);
	BOOL IsButtonHotkeyPressed();
	BOOL IsCheckBox(HWND hWnd);
	BOOL IsCheckBoxCell(HWND hWndCol, long lRow);
	BOOL IsCheckBoxCol(HWND hWndCol);
	BOOL IsCellDisabled(HWND hWndCol, long lRow);
	BOOL IsCellReadOnly(HWND hWndCol, long lRow);
	BOOL IsColMirrored(HWND hWndCol);
	BOOL IsColSubClassed(HWND hWndCol){ return GetColSubClass(hWndCol) != NULL; }
	BOOL IsContainerSubClassed(HWND hWnd);
	BOOL IsDropDownBtn(HWND hWnd);
	BOOL IsEnabledHyperlinkCell(HWND hWndCol, long lRow);
	BOOL IsExtEditSubClassed(HWND hWndExtEdit){ return GetExtEditSubClass(hWndExtEdit) != NULL; }
	BOOL IsFocusCell(HWND hWndCol, long lRow);
	BOOL IsGroupedColHdr(HWND hWndCol);
	BOOL IsInEditMode();
	BOOL IsIndented(HWND hWndCol, long lRow);
	BOOL IsListClipSubClassed(HWND hWnd);
	BOOL IsListClipCtrlSubClassed(HWND hWnd);
	BOOL IsLockedColumn(HWND hWndCol);
	BOOL IsMergeCell(HWND hWndCol, long lRow);
	BOOL IsMergedCell(HWND hWndCol, long lRow);
	BOOL IsMyWindow(HWND hWnd);
	BOOL IsOwnerDrawnItem(HWND hWndParam, long lParam, int iType);
	BOOL IsOverCornerSep(int iX, int iY);
	BOOL IsOverHyperlink(HWND hWndCol, long lRow, DWORD dwMFlags);
	BOOL IsRow(long lRow);
	int IsScrollBarV(HWND hWndScr, int iScrCtl, CMTblMetrics *ptm = NULL);
	BOOL IsSplitted();
	BOOL IsStandaloneColHdrGrpCol(HWND hWndCol);
	BOOL IsTipTypeEnabled(DWORD dwTipType);
	BOOL IsValidLineStyle(int iStyle) { return iStyle == MTLS_INVISIBLE || iStyle == MTLS_DOT || iStyle == MTLS_SOLID || iStyle == MTLS_UNDEF; };
	long KillEditEsc(HWND hWnd);
	BOOL KillEditNeutral();
	BOOL KillFocusFrameI();
	void LoadChildRows(long lRow);
	void LogLine(LPCTSTR lpsFile, LPCTSTR lpsText);
	void LogLineFocusPaint(LPTSTR lpsText);
	long LockPaint(void);
	BOOL MustPaintRowFocus();
	BOOL MustRedrawSelections();
	BOOL MustUseEndEllipsis(HWND hWndCol);
	BOOL MustShowBtn(HWND hWndCol, long lRow, BOOL bNoWithCheck = FALSE, BOOL bPermanent = FALSE);
	BOOL MustShowDDBtn(HWND hWndCol, long lRow, BOOL bNoWithCheck = FALSE, BOOL bPermanent = FALSE);
	BOOL MustSuppressRowFocusCTD();
	BOOL MustTrapColSelChanges();
	BOOL MustTrapRowSelChanges();
	BOOL NormalizeHierarchy(long lRow, int iReason);
	BOOL ObjFromPoint(POINT pt, LPLONG plRow, LPHWND phWndCol, LPDWORD pdwFlags, LPDWORD pdwMFlags);
	void ODIPrint(LPMTBLODI pODI);
	BOOL PaintBtn(LPMTBLPAINTBTN pb);
	BOOL PaintCell(MTBLPAINTCELL &pc);
	BOOL PaintDDBtn(LPMTBLPAINTDDBTN pb);
	BOOL PaintFocus(LPCRECT pRect);
	BOOL PaintNode(MTBLPAINTCELL &pc, LPRECT pr, BOOL bInverted = FALSE);
	BOOL PaintRow();
	BOOL PaintTable();
	BOOL ProcessHotkey(UINT uKeyMsg, WORD wParam, LONG lParam);
	BOOL QueryCellFlags(long lRow, HWND hWndCol, DWORD dwFlags);
	BOOL QueryColumnFlags(HWND hWndCol, DWORD dwFlags);
	BOOL QueryColumnHdrFlags(HWND hWndCol, DWORD dwFlags);
	BOOL QueryFlags(DWORD dwFlags) { return (m_dwFlags & dwFlags) != 0; };
	BOOL QueryRowFlagImgFlags(DWORD dwRowFlag, DWORD dwFlags);
	BOOL QueryRowFlags(long lRow, DWORD dwFlags);
	BOOL QueryTreeFlags(DWORD dwFlags) { return (m_dwTreeFlags & dwFlags) != 0; };
	BOOL QueryVisRange(long &lRowFrom, long &lRowTo, BOOL bSplitRows = FALSE);
	BOOL RegisterBtnClass();
	BOOL ResetFocusData();
	void ResetTip();
	long RowFromPoint(POINT pt);
	BOOL SetBtnWndData(CMTblBtn *pBtn, int iHeight);
	BOOL SetCellFlags(long lRow, HWND hWndCol, CMTblMetrics *ptm, DWORD dwFlags, BOOL bSet, BOOL bInvalidate, DWORD *pdwDiff = NULL);
	BOOL SetCellIndentLevel(long lRow, HWND hWndCol, long lIndentLevel, CMTblMetrics *ptm, BOOL bInvalidate);
	BOOL SetCellRangeFlags(LPMTBLCELLRANGE pcr, DWORD dwFlags, BOOL bSet, BOOL bInvalidate);
	BOOL SetColumnTitle(HWND hWndCol, LPCTSTR lpsTitle);
	BOOL SetCurrentCellType(HWND hWndCol, CMTblCellType *pct);
	BOOL SetDropDownListContext(HWND hWndCol, long lRow);
	BOOL SetEffCellPropDetOrder(DWORD dwOrder);
	BOOL SetFlags(DWORD dwFlags, BOOL bSet);
	BOOL SetFocusCell(long lRow, HWND hWndCol);
	BOOL SetItemHighlighted(CMTblItem &item, BOOL bSet);
	BOOL SetMergeRowFlags(LPMTBLMERGEROW pmr, DWORD dwFlags, BOOL bSet, BOOL bInternal = FALSE, long lRowExcept = TBL_Error);
	BOOL SetNodeImages(HANDLE hImg, HANDLE hImgExp);
	BOOL SetOwnerDrawnItem(HWND hWndParam, long lParam, int iType, BOOL bSet);
	void SetPaintColData(HWND hWndCol, int iColPos, BOOL bPaint, BOOL bLocked, BOOL bLastLocked, BOOL bFirstUnlocked, BOOL bFirstVis, BOOL bEmpty, BOOL bRowHdr, LPPOINT pptLog, LPPOINT pptPhys, LPMTBLPAINTCOL pPaintCol);
	BOOL SetParentRow(long lRow, long lParentRow, DWORD dwFlags);
	BOOL SetRowFlagBitmap(DWORD dwFlag, int iBitmRes, DWORD dwTranspClr);
	BOOL SetRowFlagImg(DWORD dwFlag, HANDLE hImage, BOOL bUserDefined, DWORD dwFlags = 0);
	BOOL SetRowHdrFlags(WORD wFlagsSet, BOOL bSet);
	BOOL SetTreeFlags(DWORD dwFlags, BOOL bSet);
	BOOL Sort(HARRAY hWndaCols, HARRAY naFlags, long lMinRow, long lMaxRow, long lFromLevel = -1, long lToLevel = -1, long lParentRow = TBL_Error);
	BOOL SortTree(DWORD dwFlags, BOOL bNoRedraw = FALSE);
	void StartInternalTimer(BOOL bProcessImmediate = FALSE);
	void StopInternalTimer();
	BOOL SubClassCol(HWND hWndCol);
	BOOL SubClassCols();
	BOOL SubClassContainer(HWND hWnd);
	BOOL SubClassExtEdit(HWND hWndCol, HWND hWndExtEdit);
	BOOL SubClassListClip(HWND hWnd);
	BOOL SubClassListClipCtrl(HWND hWnd);
	BOOL SwapRows(long lRow1, long lRow2);
	long UnlockPaint(void);
	BOOL UpdateFocus(long lRow = TBL_Error, HWND hWndCol = NULL, int iRedrawMode = UF_REDRAW_AUTO, int iSelChangeMode = UF_SELCHANGE_AUTO, DWORD dwFlags = 0);
	BOOL Validate(HWND hWndColHasFocus, HWND hWndColGetFocus);

	//************************************************************
	//Arbeitsbereich Alexander Hakert
	//************************************************************

	int Print(HWND hWndTbl, PRINTOPTIONS printOptions);
	BOOL GetTableInfo(HWND hWndTbl, PRINTOPTIONS printOptions, TABLEINFO &tableInfo);
	int GetColsToPrint(TABLEINFO tableInfo, VecColInfo &vecColInfo, PRINTOPTIONS printOptions);
	BOOL GetPrintMetrics(HDC hDeviceContext, PRINTOPTIONS printOptions, TABLEINFO tableInfo, VecColInfo &vecColInfo, LPPRINTMETRICS lpPrintMetrics);
	BOOL IsPageCountNeeded(PRINTOPTIONS printOptions);
	BOOL IsPageCountPhIncl(std::vector<PRINTLINE> printLines, long lItems);
	long GetPageCount(PRINTOPTIONS printOptions, PRINTMETRICS printMetrics, TABLEINFO tableInfo, VecColInfo vecColInfo, HDC hDeviceContext);
	int PrintPage(PRINTPAGEPARAMS printPageParams, long &lLastPrintedRow, long &lLastPrintedTotal);
	RECT PrintLine(PRINTLINEPARAMS printPageParams);
	int PrintLine_GetTextHeight(tstring tString, PRINTLINEPARAMS printPageParams, DWORD dwDrawFlags);
	SIZE PrintLine_GetImageSize(HANDLE hImage, PRINTLINEPARAMS printLineParams);
	int PrintLine_GetImageHeight(HANDLE hImage, PRINTLINEPARAMS printLineParams);
	void PrintLine_DrawImage(HANDLE hImage, int iJustify, RECT rect, PRINTLINEPARAMS printLineParams);
	tstring ExpandPlaceholders(tstring tString);
	void GetUIStr(PRINTOPTIONS printOptions, LPTSTR lpsKey, tstring &tsResult);
	int PixTblToPrint(int iPix, BOOL bVertical, TABLEINFO tableInfo, PRINTMETRICS printMetrics);
	int PrintCells(HDC hDeviceContext, PRINTCELLSPARAMS printCellsParams, long &lLastPrintedRow, int &iYBottom);
	void BorderFlagsFromGridType(GridType gridType, RECT &rect);
	int PrintColHeaders(HDC hDeviceContext, PRINTCOLHEADERSPARAMS colHeaderParams);
	BOOL GetColHeaderGrpText(HWND hWndTbl, LPVOID pHeaderGrp, tstring &tString);
	long PrintColHeadersGetTextAreaLeft(BOOL bHeaderGrp, PRINTCOLHEADERSPARAMS colHeaderParams, PRINTCELLPARAMS printCellParams, int iColInfoIndex, LPVOID pHeaderGrp);
	void PrintColHeadersGetImageRect(BOOL bHeaderGrp, PRINTCOLHEADERSPARAMS colHeaderParams, PRINTCELLPARAMS &printCellParams, int iColInfoIndex, LPVOID pHeaderGrp);
	long PrintColHeadersGetTextAreaRight(BOOL bHeaderGrp, PRINTCOLHEADERSPARAMS colHeaderParams, PRINTCELLPARAMS printCellParams, int iColInfoIndex, LPVOID pHeaderGrp);
	int PrintCell(PRINTCELLPARAMS printCellParams);
	long PrintCellsGetRowHeight(long lRow, PRINTCELLSPARAMS printCellsParams);
	int PrintCellsGetBestCellHeight(HWND hWndCol, long lRow, PRINTCELLSPARAMS printCellsParams);
	BOOL PrintGetCellFont(HDC hDeviceContext, COLINFO colInfo, long lRow, PRINTMETRICS printMetrics, TABLEINFO tableInfo, tstring &tsFontName, double &dFontHeight, long &lFontStyle);

	long PrintCellsGetEffCellIndent(PRINTCELLSPARAMS printCellsParams, int iColIndex, long lRow);
	long PrintCellsGetCellAreaLeft(PRINTCELLSPARAMS printCellsParams, int iEff, long lEffRow, LPCELLTOPRINT lpCellToPrint);
	long PrintCellsGetCellTextAreaLeft(PRINTCELLSPARAMS printCellsParams, int iEffIndex, long lEffRow, PRINTCELLPARAMS printCellParams, LPCELLTOPRINT lpCellToPrint, int iTextOffsetX);
	long PrintCellsGetCellTextAreaRight(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS printCellParams, int iEffIndex, long lEffRow);
	void PrintCellsGetTextRect(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS &printCellParams);
	void PrintCellsGetImageRect(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS &printCellParams, int iEffIndex, long lEffRow, LPCELLTOPRINT lpCellToPrint, int iFontHeight);
	BOOL PrintCellsCheckRowFlags(TABLEINFO tableInfo, long lRow, WORD dwFlagsOn, WORD dwFlagsOff);
	void PrintCellsGetVertTreeLinesToParent(PRINTCELLSPARAMS printCellsParams, PRINTCELLPARAMS &printCellParams, int iEffRowLevel, int iTreeIndentPerLevel, int iFontHeight, LPCELLTOPRINT lpCellToPrint);
	BOOL PrintGetNodeRect(long lCellLeft, long lTextTop, int iFontHeight, int iNodeLevel, int iCellLeadingX, int iCellLeadingY, int iNodeWid, int iNodeHt, int iTreeIndentPerLevel, RECT &nodeRect);
	BOOL PrintFontExists(tstring tsFontName);
	HFONT PrintCreateFont(HDC hDeviceContext, tstring tsFontName, double dFontHeight, long lStyle);
	DWORD GetPrintJobStatus(tstring tsPrinter, int iPrintJob);
	void FreeDevModes();

	std::vector<HWND> m_vecHwndDisabled;
	double m_dPrintScale;
	int m_iPrintJob;
	tstring m_tsPrintFonts;

	//************************************************************
}
MTBLSUBCLASS, *LPMTBLSUBCLASS;

#endif //_MTBLSUBCLASS_H