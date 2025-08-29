// mtblapi.h

#ifndef _INC_MTBLAPI
#define _INC_MTBLAPI

// Includes
#include <limits.h>
#include "ctd.h"
#include "tblflags.h"
#include "treeflags.h"
#include "cellmodeflags.h"
#include "rowflags.h"
#include "tipflags.h"
#include "chgflags.h"
#include "colflags.h"
#include "colhdrflags.h"
#include "cellflags.h"
#include "export.h"
#include "merge.h"
#include "mtimg.h"
#include "lines.h"
#include "color.h"
#include "cmtblitem.h"

// Message-Offset
//extern UINT guiMsgOffs;

// Allgemeine Konstanten
#define MTBL_NODEFAULTACTION 1

// Standardsprachen
#define MTBL_LNG_GERMAN 0
#define MTBL_LNG_ENGLISH 1

// Objekte
#define MTBL_OBJ_CELL 1
#define MTBL_OBJ_COLUMN 2
#define MTBL_OBJ_COLUMN_HEADER 3
#define MTBL_OBJ_COLUMN_HEADER_GROUP 4
#define MTBL_OBJ_ROW 5
#define MTBL_OBJ_ROW_HEADER 6
#define MTBL_OBJ_CORNER 7

#define MTBL_OBJ_HEADERS 8

// Font
#define MTBL_FONT_UNDEF_NAME _T("")
#define MTBL_FONT_UNDEF_SIZE USHRT_MAX
#define MTBL_FONT_UNDEF_ENH UCHAR_MAX

// OwnerDrawnItem
#define MTBL_ODI_CELL 1
#define MTBL_ODI_COLUMN 2
#define MTBL_ODI_COLHDR 3
#define MTBL_ODI_COLHDRGRP 4
#define MTBL_ODI_ROW 5
#define MTBL_ODI_ROWHDR 6
#define MTBL_ODI_CORNER 7

// Tip
#define MTBL_TIP_DEFAULT 0
#define MTBL_TIP_CELL 1
#define MTBL_TIP_CELL_TEXT 2
#define MTBL_TIP_CELL_COMPLETETEXT 4
#define MTBL_TIP_CELL_IMAGE 8
#define MTBL_TIP_CELL_BTN 16
#define MTBL_TIP_COLHDR 32
#define MTBL_TIP_COLHDRGRP 64
#define MTBL_TIP_ROWHDR 128
#define MTBL_TIP_CORNER 256
#define MTBL_TIP_COLHDR_BTN 512

// MTblAutoSizeColumn - Konstanten
#define MTASC_ALLROWS 1
#define MTASC_BTNSPACE 2
#define MTASC_HIDDENCOLS 4
#define MTASC_HIDDENROWS 8
#define MTASC_SPLITROWS 16
#define MTASC_INCREASE_ONLY 32

// MTblAutoSizeRow - Konstanten
#define MTASR_ALLROWS 1
#define MTASR_BTNSPACE 2
#define MTASR_HIDDENCOLS 4
#define MTASR_HIDDENROWS 8
#define MTASR_SPLITROWS 16
#define MTASR_INCREASE_ONLY 32
#define MTASR_CONTEXTROW_ONLY 64

// MTbl*EffCellPropDet
#define MTECPD_ORDER_CELL_COL_ROW 1
#define MTECPD_ORDER_CELL_ROW_COL 2
#define MTECPD_ORDER_COL_CELL_ROW 3
#define MTECPD_ORDER_COL_ROW_CELL 4
#define MTECPD_ORDER_ROW_CELL_COL 5
#define MTECPD_ORDER_ROW_COL_CELL 6

#define MTECPD_FLAG_COMB_FONT_ENH 1

// MTblCollapseRow - Konstanten
#define MTCR_REDRAW 1
#define MTCR_BY_USER 2

// MTblDefineHighlighting - Konstanten
#define MTDHL_REDRAW 1
#define MTDHL_REMOVE 2

// Erweiterte Nachrichten - Optionen
#define MTEM_CORNERSEP_EXCLUSIVE 1
#define MTEM_COLHDRSEP_EXCLUSIVE 2
#define MTEM_ROWHDRSEP_EXCLUSIVE 4

// MTblExpandRow - Konstanten
#define MTER_REDRAW 1
#define MTER_RECURSIVE 2
#define MTER_BY_USER 4

// MTblGet*Rect - Konstanten
#define MTGR_VISIBLE_PART 1
#define MTGR_EXCLUDE_COLHDRGRP 2
#define MTGR_INCLUDE_MERGED 4
#define MTGR_TO_END_OF_AREA 8

#define MTGR_RET_ERROR -1
#define MTGR_RET_INVISIBLE 0
#define MTGR_RET_PARTLY_VISIBLE 1
#define MTGR_RET_VISIBLE 2

// MTblInsertChildRow - Konstanten
#define MTICR_REDRAW 1
#define MTICR_FIRST 2

// Mousewheel-Optionen
#define MTMW_ROWS_PER_NOTCH 1
#define MTMW_SPLITROWS_PER_NOTCH 2
#define MTMW_PAGE_KEY 3

// MTblObjFromPoint - Konstanten
#define MTOFP_OVERNODE 1
#define MTOFP_OVERCELLTEXT 2 /* Nur für Kompatibilität */
#define MTOFP_OVERTEXT 2
#define MTOFP_OVERCELLIMAGE 4 /* Nur für Kompatibilität */
#define MTOFP_OVERIMAGE 4
#define MTOFP_OVERMERGEDCELL 8
#define MTOFP_OVERCOLHDRGRP 16
#define MTOFP_OVERBTN 32
#define MTOFP_OVERDDBTN 64
#define MTOFP_OVERCOLHDRSEP 128
#define MTOFP_OVERCORNERSEP 256
#define MTOFP_OVERROWHDRSEP 512
#define MTOFP_OVERMERGEDROWHDRSEP 1024

// MTblSet*Color - Konstanten
#define MTSC_REDRAW 1

// MTblSet*Font - Konstanten
#define MTSF_REDRAW 1

// MTblSet*Merge - Konstanten
#define MTSM_REDRAW 1
#define MTSM_NO_CELLMERGE 2
#define MTSM_INS_DEL_ROWS 4

// MTblSetParentRow - Konstanten
#define MTSPR_REDRAW 1
#define MTSPR_AS_ORIG 2

// MTblSetRowFlags - Konstanten
#define MTSRF_REDRAW 1

// MTblSort - Konstanten
#define MTS_ASC 0
#define MTS_DT_DEFAULT 0
#define MTS_DESC 1
#define MTS_IGNORECASE 2
#define MTS_DT_STRING 4
#define MTS_DT_NUMBER 8
#define MTS_DT_DATETIME 16
#define MTS_DT_CUSTOM 32
#define MTS_DT_CUSTOM_STRING 64
#define MTS_STRINGSORT 128

// MTblSortTree - Konstanten
#define MTST_TOPDOWN 0
#define MTST_BOTTOMUP 1

// MTblSwapRows - Konstanten
#define MTSR_REDRAW 1

// MTblSetRowHeight - Konstanten
#define MTSRH_REDRAW 1

// Nachrichten
#define MTM_First (SAM_User + MTblGetMsgOffset( ))
#define MTM_Last (MTM_First + 106)

#define MTM_Internal MTM_First

#define MTM_AreaLBtnDown (MTM_First + 1)
#define MTM_AreaLBtnUp (MTM_First + 2)
#define MTM_AreaLBtnDblClk (MTM_First + 3)
#define MTM_AreaRBtnDown (MTM_First + 4)
#define MTM_AreaRBtnUp (MTM_First + 5)
#define MTM_AreaRBtnDblClk (MTM_First + 6)

#define MTM_ColHdrLBtnDown (MTM_First + 7)
#define MTM_ColHdrLBtnUp (MTM_First + 8)
#define MTM_ColHdrLBtnDblClk (MTM_First + 9)
#define MTM_ColHdrRBtnDown (MTM_First + 10)
#define MTM_ColHdrRBtnUp (MTM_First + 11)
#define MTM_ColHdrRBtnDblClk (MTM_First + 12)

#define MTM_ColHdrSepLBtnDown (MTM_First + 13)
#define MTM_ColHdrSepLBtnUp (MTM_First + 14)
#define MTM_ColHdrSepLBtnDblClk (MTM_First + 15)
#define MTM_ColHdrSepRBtnDown (MTM_First + 16)
#define MTM_ColHdrSepRBtnUp (MTM_First + 17)
#define MTM_ColHdrSepRBtnDblClk (MTM_First + 18)

#define MTM_CornerLBtnDown (MTM_First + 19)
#define MTM_CornerLBtnUp (MTM_First + 20)
#define MTM_CornerLBtnDblClk (MTM_First + 21)
#define MTM_CornerRBtnDown (MTM_First + 22)
#define MTM_CornerRBtnUp (MTM_First + 23)
#define MTM_CornerRBtnDblClk (MTM_First + 24)

#define MTM_CornerSepLBtnDown (MTM_First + 87)
#define MTM_CornerSepLBtnUp (MTM_First + 88)
#define MTM_CornerSepLBtnDblClk (MTM_First + 89)
#define MTM_CornerSepRBtnDown (MTM_First + 90)
#define MTM_CornerSepRBtnUp (MTM_First + 91)
#define MTM_CornerSepRBtnDblClk (MTM_First + 92)

#define MTM_RowHdrLBtnDown (MTM_First + 25)
#define MTM_RowHdrLBtnUp (MTM_First + 26)
#define MTM_RowHdrLBtnDblClk (MTM_First + 27)
#define MTM_RowHdrRBtnDown (MTM_First + 28)
#define MTM_RowHdrRBtnUp (MTM_First + 29)
#define MTM_RowHdrRBtnDblClk (MTM_First + 30)

#define MTM_RowHdrSepLBtnDown (MTM_First + 94)
#define MTM_RowHdrSepLBtnUp (MTM_First + 95)
#define MTM_RowHdrSepLBtnDblClk (MTM_First + 96)
#define MTM_RowHdrSepRBtnDown (MTM_First + 97)
#define MTM_RowHdrSepRBtnUp (MTM_First + 98)
#define MTM_RowHdrSepRBtnDblClk (MTM_First + 99)

#define MTM_ColSized (MTM_First + 31)
#define MTM_ColMoved (MTM_First + 32)
#define MTM_ColPosChanged (MTM_First + 71)

#define MTM_RowsSwapped (MTM_First + 33)
#define MTM_RowDeleted (MTM_First + 34)
#define MTM_Reset (MTM_First + 35)
#define MTM_RowInserted (MTM_First + 36)

#define MTM_SplitBarLBtnDown (MTM_First + 37)
#define MTM_SplitBarLBtnUp (MTM_First + 38)
#define MTM_SplitBarLBtnDblClk (MTM_First + 39)
#define MTM_SplitBarRBtnDown (MTM_First + 40)
#define MTM_SplitBarRBtnUp (MTM_First + 41)
#define MTM_SplitBarRBtnDblClk (MTM_First + 42)

#define MTM_SplitBarMoved (MTM_First + 43)

#define MTM_CollapseRow (MTM_First + 44)
#define MTM_CollapseRowDone (MTM_First + 53)
#define MTM_ExpandRow (MTM_First + 45)
#define MTM_ExpandRowDone (MTM_First + 52)
#define MTM_LoadChildRows (MTM_First + 51)
#define MTM_QueryAutoNormHierarchy (MTM_First + 105)

#define MTM_Char (MTM_First + 78)
#define MTM_KeyDown (MTM_First + 46)
#define MTM_LButtonDown (MTM_First + 72)
#define MTM_SetCursor (MTM_First + 80)

#define MTM_Copy (MTM_First + 106)

#define MTM_QuerySortValue (MTM_First + 47)
#define MTM_QuerySortString (MTM_First + 69)

#define MTM_BtnClick (MTM_First + 48)
#define MTM_HyperlinkClick (MTM_First + 49)

#define MTM_Paint (MTM_First + 50)
#define MTM_PaintDone (MTM_First + 70)

#define MTM_ShowCellTip (MTM_First + 54)
#define MTM_ShowCellTextTip (MTM_First + 55)
#define MTM_ShowCellImageTip (MTM_First + 56)
#define MTM_ShowCellBtnTip (MTM_First + 57)
#define MTM_ShowColHdrTip (MTM_First + 58)
#define MTM_ShowColHdrBtnTip (MTM_First + 104)
#define MTM_ShowColHdrGrpTip (MTM_First + 59)
#define MTM_ShowRowHdrTip (MTM_First + 60)
#define MTM_ShowCornerTip (MTM_First + 61)
#define MTM_ShowCellCompleteTextTip (MTM_First + 62)

#define MTM_DrawItem (MTM_First + 63)
#define MTM_QueryBestCellWidth (MTM_First + 64)
#define MTM_QueryBestColHdrWidth (MTM_First + 65)
#define MTM_QueryBestCellHeight (MTM_First + 73)

#define MTM_QueryMaxColWidth (MTM_First + 68)
#define MTM_QueryMaxRowHeight (MTM_First + 74)
#define MTM_QueryMinColWidth (MTM_First + 76)
#define MTM_QueryMinRowHeight (MTM_First + 77)

#define MTM_RowSelChanged (MTM_First + 66)
#define MTM_ColSelChanged (MTM_First + 67)

#define MTM_FocusCellChanged (MTM_First + 75)
#define MTM_QueryNoFocusOnCell (MTM_First + 82)

#define MTM_DragCanAutoStart (MTM_First + 79)

#define MTM_BeforeEnterEditMode (MTM_First + 81)

#define MTM_QueryExcelNumberFormat (MTM_First + 83)
#define MTM_QueryExcelColWidth (MTM_First + 84)
#define MTM_QueryExcelRowHeight (MTM_First + 85)

#define MTM_QueryNoCellSelection (MTM_First + 86)

#define MTM_QueryExcelValue (MTM_First + 93)

#define MTM_LinesPerRowChanged (MTM_First + 100)
#define MTM_RowSized (MTM_First + 101)

#define MTM_MouseEnterItem (MTM_First + 102)
#define MTM_MouseLeaveItem (MTM_First + 103)

extern "C"
{
// API-Funktionen
BOOL WINAPI MTblApplyProperties( HWND hWndTbl );
BOOL WINAPI MTblAutoSizeColumn( HWND hWndTbl, HWND hWndCol, DWORD dwFlags );
BOOL WINAPI MTblAutoSizeRows( HWND hWndTbl, DWORD dwFlags );
int WINAPI MTblCalcRowHeight( HWND hWndTbl, int iLinesPerRow );
BOOL WINAPI MTblClearCellSelection( HWND hWndTbl );
BOOL WINAPI MTblClearExportSubstColors( HWND hWndTbl );
long WINAPI MTblCollapseRow( HWND hWndTbl, long lRow, WPARAM wFlags );
HWND WINAPI MTblColFromPoint( HWND hWndTbl, int iX, int iY );
BOOL WINAPI MTblCopyRows( HWND hWndTbl, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff );
BOOL WINAPI MTblCopyRowsAndHeaders( HWND hWndTbl, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff );
LPVOID WINAPI MTblCreateColHdrGrp( HWND hWndTbl, LPTSTR lpsText, HARRAY haCols, int iCols, DWORD dwFlags );
BOOL WINAPI MTblDefineButtonCell( HWND hWndCol, long lRow, BOOL bShow, int iWidth, LPCTSTR lpsText, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblDefineButtonColumn( HWND hWndCol, BOOL bShow, int iWidth, LPCTSTR lpsText, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblDefineCellMode( HWND hWndTbl, BOOL bActive, DWORD dwFlags );
BOOL WINAPI MTblDefineCellType( HWND hWndCol, long lRow, int iType, DWORD dwFlags, int iLines, LPCTSTR lpsCheckedVal, LPCTSTR lpsUnCheckedVal );
BOOL WINAPI MTblDefineButtonHotkey( HWND hWndTbl, int iKey, BOOL bCtrl, BOOL bShift, BOOL bAlt );
BOOL WINAPI MTblDefineColLines( HWND hWndTbl, int iStyle, COLORREF clr );
BOOL WINAPI MTblDefineEffCellPropDet( HWND hWndTbl, DWORD dwOrder, DWORD dwFlags );
BOOL WINAPI MTblDefineHighlighting( HWND hWndTbl, long lItem, long lPart, HUDV hUdv, DWORD dwFlags );
BOOL WINAPI MTblDefineHyperlinkCell( HWND hWndCol, long lRow, BOOL bEnable, DWORD dwFlags );
BOOL WINAPI MTblDefineHyperlinkHover( HWND hWndTbl, BOOL bEnabled, long lAddFontEnh, DWORD dwTextColor );
BOOL WINAPI MTblDefineItemButton( HUDV hUdvItem, HUDV hUdvBtnDef );
BOOL WINAPI MTblDefineItemButtonC( LPVOID pItemV, LPVOID pBtnDefV );
BOOL WINAPI MTblDefineItemLine(HUDV hUdvItem, HUDV hUdvLineDef, int iLineType);
BOOL WINAPI MTblDefineItemLineC(LPVOID pItemV, LPVOID pLineDefV, int iLineType);
BOOL WINAPI MTblDefineFocusFrame( HWND hWndTbl, int iThickness, int iOffsX, int iOffsY, DWORD dwColor );
BOOL WINAPI MTblDefineLastLockedColLine( HWND hWndTbl, int iStyle, COLORREF clr );
BOOL WINAPI MTblDefineRowLines( HWND hWndTbl, int iStyle, COLORREF clr );
BOOL WINAPI MTblDefineTree( HWND hWndTbl, HWND hWndTreeCol, int iNodeSize, int iIndent );
BOOL WINAPI MTblDefineTreeLines( HWND hWndTbl, int iStyle, COLORREF clr );
BOOL WINAPI MTblDeleteColHdrGrp( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags );
int WINAPI MTblDeleteDescRows( HWND hWndTbl, long lRow, WORD wMethod );
BOOL WINAPI MTblDeleteRowUserData( HWND hWndTbl, long lRow, LPTSTR lpsKey );
BOOL WINAPI MTblDisableCell( HWND hWndCol, long lRow, BOOL bDisable );
BOOL WINAPI MTblDisableRow( HWND hWndTbl, long lRow, BOOL bDisable );
BOOL WINAPI MTblEnableExtMsgs( HWND hWndTbl, BOOL bSet );
BOOL WINAPI MTblEnableMWheelScroll( HWND hWndTbl, BOOL bSet );
BOOL WINAPI MTblEnableTipType( HWND hWndTbl, DWORD dwType, BOOL bSet );
long WINAPI MTblEnumColHdrGrpCols( HWND hWndTbl, LPVOID pGrp, HARRAY haCols );
long WINAPI MTblEnumColHdrGrps( HWND hWndTbl, HARRAY haGrps );
long WINAPI MTblEnumExportSubstColors( HWND hWndTbl, HARRAY haTblClr, HARRAY haExcClr );
long WINAPI MTblEnumFonts( HARRAY haFontNames );
long WINAPI MTblEnumFontSizes( HSTRING hsFontName, HARRAY haFontSizes );
long WINAPI MTblEnumRowUserData( HWND hWndTbl, long lRow, HARRAY haKeys, HARRAY haValues );
long WINAPI MTblExpandRow( HWND hWndTbl, long lRow, WPARAM wFlags );
int WINAPI MTblExportToHTML( HWND hWndTbl, LPCTSTR lpsFile, DWORD dwHTMLFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff );
int WINAPI MTblFetchMissingRows( HWND hWndTbl );
LPVOID WINAPI MTblFindMergeCell( HWND hWndCol, long lRow, int iMode, LPVOID pMergeCells );
int WINAPI MTblFindNextCol( HWND hWndTbl, LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff );
BOOL WINAPI MTblFindNextRow( HWND hWndTbl, LPLONG plRow, long lMaxRow, WORD wFlagsOn, WORD wFlagsOff );
int WINAPI MTblFindPrevCol( HWND hWndTbl, LPHWND phWndCol, DWORD dwFlagsOn, DWORD dwFlagsOff );
BOOL WINAPI MTblFindPrevRow( HWND hWndTbl, LPLONG plRow, long lMinRow, WORD wFlagsOn, WORD wFlagsOff );
BOOL WINAPI MTblFreeMergeCells( HWND hWndTbl, LPVOID pMergeCells );
BOOL WINAPI MTblGetAltRowBackColors( HWND hWndTbl, BOOL bSplitRows, LPDWORD pdwColor1, LPDWORD pdwColor2 );
int WINAPI MTblGetBestCellHeight( HWND hWndTbl, HWND hWndCol, long lRow, long lCellLeadingY, long lCharBoxY, HFONT hUseFont = NULL, LPVOID pMergeCells = NULL, BOOL bCheckboxText = FALSE, BOOL bIgnoreButtons = FALSE );
int WINAPI MTblGetBestCellWidth( HWND hWndTbl, HWND hWndCol, long lRow, long lCellLeadingX, HFONT hUseFont = NULL, LPVOID pMergeCells = NULL, BOOL bCheckboxText = FALSE );
int WINAPI MTblGetBestColHdrWidth( HWND hWndTbl, HWND hWndCol, long lCellLeadingX, HFONT hUseFont = NULL );
BOOL WINAPI MTblGetCellBackColor( HWND hWndCol, long lRow, LPDWORD pdwColor );
BOOL WINAPI MTblGetCellFont( HWND hWndCol, long lRow, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
BOOL WINAPI MTblGetCellIndentLevel( HWND hWndCol, long lRow, LPLONG plIndentLevel );
BOOL WINAPI MTblGetCellImage( HWND hWndCol, long lRow, LPHANDLE phImage );
BOOL WINAPI MTblGetCellImageExp( HWND hWndCol, long lRow, LPHANDLE phImage );
BOOL WINAPI MTblGetCellMerge( HWND hWndCol, long lRow, LPINT piCount );
BOOL WINAPI MTblGetCellMergeEx( HWND hWndCol, long lRow, LPINT piColCount, LPINT piRowCount );
int WINAPI MTblGetCellRect( HWND hWndCol, long lRow, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom );
int WINAPI MTblGetCellRectEx( HWND hWndCol, long lRow, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom, DWORD dwFlags );
BOOL WINAPI MTblGetCellTextColor( HWND hWndCol, long lRow, LPDWORD pdwColor );
int WINAPI MTblGetCellTextLeft( HWND hWndTbl, long lCellLeft, long lRow, int iCellLeadingX );
int WINAPI MTblGetCellType( HWND hWndCol, long lRow );
int WINAPI MTblGetChildRowCount( HWND hWndTbl, long lRow );
LPVOID WINAPI MTblGetColHdrGrp( HWND hWndCol );
BOOL WINAPI MTblGetColHdrGrpBackColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor );
BOOL WINAPI MTblGetColHdrGrpFont( HWND hWndTbl, LPVOID pGrp, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
BOOL WINAPI MTblGetColHdrGrpImage( HWND hWndTbl, LPVOID pGrp, LPHANDLE phImage );
BOOL WINAPI MTblGetColHdrGrpText( HWND hWndTbl, LPVOID pGrp, LPHSTRING phsText );
BOOL WINAPI MTblGetColHdrGrpTextColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor );
int WINAPI MTblGetColHdrGrpTextLineCount( HWND hWndTbl, LPVOID pGrp );
int WINAPI MTblGetColHdrRect( HWND hWndCol, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom );
int WINAPI MTblGetColHdrRectEx( HWND hWndCol, LPINT piLeft, LPINT piTop, LPINT piRight, LPINT piBottom, DWORD dwFlags );
long WINAPI MTblGetColSelChanges( HWND hWndTbl, HARRAY haCols, HARRAY haChanges );
BOOL WINAPI MTblGetColumnBackColor( HWND hWndCol, LPDWORD pdwColor );
BOOL WINAPI MTblGetColumnFont( HWND hWndCol, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
BOOL WINAPI MTblGetColumnTextColor( HWND hWndCol, LPDWORD pdwColor );
BOOL WINAPI MTblGetColumnTitle( HWND hWndCol, LPHSTRING phsTitle );
BOOL WINAPI MTblGetColumnHdrBackColor( HWND hWndCol, LPDWORD pdwColor );
BOOL WINAPI MTblGetColumnHdrFont( HWND hWndCol, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
BOOL WINAPI MTblGetColumnHdrFreeBackColor( HWND hWndTbl, LPDWORD pdwColor );
BOOL WINAPI MTblGetColumnHdrImage( HWND hWndCol, LPHANDLE phImage );
BOOL WINAPI MTblGetColumnHdrTextColor( HWND hWndCol, LPDWORD pdwColor );
BOOL WINAPI MTblGetCornerBackColor( HWND hWndTbl, LPDWORD pdwColor );
BOOL WINAPI MTblGetCornerImage( HWND hWndTbl, LPHANDLE phImage );
HANDLE WINAPI MTblGetEditHandle(HWND hWndTbl);
BOOL WINAPI MTblGetEffCellBackColor( HWND hWndCol, long lRow, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffCellFont( HWND hWndCol, long lRow, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
HFONT WINAPI MTblGetEffCellFontHandle( HWND hWndCol, long lRow, LPBOOL pbMustDelete );
BOOL WINAPI MTblGetEffCellIndent( HWND hWndCol, long lRow, LPLONG plIndent );
HANDLE WINAPI MTblGetEffCellImage( HWND hWndCol, long lRow );
BOOL WINAPI MTblGetEffCellTextColor( HWND hWndCol, long lRow, LPDWORD pdwColor );
int WINAPI MTblGetEffCellTextJustify( HWND hWndCol, long lRow );
int WINAPI MTblGetEffCellTextVAlign( HWND hWndCol, long lRow );
BOOL WINAPI MTblGetEffColHdrGrpBackColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffColHdrGrpFont( HWND hWndTbl, LPVOID pGrp, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
HFONT WINAPI MTblGetEffColHdrGrpFontHandle( HWND hWndTbl, LPVOID pGrp, LPBOOL pbMustDelete );
BOOL WINAPI MTblGetEffColHdrGrpTextColor( HWND hWndTbl, LPVOID pGrp, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffColumnHdrBackColor( HWND hWndCol, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffColumnHdrFont( HWND hWndCol, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
BOOL WINAPI MTblGetEffColumnHdrFreeBackColor( HWND hWndTbl, LPDWORD pdwColor );
HFONT WINAPI MTblGetEffColumnHdrFontHandle( HWND hWndCol, LPBOOL pbMustDelete );
BOOL WINAPI MTblGetEffColumnHdrTextColor( HWND hWndCol, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffCornerBackColor( HWND hWndTbl, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffHighlighting( HUDV hUdvItem, HUDV hUdvHili );
BOOL WINAPI MTblGetEffRowHdrBackColor( HWND hWndTbl, long lRow, LPDWORD pdwColor );
BOOL WINAPI MTblGetEffRowHeight( HWND hWndTbl, long lRow, LPLONG plHeight );
DWORD WINAPI MTblGetExportSubstColor( HWND hWndTbl, DWORD dwColor );
long WINAPI MTblGetFirstChildRow( HWND hWndTbl, long lRow );
BOOL WINAPI MTblGetFreeColumnHdrBackColor( HWND hWndTbl, LPDWORD pdwColor );
BOOL WINAPI MTblGetHeadersBackColor( HWND hWndTbl, LPDWORD pdwColor );
BOOL WINAPI MTblGetItem( HWND hWndTbl, LPARAM lID, HUDV hUdv );
BOOL WINAPI MTblGetIndentPerLevel( HWND hWndTbl, LPLONG plIndentPerLevel );
BOOL WINAPI MTblGetLanguageFile( int iLanguage, LPHSTRING phsLanguageFile );
BOOL WINAPI MTblGetLanguageString( int iLanguage, LPCTSTR lpsSection, LPCTSTR lpsKey, LPHSTRING phsString );
int WINAPI MTblGetLanguageStringC( int iLanguage, LPCTSTR lpsSection, LPCTSTR lpsKey, LPTSTR lpsString );
BOOL WINAPI MTblGetLastMergedCell( HWND hWndCol, long lRow, LPHWND phWndLastMergedCol, LPLONG plLastMergedRow );
long WINAPI MTblGetLastMergedRow( HWND hWndTbl, long lRow );
long WINAPI MTblGetLastSortDef( HWND hWndTbl, HARRAY hWndaCols, HARRAY naSortFlags );
BOOL WINAPI MTblGetMergeCell( HWND hWndCol, long lRow, LPHWND phWndMergeCol, LPLONG plMergeRow );
LPVOID WINAPI MTblGetMergeCells( HWND hWndTbl, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff );
HWND WINAPI MTblGetMergeCol( HWND hWndCol, long lRow );
long WINAPI MTblGetMergeRow( HWND hWndTbl, long lRow );
UINT WINAPI MTblGetMsgOffset( );
int WINAPI MTblGetMWheelOption( HWND hWndTbl, int iOption );
long WINAPI MTblGetNextChildRow( HWND hWndTbl, long lRow );
BOOL WINAPI MTblGetNodeImages( HWND hWndTbl, LPHANDLE phImg, LPHANDLE phImgExp );
BOOL WINAPI MTblGetNodeRect( HWND hWndTbl, long lCellLeft, long lTextTop, int iFontHeight, int iLevel, int iCellLeadingX, int iCellLeadingY, LPRECT pr );
long WINAPI MTblGetOrigParentRow( HWND hWndTbl, long lRow );
BOOL WINAPI MTblGetOwnerDrawnItem( HWND hWndTbl, LPARAM lID, HUDV hUdv );
long WINAPI MTblGetParentRow( HWND hWndTbl, long lRow );
HSTRING WINAPI MTblGetPasswordChar( HWND hWndTbl );
TCHAR WINAPI MTblGetPasswordCharVal( HWND hWndTbl );
long WINAPI MTblGetPatchLevel( );
long WINAPI MTblGetPrevChildRow( HWND hWndTbl, long lRow );
BOOL WINAPI MTblGetRowBackColor( HWND hWndTbl, long lRow, LPDWORD pdwColor );
BOOL WINAPI MTblGetRowFlagImage( HWND hWndTbl, WORD wRowFlag, LPHANDLE phImage );
BOOL WINAPI MTblGetRowFont( HWND hWndTbl, long lRow, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
long WINAPI MTblGetRowFromID( HWND hWndTbl, LPVOID lpRow );
BOOL WINAPI MTblGetRowHdrBackColor( HWND hWndTbl, long lRow, LPDWORD pdwColor );
BOOL WINAPI MTblGetRowHeight( HWND hWndTbl, long lRow, LPLONG plHeight );
LPVOID WINAPI MTblGetRowID( HWND hWndTbl, long lRow );
BOOL WINAPI MTblGetRowMerge( HWND hWndTbl, long lRow, LPINT piCount );
int WINAPI MTblGetRowLevel( HWND hWndTbl, long lRow );
long WINAPI MTblGetRowSelChanges( HWND hWndTbl, HARRAY haRows, HARRAY haChanges );
BOOL WINAPI MTblGetRowTextColor( HWND hWndTbl, long lRow, LPDWORD pdwColor );
HSTRING WINAPI MTblGetRowUserData( HWND hWndTbl, long lRow, LPTSTR lpsKey );
long WINAPI MTblGetRowVisPos( HWND hWndTbl, long lRow );
BOOL WINAPI MTblGetSelectionColors( HWND hWndTbl, LPDWORD pdwTextColor, LPDWORD pdwBackColor );
BOOL WINAPI MTblGetSelectionDarkening(HWND hWndTbl,  LPBYTE pbySelDarkening);
HWND WINAPI MTblGetSepCol( HWND hWndTbl, int iX, int iY );
long WINAPI MTblGetSepRow( HWND hWndTbl, int iX, int iY );
BOOL WINAPI MTblGetTipBackColor( HWND hWndParam, long lParam, int iType, LPDWORD pdwColor );
BOOL WINAPI MTblGetTipDelay( HWND hWndParam, long lParam, int iType, LPINT piDelay );
BOOL WINAPI MTblGetTipFadeInTime( HWND hWndParam, long lParam, int iType, LPNUMBER pnTime );
BOOL WINAPI MTblGetTipFont( HWND hWndParam, long lParam, int iType, LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
BOOL WINAPI MTblGetTipFrameColor( HWND hWndParam, long lParam, int iType, LPDWORD pdwColor );
BOOL WINAPI MTblGetTipOffset( HWND hWndParam, long lParam, int iType, LPINT piOffsX, LPINT piOffsY );
BOOL WINAPI MTblGetTipOpacity( HWND hWndParam, long lParam, int iType, LPNUMBER pnOpacity );
BOOL WINAPI MTblGetTipPos( HWND hWndParam, long lParam, int iType, LPNUMBER pnX, LPNUMBER pnY );
BOOL WINAPI MTblGetTipRCESize( HWND hWndParam, long lParam, int iType, LPNUMBER pnWid, LPNUMBER pnHt );
BOOL WINAPI MTblGetTipSize( HWND hWndParam, long lParam, int iType, LPNUMBER pnWid, LPNUMBER pnHt );
BOOL WINAPI MTblGetTipText( HWND hWndParam, long lParam, int iType, LPHSTRING phsText );
BOOL WINAPI MTblGetTipTextColor( HWND hWndParam, long lParam, int iType, LPDWORD pdwColor );
BOOL WINAPI MTblGetTreeNodeColors( HWND hWndTbl, LPDWORD pdwOuterClr, LPDWORD pdwInnerClr, LPDWORD pdwBackClr );
HSTRING WINAPI MTblGetVersion( );
long WINAPI MTblInsertChildRow( HWND hWndTbl, LPLONG plParentRow, DWORD dwFlags );
BOOL WINAPI MTblIsCellDisabled( HWND hWndCol, long lRow );
BOOL WINAPI MTblIsExtMsgsEnabled( HWND hWndTbl );
BOOL WINAPI MTblIsHighlighted( HUDV hUdvItem );
BOOL WINAPI MTblIsMWheelScrollEnabled( HWND hWndTbl );
BOOL WINAPI MTblIsOwnerDrawnItem( HWND hWndParam, long lParam, int iType );
BOOL WINAPI MTblIsParentRow( HWND hWndTbl, long lRow );
BOOL WINAPI MTblIsRow( HWND hWndTbl, long lRow );
BOOL WINAPI MTblIsRowDescOf( HWND hWndTbl, long lRow, long lParentRow );
BOOL WINAPI MTblIsRowDisabled( HWND hWndTbl, long lRow );
BOOL WINAPI MTblIsSubClassed( HWND hWndTbl );
BOOL WINAPI MTblIsTipTypeEnabled( HWND hWndTbl, DWORD dwTipType );
long WINAPI MTblLockPaint( HWND hWndTbl );
BOOL WINAPI MTblMustRedrawSelections( HWND hWndTbl );
BOOL WINAPI MTblObjFromPoint( HWND hWndTbl, int iX, int iY, LPLONG plRow, LPHWND phWndCol, LPDWORD pdwFlags, LPDWORD pdwMFlags );
BOOL WINAPI MTblODIPrint( HWND hWndTbl, LPVOID pODI );
HSTRING WINAPI MTblPrintGetDefPrinterName();
long WINAPI MTblPrintGetPrinterNames(HARRAY arr);
int WINAPI MTblPrintSetupDlg(HWND hWndParent, HSTRING hsPrinter);
int WINAPI MTblPrint(HWND hWndTbl, HUDV hParams);
BOOL WINAPI MTblQueryButtonCell( HWND hWndCol, long lRow, LPBOOL pbShowBtn, LPINT piWidth, LPHSTRING phsText, LPHANDLE phImage, LPDWORD pdwFlags );
BOOL WINAPI MTblQueryButtonColumn( HWND hWndCol, LPBOOL pbShowBtn, LPINT piWidth, LPHSTRING phsText, LPHANDLE phImage, LPDWORD pdwFlags );
BOOL WINAPI MTblQueryButtonHotkey( HWND hWndTbl, LPINT piKey, LPBOOL pbCtrl, LPBOOL pbShift, LPBOOL pbAlt );
BOOL WINAPI MTblQueryCellImageFlags( HWND hWndCol, long lRow, DWORD dwFlags );
BOOL WINAPI MTblQueryCellImageExpFlags( HWND hWndCol, long lRow, DWORD dwFlags );
BOOL WINAPI MTblQueryCellFlags( HWND hWndCol, long lRow, DWORD dwFlags );
BOOL WINAPI MTblQueryCellMode( HWND hWndTbl, LPBOOL pbActive, LPDWORD pdwFlags );
BOOL WINAPI MTblQueryCellType( HWND hWndCol, long lRow, LPINT piType, LPDWORD pdwFlags, LPINT piLines, LPHSTRING phsCheckedVal, LPHSTRING phsUnCheckedVal );
BOOL WINAPI MTblQueryColHdrGrpFlags( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags );
BOOL WINAPI MTblQueryColHdrGrpImageFlags( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags );
BOOL WINAPI MTblQueryColLines( HWND hWndTbl, LPINT piStyle, LPDWORD pclr );
BOOL WINAPI MTblQueryColumnFlags( HWND hWndCol, DWORD dwFlags );
BOOL WINAPI MTblQueryColumnHdrFlags( HWND hWndCol, DWORD dwFlags );
BOOL WINAPI MTblQueryColumnHdrImageFlags( HWND hWndCol, DWORD dwFlags );
BOOL WINAPI MTblQueryCornerImageFlags( HWND hWndTbl, DWORD dwFlags );
BOOL WINAPI MTblQueryDropDownListContext( HWND hWndCol, LPLONG plRow );
BOOL WINAPI MTblQueryEffCellImageFlags( HWND hWndCol, long lRow, DWORD dwFlags );
BOOL WINAPI MTblQueryEffCellPropDet( HWND hWndTbl, LPDWORD pdwOrder, LPDWORD pdwFlags );
BOOL WINAPI MTblQueryEffItemLine(HUDV hUdvItem, HUDV hUdvLineDef, int iLineType);
BOOL WINAPI MTblQueryEffItemLineC(MTBLITEM& item, MTBLLINEDEF& LineDef, int iLineType);
BOOL WINAPI MTblQueryExtMsgsOptions( HWND hWndTbl, DWORD dwOptions );
BOOL WINAPI MTblQueryFlags( HWND hWndTbl, DWORD dwFlags );
BOOL WINAPI MTblQueryFocusCell( HWND hWndTbl, LPLONG plRow, LPHWND phWndCol );
BOOL WINAPI MTblQueryFocusFrame( HWND hWndTbl, LPINT piThickness, LPINT piOffsX, LPINT piOffsY, LPDWORD pdwColor );
BOOL WINAPI MTblQueryHighlighting( HWND hWndTbl, long lItem, long lPart, HUDV hUdv );
BOOL WINAPI MTblQueryHyperlinkCell( HWND hWndCol, long lRow, LPBOOL pbEnabled, LPDWORD pdwFlags );
BOOL WINAPI MTblQueryHyperlinkHover( HWND hWndTbl, LPBOOL pbEnabled, LPLONG plAddFontEnh, LPDWORD pdwTextColor );
BOOL WINAPI MTblQueryItemButton( HUDV hUdvItem, HUDV hUdvBtnDef );
BOOL WINAPI MTblQueryItemButtonC( LPVOID pItemV, LPVOID pBtnDefV );
BOOL WINAPI MTblQueryItemLine(HUDV hUdvItem, HUDV hUdvLineDef, int iLineType);
BOOL WINAPI MTblQueryItemLineC(LPVOID pItemV, LPVOID pLineDefV, int iLineType);
BOOL WINAPI MTblQueryLastLockedColLine( HWND hWndTbl, LPINT piStyle, LPDWORD pclr );
int WINAPI MTblQueryLockedColumns( HWND hWndTbl );
BOOL WINAPI MTblQueryRowFlagImageFlags( HWND hWndTbl, DWORD dwRowFlag, DWORD dwFlags );
BOOL WINAPI MTblQueryRowFlags(HWND hWndTbl, long lRow, DWORD dwFlags);
BOOL WINAPI MTblQueryRowLines(HWND hWndTbl, LPINT piStyle, LPDWORD pclr);
BOOL WINAPI MTblQueryTipFlags( HWND hWndParam, long lParam, int iType, DWORD dwFlags );
BOOL WINAPI MTblQueryTree( HWND hWndTbl, LPHWND phWndTreeCol, LPINT piNodeSize, LPINT piIndent );
BOOL WINAPI MTblQueryTreeFlags( HWND hWndTbl, DWORD dwFlags );
BOOL WINAPI MTblQueryTreeLines( HWND hWndTbl, LPINT piStyle, LPDWORD pclr );
BOOL WINAPI MTblRegisterLanguage( int iLanguage, LPTSTR lpsLanguageFile, BOOL bShowErrMsg = TRUE );
BOOL WINAPI MTblResetTip(HWND hWndTbl);
BOOL WINAPI MTblSetAltRowBackColors( HWND hWndTbl, BOOL bSplitRows, DWORD dwColor1, DWORD dwColor2 );
BOOL WINAPI MTblQueryVisibleSplitRange( HWND hWndTbl, LPLONG plMin, LPLONG plMax );
BOOL WINAPI MTblSetCellBackColor( HWND hWndCol, long lRow, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetCellFlags( HWND hWndCol, long lRow, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetCellFont( HWND hWndCol, long lRow, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags );
BOOL WINAPI MTblSetCellIndentLevel( HWND hWndCol, long lRow, long lIndentLevel );
BOOL WINAPI MTblSetCellImage( HWND hWndCol, long lRow, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblSetCellImageExp( HWND hWndCol, long lRow, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblSetCellMerge( HWND hWndCol, long lRow, int iCount, DWORD dwFlags );
BOOL WINAPI MTblSetCellMergeEx( HWND hWndCol, long lRow, int iColCount, int iRowCount, DWORD dwFlags );
BOOL WINAPI MTblSetCellTextColor( HWND hWndCol, long lRow, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColHdrGrpBackColor( HWND hWndTbl, LPVOID pGrp, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColHdrGrpCol( HWND hWndTbl, LPVOID pGrp, HWND hWndCol, BOOL bSet, DWORD dwFlags );
BOOL WINAPI MTblSetColHdrGrpFlags( HWND hWndTbl, LPVOID pGrp, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetColHdrGrpFont( HWND hWndTbl, LPVOID pGrp, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags );
BOOL WINAPI MTblSetColHdrGrpImage( HWND hWndTbl, LPVOID pGrp, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblSetColHdrGrpText( HWND hWndTbl, LPVOID pGrp, LPCTSTR lpsText, DWORD dwFlags );
BOOL WINAPI MTblSetColHdrGrpTextColor( HWND hWndTbl, LPVOID pGrp, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColumnBackColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColumnFlags( HWND hWndCol, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetColumnFont( HWND hWndCol, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags );
BOOL WINAPI MTblSetColumnTextColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColumnTitle( HWND hWndCol, LPCTSTR lpsTitle );
BOOL WINAPI MTblSetColumnHdrBackColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColumnHdrFlags( HWND hWndCol, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetColumnHdrFont( HWND hWndCol, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags );
BOOL WINAPI MTblSetColumnHdrFreeBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetColumnHdrImage( HWND hWndCol, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblSetColumnHdrTextColor( HWND hWndCol, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetCornerBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetCornerImage( HWND hWndTbl, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblSetDropDownListContext( HWND hWndCol, long lRow );
BOOL WINAPI MTblSetExportSubstColor( HWND hWndTbl, DWORD dwColor, DWORD dwExportColor );
BOOL WINAPI MTblSetExtMsgsOptions( HWND hWndTbl, DWORD dwOptions, BOOL bSet );
BOOL WINAPI MTblSetFlags( HWND hWndTbl, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetFocusCell( HWND hWndTbl, long lRow, HWND hWndCol, DWORD dwFlags );
BOOL WINAPI MTblSetFreeColumnHdrBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetHeadersBackColor( HWND hWndTbl, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetHighlighted( HUDV hUdvItem, BOOL bSet );
BOOL WINAPI MTblSetIndentPerLevel( HWND hWndTbl, long lIndentPerLevel );
BOOL WINAPI MTblSetMsgOffset( UINT uiOffs );
BOOL WINAPI MTblSetMWheelOption( HWND hWndTbl, int iOption, int iValue );
BOOL WINAPI MTblSetNodeImages( HWND hWndTbl, HANDLE hImg, HANDLE hImgExp, DWORD dwFlags );
BOOL WINAPI MTblSetOwnerDrawnItem( HWND hWndParam, long lParam, int iType, BOOL bSet );
BOOL WINAPI MTblSetParentRow( HWND hWndTbl, long lRow, long lParentRow, DWORD dwFlags );
BOOL WINAPI MTblSetPasswordChar( HWND hWndTbl, LPCTSTR lpsPwd );
BOOL WINAPI MTblSetRowBackColor( HWND hWndTbl, long lRow, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetRowFlagImage( HWND hWndTbl, WORD wRowFlag, HANDLE hImage, DWORD dwFlags );
BOOL WINAPI MTblSetRowFlags(HWND hWndTbl, long lRow, DWORD dwRowFlags, BOOL bSet, DWORD dwFlags);
BOOL WINAPI MTblSetRowFont( HWND hWndTbl, long lRow, LPTSTR lpsName, long lSize, long lEnh, DWORD dwFlags );
BOOL WINAPI MTblSetRowHdrBackColor( HWND hWndTbl, long lRow, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetRowHeight( HWND hWndTbl, long lRow, long lHeight, DWORD dwFlags );
BOOL WINAPI MTblSetRowMerge( HWND hWndTbl, long lRow, int iCount, DWORD dwFlags );
BOOL WINAPI MTblSetRowTextColor( HWND hWndTbl, long lRow, DWORD dwColor, DWORD dwFlags );
BOOL WINAPI MTblSetRowUserData(HWND hWndTbl, long lRow, LPTSTR lpsKey, LPTSTR lpsValue);
BOOL WINAPI MTblSetSelectionColors( HWND hWndTbl, DWORD dwTextColor, DWORD dwBackColor );
BOOL WINAPI MTblSetSelectionDarkening(HWND hWndTbl, BYTE byDarkening);
BOOL WINAPI MTblSetTipBackColor( HWND hWndParam, long lParam, int iType, DWORD dwColor );
BOOL WINAPI MTblSetTipDelay( HWND hWndParam, long lParam, int iType, NUMBER nDelay );
BOOL WINAPI MTblSetTipFadeInTime( HWND hWndParam, long lParam, int iType, NUMBER nTime );
BOOL WINAPI MTblSetTipFlags( HWND hWndParam, long lParam, int iType, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetTipFont( HWND hWndParam, long lParam, int iType, LPTSTR lpsName, long lSize, long lEnh );
BOOL WINAPI MTblSetTipFrameColor( HWND hWndParam, long lParam, int iType, DWORD dwColor );
BOOL WINAPI MTblSetTipOffset( HWND hWndParam, long lParam, int iType, NUMBER nOffsX, NUMBER nOffsY );
BOOL WINAPI MTblSetTipOpacity( HWND hWndParam, long lParam, int iType, NUMBER nOpacity );
BOOL WINAPI MTblSetTipPos( HWND hWndParam, long lParam, int iType, NUMBER nX, NUMBER nY );
BOOL WINAPI MTblSetTipRCESize( HWND hWndParam, long lParam, int iType, NUMBER nWid, NUMBER nHt );
BOOL WINAPI MTblSetTipSize( HWND hWndParam, long lParam, int iType, NUMBER nWid, NUMBER nHt );
BOOL WINAPI MTblSetTipText( HWND hWndParam, long lParam, int iType, LPCTSTR lpsText );
BOOL WINAPI MTblSetTipTextColor( HWND hWndParam, long lParam, int iType, DWORD dwColor );
BOOL WINAPI MTblSetTreeFlags( HWND hWndTbl, DWORD dwFlags, BOOL bSet );
BOOL WINAPI MTblSetTreeNodeColors( HWND hWndTbl, DWORD dwOuterClr, DWORD dwInnerClr, DWORD dwBackClr );
BOOL WINAPI MTblSort( HWND hWndTbl, HARRAY hWndaCols, HARRAY naFlags );
BOOL WINAPI MTblSortEx( HWND hWndTbl, HARRAY hWndaCols, HARRAY naFlags, long lMinRow, long lMaxRow, long lFromLevel, long lToLevel, long lParentRow );
BOOL WINAPI MTblSortRange( HWND hWndTbl, HARRAY hWndaCols, HARRAY naFlags, long lMinRow, long lMaxRow );
BOOL WINAPI MTblSortTree( HWND hWndTbl, DWORD dwFlags );
BOOL WINAPI MTblSubClass( HWND hWndTbl );
BOOL WINAPI MTblSwapRows( HWND hWndTbl, long lRow1, long lRow2, DWORD dwFlags );
long WINAPI MTblUnlockPaint( HWND hWndTbl );
}

#endif // _INC_MTBLAPI