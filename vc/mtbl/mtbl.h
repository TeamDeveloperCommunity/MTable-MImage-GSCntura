// mtbl.h

#ifndef _INC_MTBL
#define _INC_MTBL

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <math.h>
#include <time.h>
#include <algorithm>
#include <array>
#include <list>
#include <map>
#include <set>
#include <stack>
#include <vector>
#include "cmtblcol.h"
#include "cmtblcolhdrgrp.h"
#include "cmtblcolhdrgrps.h"
#include "cmtblcols.h"
#include "cmtblrow.h"
#include "cmtblrows.h"
#include "cmtblimg.h"
#include "cmtblmetrics.h"
#include "ctd.h"
#include "vis.h"
#include "resource.h"
#include "mtblapi.h"
#include "rowflags.h"
#include "cmtbltip.h"
#include "mimgapi.h"
#include "calphawnd.h"
#include "ctheme.h"
#include "fileversion.h"
#include "tstring.h"
#include "cmtblcounter.h"
#include "cmtblcoltag.h"
#include "cmtblsortval.h"
#include "gradrect.h"
#include "ccurcols.h"

#include <commctrl.h>
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' \
version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#ifdef USE_APIHOOK
#include "apihook.h"
#endif // USE_APIHOOK

// Allgemeine Konstanten
//#define MTBL_VERSION _T("2.1.6")
#define MTBL_PATCH_LEVEL 0

// Node-Typen
#define MTBL_NODE_PLUS 0
#define MTBL_NODE_MINUS 1

// Subclassing
#define MTBL_CONTAINER_SUBCLASS_PROP _T("MTblListClipContainerSubClass")
#define MTBL_LISTCLIP_CTRL_SUBCLASS_PROP _T("MTblListClipCtrlSubClass")
#define MTBL_LISTCLIP_SUBCLASS_PROP _T("MTblListClipSubClass")

// Fensterklassen
#define MTBL_BTN_CLASS _T("MTblBtn")
#define MTBL_EMPTYCOLWND_CLASS _T("MTblEmptyCol")
#define MTBL_TIMERWND_CLASS _T("MTblTimer")
#define MTBL_TIPWND_CLASS _T("MTblTip")

// Maximale Anzahl Spalten
#define MTBL_MAX_COLS 16384
// Buttons
#define MTBL_DDBTN_WID 16
#define MTBL_DDBTN_WID_IS 24
#define MTBL_BTN_ID 1111

// Gepunktete Linien
#define MTBL_DOTLINE_ALL 0
#define MTBL_DOTLINE_ROW 1
#define MTBL_DOTLINE_COLUMN 2
#define MTBL_DOTLINE_LASTLOCKEDCOL 3

#define MTBL_DOTLINE_LEN_ROW 100
#define MTBL_DOTLINE_LEN_COL 100

// Normalisierung Hierarchie
#define NORM_HIER_ROW_DELETE 1
#define NORM_HIER_ROW_HIDE 2
#define NORM_HIER_ROW_SHOW 3
#define NORM_HIER_ROW_DELETE_FAIL 4

// GetCellMergeRowRange
#define GCMRR_UP 1
#define GCMRR_DOWN 2

// GetDefaultColor
// TblGetDefaultColor - Konstanten
#define GDC_HDR_BKGND 1
#define GDC_HDR_LINE 2
#define GDC_HDR_3DHL 3
#define GDC_CELL_LINE 4

// Flags GetEff*Color
#define GEC_HYPERLINK 1
#define GEC_EMPTY_COL 2
#define GEC_EMPTY_ROW 4
#define GEC_HILI 8
#define GEC_HILI_PD 16

// Flags GetEff*Font
#define GEF_HYPERLINK 1
#define GEF_HILI 8
#define GEF_HILI_PD 16

// Flags GetEff*Image
#define GEI_HILI 8
#define GEI_HILI_PD 16

// GetSel* - Modus-Flags
#define GETSEL_BEFORE 1
#define GETSEL_AFTER 2
#define GETSEL_REMOVE_NEW 4

// GetFocusRgn
#define GFR_OUTER 0
#define GFR_INNER 1
#define GFR_FRAME 2

// Aktion MTM_Internal
#define IA_REMOVE_FOCUS_ROW_SELECTION 1
#define IA_KILL_EDIT_ESC 2
#define IA_REMOVE_FOCUS 3
#define IA_SUBCLASS_LISTCLIP_CTRL 4
#define IA_ADAPT_LISTCLIP 8
#define IA_CHECK_LIST_OPEN 16
#define IA_UPDATE_WINDOW 32
#define IA_AFTER_SCROLL 64
#define IA_GET_SPLITTED 128

// Typen Merge-Zelle
#define MC_NONE 0
#define MC_SIMPLE 1
#define MC_EXTENDED 2

// Aktionen Hook
#define HA_NONE 0
#define HA_MAKE_LEFT_EDIT 1

// UpdateFocus
#define UF_REDRAW_AUTO 0
#define UF_REDRAW_INVALIDATE 1
#define UF_REDRAW_NO_INVALIDATE 2

#define UF_SELCHANGE_AUTO 0
#define UF_SELCHANGE_FOCUS_ONLY 1
#define UF_SELCHANGE_FOCUS_CELL_ONLY 2
#define UF_SELCHANGE_FOCUS_ADD 3
#define UF_SELCHANGE_FOCUS_ADD_KEY_RANGE 4
#define UF_SELCHANGE_FOCUS_ADD_RANGE 5
#define UF_SELCHANGE_FOCUS_SWITCH 6
#define UF_SELCHANGE_CLEAR_ALL 7
#define UF_SELCHANGE_CLEAR_ROWS_AND_COLS 8
#define UF_SELCHANGE_CLEAR_CELLS 9
#define UF_SELCHANGE_NONE 10

#define UF_FLAG_SCROLL 1
#define UF_FLAG_SELCHANGE_ROWS 2

// Properties
#define MTPROP_VAL_BOOL_NO _T("0")
#define MTPROP_VAL_BOOL_YES _T("1")

// Windows-Definitionen
#ifndef WM_MOUSEWHEEL
#define WM_MOUSEWHEEL 0x020A
#endif

#ifndef WHEEL_DELTA
#define WHEEL_DELTA 120
#endif

#ifndef CS_DROPSHADOW
#define CS_DROPSHADOW 0x00020000
#endif

#ifndef SM_XVIRTUALSCREEN
#define SM_XVIRTUALSCREEN 76
#endif

#ifndef SM_YVIRTUALSCREEN
#define SM_YVIRTUALSCREEN 77
#endif

#ifndef SM_CXVIRTUALSCREEN
#define SM_CXVIRTUALSCREEN 78
#endif

#ifndef SM_CYVIRTUALSCREEN
#define SM_CYVIRTUALSCREEN 79
#endif

// Using
using namespace std;

// Typedefs
typedef map< DWORD, tstring> LanguageMap;
typedef list<long> LongList;
typedef LongList *PLongList;
typedef list<HWND> HwndList;
typedef HwndList *PHwndList;
typedef list<POINT> PointList;
typedef PointList *PPointList;
typedef list<RECT> RectList;
typedef RectList *PRectList;
typedef map< DWORD, pair<HICON,HICON> > RowFlagIconMap;
typedef RowFlagIconMap *PRowFlagIconMap;
typedef map< DWORD, pair<CMTblImg*,CMTblImg*> > RowFlagImgMap;
typedef RowFlagImgMap *PRowFlagImgMap;
typedef map< HWND, LPVOID > SubClassMap;
typedef map< COLORREF, COLORREF > SubstColorMap;
typedef pair< long, long > RowRange;
typedef pair< CMTblRow*, CMTblRow* > RowPtrs;
typedef pair< RowPtrs, int > RowPtrsInt;
typedef pair< CMTblRow**, long > RowArrLong;
typedef list < CMTblRow* > RowPtrList;
typedef stack < CMTblRow* > RowPtrStack;
typedef vector < CMTblRow* > RowPtrVector;
typedef vector < RowPtrsInt > RowPtrsIntVector;
typedef vector < RowArrLong > RowArrLongVector;
typedef vector < RowRange > RowRangeVector;

// Globale Variablen
/*
BOOL gbCtdHooked;
HANDLE ghInstance;
LanguageMap glm;
SubClassMap gscm;
UINT guiMsgOffs = 100;
tstring gsEmpty( _T("") );

// Funktionscontainer
CCtd Ctd;
CVis Vis;
*/

// Strukturen

// MTBLDOTLINE
typedef struct MTBLDOTLINE
{
	HDC hDC;
	RECT rCoord;
	COLORREF clr1;
	COLORREF clr2;
	BOOL bVert;
	BOOL bOdd;
}
MTBLDOTLINE, *LPMTBLDOTLINE;

// MTBLSWAPROWS
typedef struct MTBLSWAPROWS
{
	TDLONG lRow1;
	TDLONG lRow2;
}
MTBLSWAPROWS, * LPMTBLSWAPROWS;

// MTBLSETRANGE
typedef struct MTBLSETRANGE
{
	long l1;
	long l2;
	long l3;
}
MTBLSETRANGE, * LPMTBLSETRANGE;

// MTBLBTN
typedef struct MTBLBTN
{
	HWND hWndTbl;
	BOOL bPushed;
	BOOL bCaptured;
	LPCTSTR lpsText;
	HANDLE hImage;
	DWORD dwFlags;
	int iHeight;
}
MTBLBTN, * LPMTBLBTN;

// MTBLBTNDEFUDV
#pragma pack(UDV_PACK)
typedef struct MTBLBTNDEFUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	NUMBER Visible;
	NUMBER Width;
	HSTRING Text;
	NUMBER Image;
	NUMBER Flags;
}
MTBLBTNDEFUDV, *LPMTBLBTNDEFUDV;
#pragma pack()

// MTBLBTNDEF
typedef struct MTBLBTNDEF
{
	BOOL bVisible;
	int iWidth;
	tstring sText;
	HANDLE hImage;
	DWORD dwFlags;

	BOOL FromUdv( HUDV hUdv );
	BOOL ToUdv( HUDV hUdv );
}
MTBLBTNDEF, *LPMTBLBTNDEF;


// MTBLFSC
typedef struct MTBLFSC
{
	LPTSTR lpsClass;
	HWND hWndFound;
}
MTBLFSC, * LPMTBLFSC;

// MTBLHILI ( HILI = Highlighting )
typedef struct MTBLHILI
{
	DWORD dwBackColor;
	DWORD dwTextColor;
	LONG lFontEnh;
	CMTblImg Img;
	CMTblImg ImgExp;
	//HANDLE hImage;
	//HANDLE hImageExp;

	void Init( )
	{
		dwBackColor = dwTextColor = MTBL_COLOR_UNDEF;
		lFontEnh = MTBL_FONT_UNDEF_ENH;

		Img.SetHandle( NULL, NULL );
		Img.ClearFlags( );

		ImgExp.SetHandle( NULL, NULL );
		ImgExp.ClearFlags( );
	}

	BOOL IsUndef( )
	{
		return dwBackColor == MTBL_COLOR_UNDEF && dwTextColor == MTBL_COLOR_UNDEF && lFontEnh == MTBL_FONT_UNDEF_ENH && !Img.GetHandle( ) && !ImgExp.GetHandle( );
	}
}
MTBLHILI, *LPMTBLHILI;

typedef map< LONG_PAIR, MTBLHILI > HiLiDefMap;

// MTBLHILIUDV
#pragma pack(UDV_PACK)
typedef struct MTBLHILIUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	NUMBER BackColor;
	NUMBER TextColor;
	NUMBER FontEnh;
	NUMBER Image;
	NUMBER ImageExp;
}
MTBLHILIUDV, *LPMTBLHILIUDV;
#pragma pack()


// MTBLODI ( ODI = OwnerDrawnItem )
typedef struct MTBLODI
{
	long lType;
	HWND hWndParam;
	long lParam;
	HWND hWndPart;
	long lPart;
	HDC hDC;
	RECT r;
	long lSelected;
	long lXLeading;
	long lYLeading;
	long lPrinting;
	double dYResFactor;
	double dXResFactor;
}
MTBLODI, *LPMTBLODI;

// MTBLODIUDV
#pragma pack(UDV_PACK)
typedef struct MTBLODIUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	NUMBER Type;
	HWND HWndParam;
	NUMBER NParam;
	HWND HWndPart;
	NUMBER NPart;
	NUMBER DC;
	NUMBER Left;
	NUMBER Top;
	NUMBER Right;
	NUMBER Bottom;
	NUMBER Selected;
	NUMBER XLeading;
	NUMBER YLeading;
	NUMBER Printing;
	NUMBER YResFactor;
	NUMBER XResFactor;
}
MTBLODIUDV, *LPMTBLODIUDV;
#pragma pack()

// MTBLPAINTBTN
typedef struct MTBLPAINTBTN
{
	HDC hDC;
	RECT rBtn;
	LPCTSTR lpsText;
	HANDLE hFont;
	HANDLE hImage;
	DWORD dwFlags;
	BOOL bPushed;
	BOOL bDrawInverted;
	BOOL bRestoreClipRgn;
}
MTBLPAINTBTN, *LPMTBLPAINTBTN;

// MTBLPAINTDDBTN
typedef struct MTBLPAINTDDBTN
{
	HDC hDC;
	RECT rBtn;
	BOOL bPushed;
}
MTBLPAINTDDBTN, *LPMTBLPAINTDDBTN;

// MTBLPAINTCELL
typedef struct MTBLPAINTCELL
{
	// Zelle allgemein
	RECT rCell;					// Rechteck der Zelle
	RECT rCellVis;				// Sichtbares Rechteck der Zelle
	DWORD dwTopLog;				// Logische Oberkante
	DWORD dwBottomLog;			// Logische Unterkante
	COLORREF crBackColor;		// Hintergrundfarbe
	BOOL bMerge;				// Merge-Zelle?
	CMTblColHdrGrp *pColHdrGrp;	// Gruppe Spaltenheader
	BOOL bGroupColHdr;			// Gruppe Spaltenheader?
	BOOL bGroupColHdrTop;		// Oberer Bereich Spaltenheader-Gruppe
	BOOL bGroupColHdrNCH;		// Spaltenheader-Gruppe ohne Column-Header
	BOOL bFirstGroupColHdr;		// Erste Spalte einer Spaltenheader-Gruppe?
	BOOL bLastGroupColHdr;		// Letzte Spalte einer Spaltenheader-Gruppe?
	int iODIType;				// OwnerDrawnItem-Typ, 0 = <kein>
	int iCellLeadingY;			// Y-Abstand Zellinhalt
	struct tagBits
	{
		UINT Selected: 1;				// "Selektiert" ja/nein
		UINT NoSelInvBkgnd : 1;			// Hintergrund bei Selektion nicht invertieren?
		UINT NoSelInvText : 1;			// Text bei Selektion nicht invertieren?
		UINT NoSelInvImage : 1;			// Bild bei Selektion nicht invertieren?
		UINT NoSelInvNodes : 1;			// Knoten bei Selektion nicht invertieren?
		UINT NoSelInvTreeLines : 1;		// Tree-Linien bei Selektion nicht invertieren?
		UINT : 0;						// Für Speicher-Ausrichtung
	} Bits;

	// Identifikation
	MTBLITEM Item;				// Item-Definition
	BOOL bCell;					// Zelle?
	BOOL bRealCell;				// Echte Zelle?
	BOOL bColHdr;				// Spaltenkopf?
	BOOL bColHdrGrp;			// Spaltenkopfgruppe?
	BOOL bRealColHdr;			// Echter Spaltenkopf?
	BOOL bRowHdr;				// Zeilenkopf?
	BOOL bRealRowHdr;			// Echter Zeilenkopf
	BOOL bCorner;				// Ecke?

	// Highlighting
	MTBLHILI Hili;				// Highlighting

	// Linien
	BOOL bColLine;				// Spaltenlinie zeichnen ja/nein
	BOOL bRowLine;				// Zeilenlinie zeichnen ja/nein
	MTBLITEM HorzLineItem;		// Item für horizontale Linie
	MTBLITEM VertLineItem;		// Item für vertikale Linie

	// Tree
	BOOL bTreeBottomUp;			// Baum von unten nach oben
	BOOL bHorzTreeLine;			// Horizontale Tree-Linie ja/nein
	BOOL bVertTreeLineToParent;	// Vertikale Tree-Linie zur Elternzeile ja/nein;
	BOOL bVertTreeLineToChild;	// Vertikale Tree-Linie zur Kindzeile ja/nein;

	// Text
	HSTRING hsText;				// Text 
	LPTSTR lpsText;				// Text
	long lTextLen;				// Textlänge
	HFONT hTextFont;			// Textfont
	long lTextAreaLeft;			// Textbereich links
	long lTextAreaRight;		// Textbereich rechts
	RECT rText;					// Rechteck des Textes
	DWORD dwTextDrawFlags;		// DrawText() - Flags
	COLORREF crTextColor;		// Textfarbe
	BOOL bTextHidden;			// Text versteckt?
	BOOL bTextDisabled;			// Text disabled?

	// Node ( Knoten )
	BOOL bNode;					// Knoten zeichnen ja/nein
	RECT rNode;					// Rechteck des Knotens
	int iNodeType;				// Knotentyp
	COLORREF crNodeColor;		// Knotenfarbe
	COLORREF crNodeInnerColor;	// Knotenfarbe innen
	COLORREF crNodeBackColor;	// Knotenfarbe Hintergrund
	HANDLE hNodeImg;			// Knotenbild

	// Bild
	CMTblImg Img;				// Bild
	RECT rImg;					// Rechteck des Bildes
	DWORD dwImgDrawFlags;		// Zeichnungs-Flags

	// Zelle
	CMTblCol *pCell;			// Zelle
	CMTblCellType *pCellType;	// Zellentyp

	// Funktionen
	int GetCellType( ) { return pCellType ? pCellType->m_iType : COL_CellType_Standard; }
	DWORD GetCellTypeFlags( ) { return pCellType ? pCellType->m_dwFlags : 0; }
	BOOL IsCellTypeCheckedVal( LPCTSTR lps );
}
MTBLPAINTCELL, *LPMTBLPAINTCELL;

// MTBLPAINTCOL
typedef struct MTBLPAINTCOL
{
	HWND hWnd;					// Handle
	int iPos;					// Position
	int iID;					// ID
	BOOL bPaint;				// Zeichnen?
	BOOL bLocked;				// Verankert?
	BOOL bLastLocked;			// Letzte Verankerte?
	BOOL bFirstUnlocked;		// Erste nicht Verankerte?
	BOOL bSelected;				// Selektiert?
	BOOL bWordWrap;				// Word Wrap?
	BOOL bInvisible;			// Invisible?
	BOOL bIndicateOverflow;		// Überlauf anzeigen?
	BOOL bFirstVisible;			// Erste sichtbare Spalte?
	BOOL bLastVisible;			// Letzte sichtbare Spalte?
	BOOL bEmpty;				// Leere Spalte ( Bereich rechts neben der letzten echten Spalte )?
	BOOL bRowHdr;				// Row-Header?
	BOOL bPseudo;				// Pseudo-Spalte?
	BOOL bOddLine;				// Linie ( wenn gepunktet ) um 1 Pixel verschoben?
	long lJustify;				// Ausrichtung
	int iCellType;				// Zellentyp
	DWORD dwCBFlags;			// Check-Box-Flags
	tstring sCBCheckedVal;		// Check-Box Checked-Wert
	tstring sCBUncheckedVal;		// Check-Box Unchecked-Wert
	POINT ptX;					// X-Koordinaten
	POINT ptLogX;				// X-Koordinaten logisch
	POINT ptVisX;				// X-Koordinaten sichtbar
	CMTblCol *pCol;				// Spalte
	CMTblColHdrGrp *pHdrGrp;	// Header-Gruppe
}
MTBLPAINTCOL, *LPMTBLPAINTCOL;

// MTBLPAINTMERGECELL
typedef struct MTBLPAINTMERGECELL
{
	int iType;

	LPMTBLPAINTCOL pColFrom;
	CMTblRow *pRowFrom;

	LPMTBLPAINTCOL pColTo;
	CMTblRow *pRowTo;

	int iMergedCols;
	int iMergedRows;

	DWORD dwTopLog;
	DWORD dwBottomLog;

	RECT r;
}
MTBLPAINTMERGECELL, *LPMTBLPAINTMERGECELL;
typedef vector<LPMTBLPAINTMERGECELL> PaintMergeCellsVector;

// MTBLPAINTMERGEROWHDR
typedef struct MTBLPAINTMERGEROWHDR
{
	LPMTBLPAINTCOL pCol;

	CMTblRow *pRowFrom;
	CMTblRow *pRowTo;

	int iMergedRows;

	DWORD dwTopLog;
	DWORD dwBottomLog;

	RECT r;
}
MTBLPAINTMERGEROWHDR, *LPMTBLPAINTMERGEROWHDR;
typedef vector<LPMTBLPAINTMERGEROWHDR> PaintMergeRowHdrsVector;

// MTBLPAINTDATA
typedef struct MTBLPAINTDATA
{
	HDC hDC;									// Handle des Gerätekontexts, in den gezeichnet wird
	HDC hDCTbl;									// Handle Gerätekontext Tabelle
	HDC hDCDoubleBuffer;						// Handle Gerätekontext Double-Buffer
	HBITMAP hBitmDoubleBuffer;					// Bitmap Double-Buffer
	HBITMAP hOldBitmDoubleBuffer;				// Alte Bitmap Double-Buffer
	HBRUSH hBrTblBack;							// Pinsel mit der Hintergrundfarbe der Tabelle
	CMTblMetrics tm;							// Tabellenmaße
	RECT rUpd;									// Update-Rechteck
	COLORREF clrTblBackColor;					// Hintergrundfarbe der Tabelle
	COLORREF clrTblTextColor;					// Textfarbe der Tabelle
	COLORREF clrHdrBack;						// Hintergrundfarbe Header
	COLORREF clrHdrLine;						// Linienfarbe Header
	COLORREF clr3DHiLight;						// 3D HiLight-Farbe
	long lRow;									// Zeile
	long lEffPaintRow;							// Effektiv gezeichnete Zeile
	CMTblRow *pRow;								// Zeilenpointer
	CMTblRow *pEffPaintRow;						// Effektiv gezeichnete Zeile
	CMTblRow *pFirstMergeRow;					// Pointer erste Merge-Zeile
	CMTblRow *pLastMergeRow;					// Pointer letzte Merge-Zeile
	int iRowLevel;								// Zeilenebene
	BOOL bRowSelected;							// Zeile markiert?
	LPMTBLPAINTCOL PaintCols[MTBL_MAX_COLS];	// Zu zeichnende Spalten
	LPMTBLPAINTCOL pPaintCol;					// Aktuelle Spalte
	LPMTBLPAINTCOL pEffPaintCol;				// Effektiv gezeichnete Spalte
	LPMTBLPAINTCOL pFirstMergeCol;				// Erste Merge-Spalte( Master )
	LPMTBLPAINTCOL pLastMergeCol;				// Letzte Merge-Spalte
	LPMTBLPAINTCOL pFirstGroupColHdr;			// Erster Gruppen-Spaltenheader( Master )
	LPMTBLPAINTCOL pLastGroupColHdr;			// Letzter Gruppen-Spaltenheader
	int iCols;									// Anzahl Spalten
	int iNormalCols;							// Anzahl normale Spalten
	MTBLPAINTCELL pc;							// Zellendaten
	PaintMergeCellsVector PaintMergeCells;		// Zu zeichnende Merge-Zellen
	LPMTBLPAINTMERGECELL ppmc;					// Zu zeichnende Merge-Zelle
	PaintMergeRowHdrsVector PaintMergeRowHdrs;	// Zu zeichnende Merge-Rowheader
	LPMTBLPAINTMERGEROWHDR ppmrh;				// Zu zeichnender Merge-Rowheader
	LPMTBLMERGEROW pmr;							// Merge-Zeile
	long lMinSplitRow;							// Niedrigste sichtbare Split-Zeile
	long lMaxSplitRow;							// Höchste sichtbare Split-Zeile
	long lFocusRow;								// Focus-Zeile
	//long lAboveFocusRow;						// Zeile über Focus-Zeile
	//long lBelowFocusRow;						// Zeile unter Focus-Zeile
	HWND hWndFocusCol;							// Focus-Zelle
	BOOL bColHdr;								// Spaltenheader?
	BOOL bEmptyRow;								// Leere Zeile?
	BOOL bGrayedHeaders;						// Graue Header?
	BOOL bSplit;								// Tabelle geteilt?
	BOOL bNoRowLines;							// Keine Zeilenlinien?
	BOOL bSuppressLastColLine;					// Letzte Spaltenlinie unterdrücken?
	BOOL bSplitRow;								// Split-Zeile?
	BOOL bAnyRow;								// Mind. 1 Zeile?
	BOOL bAnySplitRow;							// Mind. 1 Split-Zeile?
	WORD wRowHdrFlags;							// Row-Header Flags
	BOOL bRowHdrMirror;							// Spiegelung einer Spalte im Row-Header?
	HWND hWndRowHdrCol;							// Row-Header Spiegelspalte
	BOOL bRowHdrColWordWrap;					// WordWrap Row-Header Spiegelspalte
	long lRowHdrColJustify;						// Ausrichtung Row-Header Spiegelspalte
	tstring sRowHdrTitle;						// Row-Header Titel
	int iScrPosV;								// Vert. Scroll-Position
	int iScrPosVSplit;							// Vert. Scroll-Position Split-Bereich
	DWORD dwRowTop;								// Obere Position der Zeile
	DWORD dwRowBottom;							// Untere Position der Zeile
	DWORD dwRowTopLog;							// Logische obere Position der Zeile
	DWORD dwRowBottomLog;						// Logische untere Position der Zeile
	DWORD dwRowTopVis;							// Sichtbare obere Position der Zeile
	DWORD dwRowBottomVis;						// Sichtbare untere Position der Zeile
	HTHEME hTheme;								// Theme-Handle, falls App themed
	BOOL bAnySelColors;							// Irgendeine Selektionsfarbe gesetzt?	
	BOOL bClipFocusArea;						// Focus-Bereich "ausschneiden"?
	BOOL bPaintFocus;							// Focus zeichnen?
	int iPaintColFocusCell;						// Index Paint-Spalte der Focus-Zelle
	RECT rFocusClip;							// Clipping-Rechteck Focus

	MTBLPAINTDATA( );
	LPMTBLPAINTCOL EnsureValidPaintCol( int iIndex );
}
MTBLPAINTDATA, *LPMTBLPAINTDATA;

// MTBLCOLSUBCLASS
typedef struct MTBLCOLSUBCLASS
{
	HWND hWndTbl;
	WNDPROC wpOld;
	BOOL bFixingJustify;
	BOOL bSettingCellType;
}
MTBLCOLSUBCLASS, *LPMTBLCOLSUBCLASS;

// MTBLEXTEDITSUBCLASS
typedef struct MTBLEXTEDITSUBCLASS
{
	HWND hWndCol;
	WNDPROC wpOld;
}
MTBLEXTEDITSUBCLASS, *LPMTBLEXTEDITSUBCLASS;

// MTBLCOLINFO
typedef struct MTBLCOLINFO
{
	HWND hWnd;
	int iID;
	long lType;
	long lWidth;
	long lJustify;
	CMTblColHdrGrp *pColHdrGrp;
	BOOL bColHdrGrpCols;
} MTBLCOLINFO, *LPMTBLCOLINFO;

// MTBLCOLINFO2
typedef struct MTBLCOLINFO2
{
	HWND hWnd;
	int iPos;
	int iCellType;
	BOOL bLocked;
} MTBLCOLINFO2, *LPMTBLCOLINFO2;

// MTBLERTHPARAMS ( ERTH = Export Row To HTML )
typedef struct MTBLERTHPARAMS
{
	HANDLE hFile;
	long lRow;
	WORD wRowFlagsOn;
	WORD wRowFlagsOff;
	DWORD dwHTMLFlags;
	DWORD dwExpFlags;
	int iHt;
	LPMTBLCOLINFO Cols;
	int iCols;
	CMTblColTag *pColTag;
	LPMTBLMERGECELLS pMergeCells;

} MTBLERTHPARAMS, *LPMTBLERTHPARAMS;

// MTBLCELLRANGE
typedef struct MTBLCELLRANGE
{
	HWND hWndColFrom;
	long lRowFrom;

	HWND hWndColTo;
	long lRowTo;

	void CopyTo( MTBLCELLRANGE &cr );
	int GetColFromPos( ){return Ctd.TblQueryColumnPos( hWndColFrom );}
	int GetColToPos( ){return Ctd.TblQueryColumnPos( hWndColTo );}
	int GetDiff( HWND hWndTbl, MTBLCELLRANGE &cr1, MTBLCELLRANGE &cr2, MTBLCELLRANGE craDiff[4] );
	BOOL IsEqualTo( MTBLCELLRANGE &cr );
	BOOL IsEmpty( ){return hWndColFrom == NULL && lRowFrom == TBL_Error && hWndColTo == NULL && lRowTo == TBL_Error;}
	BOOL IsValid( );
	void Normalize( );
	void SetEmpty( ){hWndColFrom = NULL; lRowFrom = TBL_Error; hWndColTo = NULL; lRowTo = TBL_Error;}
} MTBLCELLRANGE;

typedef MTBLCELLRANGE * LPMTBLCELLRANGE;

// MTBLENUM
typedef struct MTBLENUM
{
	HARRAY ha;
	long lLower;
	long lItems;
} MTBLENUM;

typedef MTBLENUM * LPMTBLENUM;

// MTBLROWRANGE
typedef struct MTBLROWRANGE
{
	long lRowFrom;
	long lRowTo;

	void CopyTo( MTBLROWRANGE &rr ){rr.lRowFrom = lRowFrom;	rr.lRowTo = lRowTo;};
	BOOL IsEqualTo( MTBLROWRANGE &rr ){return rr.lRowFrom == lRowFrom && rr.lRowTo == lRowTo;};
	BOOL IsEmpty( ){return lRowFrom == TBL_Error && lRowTo == TBL_Error;}
	BOOL IsValid( );
	void Normalize( );
	void SetEmpty( ){lRowFrom = TBL_Error; lRowTo = TBL_Error;}
} MTBLROWRANGE;

typedef MTBLROWRANGE * LPMTBLROWRANGE;

// SORTCOLINFO
typedef struct SORTCOLINFO
{
	HWND hWnd;
	int iID;
	int iDataType;
	int iColDataType;
	int iSortOrder;
	BOOL bIgnoreCase;
	BOOL bStringSort;
	BOOL bQuery;
}
SORTCOLINFO;

// SORTROW
typedef struct SORTROW
{
	CMTblSortVal *pSortVals;
	SORTCOLINFO *pSortColInfos;
	long lCols;
	long lRow;
}
SORTROW;

struct MTBLSUBCLASS;
typedef MTBLSUBCLASS *LPMTBLSUBCLASS;

// MTBLCONTAINERSUBCLASS
typedef struct MTBLCONTAINERSUBCLASS
{
	HWND hWndTbl;
	WNDPROC wpOld;
}
MTBLCONTAINERSUBCLASS, *LPMTBLCONTAINERSUBCLASS;

// MTBLLISTCLIPSUBCLASS
typedef struct MTBLLISTCLIPSUBCLASS
{
	HWND hWndTbl;
	WNDPROC wpOld;
}
MTBLLISTCLIPSUBCLASS, *LPMTBLLISTCLIPSUBCLASS;

// MTBLLISTCLIPCONTSUBCLASS
typedef struct MTBLLISTCLIPCONTSUBCLASS
{
	HWND hWndTbl;
	WNDPROC wpOld;
}
MTBLLISTCLIPCONTSUBCLASS, *LPMTBLLISTCLIPCONTSUBCLASS;

// MTBLLISTCLIPCTRLSUBCLASS
typedef struct MTBLLISTCLIPCTRLSUBCLASS
{
	HWND hWndTbl;
	WNDPROC wpOld;
	DWORD dwData;
}
MTBLLISTCLIPCTRLSUBCLASS, *LPMTBLLISTCLIPCTRLSUBCLASS;

// Typedefs Funktionen
typedef int (WINAPI *FLATSB_SETSCROLLPOS)(HWND, int, int, BOOL);
typedef int (WINAPI *FLATSB_SETSCROLLRANGE)(HWND, int, int, int, BOOL);
typedef BOOL (WINAPI *INVERTRECT)(HDC, CONST RECT*);
typedef BOOL (WINAPI *PATBLT)(HDC, int, int, int, int, DWORD);
typedef BOOL (WINAPI *SCROLLWINDOW)(HWND, int, int, CONST RECT*, CONST RECT*);
typedef int (WINAPI *SETSCROLLPOS)(HWND, int, int, BOOL);
typedef BOOL (WINAPI *SETSCROLLRANGE)(HWND, int, int, int, BOOL);

// Funktionsprototypen
int CALLBACK MTblEnumFontsProc( ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, int iFontType, LPARAM lParam );
int CALLBACK MTblEnumFontSizesProc( ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, int iFontType, LPARAM lParam ); 
HWND MTblGetColHdr( HWND hWndTbl, int iX, int iY );
int MTblGetColXCoord( HWND hWndTbl, HWND hWndCol, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis );
HWND MTblGetFocusCol( HWND hWndTbl );
int MTblGetRowYCoord( HWND hWndTbl, long lRow, CMTblMetrics * ptm, LPPOINT ppt, LPPOINT pptVis, BOOL bInVisRangeOnly = TRUE );
LPMTBLSUBCLASS MTblGetSubClass( HWND hWndTbl );
LPMTBLSUBCLASS MTblGetSubClassFromCol( HWND hWndCol );
LPMTBLSUBCLASS MTblGetSubClassFromItem( LPMTBLITEM pItem );
LPMTBLSUBCLASS MTblGetSubClass_GetTip( HWND hWndParam, long lParam, int iType );
LPMTBLSUBCLASS MTblGetSubClass_SetTip( HWND hWndParam, long lParam, int iType );
BOOL MTblIsListOpen( HWND hWndTbl, HWND hWndCol );
BOOL MTblPropStrToBool( HSTRING hs );
LRESULT CALLBACK MTblWndProc( HWND hWndTbl, UINT uMsg, WPARAM wParam, LPARAM lParam );

// ApiHook
#ifdef USE_APIHOOK

typedef struct HOOKDESC
{
	char szImpDll[MAX_PATH];
	HOOKFUNCDESCA HookFunc;
}
HOOKDESC, *LPHOOKDESC;

class CHookDesc
{
public:
	// Variablen
	HMODULE m_hModule;
	TCHAR m_szModule[MAX_PATH];
	char m_szImpDll[MAX_PATH];
	HOOKFUNCDESCA m_HookFunc;

	// Konstruktor
	CHookDesc( );

	// Operatoren
	CHookDesc & operator=( const CHookDesc &hd );
};

int WINAPI MyFlatSB_SetScrollPos(HWND hWnd, int fnBar, int nPos, BOOL fRedraw);
int WINAPI MyFlatSB_SetScrollRange(HWND hWnd, int nBar, int nMinPos, int nMaxPos, BOOL bRedraw);
BOOL WINAPI MyInvertRect( HDC hDC, CONST RECT *lprc );
BOOL WINAPI MyPatBlt( HDC hDC, int iXLeft, int iYLeft, int iWidth, int iHeight, DWORD dwRop );
BOOL WINAPI MyScrollWindow( HWND hWnd, int XAmount, int YAmount, CONST RECT *lpRect, CONST RECT *lpClipRect );
int WINAPI MySetScrollPos( HWND hWnd, int fnBar, int nPos, BOOL fRedraw );
BOOL WINAPI MySetScrollRange( HWND hWnd, int nBar, int nMinPos, int nMaxPos, BOOL bRedraw );

BOOL Hook( BOOL bHook );
BOOL HookTabliXX( BOOL bHook );
#endif // USE_APIHOOK

#endif // _INC_MTBL
