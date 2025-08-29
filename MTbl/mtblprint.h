#ifndef _MTBLPRINT_H
#define _MTBLPRINT_H

//*********************************************************
// Alexander Hakert Arbeitsbereich
//*********************************************************

#include "mtbl.h"

#ifndef UNICODE
#define to_tstring	std::to_string
#else
#define to_tstring	std::to_wstring
#endif

//******** Preview-Part ********

//Schon vorhanden?
#pragma comment(lib, "ComCtl32.lib")

#pragma comment(linker, \
"\"/manifestdependency:type='win32' "\
"name='Microsoft.Windows.Common-Controls' "\
"version='6.0.0.0' "\
"processorArchitecture='*' "\
"publicKeyToken='6595b64144ccf1df' "\
"language='*'\"")

//******** Preview-Part Ende ********

//Fehlerkonstanten
#define ERR_MARGINS_TOO_BIG	-1
#define ERR_NO_COLS_TO_PRINT	-2
#define ERR_NO_PRINTER	-3
#define ERR_INVALID_PRINTER	-4
#define ERR_NOT_SUBCLASSED	-6
#define ERR_INVALID_LANG	-7

#define MTPL_NO_IMAGE_SCALE	1
#define MTPL_NO_TEXT_SCALE	2

const tstring mtblPrint_Placeholders[] = { _T("%PAGENR%"), _T("%PAGECOUNT%") };
const int mtblPrint_PlaceholderCount = 2;

typedef struct MTBLPRINT_PRINTERDEVMODE
{
	tstring tsPrinter;
	LPDEVMODE lpDevMode;
}
MTBLPRINT_PRINTERDEVMODE, *LPMTBLPRINT_PRINTERDEVMODE;

/*
long mtblPrint_lCurrentPage;
HANDLE mtblPrint_hDllCentura;
BOOL mtblPrint_bInitialized;
long mtblPrint_PageCount;
std::vector<MTBLPRINT_PRINTERDEVMODE> haMTblPrinterDevModes;
*/
BOOL SetPrinterDevMode(tstring tsPrinter, LPDEVMODE lpDevMode);
LPDEVMODE GetPrinterDevMode(tstring tsPrinter);
BOOL PrinterExists(tstring tsPrinter);
void GetDesktopResolution(int &horizontal, int &vertical);
LPDEVMODE GetPrinterDevModeCopy(tstring tsPrinter);

//******************** Druckstrukturen ********************

enum Justify { jfyLeft = 0, jfyCenter = 1, jfyRight = 2 };
enum MTblPrintOrientation { mtblOriPortrait = 0, mtblOriLandscape = 1 };
enum GridType { gtNo = 0, gtStandard = 1, gtTable = 2, gtStandard2 = 3, gtStandard3 = 4, gtStandard4 = 5 };

typedef struct PRINTLINEPOSTEXT
{
	tstring tsText;			//Text
	double dLeft;			//Linke Position
	double dWidth;			//Breite
	Justify justify;		//Ausrichtung

	BOOL FromUdv(HUDV hUdv);
}
PRINTLINEPOSTEXT, *LPPRINTLINEPOSTEXT;

#pragma pack(UDV_PACK)
typedef struct PRINTLINEPOSTEXTUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	HSTRING hsText;			//Text
	NUMBER nLeft;			//Linke Position
	NUMBER nWidth;			//Breite
	NUMBER nJustify;		//Ausrichtung
}
PRINTLINEPOSTEXTUDV, *LPPRINTLINEPOSTEXTUDV;
#pragma pack()

typedef struct PRINTLINE
{
	tstring tsLeftText, tsCenterText, tsRightText;	//Linker, mittlerer und rechter Text
	//LPPRINTLINEPOSTEXT plptPosText;						//Positions-Texte
	std::vector<PRINTLINEPOSTEXT> vecplPosText;
	int iPosTextCount;									//Anzahl Positionstexte
	double dFontSizeMulti;								//Fontgrößen-Multiplikator
	long lFontEnh;								//Fontstile bold, italic, underline, strikeout
	COLORREF crFontColor;									//Fontfarbe
	HANDLE hLeftImage, hCenterImage, hRightImage;		//Linkes, mittleres und rechtes Bild
	DWORD dwFlags;										//Flags

	BOOL FromUdv(HUDV hUdv);
}
PRINTLINE, *LPPRINTLINE;

#pragma pack(UDV_PACK)
typedef struct PRINTLINEUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	HSTRING hsLeftText, hsCenterText, hsRightText;		//Linker, mittlerer und rechter Text
	HARRAY haPosTexts;									//Positions-Texte
	NUMBER nPosTextCount;								//Anzahl Positionstexte
	NUMBER nFontSizeMulti;								//Fontgrößen-Multiplikator
	NUMBER nFontEnh;									//Fontstile bold, italic, underline, strikeout
	NUMBER nFontColor;									//Fontfarbe
	NUMBER nLeftImage, nCenterImage, nRightImage;		//Linkes, mittleres und rechtes Bild
	NUMBER nFlags;										//Flags
}
PRINTLINEUDV, *LPPRINTLINEUDV;
#pragma pack()

typedef struct PRINTOPTIONS
{
	tstring tsPrinterName;			//Name des Druckers. Wenn leer, wird der Standarddrucker verwendet
	tstring tsPrevWndTitle;			//Titel des Vorschaufensters, leer = Standardtitel
	tstring tsDocName;				//Dokumentname, leer = Standardname
	std::vector<PRINTLINE> vecTitles;	//Titelzeilen
	std::vector<PRINTLINE> vecTotals;	//Summenzeilen
	std::vector<PRINTLINE> vecPageHeaders;	//Seitenkopfzeilen
	std::vector<PRINTLINE> vecPageFooters;	//Seitenfußzeilen
	int iTitleCount;				//Anzahl Titelzeilen
	int iTotalCount;				//Anzahl Summenzeilen
	int iPageHeaderCount;			//Anzahl Seitenkopfzeilen
	int iPageFooterCount;			//Anzahl Seitenfußzeilen
	BOOL bPageNrs;					//Seitenzahlen ja/nein
	BOOL bNoPageNrsScale;			//Seitenzahlen nicht skalieren ja/nein
	int iCopies;					//Anzahl Kopien
	double dScale;					//Skalierungsfaktor
	MTblPrintOrientation orientation;	//Hoch- oder Querformat
	GridType gridType;				//Gitternetz-Typ
	int iLeftMargin, iTopMargin, iRightMargin, iBottomMargin;	//Seitenränder in mm
	WORD dwRowFlagsOn;				//Flags, von denen eine Zeile mindestens eines haben muss, um gedruckt zu werden
	WORD dwRowFlagsOff;			//Flags, von denen eine Zeile keines haben darf, um gedruckt zu werden
	WORD dwColFlagsOn;				//Flags, von denen eine Spalte mindestens eines haben muss, um gedruckt zu werden
	WORD dwColFlagsOff;			//Flags, von denen eine Spalte keines haben darf, um gedruckt zu werden
	int iPrevWndState;				//Fensterstatus des Verschaufensters
	int iPrevWndLeft;				//Position links des Vorschaufensters
	int iPrevWndTop;				//Position obens des Vorschaufensters
	int iPrevWndWidth;				//Breite des Vorschaufensters
	int iPrevWndHeight;				//Höhe des Vorschaufensters
	BOOL bPrevWndFullPage;			//Ganze Seite Vorschaufenster
	BOOL bFitCellSize;				//Optimale Zellengröße ja/nein
	BOOL bColHeaders;				//Spaltenköpfe drucken ja/nein
	int iLanguage;					//Sprache
	BOOL bSpreadCols;				//Spalten auf mehrere Seiten verteilen ja/nein
	BOOL bPreview;					//Vorschau ja/nein

	BOOL FromUdv(HUDV hUdv);
}
PRINTOPTIONS, *LPPRINTOPTIONS;

#pragma pack(UDV_PACK)
typedef struct PRINTOPTIONSUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	HSTRING hsPrinterName;		//Name des Druckers. Wenn leer, wird der Standarddrucker verwendet
	HSTRING hsPrevWndTitle;		//Titel des Vorschaufensters, leer = Standardtitel
	HSTRING hsDocName;			//Dokumentname, leer = Standardname
	HARRAY haTitles;			//Titelzeilen
	HARRAY haTotals;			//Summenzeilen
	HARRAY haPageHeaders;		//Seitenkopfzeilen
	HARRAY haPageFooters;		//Seitenfußzeilen
	NUMBER nTitleCount;			//Anzahl Titelzeilen
	NUMBER nTotalCount;			//Anzahl Summenzeilen
	NUMBER nPageHeaderCount;	//Anzahl Seitenkopfzeilen
	NUMBER nPageFooterCount;	//Anzahl Seitenfußzeilen
	NUMBER nPageNrs;			//Seitenzahlen ja/nein
	NUMBER nNoPageNrsScale;		//Seitenzahlen nicht skalieren ja/nein
	NUMBER nCopies;				//Anzahl Kopien
	NUMBER nScale;				//Skalierungsfaktor
	NUMBER nOrientation;		//Hoch- oder Querformat
	NUMBER nGridType;			//Gitternetz-Typ
	NUMBER nLeftMargin, nTopMargin, nBottomMargin, nRightMargin;	//Seitenränder in mm
	NUMBER nRowFlagsOn;			//Flags, von denen eine Zeile mindestens eines haben muss, um gedruckt zu werden
	NUMBER nRowFlagsOff;		//Flags, von denen eine Zeile keines haben darf, um gedruckt zu werden
	NUMBER nColFlagsOn;			//Flags, von denen eine Spalte mindestens eines haben muss, um gedruckt zu werden
	NUMBER nColFlagsOff;		//Flags, von denen eine Spalte keines haben darf, um gedruckt zu werden
	NUMBER nPrevWndState;		//Fensterstatus des Verschaufensters
	NUMBER nPrevWndLeft;		//Position links des Vorschaufensters
	NUMBER nPrevWndTop;			//Position obens des Vorschaufensters
	NUMBER nPrevWndWidth;		//Breite des Vorschaufensters
	NUMBER nPrevWndHeight;		//Höhe des Vorschaufensters
	NUMBER nPrevWndFullPage;	//Ganze Seite Vorschaufenster
	NUMBER nFitCellSize;		//Optimale Zellengröße ja/nein
	NUMBER nColHeaders;			//Spaltenköpfe drucken ja/nein
	NUMBER nLanguage;			//Sprache
	NUMBER nSpreadCols;			//Spalten auf mehrere Seiten verteilen ja/nein
	NUMBER nPreview;			//Vorschau ja/nein
}
PRINTOPTIONSUDV, *LPPRINTOPTIONSUDV;
#pragma pack()

typedef struct PRINTMETRICS
{
	int iPixelPerInchX, iPixelPerInchY;		//Pixel pro Inch X/Y
	double dPixelPerMMX, dPixelPerMMY;		//Pixel pro mm X/Y
	int iPhysWidth, iPhysHeight;			//Physikalische Breite/Höhe
	int iLeftOffset, iTopOffset, iRightOffset, iBottomOffset;	//Rand-Offsets
	int iLeftMargin, iTopMargin, iRightMargin, iBottomMargin;	//Ränder
	RECT clipRect;							//Clipping-Rechteck
	int iTotalColWidth;						//Gesamtbreite aller zu druckenden Spalten
	tstring tsFontName;						//Name des Fonts
	double dFontHeight;						//Absolute Font-Höhe
	double dFontHeightNS;					//Absolute Font-Höhe nicht skaliert

	long lFontStyle;						//Font-Style
	double dScale;							//Skalierung
}
PRINTMETRICS, *LPPRINTMETRICS;

typedef struct TABLEINFO
{
	HWND handle;							//Window Handle
	tstring tsFontName;						//Fontname
	int iFontSize;							//Fontgröße
	DWORD dwFontEnh;						//Font-Attribute
	COLORREF crTextColor;					//Textfarbe
	COLORREF crBackgroundColor;				//Hintergrundfarbe
	int iLinesPerRow;						//Linien pro Zeile
	BOOL bGrayedHeaders;					//Graue Spaltenköpfe ja/nein
	HWND hwTreeCol;							//Tree-Spalte
	DWORD dwTreeFlags;						//Tree-Flags
	int iTreeIndentPerLevel;				//Einrückung pro Level Tree-Spalte
	int iNodeSize;							//Knoten-Größe
	DWORD dwNodeOuterColor;					//Knoten-Farbe außen
	DWORD dwNodeInnerColor;					//Knoten-Farbe innen
	DWORD dwNodeBackColor;					//Knoten-Farbe Hintergrund
	long iIndentPerLevel;					//Einrückung pro Level
	TCHAR tcPasswordChar;					//Passwort-Zeichen
	MTBLGETMETRICS metrics;					//Tabellen-Maße
	int iPixelPerInchX, iPixelPerInchY;		//Pixel pro Inch X/Y
	BOOL bNoRowLines;						//Keine Zeilenlinien ja/nein
	BOOL bNoColLines;						//Keine Spaltenlinien ja/nein
	LPMTBLMERGECELLS pMergeCells;			//Zeiger auf Merge-Zellen
}
TABLEINFO, *LPTABLEINFO;

typedef struct COLINFO
{
	HWND handle;		//Window Handle
	int iID;			//ID
	int iCellType;		//Zellentyp
	int iWidthPix;		//Breite in Pixel
	int iJustify;		//Ausrichtung
	BOOL bWordWrap;		//Word Wrap ja/nein
	tstring tsTitle;	//Titel
	BOOL bTreeCol;		//Tree-Spalte ja/nein
	BOOL bLocked;		//Gelockt ja/nein
	BOOL bInvisible;	//Passwort-Spalte ja/nein
	BOOL bNewPage;		//Spalte, ab der eine neue Seite beginnt ja/nein
	LPVOID HdrGrp;		//Header-Gruppe
	BOOL bNoRowLines;	//Keine Zeilenlinien ja/nein
	BOOL bNoColLine;	//Kein Spaltenlinie ja/nein
}
COLINFO, *LPCOLINFO;

typedef struct OWNERDRAWNITEM
{
	long lType;
	HWND hWndParam;
	long lParam;
	HWND hWndPart;
	long lPart;
	HANDLE hDC;
	RECT rectOdi;
	long lSelected;
	long lXLeading;
	long lYLeading;
	long lPrinting;
	double dYResFactor;
	double dXResFactor;
}
OWNERDRAWNITEM, *LPOWNERDRAWNITEM;

typedef struct PRINTCELL
{
	tstring tsText;						//Zu druckender Text
	int iTextJustify;					//Horizontale Ausrichtung des Texts(DT_LEFT, DT_CENTER, DT_RIGHT)
	int iTextVAlign;					//Vertikale Ausrichtung des Texts(DT_LEFT, DT_CENTER, DT_RIGHT)
	int iTextOffsetX, iTextOffsetY;		//Text-Abstände X/Y
	RECT rTextArea;						//Textausgabe-Bereich
	RECT rTextRect;						//Textausgabe-Rechteck
	DWORD dwTextDrawFlags;				//Textausgabe-Flags
	BOOL bWordBreak;					//Automatischer Zeilenumbruch ja/nein
	BOOL bSingleLine;					//Einzeilige Ausgabe ja/nein
	RECT rect;							//Umgebendes Rechteck
	RECT clipRect;						//Clipping-Rechteck
	RECT rBorders;						//Linienflags (0 = keine Linie, 1 = Linie)
	tstring tsFontName;					//Name des Fonts
	double dFontHeight;					//Höhe des Fonts

	long lFontStyle;					//Fontstyle
	COLORREF crFontColor;				//Farbe des Fonts
	COLORREF crBackgroundColor;			//Farbe des Hintergrundes (clWhite = transparent)
	BOOL bPaintNode;					//Knoten zeichnen ja/nein
	int iNodeType;						//Knoten-Typ
	RECT nodeRect;						//Knoten-Rechteck
	COLORREF crNodeOuterColor;			//Knoten-Farbe außen
	COLORREF crNodeInnerColor;			//Knoten-Farbe innen
	COLORREF crNodeBackColor;			//Knoten-Farbe Hintergrund
	HANDLE hNodeImageHandle;			//Handle Knoten-Bild
	SIZE siNodeImageSize;				//Größe Knoten-Bild
	double dTblPixSize;					//Größe eines Tabellenpixels für die Ausgabe
	HANDLE hImage;						//Bild-Handle
	int iImagePosX;						//Bildposition vom linken Rand
	SIZE siImageSize;					//Bildgröße
	RECT rImageRect;					//Bildausgabe-Rechteck
	BOOL bColHdr;						//Spaltenkopf ja/nein
	OWNERDRAWNITEM odi;					//Owner Drawn Item Daten
	BOOL bTreeBottomUp;					//Baum von unten nach oben ja/nein
	BOOL bHorzTreeLine;					//Horizontale Tree-Linie ja/nein
	RECT rHTL_Pos;						//Position horizontale Tree-Linie
	BOOL bVertTreeLineToParent;			//Vertikale Tree-Linien zur Elternzeile ja/nein
	std::vector<RECT> vecVTLTP_Pos;		//Positionen vertikaler Tree-Linien zur Elternzeile
	BOOL bVertTreeLineToChild;			//Vertikale Tree-Linie zur Kindzeile ja/nein
	RECT rVTLTC_Pos;					//Position vertikale Tree-Linie zur Kindzeile
	int iTreeLineStyle;					//Tree-Linien-Stil
	COLORREF crTreeLineColor;			//Farbe Tree-Linien
}
PRINTCELL, *LPPRINTCELL;

typedef struct PRINTCELLPARAMS
{
	HDC hdcCanvas;			//Handle Canvas
	HWND hwTable;			//Handle der Tabelle
	PRINTCELL cellData;		//Zelldaten
}
PRINTCELLPARAMS, *LPPRINTCELLPARAMS;

typedef struct CELLTOPRINT
{
	long lFromRow, lToRow;
	long lRows;
	HWND hwFromCol, hwToCol;
	long lFromColPos, lToColPos;
	long lWidth, lHeight;
	long lX;
	long lY;
	RECT clipRect;
	PRINTCELLPARAMS pcp;
}
CELLTOPRINT, *LPCELLTOPRINT;

typedef std::vector<COLINFO> VecColInfo;

typedef struct PRINTCOLHEADERSPARAMS
{
	HDC hdcCanvas;					//Handle Canvas
	int iYStart;					//Startpunkt Y
	int iYMax;						//Maximaler Endpunkt Y
	PRINTOPTIONS printOptions;		//Druckoptionen
	PRINTMETRICS printMetrics;		//Druckerabmessungen
	TABLEINFO tableInfo;			//Tabelleninformationen
	VecColInfo vecColsInfo;			//Spalteninformationen
	int iFromCol, iToCol;			//Start- und Endspalte
}
PRINTCOLHEADERSPARAMS, *LPPRINTCOLHEADERSPARAMS;

typedef struct PRINTCELLSPARAMS
{
	HDC hdcCanvas;					//Handle Canvas
	long lStartRow;					//Startzeile
	int iYStart;					//Startpunkt Y
	int iYMax;						//Maximaler Endpunkt Y
	PRINTOPTIONS printOptions;		//Druckoptionen
	PRINTMETRICS printMetrics;		//Druckerabmessungen
	TABLEINFO tableInfo;			//Tabelleninformationen
	VecColInfo vecColsInfo;			//Spalteninformationen
	int iFromCol, iToCol;			//Start- und Endspalte
}
PRINTCELLSPARAMS, *LPPRINTCELLSPARAMS;

typedef struct PRINTLINEPARAMS
{
	HDC hdcCanvas;					//Handle Canvas
	int iYStart;					//Startpunkt Y
	RECT clipRect;					//Clipping-Rechteck
	PRINTLINE line;					//Zu druckende Zeile
	tstring tsFontName;				//Name des Fonts
	double dFontHeight;				//Höhe des Fonts

	long lFontStyle;				//Fontstyle
	COLORREF crFontColor;			//Farbe des Fonts
	LPPRINTMETRICS lpPrintMetrics;	//Druckerabmessungen
	LPTABLEINFO lpTableInfo;		//Tabelleninformationen
	BOOL bYReversed;				//Wenn true, ist YStart die Unterkante
	BOOL bSimulation;				//Nur simulieren ja/nein
}
PRINTLINEPARAMS, *LPPRINTLINEPARAMS;

typedef struct PRINTPAGEPARAMS
{
	HDC hdcCanvas;					//Handle Canvas
	int iPageNr;					//Seitennummer
	long lStartRow;					//Startzeile
	PRINTOPTIONS printOptions;		//Druckoptionen
	PRINTMETRICS printMetrics;		//Druckerabmessungen
	TABLEINFO tableInfo;			//Tabelleninformationen
	VecColInfo vecColsInfo;			//Spalteninformationen
	int iFromCol, iToCol;			//Start- und Endspalte
	BOOL bColHeaders;				//Spaltenköpfe ja/nein
	BOOL bCells;					//Zellen ja/nein
}
PRINTPAGEPARMAS, *LPPRINTPAGEPRAMS;
//*********************************************************

//***** Preview Part *****
#define IDC_METAFILEVIEW					10000
#define IDC_PANEL							10001
#define IDC_DIALOG_CONTENT					10002
#define IDC_MENU_PANEL						10003
#define IDC_MENU_BTN_CLOSE					10004
#define IDC_MENU_BTN_PRINT					10005
#define IDC_MENU_BTN_SCALING				10006
#define IDC_MENU_BTN_FULLPAGE				10007
#define IDC_MENU_BTN_FIRSTPAGE				10008
#define IDC_MENU_BTN_PREVPAGE				10009
#define IDC_MENU_BTN_NEXTPAGE				10010
#define IDC_MENU_BTN_LASTPAGE				10011
#define IDC_TOOLBAR							10012
#define IDC_MENU_ITEM_PRINT_ALL				10013
#define IDC_MENU_ITEM_PRINT_CURRENT_PAGE	10014
#define IDC_MENU_ITEM_SCALE_USER_DEFINED	10015
#define IDC_MENU_ITEM_SCALE_10				10016
#define IDC_MENU_ITEM_SCALE_25				10017
#define IDC_MENU_ITEM_SCALE_50				10018
#define IDC_MENU_ITEM_SCALE_75				10019
#define IDC_MENU_ITEM_SCALE_100				10020
#define IDC_MENU_ITEM_SCALE_150				10021
#define IDC_MENU_ITEM_SCALE_200				10022
#define IDC_MENU_ITEM_SCALE_TO_PAGE_WIDTH	10023
#define IDC_PRINTPREVIEW_STATUSBAR			10024
#define PRINTPREVIEW_UNKNOWN	INT_MAX

typedef struct PRINTMETAFILE
{
	HENHMETAFILE hMetafile;
	int iRows;
}
PRINTMETAFILE, *LPPRINTMETAFILE;

typedef struct PRINTPREVIEW
{
	PRINTPREVIEW(LPMTBLSUBCLASS lpRefSubClass, PRINTOPTIONS printOptions, PRINTMETRICS printMetrics, TABLEINFO tableInfo, VecColInfo vecColsInfo, HDC hdcPrinter, LPDEVMODE lpDevMode);

	double dMinScale;		//Minimale Skalierung in %
	double dMaxScale;		//Maximale Skalierung in %
	int iUnknown;			//Wert für iLastPage, wenn die letzte Seite unbekannt ist
	int iImgOffset;			//Abstand links/oben von Metafile zu Panel
	LPMTBLSUBCLASS lpSubclassRef;	//Referenz auf die aufgerufende Subclass, verwendet zum Anzeigen/Drucken von Seiten und Abfrage von sprachabhängigen Texten
	PRINTOPTIONS printOptions;		//Verwendete Druckoptionen
	PRINTMETRICS printMetrics;		//Verwendete Druckmetriken
	TABLEINFO tableInfo;			//Tabelleninformationen
	VecColInfo vecColsInfo;			//Anzuzeigene bzw zu druckene Spalten
	HWND hWndDialog;		//Window Handle für den Vorschaudialog
	HWND hWndDialogContent;	//Window Handle für den Inhaltsbereich des Vorschaudialog, hierin befindet sich das Panel
	HWND hWndPanel;			//Window Handle für das Panel, das den Rahmen der Metafile darstellt
	HWND hWndMetafileView;	//Window Handle für das Fenster, das die Metafile beinhaltet
	int iMetafileViewWidth, iMetafileViewHeight;	//Metafile-Breite/Höhe
	HDC hdcPrinter;			//Device Context des Druckers
	int iMenuPanelHeight;	//Toolbar-Höhe
	int iMenuButtonWidth;	//Toolbar Button-Breite
	BOOL bFullPage;			//Ganze Seite anzeigen?

	HWND hWndToolbar;		//Window Handle für die Toolbar
	HIMAGELIST hImageList;	//Handle der ImageList für die Icons der Toolbar
	HMENU hMenuPrint, hMenuScale;	//Menu Handle für die DropDowns für Druck und Skalierung
	LPDEVMODE pDevMode;		//Device Mode des Druckers

	HWND hWndStatusbar;		//Window Handle für die Statusbar
	int iStatusbarHeight;	//Höhe der Statusbar

	/*
	Window Handle für die Fenster, die bei Start des Vorschaufensters deaktiviert wurden, damit das
	Vorschaufenster modal ist und keine Änderungen an der Tabelle oder den Druckoptionen vorgenommen werden können.
	*/
	std::vector<HWND> vecHwndDisabled;
	HWND hWndLastActiveWindow;		//Aktives Fenster, als das Vorschaufenster erstellt wurde. Genutzt um dieses hinterher wieder zu aktivieren

	//Scroll
	int xMinScroll;		//Minimaler horizontaler Scrollwert
	int xCurrentScroll;	//Aktueller horizontaler Scrollwert
	int xMaxScroll;		//Maximaler horizontaler Scrollwert
	int yMinScroll;		//Minimaler vertikaler Scrollwert
	int yCurrentScroll;	//Aktueller vertikaler Scrollwert
	int yMaxScroll;		//Maximaler vertikaler Scrollwert

	int iCurrentPage;	//Aktuelle Seite
	int iLastPage;		//Letzte Seite
	std::vector<int> vecPageStartRows;		//Startzeilen der Seiten
	std::vector<int> vecPageFromCols;		//Von-Spalten der Seiten
	std::vector<int> vecPageToCols;			//Bis-Spalten der Seiten
	std::vector<int> vecPageLastTotals;		//Letzte Totalzeilen der Seiten
	std::vector<PRINTMETAFILE> vecPageMetafiles;	//Metafiles der Seiten
	PRINTPAGEPARAMS printPageParams;		//Parameter für den Seitendruck

	BOOL bWaitCursor;		//Sanduhr-Cursor an/aus

	void OnDialogInit(HWND hWndDialog);
	void OnDlgUserDefinedScaleInit(HWND hWndUserDefinedScaleDlg);
	void OnDlgUserDefinedScaleOkClick(HWND hWndUserDefinedScaleDlg);
	void CloseBtnClick();
	void PrintBtnClick(UINT uiId);
	void UserDefinedScaleClick();
	void ScaleMenuItemClick(UINT uiId);
	void FullPageBtnClick();
	void FirstPageBtnClick();
	void PrevPageBtnClick();
	void NextPageBtnClick();
	void LastPageBtnClick();
	void DropDownClick(LPNMTOOLBAR lpnmtb);
	void OnMouseWheelScroll(short sDeltaScroll);
	void OnVerticalScroll(WORD wScrollRequest, WORD wScrollboxPosition);
	void OnHorizontalScroll(WORD wScrollRequest, WORD wScrollboxPosition);
	void OnContentSizeChanged(WORD wNewWidth, WORD wNewHeight);

private:
	void CreatePreviewWindow();
	void CreateToolbar(HWND hWndParent);
	void CreatePopupMenus();
	void InitWindowSettings();
	void StartPreview();
	void SetWaitCursor(BOOL bWait);
	int PrintCreateEnhMetafile(int iPageNr);
	int ShowPage(int iPageNr);
	void SizeObjects();
	BOOL Scale(double dScale);
	void ScaleToPageWidthClick();
	void ScaleContentToSize();
	void CheckScaleMenuItem(UINT uiId);
	void PrintCurrentPage();
	void PageChanged();
	void EnableButtons();
	void SetStatusbar();
	void UpdateScrollbar(int iWndWidth, int iWndHeight);
	void FreeMetafiles();

	void ReplaceSubstring(tstring &tString, tstring tsToReplace, tstring tsReplaceWith);
}
PRINTPREVIEW, *LPPRINTPREVIEW;

//***** Preview Part End *****

//***** Statusform Part Start *****

typedef struct PRINTSTATUSFORM
{
	//Konstruktor nur für Initialisierung von Werten
	PRINTSTATUSFORM(PRINTOPTIONS printOptions, LPMTBLSUBCLASS lpSubclassRef, LPBOOL lpBoolCancel);
	~PRINTSTATUSFORM();

	HWND hWndStatusForm;		//Window Handle für das Statusfenster
	HWND hWndBtnCancel;			//Window Handle für den Abbruchbutton
	HWND hWndLabel;				//Window Handle für das Label des Statusfenster, das die Information anzeigt
	LPMTBLSUBCLASS lpSubclassRef;	//Referenz auf die aufrufende Subclass
	PRINTOPTIONS printOptions;		//Druckoptionen
	LPBOOL lpBoolCancel;			//Pointer auf BOOL, der von der Druckfunktion verwendet wird, um festzustellen, ob der Druck abgebrochen werden soll
	tstring tsTitle;				//Title des Statusfensters
	tstring tsButtonText;			//Text des Buttons im Statusfensters
	HANDLE hStatusThread;			//Handle für den Thread, in dem das Statusfenster erstellt wird

	void OnStatusFormInit(HWND hWndStatusForm);
	void CreateStatusForm();
	void OnCancelClick();
}
PRINTSTATUSFORM, *LPPRINTSTATUSFORM;

//***** Statusform Part Ende *****
//************************************************************

#endif // _MTBLPRINT_H