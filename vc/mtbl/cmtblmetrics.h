// cmtblmetrics.h

#ifndef _INC_CMTBLMETRICS
#define _INC_CMTBLMETRICS

// Includes
#include <windows.h>
#include "ctd.h"

// Globale Variablen
extern CCtd Ctd;
extern UINT guiMsgOffs;

// Struktur MTBLGETMETRICS
typedef struct tag_MTBLGETMETRICS
{
	int iHeaderHeight;
	int iCellHeight;
	POINT ptCharBox;
	POINT ptLineSize;
	POINT ptCellLeading;
	int iSplitBarTop;
	int iSplitBarHeight;
	int iLockedColumnsWidth;
}
MTBLGETMETRICS, *LPMTBLGETMETRICS;

// Klasse CMTblMetrics
class CMTblMetrics
{
public:
	// Werte
	RECT m_ClientRect;
	int m_HeaderHeight;
	int m_CellHeight;
	int m_CharBoxX;
	int m_CharBoxY;
	int m_LineSizeX;
	int m_LineSizeY;
	int m_CellLeadingX;
	int m_CellLeadingY;
	int m_SplitBarTop;
	int m_SplitBarHeight;
	int m_LockedColumnsWidth;
	int m_LinesPerRow;
	int m_RowHeight;
	int m_RowsTop;
	int m_RowsBottom;
	int m_SplitRowsTop;
	int m_SplitRowsBottom;
	int m_RowHeaderLeft;
	int m_RowHeaderRight;
	int m_LockedColumns;
	int m_ColumnsLeft;
	int m_ColumnsRight;
	int m_LockedColumnsLeft;
	int m_LockedColumnsRight;
	int m_ColumnOrigin;
	long m_MinVisibleRow;
	long m_MaxVisibleRow;
	// Konstruktor
	CMTblMetrics( void );
	CMTblMetrics( HWND hWndTbl );
	// Destruktor
	~CMTblMetrics( void );
	// Funktionen
	int CalcRowHeight( int iLinesPerRow );
	void Get( HWND hWndTbl, BOOL bNoVisRange = FALSE );
	// Statische Funktionen
	static long QueryCheckBoxSize( HWND hWndTbl );
};

#endif // _INC_CMTBLMETRICS