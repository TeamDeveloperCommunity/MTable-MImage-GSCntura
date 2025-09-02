// cmtblrows.h

#ifndef _INC_CMTBLROWS
#define _INC_CMTBLROWS

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include "ctd.h"
#include "cmtblrow.h"

// Defines
#define MTBL_MAXROWSINARRAY 1048576

// Globale Variablen
extern tstring gsEmpty;

// Klasse CMTblRows
class CMTblRows
{
private:
	// Variablen
	CMTblRow ** m_Rows;				// Zeilen-Array
	CMTblRow ** m_SplitRows;		// Split-Zeilen-Array
	long m_UpperBound;				// Obergrenze Zeilen-Array
	long m_SplitUpperBound;			// Obergrenze Split-Zeilen-Array
	long m_LastValidPos;			// Höchste gültige Zeilenposition
	long m_LastValidSplitPos;		// Höchste gültige Split-Zeilenposition
	long m_ChildRows;				// Anzahl Kindzeilen
	CMTblCounter *m_pCounter;		// Zähler

	// Funktionen
	void InitMembers( );
public:
	// Konstruktor
	CMTblRows( );
	CMTblRows( CMTblCounter *pCounter );

	// Destruktor
	~CMTblRows( );

	// Funktionen
	BOOL AnyMergeCellInCol( HWND hWndCol );
	CMTblRow * CreateRow( long lRow );
	BOOL DeleteRow( long lRow, int iMode );
	BOOL DelRowUserData( long lRow, tstring &sKey );
	CMTblRow * EnsureValid( long lRow );
	BOOL EnsureValidPos( long lPos );
	BOOL EnsureValidSplitPos( long lPos );
	long EnumRowUserData( long lRow, HARRAY haKeys, HARRAY haValues );
	CMTblRow * FindNextRow( LPLONG plRow );
	CMTblRow * FindNextSplitRow( LPLONG plRow );
	CMTblRow * FindPrevRow( LPLONG plRow );
	CMTblRow * FindPrevSplitRow( LPLONG plRow );
	BOOL FreeRows( );
	BOOL FreeSplitRows( );
	long GetChildRows( CMTblRow *pRow, RowPtrVector *pRows, BOOL bBottomUp = FALSE );
	CMTblImg * GetColImagePtr( long lRow, HWND hWndCol );
	CMTblCol * GetColPtr( long lRow, HWND hWndCol );
	CMTblRow * GetFirstChildRow( CMTblRow * pRow, BOOL bBottomUp = FALSE );
	CMTblRow * GetLastChildRow( CMTblRow * pRow, BOOL bBottomUp = FALSE );
	CMTblRow * GetLastDescendantRow( CMTblRow * pRow, BOOL bBottomUp = FALSE );
	CMTblRow * GetNextChildRow( CMTblRow * pRow, BOOL bBottomUp = FALSE );
	long GetOrigParentRow( long lRow );
	long GetParentRow( long lRow );
	CMTblRow * GetPrevChildRow( CMTblRow * pRow, BOOL bBottomUp = FALSE );
	CMTblRow ** GetRowArray( long lRow, LPLONG plUpperBound, LPLONG plLastValidPos );
	long GetRowHeight( long lRow );
	int GetRowLevel( long lRow );
	long GetRowLogX( long lRow, long lDefRowHt );
	CMTblRow * GetRowPtr( long lRow );
	tstring & GetRowUserData( long lRow, tstring &sKey );
	long GetRowVisPos( long lRow );
	long GetVisPosRow( long lVisPos, BOOL bSplitRows = FALSE );
	BOOL InsertRow( long lRow );
	BOOL IsParentRow( long lRow );
	BOOL IsRowDescendantOf( long lRow, long lParentRow );
	BOOL IsRowDelAdj( long lRow );
	BOOL IsRowDisCol( long lRow, HWND hWndCol );
	BOOL IsRowHyperlinkCol( long lRow, HWND hWndCol );
	BOOL IsRowDisabled( long lRow );
	BOOL IsValidRowPtr( CMTblRow *pRow );
	BOOL QueryRowFlags( long lRow, DWORD dwFlags );
	BOOL Reset( );
	BOOL ResetRowHeights( BOOL bSplitRows = FALSE );
	BOOL SetCellDisabled( long lRow, HWND hWndCol, BOOL bSet );
	BOOL SetCellFlagsAnyRows( HWND hWndCol, DWORD dwFlags, BOOL bSet );
	BOOL SetCellHyperlink( long lRow, HWND hWndCol, CMTblHyLnk & hl );
	BOOL SetMerged( long lRow, long lMerged );
	BOOL SetParentRow( CMTblRow *pRow, CMTblRow *pParentRow, BOOL bAsOrig = FALSE );
	BOOL SetParentRow( long lRow, long lParentRow, BOOL bAsOrig = FALSE );
	BOOL SetRowDelAdj( long lRow );
	BOOL SetRowDisabled( long lRow, BOOL bSet );
	DWORD SetRowFlags( long lRow, DWORD dwFlags, BOOL bSet );
	BOOL SetRowHeight( long lRow, long lHeight );
	BOOL SetRowHidden( long lRow, BOOL bSet );
	BOOL SetRowsInternalFlags( BOOL bSplitRows, DWORD dwFlags, BOOL bSet );
	BOOL SetRowUserData(long lRow, tstring &sKey, tstring &sValue);
	BOOL SwapRows( long lRow1, long lRow2 );

	// Inline-Funktionen
	BOOL AnyChildRows( ) { return ( m_ChildRows > 0 ); }
	long GetSplitPosRow( long lPos ) { return lPos + TBL_MinSplitRow; }
	long GetSplitRowPos( long lRow ) { return lRow - TBL_MinSplitRow; }
	BOOL IsValidPos( long lPos ) { return lPos >= 0 && lPos <= m_UpperBound; }
	BOOL IsValidRow( long lRow ) { return GetRowPtr( lRow ) != NULL; }
	BOOL IsValidSplitPos( long lPos ) { return lPos >= 0 && lPos <= m_SplitUpperBound; }
};

#endif // _INC_CMTBLROWS