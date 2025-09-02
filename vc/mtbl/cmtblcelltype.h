// cmtblcelltype.h

#ifndef _INC_CMTBLCELLTYPE
#define _INC_CMTBLCELLTYPE

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <vector>
#include "ctd.h"
#include "tstring.h"
// Using
using namespace std;

//Typedefs
typedef vector< pair<tstring,long> > CellTypeListEntries;

// Klasse CMTblCellType
#pragma pack(1)
class CMTblCellType
{
public:
	// Variablen
	int m_iType;							// Typ, eine der COL_CellType_* Konstanten
	DWORD m_dwFlags;						// Flags
	int m_iLines;							// Anzahl Zeilen
	tstring m_sCheckedVal;					// String "Checkbox an"
	tstring m_sUnCheckedVal;				// String "Checkbox aus"
	CellTypeListEntries m_ListEntries;		// Listeneinträge

	BOOL m_bSettingColumn;					// Spalte wird gerade gesetzt
	
	// Konstruktoren
	CMTblCellType( );

	// Destruktor
	~CMTblCellType( );

	// Operatoren
	CMTblCellType & operator=( const CMTblCellType &ct );
	BOOL operator==( CMTblCellType &ct );
	BOOL operator!=( CMTblCellType &ct );

	// Funktionen
	void DefineCheckBox( DWORD dwFlags, tstring &sCheckedVal, tstring &sUnCheckedVal );
	void DefineDropDownList( DWORD dwFlags, int iLines, CellTypeListEntries *pe );
	BOOL DefineFromColumn( HWND hWndCol );
	void DefinePopupEdit( DWORD dwFlags, int iLines );
	void DefineStandard( );
	void InitMembers( BOOL bInitListData = TRUE );
	BOOL Get( LPINT piType, LPDWORD pdwFlags, LPINT piLines, LPHSTRING phsCheckedVal, LPHSTRING phsUnCheckedVal );
	int GetListData( HWND hWndCol, CellTypeListEntries *pe );
	BOOL IsCheckBox( ) { return m_iType == COL_CellType_CheckBox; }
	BOOL IsDropDownList( ) { return m_iType == COL_CellType_DropDownList; }
	BOOL IsDropDownListWndCls( HWND hWnd );
	BOOL IsPopupEdit( ) { return m_iType == COL_CellType_PopupEdit; }
	BOOL IsPopupEditWndCls( HWND hWnd );
	BOOL IsMyTypeWndCls( HWND hWnd );
	BOOL IsStandard( ) { return m_iType == COL_CellType_Standard; }
	BOOL SetColumn( HWND hWndCol );
	void SortListEntries( );
};
#pragma pack()

#endif // _INC_CMTBLCELLTYPE