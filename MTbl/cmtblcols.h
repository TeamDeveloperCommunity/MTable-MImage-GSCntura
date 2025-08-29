// cmtblcols.h

#ifndef _INC_CMTBLCOLS
#define _INC_CMTBLCOLS

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <list>
#include <map>
#include <stack>
#include "tstring.h"
#include "cmtblcol.h"

// Using
using namespace std;

// Typedefs
typedef map<HWND,CMTblCol> MTblColList;
typedef MTblColList *PMTblColList;

// Klasse CMTblCols
#pragma pack(1)
class CMTblCols
{
public:
	MTblColList m_List;				// Liste Spalten
	MTblColList m_HdrList;			// Liste Spaltenköpfe
	// Konstruktoren
	CMTblCols( );
	CMTblCols( BYTE nColType, CMTblCounter *pCounter );
	// Destruktor
	~CMTblCols( );
	// Funktionen
	CMTblCol * EnsureValid( HWND hWndCol );
	CMTblCol * EnsureValidHdr( HWND hWndCol );
	CMTblCol * Find( HWND hWndCol );
	CMTblCol * FindHdr( HWND hWndCol );
	CMTblBtn * GetBtn( HWND hWndCol );
	BOOL GetBtnPushed( HWND hWndCol );
	CMTblImg * GetColHdrImagePtr( HWND hWndCol );
	BOOL GetDisabled( HWND hWndCol );
	CMTblFont * GetFont( HWND hWndCol );
	CMTblBtn * GetHdrBtn( HWND hWndCol );
	BOOL GetHdrBtnPushed( HWND hWndCol );
	CMTblLineDef * GetHdrLineDef(HWND hWndCol, int iLineType);
	CMTblHyLnk * GetHyLnk( HWND hWndCol );
	CMTblLineDef * GetLineDef(HWND hWndCol, int iLineType);
	long GetMerged( HWND hWndCol );
	long GetMergedRows( HWND hWndCol );
	BOOL SetBtn( HWND hWndCol, CMTblBtn & b );
	BOOL SetBtnPushed( HWND hWndCol, BOOL bSet );
	BOOL SetColsFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetColsInternalFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetDisabled( HWND hWndCol, BOOL bSet );
	BOOL SetFont ( HWND hWndCol, CMTblFont & f );
	BOOL SetHdrBtn( HWND hWndCol, CMTblBtn & b );
	BOOL SetHdrBtnPushed( HWND hWndCol, BOOL bSet );
	BOOL SetHdrLineDef(HWND hWndCol, CMTblLineDef & ld, int iLineType);
	BOOL SetHyLnk( HWND hWndCol, CMTblHyLnk & hl );
	BOOL SetImage( HWND hWndCol, HANDLE hImage );
	BOOL SetImage2( HWND hWndCol, HANDLE hImage );
	BOOL SetLineDef(HWND hWndCol, CMTblLineDef & ld, int iLineType);
	BOOL SetMerged( HWND hWndCol, long lMerged );
	BOOL SetMergedRows( HWND hWndCol, long lMergedRows );
private:
	BYTE m_nColType;				// Typ der Spalten
	CMTblCounter *m_pCounter;		// Zähler
	// Funktionen
	void InitMembers( );
};
#pragma pack()

#endif // _INC_CMTBLCOLS