// cmtblrow.h

#ifndef _INC_CMTBLROW
#define _INC_CMTBLROW

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <map>
#include "tstring.h"
#include "cmtblcols.h"
#include "cmtbltip.h"
#include "ctd.h"
#include "rowflags.h"
#include "lines.h"

// Interne Row-Flags
#define ROW_IFLAG_SORTVISIT 1
#define ROW_IFLAG_LOADCHILDVISIT 2
#define ROW_IFLAG_SELECTED_BEFORE 4
#define ROW_IFLAG_SELECTED_AFTER 8
#define ROW_IFLAG_HIDE 16
#define ROW_IFLAG_HIDE_REDRAW 32
#define ROW_IFLAG_DELETE 64
#define ROW_IFLAG_DELETE_ADJUST 128
#define ROW_IFLAG_LOAD_CHILD_ROWS 256

// Using
using namespace std;

// Vorabdeklarationen
class CMTblRow;

// Typedefs
typedef map< tstring, tstring, less<tstring> > StrMap;
typedef StrMap * PStrMap;
typedef StrMap::value_type STRMAPVAL;
typedef StrMap::iterator STRMAPIT;
typedef vector< CMTblRow* > RowPtrVector;

// Globale Variablen
extern tstring gsEmpty;
extern CCtd Ctd;

// Klasse CMTblRow
#pragma pack(1)
class CMTblRow
{
private:
	// Variablen
	long m_Nr;							// Nummer der Zeile
	struct tagBits						// Bits
	{
		BYTE DelAdj : 1;				// Mit TBL_Adjust gelöscht ja/nein
		BYTE Disabled : 1;				// Zeile gesperrt ja/nein
		BYTE Ownerdrawn : 1;			// Benutzerdefiniert gezeichnet ja/nein
		BYTE HdrOwnerdrawn : 1;			// Header Benutzerdefiniert gezeichnet ja/nein
		BYTE HiddenFlag : 1;			// ROW_Hidden gesetzt ja / nein
		BYTE : 0;						// Für Speicher-Ausrichtung
	} m_Bits;				
	PStrMap m_UserData;					// Map der Benutzerdaten
	CMTblRow * m_ParentRow;				// Adresse der Elternzeile
	CMTblRow * m_OrigParentRow;			// Adresse der ursprünglichen Elternzeile
	DWORD m_Flags;						// Flags
	long m_ChildRows;					// Anzahl Kindzeilen
	long m_ChildRowsVis;				// Anzahl sichtbare Kindzeilen (MTBL_ROW_HIDDEN nicht gesetzt)
	WORD m_InternalFlags;				// Interne Flags für unterschiedliche Verwendungszwecke
	CMTblCounter * m_pCounter;			// Zähler
	long m_Height;						// Höhe, 0 = Standard

	// Funktionen
	void CreateMembers( );
	void InitMembers( );

public:
	CMTblColor * m_TextColor;			// Textfarbe
	CMTblColor * m_BackColor;			// Hintergrundfarbe
	CMTblColor * m_HdrBackColor;		// Header-Hintergrundfarbe
	CMTblCols * m_Cols;					// Spalten ( = Zellen )
	CMTblFont * m_Font;					// Font
	CMTblTip * m_HdrTip;				// Row-Header-Tip
	long m_Merged;						// Anzahl verbundener Zeilen
	CMTblLineDef* m_HorzLineDef;		// Definition horizontale Linien
	CMTblLineDef* m_VertLineDef;		// Definition vertikale Linien
	CMTblLineDef* m_HdrHorzLineDef;		// Definition horizontale Linien Header
	CMTblLineDef* m_HdrVertLineDef;		// Definition vertikale Linien Header

	// Konstruktoren
	CMTblRow( );
	CMTblRow( CMTblCounter *pCounter );

	// Destruktor
	~CMTblRow( );

	// Funktionen
	BOOL DelUserData( tstring &sKey );
	long EnumUserData( HARRAY haKeys, HARRAY haValues );
	CMTblImg * GetColImagePtr( HWND hWndCol );
	CMTblLineDef * GetHdrLineDef(int iLineType);
	long GetHeight( ) { return m_Height; }
	int GetLevel();
	CMTblLineDef * GetLineDef(int iLineType);
	tstring & GetUserData( tstring &sKey );
	BOOL IsDescendantOf( CMTblRow * pRow );
	BOOL IsOrigDescendantOf( CMTblRow * pRow );
	BOOL QueryInternalFlags( DWORD dwFlags );
	void SetDelAdj( );
	DWORD SetFlags(DWORD dwFlags, BOOL bSet);
	BOOL SetHdrLineDef(CMTblLineDef& ld, int iLineType);
	DWORD SetInternalFlags(DWORD dwFlags, BOOL bSet);
	BOOL SetLineDef(CMTblLineDef& ld, int iLineType);
	BOOL SetOrigParentRow( CMTblRow * pRow );
	BOOL SetParentRow( CMTblRow * pRow, BOOL bAsOrig = FALSE );
	BOOL SetUserData(tstring &sKey, tstring &sValue);
	DWORD SetUserFlags(DWORD dwFlags, BOOL bSet);

	// Inline-Funktionen
	int GetChildCount( ) { return m_ChildRows; }
	long GetNr( ) { return m_Nr; }
	CMTblRow * GetOrigParentRow( ) { return m_OrigParentRow; }
	CMTblRow * GetParentRow( ) { return m_ParentRow; }
	BOOL HasHiddenParentRow( ) { return m_ParentRow ? m_ParentRow->QueryFlags( MTBL_ROW_HIDDEN ) : FALSE; }
	BOOL HasParentRowChanged( ) { return m_ParentRow != m_OrigParentRow; }
	BOOL HasVisChildRows( ) { return m_ChildRowsVis > 0; }
	BOOL IsChild( ) { return m_ParentRow != NULL; }
	BOOL IsDelAdj( ) { return m_Bits.DelAdj; }
	BOOL IsDisabled( ) { return m_Bits.Disabled; }
	BOOL IsHdrOwnerdrawn( ) { return m_Bits.HdrOwnerdrawn; }
	BOOL IsNoSelInvImageSet( ) { return QueryFlags( MTBL_ROW_NOSELINV_IMAGE ); }
	BOOL IsNoSelInvTextSet( ) { return QueryFlags( MTBL_ROW_NOSELINV_TEXT ); }
	BOOL IsOwnerdrawn( ) { return m_Bits.Ownerdrawn; }
	BOOL IsParent( ) { return ( m_ChildRows > 0 ); }
	BOOL IsVisible( ) { return !( IsDelAdj( ) || QueryHiddenFlag( ) ); };
	void ParentRowDeleted( ) { m_ParentRow = NULL; }
	BOOL QueryFlags( DWORD dwFlags ) { return ( m_Flags & dwFlags ) != 0; }
	BOOL QueryHiddenFlag( ) { return m_Bits.HiddenFlag; }
	void SetDisabled( BOOL bSet ) { m_Bits.Disabled = ( bSet ? 1 : 0 ); }
	void SetHdrOwnerdrawn( BOOL bSet ) { m_Bits.HdrOwnerdrawn = ( bSet ? 1 : 0 ); }
	void SetHeight( long lHeight );
	void SetHiddenFlag( BOOL bSet ) { m_Bits.HiddenFlag = ( bSet ? 1 : 0 ); }
	void SetMerged( long lMerged );
	void SetNr( long lNr ) { m_Nr = lNr; }
	void SetOwnerdrawn( BOOL bSet ) { m_Bits.Ownerdrawn = ( bSet ? 1 : 0 ); }

};
#pragma pack()

#endif // _INC_CMTBLROW