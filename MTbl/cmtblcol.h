// cmtblcol.h

#ifndef _INC_CMTBLCOL
#define _INC_CMTBLCOL

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <list>
#include "tstring.h"
#include "cmtblfont.h"
#include "cmtblbtn.h"
#include "cmtblhylnk.h"
#include "cmtblimg.h"
#include "cmtblcolor.h"
#include "cmtbltip.h"
#include "cmtblcounter.h"
#include "colflags.h"
#include "cellflags.h"
#include "cmtblcelltype.h"
#include "lines.h"
#include "color.h"

// Using
using namespace std;

// Typen
#define COL_ITYPE_COL 0
#define COL_ITYPE_COLHDR 1
#define COL_ITYPE_CELL 2

// Interne Column-Flags
#define COL_IFLAG_SELECTED_BEFORE 1
#define COL_IFLAG_SELECTED_AFTER 2

// Klasse CMTblCol
#pragma pack(1)
class CMTblCol
{
public:
	HWND m_hWnd;					// Window-Handle
	BYTE m_nType;					// Typ
	CMTblCellType* m_pCellType;		// Zelltyp der Spalte
	CMTblCellType* m_pCellTypeCurr;	// Aktueller Zelltyp der Spalte
	CMTblColor m_TextColor;			// Textfarbe
	CMTblColor m_TextColorSal;		// Textfarbe Sal
	CMTblColor m_BackColor;			// Hintergrundfarbe
	CMTblBtn m_Btn;					// Button
	CMTblFont m_Font;				// Font
	CMTblHyLnk m_HyLnk;				// Hyperlink
	CMTblImg m_Image;				// Bild
	CMTblImg m_Image2;				// Bild 2
	DWORD m_dwFlags;				// Flags
	long m_DropDownListContext;		// Kontext DropDownList
	long m_Merged;					// Anzahl verbundener Spalten
	long m_MergedRows;				// Anzahl verbundener Zeilen
	CMTblTip* m_pTip;				// Tip
	CMTblTip* m_pTip_Text;			// Tip für Text
	CMTblTip* m_pTip_CompleteText;	// Tip für Anzeige des kompletten Texts
	CMTblTip* m_pTip_Image;			// Tip für Image
	CMTblTip* m_pTip_Button;		// Tip für Button
	long m_IndentLevel;				// Einrückungs-Ebene
	WORD m_InternalFlags;			// Interne Flags für unterschiedliche Verwendungszwecke
	CMTblLineDef* m_pHorzLineDef;	// Definition horizontale Linien
	CMTblLineDef* m_pVertLineDef;	// Definition vertikale Linien

	
	// Konstruktoren
	CMTblCol( );
	CMTblCol( HWND hWndCol, BYTE nType, CMTblCounter *pCounter );
	// Destruktor
	~CMTblCol( );
	// == Operatoren
	BOOL operator==( CMTblCol c );
	BOOL operator==( HWND hWndCol );
	// Funktionen
	long GetDropDownListContext( ) { return m_DropDownListContext; }
	CMTblLineDef * GetLineDef(int iLineType);
	BOOL IsBtnPushed( ) { return m_Bits.BtnPushed; }
	BOOL IsDisabled( ) { return m_Bits.Disabled; }
	BOOL IsNoSelInvBkgndSet( );
	BOOL IsNoSelInvImageSet( );
	BOOL IsNoSelInvTextSet( );
	BOOL IsOwnerdrawn( ) { return m_Bits.Ownerdrawn; }
	int QueryCellType( ) { return ( m_pCellType ? m_pCellType->m_iType : COL_CellType_Undefined ); }
	BOOL QueryFlags( DWORD dwFlags ) { return ( m_dwFlags & dwFlags ) != 0; }
	BOOL QueryInternalFlags( DWORD dwFlags ) { return ( m_InternalFlags & dwFlags ) != 0; }
	BOOL SetBtnPushed( BOOL bSet ) { m_Bits.BtnPushed = ( bSet ? 1 : 0 ); return TRUE; }
	BOOL SetDisabled( BOOL bSet ) { m_Bits.Disabled = ( bSet ? 1 : 0 ); return TRUE; }
	void SetDropDownListContext( long lRow ) { m_DropDownListContext = lRow; }
	DWORD SetFlags( DWORD dwFlags, BOOL bSet );
	DWORD SetInternalFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetLineDef(CMTblLineDef& ld, int iLineType);
	void SetMerged( long lMerged );
	void SetMergedRows( long lMergedRows );
	BOOL SetOwnerdrawn( BOOL bSet ) { m_Bits.Ownerdrawn = ( bSet ? 1 : 0 ); return TRUE; }
protected:
	struct tagBits
	{
		BYTE Disabled : 1;		// Durch M!Table-Funktionalität disabled ja/nein
		BYTE Ownerdrawn : 1;	// Benutzerdefiniert gezeichnet ja/nein
		BYTE BtnPushed : 1;     // M!Table-Button gedrückt?
		BYTE : 0;				// Für Speicher-Ausrichtung 
	}m_Bits;
	CMTblCounter *m_pCounter;	// Zähler
	// Funktionen
	void InitMembers( );
};
#pragma pack()

#endif // _INC_CMTBLCOL