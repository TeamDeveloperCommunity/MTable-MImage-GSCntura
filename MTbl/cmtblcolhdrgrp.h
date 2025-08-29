// cmtblcolhdrgrp.h

#ifndef _INC_CMTBLCOLHDRGRP
#define _INC_CMTBLCOLHDRGRP

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <list>
#include "tstring.h"
#include "cmtblbtn.h"
#include "cmtblfont.h"
#include "cmtblcolor.h"
#include "cmtbltip.h"
#include "cmtblimg.h"
#include "lines.h"

// Using
using namespace std;

// Klasse CMTblColHdrGrp
class CMTblColHdrGrp
{
private:
	// Funktionen
	void InitMemberVars( );
public:
	tstring m_Text;					// Text
	CMTblColor m_TextColor;			// Textfarbe
	CMTblColor m_BackColor;			// Hintergrundfarbe
	CMTblBtn m_Btn;					// Button
	CMTblFont m_Font;				// Font
	CMTblImg m_Image;				// Bild
	CMTblTip* m_pTip;				// Tip
	CMTblLineDef* m_HorzLineDef;	// Definition horizontale Linien
	CMTblLineDef* m_VertLineDef;	// Definition vertikale Linien
	list<HWND> m_Cols;				// Spalten
	DWORD m_dwFlags;				// Flags

	// Konstruktoren
	CMTblColHdrGrp( );
	CMTblColHdrGrp( tstring &sText, list<HWND> &liCols, DWORD dwFlags = 0 );
	// Destruktor
	~CMTblColHdrGrp( );
	// Funktionen
	long EnumCols(HARRAY haCols);
	CMTblLineDef * GetLineDef(int iLineType);
	int GetNrOfTextLines( );
	BOOL IsColOf( HWND hWndCol );
	BOOL IsOwnerdrawn( ) { return m_bOwnerdrawn; }
	BOOL QueryFlags( DWORD dwFlags ) { return m_dwFlags & dwFlags; };
	BOOL SetCol( HWND hWndCol, BOOL bSet );
	DWORD SetFlags(DWORD dwFlags, BOOL bSet);
	BOOL SetLineDef(CMTblLineDef& ld, int iLineType);
	void SetOwnerdrawn( BOOL bSet ) { m_bOwnerdrawn = bSet ? true : false; }
private:	
	bool m_bOwnerdrawn;			// Benutzerdef. gezeichnet ja/nein
};

#endif // _INC_CMTBLCOLHDRGRP