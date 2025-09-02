// cmtblcolor.h

#ifndef _INC_CMTBLCOLOR
#define _INC_CMTBLCOLOR

// Includes
#include <windows.h>
#include <limits.h>
#include "ctd.h"

// Globale Variablen
extern CCtd Ctd;

// Klasse CMTblColor
class CMTblColor
{
private:
	COLORREF m_Color;		// Farbe
public:
	// Konstanten
	static const DWORD UNDEF;
	// Konstruktor
	CMTblColor( );
	// Destruktor
	~CMTblColor( );

	// Operatoren
	BOOL operator ==( CMTblColor & clr );

	// Funktionen
	COLORREF Get( ) { return m_Color; };
	BOOL IsUndef( ) { return m_Color == UNDEF; };
	void Set( COLORREF c ) { m_Color = c; };
};

#endif // _INC_CMTBLCOLOR