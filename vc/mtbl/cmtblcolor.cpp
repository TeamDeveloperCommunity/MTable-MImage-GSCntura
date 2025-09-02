// cmtblcolor.cpp

// Includes
#include "cmtblcolor.h"

// Konstanten initialisieren
const DWORD CMTblColor::UNDEF = ULONG_MAX;

// Konstruktor
CMTblColor::CMTblColor( ) : m_Color(UNDEF)
{

}

// Destruktor
CMTblColor::~CMTblColor( )
{

}

// Operatoren
BOOL CMTblColor::operator ==( CMTblColor & clr )
{
	if ( this == &clr )
		return TRUE;

	return m_Color == clr.m_Color;
}
