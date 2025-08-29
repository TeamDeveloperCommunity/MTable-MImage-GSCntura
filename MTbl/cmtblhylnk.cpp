// cmtblhylnk.cpp

// Includes
#include "cmtblhylnk.h"

// Standard-Konstruktor
CMTblHyLnk::CMTblHyLnk( ) : m_wFlags(0), m_bEnabled(FALSE)
{

}

// Destruktor
CMTblHyLnk::~CMTblHyLnk( )
{

}

// Operator =
CMTblHyLnk & CMTblHyLnk::operator =( CMTblHyLnk & hl )
{
	if ( this == &hl )
		return *this;

	m_bEnabled = hl.m_bEnabled;
	m_wFlags = hl.m_wFlags;

	return *this;
}
