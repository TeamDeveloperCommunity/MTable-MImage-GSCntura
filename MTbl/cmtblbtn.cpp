// cmtblbtn.cpp

// Includes
#include "cmtblbtn.h"

// Standard-Konstruktor
CMTblBtn::CMTblBtn( ) : m_bShow(FALSE), m_iWidth(0), m_sText(_T("")), m_hImage(NULL), m_dwFlags(0)
{

}

// Destruktor
CMTblBtn::~CMTblBtn( )
{

}

// Operator =
CMTblBtn & CMTblBtn::operator =( CMTblBtn & b )
{
	if ( this == &b )
		return *this;

	m_bShow = b.m_bShow;
	m_iWidth = b.m_iWidth;
	m_sText = b.m_sText;
	m_hImage = b.m_hImage;
	m_dwFlags = b.m_dwFlags;

	return *this;
}