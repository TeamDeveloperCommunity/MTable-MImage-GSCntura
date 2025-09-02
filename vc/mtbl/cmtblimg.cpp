// cmtblimg.cpp

// Includes
#include "cmtblimg.h"

// Konstruktoren
CMTblImg::CMTblImg( )
{
	InitMembers( );
}

// Destruktor
CMTblImg::~CMTblImg( )
{
	if ( m_hImage && m_pCounter )
	{
		if ( IsNoSelInvSet( ) )
			m_pCounter->DecNoSelInv( );
	}
}

CMTblImg & CMTblImg::operator =( CMTblImg & i )
{
	if ( &i != this )
	{
		m_hImage = i.m_hImage;
		m_wFlags = i.m_wFlags;
		m_pCounter = i.m_pCounter;
	}

	return *this;
}

// ClearFlags
void CMTblImg::ClearFlags( )
{
	BOOL bNoSelInvBefore = IsNoSelInvSet( );
	m_wFlags = 0;

	if ( IsValidImageHandle( m_hImage ) && m_pCounter )
	{
		if ( bNoSelInvBefore )
			m_pCounter->DecNoSelInv( );
	}
}

// Draw
BOOL CMTblImg::Draw( HDC hDC, LPRECT lpRect, DWORD dwFlags )
{
	if ( !IsValidImageHandle( m_hImage ) ) return FALSE;

	return MImgDraw( m_hImage, hDC, lpRect, dwFlags );
}

// GetSize
BOOL CMTblImg::GetSize( SIZE & s )
{
	if ( !IsValidImageHandle( m_hImage ) ) return FALSE;

	return MImgGetSize( m_hImage, &s );
}

// HasOpacity
BOOL CMTblImg::HasOpacity( )
{
	if ( !IsValidImageHandle( m_hImage ) ) return FALSE;

	return MImgHasOpacity( m_hImage );
}

// InitMembers
void CMTblImg::InitMembers( )
{
	m_hImage = NULL;
	m_wFlags = 0;
	m_pCounter = NULL;
}

// SetDrawClipRect
BOOL CMTblImg::SetDrawClipRect( LPRECT pRect )
{
	if ( !IsValidImageHandle( m_hImage ) ) return FALSE;
	/*if (!MImgIsValidHandle(m_hImage))
		return FALSE;*/
	return MImgSetDrawClipRect( m_hImage, pRect );
}

// SetFlags
BOOL CMTblImg::SetFlags( DWORD dwFlags, BOOL bSet )
{
	BOOL bNoSelInvBefore = IsNoSelInvSet( );
	m_wFlags = ( bSet ? ( m_wFlags | WORD(dwFlags) ) : ( ( m_wFlags & WORD(dwFlags) ) ^ m_wFlags ) );
	BOOL bNoSelInvAfter = IsNoSelInvSet( );

	if ( IsValidImageHandle( m_hImage ) && m_pCounter )
	{
		if ( bNoSelInvBefore && !bNoSelInvAfter )
			m_pCounter->DecNoSelInv( );
		else if ( !bNoSelInvBefore && bNoSelInvAfter )
			m_pCounter->IncNoSelInv( );
	}

	return TRUE;
}

// SetHandle
BOOL CMTblImg::SetHandle( HANDLE hImage, CMTblCounter *pCounter )
{
	m_pCounter = pCounter;
	m_hImage = hImage;
	return TRUE;
}
