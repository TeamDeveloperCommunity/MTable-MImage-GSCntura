// cmtbltip.cpp

// Includes
#include "cmtbltip.h"

// Konstruktoren
CMTblTip::CMTblTip( )
{
	m_iPosX = INT_MIN;
	m_iPosY = INT_MIN;
	m_iWid = INT_MIN;
	m_iHt = INT_MIN;
	m_iDelay = INT_MIN;
	m_iOffsX = INT_MIN;
	m_iOffsY = INT_MIN;
	m_iRCEWid = INT_MIN;
	m_iRCEHt = INT_MIN;
	m_iOpacity = INT_MIN;
	m_iFadeInTime = INT_MIN;
	m_wFlags = 0;
	m_pDefTip = NULL;
}

// Methoden
HFONT CMTblTip::CreateFont( HDC hDC, BOOL bMyValue )
{
	if ( !hDC ) return NULL;

	CMTblFont f = m_Font;
	if ( !bMyValue && m_pDefTip  )
	{		
		f.SubstUndef( m_pDefTip->m_Font );
	}

	return f.Create( hDC );
}

DWORD CMTblTip::GetBackColor( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_BackColor.IsUndef( ) )
		return m_pDefTip->m_BackColor.Get( );
	else	
		return m_BackColor.Get( );
}

int & CMTblTip::GetDelay( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iDelay == INT_MIN )
		return m_pDefTip->m_iDelay;
	else
		return m_iDelay;
}

__int64 CMTblTip::GetDelayI64( BOOL bMyValue )
{
	int iDelay = GetDelay( bMyValue );
	return (__int64)iDelay * 10000I64;
}

int & CMTblTip::GetFadeInTime( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iFadeInTime == INT_MIN )
		return m_pDefTip->m_iFadeInTime;
	else
		return m_iFadeInTime;
}

DWORD CMTblTip::GetFrameColor( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_FrameColor.IsUndef( ) )
		return m_pDefTip->m_FrameColor.Get( );
	else	
		return m_FrameColor.Get( );
}

int & CMTblTip::GetHt( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iHt == INT_MIN )
		return m_pDefTip->m_iHt;
	else
		return m_iHt;
}

int & CMTblTip::GetOffsX( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iOffsX == INT_MIN )
		return m_pDefTip->m_iOffsX;
	else
		return m_iOffsX;
}

int & CMTblTip::GetOffsY( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iOffsY == INT_MIN )
		return m_pDefTip->m_iOffsY;
	else
		return m_iOffsY;
}

int & CMTblTip::GetOpacity( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iOpacity == INT_MIN )
		return m_pDefTip->m_iOpacity;
	else
		return m_iOpacity;
}

int & CMTblTip::GetPosX( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iPosX == INT_MIN )
		return m_pDefTip->m_iPosX;
	else
		return m_iPosX;
}

int & CMTblTip::GetPosY( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iPosY == INT_MIN )
		return m_pDefTip->m_iPosY;
	else
		return m_iPosY;
}

int & CMTblTip::GetRCEHt( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iRCEHt == INT_MIN )
		return m_pDefTip->m_iRCEHt;
	else
		return m_iRCEHt;
}

int & CMTblTip::GetRCEWid( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iRCEWid == INT_MIN )
		return m_pDefTip->m_iRCEWid;
	else
		return m_iRCEWid;
}

tstring & CMTblTip::GetText( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_sText.size( ) == 0 )
		return m_pDefTip->m_sText;
	else	
		return m_sText;
}

DWORD CMTblTip::GetTextColor( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_TextColor.IsUndef( ) )
		return m_pDefTip->m_TextColor.Get( );
	else	
		return m_TextColor.Get( );
}

int & CMTblTip::GetWid( BOOL bMyValue )
{
	if ( !bMyValue && m_pDefTip && m_iWid == INT_MIN )
		return m_pDefTip->m_iWid;
	else
		return m_iWid;
}

BOOL CMTblTip::MustShow( )
{
	return ( GetText( ).size( ) > 0 );
}

BOOL CMTblTip::QueryFlags( DWORD dwFlags, BOOL bMyValue )
{
	if ( !bMyValue && m_wFlags & MTBL_TIP_FLAG_NODEFFLAGS )
		bMyValue = TRUE;

	if ( !bMyValue && m_pDefTip )
		return ( m_wFlags | m_pDefTip->m_wFlags ) & dwFlags;
	else
		return m_wFlags & dwFlags;
}

void CMTblTip::SetDefTip( CMTblTip *pDefTip )
{
	m_pDefTip = pDefTip;
}

BOOL CMTblTip::SetFlags( DWORD dwFlags, BOOL bSet )
{
	m_wFlags = ( bSet ? ( m_wFlags | dwFlags ) : ( ( m_wFlags & dwFlags ) ^ m_wFlags ) );
	return TRUE;
}
