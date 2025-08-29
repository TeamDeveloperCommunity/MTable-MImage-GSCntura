// cmtblcolhdrgrp.cpp

// Includes
#include "cmtblcolhdrgrp.h"

// Standard-Konstruktor
CMTblColHdrGrp::CMTblColHdrGrp( )
{
	InitMemberVars( );
}

// Weitere Konstruktoren
CMTblColHdrGrp::CMTblColHdrGrp( tstring &sText, list<HWND> &liCols, DWORD dwFlags /*=0*/ )
{
	InitMemberVars( );
	m_Text = sText;
	m_Cols.assign( liCols.begin( ), liCols.end( ) );
	m_dwFlags = dwFlags;
}

// Destruktor
CMTblColHdrGrp::~CMTblColHdrGrp( )
{
	if ( m_pTip )
		delete m_pTip;

	if (m_HorzLineDef)
		delete m_HorzLineDef;

	if (m_VertLineDef)
		delete m_VertLineDef;
}

// EnumCols
long CMTblColHdrGrp::EnumCols( HARRAY haCols )
{
	if ( !Ctd.IsPresent( ) ) return -1;

	// Check Array
	long lDims;
	if ( !Ctd.ArrayDimCount( haCols, &lDims ) ) return -1;
	if ( lDims > 1 ) return -1;

	// Untergrenzen Array ermitteln
	long lLow;
	if ( !Ctd.ArrayGetLowerBound( haCols, 1, &lLow ) ) return -1;

	// Daten ermitteln
	long lCount;
	list<HWND>::iterator it;
	for ( it = m_Cols.begin( ), lCount = 0; it != m_Cols.end( ); ++it )
	{
		if ( !Ctd.MDArrayPutHandle( haCols, HANDLE( *it ), lCount + lLow ) ) return -1;

		++lCount;
	}

	return lCount;
}

// GetLineDef
// Liefert eine Liniendefinition
LPMTBLLINEDEF CMTblColHdrGrp::GetLineDef(int iLineType)
{
	if (iLineType == MTLT_HORZ)
		return m_HorzLineDef;
	else if (iLineType == MTLT_VERT)
		return m_VertLineDef;
	else
		return NULL;
}

// GetNrOfTextLines
int CMTblColHdrGrp::GetNrOfTextLines( )
{
	int iLines = 1;
	int iSize = m_Text.size( );
	LPCTSTR lpsText = m_Text.c_str( );
	for ( int i = 0; i < iSize; ++i )
	{
		if ( lpsText[i] == 13 )
			++iLines;
	}

	return iLines;
}

// InitMemberVars
void CMTblColHdrGrp::InitMemberVars( )
{
	m_bOwnerdrawn = FALSE;
	m_dwFlags = 0;
	m_pTip = NULL;
	m_HorzLineDef = NULL;
	m_VertLineDef = NULL;
}

// IsColOf
BOOL CMTblColHdrGrp::IsColOf( HWND hWndCol )
{
	return find( m_Cols.begin( ), m_Cols.end( ), hWndCol ) != m_Cols.end( );
}

// SetCol
BOOL CMTblColHdrGrp::SetCol( HWND hWndCol, BOOL bSet )
{
	if ( !hWndCol )
		return FALSE;

	if ( bSet && !IsColOf( hWndCol ) )
		m_Cols.push_back( hWndCol );
	else if ( !bSet )
		m_Cols.remove( hWndCol );	

	return TRUE;
}

// SetFlags
DWORD CMTblColHdrGrp::SetFlags( DWORD dwFlags, BOOL bSet )
{
	DWORD dwOldFlags = m_dwFlags;
	m_dwFlags = ( bSet ? ( m_dwFlags | dwFlags ) : ( ( m_dwFlags & dwFlags ) ^ m_dwFlags ) );

	return ( dwOldFlags ^ m_dwFlags );
}

// SetLineDef
// Setzt eine Liniendefinition
BOOL CMTblColHdrGrp::SetLineDef(CMTblLineDef & ld, int iLineType)
{
	BOOL bOk = TRUE;

	if (iLineType == MTLT_HORZ)
	{
		if (!m_HorzLineDef)
			m_HorzLineDef = new CMTblLineDef;
		*m_HorzLineDef = ld;
	}

	else if (iLineType == MTLT_VERT)
	{
		if (!m_VertLineDef)
			m_VertLineDef = new CMTblLineDef;
		*m_VertLineDef = ld;
	}
	else
		bOk = FALSE;

	return bOk;
}
