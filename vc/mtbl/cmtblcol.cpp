// cmtblcol.cpp

// Includes
#include "cmtblcol.h"

// Standard-Konstruktor
CMTblCol::CMTblCol( )
{
	InitMembers( );
}

// Weitere Konstruktoren
CMTblCol::CMTblCol( HWND hWndCol, BYTE nType, CMTblCounter *pCounter )
{
	InitMembers( );
	m_hWnd = hWndCol;
	m_nType = nType;
	m_pCounter = pCounter;
}

// Destruktor
CMTblCol::~CMTblCol( )
{
	if ( m_pCounter )
	{
		if ( IsNoSelInvBkgndSet( ) )
			m_pCounter->DecNoSelInv( );

		if ( IsNoSelInvImageSet( ) )
			m_pCounter->DecNoSelInv( );

		if ( IsNoSelInvTextSet( ) )
			m_pCounter->DecNoSelInv( );

		if ( m_Merged > 0 )
			m_pCounter->DecMergeCells( );

		if ( m_MergedRows > 0 )
			m_pCounter->DecMergeCells( );
	}

	if ( m_pCellType )
		delete m_pCellType;

	if ( m_pTip )
		delete m_pTip;

	if ( m_pTip_Button )
		delete m_pTip_Button;

	if (m_pTip_CompleteText)
		delete m_pTip_CompleteText;
		

	if ( m_pTip_Image )
		delete m_pTip_Image;

	if ( m_pTip_Text )
		delete m_pTip_Text;


	if (m_pHorzLineDef)
		delete m_pHorzLineDef;

	if (m_pVertLineDef)
		delete m_pVertLineDef;
}

// == Operatoren
BOOL CMTblCol::operator==( CMTblCol c )
{
	if ( this == &c ) return TRUE;
	return ( m_hWnd == c.m_hWnd );
}

BOOL CMTblCol::operator==( HWND hWndCol )
{
	return ( m_hWnd == hWndCol );
}


// GetLineDef
// Liefert eine Liniendefinition
LPMTBLLINEDEF CMTblCol::GetLineDef(int iLineType)
{
	if (iLineType == MTLT_HORZ)
		return m_pHorzLineDef;
	else if (iLineType == MTLT_VERT)
		return m_pVertLineDef;
	else
		return NULL;
}

// SetFlags
DWORD CMTblCol::SetFlags( DWORD dwFlags, BOOL bSet )
{
	BOOL bNoSelInvBkgndBefore = IsNoSelInvBkgndSet( );
	BOOL bNoSelInvImageBefore = IsNoSelInvImageSet( );
	BOOL bNoSelInvTextBefore = IsNoSelInvTextSet( );

	DWORD dwOldFlags = m_dwFlags;
	m_dwFlags = ( bSet ? ( m_dwFlags | dwFlags ) : ( ( m_dwFlags & dwFlags ) ^ m_dwFlags ) );
	DWORD dwDiff = ( dwOldFlags ^ m_dwFlags );

	if ( dwDiff && m_pCounter )
	{
		BOOL bNoSelInvBkgnd = IsNoSelInvBkgndSet( );
		if ( bNoSelInvBkgnd != bNoSelInvBkgndBefore )
		{
			long lInc = ( bNoSelInvBkgnd && !bNoSelInvBkgndBefore ) ? 1 : -1;
			m_pCounter->IncNoSelInv( lInc );
		}

		BOOL bNoSelInvImage = IsNoSelInvImageSet( );
		if ( bNoSelInvImage != bNoSelInvImageBefore )
		{
			long lInc = ( bNoSelInvImage && !bNoSelInvImageBefore ) ? 1 : -1;
			m_pCounter->IncNoSelInv( lInc );
		}

		BOOL bNoSelInvText = IsNoSelInvTextSet( );
		if ( bNoSelInvText != bNoSelInvTextBefore )
		{
			long lInc = ( bNoSelInvText && !bNoSelInvTextBefore ) ? 1 : -1;
			m_pCounter->IncNoSelInv( lInc );
		}
	}
	
	return dwDiff;
}

// SetInternalFlags
DWORD CMTblCol::SetInternalFlags( DWORD dwFlags, BOOL bSet )
{
	WORD wOldFlags = m_InternalFlags;
	m_InternalFlags = ( bSet ? ( m_InternalFlags | dwFlags ) : ( ( m_InternalFlags & dwFlags ) ^ m_InternalFlags ) );

	return ( wOldFlags ^ m_InternalFlags );
}


// SetLineDef
// Setzt eine Liniendefinition
BOOL CMTblCol::SetLineDef(CMTblLineDef & ld, int iLineType)
{
	BOOL bOk = TRUE;

	if (iLineType == MTLT_HORZ)
	{
		if (!m_pHorzLineDef)
			m_pHorzLineDef = new CMTblLineDef;
		*m_pHorzLineDef = ld;
	}
		
	else if (iLineType == MTLT_VERT)
	{
		if (!m_pVertLineDef)
			m_pVertLineDef = new CMTblLineDef;
		*m_pVertLineDef = ld;
	}		
	else
		bOk = FALSE;

	return bOk;
}

void CMTblCol::SetMerged( long lMerged )
{
	lMerged = max( lMerged, 0 );

	if ( lMerged != m_Merged )
	{
		if ( m_pCounter )
		{
			if ( lMerged > 0 && m_Merged <= 0 )
				m_pCounter->IncMergeCells( );
			else if ( lMerged <= 0 && m_Merged > 0 )
				m_pCounter->DecMergeCells( );
		}
		m_Merged = lMerged;
	}
}

void CMTblCol::SetMergedRows( long lMergedRows )
{
	lMergedRows = max( lMergedRows, 0 );

	if ( lMergedRows != m_MergedRows )
	{
		if ( m_pCounter )
		{
			if ( lMergedRows > 0 && m_MergedRows <= 0 )
				m_pCounter->IncMergeCells( );
			else if ( lMergedRows <= 0 && m_MergedRows > 0 )
				m_pCounter->DecMergeCells( );

		}
		m_MergedRows = lMergedRows;
	}
}

// InitMemberVars
void CMTblCol::InitMembers( )
{
	m_hWnd = NULL;
	m_nType = COL_ITYPE_COL;
	ZeroMemory( &m_Bits, sizeof( m_Bits ) );
	m_dwFlags = 0;
	m_DropDownListContext = TBL_Error;
	m_Merged = m_MergedRows = 0;
	m_InternalFlags = 0;
	m_pCounter = NULL;
	m_pCellType = m_pCellTypeCurr = NULL;
	m_pTip = NULL;
	m_pTip_Button = NULL;
	m_pTip_CompleteText = NULL;
	m_pTip_Image = NULL;
	m_pTip_Text = NULL;
	m_IndentLevel = 0;
	m_pHorzLineDef = NULL;
	m_pVertLineDef = NULL;
}

// IsNoSelInvBkgndSet
BOOL CMTblCol::IsNoSelInvBkgndSet( )
{
	switch ( m_nType )
	{
	case COL_ITYPE_COL:
		return QueryFlags( MTBL_COL_FLAG_NOSELINV_BKGND );
	case COL_ITYPE_CELL:
		return QueryFlags( MTBL_CELL_FLAG_NOSELINV_BKGND );
	default:
		return FALSE;
	}
}

// IsNoSelInvImageSet
BOOL CMTblCol::IsNoSelInvImageSet( )
{
	switch ( m_nType )
	{
	case COL_ITYPE_COL:
		return QueryFlags( MTBL_COL_FLAG_NOSELINV_IMAGE );
	case COL_ITYPE_CELL:
		return QueryFlags( MTBL_CELL_FLAG_NOSELINV_IMAGE );
	default:
		return FALSE;
	}
}

// IsNoSelInvTextSet
BOOL CMTblCol::IsNoSelInvTextSet( )
{
	switch ( m_nType )
	{
	case COL_ITYPE_COL:
		return QueryFlags( MTBL_COL_FLAG_NOSELINV_TEXT );
	case COL_ITYPE_CELL:
		return QueryFlags( MTBL_CELL_FLAG_NOSELINV_TEXT );
	default:
		return FALSE;
	}
}
