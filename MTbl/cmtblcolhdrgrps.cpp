// cmtblcolhdrgrps.cpp

// Includes
#include "cmtblcolhdrgrps.h"

// Standard-Konstruktor
CMTblColHdrGrps::CMTblColHdrGrps( )
{

}

// Destruktor
CMTblColHdrGrps::~CMTblColHdrGrps( )
{
	DeleteAll( );
}

// CreateGrp
CMTblColHdrGrp * CMTblColHdrGrps::CreateGrp( tstring &sText, list<HWND> &liCols, DWORD dwFlags /*=0*/ )
{
	CMTblColHdrGrp *pGrp;

	// Eine der Spalten schon in einer Gruppe vorhanden?
	list<HWND>::iterator it, itEnd = liCols.end( );
	for ( it = liCols.begin( ); it != itEnd; ++it )
	{
		if ( pGrp = GetGrpOfCol( *it ) )
			pGrp->SetCol( *it, FALSE );
	}

	pGrp = new CMTblColHdrGrp( sText, liCols, dwFlags );
	m_List.push_back( pGrp );

	return pGrp;
}

// DeleteAll
void CMTblColHdrGrps::DeleteAll( )
{
	MTblColHdrGrpList::iterator it, itEnd = m_List.end( );

	for ( it = m_List.begin( ); it != itEnd; ++it )
	{
		delete *it;
	}

	m_List.clear( );
}

// DeleteGrp
BOOL CMTblColHdrGrps::DeleteGrp( CMTblColHdrGrp *pGrp )
{
	if ( !pGrp ) return FALSE;

	MTblColHdrGrpList::iterator it = find( m_List.begin( ), m_List.end( ), pGrp );
	if ( it == m_List.end( ) ) return FALSE;

	delete *it;

	m_List.erase( it );

	return TRUE;
}

// EnumCols
long CMTblColHdrGrps::EnumGrps( HARRAY haGrps )
{
	if ( !Ctd.IsPresent( ) ) return -1;

	// Check Array
	long lDims;
	if ( !Ctd.ArrayDimCount( haGrps, &lDims ) ) return -1;
	if ( lDims > 1 ) return -1;

	// Untergrenzen Array ermitteln
	long lLow;
	if ( !Ctd.ArrayGetLowerBound( haGrps, 1, &lLow ) ) return -1;

	// Daten ermitteln
	long lCount;
	MTblColHdrGrpList::iterator it;
	NUMBER n;
	for ( it = m_List.begin( ), lCount = 0; it != m_List.end( ); ++it )
	{
		if ( !Ctd.CvtULongToNumber( ULONG( *it ), &n ) ) return -1;
		if ( !Ctd.MDArrayPutNumber( haGrps, &n, lCount + lLow ) ) return -1;

		++lCount;
	}

	return lCount;
}

// GetGrpOfCol
CMTblColHdrGrp * CMTblColHdrGrps::GetGrpOfCol( HWND hWndCol )
{
	CMTblColHdrGrp *pGrp = NULL;

	MTblColHdrGrpList::iterator it, itEnd = m_List.end( );
	for ( it = m_List.begin( ); it != itEnd; ++it )
	{
		if ( (*it)->IsColOf( hWndCol ) )
		{
			pGrp = *it;
			break;
		}
	}

	return pGrp;
}

// IsValidGrp
BOOL CMTblColHdrGrps::IsValidGrp( CMTblColHdrGrp *pGrp )
{
	if ( !pGrp ) return FALSE;
	return ( find( m_List.begin( ), m_List.end( ), pGrp ) != m_List.end( ) );
}

// SetCol
BOOL CMTblColHdrGrps::SetCol( CMTblColHdrGrp *pGrp, HWND hWndCol, BOOL bSet )
{
	if ( !IsValidGrp( pGrp ) )
		return FALSE;

	if ( !hWndCol )
		return FALSE;

	if ( bSet )
	{
		CMTblColHdrGrp *pCurGrp = GetGrpOfCol( hWndCol );
		if ( pCurGrp && pCurGrp != pGrp )
			pCurGrp->SetCol( hWndCol, FALSE );

		if ( !pGrp->SetCol( hWndCol, TRUE ) )
			return FALSE;
	}
	else
		if ( !pGrp->SetCol( hWndCol, FALSE ) )
			return FALSE;

	return TRUE;
}

