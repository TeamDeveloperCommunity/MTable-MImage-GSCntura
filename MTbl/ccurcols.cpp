// ccurcols.cpp

// Includes
#include "ccurcols.h"

// Klasse CCurCols
CCurCols::CCurCols()
{
	m_hWndTbl = NULL;
	m_bValid = FALSE;
}

CCurCols::CCurCols( HWND hWndTbl )
{
	m_hWndTbl = hWndTbl;
	m_bValid = FALSE;
	Refresh( );
}

int CCurCols::GetPos( HWND hWndCol )
{
	if ( !m_bValid )
		Refresh( );

	CURCOLS_POS_MAP::iterator it = m_pm.find( hWndCol );
	if ( it == m_pm.end( ) )
		return -1;
	else
		return (*it).second;
}


CURCOLS_HANDLE_VECTOR & CCurCols::GetVectorHandle( )
{
	if ( !m_bValid )
		Refresh( );

	return m_vh;
}

CURCOLS_WIDTH_VECTOR & CCurCols::GetVectorWidth( )
{
	if ( !m_bValid )
		Refresh( );

	return m_vw;
}

void CCurCols::Refresh( )
{
	m_vh.clear( );
	m_vw.clear( );
	m_pm.clear( );

	HWND hWndCol;
	int iPos = 1;
	long lWidth;
	while( hWndCol = Ctd.TblGetColumnWindow( m_hWndTbl, iPos, COL_GetPos ) )
	{
		m_vh.push_back( hWndCol );

		lWidth = SendMessage( hWndCol, TIM_QueryColumnWidth, 0, 0 );
		m_vw.push_back( lWidth );

		m_pm.insert( CURCOLS_POS_MAP::value_type( hWndCol, iPos ) );

		++iPos;
	}

	m_bValid = TRUE;
}

void CCurCols::UpdateWidth( HWND hWndCol )
{
	if ( hWndCol )
	{
		int iPos = GetPos( hWndCol );
		if ( iPos > 0 )
		{
			m_vw[iPos-1] = SendMessage( hWndCol, TIM_QueryColumnWidth, 0, 0 );
		}
	}
}
