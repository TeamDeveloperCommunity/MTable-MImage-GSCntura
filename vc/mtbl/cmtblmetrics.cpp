// cmtblmetrics.cpp

// Includes
#include "cmtblmetrics.h"

// Konstruktoren
CMTblMetrics::CMTblMetrics( void )
{

}

CMTblMetrics::CMTblMetrics( HWND hWndTbl )
{
	Get( hWndTbl );
}


// Destruktor
CMTblMetrics::~CMTblMetrics( void )
{

}


// CalcRowHeight
// Berechnet die Zeilenhöhe nahnd einer best. Anzahl LinesPerRow
int CMTblMetrics::CalcRowHeight( int iLinesPerRow )
{
	return ( max( iLinesPerRow, 1 ) * m_CharBoxY ) + ( m_CellHeight - m_CharBoxY );
}


// Get
// Ermittelt die Werte für eine Tabelle
void CMTblMetrics::Get( HWND hWndTbl, BOOL bNoVisRange /*= FALSE*/ )
{
	if ( !GetClientRect( hWndTbl, &m_ClientRect ) ) return;

	MTBLGETMETRICS m;
	SendMessage( hWndTbl, TIM_GetMetrics, 0, (LPARAM)&m );
	m_HeaderHeight = m.iHeaderHeight;
	m_CellHeight = m.iCellHeight;
	m_CharBoxX = m.ptCharBox.x;
	m_CharBoxY = m.ptCharBox.y;
	m_LineSizeX = m.ptLineSize.x;
	m_LineSizeY = m.ptLineSize.y;
	m_CellLeadingX = m.ptCellLeading.x;
	m_CellLeadingY = m.ptCellLeading.y;
	m_SplitBarTop = m.iSplitBarTop;
	m_SplitBarHeight = m.iSplitBarHeight;
	m_LockedColumnsWidth = m.iLockedColumnsWidth;
	
	BOOL bSplitted = FALSE;

	long lMsgRet = SendMessage( hWndTbl, SAM_User + guiMsgOffs /*MTM_Internal*/, 128 /*IA_GET_SPLITTED*/, 0 );
	if ( lMsgRet )
	{
		bSplitted = ( lMsgRet > 0 ? TRUE : FALSE );
	}
	else
	{
		BOOL bDragAdjust;
		int iSplitRows = 0;
		Ctd.TblQuerySplitWindow( hWndTbl, &iSplitRows, &bDragAdjust );
		bSplitted = ( iSplitRows > 0 );
	}
	
	if ( bSplitted )
	{
		m_ClientRect.right -= GetSystemMetrics( SM_CXVSCROLL );
	}
	else
	{		
		m_SplitBarTop = 0;
		m_SplitBarHeight = 0;
	}

	Ctd.TblQueryLinesPerRow( hWndTbl, &m_LinesPerRow );

	m_RowHeight = CalcRowHeight( m_LinesPerRow );

	m_RowsTop = m_HeaderHeight;
	m_RowsBottom = ( m_SplitBarTop > 0 ) ? m_SplitBarTop : m_ClientRect.bottom;

	m_SplitRowsTop = m_SplitBarTop + m_SplitBarHeight;
	m_SplitRowsBottom = ( m_SplitBarTop > 0 ) ? m_ClientRect.bottom : 0;

	HSTRING hs = 0;
	HWND hWndCol;
	int iWidth;
	WORD wFlags;
	Ctd.TblQueryRowHeader( hWndTbl, &hs, 1, &iWidth, &wFlags, &hWndCol );
	m_RowHeaderLeft = 0;
	m_RowHeaderRight = ( wFlags & TBL_RowHdr_Visible ) ? iWidth : 0;
	if ( hs )
		Ctd.HStringUnRef( hs );

	m_LockedColumns = Ctd.TblQueryLockedColumns( hWndTbl );

	if ( m_LockedColumnsWidth > 0 )
	{
		m_LockedColumnsLeft = m_RowHeaderRight;
		m_LockedColumnsRight = m_RowHeaderRight + m_LockedColumnsWidth - 2;
		m_ColumnsLeft = m_RowHeaderRight + m_LockedColumnsWidth - 1;
		m_ColumnsRight = m_ClientRect.right;
	}
	else
	{
		m_LockedColumnsLeft = 0;
		m_LockedColumnsRight = 0;
		m_ColumnsLeft = m_RowHeaderRight;
		m_ColumnsRight = m_ClientRect.right;
	}

	LRESULT lRet = SendMessage( hWndTbl, TIM_QueryPixelOrigin, 0, 0 );
	m_ColumnOrigin = (int)LOWORD( lRet );

	if ( !bNoVisRange )
	{
		Ctd.TblQueryVisibleRange( hWndTbl, &m_MinVisibleRow, &m_MaxVisibleRow );
		m_MaxVisibleRow += 1;
	}
	else
		m_MinVisibleRow = m_MaxVisibleRow = 0;

}

long CMTblMetrics::QueryCheckBoxSize( HWND hWndTbl )
{
	static long s_lSize = 0;

	if ( s_lSize == 0 )
	{
		s_lSize = 13;
		HDC hDC = GetDC( NULL );
		if ( hDC )
		{
			int iLPY = GetDeviceCaps( hDC, LOGPIXELSY );
			if ( iLPY )
			{
				s_lSize = long(13.0 * double(iLPY) / 96.0);
			}
			ReleaseDC( NULL, hDC );
		}
	}

	return s_lSize;
}