// cmtblrows.cpp

// Includes
#include "cmtblrows.h"

// Konstruktoren
CMTblRows::CMTblRows( )
{
	InitMembers( );
}

CMTblRows::CMTblRows( CMTblCounter *pCounter )
{
	InitMembers( );
	m_pCounter = pCounter;
}

// Destruktor
CMTblRows::~CMTblRows( )
{
	// Zeilen freigeben
	FreeRows( );
	FreeSplitRows( );
}


// AnyMergeCellInCol
// Ermittelt, ob irgendeine Zelle gemergt ist
BOOL CMTblRows::AnyMergeCellInCol( HWND hWndCol )
{
	if ( !hWndCol )
		return FALSE;

	CMTblCol *pCol;
	int i;

	for ( i = 0; i <= m_LastValidPos; ++i )
	{
		if ( m_Rows[i] )
		{
			pCol = m_Rows[i]->m_Cols->Find( hWndCol );
			if ( pCol && pCol->m_Merged > 0 )
				return TRUE;
		}
	}

	for ( i = 0; i <= m_LastValidSplitPos; ++i )
	{
		if ( m_SplitRows[i] )
		{
			pCol = m_SplitRows[i]->m_Cols->Find( hWndCol );
			if ( pCol && pCol->m_Merged > 0 )
				return TRUE;
		}
	}

	return FALSE;
}


// Create Row
// Erzeugt eine Zeile
CMTblRow * CMTblRows::CreateRow( long lRow )
{
	long lPos;
	if ( lRow >= 0 )
	{
		lPos = lRow;
		// Gültige Position?
		if ( !IsValidPos( lPos ) ) return NULL;

		if ( !m_Rows[lPos] )
		{
			// Erzeugen
			m_Rows[lPos] = new CMTblRow( m_pCounter );
			if ( !m_Rows[lPos] ) return NULL;

			// Nr setzen
			m_Rows[lPos]->SetNr( lRow );

			// Höchste gültige Zeile setzen
			if ( lPos > m_LastValidPos )
				m_LastValidPos = lPos;
		}

		return m_Rows[lPos];
	}
	else
	{
		lPos = GetSplitRowPos( lRow );
		// Gültige Position?
		if ( !IsValidSplitPos( lPos ) ) return NULL;

		if ( !m_SplitRows[lPos] )
		{
			// Erzeugen
			m_SplitRows[lPos] = new CMTblRow( m_pCounter );
			if ( !m_SplitRows[lPos] ) return NULL;

			// Nr setzen
			m_SplitRows[lPos]->SetNr( lRow );

			// Höchste gültige Zeile setzen
			if ( lPos > m_LastValidSplitPos )
				m_LastValidSplitPos = lPos;
		}

		return m_SplitRows[lPos];
	}
}


// DeleteRow
// Löscht eine Zeile
BOOL CMTblRows::DeleteRow( long lRow, int iMode )
{
	CMTblRow * pDelRow = NULL;
	int i;

	// Anzahl Kindzeilen anpassen
	if ( GetParentRow( lRow ) != TBL_Error )
		m_ChildRows--;

	if ( iMode == TBL_Adjust )
	{
		// Sicherstellen, daß Zeile da ist
		if ( !EnsureValid( lRow ) ) return FALSE;
		// Flag setzen
		if ( SetRowDelAdj( lRow ) )
		{
			// Zeiger auf gelöschte Zeile ermitteln
			pDelRow = GetRowPtr( lRow );

			// Hat keine Elternzeile mehr
			pDelRow->SetParentRow( NULL );
		}
		else
			return FALSE;
	}
	else
	{
		// Zeile schon mit TBL_Adjust gelöscht?
		if ( IsRowDelAdj( lRow ) ) return FALSE;

		long lPos;
		int j;

		if ( lRow >= 0 )
		{
			lPos = lRow;
			// Größer als letzte gültige Zeilenposition?
			if ( lPos > m_LastValidPos ) return TRUE;
			// Gültige Position?
			if ( !IsValidPos( lPos ) ) return FALSE;
			// Zeile löschen
			if ( m_Rows[lPos] )
			{
				// Zeiger auf gelöschte Zeile merken
				pDelRow = m_Rows[lPos];

				// Hat keine Elternzeile mehr
				pDelRow->SetParentRow( NULL );

				// Löschen
				delete m_Rows[lPos];
				m_Rows[lPos] = NULL;
			}
			// Nachfolgende Zeilen nach vorne verschieben
			for ( i = lPos; i <= m_LastValidPos; ++i )
			{
				if ( !IsRowDelAdj( i ) )
				{
					for ( j = i + 1; j <= m_LastValidPos; ++j )
					{
						if ( !IsRowDelAdj( j ) ) break;
					}

					if ( j <= m_LastValidPos )
					{
						m_Rows[i] = m_Rows[j];
						m_Rows[j] = NULL;

						if ( m_Rows[i] )
							m_Rows[i]->SetNr( i );
					}
				}
			}
			// Letzte gültige Position ermitteln
			for ( i = m_LastValidPos; i >= 0; --i )
			{
				if ( m_Rows[i] )
					break;
			}
			m_LastValidPos = i;
		}
		else
		{
			lPos = GetSplitRowPos( lRow );
			// Größer als letzte gültige Zeilenposition?
			if ( lPos > m_LastValidSplitPos ) return TRUE;
			// Gültige Position?
			if ( !IsValidSplitPos( lPos ) ) return FALSE;
			// Zeile löschen
			if ( m_SplitRows[lPos] )
			{
				// Zeiger auf gelöschte Zeile merken
				pDelRow = m_SplitRows[lPos];

				// Hat keine Elternzeile mehr
				pDelRow->SetParentRow( NULL );

				// Löschen
				delete m_SplitRows[lPos];
				m_SplitRows[lPos] = NULL;
			}
			// Nachfolgende Zeilen nach vorne verschieben
			for ( i = lPos; i <= m_LastValidSplitPos; ++i )
			{
				if ( !IsRowDelAdj( GetSplitPosRow( i ) ) )
				{
					for ( j = i + 1; j <= m_LastValidSplitPos; ++j )
					{
						if ( !IsRowDelAdj( GetSplitPosRow( j ) ) ) break;
					}

					if ( j <= m_LastValidSplitPos )
					{
						m_SplitRows[i] = m_SplitRows[j];
						m_SplitRows[j] = NULL;

						if ( m_SplitRows[i] )
							m_SplitRows[i]->SetNr( GetSplitPosRow( i ) );
					}
				}
			}
			// Letzte gültige Position ermitteln
			for ( i = m_LastValidSplitPos; i >= 0; --i )
			{
				if ( m_SplitRows[i] )
					break;
			}
			m_LastValidSplitPos = i;
		}
	}

	if ( pDelRow && m_ChildRows > 0 )
	{
		// Kindzeilen bzw. ursprüngliche Kindzeilen anpassen
		long lLastValidPos, lUpperBound;
		CMTblRow *pRow;
		CMTblRow **RowArr = GetRowArray( lRow, &lUpperBound, &lLastValidPos );
		if ( RowArr )
		{
			for ( i = 0; i < lLastValidPos && m_ChildRows > 0; i++ )
			{
				if ( pRow = RowArr[i] )
				{
					if ( pRow->GetParentRow( ) == pDelRow )
					{
						if ( iMode == TBL_Adjust )
							pRow->SetParentRow( NULL );
						else
							pRow->ParentRowDeleted( );
						m_ChildRows--;
					}

					if ( pRow->GetOrigParentRow( ) == pDelRow )
					{
						// Aktuelle Elternzeile als ursprüngliche Elternzeile setzen
						pRow->SetOrigParentRow( pRow->GetParentRow( ) );
					}
					
				}
			}
		}


		/*if ( lRow >= 0 )
		{
			for ( i = 0; i < m_LastValidPos && m_ChildRows > 0; i++ )
			{
				if ( m_Rows[i] )
				{
					if ( m_Rows[i]->GetParentRow( ) == pDelRow )
					{
						m_Rows[i]->SetParentRow( NULL );
						m_ChildRows--;
					}

					if ( m_Rows[i]->GetOrigParentRow( ) == pDelRow )
					{
						m_Rows[i]->SetOrigParentRow( m_Rows[i]->GetParentRow( ) );
					}
					
				}
			}
		}
		else
		{
			for ( i = 0; i < m_LastValidSplitPos && m_ChildRows > 0; i++ )
			{
				if ( m_SplitRows[i] )
				{
					if ( m_SplitRows[i]->GetParentRow( ) == pDelRow )
					{
						m_SplitRows[i]->SetParentRow( NULL );
						m_ChildRows--;
					}

					if ( m_SplitRows[i]->GetOrigParentRow( ) == pDelRow )
					{
						m_SplitRows[i]->SetOrigParentRow( m_SplitRows[i]->GetParentRow( ) );
					}
					
				}
			}
		}*/
	}

	return TRUE;
}


// DelRowUserData
// Löscht Benutzerdaten einer Zeile
BOOL CMTblRows::DelRowUserData( long lRow, tstring &sKey )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return TRUE;

	return pRow->DelUserData( sKey );
}


// EnsureValid
// Stellt sicher, daß eine gültige Zeile da ist
CMTblRow * CMTblRows::EnsureValid( long lRow )
{
	if ( lRow >= 0 )
	{
		if ( !EnsureValidPos( lRow ) )
			return NULL;
	}
	else
	{
		if ( !EnsureValidSplitPos( GetSplitRowPos( lRow ) ) )
			return NULL;
	}

	return CreateRow( lRow );
}


// EnsureValidPos
// Stellt sicher, daß eine Zeilenposition gültig ist
BOOL CMTblRows::EnsureValidPos( long lPos )
{
	// Position schon gültig?
	if ( IsValidPos( lPos ) ) return TRUE;

	// Position positiv?
	if ( lPos >= 0 )
	{
		// Noch kein Speicher reserviert?
		if ( !m_Rows )
		{
			m_Rows = (CMTblRow**)VirtualAlloc( NULL, MTBL_MAXROWSINARRAY * sizeof( CMTblRow* ), MEM_RESERVE, PAGE_READWRITE );
			if ( !m_Rows ) return FALSE;
		}

		// Neue Anzahl Zeilen errechnen
		long lRows = ( ( lPos / 100 ) + 1 ) * 100;
		if ( lRows > MTBL_MAXROWSINARRAY ) return FALSE;

		// Gültigen Speicher besorgen
		LPVOID pStart = m_Rows + m_UpperBound + 1;
		DWORD dwSize = ( lRows - ( m_UpperBound + 1 ) ) * sizeof( CMTblRow* );
		if ( !VirtualAlloc( pStart, dwSize, MEM_COMMIT, PAGE_READWRITE ) ) return FALSE;

		// Neue Höchstgrenze setzen
		m_UpperBound = lRows - 1;

		return TRUE;
	}
	else
		return FALSE;
}


// EnsureValidSplitPos
// Stellt sicher, daß eine Split-Zeilenposition gültig ist
BOOL CMTblRows::EnsureValidSplitPos( long lPos )
{
	// Position schon gültig?
	if ( IsValidSplitPos( lPos ) ) return TRUE;

	// Position positiv?
	if ( lPos >= 0 )
	{
		// Noch kein Speicher reserviert?
		if ( !m_SplitRows )
		{
			m_SplitRows = (CMTblRow**)VirtualAlloc( NULL, MTBL_MAXROWSINARRAY * sizeof( CMTblRow* ), MEM_RESERVE, PAGE_READWRITE );
			if ( !m_SplitRows ) return FALSE;
		}

		// Neue Anzahl Zeilen errechnen
		long lRows = ( ( lPos / 100 ) + 1 ) * 100;
		if ( lRows > MTBL_MAXROWSINARRAY ) return FALSE;

		// Gültigen Speicher besorgen
		LPVOID pStart = m_SplitRows + m_SplitUpperBound + 1;
		DWORD dwSize = ( lRows - ( m_SplitUpperBound + 1 ) ) * sizeof( CMTblRow* );
		if ( !VirtualAlloc( pStart, dwSize, MEM_COMMIT, PAGE_READWRITE ) ) return FALSE;

		// Neue Höchstgrenze setzen
		m_SplitUpperBound = lRows - 1;

		return TRUE;
	}
	else
		return FALSE;
}



long CMTblRows::EnumRowUserData( long lRow, HARRAY haKeys, HARRAY haValues )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return FALSE;

	return pRow->EnumUserData( haKeys, haValues );
}


// FindNextRow
//Sucht die nächste Zeile
CMTblRow * CMTblRows::FindNextRow( LPLONG plRow )
{
	if ( !plRow )
		return NULL;

	for ( long lPos = ( *plRow < 0 ? 0 : *plRow + 1 ); lPos <= m_LastValidPos; ++lPos )
	{
		if ( m_Rows[lPos] && !m_Rows[lPos]->IsDelAdj( ) )
		{
			*plRow = lPos;
			return m_Rows[lPos];
		}
	}

	return NULL;
}


// FindNextSplitRow
// Sucht die nächste Split-Zeile
CMTblRow * CMTblRows::FindNextSplitRow( LPLONG plRow )
{
	if ( !plRow )
		return NULL;
	
	long lPos = GetSplitRowPos( *plRow );
	if ( lPos < 0 )
		lPos = 0;
	else
		++lPos;

	for ( ; lPos <= m_LastValidSplitPos; ++lPos )
	{
		if ( m_SplitRows[lPos] && !m_SplitRows[lPos]->IsDelAdj( ) )
		{
			*plRow = GetSplitPosRow( lPos );
			return m_SplitRows[lPos];
		}
	}

	return NULL;
}


// FindPrevRow
// Sucht die vorherige Zeile
CMTblRow * CMTblRows::FindPrevRow( LPLONG plRow )
{
	if ( !plRow )
		return NULL;

	for ( long lPos = ( *plRow > m_LastValidPos ? m_LastValidPos : *plRow - 1 ); lPos >= 0; --lPos )
	{
		if ( m_Rows[lPos] && !m_Rows[lPos]->IsDelAdj( ) )
		{
			*plRow = lPos;
			return m_Rows[lPos];
		}
	}

	return NULL;
}


// FindPrevSplitRow
// Sucht die vorherige Split-Zeile
CMTblRow * CMTblRows::FindPrevSplitRow( LPLONG plRow )
{
	if ( !plRow )
		return NULL;
	
	long lPos = GetSplitRowPos( *plRow );
	if ( lPos < 0 )
		return NULL;
	if ( lPos > m_LastValidSplitPos )
		lPos = m_LastValidSplitPos;
	else
		--lPos;

	for ( ; lPos >= 0; --lPos )
	{
		if ( m_SplitRows[lPos] && !m_SplitRows[lPos]->IsDelAdj( ) )
		{
			*plRow = GetSplitPosRow( lPos );
			return m_SplitRows[lPos];
		}
	}

	return NULL;
}


// FreeRows
// Gibt Speicher für Zeilen frei
BOOL CMTblRows::FreeRows( )
{
	if ( m_Rows )
	{
		// Zeilen löschen
		for ( int i = 0; i <= m_LastValidPos; ++i )
		{
			if ( m_Rows[i] )
			{
				delete m_Rows[i];
				m_Rows[i] = NULL;
			}
		}

		// Array freigeben
		VirtualFree( m_Rows, ( m_UpperBound + 1 ) * sizeof( CMTblRow* ), MEM_DECOMMIT );
		VirtualFree( m_Rows, 0, MEM_RELEASE );

		// Variablen zurücksetzen
		m_Rows = NULL;
		m_UpperBound = -1;
		m_LastValidPos = -1;
	}


	return TRUE;
}

// FreeSplitRows
// Gibt Speicher für Split-Zeilen frei
BOOL CMTblRows::FreeSplitRows( )
{
	if ( m_SplitRows )
	{
		// Zeilen löschen
		for ( int i = 0; i <= m_LastValidSplitPos; ++i )
		{
			if ( m_SplitRows[i] )
			{
				delete m_SplitRows[i];
				m_SplitRows[i] = NULL;
			}
		}

		// Array freigeben
		VirtualFree( m_SplitRows, ( m_SplitUpperBound + 1 ) * sizeof( CMTblRow* ), MEM_DECOMMIT );
		VirtualFree( m_SplitRows, 0, MEM_RELEASE );

		// Variablen zurücksetzen
		m_SplitRows = NULL;
		m_SplitUpperBound = -1;
		m_LastValidSplitPos = -1;
	}


	return TRUE;
}

// GetChildRows
// Liefert die Kindzeilen einer Zeile
long CMTblRows::GetChildRows( CMTblRow *pRow, RowPtrVector *pRows, BOOL bBottomUp /*= FALSE*/ )
{
	// Zeile ok?
	if ( !pRow && pRows ) return 0;

	// Zeilennummer ermitteln
	long lNr = pRow->GetNr( );
	if ( lNr == TBL_Error ) return 0;

	// Suchen
	CMTblRow **Rows;
	long lLastValidPos;
	if ( lNr >= 0 )
	{
		Rows = m_Rows;
		lLastValidPos = m_LastValidPos;
	}
	else
	{
		Rows = m_SplitRows;
		lLastValidPos = m_LastValidSplitPos;
	}

	long lCount = 0;
	if ( bBottomUp )
	{
		for ( long i = lLastValidPos; i >= 0; --i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pRow )
				{
					pRows->push_back( Rows[i] );
					lCount++;
				}
			}
		}
	}
	else
	{
		for ( long i = 0; i <= lLastValidPos; ++i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pRow )
				{
					pRows->push_back( Rows[i] );
					lCount++;
				}
			}
		}
	}

	return lCount;
}


// GetColImagePtr
// Liefert Spalten-Bildzeiger einer Zeile ( = Zelle )
CMTblImg * CMTblRows::GetColImagePtr( long lRow, HWND hWndCol )
{
	CMTblRow * pRow = GetRowPtr( lRow );
	if ( !pRow ) return NULL;
	return pRow->GetColImagePtr( hWndCol );
}


// GetColPtr
// Liefert Spaltenzeiger einer Zeile ( = Zelle )
CMTblCol * CMTblRows::GetColPtr( long lRow, HWND hWndCol )
{
	CMTblRow * pRow = GetRowPtr( lRow );
	if ( !pRow ) return NULL;
	CMTblCol * pCol = pRow->m_Cols->Find( hWndCol );
	return pCol;	
}

// GetFirstChildRow
// Ermittelt die erste Kindzeile einer Zeile
CMTblRow * CMTblRows::GetFirstChildRow( CMTblRow * pRow, BOOL bBottomUp /*=FALSE*/ )
{
	// Zeile ok?
	if ( !pRow ) return NULL;

	// Zeilennummer ermitteln
	long lNr = pRow->GetNr( );
	if ( lNr == TBL_Error ) return NULL;

	// Suchen
	CMTblRow ** Rows;
	long lLastValidPos;
	if ( lNr >= 0 )
	{
		Rows = m_Rows;
		lLastValidPos = m_LastValidPos;
	}
	else
	{
		Rows = m_SplitRows;
		lLastValidPos = m_LastValidSplitPos;
	}

	if ( bBottomUp )
	{
		for ( long i = lLastValidPos; i >= 0; --i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pRow ) return Rows[i];
			}
		}
	}
	else
	{
		for ( long i = 0; i <= lLastValidPos; ++i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pRow ) return Rows[i];
			}
		}
	}

	return NULL;
}


// GetLastChildRow
// Ermittelt die letzte Kindzeile einer Zeile
CMTblRow * CMTblRows::GetLastChildRow( CMTblRow * pRow, BOOL bBottomUp /*=FALSE*/ )
{
	// Zeile ok?
	if ( !pRow ) return NULL;

	// Zeilennummer ermitteln
	long lNr = pRow->GetNr( );
	if ( lNr == TBL_Error ) return NULL;

	// Suchen
	CMTblRow **Rows;
	CMTblRow *pLastChildRow = NULL;
	long lLastValidPos;
	if ( lNr >= 0 )
	{
		Rows = m_Rows;
		lLastValidPos = m_LastValidPos;
	}
	else
	{
		Rows = m_SplitRows;
		lLastValidPos = m_LastValidSplitPos;
	}

	if ( bBottomUp )
	{
		for ( long i = lLastValidPos; i >= 0; --i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pRow )
					pLastChildRow = Rows[i];
			}
		}
	}
	else
	{
		for ( long i = 0; i <= lLastValidPos; ++i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pRow )
					pLastChildRow = Rows[i];
			}
		}
	}

	return pLastChildRow;
}


// GetLastDescendantRow
// Ermittelt den letzten Nachkommen einer Zeile
CMTblRow * CMTblRows::GetLastDescendantRow( CMTblRow * pRow, BOOL bBottomUp /*=FALSE*/ )
{
	// Zeile ok?
	if ( !pRow ) return NULL;

	// Zeilennummer ermitteln
	long lNr = pRow->GetNr( );
	if ( lNr == TBL_Error ) return NULL;

	// Suchen
	CMTblRow **Rows;
	CMTblRow *pLastDescRow = NULL;
	long lLastValidPos;
	if ( lNr >= 0 )
	{
		Rows = m_Rows;
		lLastValidPos = m_LastValidPos;
	}
	else
	{
		Rows = m_SplitRows;
		lLastValidPos = m_LastValidSplitPos;
	}

	if ( bBottomUp )
	{
		for ( long i = lLastValidPos; i >= 0; --i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->IsDescendantOf( pRow ) )
					pLastDescRow = Rows[i];
			}
		}
	}
	else
	{
		for ( long i = 0; i <= lLastValidPos; ++i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->IsDescendantOf( pRow ) )
					pLastDescRow = Rows[i];
			}
		}
	}

	return pLastDescRow;
}


// GetNextChildRow
// Ermittelt die nächste Kindzeile
CMTblRow * CMTblRows::GetNextChildRow( CMTblRow * pRow, BOOL bBottomUp /*=FALSE*/ )
{
	// Zeile ok?
	if ( !pRow ) return NULL;

	// Zeilennummer ermitteln
	long lNr = pRow->GetNr( );
	if ( lNr == TBL_Error ) return NULL;

	// Elternzeile ermitteln
	CMTblRow * pParentRow = pRow->GetParentRow( );
	if ( !pParentRow ) return NULL;

	// Suchen
	CMTblRow ** Rows;
	long lLastValidPos, lPos;
	if ( lNr >= 0 )
	{
		Rows = m_Rows;
		lLastValidPos = m_LastValidPos;
		lPos = lNr;
	}
	else
	{
		Rows = m_SplitRows;
		lLastValidPos = m_LastValidSplitPos;
		lPos = GetSplitRowPos( lNr );
	}

	if ( bBottomUp )
	{
		for ( long i = lPos - 1; i >= 0; --i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pParentRow ) return Rows[i];
			}
		}
	}
	else
	{
		for ( long i = lPos + 1; i <= lLastValidPos; ++i )
		{
			if ( Rows[i] )
			{
				if ( Rows[i]->GetParentRow( ) == pParentRow ) return Rows[i];
			}
		}
	}

	return NULL;
}


// GetOrigParentRow
// Liefert die ursprüngliche Elternzeile einer Zeile
long CMTblRows::GetOrigParentRow( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return TBL_Error;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return TBL_Error;

	// Nr. der ursprünglichen Elternzeile ermitteln
	CMTblRow * pParentRow = pRow->GetOrigParentRow( );
	return pParentRow ? pParentRow->GetNr( ) : TBL_Error;
}


// GetParentRow
// Liefert die Elternzeile einer Zeile
long CMTblRows::GetParentRow( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return TBL_Error;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return TBL_Error;

	// Nr. der Elternzeile ermitteln
	CMTblRow * pParentRow = pRow->GetParentRow( );
	return pParentRow ? pParentRow->GetNr( ) : TBL_Error;
}



// GetPrevChildRow
// Ermittelt die vorherige Kindzeile
CMTblRow * CMTblRows::GetPrevChildRow( CMTblRow * pRow, BOOL bBottomUp /*=FALSE*/ )
{
	return GetNextChildRow( pRow, !bBottomUp );
}


// GetRowArray
// Liefert den internen Zeilen-Array für eine Zeile
CMTblRow ** CMTblRows::GetRowArray( long lRow, LPLONG plUpperBound, LPLONG plLastValidPos )
{
	if ( lRow == TBL_Error ) return NULL;

	if ( lRow >= 0 )
	{
		*plUpperBound = m_UpperBound;
		*plLastValidPos = m_LastValidPos;
		return m_Rows;
	}
	else
	{
		*plUpperBound = m_SplitUpperBound;
		*plLastValidPos = m_LastValidSplitPos;
		return m_SplitRows;
	}
}


// GetRowHeight
// Liefert die eingestellte Höhe einer Zeile
long CMTblRows::GetRowHeight( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return 0;

	return pRow->GetHeight( );
}


// GetRowLevel
// Liefert die Ebene einer Zeile
int CMTblRows::GetRowLevel( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return 0;

	return pRow->GetLevel( );
}


// GetRowLogX
// Liefert die logische X-Position einer Zeile in Pixeln
long CMTblRows::GetRowLogX( long lRow, long lDefRowHt )
{
	BOOL bSplitRows = ( lRow < 0 );
	long lHt, lLastValidPos, lUpperBound, lx = 0;
	CMTblRow ** Rows = GetRowArray( (bSplitRows ? -1 : 0), &lUpperBound, &lLastValidPos );
	if ( !Rows )
		return TBL_Error;

	long lPos = 0;
	long lPosRow = bSplitRows ? GetSplitRowPos( lRow ) : lRow;
	while ( lPos < lPosRow )
	{
		if ( lPos <= lUpperBound && Rows[lPos] )
		{
			if ( Rows[lPos]->IsVisible( ) )
			{
				lHt = Rows[lPos]->GetHeight( );
				lx += (lHt == 0 ? lDefRowHt : lHt);
			}
		}
		else
			lx += lDefRowHt;

		++lPos;
	}

	return lx;
}


// GetRowPtr
// Liefert den Zeiger einer Zeile anhand der Nummer
CMTblRow * CMTblRows::GetRowPtr( long lRow )
{
	if ( lRow == TBL_Error )
	{
		return NULL;
	}
	else if ( lRow >= 0 )
	{
		if ( !IsValidPos( lRow ) ) return NULL;
		return m_Rows[lRow];
	}
	else
	{
		long lPos = GetSplitRowPos( lRow );
		if ( !IsValidSplitPos( lPos ) ) return NULL;
		return m_SplitRows[lPos];
	}
}


// GetRowUserData
// Liefert Benutzerdaten einer Zeile
tstring & CMTblRows::GetRowUserData( long lRow, tstring &sKey )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return gsEmpty;

	return pRow->GetUserData( sKey );
}


// GetRowVisPos
// Liefert die sichtbare Position einer Zeile
long CMTblRows::GetRowVisPos( long lRow )
{
	CMTblRow *pRow = GetRowPtr( lRow );
	if ( pRow && !pRow->IsVisible( ) )
		return TBL_Error;

	BOOL bSplitRow = ( lRow < 0 );
	long lLastValidPos, lUpperBound;
	CMTblRow ** Rows = GetRowArray( lRow, &lUpperBound, &lLastValidPos );
	/*if ( !Rows )
		return TBL_Error;*/

	long l, lPos = bSplitRow ? GetSplitRowPos( lRow ) : lRow, lVisPos = 0;

	if ( lPos > lLastValidPos )
	{
		lVisPos = ( lPos - lLastValidPos - 1 );
		lPos = ( lLastValidPos + 1 );
	}

	for ( l = lPos - 1; l >= 0; --l )
	{
		if ( Rows[l] )
		{
			if ( Rows[l]->IsVisible( ) )
				++lVisPos;
		}
		else
			++lVisPos;
	}

	return lVisPos;
}


// GetVisPosRow
// Liefert die Zeile anhand einer sichtbaren Position
long CMTblRows::GetVisPosRow( long lVisPos, BOOL bSplitRows /*=FALSE*/ )
{
	if ( lVisPos < 0 || lVisPos > TBL_MaxRow )
		return TBL_Error;

	long lLastValidPos, lUpperBound;
	CMTblRow ** Rows = GetRowArray( (bSplitRows ? -1 : 0), &lUpperBound, &lLastValidPos );
	/*if ( !Rows )
		return TBL_Error;*/

	long lPos = 0, lVisPosCur = -1;
	while ( TRUE )
	{
		if ( lPos <= lUpperBound && Rows[lPos] )
		{
			if ( Rows[lPos]->IsVisible( ) )
				++lVisPosCur;
		}
		else
			++lVisPosCur;

		if ( lVisPosCur == lVisPos )
			break;

		++lPos;
	}

	return (bSplitRows ? GetSplitPosRow( lPos ) : lPos);
}


// InitMembers
// Initialisierung
void CMTblRows::InitMembers( )
{
	// Alles mit 0 initialiseren
	memset( this, 0, sizeof(CMTblRows) );

	// Obergrenze Zeilen-Array setzen
	m_UpperBound = -1;

	// Höchste gültige Zeile setzen
	m_LastValidPos = -1;

	// Obergrenze Split-Zeilen-Array setzen
	m_SplitUpperBound = -1;

	// Höchste gültige Split-Zeile setzen
	m_LastValidSplitPos = -1;

}


// InsertRow
// Fügt eine Zeile ein
BOOL CMTblRows::InsertRow( long lRow )
{
	int i, iMaxMovePos = 0, j, jMax;
	long lPos;

	if ( lRow >= 0 )
	{
		lPos = lRow;
		// Größer als letzte gültige Zeile?
		if ( lPos > m_LastValidPos )
			return TRUE;
		else
			// Sicherstellen, daß hinter die letzte gültige Zeile verschoben werden kann
			if ( !EnsureValidPos( m_LastValidPos + 1 ) ) return FALSE;

		// Zeilen nach hinten verschieben
		jMax = m_LastValidPos + 1;
		for ( i = m_LastValidPos; i >= lPos; --i )
		{
			if ( !IsRowDelAdj( i ) )
			{
				// Position ermitteln, wohin verschoben werden muß
				for ( j = i + 1; j <= jMax; ++j )
				{
					if ( !IsRowDelAdj( j ) ) break;
				}

				if ( j <= jMax )
				{
					// Verschieben
					m_Rows[j] = m_Rows[i];
					m_Rows[i] = NULL;

					if ( m_Rows[j] )
					{
						m_Rows[j]->SetNr( j );

						if( j > iMaxMovePos )
							iMaxMovePos = j;
					}
				}
			}
		}

		m_Rows[lPos] = NULL;

		if ( iMaxMovePos > m_LastValidPos )
			m_LastValidPos = iMaxMovePos;
	}
	else
	{
		lPos = GetSplitRowPos( lRow );
		// Größer als letzte gültige Zeile?
		if ( lPos > m_LastValidSplitPos )
			return TRUE;
		else
			// Sicherstellen, daß hinter die letzte gültige Zeile verschoben werden kann
			if ( !EnsureValidSplitPos( m_LastValidSplitPos + 1 ) ) return FALSE;

		// Zeilen nach hinten verschieben
		jMax = m_LastValidSplitPos + 1;
		for ( i = m_LastValidSplitPos; i >= lPos; --i )
		{
			if ( !IsRowDelAdj( GetSplitPosRow( i ) ) )
			{
				// Position ermitteln, wohin verschoben werden muß
				for ( j = i + 1; j <= jMax; ++j )
				{
					if ( !IsRowDelAdj( GetSplitPosRow( j ) ) ) break;
				}

				if ( j <= jMax )
				{
					// Verschieben
					m_SplitRows[j] = m_SplitRows[i];
					m_SplitRows[i] = NULL;

					if ( m_SplitRows[j] )
					{
						m_SplitRows[j]->SetNr( GetSplitPosRow( j ) );

						if( j > iMaxMovePos )
							iMaxMovePos = j;
					}
				}
			}
		}

		m_SplitRows[lPos] = NULL;

		if ( iMaxMovePos > m_LastValidSplitPos )
			m_LastValidSplitPos = iMaxMovePos;
	}


	return TRUE;
}


// IsParentRow
// Ermittelt, ob eine Zeile Kindzeilen hat
BOOL CMTblRows::IsParentRow( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	return pRow->IsParent( );
}


// IsRowDescendantOf
BOOL CMTblRows::IsRowDescendantOf( long lRow, long lParentRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	// Elternzeile ermitteln
	CMTblRow * pParentRow;
	if ( !( pParentRow = GetRowPtr( lParentRow ) ) ) return FALSE;

	return pRow->IsDescendantOf( pParentRow );
}


// IsRowDelAdj
// Ermittelt, ob eine Zeile mit TBL_Adjust gelöscht wurde
BOOL CMTblRows::IsRowDelAdj( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	return pRow->IsDelAdj( );
}


// IsRowDisCol
// Ermittelt, ob eine Spalte einer Zeile gesperrt ist
BOOL CMTblRows::IsRowDisCol( long lRow, HWND hWndCol )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	return pRow->m_Cols->GetDisabled( hWndCol );
}


// IsRowHyperlinkCol
// Ermittelt, ob eine Spalte einer Zeile eine Hyperlink-Spalte ist
BOOL CMTblRows::IsRowHyperlinkCol( long lRow, HWND hWndCol )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	CMTblHyLnk * pHyLnk = pRow->m_Cols->GetHyLnk( hWndCol );
	if ( !pHyLnk ) return FALSE;

	return pHyLnk->IsEnabled( );
}


// IsRowDisabled
// Ermittelt, ob eine Zeile gesperrt ist
BOOL CMTblRows::IsRowDisabled( long lRow )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	return pRow->IsDisabled( );
}


// IsValidRowPtr
// Ermittelt, ob ein Zeilen-Zeiger gültig ist
BOOL CMTblRows::IsValidRowPtr( CMTblRow *pRow )
{
	if ( !pRow ) return FALSE;

	int i;

	// In "normalen" Zeilen suchen
	for ( i = 0; i <= m_LastValidPos; ++i )
	{
		if ( pRow == m_Rows[i] )
			return TRUE;
	}

	// In Split-Zeilen suchen
	for ( i = 0; i <= m_LastValidSplitPos; ++i )
	{
		if ( pRow == m_SplitRows[i] )
			return TRUE;
	}

	return FALSE;
}


// QueryRowFlags
// Ermittelt, ob eine Zeile best. Flags hat
BOOL CMTblRows::QueryRowFlags( long lRow, DWORD dwFlags )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = GetRowPtr( lRow ) ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return FALSE;

	return pRow->QueryFlags( dwFlags );
}


// Reset
// Zurücksetzen
BOOL CMTblRows::Reset( )
{
	// Zeilen freigeben
	FreeRows( );
	FreeSplitRows( );

	// Anzahl Kindzeilen zurücksetzen
	m_ChildRows = 0;

	return TRUE;
}


// ResetRowHeights
// Setzt die Höhen aller Zeilen oder Split-Zeilen auf 0 zurück
BOOL CMTblRows::ResetRowHeights( BOOL bSplitRows /*= FALSE*/ )
{
	long lLastValidPos, lUpperBound;
	CMTblRow **pRows = GetRowArray( bSplitRows ? -1 :0, &lUpperBound, &lLastValidPos );

	if ( pRows )
	{
		for ( long l = 0; l <= lLastValidPos; ++l )
		{
			if ( pRows[l] && !pRows[l]->IsDelAdj( ) )
				pRows[l]->SetHeight( 0 );
		}
	}

	return TRUE;
}


// SetCellDisabled
// Setzt Zelle disabled ja/nein
BOOL CMTblRows::SetCellDisabled( long lRow, HWND hWndCol, BOOL bSet )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( bSet )
	{
		if ( !( pRow = EnsureValid( lRow ) ) )
			return FALSE;
	}
	else
	{
		if ( !( pRow = GetRowPtr( lRow ) ) )
			return FALSE;
	}

	return pRow->m_Cols->SetDisabled( hWndCol, bSet );
}


// SetCellFlagsAnyRows
// Setzt Cell-Flags einer Spalte in allen Zeilen
BOOL CMTblRows::SetCellFlagsAnyRows( HWND hWndCol, DWORD dwFlags, BOOL bSet )
{
	if ( !hWndCol )
		return FALSE;

	CMTblCol *pCol;
	int i;

	for ( i = 0; i <= m_LastValidPos; ++i )
	{
		if ( m_Rows[i] )
		{
			pCol = bSet ? m_Rows[i]->m_Cols->EnsureValid( hWndCol ) : m_Rows[i]->m_Cols->Find( hWndCol );
			if ( pCol )
				pCol->SetFlags( dwFlags, bSet );
		}
	}

	for ( i = 0; i <= m_LastValidSplitPos; ++i )
	{
		if ( m_SplitRows[i] )
		{
			pCol = bSet ? m_SplitRows[i]->m_Cols->EnsureValid( hWndCol ) : m_SplitRows[i]->m_Cols->Find( hWndCol );
			if ( pCol )
				pCol->SetFlags( dwFlags, bSet );
		}
	}

	return FALSE;
}


// SetCellHyperlink
// Setzt Hyperlink einer Zelle
BOOL CMTblRows::SetCellHyperlink( long lRow, HWND hWndCol, CMTblHyLnk & hl )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return FALSE;

	return pRow->m_Cols->SetHyLnk( hWndCol, hl );
}


// SetDelAdj
// Setzt eine Zeile als mit TBL_Adjust gelöscht
BOOL CMTblRows::SetRowDelAdj( long lRow )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return FALSE;

	pRow->SetDelAdj( );

	return TRUE;
}


// SetMerged
// Setzt die Anzahl verbundener Zeilen
BOOL CMTblRows::SetMerged( long lRow, long lMerged )
{
	CMTblRow *pRow = EnsureValid( lRow );
	if ( !pRow ) return FALSE;
	pRow->SetMerged( lMerged );
	return TRUE;
}


// SetParentRow
// Setzt die Elternzeile einer Zeile
BOOL CMTblRows::SetParentRow( CMTblRow *pRow, CMTblRow *pParentRow, BOOL bAsOrig /*=FALSE*/ )
{
	// Check
	if ( !pRow )
		return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return FALSE;

	BOOL bOk = FALSE;
	long lMoreChildRows = 0;

	if ( pParentRow )
	{
		// Sicherstellen, dass es sich um Zeilen im selben Bereich handelt
		if ( ( pRow->GetNr( ) >= 0 && pParentRow->GetNr( ) < 0 ) || ( pRow->GetNr( ) < 0 && pParentRow->GetNr( ) >= 0 ) ) return FALSE;

		// Elternzeile mit TBL_Adjust gelöscht?
		if ( pParentRow->IsDelAdj( ) ) return FALSE;

		// Rekursive Zuweisung verhindern
		CMTblRow * pCheckRow = pParentRow;
		while ( pCheckRow )
		{
			pCheckRow = pCheckRow->GetParentRow( );
			if ( pCheckRow == pRow ) return FALSE;
		}

		// Zuweisen
		if ( !pRow->GetParentRow( ) )
			lMoreChildRows = 1;
		bOk = pRow->SetParentRow( pParentRow, bAsOrig );

	}
	else
	{
		// Zuweisen
		if ( pRow->GetParentRow( ) )
			lMoreChildRows = -1;
		bOk = pRow->SetParentRow( NULL );
	}

	if ( bOk )
	{
		// Anzahl Kindzeilen aktualisieren
		m_ChildRows += lMoreChildRows;
	}

	return bOk;
}

// SetParentRow
// Setzt die Elternzeile einer Zeile
BOOL CMTblRows::SetParentRow( long lRow, long lParentRow, BOOL bAsOrig /*=FALSE*/ )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) )
		return FALSE;

	// Sicherstellen, daß Elternzeile da ist
	CMTblRow * pParentRow = NULL;
	if ( lParentRow != TBL_Error )
	{
		if ( !( pParentRow = EnsureValid( lParentRow ) ) )
			return FALSE;
	}

	// Setzen
	return SetParentRow( pRow, pParentRow, bAsOrig );	
}


// SetRowDisabled
// Setzt/Entfernt den Gesperrt-Status einer Zeile
BOOL CMTblRows::SetRowDisabled( long lRow, BOOL bSet )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return FALSE;

	pRow->SetDisabled( bSet );

	return TRUE;
}


// SetRowFlags
// Setzt oder entfernt Flags einer Zeile
DWORD CMTblRows::SetRowFlags( long lRow, DWORD dwFlags, BOOL bSet )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return 0;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return 0;

	return pRow->SetFlags( dwFlags, bSet );
}


// SetRowHeight
// Setzt die Höhe einer Zeile, 0 = Standardhöhe
BOOL CMTblRows::SetRowHeight( long lRow, long lHeight )
{
	// Zeile ermitteln
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return FALSE;

	pRow->SetHeight( lHeight );
	return TRUE;
}


// SetRowHidden
// Setzt das Hidden-Flag einer Zeile
BOOL CMTblRows::SetRowHidden( long lRow, BOOL bSet )
{
	CMTblRow *pRow = EnsureValid( lRow );
	if ( !pRow )
		return FALSE;

	pRow->SetHiddenFlag( bSet );
	if ( bSet )
		pRow->m_Cols->SetColsFlags( MTBL_CELL_FLAG_SELECTED, FALSE );

	return TRUE;
}


// SetRowsInternalFlags
// Setzt interne Flags für alle Rows
BOOL CMTblRows::SetRowsInternalFlags( BOOL bSplitRows, DWORD dwFlags, BOOL bSet )
{
	CMTblRow ** Rows = ( bSplitRows ? m_SplitRows : m_Rows );
	long lLastValidPos = ( bSplitRows ? m_LastValidSplitPos : m_LastValidPos );

	if ( Rows )
	{
		for ( int i = 0; i <= lLastValidPos; ++i )
		{
			if ( Rows[i] )
			{
				Rows[i]->SetInternalFlags( dwFlags, bSet );
			}
		}
		return TRUE;
	}
	else
		return FALSE;
}


// SetRowUserData
// Setzt Benutzerdaten einer Zeile
BOOL CMTblRows::SetRowUserData( long lRow, tstring &sKey, tstring &sValue )
{
	// Sicherstellen, daß Zeile da ist
	CMTblRow * pRow;
	if ( !( pRow = EnsureValid( lRow ) ) ) return FALSE;

	// Zeile mit TBL_Adjust gelöscht?
	if ( pRow->IsDelAdj( ) ) return FALSE;

	return pRow->SetUserData( sKey, sValue );
}


// SwapRows
// Vertauscht zwei Zeilen
BOOL CMTblRows::SwapRows( long lRow1, long lRow2 )
{
	if ( lRow1 == lRow2 )
		return TRUE;

	if ( lRow1 < 0 || lRow2 < 0 )
		return FALSE;

	// Sicherstellen, dass Row1 kleiner als Row2 ist
	if ( lRow1 > lRow2 )
	{
		long lRow = lRow1;
		lRow1 = lRow2;
		lRow2 = lRow;
	}

	// Gültige Zeilenpositionen sicherstellen
	if ( !EnsureValidPos( lRow1 ) || !EnsureValidPos( lRow2 ) ) return FALSE;

	// Vertauschen
	CMTblRow * TempRow = m_Rows[lRow1];
	m_Rows[lRow1] = m_Rows[lRow2];
	m_Rows[lRow2] = TempRow;

	if ( m_Rows[lRow1] )
		m_Rows[lRow1]->SetNr( lRow1 );

	if ( m_Rows[lRow2] )
		m_Rows[lRow2]->SetNr( lRow2 );

	return TRUE;
}
