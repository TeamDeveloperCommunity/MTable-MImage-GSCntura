// cmtblrow.cpp

// Includes
#include "cmtblrow.h"

// Konstruktoren
CMTblRow::CMTblRow( )
{
	InitMembers( );
	CreateMembers( );
}

CMTblRow::CMTblRow( CMTblCounter *pCounter )
{
	InitMembers( );
	m_pCounter = pCounter;
	CreateMembers( );
}

// Destruktor
CMTblRow::~CMTblRow( )
{
	if ( m_pCounter )
	{
		if ( IsNoSelInvImageSet( ) )
			m_pCounter->DecNoSelInv( );

		if ( IsNoSelInvTextSet( ) )
			m_pCounter->DecNoSelInv( );

		if ( m_Merged > 0 )
			m_pCounter->DecMergeRows( );

		if ( GetHeight( ) != 0 )
			m_pCounter->DecRowHeights( );
	}

	if ( m_Cols )
	{
		delete m_Cols;
		m_Cols = NULL;
	}

	if ( m_TextColor )
	{
		delete m_TextColor;
		m_TextColor = NULL;
	}


	if ( m_BackColor )
	{
		delete m_BackColor;
		m_BackColor = NULL;
	}

	if ( m_HdrBackColor )
	{
		delete m_HdrBackColor;
		m_HdrBackColor = NULL;
	}

	if ( m_Font )
	{
		delete m_Font;
		m_Font = NULL;
	}

	if ( m_HdrTip )
	{
		delete m_HdrTip;
		m_HdrTip = NULL;
	}

	if ( m_UserData )
	{
		delete m_UserData;
		m_UserData = NULL;
	}

	if (m_HorzLineDef)
		delete m_HorzLineDef;

	if (m_VertLineDef)
		delete m_VertLineDef;

	if (m_HdrHorzLineDef)
		delete m_HdrHorzLineDef;

	if (m_HdrVertLineDef)
		delete m_HdrVertLineDef;
}

// CreateMembers
// Erzeugt Member
void CMTblRow::CreateMembers( )
{
	// Spaltenliste erzeugen
	m_Cols = new CMTblCols( COL_ITYPE_CELL, m_pCounter );

	// Textfarbe erzeugen
	m_TextColor = new CMTblColor;

	// Hintergrundfarbe erzeugen
	m_BackColor = new CMTblColor;

	// Header-Hintergrundfarbe erzeugen
	m_HdrBackColor = new CMTblColor;

	// Font erzeugen
	m_Font = new CMTblFont;
}


// DelUserData
// Löscht Benutzerdaten
BOOL CMTblRow::DelUserData( tstring &sKey )
{
	if ( !m_UserData ) return FALSE;

	// Gültiger Key?
	if ( sKey.length( ) < 1 ) return FALSE;

	return ( m_UserData->erase( sKey ) > 0 );
}

// EnumUserData
// Liefert alle Benutzerdaten
long CMTblRow::EnumUserData( HARRAY haKeys, HARRAY haValues )
{
	if ( !Ctd.IsPresent( ) ) return -1;
	if ( !m_UserData ) return 0;

	// Check Arrays
	long lDims;
	if ( !Ctd.ArrayDimCount( haKeys, &lDims ) ) return -1;
	if ( lDims > 1 ) return -1;

	if ( !Ctd.ArrayDimCount( haValues, &lDims ) ) return -1;
	if ( lDims > 1 ) return -1;

	// Untergrenzen der Arrays ermitteln
	long lLowKeys;
	if ( !Ctd.ArrayGetLowerBound( haKeys, 1, &lLowKeys ) ) return -1;

	long lLowValues;
	if ( !Ctd.ArrayGetLowerBound( haValues, 1, &lLowValues ) ) return -1;

	// Daten ermitteln
	BOOL bError = FALSE;
	HSTRING hs;
	LPTSTR lps;
	long lLen, lCount;
	STRMAPIT it;
	for ( it = m_UserData->begin( ), lCount = 0; it != m_UserData->end( ); ++it )
	{
		// Schlüssel
		hs = 0;
		Ctd.InitLPHSTRINGParam( &hs, Ctd.BufLenFromStrLen( it->first.length( ) ) );
		lps = Ctd.StringGetBuffer( hs, &lLen );
		if (lps)
			lstrcpy(lps, it->first.c_str());
		else
			bError = TRUE;

		Ctd.HStrUnlock(hs);
		if (bError)
		{
			Ctd.HStringUnRef(hs);
			return -1;
		}			

		if ( !Ctd.MDArrayPutHString( haKeys, hs, lCount + lLowKeys ) ) return -1;

		// Wert
		hs = 0;
		Ctd.InitLPHSTRINGParam( &hs, Ctd.BufLenFromStrLen( it->second.length( ) ) );
		lps = Ctd.StringGetBuffer( hs, &lLen );
		if (lps)
			lstrcpy(lps, it->second.c_str());
		else
			bError = TRUE;

		Ctd.HStrUnlock(hs);
		if (bError)
		{
			Ctd.HStringUnRef(hs);
			return -1;
		}

		if ( !Ctd.MDArrayPutHString( haValues, hs, lCount + lLowValues ) ) return -1;

		++lCount;
	}

	return lCount;
}


// GetColImagePtr
// Liefert Zeiger auf das anzuzeigende Bild
CMTblImg * CMTblRow::GetColImagePtr( HWND hWndCol )
{
	CMTblCol *pCol = m_Cols->Find( hWndCol );
	if ( pCol )
	{
		if ( QueryFlags( MTBL_ROW_ISEXPANDED ) )
		{
			if ( pCol->m_Image2.GetHandle( ) )
				return &pCol->m_Image2;
			else if ( pCol->m_Image.GetHandle( ) )
				return &pCol->m_Image;
		}
		else
		{
			if ( pCol->m_Image.GetHandle( ) )
				return &pCol->m_Image;
		}
	}

	return NULL;
}


// GetHdrLineDef
// Liefert eine Liniendefinition des Headers
LPMTBLLINEDEF CMTblRow::GetHdrLineDef(int iLineType)
{
	if (iLineType == MTLT_HORZ)
		return m_HdrHorzLineDef;
	else if (iLineType == MTLT_VERT)
		return m_HdrVertLineDef;
	else
		return NULL;
}

// GetLevel
// Liefert die Ebene der Zeile
int CMTblRow::GetLevel( )
{
	int iLevel = 0;
	CMTblRow * pRow = this;

	while ( pRow != NULL && ( pRow = pRow->GetParentRow( ) ) != NULL )
	{
		++iLevel;
	}

	return iLevel;
}


// GetLineDef
// Liefert eine Liniendefinition
LPMTBLLINEDEF CMTblRow::GetLineDef(int iLineType)
{
	if (iLineType == MTLT_HORZ)
		return m_HorzLineDef;
	else if (iLineType == MTLT_VERT)
		return m_VertLineDef;
	else
		return NULL;
}


// GetUserData
// Liefert Benutzerdaten
tstring & CMTblRow::GetUserData( tstring &sKey )
{
	if ( !m_UserData ) return gsEmpty;

	// Gültiger Key?
	if ( sKey.length( ) < 1 ) return gsEmpty;

	STRMAPIT it;
	it =  m_UserData->find( sKey );
	if ( it != m_UserData->end( ) )
		return it->second;
	else
		return gsEmpty;
}


// InitMembers
// Initialisierung
void CMTblRow::InitMembers( )
{
	// Alles mit 0 initialiseren
	memset( this, 0, sizeof(CMTblRow) );

	// Nr. setzen
	SetNr( TBL_Error );
}


// IsDecendantOf
// Ermittelt, ob die Zeile Nachkomme einer Zeile ist
BOOL CMTblRow::IsDescendantOf( CMTblRow * pRow )
{
	if ( pRow )
	{
		CMTblRow * pCurRow = this;
		while ( pCurRow != NULL && ( pCurRow = pCurRow->GetParentRow( ) ) != NULL )
		{
			if ( pCurRow == pRow ) return TRUE;
		}
	}
	
	return FALSE;
}


// IsOrigDecendantOf
// Ermittelt, ob die Zeile Original-Nachkomme einer Zeile ist
BOOL CMTblRow::IsOrigDescendantOf( CMTblRow * pRow )
{
	if ( pRow )
	{
		CMTblRow * pCurRow = this;
		while ( pCurRow != NULL && ( pCurRow = pCurRow->GetOrigParentRow( ) ) != NULL )
		{
			if ( pCurRow == pRow ) return TRUE;
		}
	}
	
	return FALSE;
}


// QueryInternalFlags
BOOL CMTblRow::QueryInternalFlags( DWORD dwFlags )
{
	return ( m_InternalFlags & dwFlags ) != 0;
}


// SetDelAdj
// Setzt die Zeile als mit TBL_Adjust gelöscht
void CMTblRow::SetDelAdj( )
{
	if ( !IsDelAdj( ) )
	{
		SetParentRow( NULL );
		if ( m_pCounter )
		{
			if ( IsNoSelInvImageSet(  ) )
				m_pCounter->DecNoSelInv( );

			if ( IsNoSelInvTextSet(  ) )
				m_pCounter->DecNoSelInv( );

			if ( m_Merged > 0 )
				m_pCounter->DecMergeRows( );

			if ( GetHeight( ) != 0 )
				m_pCounter->DecRowHeights( );
		}
	}

	m_Bits.DelAdj = TRUE;
}


// SetFlags
// Setzt oder entfernt Flags
DWORD CMTblRow::SetFlags( DWORD dwFlags, BOOL bSet )
{
	BOOL bNoSelInvImageBefore = IsNoSelInvImageSet( );
	BOOL bNoSelInvTextBefore = IsNoSelInvTextSet( );

	DWORD dwOldFlags = m_Flags;
	m_Flags = ( bSet ? ( m_Flags | dwFlags ) : ( ( m_Flags & dwFlags ) ^ m_Flags ) );

	DWORD dwDiff = ( dwOldFlags ^ m_Flags );

	if ( dwDiff && m_pCounter )
	{
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

	if ( dwDiff & MTBL_ROW_HIDDEN )
	{
		if ( m_ParentRow )
		{
			bSet ? m_ParentRow->m_ChildRowsVis-- : m_ParentRow->m_ChildRowsVis++;
			if ( m_ParentRow->m_ChildRowsVis > 0 )
				m_ParentRow->SetFlags( MTBL_ROW_CANEXPAND, TRUE );
			else
				m_ParentRow->SetFlags( MTBL_ROW_CANEXPAND | MTBL_ROW_ISEXPANDED, FALSE );
		}
	}

	return dwDiff;
}


// SetHdrLineDef
// Setzt eine Liniendefinition des Headers
BOOL CMTblRow::SetHdrLineDef(CMTblLineDef & ld, int iLineType)
{
	BOOL bOk = TRUE;

	if (iLineType == MTLT_HORZ)
	{
		if (!m_HdrHorzLineDef)
			m_HdrHorzLineDef = new CMTblLineDef;
		*m_HdrHorzLineDef = ld;
	}

	else if (iLineType == MTLT_VERT)
	{
		if (!m_HdrVertLineDef)
			m_HdrVertLineDef = new CMTblLineDef;
		*m_HdrVertLineDef = ld;
	}
	else
		bOk = FALSE;

	return bOk;
}


// SetHeight
// Setzt die Höhe
void CMTblRow::SetHeight( long lHeight )
{
	lHeight = max( lHeight, 0 );

	if ( lHeight != m_Height )
	{
		if ( m_pCounter )
		{
			if ( lHeight > 0 && m_Height <= 0 )
				m_pCounter->IncRowHeights( );
			else if ( lHeight <= 0 && m_Height > 0 )
				m_pCounter->DecRowHeights( );
		}
		m_Height = lHeight;
	}
}


// SetInternalFlags
// Setzt oder entfernt interne Flags
DWORD CMTblRow::SetInternalFlags( DWORD dwFlags, BOOL bSet )
{
	DWORD dwOldFlags = m_InternalFlags;
	m_InternalFlags = ( bSet ? ( m_InternalFlags | dwFlags ) : ( ( m_InternalFlags & dwFlags ) ^ m_InternalFlags ) );

	return ( dwOldFlags ^ m_InternalFlags );
}


// SetLineDef
// Setzt eine Liniendefinition
BOOL CMTblRow::SetLineDef(CMTblLineDef & ld, int iLineType)
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


// SetMerged
// Setzt die Anzahl verbundener Zeilen
void CMTblRow::SetMerged( long lMerged )
{
	lMerged = max( lMerged, 0 );

	if ( lMerged != m_Merged )
	{
		if ( m_pCounter )
		{
			if ( lMerged > 0 && m_Merged <= 0 )
				m_pCounter->IncMergeRows( );
			else if ( lMerged <= 0 && m_Merged > 0 )
				m_pCounter->DecMergeRows( );
		}
		m_Merged = lMerged;
	}
}

// SetOrigParentRow
// Setzt die Adresse der ursprünglichen Elternzeile
BOOL CMTblRow::SetOrigParentRow( CMTblRow * pRow )
{
	if ( pRow == this ) return FALSE;

	// Neue ursprüngliche Elternzeile setzen
	m_OrigParentRow = pRow;

	return TRUE;
}


// SetParentRow
// Setzt die Adresse der Elternzeile
BOOL CMTblRow::SetParentRow( CMTblRow * pRow, BOOL bAsOrig /*=FALSE*/ )
{
	if ( pRow == this ) return FALSE;

	// Als ursprüngliche Elternzeile setzen?
	if ( bAsOrig )			
		SetOrigParentRow( pRow );

	// Alte Elternzeile merken
	CMTblRow * pOldParentRow = m_ParentRow;

	// Neue Elternzeile setzen
	m_ParentRow = pRow;

	if ( pOldParentRow != m_ParentRow )
	{
		BOOL bHidden = this->QueryFlags( MTBL_ROW_HIDDEN );
		// Anzahl Kindzeilen und Flags der alten und neuen Elternzeile aktualisieren
		if ( pOldParentRow )
		{
			pOldParentRow->m_ChildRows--;
			if ( !bHidden )
				pOldParentRow->m_ChildRowsVis--;

			if ( pOldParentRow->m_ChildRowsVis < 1 )
				pOldParentRow->SetFlags( MTBL_ROW_CANEXPAND | MTBL_ROW_ISEXPANDED, FALSE );
		}
		if ( m_ParentRow )
		{
			m_ParentRow->m_ChildRows++;
			if ( !bHidden )
				m_ParentRow->m_ChildRowsVis++;

			if ( m_ParentRow->m_ChildRowsVis > 0 )
				m_ParentRow->SetFlags( MTBL_ROW_CANEXPAND, TRUE );
		}
	}

	return TRUE;
}


// SetUserData
// Setzt Benutzerdaten
BOOL CMTblRow::SetUserData( tstring &sKey, tstring &sValue )
{
	if ( !m_UserData )
		m_UserData = new StrMap;

	// Gültiger Key?
	if ( sKey.length( ) < 1 ) return FALSE;

	STRMAPIT it;
	it = m_UserData->find( sKey );
	if ( it == m_UserData->end( ) )
		m_UserData->insert( STRMAPVAL( sKey, sValue ) );
	else
		it->second = sValue;
	return TRUE;
}
