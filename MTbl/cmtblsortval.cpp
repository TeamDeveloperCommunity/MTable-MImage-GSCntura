// cmtblsortval.cpp

// Includes
#include "cmtblsortval.h"

// Konstruktor(en)
CMTblSortVal::CMTblSortVal( )
{
	InitMembers( );
}

// Destruktor
CMTblSortVal::~CMTblSortVal( )
{
	FreeData( );
}

// Funktionen
/*int CMTblSortVal::CmpDate( DATETIME &dt1, DATETIME &dt2 )
{
	int iTo = min( dt1.DATETIME_Length, dt2.DATETIME_Length );
	for ( int i = 0; i < iTo; ++i )
	{
		if ( dt1.DATETIME_Buffer[i] < dt2.DATETIME_Buffer[i] )
			return -1;
		else if ( dt1.DATETIME_Buffer[i] > dt2.DATETIME_Buffer[i] )
			return 1;
	}

	if ( dt1.DATETIME_Length < dt2.DATETIME_Length )
		return -1;
	else if ( dt1.DATETIME_Length > dt2.DATETIME_Length )
		return 1;
	else
		return 0;
}*/

int CMTblSortVal::CmpDate( DATETIME &dt1, DATETIME &dt2 )
{
	NUMBER n1, n2;

	memcpy( &n1, &dt1, sizeof(NUMBER) );
	n1.numLength &= DTLENGTHMASK;

	memcpy( &n2, &dt2, sizeof(NUMBER) );
	n2.numLength &= DTLENGTHMASK;

	return CmpNumber( n1, n2 );
}

/*int CMTblSortVal::CmpNumber( NUMBER &n1, NUMBER &n2 )
{
#ifdef _DEBUG
	double d1 = 0.0, d2 = 0.0;
	Ctd.CvtNumberToDouble( &n1, &d1 );
	Ctd.CvtNumberToDouble( &n2, &d2 );
#endif
	int iTo = min( n1.numLength, n2.numLength );
	for ( int i = 0; i < iTo; ++i )
	{
		if ( n1.numValue[i] < n2.numValue[i] )
			return -1;
		else if ( n1.numValue[i] > n2.numValue[i] )
			return 1;
	}

	if ( n1.numLength < n2.numLength )
		return -1;
	else if ( n1.numLength > n2.numLength )
		return 1;
	else
		return 0;
}*/

int CMTblSortVal::CmpNumber( NUMBER &n1, NUMBER &n2 )
{
	int iTo = min( n1.numLength, n2.numLength );
	// None is NUMBER_Null
	if ( iTo > 0 )
	{
		BOOL b1Positive = ( ( n1.numValue[0] & 128 ) != 0 );
		BOOL b2Positive = ( ( n2.numValue[0] & 128 ) != 0 );
		if ( b1Positive != b2Positive )
		{
			return b1Positive ? 1 : -1;
		}

		// None is 0
		if ( iTo > 1 )
		{
			int i1DecPos = int( ( ( n1.numValue[0] & 128 ) ^ n1.numValue[0] ) ) - 64;
			int i2DecPos = int( ( ( n2.numValue[0] & 128 ) ^ n2.numValue[0] ) ) - 64;
			if ( i1DecPos != i2DecPos )
			{
				return i1DecPos > i2DecPos ? 1 : -1;
			}

			for ( int i = 1; i < iTo; ++i )
			{			
				if ( n1.numValue[i] < n2.numValue[i] )
					return -1;
				else if ( n1.numValue[i] > n2.numValue[i] )
					return 1;
			}
		}		
	}

	if ( n1.numLength < n2.numLength )
		return -1;
	else if ( n1.numLength > n2.numLength )
		return 1;
	else
		return 0;
}

int CMTblSortVal::CmpString( LPCTSTR lpcts1, LPCTSTR lpcts2, BOOL bCaseInsensitive /*=FALSE*/, BOOL bStringSort /*=FALSE*/ )
{
	DWORD dwCmpFlags = 0;
	if ( bCaseInsensitive )
		dwCmpFlags |= NORM_IGNORECASE;
	if ( bStringSort )
		dwCmpFlags |= SORT_STRINGSORT;

	int iCmp = CompareString( LOCALE_USER_DEFAULT, dwCmpFlags, lpcts1, -1, lpcts2, -1 );

	if ( iCmp == CSTR_EQUAL )
		return 0;
	else if ( iCmp == CSTR_GREATER_THAN )
		return 1;
	else
		return -1;
}

int CMTblSortVal::Compare( CMTblSortVal &Val, BOOL bCaseInsensitive /*=FALSE*/, BOOL bStringSort /*=FALSE*/ )
{
	if ( m_iType == Val.m_iType && m_pData && Val.m_pData )
	{
		switch ( m_iType )
		{
		case DT_String:
			return CmpString( LPCTSTR(m_pData), LPCTSTR(Val.m_pData), bCaseInsensitive, bStringSort );
		case DT_Number:
			return CmpNumber( *LPNUMBER(m_pData), *LPNUMBER(Val.m_pData) );
		case DT_DateTime:
			return CmpDate( *LPDATETIME(m_pData), *LPDATETIME(Val.m_pData) );
		}
	}

	return 0;
}

void CMTblSortVal::FreeData( )
{
	if ( m_pData )
	{
		free( m_pData );
		m_pData = NULL;
	}
}

void CMTblSortVal::InitMembers( )
{
	m_iType = 0;
	m_pData = NULL;
}

BOOL CMTblSortVal::SetDate( DATETIME &dtVal )
{
	FreeData( );
	m_iType = DT_DateTime;

	if ( !( m_pData = malloc( sizeof(dtVal) ) ) )
		return FALSE;
	memcpy( m_pData, &dtVal, sizeof(dtVal) );

	return TRUE;
}

BOOL CMTblSortVal::SetNumber( NUMBER &nVal )
{
	FreeData( );
	m_iType = DT_Number;

	if ( !( m_pData = malloc( sizeof(nVal) ) ) )
		return FALSE;
	memcpy( m_pData, &nVal, sizeof(nVal) );

	return TRUE;
}

BOOL CMTblSortVal::SetString( LPCTSTR lpctsVal )
{
	FreeData( );
	m_iType = DT_String;

	if ( !lpctsVal )
	{
		if ( !( m_pData = malloc( sizeof(TCHAR) ) ) )
			return FALSE;
		*LPTSTR(m_pData) = 0;
	}
	else
	{
		long lLen = lstrlen( lpctsVal ) + sizeof(TCHAR);
		if ( !( m_pData = malloc( lLen * sizeof(TCHAR) ) ) )
			return FALSE;
		if ( !lstrcpy( LPTSTR(m_pData), lpctsVal ) )
			return FALSE;
	}

	return TRUE;
}