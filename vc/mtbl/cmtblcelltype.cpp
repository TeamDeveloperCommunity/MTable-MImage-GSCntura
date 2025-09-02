// cmtblcelltype.cpp

// Includes
#include "cmtblcelltype.h"

// Globale Variablen
extern CCtd Ctd;

// Konstruktoren
CMTblCellType::CMTblCellType( )
{
	InitMembers( );
}

// Destruktor
CMTblCellType::~CMTblCellType( )
{

}

// Operatoren
CMTblCellType & CMTblCellType::operator=( const CMTblCellType &ct )
{
	if ( this != &ct )
	{
		m_iType = ct.m_iType;
		m_dwFlags = ct.m_dwFlags;
		m_iLines = ct.m_iLines;
		m_sCheckedVal = ct.m_sCheckedVal;
		m_sUnCheckedVal = ct.m_sUnCheckedVal;
		m_ListEntries = ct.m_ListEntries;
	}

	return *this;
}


BOOL CMTblCellType::operator==( CMTblCellType &ct )
{
	if ( this == &ct )
		return TRUE;
	else
	{
		BOOL bEqual = m_iType == ct.m_iType &&
					  m_dwFlags == ct.m_dwFlags &&
					  m_iLines == ct.m_iLines &&
					  m_sCheckedVal == ct.m_sCheckedVal &&
					  m_sUnCheckedVal == ct.m_sUnCheckedVal &&
					  m_ListEntries == ct.m_ListEntries;

		return bEqual;
	}
}


BOOL CMTblCellType::operator!=( CMTblCellType &ct )
{
	return !(*this == ct );
}

// Funktionen
void CMTblCellType::DefineCheckBox( DWORD dwFlags, tstring &sCheckedVal, tstring &sUnCheckedVal )
{
	InitMembers( );
	m_iType = COL_CellType_CheckBox;
	m_dwFlags = dwFlags;
	m_sCheckedVal = sCheckedVal;
	m_sUnCheckedVal = sUnCheckedVal;
}


void CMTblCellType::DefineDropDownList( DWORD dwFlags, int iLines, CellTypeListEntries *pe )
{
	InitMembers( FALSE );
	m_iType = COL_CellType_DropDownList;
	m_dwFlags = dwFlags;
	m_iLines = iLines;
	if ( pe )
		m_ListEntries = *pe;
}


BOOL CMTblCellType::DefineFromColumn( HWND hWndCol )
{
	InitMembers( );

	int iCellType = Ctd.TblGetCellType( hWndCol );

	BOOL bOk = FALSE;
	DWORD dwFlags;
	HSTRING hs1 = 0, hs2 = 0;
	int iLines;

	switch ( iCellType )
	{
	case COL_CellType_Standard:
		DefineStandard( );
		bOk = TRUE;
		break;
	case COL_CellType_CheckBox:
		if ( Ctd.TblQueryCheckBoxColumn( hWndCol, &dwFlags, &hs1, &hs2 ) )
		{
			long lLen;
			tstring sCheckedVal( Ctd.StringGetBuffer( hs1, &lLen ) );
			Ctd.HStrUnlock(hs1);

			tstring sUnCheckedVal( Ctd.StringGetBuffer( hs2, &lLen ) );
			Ctd.HStrUnlock(hs2);

			DefineCheckBox( dwFlags, sCheckedVal, sUnCheckedVal );
			bOk = TRUE;
		}
		break;
	case COL_CellType_DropDownList:
		if ( Ctd.TblQueryDropDownListColumn( hWndCol, &dwFlags, &iLines ) )
		{
			CellTypeListEntries e;
			GetListData( hWndCol, &e );

			DefineDropDownList( dwFlags, iLines, &e );
			bOk = TRUE;
		}
		break;
	case COL_CellType_PopupEdit:
		if ( Ctd.TblQueryPopupEditColumn( hWndCol, &dwFlags, &iLines ) )
		{
			DefinePopupEdit( dwFlags, iLines );
			bOk = TRUE;
		}
		break;
	}

	if (hs1)
		Ctd.HStringUnRef(hs1);
		
	if ( hs2 )
		Ctd.HStringUnRef(hs2);

	return bOk;
}


void CMTblCellType::DefinePopupEdit( DWORD dwFlags, int iLines )
{
	InitMembers( );
	m_iType = COL_CellType_PopupEdit;
	m_iLines = iLines;
}


void CMTblCellType::DefineStandard( )
{
	InitMembers( );
	m_iType = COL_CellType_Standard;
}


BOOL CMTblCellType::Get( LPINT piType, LPDWORD pdwFlags, LPINT piLines, LPHSTRING phsCheckedVal, LPHSTRING phsUnCheckedVal )
{
	*piType = m_iType;
	*pdwFlags = m_dwFlags;
	*piLines = m_iLines;

	BOOL bOk = TRUE;

	long lLen = Ctd.BufLenFromStrLen( m_sCheckedVal.size( ) );
	if ( !Ctd.InitLPHSTRINGParam( phsCheckedVal, lLen ) )
		return FALSE;
	LPTSTR lps = Ctd.StringGetBuffer( *phsCheckedVal, &lLen );
	if (lps)
		lstrcpy(lps, m_sCheckedVal.c_str());
	else
		bOk = FALSE;
	Ctd.HStrUnlock(*phsCheckedVal);

	if (bOk)
	{
		lLen = Ctd.BufLenFromStrLen(m_sUnCheckedVal.size());
		if (!Ctd.InitLPHSTRINGParam(phsUnCheckedVal, lLen))
			return FALSE;
		lps = Ctd.StringGetBuffer(*phsUnCheckedVal, &lLen);
		if (lps)
			lstrcpy(lps, m_sUnCheckedVal.c_str());
		else
			bOk = FALSE;
		Ctd.HStrUnlock(*phsUnCheckedVal);
	}	

	return bOk;
}


int CMTblCellType::GetListData( HWND hWndCol, CellTypeListEntries *pe )
{
	HSTRING hs = 0;
	int iCount, iLen;
	LPTSTR lps;
	long lLen;

	if ( !pe )
		pe = &m_ListEntries;

	pe->clear( );

	iCount = Ctd.ListQueryCount( hWndCol );
	for ( int i = 0; i < iCount; i++ )
	{
		if ( ( iLen = Ctd.ListQueryText( hWndCol, i, &hs ) ) > 0 )
		{
			if ( lps = Ctd.StringGetBuffer( hs, &lLen ) )
			{
				CellTypeListEntries::value_type val;
				val.first = lps;
				val.second = SendMessage( hWndCol, LB_GETITEMDATA, (WPARAM)i, 0 );
				pe->push_back(val);
			}
			Ctd.HStrUnlock(hs);
		}
	}

	if ( hs )
		Ctd.HStringUnRef( hs );

	return iCount;
}


void CMTblCellType::InitMembers( BOOL bInitListData /*=TRUE*/ )
{
	m_iType = COL_CellType_Undefined;
	m_dwFlags = 0;
	m_iLines = 0;
	m_sCheckedVal = _T("");
	m_sUnCheckedVal = _T("");
	if ( bInitListData )
	{
		m_ListEntries.clear();
	}
	m_bSettingColumn = FALSE;
}


BOOL CMTblCellType::IsDropDownListWndCls( HWND hWnd )
{
	if ( !IsWindow( hWnd ) )
		return FALSE;

	TCHAR szClsName[MAX_PATH];

	if ( !GetClassName( hWnd, szClsName, MAX_PATH - 1 ) )
		return FALSE;

	return lstrcmp( szClsName, _T("ListBox") ) == 0;
}


BOOL CMTblCellType::IsMyTypeWndCls( HWND hWnd )
{
	if ( !IsWindow( hWnd ) )
		return FALSE;

	BOOL bIs = FALSE;
	switch ( m_iType )
	{
	case COL_CellType_DropDownList:
		bIs = IsDropDownListWndCls( hWnd );
		break;
	case COL_CellType_PopupEdit:
		bIs = IsPopupEditWndCls( hWnd );
		break;
	}

	return bIs;
}


BOOL CMTblCellType::IsPopupEditWndCls( HWND hWnd )
{
	if ( !IsWindow( hWnd ) )
		return FALSE;

	TCHAR szClsName[MAX_PATH];

	if ( !GetClassName( hWnd, szClsName, MAX_PATH - 1 ) )
		return FALSE;

	return lstrcmp( szClsName, _T("Edit") ) == 0;
}


BOOL CMTblCellType::SetColumn( HWND hWndCol )
{
	BOOL bOk = FALSE;
	BOOL bSet = FALSE, bSetListEntriesOnly = FALSE;

	m_bSettingColumn = TRUE;
	SendMessage( hWndCol, CIM_DefineCellType, 1, 0 );

	int iCurrentType = Ctd.TblGetCellType( hWndCol );
	switch ( m_iType )
	{
	case COL_CellType_Standard:
		if ( iCurrentType == COL_CellType_Standard )
			bOk = TRUE;
		else
			bOk = SendMessage( hWndCol, WM_USER + 122, 0, 0 );
		break;

	case COL_CellType_CheckBox:
		if ( iCurrentType == COL_CellType_CheckBox )
		{
			CMTblCellType ct;
			ct.DefineFromColumn( hWndCol );
			if ( ct != *this )
				bSet = TRUE;
		}
		else
			bSet = TRUE;

		if ( bSet )
			bOk = Ctd.TblDefineCheckBoxColumn( hWndCol, m_dwFlags, LPTSTR( m_sCheckedVal.c_str( ) ), LPTSTR( m_sUnCheckedVal.c_str( ) ) );
		else
			bOk = TRUE;
		break;

	case COL_CellType_DropDownList:
		if ( iCurrentType == COL_CellType_DropDownList )
		{
			CMTblCellType ct;
			ct.DefineFromColumn( hWndCol );
			if ( ct != *this )
			{
				bSet = TRUE;
				if ( m_dwFlags == ct.m_dwFlags && m_iLines == ct.m_iLines )
					bSetListEntriesOnly = TRUE;
			}
				
		}
		else
			bSet = TRUE;

		if ( bSet )
		{
			if ( !bSetListEntriesOnly )
				bOk = Ctd.TblDefineDropDownListColumn( hWndCol, m_dwFlags, m_iLines );
			else
				bOk = TRUE;

			if ( bOk )
			{
				int iIdx;
				Ctd.ListClear( hWndCol );
				CellTypeListEntries::iterator it, itEnd = m_ListEntries.end( );
				for ( it = m_ListEntries.begin( ); it != itEnd; ++it )
				{
					iIdx = Ctd.ListAdd( hWndCol, LPTSTR( (*it).first.c_str( ) ) );
					if ( iIdx >= 0 )
					{
						SendMessage( hWndCol, LB_SETITEMDATA, (WPARAM)iIdx, (*it).second );
					}
				}
			}
		}
		else
			bOk = TRUE;
		break;

	case COL_CellType_PopupEdit:
		if ( iCurrentType == COL_CellType_PopupEdit )
		{
			CMTblCellType ct;
			ct.DefineFromColumn( hWndCol );
			if ( ct != *this )
				bSet = TRUE;
		}
		else
			bSet = TRUE;

		if ( bSet )
			bOk = Ctd.TblDefinePopupEditColumn( hWndCol, m_dwFlags, m_iLines );
		else
			bOk = TRUE;
		break;
	}

	m_bSettingColumn = FALSE;
	SendMessage( hWndCol, CIM_DefineCellType, 2, 0 );

	return bOk;
}

void CMTblCellType::SortListEntries( )
{
	sort( m_ListEntries.begin( ), m_ListEntries.end( ) );
}