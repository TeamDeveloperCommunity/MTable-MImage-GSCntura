// cmtblcols.cpp

// Includes
#include "cmtblcols.h"

// Standard-Konstruktor
CMTblCols::CMTblCols( )
{
	InitMembers( );
}

// Weitere Konstruktoren
CMTblCols::CMTblCols( BYTE nColType, CMTblCounter *pCounter )
{
	InitMembers( );
	m_nColType = nColType;
	m_pCounter = pCounter;
}

// Destruktor
CMTblCols::~CMTblCols( )
{

}

// EnsureValid
CMTblCol * CMTblCols::EnsureValid( HWND hWndCol )
{
	if ( !hWndCol ) return NULL;

	CMTblCol c( hWndCol, m_nColType, m_pCounter );
	pair<MTblColList::iterator,bool> ret;
	ret = m_List.insert( MTblColList::value_type( hWndCol, c ) );

	return &((*ret.first).second);
}

// EnsureValidHdr
CMTblCol * CMTblCols::EnsureValidHdr( HWND hWndCol )
{
	if ( !hWndCol ) return NULL;

	CMTblCol c( hWndCol, COL_ITYPE_COLHDR, m_pCounter );
	pair<MTblColList::iterator,bool> ret;
	ret = m_HdrList.insert( MTblColList::value_type( hWndCol, c ) );

	return &((*ret.first).second);
}

// Find
CMTblCol * CMTblCols::Find( HWND hWndCol )
{
	MTblColList::iterator it, itEnd = m_List.end( );

	it = m_List.find( hWndCol );
	if ( it != itEnd )
		return &(*it).second;
	else
		return NULL;
}

// FindHdr
CMTblCol * CMTblCols::FindHdr( HWND hWndCol )
{
	MTblColList::iterator it, itEnd = m_HdrList.end( );

	it = m_HdrList.find( hWndCol );
	if ( it != itEnd )
		return &(*it).second;
	else
		return NULL;
}

// GetBtn
CMTblBtn * CMTblCols::GetBtn( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	return pCol ? &pCol->m_Btn : NULL;	
}

// GetBtnPushed
BOOL CMTblCols::GetBtnPushed( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	return pCol ? pCol->IsBtnPushed( ) : FALSE;
}

// GetColHdrImagePtr
CMTblImg * CMTblCols::GetColHdrImagePtr( HWND hWndCol )
{
	CMTblCol *pCol = FindHdr( hWndCol );
	if ( pCol && pCol->m_Image.GetHandle( ) )
		return &pCol->m_Image;

	return NULL;
}

// GetDisabled
BOOL CMTblCols::GetDisabled( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	//return pCol ? pCol->m_bDisabled : FALSE;
	return pCol ? pCol->IsDisabled( ) : FALSE;
}

// GetFont
CMTblFont * CMTblCols::GetFont( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	return pCol ? &pCol->m_Font : NULL;	
}

// GetHdrBtn
CMTblBtn * CMTblCols::GetHdrBtn( HWND hWndCol )
{
	CMTblCol *pCol = FindHdr( hWndCol );
	return pCol ? &pCol->m_Btn : NULL;	
}

// GetHdrBtnPushed
BOOL CMTblCols::GetHdrBtnPushed( HWND hWndCol )
{
	CMTblCol *pCol = FindHdr( hWndCol );
	return pCol ? pCol->IsBtnPushed( ) : FALSE;
}

// GetHdrLineDef
LPMTBLLINEDEF CMTblCols::GetHdrLineDef(HWND hWndCol, int iLineType)
{
	CMTblCol *pCol = FindHdr(hWndCol);
	return pCol ? pCol->GetLineDef(iLineType) : NULL;
}

// GetHyLnk
CMTblHyLnk * CMTblCols::GetHyLnk( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	return pCol ? &pCol->m_HyLnk : NULL;	
}

// GetLineDef
LPMTBLLINEDEF CMTblCols::GetLineDef(HWND hWndCol, int iLineType)
{
	CMTblCol *pCol = Find(hWndCol);
	return pCol ? pCol->GetLineDef( iLineType ) : NULL;
}

// GetMerged
long CMTblCols::GetMerged( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	return pCol ? pCol->m_Merged : 0;
}

// GetMergedRows
long CMTblCols::GetMergedRows( HWND hWndCol )
{
	CMTblCol *pCol = Find( hWndCol );
	return pCol ? pCol->m_MergedRows : 0;	
}

// InitMembers
void CMTblCols::InitMembers( )
{
	m_pCounter = NULL;
	m_nColType = COL_ITYPE_COL;
}

// SetBtn
BOOL CMTblCols::SetBtn( HWND hWndCol, CMTblBtn & b )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	pCol->m_Btn = b;
	return TRUE;
}

// SetBtnPushed
BOOL CMTblCols::SetBtnPushed( HWND hWndCol, BOOL bSet )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	return pCol->SetBtnPushed( bSet );
}

// SetColsFlags
BOOL CMTblCols::SetColsFlags( DWORD dwFlags, BOOL bSet )
{
	MTblColList::iterator it, itEnd = m_List.end( );

	for ( it = m_List.begin( ); it != itEnd; ++it )
	{
		(*it).second.SetFlags( dwFlags, bSet );
	}

	return TRUE;
}

// SetColsInternalFlags
BOOL CMTblCols::SetColsInternalFlags( DWORD dwFlags, BOOL bSet )
{
	MTblColList::iterator it, itEnd = m_List.end( );

	for ( it = m_List.begin( ); it != itEnd; ++it )
	{
		(*it).second.SetInternalFlags( dwFlags, bSet );
	}

	return TRUE;
}

// SetDisabled
BOOL CMTblCols::SetDisabled( HWND hWndCol, BOOL bSet )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	return pCol->SetDisabled( bSet );
}

// SetFont
BOOL CMTblCols::SetFont( HWND hWndCol, CMTblFont & f )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	pCol->m_Font = f;
	return TRUE;
}

// SetHdrBtn
BOOL CMTblCols::SetHdrBtn( HWND hWndCol, CMTblBtn & b )
{
	CMTblCol *pCol = EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;
	pCol->m_Btn = b;
	return TRUE;
}

// SetHdrBtnPushed
BOOL CMTblCols::SetHdrBtnPushed( HWND hWndCol, BOOL bSet )
{
	CMTblCol *pCol = EnsureValidHdr( hWndCol );
	if ( !pCol ) return FALSE;
	return pCol->SetBtnPushed( bSet );
}

// SetHdrLineDef
BOOL CMTblCols::SetHdrLineDef(HWND hWndCol, CMTblLineDef & ld, int iLineType)
{
	CMTblCol *pCol = EnsureValidHdr(hWndCol);
	if (!pCol) return FALSE;
	return pCol->SetLineDef(ld, iLineType);
}

// SetHyLnk
BOOL CMTblCols::SetHyLnk( HWND hWndCol, CMTblHyLnk & hl )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	pCol->m_HyLnk = hl;
	return TRUE;
}

// SetImage
BOOL CMTblCols::SetImage( HWND hWndCol, HANDLE hImage )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	return pCol->m_Image.SetHandle( hImage, m_pCounter );
}

// SetImage2
BOOL CMTblCols::SetImage2( HWND hWndCol, HANDLE hImage )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	return pCol->m_Image2.SetHandle( hImage, m_pCounter );
}

// SetLineDef
BOOL CMTblCols::SetLineDef(HWND hWndCol, CMTblLineDef & ld, int iLineType)
{
	CMTblCol *pCol = EnsureValid(hWndCol);
	if (!pCol) return FALSE;
	return pCol->SetLineDef(ld, iLineType);
}

// SetMerged
BOOL CMTblCols::SetMerged( HWND hWndCol, long lMerged )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	pCol->SetMerged( lMerged );
	return TRUE;
}

// SetMergedRows
BOOL CMTblCols::SetMergedRows( HWND hWndCol, long lMergedRows )
{
	CMTblCol *pCol = EnsureValid( hWndCol );
	if ( !pCol ) return FALSE;
	pCol->SetMergedRows( lMergedRows );
	return TRUE;
}