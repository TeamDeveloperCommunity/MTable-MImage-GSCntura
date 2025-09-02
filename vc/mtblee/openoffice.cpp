#include "stdafx.h"
#include "openoffice.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

// Funktionen
void AttachRange( OOSpreadsheet & Sheet, OOCellRange & R, int iColFrom, int iRowFrom, int iColTo /*= -1*/, int iRowTo /*= -1*/ )
{
	if ( iColTo < 0 )
		iColTo = iColFrom;

	if ( iRowTo < 0 )
		iRowTo = iRowFrom;

	LPDISPATCH pDisp = Sheet.GetCellRangeByPosition( iColFrom, iRowFrom, iColTo, iRowTo );
	R.AttachDispatch( pDisp );
}

long MakeColor( long lColor )
{
	long lR = GetRValue( lColor );
	lR <<= 16;

	long lG = GetGValue( lColor );
	lG <<= 8;

	long lB = GetBValue( lColor );

	return lR | lG | lB;
}

void SetFont( OOCellRange & R, FONTINFO & fi )
{
	
	R.SetFontName( fi.sName );
	R.SetFontSize( float( fi.iPointSize ) );
	R.SetFontWeight( fi.bBold ? FW_Bold : FW_Normal );
	R.SetFontStyle( fi.bItalic ? FS_Italic : FS_None );
	R.SetUnderline( fi.bUnderline ? UL_Single : UL_None );
	R.SetStrikeout( fi.bStrikeOut ? SO_Single : SO_None );
}

// Klassen

// OOBorderLine
BOOL OOBorderLine::SetColor( long lValue )
{
	DISPID did;
	LPOLESTR lposName = L"Color";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	long lColor = MakeColor( lValue );
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lColor );

	return TRUE;
}

BOOL OOBorderLine::SetInnerLineWidth( short shValue )
{
	DISPID did;
	LPOLESTR lposName = L"InnerLineWidth";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I2;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, shValue );

	return TRUE;
}

BOOL OOBorderLine::SetLineDistance( short shValue )
{
	DISPID did;
	LPOLESTR lposName = L"LineDistance";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I2;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, shValue );

	return TRUE;
}

BOOL OOBorderLine::SetOuterLineWidth( short shValue )
{
	DISPID did;
	LPOLESTR lposName = L"OuterLineWidth";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I2;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, shValue );

	return TRUE;
}


// OOCell
BOOL OOCell::SetNumberFormat( long lFmtIndex )
{
	DISPID did;
	LPOLESTR lposName = L"NumberFormat";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lFmtIndex );

	return TRUE;
}

BOOL OOCell::SetValue( LPCTSTR lpsValue, BOOL bFormula /*=FALSE*/ )
{
	DISPID did;
	LPOLESTR lposName = NULL;
	if ( bFormula )
		lposName = L"setFormula";
	else
		lposName = L"setString";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lpsValue );

	return TRUE;
}

BOOL OOCell::SetValue( double dValue )
{
	DISPID did;
	LPOLESTR lposName = L"setValue";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_R8;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, dValue );

	return TRUE;
}

// OOCellRange
BOOL OOCellRange::ClearContents( long lFlags )
{
	DISPID did;
	LPOLESTR lposName = L"clearContents";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lFlags );

	return TRUE;
}

LPDISPATCH OOCellRange::GetBorderLine( int iType )
{
	DISPID did;
	LPOLESTR lposName = NULL;
	OOTableBorder ooTableBorder;

	switch ( iType )
	{
	case BL_TopLeft_BottomRight:
		lposName = L"DiagonalTLBR";
		break;
	case BL_BottomLeft_TopRight:
		lposName = L"DiagonalBLTR";
		break;
	default:
		return NULL;
	}
	
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OOCellRange::GetCellByPosition( long lCol, long lRow )
{
	DISPID did;
	LPOLESTR lposName = L"getCellByPosition";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_I4 VTS_I4;

	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lCol, lRow );

	return lpd;
}

short OOCellRange::GetParaIndent( )
{
	DISPID did;
	LPOLESTR lposName = L"ParaIndent";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return -1;

	short sh = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_I2, &sh, NULL );

	return sh;
}

LPDISPATCH OOCellRange::GetRangeAddress( )
{
	DISPID did;
	LPOLESTR lposName = L"getRangeAddress";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OOCellRange::GetSpreadsheet( )
{
	DISPID did;
	LPOLESTR lposName = L"getSpreadsheet";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OOCellRange::GetStartCell( )
{
	LPDISPATCH pCell = GetCellByPosition( 0, 0 );
	return pCell;
}

LPDISPATCH OOCellRange::GetTableBorder( )
{
	DISPID did;
	LPOLESTR lposName = L"TableBorder";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_DISPATCH, &lpd, NULL );

	return lpd;
}


BOOL OOCellRange::Indent( long lSteps )
{
	if ( lSteps )
	{
		DISPID did;
		LPOLESTR lposName;

		if ( lSteps > 0 )
			lposName = L"incrementIndent";
		else
			lposName = L"decrementIndent";

		HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
		if ( FAILED( hr ) )
			return FALSE;

		for ( long l = 0; l < abs(lSteps); l++ )
		{
			InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, NULL );
		}
	}

	return TRUE;
}

BOOL OOCellRange::Merge( BOOL bMerge )
{
	DISPID did;
	LPOLESTR lposName = L"merge";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, bMerge );

	return TRUE;
}


BOOL OOCellRange::SetBackColor( long lValue )
{
	DISPID did;
	LPOLESTR lposName = L"CellBackColor";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	long lColor = MakeColor( lValue );
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lColor );

	return TRUE;
}

BOOL OOCellRange::SetBorderLine( int iType, LPDISPATCH lpDisp )
{
	DISPID did;
	LPOLESTR lposName;

	switch ( iType )
	{
	case BL_TopLeft_BottomRight:
		lposName = L"DiagonalTLBR";
		break;
	case BL_BottomLeft_TopRight:
		lposName = L"DiagonalBLTR";
		break;
	default:
		return FALSE;
	}

	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lpDisp );

	return TRUE;
}

BOOL OOCellRange::SetFontName( LPCTSTR lpsName )
{
	DISPID did;
	LPOLESTR lposName = L"CharFontName";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lpsName );

	return TRUE;
}

BOOL OOCellRange::SetFontSize( float fSize )
{
	DISPID did;
	LPOLESTR lposName = L"CharHeight";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_R4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, fSize );

	return TRUE;
}

BOOL OOCellRange::SetFontStyle( long lValue )
{
	DISPID did;
	LPOLESTR lposName = L"CharPosture";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lValue );

	return TRUE;
}

BOOL OOCellRange::SetFontWeight( float fWeight )
{
	DISPID did;
	LPOLESTR lposName = L"CharWeight";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_R4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, fWeight );

	return TRUE;
}

BOOL OOCellRange::SetHoriJustify( long lValue )
{
	DISPID did;
	LPOLESTR lposName = L"HoriJustify";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lValue );

	return TRUE;
}

BOOL OOCellRange::SetIndentLevel( long lValue )
{
	if ( !SetParaIndent( 0 ) )
		return FALSE;

	if ( lValue > 0 )
	{
		if ( !Indent( lValue ) )
			return FALSE;
	}

	return TRUE;
}

BOOL OOCellRange::SetNumberFormat( long lFmtIndex )
{
	DISPID did;
	LPOLESTR lposName = L"NumberFormat";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lFmtIndex );

	return TRUE;
}

BOOL OOCellRange::SetParaIndent( short shValue )
{
	DISPID did;
	LPOLESTR lposName = L"ParaIndent";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I2;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, shValue );

	return TRUE;
}

BOOL OOCellRange::SetStrikeout( short shValue )
{
	DISPID did;
	LPOLESTR lposName = L"CharStrikeout";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I2;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, shValue );

	return TRUE;
}

BOOL OOCellRange::SetTableBorder( LPDISPATCH lpDisp )
{
	DISPID did;
	LPOLESTR lposName = L"TableBorder";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lpDisp );

	return TRUE;
}

BOOL OOCellRange::SetTextColor( long lValue )
{
	DISPID did;
	LPOLESTR lposName = L"CharColor";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	long lColor = MakeColor( lValue );
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lColor );

	return TRUE;
}

BOOL OOCellRange::SetTextWrap( BOOL bValue )
{
	DISPID did;
	LPOLESTR lposName = L"IsTextWrapped";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, bValue );

	return TRUE;
}

BOOL OOCellRange::SetUnderline( short shValue )
{
	DISPID did;
	LPOLESTR lposName = L"CharUnderline";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I2;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, shValue );

	return TRUE;
}

BOOL OOCellRange::SetValue( LPCTSTR lpsValue, BOOL bFormula /*=FALSE*/ )
{
	LPDISPATCH pCell = GetStartCell( );
	if ( !pCell )
		return FALSE;

	OOCell ooCell( pCell );
	return ooCell.SetValue( lpsValue, bFormula );
}

BOOL OOCellRange::SetValue( double dValue )
{
	LPDISPATCH pCell = GetStartCell( );
	if ( !pCell )
		return FALSE;

	OOCell ooCell( pCell );
	return ooCell.SetValue( dValue );
}

BOOL OOCellRange::SetVertJustify( long lValue )
{
	DISPID did;
	LPOLESTR lposName = L"VertJustify";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lValue );

	return TRUE;
}


// OOCellRangeAddress
long OOCellRangeAddress::EndColumn( )
{
	DISPID did;
	LPOLESTR lposName = L"EndColumn";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return -1;

	long l = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_I4, &l, NULL );

	return l;
}

long OOCellRangeAddress::EndRow( )
{
	DISPID did;
	LPOLESTR lposName = L"EndRow";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return -1;

	long l = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_I4, &l, NULL );

	return l;
}

long OOCellRangeAddress::StartColumn( )
{
	DISPID did;
	LPOLESTR lposName = L"StartColumn";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return -1;

	long l = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_I4, &l, NULL );

	return l;
}

long OOCellRangeAddress::StartRow( )
{
	DISPID did;
	LPOLESTR lposName = L"StartRow";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return -1;

	long l = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_I4, &l, NULL );

	return l;
}


// OOComponents
LPDISPATCH OOComponents::CreateEnumeration( )
{
	DISPID did;
	LPOLESTR lposName = L"createEnumeration";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}


// OOConfirmer
OOConfirmer::OOConfirmer()
{
	InitializeCriticalSection( &m_csStop );
	SetStop( FALSE );
}

long OOConfirmer::Confirm( )
{
	long lConfirmed = 0;
	EnumWindows( ConfirmProc, (LPARAM)&lConfirmed );
	return lConfirmed;
}

BOOL CALLBACK OOConfirmer::ConfirmProc( HWND hWnd, LPARAM lParam )
{
	TCHAR szClsName[255] = _T("");
	GetClassName( hWnd, szClsName, sizeof( szClsName ) - 1 );
	if ( lstrcmpi( szClsName, _T("SALSUBFRAME") ) == 0 )
	{
		TCHAR szText[255] = _T("");
		long lLen = GetWindowText( hWnd, szText, sizeof( szText ) );
		if ( lLen )
		{
			SetForegroundWindow( hWnd );
			SendMessage( hWnd, WM_KEYDOWN, VK_RETURN, 0 );
			SendMessage( hWnd, WM_KEYUP, VK_RETURN, 0 );

			(*(long*)lParam)++;
		}
	}
	return TRUE;
}

BOOL OOConfirmer::IsStopped( )
{
	EnterCriticalSection( &m_csStop );
	BOOL bStop = m_bStop;
	LeaveCriticalSection( &m_csStop );
	return bStop;
}

void OOConfirmer::SetStop( BOOL bSet )
{
	EnterCriticalSection( &m_csStop );
	m_bStop = bSet;
	LeaveCriticalSection( &m_csStop );
}

UINT OOConfirmer::Start( LPVOID pParam )
{
	if ( !pParam )
		return 1;

	OOConfirmer *pPC = (OOConfirmer*)pParam;
	pPC->SetStop( FALSE );

	long lConfirmed;

	while ( TRUE )
	{
		Sleep( 100 );

		if ( pPC->IsStopped( ) )
			break;

		lConfirmed = Confirm( );
		if ( lConfirmed )
			break;		
	}

	pPC->SetStop( TRUE );
	return 0;
}

void OOConfirmer::Stop()
{
	EnterCriticalSection( &m_csStop );
	m_bStop = TRUE;
	LeaveCriticalSection( &m_csStop );
}


// OODesktop
LPDISPATCH OODesktop::GetActiveFrame( )
{
	DISPID did;
	LPOLESTR lposName = L"getActiveFrame";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OODesktop::GetComponents( )
{
	DISPID did;
	LPOLESTR lposName = L"getComponents";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

// OODesktop
LPDISPATCH OODesktop::GetCurrentFrame( )
{
	DISPID did;
	LPOLESTR lposName = L"getCurrentFrame";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OODesktop::LoadComponentFromUrl( LPCTSTR lpsUrl, LPCTSTR lpsFrameName, long lSearchFlags )
{
	DISPID did;
	LPOLESTR lposName = L"loadComponentFromURL";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	VARIANT v;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR VTS_I4 VTS_VARIANT;

	VariantInit( &v );
	v.vt = VT_ARRAY | VT_VARIANT;
	v.pparray = NULL;
	v.parray = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lpsUrl, lpsFrameName, lSearchFlags, &v );

	return lpd;
}


// OODispatchHelper
VARIANT OODispatchHelper::ExecuteDispatch( LPDISPATCH lpDispProvider, LPCTSTR lpsUrl, LPCTSTR lpsFrame, long lSearchFlags, COleSafeArray &PropArray )
{
	VARIANT v;
	VariantInit( &v );

	DISPID did;
	LPOLESTR lposName = L"executeDispatch";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( !FAILED( hr ) )
	{
		static BYTE parms[] =
			VTS_DISPATCH VTS_BSTR VTS_BSTR VTS_I4 VTS_VARIANT;	

		VARIANT *pv = (LPVARIANT)PropArray;
		InvokeHelper( did, DISPATCH_METHOD, VT_VARIANT, &v, parms, lpDispProvider, lpsUrl, lpsFrame, lSearchFlags, pv );
	}

	return v;
}


// OOEnumeration
BOOL OOEnumeration::HasMoreElements( )
{
	DISPID did;
	LPOLESTR lposName = L"hasMoreElements";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	BOOL b = FALSE;
	InvokeHelper( did, DISPATCH_METHOD, VT_BOOL, &b, NULL );

	return b;
}

VARIANT OOEnumeration::NextElement( )
{
	VARIANT v;
	VariantInit( &v );

	DISPID did;
	LPOLESTR lposName = L"nextElement";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( !FAILED( hr ) )	
		InvokeHelper( did, DISPATCH_METHOD, VT_VARIANT, &v, NULL );

	return v;
}


// OOFrame

CString OOFrame::GetName( )
{
	CString sName( _T("") );

	DISPID did;
	LPOLESTR lposName = L"getName";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( !FAILED( hr ) )
	{
		InvokeHelper( did, DISPATCH_METHOD, VT_BSTR, &sName, NULL );
	}

	return sName;
}


// OOIdlClass
LPDISPATCH OOIdlClass::CreateObject( )
{	
	DISPID did;
	LPOLESTR lposName = L"createObject";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	VARIANT v;
	VariantInit( &v );
	static BYTE parms[] = { VT_VARIANT | VT_MFCBYREF, 0 };
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &v );

	if ( v.vt == VT_DISPATCH )
		return v.pdispVal;
	else
	{
		VariantClear( &v );
		return NULL;
	}
}


// OOIndexAccess
VARIANT OOIndexAccess::GetByIndex( long lIndex )
{
	VARIANT v;
	VariantInit( &v );

	DISPID did;
	LPOLESTR lposName = L"getByIndex";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( !FAILED( hr ) )
	{
		static BYTE parms[] =
		VTS_I4;
		InvokeHelper( did, DISPATCH_METHOD, VT_VARIANT, &v, parms, lIndex );
	}

	return v;
}


// OOModel
LPDISPATCH OOModel::GetCurrentController( )
{
	DISPID did;
	LPOLESTR lposName = L"getCurrentController";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}


// OONumberFormatter
BOOL OONumberFormatter::AttachNumberFormatsSupplier( LPDISPATCH lpDisp )
{
	DISPID did;
	LPOLESTR lposName = L"attachNumberFormatsSupplier";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lpDisp );

	return TRUE;	
}


long OONumberFormatter::DetectNumberFormat( long lKey, LPCTSTR lpsValue )
{
	DISPID did;
	LPOLESTR lposName = L"detectNumberFormat";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4 VTS_BSTR;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lKey, lpsValue );

	return TRUE;
}


// OOPropertyValue
BOOL OOPropertyValue::SetName( LPCTSTR lpsName )
{
	DISPID did;
	LPOLESTR lposName = L"Name";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lpsName );

	return TRUE;
}

BOOL OOPropertyValue::SetValue( long Value )
{
	DISPID did;
	LPOLESTR lposName = L"Value";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_I4;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, Value );

	return TRUE;
}

BOOL OOPropertyValue::SetValue( LPCTSTR Value )
{
	DISPID did;
	LPOLESTR lposName = L"Value";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, Value );

	return TRUE;
}


// OOReflection
LPDISPATCH OOReflection::ForName( LPCTSTR lpsName )
{
	DISPID did;
	LPOLESTR lposName = L"forName";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lpsName );

	return lpd;
}


// OOServiceManager
LPDISPATCH OOServiceManager::CreateInstance( LPCTSTR lpsName )
{
	DISPID did;
	LPOLESTR lposName = L"createInstance";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lpsName );

	return lpd;
}


// OOSpreadsheet
LPDISPATCH OOSpreadsheet::GetCellRangeByName( LPCTSTR lpsName )
{
	DISPID did;
	LPOLESTR lposName = L"getCellRangeByName";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_BSTR;

	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lpsName );

	return lpd;
}

LPDISPATCH OOSpreadsheet::GetCellByPosition( long lCol, long lRow )
{
	DISPID did;
	LPOLESTR lposName = L"getCellByPosition";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_I4 VTS_I4;

	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lCol, lRow );

	return lpd;
}

LPDISPATCH OOSpreadsheet::GetCellRangeByPosition( long lColLeft, long lRowTop, long lColRight, long lRowBottom )
{
	DISPID did;
	LPOLESTR lposName = L"getCellRangeByPosition";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_I4 VTS_I4;

	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, parms, lColLeft, lRowTop, lColRight, lRowBottom );

	return lpd;
}

LPDISPATCH OOSpreadsheet::GetColumns( )
{
	DISPID did;
	LPOLESTR lposName = L"getColumns";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;

	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OOSpreadsheet::GetRows( )
{
	DISPID did;
	LPOLESTR lposName = L"getRows";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;

	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}


// OOSpreadsheetDocument
LPDISPATCH OOSpreadsheetDocument::GetSheets( )
{
	DISPID did;
	LPOLESTR lposName = L"getSheets";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

BOOL OOSpreadsheetDocument::IsSpreadsheetDocument( LPDISPATCH lpDisp )
{
	if ( !lpDisp )
		return FALSE;

	DISPID did;
	LPOLESTR lposName = L"getSheets";
	HRESULT hr = lpDisp->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	return !FAILED( hr );
}


// OOSpreadsheetView
LPDISPATCH OOSpreadsheetView::GetActiveSheet( )
{
	DISPID did;
	LPOLESTR lposName = L"getActiveSheet";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OOSpreadsheetView::GetFrame( )
{
	DISPID did;
	LPOLESTR lposName = L"getFrame";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

LPDISPATCH OOSpreadsheetView::GetSelection( )
{
	DISPID did;
	LPOLESTR lposName = L"getSelection";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	VARIANT v;
	VariantInit( &v );

	InvokeHelper( did, DISPATCH_METHOD, VT_VARIANT, &v, NULL );

	lpd = v.pdispVal;
	if ( lpd )
		lpd->AddRef( );
	VariantClear( &v );

	return lpd;
}

BOOL OOSpreadsheetView::InsertTransferable( LPDISPATCH lpDisp )
{
	DISPID did;
	LPOLESTR lposName = L"insertTransferable";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lpDisp );

	return TRUE;
}

BOOL OOSpreadsheetView::SetActiveSheet( LPDISPATCH pDisp )
{
	DISPID did;
	LPOLESTR lposName = L"setActiveSheet";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, pDisp );

	return TRUE;
}

BOOL OOSpreadsheetView::Select( LPDISPATCH pDisp )
{
	DISPID did;
	LPOLESTR lposName = L"select";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	BOOL bRet = FALSE;
	static BYTE parms[] =
		VTS_VARIANT;
	VARIANT v;

	VariantInit( &v );
	v.vt = VT_DISPATCH;
	v.pdispVal = pDisp;
	InvokeHelper( did, DISPATCH_METHOD, VT_BOOL, &bRet, parms, &v );

	return bRet;
}


// OOSpreadsheets
long OOSpreadsheets::Add( int iLanguage )
{
	long lSheets = GetCount( );
	if ( lSheets < 0 )
		return -1;

	//CString sName = iLanguage == MTE_LNG_GERMAN ? _T("Tabelle") : _T("Sheet");
	CString sName = _T("");
	GetUIStr( iLanguage, _T("OOCalc.NewSheet.Name"), sName );
	long lNr = lSheets + 1;
	TCHAR szNr[10];
	_itot( lNr, szNr, 10 );
	sName += szNr;

	if( !InsertNewByName( sName, (short)lNr ) )
		return -1;

	return lNr;
}

LPDISPATCH OOSpreadsheets::GetByIndex( long lIndex )
{
	DISPID did;
	LPOLESTR lposName = L"getByIndex";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	static BYTE parms[] =
		VTS_I4;

	VARIANT v;
	VariantInit( &v );

	InvokeHelper( did, DISPATCH_METHOD, VT_VARIANT, &v, parms, lIndex );

	lpd = v.pdispVal;
	if ( lpd )
		lpd->AddRef( );
	VariantClear( &v );

	return lpd;
}

long OOSpreadsheets::GetCount( )
{
	DISPID did;
	LPOLESTR lposName = L"getCount";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return -1;

	long lCount = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_I4, &lCount, NULL );

	return lCount;
}

VARIANT OOSpreadsheets::GetElementNames( )
{
	VARIANT v;
	VariantInit( &v );

	DISPID did;
	LPOLESTR lposName = L"getElementNames";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( !FAILED( hr ) )	
		InvokeHelper( did, DISPATCH_METHOD, VT_VARIANT, &v, NULL );

	return v;
}

BOOL OOSpreadsheets::InsertNewByName( LPCTSTR lpsName, short nPos )
{
	DISPID did;
	LPOLESTR lposName = L"insertNewByName";
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BSTR VTS_I2;
	InvokeHelper( did, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lpsName, nPos );

	return TRUE;
}


// OOSystemClipboard
LPDISPATCH OOSystemClipboard::GetContents( )
{
	DISPID did;
	LPOLESTR lposName = L"getContents";	
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_METHOD, VT_DISPATCH, &lpd, NULL );

	return lpd;
}


// OOTableBorder
LPDISPATCH OOTableBorder::GetBorderLine( int iType )
{
	DISPID did;
	LPOLESTR lposName = NULL;

	switch ( iType )
	{
	case BL_Top:
		lposName = L"TopLine";
		break;
	case BL_Left:
		lposName = L"LeftLine";
		break;
	case BL_Right:
		lposName = L"RightLine";
		break;
	case BL_Bottom:
		lposName = L"BottomLine";
		break;
	case BL_Horizontal:
		lposName = L"HorizontalLine";
		break;
	case BL_Vertical:
		lposName = L"VerticalLine";
		break;
	default:
		return NULL;
	}
	
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return NULL;

	LPDISPATCH lpd = NULL;
	InvokeHelper( did, DISPATCH_PROPERTYGET, VT_DISPATCH, &lpd, NULL );

	return lpd;
}

BOOL OOTableBorder::SetBorderLine( int iType, LPDISPATCH lpDisp )
{
	DISPID did;
	LPOLESTR lposName;

	switch ( iType )
	{
	case BL_Top:
		lposName = L"TopLine";
		break;
	case BL_Left:
		lposName = L"LeftLine";
		break;
	case BL_Right:
		lposName = L"RightLine";
		break;
	case BL_Bottom:
		lposName = L"BottomLine";
		break;
	case BL_Horizontal:
		lposName = L"HorizontalLine";
		break;
	case BL_Vertical:
		lposName = L"VerticalLine";
		break;
	default:
		return FALSE;
	}

	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, lpDisp );

	return TRUE;
}

BOOL OOTableBorder::SetLineValid( int iType, BOOL bValid )
{
	DISPID did;
	LPOLESTR lposName = NULL;

	switch ( iType )
	{
	case BL_Top:
		lposName = L"IsTopLineValid";
		break;
	case BL_Left:
		lposName = L"IsLeftLineValid";
		break;
	case BL_Right:
		lposName = L"IsRightLineValid";
		break;
	case BL_Bottom:
		lposName = L"IsBottomLineValid";
		break;
	case BL_Horizontal:
		lposName = L"IsHorizontalLineValid";
		break;
	case BL_Vertical:
		lposName = L"IsVerticalLineValid";
		break;
	default:
		return NULL;
	}
	
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, bValid );

	return TRUE;
}


// OOTableColumn
BOOL OOTableColumn::SetOptimalWidth( BOOL bSet )
{
	DISPID did;
	LPOLESTR lposName = L"OptimalWidth";
	
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, bSet );

	return TRUE;
}


// OOTableRow
BOOL OOTableRow::SetOptimalHeight( BOOL bSet )
{
	DISPID did;
	LPOLESTR lposName = L"OptimalHeight";
	
	HRESULT hr = m_lpDispatch->GetIDsOfNames( IID_NULL, &lposName, 1, LOCALE_USER_DEFAULT, &did );
	if ( FAILED( hr ) )
		return FALSE;

	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper( did, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, bSet );

	return TRUE;
}