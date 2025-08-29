// CMImgCtl.cpp

// Includes
#include "CMImgCtl.h"

// Konstruktoren
CMImgCtl::CMImgCtl( )
{

}

CMImgCtl::CMImgCtl( HWND hWnd )
{
	Init( );

	SetWindowLongPtr( hWnd, 0, LONG_PTR( this ) );
	m_hWnd = hWnd;
}

// GetBackBrush
BOOL CMImgCtl::GetBackBrush( LPINT piStyle, LPCOLORREF pclr )
{
	if ( !piStyle || !pclr ) return FALSE;

	*piStyle = m_iBrushStyle;
	*pclr = m_clrBrush;

	return TRUE;
}

// GetContentsExtent
BOOL CMImgCtl::GetContentsExtent( SIZE &s )
{
	s.cx = s.cy = 0;

	ImgGetEffSize( &s );
	if ( !( QueryAlignFlags( MIMG_ALIGN_HCENTER ) || QueryAlignFlags( MIMG_ALIGN_RIGHT ) ) )
		s.cx += m_ptImage.x;
	if ( !( QueryAlignFlags( MIMG_ALIGN_VCENTER ) || QueryAlignFlags( MIMG_ALIGN_BOTTOM ) ) )
		s.cy += m_ptImage.y;

	if ( HasText( ) )
	{
		SIZE st;
		GetTextExtent( st );
		st.cx += m_ptText.x;
		st.cy += m_ptText.y;
		if ( st.cx > s.cx )
			s.cx = st.cx;
		if ( st.cy > s.cy )
			s.cy = st.cy;
	}

	return TRUE;
}

// GetEffTextRect
void CMImgCtl::GetEffTextRect( RECT &r )
{
	SIZE s;
	GetTextExtent( s );
	
	r.left = m_ptText.x - m_sih.nPos;
	r.top = m_ptText.y - m_siv.nPos;
	r.right = r.left + s.cx;
	r.bottom = r.top + s.cy;
}

// GetFullClientRect
BOOL CMImgCtl::GetFullClientRect( RECT &r )
{
	if ( !GetClientRect( m_hWnd, &r ) )
		return FALSE;

	if ( m_siv.nMax > 0 )
		r.right += GetSystemMetrics( SM_CXVSCROLL );
	if ( m_sih.nMax > 0 )
		r.bottom += GetSystemMetrics( SM_CXHSCROLL );

	return TRUE;
}

// GetImagePixFromPoint
BOOL CMImgCtl::GetImagePixFromPoint( POINT pt, LPPOINT pPt )
{
	if ( !pPt ) return FALSE;

	RECT r;
	if ( !ImgGetEffRect( &r ) ) return FALSE;

	if ( pt.x > r.right || pt.x < r.left ) return FALSE;
	if ( pt.y > r.bottom || pt.y < r.top ) return FALSE;

	pPt->x = pt.x - r.left;
	pPt->y = pt.y - r.top;

	double dDivX = 0.0, dDivY = 0.0;

	if ( m_iFitMode != MIMG_FIT_NONE )
	{
		SIZE s;
		if ( !ImgMan.GetImgSize( m_hImage, &s ) ) return FALSE;
		dDivX = double(r.right - r.left) / double(s.cx);
		dDivY = double(r.bottom - r.top) / double(s.cy);
	}
	else
	{
		if ( m_dHScale != 1.0 )
			dDivX = m_dHScale;
		if ( m_dVScale != 1.0 )
			dDivY = m_dVScale;
	}

	if ( dDivX )
		pPt->x = double(pPt->x) / dDivX;
	if ( dDivY )
		pPt->y = double(pPt->y) / dDivY;

	return TRUE;
}

// GetText
void CMImgCtl::GetText( tstring &s )
{
	s = _T("");
	long lLen = SendMessage( m_hWnd, WM_GETTEXTLENGTH, 0, 0 );
	if ( lLen > 0 )
	{
		LPTSTR lpsText = new TCHAR[lLen + 1];
		if ( lpsText && SendMessage( m_hWnd, WM_GETTEXT, WPARAM( lLen + sizeof(TCHAR) ), LPARAM( lpsText ) ) )
		{
			s = lpsText;
			delete []lpsText;
		}
	}
}

// GetTextSize
void CMImgCtl::GetTextExtent( SIZE &s )
{
	s.cx = s.cy = 0;

	tstring sText;
	GetText( sText );
	if ( sText.size( ) > 0 )
	{
		HDC hDC = GetDC( m_hWnd );
		if ( hDC )
		{
			HGDIOBJ hOldFont = SelectObject( hDC, m_hFont );

			RECT r = { 0, 0, 0, 0 };
			DrawText( hDC, sText.c_str( ), -1, &r, DT_NOPREFIX | DT_CALCRECT );
			s.cx = r.right - r.left;
			s.cy = r.bottom - r.top;

			SelectObject( hDC, hOldFont );

			ReleaseDC( m_hWnd, hDC );
		}
	}	
}

// ImgDelete
BOOL CMImgCtl::ImgDelete( )
{
	if ( !ImgMan.IsValidImgHandle( m_hImage ) ) return TRUE;
	return ImgMan.DeleteImg( m_hImage );
}

// ImgGetEffPos
BOOL CMImgCtl::ImgGetEffPos( LPPOINT pPos )
{
	if ( !ImgGetLogPos( pPos ) ) return FALSE;
	
	pPos->x -= m_sih.nPos;
	pPos->y -= m_siv.nPos;

	return TRUE;
}

// ImgGetEffRect
BOOL CMImgCtl::ImgGetEffRect( LPRECT pRect )
{
	POINT p;
	if ( !ImgGetEffPos( &p ) ) return FALSE;

	SIZE s;
	if ( !ImgGetEffSize( &s ) ) return FALSE;

	pRect->left = p.x;
	pRect->top = p.y;
	pRect->right = p.x + s.cx;
	pRect->bottom = p.y + s.cy;

	return TRUE;
}

// ImgGetEffSize
BOOL CMImgCtl::ImgGetEffSize( LPSIZE pSize )
{
	if ( m_hImage && ImgMan.GetImgSize( m_hImage, pSize ) )
	{
		int iFitMode = GetFitMode( );
		if ( iFitMode != MIMG_FIT_NONE )
		{
			RECT r;
			if ( !GetFullClientRect( r ) ) return FALSE;
			int iClWid = r.right - r.left;
			if ( !( QueryAlignFlags( MIMG_ALIGN_HCENTER ) || QueryAlignFlags( MIMG_ALIGN_RIGHT ) ) )
				iClWid -= m_ptImage.x;
			int iClHt = r.bottom - r.top;
			if ( !( QueryAlignFlags( MIMG_ALIGN_VCENTER ) || QueryAlignFlags( MIMG_ALIGN_BOTTOM ) ) )
				iClHt -= m_ptImage.y;

			if ( iFitMode == MIMG_FIT_NORMAL )
			{
				pSize->cx = iClWid;
				pSize->cy = iClHt;
			}
			else if ( iFitMode == MIMG_FIT_BEST )
			{
				double dRatio = double(pSize->cx) / double(pSize->cy);
				double dRatioCl = double(iClWid) / double(iClHt);
				if ( dRatio >= dRatioCl )
				{
					pSize->cx = iClWid;
					pSize->cy = double(pSize->cx) / dRatio;
				}
				else
				{
					pSize->cy = iClHt;
					pSize->cx = double(pSize->cy) * dRatio;
				}
			}
			else if ( iFitMode == MIMG_FIT_WIDTH )
			{
				double dRatio = double(pSize->cx) / double(pSize->cy);
				pSize->cx = iClWid;
				pSize->cy = double(pSize->cx) / dRatio;
			}
			else if ( iFitMode == MIMG_FIT_HEIGHT )
			{
				double dRatio = double(pSize->cx) / double(pSize->cy);
				pSize->cy = iClHt;
				pSize->cx = double(pSize->cy) * dRatio;
			}
		}
		else
		{
			if ( m_dHScale != 1.0 )
				pSize->cx = double(pSize->cx) * m_dHScale;
			if ( m_dVScale != 1.0 )
				pSize->cy = double(pSize->cy) * m_dVScale;
		}
	}
	else
	{
		pSize->cx = pSize->cy = 0;
	}

	return TRUE;
}

// ImgGetLogPos
BOOL CMImgCtl::ImgGetLogPos( LPPOINT pPos )
{
	if ( !pPos ) return FALSE;

	RECT r;
	if ( !GetClientRect( m_hWnd, &r ) ) return FALSE;

	SIZE s;
	if ( !ImgGetEffSize( &s ) ) return FALSE;

	if ( QueryAlignFlags( MIMG_ALIGN_HCENTER ) )
		pPos->x = max( r.left, ( ( r.right - r.left ) - s.cx ) / 2 );
	else if ( QueryAlignFlags( MIMG_ALIGN_RIGHT ) )
		pPos->x = max( r.left, r.right - s.cx );
	else
		pPos->x = r.left + m_ptImage.x;

	if ( QueryAlignFlags( MIMG_ALIGN_VCENTER ) )
		pPos->y = max( r.top, ( ( r.bottom - r.top ) - s.cy ) / 2 );
	else if ( QueryAlignFlags( MIMG_ALIGN_BOTTOM ) )
		pPos->y = max( r.top, r.bottom - s.cy );
	else
		pPos->y = r.top + m_ptImage.y;

	return TRUE;
}

// Init
void CMImgCtl::Init( )
{
	m_hWnd = NULL;
	m_hImage = NULL;
	m_dwAlignFlags = 0;
	m_dHScale = m_dVScale = 1.0;
	m_iFitMode = 0;
	m_clrBrush = 0;
	m_iBrushStyle = MIMG_BRUSH_NONE;
	m_hFont = NULL;
	m_ptImage.x = m_ptImage.y = 0;
	m_ptText.x = m_ptText.y = 0;
	m_dwFlags = 0;

	ZeroMemory( &m_sih, sizeof( SCROLLINFO ) );
	ZeroMemory( &m_siv, sizeof( SCROLLINFO ) );
	m_sih.cbSize = m_siv.cbSize = sizeof( SCROLLINFO );

	m_bDesignMode = FALSE;
	m_bOwnImage = FALSE;
	m_bIgnoreSizeMsg = FALSE;
}

// OnDestroy
LRESULT CMImgCtl::OnDestroy( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	if ( m_bOwnImage )
		ImgDelete( );
	delete this;
	return 0;
}

// OnEraseBkgnd
LRESULT CMImgCtl::OnEraseBkgnd( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	HDC hDC = HDC( wParam );

	RECT r;
	GetClientRect( hWnd, &r );
	HBRUSH hBrush = HBRUSH( SendMessage( GetParent( hWnd ), WM_CTLCOLORSTATIC, WPARAM( hDC ), LPARAM( hWnd ) ) );
	FillRect( hDC, &r, hBrush );

	if ( m_iBrushStyle != MIMG_BRUSH_NONE )
	{
		hBrush = CreateHatchBrush( m_iBrushStyle, m_clrBrush );
		if ( hBrush )
		{
			POINT ptOld;
			SetBrushOrgEx( hDC, r.left - m_sih.nPos, r.top - m_siv.nPos, &ptOld );

			FillRect( hDC, &r, hBrush );

			DeleteObject( hBrush );
		}
	}

	return 1;
}

// OnGetFont
LRESULT CMImgCtl::OnGetFont( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	return LRESULT( m_hFont );
}

// OnHScroll
LRESULT CMImgCtl::OnHScroll( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	int iLineSize = 10, iScroll = 0;
	SCROLLINFO si;

	switch ( LOWORD( wParam ) )
	{
	case SB_LINEDOWN:
		iScroll = max( 0, min( iLineSize, m_sih.nMax - m_sih.nPage - m_sih.nPos + 1 ) );
		break;
	case SB_LINEUP:
		iScroll = max( -iLineSize, -m_sih.nPos );
		break;
	case SB_PAGEDOWN:
		iScroll = max( 0, min( m_sih.nPage, m_sih.nMax - m_sih.nPage - m_sih.nPos + 1 ) );
		break;
	case SB_PAGEUP:
		iScroll = max( -m_sih.nPage, -m_sih.nPos );
		break;
	case SB_THUMBTRACK:
		ZeroMemory( &si, sizeof( SCROLLINFO ) );
		si.cbSize = sizeof( SCROLLINFO );
		si.fMask = SIF_TRACKPOS;
		GetScrollInfo( hWnd, SB_HORZ, &si );
		iScroll = si.nTrackPos - m_sih.nPos;
		break;
	}

	if ( iScroll )
	{
		ScrollWindowEx( hWnd, -iScroll, 0, 0, 0, 0, 0, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN );
		m_sih.nPos += iScroll;
		m_sih.fMask = SIF_POS;
		SetScrollInfo( hWnd, SB_HORZ, &m_sih, TRUE );
	}

	return 0;
}

// OnLButtonDown
LRESULT CMImgCtl::OnLButtonDown( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	if ( !m_bDesignMode && QueryFlags( MIMG_FLAG_SET_FOCUS_ON_CLICK ) )
	{
		SetFocus( hWnd );
		return 0;
	}
	else
		return DefWindowProc( hWnd, WM_LBUTTONDOWN, wParam, lParam );
}

// OnMouseWheel
LRESULT CMImgCtl::OnMouseWheel( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	int iDelta = (short)HIWORD(wParam);
	int iLineSize = 30, iScroll = 0;

	if ( !m_bDesignMode && m_siv.nMax && iDelta )
	{
		iScroll = ( abs( iDelta ) / WHEEL_DELTA ) * iLineSize;
		if ( iScroll )
		{
			if ( iDelta < 0 )
				iScroll = max( 0, min( iScroll, m_siv.nMax - m_siv.nPage - m_siv.nPos + 1 ) );
			else if ( iDelta > 0 )
				iScroll = max( -iScroll, -m_siv.nPos );

			if ( iScroll )
			{
				ScrollWindowEx( hWnd, 0, -iScroll, 0, 0, 0, 0, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN );
				m_siv.nPos += iScroll;
				m_siv.fMask = SIF_POS;
				SetScrollInfo( hWnd, SB_VERT, &m_siv, TRUE );
			}
		}
	}

	return 0;
}

// OnNCCreate
LRESULT CMImgCtl::OnNCCreate( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	m_hWnd = hWnd;

	LPCREATESTRUCT pcs = LPCREATESTRUCT( lParam );
	LPSWCCCREATESTRUCT pccs = LPSWCCCREATESTRUCT( pcs->lpCreateParams );
	if ( pccs && !pccs->fUserMode )
		m_bDesignMode = TRUE;

	if ( !m_bDesignMode )
	{
		ULONG_PTR lStyle = GetClassLongPtr( hWnd, GCL_STYLE );
		if ( lStyle & CS_HREDRAW || lStyle & CS_VREDRAW )
			SetClassLongPtr( hWnd, GCL_STYLE,  lStyle - CS_HREDRAW - CS_VREDRAW );

		lStyle = GetWindowLongPtr( hWnd, GWL_STYLE );
		if ( !( lStyle & WS_HSCROLL ) )
			SetFlags( MIMG_FLAG_NOHSCROLL, TRUE );
		if ( !( lStyle & WS_VSCROLL ) )
			SetFlags( MIMG_FLAG_NOVSCROLL, TRUE );
	}

	return 1;
}

// OnPaint
LRESULT CMImgCtl::OnPaint( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	PAINTSTRUCT ps;
	BeginPaint( hWnd, &ps );

	if ( ImgMan.IsValidImgHandle( m_hImage ) )
	{
		ImgMan.SetImgDrawClipRect( m_hImage, &ps.rcPaint );

		RECT r;
		ImgGetEffRect( &r );

		DWORD dwFlags = MIMG_DRAW_USECLIPRECT;
		ImgMan.DrawImg( m_hImage, ps.hdc, &r, dwFlags );
	}

	tstring sText;
	GetText( sText );
	if ( sText.size( ) > 0 )
	{
		HGDIOBJ hOldFont = SelectObject( ps.hdc, m_hFont );

		int iOldBkMode = SetBkMode( ps.hdc, TRANSPARENT );

		int iX = m_ptText.x - m_sih.nPos;
		int iY = m_ptText.y - m_siv.nPos;

		RECT r = {iX,iY,iX,iY};
		DrawText( ps.hdc, sText.c_str( ), -1, &r, DT_NOPREFIX | DT_NOCLIP );

		SetBkMode( ps.hdc, iOldBkMode );
		SelectObject( ps.hdc, hOldFont );
	}

	EndPaint( hWnd, &ps );

	return 0;
}

// OnSize
LRESULT CMImgCtl::OnSize( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	if ( m_bIgnoreSizeMsg )
		return DefWindowProc( hWnd, WM_SIZE, wParam, lParam );

	SetScrBars( );

	if ( ImgMan.IsValidImgHandle( m_hImage ) )
	{
		if ( QueryAlignFlags( MIMG_ALIGN_HCENTER | MIMG_ALIGN_RIGHT | MIMG_ALIGN_VCENTER | MIMG_ALIGN_BOTTOM ) ||
			 GetFitMode( ) != MIMG_FIT_NONE
		   )
			WndInvalidate( );
	}

	return 0;
}

// OnSetFont
LRESULT CMImgCtl::OnSetFont( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	HFONT hFont = HFONT( wParam );
	if ( hFont != m_hFont )
	{
		m_hFont = hFont;

		SetScrBars( );
		if ( lParam && HasText( ) )
			RedrawWindow( hWnd, NULL, NULL, RDW_ERASE | RDW_INVALIDATE );
	}

	return 0;
}

// OnSetText
LRESULT CMImgCtl::OnSetText( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	LRESULT lRet = DefWindowProc( hWnd, WM_SETTEXT, wParam, lParam );
	Refresh( );
	return lRet;
}

// OnVScroll
LRESULT CMImgCtl::OnVScroll( HWND hWnd, WPARAM wParam, LPARAM lParam )
{
	int iLineSize = 10, iScroll = 0;
	SCROLLINFO si;

	switch ( LOWORD( wParam ) )
	{
	case SB_LINEDOWN:
		iScroll = max( 0, min( iLineSize, m_siv.nMax - m_siv.nPage - m_siv.nPos + 1 ) );
		break;
	case SB_LINEUP:
		iScroll = max( -iLineSize, -m_siv.nPos );
		break;
	case SB_PAGEDOWN:
		iScroll = max( 0, min( m_siv.nPage, m_siv.nMax - m_siv.nPage - m_siv.nPos + 1 ) );
		break;
	case SB_PAGEUP:
		iScroll = max( -m_siv.nPage, -m_siv.nPos );
		break;
	case SB_THUMBTRACK:
		ZeroMemory( &si, sizeof( SCROLLINFO ) );
		si.cbSize = sizeof( SCROLLINFO );
		si.fMask = SIF_TRACKPOS;
		GetScrollInfo( hWnd, SB_VERT, &si );
		iScroll = si.nTrackPos - m_siv.nPos;
		break;
	}

	if ( iScroll )
	{
		ScrollWindowEx( hWnd, 0, -iScroll, 0, 0, 0, 0, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN );
		m_siv.nPos += iScroll;
		m_siv.fMask = SIF_POS;
		SetScrollInfo( hWnd, SB_VERT, &m_siv, TRUE );
	}

	return 0;
}

// Refresh
BOOL CMImgCtl::Refresh( )
{
	SetScrBars( );
	WndInvalidate( );
	return TRUE;
}

// SetAlignFlags
BOOL CMImgCtl::SetAlignFlags( DWORD dwFlags, BOOL bSet )
{
	DWORD dwOldFlags = m_dwAlignFlags;

	m_dwAlignFlags = ( bSet ? ( m_dwAlignFlags | dwFlags ) : ( ( m_dwAlignFlags & dwFlags ) ^ m_dwAlignFlags ) );
	if ( dwOldFlags != m_dwAlignFlags )
		Refresh( );

	return TRUE;
}

// SetBackBrush
BOOL CMImgCtl::SetBackBrush( int iStyle, COLORREF clr )
{
	BOOL bChanged = FALSE;

	if ( iStyle != m_iBrushStyle )
	{
		m_iBrushStyle = iStyle;
		bChanged = TRUE;
	}

	if ( clr != m_clrBrush )
	{
		m_clrBrush = clr;
		bChanged = TRUE;
	}

	if ( bChanged )
		Refresh( );

	return TRUE;
}

// SetFitMode
BOOL CMImgCtl::SetFitMode( int iFitMode )
{
	if ( iFitMode != m_iFitMode )
	{
		m_iFitMode = iFitMode;
		Refresh( );
	}

	return TRUE;
}

// SetFlags
BOOL CMImgCtl::SetFlags( DWORD dwFlags, BOOL bSet )
{
	DWORD dwOldFlags = m_dwFlags;

	m_dwFlags = ( bSet ? ( m_dwFlags | dwFlags ) : ( ( m_dwFlags & dwFlags ) ^ m_dwFlags ) );
	if ( dwOldFlags != m_dwFlags )
		Refresh( );

	return TRUE;
}

// SetImage
BOOL CMImgCtl::SetImage( HANDLE hImg, BOOL bOwn )
{
	if ( hImg && !ImgMan.IsValidImgHandle( hImg ) ) return FALSE;

	if ( hImg != m_hImage )
	{
		if ( m_bOwnImage )
			ImgDelete( );

		m_hImage = hImg;
		m_bOwnImage = bOwn;

		Refresh( );
	}

	return TRUE;
}

// SetImagePos
BOOL CMImgCtl::SetImagePos( POINT &pt )
{
	if ( pt.x != m_ptImage.x || pt.y != m_ptImage.y )
	{
		m_ptImage = pt;
		Refresh( );
	}

	return TRUE;
}

// SetScale
BOOL CMImgCtl::SetScale( int iWhat, double dScale )
{
	if ( dScale <= 0.0 ) return FALSE;

	BOOL bHorz = ( iWhat == MIMG_SCALE_HORIZONTAL ) || ( iWhat == MIMG_SCALE_BOTH );
	BOOL bVert = ( iWhat == MIMG_SCALE_VERTICAL ) || ( iWhat == MIMG_SCALE_BOTH );
	BOOL bRefresh = FALSE;

	if ( bHorz && dScale != m_dHScale )
	{
		m_dHScale = dScale;
		bRefresh = TRUE;
	}

	if ( bVert && dScale != m_dVScale )
	{
		m_dVScale = dScale;
		bRefresh = TRUE;
	}

	if ( bRefresh )
		Refresh( );

	return TRUE;
}

// SetScrBars
BOOL CMImgCtl::SetScrBars( )
{
	BOOL bFitHt, bFitWid, bHScroll, bNoHScroll, bNoScroll, bNoVScroll, bVScroll;
	int iFitMode, iHt, iHScrollHt, iScroll, iVScrollWid, iWid;
	SIZE si;
	RECT r;

	if ( !GetFullClientRect( r ) ) return FALSE;

	iVScrollWid = GetSystemMetrics( SM_CXVSCROLL );
	iHScrollHt = GetSystemMetrics( SM_CYHSCROLL );

	iWid = r.right - r.left;

	iHt = r.bottom - r.top;

	iFitMode = GetFitMode( );
	bFitHt = iFitMode == MIMG_FIT_HEIGHT || iFitMode == MIMG_FIT_NORMAL || iFitMode == MIMG_FIT_BEST;
	bFitWid = iFitMode == MIMG_FIT_WIDTH || iFitMode == MIMG_FIT_NORMAL || iFitMode == MIMG_FIT_BEST;
	bNoHScroll = QueryFlags( MIMG_FLAG_NOHSCROLL );
	bNoVScroll = QueryFlags( MIMG_FLAG_NOVSCROLL );
	bNoScroll = ( bNoHScroll && bNoVScroll );

	if ( m_hImage && ImgMan.IsValidImgHandle( m_hImage ) && !bNoScroll )
	{
		if ( !ImgGetEffSize( &si ) ) return FALSE;

		if ( bNoHScroll || bFitWid /*|| bFitHt*/ )
			si.cx = 0;
		else if ( !( QueryAlignFlags( MIMG_ALIGN_HCENTER ) || QueryAlignFlags( MIMG_ALIGN_RIGHT ) ) )
			si.cx += m_ptImage.x;

		if ( bNoVScroll || bFitHt /*|| bFitWid*/ )
			si.cy = 0;
		else if ( !( QueryAlignFlags( MIMG_ALIGN_VCENTER ) || QueryAlignFlags( MIMG_ALIGN_BOTTOM ) ) )
			si.cy += m_ptImage.y;
	}
	else
	{
		si.cx = 0;
		si.cy = 0;
	}

	if ( HasText( ) && !bNoScroll )
	{
		SIZE st;
		GetTextExtent( st );
		st.cx += m_ptText.x;
		st.cy += m_ptText.y;
		if ( st.cx > si.cx && !bNoHScroll )
			si.cx = st.cx;
		if ( st.cy > si.cy && !bNoVScroll )
			si.cy = st.cy;
	}

	bVScroll = ( si.cy > iHt );
	bHScroll = ( si.cx > iWid );

	if ( bVScroll )
	{
		iWid -= iVScrollWid;
		bHScroll = ( si.cx > iWid );
	}
	if ( bHScroll )
	{
		iHt -= iHScrollHt;
		if ( si.cy > iHt && !bVScroll )
		{
			bVScroll = TRUE;
			iWid -= iVScrollWid;
		}
	}

	// Vertikal
	if ( bVScroll )
		m_siv.nMax = si.cy;
	else
		m_siv.nMax = 0;

	if ( m_siv.nPos && ( si.cy - m_siv.nPos ) < iHt )
	{
		iScroll = min( m_siv.nPos, iHt - ( si.cy - m_siv.nPos ) );
		if ( iScroll > 0 )
		{
			ScrollWindowEx( m_hWnd, 0, iScroll, 0, 0, 0, 0, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN );
			m_siv.nPos -= iScroll;
		}
	}
	else
		m_siv.nPos = min( m_siv.nPos, m_siv.nMax );

	m_siv.nPage = iHt + 1;

	m_siv.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
	m_bIgnoreSizeMsg = TRUE;
	SetScrollInfo( m_hWnd, SB_VERT, &m_siv, TRUE );
	m_bIgnoreSizeMsg = FALSE;

	// Horizontal
	if ( bHScroll )
		m_sih.nMax = si.cx;
	else
		m_sih.nMax = 0;

	if ( m_sih.nPos && ( si.cx - m_sih.nPos ) < iWid )
	{
		iScroll = min( m_sih.nPos, iWid - ( si.cx - m_sih.nPos ) );
		if ( iScroll > 0 )
		{
			ScrollWindowEx( m_hWnd, iScroll, 0, 0, 0, 0, 0, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN );
			m_sih.nPos -= iScroll;
		}
	}
	else
		m_sih.nPos = min( m_sih.nPos, m_sih.nMax );

	m_sih.nPage = iWid + 1;

	m_sih.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
	m_bIgnoreSizeMsg = TRUE;
	SetScrollInfo( m_hWnd, SB_HORZ, &m_sih, TRUE );
	m_bIgnoreSizeMsg = FALSE;

	return TRUE;
}

// SetTextPos
BOOL CMImgCtl::SetTextPos( POINT &pt )
{
	if ( pt.x != m_ptText.x || pt.y != m_ptText.y )
	{
		m_ptText = pt;
		Refresh( );
	}

	return TRUE;
}

// SizeToImage
BOOL CMImgCtl::SizeToContents( )
{
	SIZE s;
	GetContentsExtent( s );

	RECT r = {0, 0, s.cx, s.cy};
	if ( !AdjustWindowRectEx( &r, GetWindowLongPtr( m_hWnd, GWL_STYLE ), FALSE, GetWindowLongPtr( m_hWnd, GWL_EXSTYLE ) ) ) return FALSE;

	return SetWindowPos( m_hWnd, NULL, 0, 0, r.right - r.left, r.bottom - r.top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE );
}

// WndInvalidate
BOOL CMImgCtl::WndInvalidate( BOOL bErase /*=TRUE*/ )
{
	return InvalidateRect( m_hWnd, NULL, bErase );
}

// WndProc
LRESULT CMImgCtl::WndProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	switch ( uMsg )
	{
	case WM_DESTROY:
		return OnDestroy( hWnd, wParam, lParam );
	case WM_ERASEBKGND:
		return OnEraseBkgnd( hWnd, wParam, lParam );
	case WM_GETFONT:
		return OnGetFont( hWnd, wParam, lParam );
	case WM_HSCROLL:
		return OnHScroll( hWnd, wParam, lParam );
	case WM_LBUTTONDOWN:
		return OnLButtonDown( hWnd, wParam, lParam );
	case WM_MOUSEWHEEL:
		return OnMouseWheel( hWnd, wParam, lParam );
	case WM_NCCREATE:
		return OnNCCreate( hWnd, wParam, lParam );
	case WM_PAINT:
		return OnPaint( hWnd, wParam, lParam );
	case WM_SETFONT:
		return OnSetFont( hWnd, wParam, lParam );
	case WM_SETTEXT:
		return OnSetText( hWnd, wParam, lParam );
	case WM_SIZE:
		return OnSize( hWnd, wParam, lParam );
	case WM_VSCROLL:
		return OnVScroll( hWnd, wParam, lParam );
	default:
		return DefWindowProc( hWnd, uMsg, wParam, lParam );
	}
}
