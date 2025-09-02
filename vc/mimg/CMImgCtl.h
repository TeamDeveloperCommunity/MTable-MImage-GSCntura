// CMImgCtl.h

#ifndef _INC_CMIMGCTL
#define _INC_CMIMGCTL

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <tchar.h>
#include "CMImgMan.h"
#include "swcc.h"

// Windows-Definitionen
#ifndef WM_MOUSEWHEEL
#define WM_MOUSEWHEEL 0x020A
#endif

#ifndef WHEEL_DELTA
#define WHEEL_DELTA 120
#endif

// Globale Objekte
extern CMImgMan ImgMan;

// Klasse CMImgCtl
class CMImgCtl
{
private:
	HWND m_hWnd;
	HANDLE m_hImage;
	DWORD m_dwAlignFlags;
	double m_dHScale;
	double m_dVScale;
	int m_iFitMode;
	COLORREF m_clrBrush;
	int m_iBrushStyle;
	HFONT m_hFont;
	POINT m_ptImage;
	POINT m_ptText;
	DWORD m_dwFlags;

	SCROLLINFO m_sih, m_siv;
	BOOL m_bDesignMode;
	BOOL m_bOwnImage;
	BOOL m_bIgnoreSizeMsg;

	// Konstruktoren
	CMImgCtl( );

	// Funktionen
	void GetEffTextRect( RECT &r );
	BOOL GetFullClientRect( RECT &r );
	void GetText( tstring &s );
	void GetTextExtent( SIZE &s );
	BOOL HasText( ){return SendMessage( m_hWnd, WM_GETTEXTLENGTH, 0, 0 ) > 0;}
	BOOL ImgDelete( );
	BOOL ImgGetEffPos( LPPOINT pPos );
	BOOL ImgGetEffRect( LPRECT pRect );
	BOOL ImgGetEffSize( LPSIZE pSize );
	BOOL ImgGetLogPos( LPPOINT pPos );
	void Init( );
	BOOL SetScrBars( );
	BOOL WndInvalidate( BOOL bErase = TRUE );

	// Message-Handler
	LRESULT OnDestroy( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnEraseBkgnd( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnGetFont( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnHScroll( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnLButtonDown( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnMouseWheel( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnNCCreate( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnPaint( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnSetFont( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnSetText( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnSize( HWND hWnd, WPARAM wParam, LPARAM lParam );
	LRESULT OnVScroll( HWND hWnd, WPARAM wParam, LPARAM lParam );
public:
	// Konstruktoren
	CMImgCtl( HWND hWnd );

	// Funktionen
	BOOL GetBackBrush( LPINT piStyle, LPCOLORREF pclr );
	BOOL GetContentsExtent( SIZE &s );
	int GetFitMode( ){return m_iFitMode;};
	HANDLE GetImage( ){return m_hImage;};
	BOOL GetImagePixFromPoint( POINT pt, LPPOINT pPt );
	BOOL GetImagePos( POINT &pt ){pt = m_ptImage;return TRUE;}
	double GetScale( int iWhat ){return ( iWhat == MIMG_SCALE_HORIZONTAL ) ? m_dHScale : ( iWhat == MIMG_SCALE_VERTICAL ? m_dVScale : 0.0 );}
	BOOL GetTextPos( POINT &pt ){pt = m_ptText;return TRUE;}
	BOOL QueryAlignFlags( DWORD dwFlags ){return dwFlags & m_dwAlignFlags;}
	BOOL QueryFlags( DWORD dwFlags ){return dwFlags & m_dwFlags;}
	BOOL Refresh( );
	BOOL SetAlignFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetBackBrush( int iStyle, COLORREF clr );
	BOOL SetFitMode( int iFitMode );
	BOOL SetFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetImage( HANDLE hImg, BOOL bOwn );
	BOOL SetImagePos( POINT &pt );
	BOOL SetScale( int iWhat, double dScale );
	BOOL SetTextPos( POINT &pt );
	BOOL SizeToContents( );
	LRESULT WndProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam );
};

#endif // _INC_CMIMGCTL