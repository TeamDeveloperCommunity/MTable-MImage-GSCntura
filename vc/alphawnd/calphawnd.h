// calphawnd.h

#ifndef _INC_CALPHAWND
#define _INC_CALPHAWND

// Includes
#include <windows.h>
#include <tchar.h>
#include "winversion.h"

// Konstanten
#ifndef WS_EX_LAYERED
#define WS_EX_LAYERED 0x80000
#endif

#ifndef LWA_COLORKEY
#define LWA_COLORKEY 1
#endif

#ifndef LWA_ALPHA
#define LWA_ALPHA 2
#endif

#ifndef AW_HOR_POSITIVE
#define AW_HOR_POSITIVE 0x00000001
#endif

#ifndef AW_HOR_NEGATIVE
#define AW_HOR_NEGATIVE 0x00000002
#endif

#ifndef AW_VER_POSITIVE
#define AW_VER_POSITIVE 0x00000004
#endif

#ifndef AW_VER_NEGATIVE
#define AW_VER_NEGATIVE 0x00000008
#endif

#ifndef AW_CENTER
#define AW_CENTER 0x00000010
#endif

#ifndef AW_HIDE
#define AW_HIDE 0x00010000
#endif

#ifndef AW_ACTIVATE
#define AW_ACTIVATE 0x00020000
#endif

#ifndef AW_SLIDE
#define AW_SLIDE 0x00040000
#endif

#ifndef AW_BLEND
#define AW_BLEND 0x00080000
#endif

// Typedefs
typedef BOOL (WINAPI *PSLWA)(HWND, DWORD, BYTE, DWORD);
typedef DWORD (WINAPI *PAW)(HWND, DWORD, DWORD);

// Klasse CAlphaWnd
class CAlphaWnd
{
public:
	// Konstruktor
	CAlphaWnd( );
	// Destruktor
	~CAlphaWnd( );
	// Funktionen
	BOOL Animate( HWND hWnd, DWORD dwTime, DWORD dwFlags );
	BOOL CanUse( ){ return m_bCanUse; };	
	BOOL SetAlpha( HWND hWnd, BYTE byAlpha ); 
	BOOL SetLayeredStyle( HWND hWnd, BOOL bSet );
private:
	// Variablen
	BOOL m_bCanUse;
	HINSTANCE m_hUserDLL;
	// Funktionsvariablen
	PAW AnimateWindow;
	PSLWA SetLayeredWindowAttributes;
};


// Funktionen

// theAlphaWnd
CAlphaWnd & theAlphaWnd( );

#endif // _INC_CALPHAWND