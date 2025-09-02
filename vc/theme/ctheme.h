// ctheme.h

#ifndef _INC_CTHEME
#define _INC_CTHEME

// Includes
#include <windows.h>
#include <tchar.h>
#include "winversion.h"

// Konstanten
#define BP_CHECKBOX 3
#define BP_PUSHBUTTON 1
#define CBS_CHECKEDNORMAL 5
#define CBS_CHECKEDPRESSED 7
#define CBS_MIXEDNORMAL 9
#define CBS_MIXEDPRESSED 11
#define CBS_UNCHECKEDNORMAL 1
#define CBS_UNCHECKEDPRESSED 3
#define CBXS_NORMAL 1
#define CBXS_HOT 2;
#define CBXS_PRESSED 3
#define CP_DROPDOWNBUTTON 1
#define PBS_NORMAL 1
#define PBS_HOT 2
#define PBS_PRESSED 3
#define PBS_DISABLED 4

#define HP_HEADERITEM 1
#define HIS_NORMAL 1

#define WP_DIALOG 29
#define FS_ACTIVE 1

#define TMT_TEXTCOLOR 0x0EDB

// Typedefs
typedef HANDLE HTHEME;
typedef enum {
    TS_MIN,
    TS_TRUE,
    TS_DRAW
} THEMESIZE;

typedef HRESULT (WINAPI *CLOSE_THEME_DATA)(HTHEME hTheme);
typedef HRESULT (WINAPI *DRAW_THEME_BACKGROUND)(HTHEME hTheme, HDC hDC, int iPartID, int iStateID, const RECT *pRect,const RECT *pClipRect);
typedef COLORREF (WINAPI *GET_THEME_COLOR)(HTHEME hTheme, int iPartId, int iStateId, int iPropId, COLORREF *pClr);
typedef COLORREF (WINAPI *GET_THEME_SYS_COLOR)(HTHEME hTheme, int iColorId);
typedef HRESULT (WINAPI *GET_THEME_PART_SIZE)(HTHEME hTheme, HDC hdc, int iPartId, int iStateId, RECT *prc, THEMESIZE eSize, SIZE *psz);

typedef BOOL (WINAPI *IS_APP_THEMED)(void);
typedef HTHEME (WINAPI *OPEN_THEME_DATA)(HWND hWnd, LPCWSTR pszClassList);


// Klasse CTheme
class CTheme
{
public:
	// Konstruktor
	CTheme( );
	// Destruktor
	~CTheme( );
	// Funktionen
	BOOL CanUse( ){ return m_bCanUse; };
	// Funktionsvariablen
	CLOSE_THEME_DATA CloseThemeData;
	DRAW_THEME_BACKGROUND DrawThemeBackground;
	GET_THEME_COLOR GetThemeColor;
	GET_THEME_SYS_COLOR GetThemeSysColor;
	GET_THEME_PART_SIZE GetThemePartSize;
	IS_APP_THEMED IsAppThemed;
	OPEN_THEME_DATA OpenThemeData;
private:
	// Variablen
	BOOL m_bCanUse;
	HINSTANCE m_hUxThemeDLL;	
};


// Funktionen

// theTheme
CTheme & theTheme( );

#endif // _INC_CTHEME