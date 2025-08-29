// gradrect.h

#ifndef _INC_GRADRECT
#define _INC_GRADRECT

// Includes
#include <windows.h>

// Defines
// Füllmodi
#define GRFM_LR		0	// von links nach rechts
#define GRFM_RL		1	// von rechts nach links
#define GRFM_TB		2	// von oben nach unten
#define GRFM_BT		3	// von unten nach oben
#define GRFM_TLBR	4	// von oben links nach unten rechts
#define GRFM_TRBL	5	// von oben rechts nach unten links
#define GRFM_BRTL	6	// von unten rechts nach oben links
#define GRFM_BLTR	7	// von unten links nach oben rechts

// Funktionen
inline COLOR16 GetBValue16( COLORREF c ) { return GetBValue( c ) << 8; };
inline COLOR16 GetBValue16( COLORREF c0, COLORREF c1 ) { return ( ( GetBValue( c0 ) + GetBValue( c1 ) ) / 2 ) << 8; };
inline COLOR16 GetGValue16( COLORREF c ) { return GetGValue( c ) << 8; };
inline COLOR16 GetGValue16( COLORREF c0, COLORREF c1 ) { return ( ( GetGValue( c0 ) + GetGValue( c1 ) ) / 2 ) << 8; };
inline COLOR16 GetRValue16( COLORREF c ) { return GetRValue( c ) << 8; };
inline COLOR16 GetRValue16( COLORREF c0, COLORREF c1 ) { return ( ( GetRValue( c0 ) + GetRValue( c1 ) ) / 2 ) << 8; };

BOOL GradientRect( HDC hDC, int x0, int y0, int x1, int y1, COLORREF c0, COLORREF c1, int iFillMode );
BOOL GradientRect( HDC hDC, RECT &r, COLORREF c0, COLORREF c1, int iFillMode );

#endif // _INC_GRADRECT