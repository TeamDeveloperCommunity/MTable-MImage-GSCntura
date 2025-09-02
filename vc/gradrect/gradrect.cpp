// gradrect.cpp

// Includes
#include "gradrect.h"

// Funktionen
BOOL GradientRect( HDC hDC, int x0, int y0, int x1, int y1, COLORREF c0, COLORREF c1, int iFillMode )
{
	BOOL bSwitchColors =
		iFillMode == GRFM_RL ||
		iFillMode == GRFM_BT ||
		iFillMode == GRFM_BRTL ||
		iFillMode == GRFM_BLTR;

	if ( bSwitchColors )
	{
		COLORREF c = c0;
		c0 = c1;
		c1 = c;
	}

	ULONG index[] = { 0, 1, 2, 0, 1, 3 };
	TRIVERTEX vert[4] = 
	{
		{ x0, y0, GetRValue16( c0 ), GetGValue16( c0 ), GetBValue16( c0 ), 0 },
		{ x1, y1, GetRValue16( c1 ), GetGValue16( c1 ), GetBValue16( c1 ), 0 },
		{ x0, y1, GetRValue16( c0, c1 ), GetGValue16( c0, c1 ), GetBValue16( c0, c1 ), 0 },
		{ x1, y0, GetRValue16( c0, c1 ), GetGValue16( c0, c1 ), GetBValue16( c0, c1 ), 0 }
	};

	switch ( iFillMode )
	{
	case GRFM_LR:
	case GRFM_RL:
		return GradientFill( hDC, vert, 2, index, 1, GRADIENT_FILL_RECT_H );
	case GRFM_TB:
	case GRFM_BT:
		return GradientFill( hDC, vert, 2, index, 1, GRADIENT_FILL_RECT_V );
	case GRFM_TLBR:
	case GRFM_BRTL:
		return GradientFill( hDC, vert, 4, index, 2, GRADIENT_FILL_TRIANGLE );
	case GRFM_TRBL:
	case GRFM_BLTR:
		vert[0].x = x1;
		vert[3].x = x0;
		vert[1].x = x0;
		vert[2].x = x1;
		return GradientFill( hDC, vert, 4, index, 2, GRADIENT_FILL_TRIANGLE );
	default:
		return FALSE;
	}

	return TRUE;
}

BOOL GradientRect( HDC hDC, RECT &r, COLORREF c0, COLORREF c1, int iFillMode )
{
	return GradientRect( hDC, r.left, r.top, r.right, r.bottom, c0, c1, iFillMode );
}

