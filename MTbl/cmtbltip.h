// cmtbltip.h

#ifndef _INC_CMTBLTIP
#define _INC_CMTBLTIP

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <limits.h>
#include "tstring.h"
#include "ctd.h"
#include "cmtblcolor.h"
#include "cmtblfont.h"
#include "tipflags.h"

// Using
using namespace std;

// Klasse CMTblTip
#pragma pack(1)
class CMTblTip
{
public:
	// Member-Variablen
	tstring m_sText;			// Text
	int m_iPosX;			// Position X
	int m_iPosY;			// Position Y
	int m_iWid;				// Breite
	int m_iHt;				// Höhe
	CMTblFont m_Font;		// Font
	CMTblColor m_TextColor;	// Textfarbe
	CMTblColor m_BackColor;	// Hintergrundfarbe
	CMTblColor m_FrameColor;// Rahmenfarbe
	int m_iDelay;			// Anzeigeverzögerung in Millisekunden
	int m_iOffsX;			// Maus-Offset X
	int m_iOffsY;			// Maus-Offset Y
	int m_iRCEWid;			// Breite Round Corner Ellipse
	int m_iRCEHt;			// Höhe Round Corner Ellipse
	int m_iOpacity;			// Opazität
	int m_iFadeInTime;		// Einblendzeit in Millisekunden
	WORD m_wFlags;			// Flags

	CMTblTip * m_pDefTip;	// Zeiger auf Default-Tip

	// Konstruktoren
	CMTblTip( );

	// Methoden
	HFONT CreateFont( HDC hDC, BOOL bMyValue = FALSE );
	DWORD GetBackColor( BOOL bMyValue = FALSE );
	int & GetDelay( BOOL bMyValue = FALSE );
	__int64 GetDelayI64( BOOL bMyValue = FALSE );
	int & GetFadeInTime( BOOL bMyValue = FALSE );
	DWORD GetFrameColor( BOOL bMyValue = FALSE );
	int & GetHt( BOOL bMyValue = FALSE );
	int & GetOffsX( BOOL bMyValue = FALSE );
	int & GetOffsY( BOOL bMyValue = FALSE );
	int & GetOpacity( BOOL bMyValue = FALSE );
	int & GetPosX( BOOL bMyValue = FALSE );
	int & GetPosY( BOOL bMyValue = FALSE );
	int & GetRCEHt( BOOL bMyValue = FALSE );
	int & GetRCEWid( BOOL bMyValue = FALSE );
	tstring & GetText( BOOL bMyValue = FALSE );
	DWORD GetTextColor( BOOL bMyValue = FALSE );
	int & GetWid( BOOL bMyValue = FALSE );
	BOOL MustShow( );
	BOOL QueryFlags( DWORD dwFlags, BOOL bMyValue = FALSE );
	void SetDefTip( CMTblTip *pDefTip );
	BOOL SetFlags( DWORD dwFlags, BOOL bSet );
};
#pragma pack()

#endif // _INC_CMTBLTIP