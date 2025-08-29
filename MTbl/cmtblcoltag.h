// cmtblcoltag.h

#ifndef _INC_CMTBLCOLTAG
#define _INC_CMTBLCOLTAG

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include "tstring.h"
#include "cmtblcolor.h"
#include "cmtblfont.h"

// Using
using namespace std;

// Klasse CMTblColTag
class CMTblColTag
{
public:
	// Member-Variablen
	CMTblFont m_Font, m_FontTbl;
	DWORD m_dwBackColor, m_dwBackColorTbl, m_dwTextColor, m_dwTextColorTbl;
	long m_lColSpan, m_lRowSpan, m_lTextIndent, m_lWidth, m_lHeight;
	tstring m_sAlign, m_sText, m_sVAlign;

	// Konstruktor
	CMTblColTag( );

	// Destruktor
	~CMTblColTag( );

	// Member-Funktionen
	void GetString( tstring &s );
	void Init( );
	void UseSpecialCharConversion( BOOL bUse );

private:
	// Member-Variablen
	BOOL m_bUseSpecialCharConversion;
	// Member-Funktionen
	void GetHTMLText( tstring &sText );
};


#endif // _INC_CMTBLCOLTAG