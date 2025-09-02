// cmtblbtn.h

#ifndef _INC_CMTBLBTN
#define _INC_CMTBLBTN

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include "tstring.h"
#include "btnflags.h"

// Using
using namespace std;

// Klasse CMTblBtn
#pragma pack(1)
class CMTblBtn
{
public:
	bool m_bShow;				// Anzeigen ja/nein
	int m_iWidth;				// Breite
	tstring m_sText;			// Text
	HANDLE m_hImage;			// Bild
	DWORD m_dwFlags;			// Flags

	// Konstruktoren
	CMTblBtn( );
	// Destruktor
	~CMTblBtn( );
	// Funktionen
	BOOL IsDisabled( ) { return m_dwFlags & MTBL_BTN_DISABLED ? TRUE : FALSE; }
	BOOL IsLeft( ) { return m_dwFlags & MTBL_BTN_LEFT ? TRUE : FALSE; }
	// Operatoren
	CMTblBtn & operator =( CMTblBtn & b );
};
#pragma pack()

#endif // _INC_CMTBLBTN