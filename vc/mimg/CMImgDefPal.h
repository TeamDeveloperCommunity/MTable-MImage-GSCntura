// CMImgDefPal.h

#ifndef _INC_CMIMGDEFPAL
#define _INC_CMIMGDEFPAL

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>

// Klasse CMImgDefPal
class CMImgDefPal
{
private:
	HPALETTE m_hPal;
	// Methoden
	void CreatePalette( );
	void DeletePalette( );
public:
	// Konstruktoren
	CMImgDefPal( );
	// Destruktor
	~CMImgDefPal( );
	// Methoden
	HPALETTE GetHandle( );
};

// Funktionen
HPALETTE theMImgDefPal( );

#endif // _INC_CMIMGDEFPAL
