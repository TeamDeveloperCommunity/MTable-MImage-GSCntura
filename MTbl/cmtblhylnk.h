// cmtblhylnk.h

#ifndef _INC_CMTBLHYLNK
#define _INC_CMTBLHYLNK

// Includes
#include <windows.h>

// Klasse CMTblHyLnk
#pragma pack(1)
class CMTblHyLnk
{
public:
	WORD m_wFlags;			// Flags

	// Konstruktoren
	CMTblHyLnk( );
	// Destruktor
	~CMTblHyLnk( );

	// Operatoren
	CMTblHyLnk & operator =( CMTblHyLnk & hl );

	// Funktionen
	BOOL IsEnabled( ) { return m_bEnabled; }
	void SetEnabled( BOOL bSet ) { m_bEnabled = bSet ? true : false; }
private:
	bool m_bEnabled;			// Enabled ja/nein

};
#pragma pack()

#endif // _INC_CMTBLHYLNK