// cmtblfont.h

#ifndef _INC_CMTBLFONT
#define _INC_CMTBLFONT

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <limits.h>
#include <algorithm>
#include "tstring.h"
#include "ctd.h"

// Using
using namespace std;

// Globale Variablen
extern tstring gsEmpty;
extern CCtd Ctd;

// Klasse CMTblFont
#pragma pack(1)
class CMTblFont
{
private:
	tstring m_sName;				// Name
	USHORT m_nSize;					// Größe
	BYTE m_nEnh;					// Stil(e)
	struct tagBits
	{
		BYTE NameUndef : 1;			// Name nicht definiert?
		BYTE SizeUndef : 1;			// Größe nicht definiert?
		BYTE EnhUndef : 1;			// Stil nicht definiert?
		BYTE : 0;					// Für Speicher-Ausrichtung
	}m_Bits;
public:
	// Konstanten
	static const tstring UNDEF_NAME;
	static const USHORT UNDEF_SIZE;
	static const BYTE UNDEF_ENH;

	// Konstruktor
	CMTblFont( );
	// Destruktor
	~CMTblFont( );

	// Operatoren
	CMTblFont & operator =( CMTblFont & f );
	BOOL operator ==( CMTblFont & f );
	BOOL operator !=( CMTblFont & f );

	// Funktionen
	BOOL AddEnh( long lEnh );
	BOOL AnyUndef( ) { return IsNameUndef( ) || IsSizeUndef( ) || IsEnhUndef( ); };
	HFONT Create( HDC hDC );
	BOOL Get( LPHSTRING phsName, LPLONG plSize, LPLONG plEnh );
	BOOL Get( tstring & sName, long & lSize, long & lEnh );
	long GetEnh( ) { return m_nEnh; }
	tstring &GetName( ) { return m_sName; }
	long GetSize( ) { return m_nSize; }
	BOOL IsEnhUndef( ) { return m_Bits.EnhUndef; }
	BOOL IsNameUndef( ) { return m_Bits.NameUndef; }
	BOOL IsSizeUndef( ) { return m_Bits.SizeUndef; }
	BOOL IsUndef( ) { return IsNameUndef( ) && IsSizeUndef( ) && IsEnhUndef( ); };
	BOOL Set( LPTSTR lpsName, long lSize, long lEnh );
	BOOL Set( HDC hDC, HFONT hFont );
	void SetEnhUndef( ) { m_nEnh = UNDEF_ENH; m_Bits.EnhUndef = 1; }
	void SetNameUndef( ) { m_sName = UNDEF_NAME; m_Bits.NameUndef = 1; }
	void SetSizeUndef( ) { m_nSize = UNDEF_SIZE; m_Bits.SizeUndef = 1; }
	void SetUndef( ) { SetNameUndef( ); SetSizeUndef( ); SetEnhUndef( ); }
	BOOL SubstUndef( CMTblFont & f );
};
#pragma pack()

#endif // _INC_CMTBLFONT