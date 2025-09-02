// cmtblimg.h

#ifndef _INC_CMTBLIMG
#define _INC_CMTBLIMG

// Includes
#include <windows.h>
#include <limits.h>
#include "MImgApi.h"
#include "mtimg.h"
#include "cmtblcounter.h"

// Klasse CMTblImg
class CMTblImg
{
private:
	// Member-Variablen
	HANDLE m_hImage;
	WORD m_wFlags;
	CMTblCounter *m_pCounter;

	// Funktionen
	void InitMembers( );
public:
	// Konstruktoren
	CMTblImg( );
	// Destruktor
	~CMTblImg( );

	// Operatoren
	CMTblImg & operator =( CMTblImg & f );

	// Funktionen
	BOOL Draw( HDC hDC, LPRECT lpRect, DWORD dwFlags = 0 );
	BOOL GetSize( SIZE & s );
	// Funktionen
	void ClearFlags( );
	WORD GetFlags( ) { return m_wFlags; };
	BOOL HasOpacity( );
	BOOL IsNoSelInvSet( ) { return QueryFlags( MTSI_NOSELINV ); };
	BOOL QueryFlags( DWORD dwFlags ){ return ( m_wFlags & dwFlags ) ? TRUE : FALSE; };
	BOOL SetDrawClipRect( LPRECT pRect );
	BOOL SetFlags( DWORD dwFlags, BOOL bSet );
	BOOL SetHandle( HANDLE hImage, CMTblCounter *pCounter );
	// Statische Funktionen
	static BOOL IsValidImageHandle( HANDLE hImage ) { return DWORD(hImage) != 0 && DWORD(hImage) != MTBL_HIMG_NODEFAULT; };

	// Inline-Funktionen
	HANDLE GetHandle( ) { return m_hImage; };
};

#endif // _INC_CMTBLIMG