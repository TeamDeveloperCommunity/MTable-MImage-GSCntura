// vis.h

#ifndef _INC_VIS
#define _INC_VIS

// Pragmas
#pragma warning(disable:4786)

// Includes
#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif
#include <stdio.h>
#include <tchar.h>

// Vorab-Deklarationen
struct  TREE;
typedef TREE FAR* LPTREE;

// Definitionen
#define PIC_LoadSWinStr				0x0100
#define PIC_LoadResource			0x0200
#define PIC_LoadFile				0x0400
#define PIC_LoadSWinRes				0x0800
#define PIC_LoadSmallIcon			0x1000
#define PIC_LoadLargeIcon			0x2000

// Typedefs Funktionen
typedef LPTREE(WINAPI* VISLISTGETITEMHANDLE)(HWND, INT);
typedef LPVOID (WINAPI *VISPICLOAD)(WORD,LPCTSTR,LPCTSTR);
typedef LONG (WINAPI* VISTBLFINDSTRING)(HWND, LONG, HWND, LPTSTR);

// Klassen
class CVis
{
private:
	// Variablen
	HINSTANCE m_hInst;
	BOOL m_Initialized;
public:
	// Funktionsvariablen
	VISLISTGETITEMHANDLE ListGetItemHandle;
	VISPICLOAD PicLoad;
	VISTBLFINDSTRING TblFindString;
	// Konstruktor
	CVis( void );
	// Destruktor
	~CVis( void );

	// Funktionen
	BOOL IsPresent( ) { return m_Initialized; };
};

#endif // _INC_VIS