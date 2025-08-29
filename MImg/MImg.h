// MImg.h

#ifndef _INC_MIMG
#define _INC_MIMG

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <list>
#include "tstring.h"
#include "resource.h"
#include "MImgApi.h"
#include "MImgConst.h"
#include "CMImgMan.h"
#include "CMImgCtl.h"

// Globale Variablen
HANDLE ghInstance;

// Globale Objekte
CCtd Ctd;
CVis Vis;
CMImgMan ImgMan;

// Funktionen
CMImgCtl * MImgCtlGet( HWND hWndImg );
static LRESULT CALLBACK MImgCtlWndProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam );
BOOL MImgCtlRegisterClass( );

#endif // _INC_MIMG