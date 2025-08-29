// MImgCtl.h

#ifndef _INC_MIMGCTL
#define _INC_MIMGCTL

// Pragmas
#pragma warning(disable:4786)

// Includes
#include <windows.h>
#include <algorithm>
#include <string>
#include <list>
#include "resource.h"
#include "MImgConst.h"
#include "CMImgMan.h"

// Globale Variablen
HANDLE ghInstance;

// Globale Objekte
CCtd Ctd;
CVis Vis;
CMImgMan ImgMan;

#endif // _INC_MIMG