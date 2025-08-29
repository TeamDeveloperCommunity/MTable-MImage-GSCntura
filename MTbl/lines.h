// lines.h
#ifndef _INC_LINES
#define _INC_LINES

#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif
#include "ctd.h"
#include "color.h"

#define MTLS_UNDEF -1
#define MTLS_INVISIBLE 0
#define MTLS_DOT 1
#define MTLS_SOLID 2

#define MTLT_HORZ 0
#define MTLT_VERT 1

#define MTBL_LINE_UNDEF_THICKNESS -1

// MTBLLINEDEFUDV
#pragma pack(UDV_PACK)
typedef struct MTBLLINEDEFUDV
{
#ifndef CTD1x
	BYTE bOffset[CLASS_UDV_OFFSET_SIZE];
#endif
	NUMBER Style;
	NUMBER Color;
	NUMBER Thickness;
}
MTBLLINEDEFUDV, *LPMTBLLINEDEFUDV;
#pragma pack()

// CMTblLineDef
class CMTblLineDef
{
public:
	int Style;
	COLORREF Color;
	int Thickness;

	// Kontruktoren
	CMTblLineDef();
	CMTblLineDef( CMTblLineDef& ld);

	// Destruktor
	CMTblLineDef::~CMTblLineDef();

	// Operatoren
	CMTblLineDef & operator=(CMTblLineDef &ld);
	BOOL operator==(CMTblLineDef &ld);
	BOOL operator!=(CMTblLineDef &ld);
	void operator+=(CMTblLineDef &ld);

	// Funktionen
	BOOL AnyUndef();
	BOOL FromUdv(HUDV hUdv);
	void Init();
	void ReplaceUndefinedFrom(CMTblLineDef &ld);
	BOOL ToUdv(HUDV hUdv);
};

typedef CMTblLineDef MTBLLINEDEF;
typedef CMTblLineDef* LPMTBLLINEDEF;

#endif // _INC_LINES