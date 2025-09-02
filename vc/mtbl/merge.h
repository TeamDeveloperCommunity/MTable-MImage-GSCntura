// merge.h

#ifndef _INC_MERGE
#define _INC_MERGE

#pragma warning(disable:4786)

// Includes
#ifdef _AFXDLL
#include "stdafx.h"
#else
#include <windows.h>
#endif
#include <map>
#include <set>
#include <vector>

// Using
using namespace std;

// Find*MergeCell - Modi
#define FMC_MASTER 0
#define FMC_SLAVE 1
#define FMC_ANY 2

// Find*MergeRow - Modi
#define FMR_MASTER 0
#define FMR_SLAVE 1
#define FMR_ANY 2

// Typedefs
typedef pair< long,long > LONG_PAIR;
typedef pair< LONG_PAIR,LONG_PAIR > LONG_PAIR_PAIR;

// MTBLMERGECELL
#pragma pack(1)
typedef struct MTBLMERGECELL
{
	int iType;

	HWND hWndColFrom;
	int iColPosFrom;
	long lRowFrom;
	long lRowVisPosFrom;

	HWND hWndColTo;
	int iColPosTo;
	long lRowTo;
	long lRowVisPosTo;

	int iMergedCols;
	int iMergedRows;
}
MTBLMERGECELL, *LPMTBLMERGECELL;

typedef map< LONG_PAIR,LPMTBLMERGECELL > MTBLMERGECELL_MAP;
typedef set< LPMTBLMERGECELL > MTBLMERGECELL_SET;
typedef vector< LPMTBLMERGECELL > MTBLMERGECELL_VECTOR;

typedef struct MTBLMERGECELLS
{
	MTBLMERGECELL_MAP mcm;

	WORD wRowFlagsOn;
	WORD wRowFlagsOff;

	DWORD dwColFlagsOn;
	DWORD dwColFlagsOff;
}
MTBLMERGECELLS, *LPMTBLMERGECELLS;
#pragma pack()

// MTBLMERGEROW
#pragma pack(1)
typedef struct MTBLMERGEROW
{
	long lRowFrom;
	long lRowVisPosFrom;

	long lRowTo;
	long lRowVisPosTo;

	int iMergedRows;

	void CopyTo( MTBLMERGEROW &mr );
	void Init( );
}
MTBLMERGEROW, *LPMTBLMERGEROW;

typedef map< long,LPMTBLMERGEROW > MTBLMERGEROW_MAP;
typedef set< LPMTBLMERGEROW > MTBLMERGEROW_SET;
typedef vector< LPMTBLMERGEROW > MTBLMERGEROW_VECTOR;

typedef struct MTBLMERGEROWS
{
	MTBLMERGEROW_MAP mrm;

	WORD wRowFlagsOn;
	WORD wRowFlagsOff;
}
MTBLMERGEROWS, *LPMTBLMERGEROWS;
#pragma pack()

#endif // _INC_MERGE
