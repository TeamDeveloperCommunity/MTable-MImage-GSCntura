// mtexcexpapi.h

#ifndef _INC_MTEXCEXPAPI
#define _INC_MTEXCEXPAPI

#include "export.h"

extern "C"
{
// Funktionen
int WINAPI MTblExportToExcel( HWND hWndTbl, int iLanguage, LPCTSTR lpsExcelStartCell, DWORD dwExcelFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff );
int WINAPI MTblExportToOOCalc( HWND hWndTbl, int iLanguage, LPCTSTR lpsCalcStartCell, DWORD dwCalcFlags, DWORD dwExpFlags, WORD wRowFlagsOn, WORD wRowFlagsOff, DWORD dwColFlagsOn, DWORD dwColFlagsOff );
}

#endif // _INC_MTEXCEXPAPI