#include "mtblprint.h"
#include "mtblsubclass.h"

extern BOOL gbCtdHooked;
extern HANDLE ghInstance;
extern LanguageMap glm;
extern SubClassMap gscm;
extern UINT guiMsgOffs;
extern tstring gsEmpty;

// Funktionscontainer
extern CCtd Ctd;

/*
Liefert den Namen des aktuellen Standarddruckers als HSTRING
*/
HSTRING WINAPI MTblPrintGetDefPrinterName()
{
	DWORD dwBuffer = 0;
	LPTSTR lpsBuffer = NULL;
	HSTRING hs = 0;

	GetDefaultPrinter(NULL, &dwBuffer);
	if (dwBuffer)
	{
		lpsBuffer = new TCHAR[dwBuffer];
		if (GetDefaultPrinter(lpsBuffer, &dwBuffer))
		{
			Ctd.InitLPHSTRINGParam(&hs, Ctd.BufLenFromStrLen(dwBuffer));
			long lLen;
			LPTSTR lps = Ctd.StringGetBuffer(hs, &lLen);
			if (lps)
				lstrcpy(lps, lpsBuffer);
			Ctd.HStrUnlock(hs);
		}

		delete[]lpsBuffer;
	}

	return hs;
}

/*
Ermittelt verfügbare Drucker und fügt die Namen der Drucker zu dem übergebenen Array hinzu.
Gibt die Anzahl der verfügbaren Drucker zurück, -1 beim Auftreten eines Fehlers.
*/
long WINAPI MTblPrintGetPrinterNames(HARRAY arr)
{
	DWORD dwBufferSizeNeeded = 0;
	DWORD dwPrinterCount = 0;
	LPBYTE lpbPrinterStructs;
	HSTRING hs;
	PRINTER_INFO_4* ppi;
	long lLower;
	long result = -1;

	//Gütiges Array?
	if (!Ctd.ArrayGetLowerBound(arr, 1, &lLower)){
		return result;
	}

	//Benötigte Buffergröße ermitteln
	EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 4, NULL, 0, &dwBufferSizeNeeded, &dwPrinterCount);

	lpbPrinterStructs = new BYTE[dwBufferSizeNeeded];
	//Drucker Info Structs für alle verfügbaren Drucker ermitteln
	if (EnumPrinters(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, NULL, 4, lpbPrinterStructs, dwBufferSizeNeeded, &dwBufferSizeNeeded, &dwPrinterCount))
	{
		for (int i = 0; i < dwPrinterCount; i++)
		{
			//Druckername extrahieren
			ppi = (PRINTER_INFO_4*)(lpbPrinterStructs + (i * sizeof(PRINTER_INFO_4)));
			hs = 0;
			if (!Ctd.InitLPHSTRINGParam(&hs, Ctd.BufLenFromStrLen(lstrlen(ppi->pPrinterName)))){
				result = -1;
				break;
			}
			long lLen;
			LPTSTR lps = Ctd.StringGetBuffer(hs, &lLen);
			if (lps)
				lstrcpy(lps, ppi->pPrinterName);
			Ctd.HStrUnlock(hs);
			if (!lps)
			{
				result = -1;
				break;
			}

			//Druckername zum Array hinzufügen
			if (!Ctd.MDArrayPutHString(arr, hs, i + lLower))
			{
				result = -1;
				break;
			}
		}
		result = dwPrinterCount;
	}
	else
		result = -1;

	delete[] lpbPrinterStructs;
	return result;
}

/*
Sucht nach gespeicherten Druckereinstellungen und setzt diese auf dem übergebenen Drucker.
Änderungen der Druckereinstellungen werden ermittelt und als neue Druckereinstellung gespeichert.
Return 0, wenn erfolgreich, sonst error.
*/
int WINAPI MTblPrintSetupDlg(HWND hWndParent, HSTRING hsPrinter)
{
	int iResult = 0;
	LPTSTR lpsPrinter;
	long lprinterNameLength;
	HANDLE hPrinter;	//Handle für einen Drucker
	HSTRING hsDefPrinter = 0;
	PRINTER_INFO_2* ppi;
	DWORD dwBufferSizeNeeded = 0;
	LPBYTE lpbPrinterInfoStruct = NULL;
	LPDEVMODE lpBuf = NULL;
	long lLen, lRet = 0;
	LPDEVMODE lpDevModeBuffer = NULL;
	LPDEVMODE lpdmCurrentDevMode;
	tstring sPrinter = _T("");

	lpsPrinter = Ctd.StringGetBuffer(hsPrinter, &lprinterNameLength);
	if (lpsPrinter)
		sPrinter = lpsPrinter;
	Ctd.HStrUnlock(hsPrinter);

	if (!PrinterExists((LPTSTR)sPrinter.c_str()))
	{		
		return ERR_INVALID_PRINTER;
	}

	//Kein Drucker übergeben?
	if (sPrinter.size() == 0)
	{
		//Verwende Default Drucker
		hsDefPrinter = MTblPrintGetDefPrinterName();
		lpsPrinter = Ctd.StringGetBuffer(hsDefPrinter, &lprinterNameLength);
		if (lpsPrinter)
			sPrinter = lpsPrinter;
		Ctd.HStrUnlock(hsDefPrinter);

		//Kein Default Drucker gefunden?
		if (sPrinter.size() == 0)
		{
			Ctd.HStringUnRef(hsDefPrinter);
			return ERR_INVALID_PRINTER;
		}
	}

	//Handle für Drucker ermitteln
	OpenPrinter((LPTSTR)sPrinter.c_str(), &hPrinter, NULL);

	//Drucker DevMode ermitteln
	GetPrinter(hPrinter, 2, NULL, 0, &dwBufferSizeNeeded);
	lpbPrinterInfoStruct = new BYTE[dwBufferSizeNeeded];

	//Gucken, ob schon ein Device Mode für den übergebenen Drucker gespeichert ist.
	lpdmCurrentDevMode = GetPrinterDevMode(sPrinter.c_str());

	if (GetPrinter(hPrinter, 2, lpbPrinterInfoStruct, dwBufferSizeNeeded, &dwBufferSizeNeeded))
	{
		ppi = (PRINTER_INFO_2*)(lpbPrinterInfoStruct);
		if (lpdmCurrentDevMode != NULL)
		{
			lpBuf = (LPDEVMODE)GlobalLock(lpdmCurrentDevMode);
		}

		//Benötigte Länge des Buffers ermitteln
		lLen = DocumentProperties(hWndParent, hPrinter, ppi->pPrinterName, NULL, NULL, 0);
		if (lLen > 0)
		{
			lpDevModeBuffer = (LPDEVMODE)GlobalAlloc(0, lLen);
			//Druckerinitialisierung setzen
			if (lpBuf != NULL)
			{
				lRet = DocumentProperties(hWndParent, hPrinter, ppi->pPrinterName, lpDevModeBuffer, lpBuf, DM_IN_BUFFER | DM_IN_PROMPT | DM_OUT_BUFFER);
			}
			else
				lRet = DocumentProperties(hWndParent, hPrinter, ppi->pPrinterName, lpDevModeBuffer, NULL, DM_IN_PROMPT | DM_OUT_BUFFER);
		}
		ClosePrinter(hPrinter);
	}
	if (lpdmCurrentDevMode != NULL)
	{
		GlobalUnlock(lpdmCurrentDevMode);
	}

	//Druckerinitialisierung erfolgreich? Dann DevMode in Array hinzufügen
	if ((lRet == IDOK) && (lpDevModeBuffer != NULL))
	{
		lpdmCurrentDevMode = (LPDEVMODE)GlobalAlloc(GHND, lLen);
		if (lpdmCurrentDevMode != NULL)
		{
			lpBuf = (LPDEVMODE)GlobalLock(lpdmCurrentDevMode);
			CopyMemory(lpBuf, lpDevModeBuffer, lLen);
			GlobalUnlock(lpdmCurrentDevMode);
			SetPrinterDevMode(sPrinter, lpBuf);
		}
	}

	if (lpDevModeBuffer != NULL)
	{
		GlobalFree(lpDevModeBuffer);
	}

	delete[] lpbPrinterInfoStruct;
	if (hsDefPrinter != 0)
		Ctd.HStringUnRef(hsDefPrinter);
	if (hsPrinter != 0)
		Ctd.HStringUnRef(hsPrinter);
	return iResult;
}

/*
Druckt die übergebene Tabelle mit den übergebenen Parametern.
Return 1 wenn erfolgreich, sonst error.
*/
int WINAPI MTblPrint(HWND hWndTbl, HUDV hParams)
{
	LPMTBLSUBCLASS lpSubClass = MTblGetSubClass(hWndTbl);
	if (!lpSubClass)
	{
		return ERR_NOT_SUBCLASSED;
	}

	//Druckoptionen aus Parametern erstellen
	PRINTOPTIONS printOptions;
	int iResult = 0;

	if (printOptions.FromUdv(hParams))
	{
		//Drucken
		iResult = lpSubClass->Print(hWndTbl, printOptions);
	}

	return iResult;
}