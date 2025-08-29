// GSCNTURA.CPP

// Includes
#include "gscntura.h"

// Interne Funktionen
int ArrGetItemCount( HARRAY hArr, int iDim )
{
	long lLower, lUpper;

	if ( Ctd.ArrayIsEmpty( hArr ) )
		return 0;

	if ( ! Ctd.ArrayGetLowerBound( hArr, iDim, &lLower ) )
		return -1;

	if ( ! Ctd.ArrayGetUpperBound( hArr, iDim, &lUpper ) )
		return -1;

	return ( lUpper - lLower ) + 1;

}

BOOL DoubleArrToNumArr( int iItems, double * Arr, HARRAY hArr )
{
	BOOL bError = FALSE;
	int i;
	long lLower;
	NUMBER n;

	if ( Ctd.ArrayGetLowerBound( hArr, 1, &lLower ) )
	{
		for( i = 0; i < iItems; i++ )
		{
			if ( ! Ctd.CvtDoubleToNumber(*(Arr + i), &n) )
			{
				bError = TRUE;
				break;
			}
			if ( ! Ctd.MDArrayPutNumber( hArr, &n, i + lLower ) )
			{
				bError = TRUE;
				break;
			}
		}
	}
	else
		bError = TRUE;

	return ( ! bError );
}

void FreeNewLPSTRArr( LPSTR *Arr, int iItems )
{
	if ( Arr )
	{
		for ( int i = 0; i < iItems; i++ )
		{
			if ( Arr[i] )
				free( (void*)Arr[i] );
		}
		free( (void*)Arr );
	}
}

double * NumArrToNew2DimDoubleArr( int iPoints, int iSets, HARRAY hArr )
{
	BOOL bError = FALSE;
	double * Arr, d;
	int i, j, k;
	long lLower;
	NUMBER n;

	// Untergrenze des Arrays ermitteln
	if ( ! Ctd.ArrayGetLowerBound( hArr, 1, &lLower ) )
		return NULL;

	// Speicher reservieren
	Arr = (double*)malloc( sizeof(double) * iPoints * iSets );
	if ( ! Arr )
		return NULL;
	
	// Daten übertragen
	for ( i = 0; i < iPoints; i++ )
	{
		for ( j = 0; j < iSets; j++ )
		{
			k = ( i * iSets ) + j;
			if ( ! Ctd.MDArrayGetNumber( hArr, &n, k + lLower ) )
			{
				bError = TRUE;
				break;
			}
			if ( ! Ctd.CvtNumberToDouble(&n, &d) )
			{
				bError = TRUE;
				break;
			}
			*(Arr + k) = d;
		}
	}

	// Fehler?
	if ( bError )
	{
		if ( Arr )
			free( (void*)Arr );
		return NULL;
	}

	return Arr;
}

double * NumArrToNewDoubleArr( int iItems, HARRAY hArr )
{
	BOOL bError = FALSE;
	double *Arr, v;
	int i;
	long lLower;
	NUMBER n;

	// Untergrenze des Arrays ermitteln
	if ( ! Ctd.ArrayGetLowerBound( hArr, 1, &lLower ) )
	{
		#ifdef SHOW_ERRORS
		MessageBox( NULL, "SalArrayGetLowerBound() failed", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
		#endif

		return NULL;
	}

	// Speicher reservieren
	Arr = (double*)malloc( sizeof(double) * iItems );
	if ( ! Arr )
	{
		#ifdef SHOW_ERRORS
		MessageBox( NULL, "Error allocating memory for array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
		#endif

		return NULL;
	}

	// Daten übertragen
	for ( i = 0; i < iItems; i++ )
	{
		if ( ! Ctd.MDArrayGetNumber( hArr, &n, i + lLower ) )
		{
			#ifdef SHOW_ERRORS
			char szMsg[255];
			wsprintf( szMsg, "SWinMDArrayGetNumber() failed at index %d", i + lLower );
			MessageBox( NULL, szMsg, "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			bError = TRUE;
			break;
		}
		if ( ! Ctd.CvtNumberToDouble(&n, &v) )
		{
			#ifdef SHOW_ERRORS
			MessageBox( NULL, "SWinCvtNumberToDouble() failed", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			bError = TRUE;
			break;
		}
		*(Arr + i) = v;
	}

	// Fehler?
	if ( bError )
	{
		if ( Arr )
			free( (void*)Arr );
		return NULL;
	}

	return Arr;
}

int * NumArrToNewIntArr( int iItems, HARRAY hArr )
{
	BOOL bError = FALSE;
	int * Arr, i, v;
	long lLower;
	NUMBER n;

	// Untergrenze des Arrays ermitteln
	if ( ! Ctd.ArrayGetLowerBound( hArr, 1, &lLower ) )
		return NULL;

	// Speicher reservieren
	Arr = (int*)malloc( sizeof(int) * iItems );
	if ( ! Arr )
		return NULL;

	// Daten übertragen
	for ( i = 0; i < iItems; i++ )
	{
		if ( ! Ctd.MDArrayGetNumber( hArr, &n, i + lLower ) )
		{
			bError = TRUE;
			break;
		}
		if ( ! Ctd.CvtNumberToInt(&n, &v) )
		{
			bError = TRUE;
			break;
		}
		*(Arr + i) = v;
	}

	// Fehler?
	if ( bError )
	{
		if ( Arr )
			free( (void*)Arr );
		return NULL;
	}

	return Arr;
}

WORD * NumArrToNewWORDArr( int iItems, HARRAY hArr )
{
	BOOL bError = FALSE;
	int i;
	long lLower;
	NUMBER n;
	WORD * Arr, v;

	// Untergrenze des Arrays ermitteln
	if ( ! Ctd.ArrayGetLowerBound( hArr, 1, &lLower ) )
		return NULL;

	// Speicher reservieren
	Arr = (WORD*)malloc( sizeof(WORD) * iItems );
	if ( ! Arr )
		return NULL;

	// Daten übertragen
	for ( i = 0; i < iItems; i++ )
	{
		if ( ! Ctd.MDArrayGetNumber( hArr, &n, i + lLower ) )
		{
			bError = TRUE;
			break;
		}
		if ( ! Ctd.CvtNumberToWord(&n, &v) )
		{
			bError = TRUE;
			break;
		}
		*(Arr + i) = v;
	}

	// Fehler?
	if ( bError )
	{
		if ( Arr )
			free( (void*)Arr );
		return NULL;
	}

	return Arr;
}

LPSTR * StrArrToNewLPSTRArr( int iItems, HARRAY hArr )
{
	BOOL bError = FALSE;
	HSTRING hs = 0;
	int i;
	LPSTR * Arr, lps, lpsbuf;
	long lLower, lLen;

	// Untergrenze des Arrays ermitteln
	if ( ! Ctd.ArrayGetLowerBound( hArr, 1, &lLower ) )
		return NULL;

	// Speicher reservieren
	Arr = (LPSTR*)calloc( iItems, sizeof(LPSTR) );
	if ( ! Arr )
		return NULL;

	// Daten übertragen
	for ( i = 0; i < iItems; i++ )
	{
		if ( ! Ctd.MDArrayGetHString( hArr, &hs, i + lLower ) )
		{
			bError = TRUE;
			break;
		}

		lps = NULL;

#if defined(TDASCII)
		lpsbuf = Ctd.StringGetBuffer( hs, &lLen );
		if (lpsbuf)
		{
			lLen = WideCharToMultiByte( CP_THREAD_ACP, 0, (LPWSTR)lpsbuf, -1, NULL, 0, NULL, NULL );
			if ( lLen != 0 )
			{
				lps = (LPSTR)malloc( lLen );
				if ( WideCharToMultiByte( CP_THREAD_ACP, 0, (LPWSTR)lpsbuf, -1, lps, lLen, NULL, NULL ) == 0 )
				{
					free( (void*)lps );
					lps = NULL;
				}
			}
		}
		
#else
		lpsbuf = Ctd.StringGetBuffer( hs, &lLen );
		if ( lpsbuf )
		{
			lLen = strlen( lpsbuf );
			lps = (LPSTR)malloc( lLen + 1 );
			if ( lps )
				strcpy( lps, lpsbuf );
		}
#endif
		if (lpsbuf)
			Ctd.HStrUnlock(hs);

		if ( !lps )
		{
			bError = TRUE;
			break;
		}
		*(Arr + i) = lps;
	}

	// Fehler?
	if ( bError )
	{
		if ( Arr )
			FreeNewLPSTRArr( Arr, iItems );
		return NULL;
	}

	return Arr;
}

// Exportierte Funktionen

/////////////////
//AG-Funktionen//
/////////////////

int WINAPI AGAmpCTD( int iPoints, int iSets, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iPoints < 0 || iSets < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNew2DimDoubleArr( iPoints, iSets, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGAmp( iPoints, iSets, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGAmpErrorCTD( int iPoints, int iSets, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iPoints < 0 || iSets < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNew2DimDoubleArr( iPoints, iSets, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGAmpError( iPoints, iSets, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGAuxCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGAux( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGClrCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGClr( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGDataLabelsCTD( int iMode, int iItems, HARRAY hArr )
{
	int iRet;
	LPSTR * Arr;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGDataLabels( iMode, iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI AGDataZCTD( int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGDataZ( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGDistCTD( int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGDist( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGDistErrorCTD( int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGDistError( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGFFTCTD( int iItems, HARRAY hArr, int iMode )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGFFT( iItems, Arr, iMode );

	if ( iRet == SUCCESS )
	{
		// Werte in Centura-Array zurückschreiben
		if ( ! DoubleArrToNumArr( iItems, Arr, hArr ) )
			iRet = FAILURE;		
	}

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGLabelsCTD( int iItems, HARRAY hArr )
{
	int iRet;
	LPSTR * Arr;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGLabels( iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI AGLabelYCTD( int iMode, int iItems, HARRAY hArr )
{
	int iRet;
	LPSTR * Arr;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGLabelY( iMode, iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI AGLabelZCTD( int iMode, int iItems, HARRAY hArr )
{
	int iRet;
	LPSTR * Arr;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGLabelZ( iMode, iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI AGLegendCTD( int iItems, HARRAY hArr )
{
	int iRet;
	LPSTR * Arr;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGLegend( iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI AGMissingLineStyleCTD( int iMode, int iItems, HARRAY hArrPatt, HARRAY hArrClr )
{
	int * ArrPatt, * ArrClr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	if ( iItems > 0 )
	{
		// Pattern-Daten in C-Array übertragen
		ArrPatt = NumArrToNewIntArr( iItems, hArrPatt );
		if ( ! ArrPatt )
			return FAILURE;

		// Color-Daten in C-Array übertragen
		ArrClr = NumArrToNewIntArr( iItems, hArrClr );
		if ( ! ArrClr )
			return FAILURE;
	}
	else
	{
		ArrPatt = NULL;
		ArrClr = NULL;
	}

	// GS-Funktion aufrufen
	iRet = AGMissingLineStyle( iMode, iItems, ArrPatt, ArrClr );

	// Speicher freigeben
	if ( ArrPatt )
		free( (void*)ArrPatt );
	if ( ArrClr )
		free( (void*)ArrClr );

	return iRet;
}

int WINAPI AGPattCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGPatt( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGSymCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGSym( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGTimeUpdateCTD( int iMode, int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGTimeUpdate( iMode, iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI AGTrendDataSetCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = AGTrendDataSet( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}


/////////////////
//GS-Funktionen//
/////////////////
int WINAPI GSDataAmpCTD( int iPoints, int iSets, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iPoints < 0 || iSets < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNew2DimDoubleArr( iPoints, iSets, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataAmp( iPoints, iSets, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataAmpErrCTD( int iPoints, int iSets, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iPoints < 0 || iSets < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNew2DimDoubleArr( iPoints, iSets, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataAmpErr( iPoints, iSets, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataAuxCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataAux( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataClrCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataClr( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataDistCTD( int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataDist( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataDistErrCTD( int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataDistErr( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataLabelsCTD( int iMode, int iPrec, int iCSet, int iTMode, int iClr, double dDataOffset, int iItems, HARRAY hArr )
{
	LPSTR * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataLabels( iMode, iPrec,iCSet, iTMode, iClr, dDataOffset, iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI GSDataPattCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataPatt( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataSymCTD( int iItems, HARRAY hArr )
{
	int * Arr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewIntArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataSym( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDataTransCTD( int iPoints, int iSets, HARRAY hArrAmp, HARRAY hArrDist, HARRAY hArrPatt, HARRAY hArrSym, HARRAY hArrAux, HARRAY hArrClr )
{
	double * ArrAmp, *ArrDist;
	int * ArrPatt, * ArrSym, * ArrAux, * ArrClr, iItems, iRet;

	// Parameter-Check
	if ( iPoints < 0 || iSets < 0 )
	{
		#ifdef SHOW_ERRORS
		MessageBox( NULL, "Invalid value for points or sets", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
		#endif

		return FAILURE;
	}

	// Amplitude-Array
	// Daten in C-Array übertragen
	ArrAmp = NumArrToNew2DimDoubleArr( iPoints, iSets, hArrAmp );
	if ( ! ArrAmp )
	{
		#ifdef SHOW_ERRORS
		MessageBox( NULL, "Error creating amplitude array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
		#endif

		return FAILURE;
	}


	// Distance-Array
	if ( Ctd.ArrayIsEmpty( hArrDist ) )
		ArrDist = NULL;
	else
	{
		iItems = iPoints;
		// Daten in C-Array übertragen
		ArrDist = NumArrToNewDoubleArr( iItems, hArrDist );
		if ( ! ArrDist )
		{
			#ifdef SHOW_ERRORS
			MessageBox( NULL, "Error creating distance array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			return FAILURE;
		}
	}

	// Pattern-Array
	if ( Ctd.ArrayIsEmpty( hArrPatt ) )
		ArrPatt = NULL;
	else
	{
		iItems = max( iPoints, iSets );
		// Daten in C-Array übertragen
		ArrPatt = NumArrToNewIntArr( iItems, hArrPatt );
		if ( ! ArrPatt )
		{
			#ifdef SHOW_ERRORS
			MessageBox( NULL, "Error creating pattern array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			return FAILURE;
		}
	}

	// Symbol-Array
	if ( Ctd.ArrayIsEmpty( hArrSym ) )
		ArrSym = NULL;
	else
	{
		iItems = max( iPoints, iSets );
		// Daten in C-Array übertragen
		ArrSym = NumArrToNewIntArr( iItems, hArrSym );
		if ( ! ArrSym )
		{
			#ifdef SHOW_ERRORS
			MessageBox( NULL, "Error creating symbol array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			return FAILURE;
		}
	}

	// Auxiliary-Array
	if ( Ctd.ArrayIsEmpty( hArrAux ) )
		ArrAux = NULL;
	else
	{
		iItems = max( iPoints, iSets );
		// Daten in C-Array übertragen
		ArrAux = NumArrToNewIntArr( iItems, hArrAux );
		if ( ! ArrAux )
		{
			#ifdef SHOW_ERRORS
			MessageBox( NULL, "Error creating auxiliary array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			return FAILURE;
		}
	}


	// Color-Array
	if ( Ctd.ArrayIsEmpty( hArrClr ) )
		ArrClr = NULL;
	else
	{
		iItems = max( iPoints, iSets );
		// Daten in C-Array übertragen
		ArrClr = NumArrToNewIntArr( iItems, hArrClr );
		if ( ! ArrClr )
		{
			#ifdef SHOW_ERRORS
			MessageBox( NULL, "Error creating color array", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
			#endif

			return FAILURE;
		}
	}

	// GS-Funktion aufrufen
	iRet = GSDataTrans( iPoints, iSets, ArrAmp, ArrDist, ArrPatt, ArrSym, ArrAux, ArrClr );

	#ifdef SHOW_ERRORS
	if ( iRet != SUCCESS )
		MessageBox( NULL, "GSDataTrans( ) failed", "MICSTO - GSCNTURA.DLL", MB_ICONERROR | MB_OK );
	#endif


	// Speicher freigeben
	free( (void*)ArrAmp );
	if ( ArrDist )
		free( (void*)ArrDist );
	if ( ArrPatt )
		free( (void*)ArrPatt );
	if ( ArrSym )
		free( (void*)ArrSym );
	if ( ArrAux )
		free( (void*)ArrAux );
	if ( ArrClr )
		free( (void*)ArrClr );

	return iRet;
}

int WINAPI GSDataZCTD( int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDataZ( iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSDefPattCTD( int iBitmap, HARRAY hArr )
{
	int iItems, iRet;
	WORD * Arr;

	// Parameter-Check
	if ( iBitmap < 0 || iBitmap > BRBITMAPMAX - 1 )
		return FAILURE;

	// Anzahl Elemente festlegen
	iItems = 8;

	// Daten in C-Array übertragen
	Arr = NumArrToNewWORDArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSDefPatt( iBitmap, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSLabelPieCTD( double dXOff, double dRad, double dWid, double dHt, int iItems, int iMode, int iCSet, int iTMode, int iClr, HARRAY hArr )
{
	LPSTR * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSLabelPie( dXOff, dRad, dWid, dHt, iItems, iMode, iCSet, iTMode, iClr, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI GSLabelXCTD( double dXOrg, double dYOrg, double dInc, double dWid, double dHt, int iItems, int iCSet, int iTMode, int iClr, HARRAY hArr )
{
	LPSTR * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSLabelX( dXOrg, dYOrg, dInc, dWid, dHt, iItems, iCSet, iTMode, iClr, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI GSLabelYCTD( double dXOrg, double dYOrg, double dInc, double dWid, double dHt, int iItems, int iCSet, int iTMode, int iClr, HARRAY hArr )
{
	LPSTR * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = StrArrToNewLPSTRArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSLabelY( dXOrg, dYOrg, dInc, dWid, dHt, iItems, iCSet, iTMode, iClr, (LPCSTR*)Arr );

	// Speicher freigeben
	FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}

int WINAPI GSLegendCTD( double dXOrg, double dYOrg, double dWid, double dHt, int iItems, int iRows, int iMode, int iCSet, int iTMode, int iClr, HARRAY hArrBClr, HARRAY hArrBPatt, HARRAY hArrLegs )
{
	LPSTR * ArrLegs;
	int * ArrBClr, * ArrBPatt, i, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Legenden-Texte in C-Array übertragen
	ArrLegs = StrArrToNewLPSTRArr( iItems, hArrLegs );
	if ( ! ArrLegs )
		return FAILURE;

	// Hintergrundfarben in C-Array übertragen
	i = ArrGetItemCount( hArrBClr, 1 );
	if ( i > 0 )
		ArrBClr = NumArrToNewIntArr( i, hArrBClr );
	else
		ArrBClr = NULL;

	// Hintergrundpattern in C-Array übertragen
	i = ArrGetItemCount( hArrBPatt, 1 );
	if ( i > 0 )
		ArrBPatt = NumArrToNewIntArr( i, hArrBPatt );
	else
		ArrBPatt = NULL;


	// GS-Funktion aufrufen
	iRet = GSLegend( dXOrg, dYOrg, dWid, dHt, iItems, iRows, iMode, iCSet, iTMode, iClr, (const int*)ArrBClr, (const int*)ArrBPatt, (LPCSTR*)ArrLegs );

	// Speicher freigeben
	FreeNewLPSTRArr( ArrLegs, iItems );
	if ( ArrBClr )
		free( (void*)ArrBClr );
	if ( ArrBPatt )
		free( (void*)ArrBPatt );

	return iRet;
}

int WINAPI GSMissingLineStyleCTD( int iMode, int iItems, HARRAY hArrPatt, HARRAY hArrClr )
{
	int * ArrPatt, * ArrClr, iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	if ( iItems > 0 )
	{
		// Pattern-Daten in C-Array übertragen
		ArrPatt = NumArrToNewIntArr( iItems, hArrPatt );
		if ( ! ArrPatt )
			return FAILURE;

		// Color-Daten in C-Array übertragen
		ArrClr = NumArrToNewIntArr( iItems, hArrClr );
		if ( ! ArrClr )
			return FAILURE;
	}
	else
	{
		ArrPatt = NULL;
		ArrClr = NULL;
	}

	// GS-Funktion aufrufen
	iRet = GSMissingLineStyle( iMode, iItems, ArrPatt, ArrClr );

	// Speicher freigeben
	if ( ArrPatt )
		free( (void*)ArrPatt );
	if ( ArrClr )
		free( (void*)ArrClr );

	return iRet;
}

int WINAPI GSTimeUpdateCTD( int iMode, int iItems, HARRAY hArr )
{
	double * Arr;
	int iRet;

	// Parameter-Check
	if ( iItems < 0 )
		return FAILURE;

	// Daten in C-Array übertragen
	Arr = NumArrToNewDoubleArr( iItems, hArr );
	if ( ! Arr )
		return FAILURE;

	// GS-Funktion aufrufen
	iRet = GSTimeUpdate( iMode, iItems, Arr );

	// Speicher freigeben
	free( (void*)Arr );

	return iRet;
}

int WINAPI GSWriteRegionFileCTD( int iMode, LPCSTR lpcsFile, LPCSTR lpcsTemplate, LPCSTR lpcsPolySpec, LPCSTR lpcsRectSpec, int iItems, HARRAY hArr )
{
	LPSTR * Arr;
	int iRet;

	// Daten in C-Array übertragen
	if ( iItems > 0 )
	{
		Arr = StrArrToNewLPSTRArr( iItems, hArr );
		if ( ! Arr )
			return FAILURE;
	}
	else
		Arr = NULL;

	// GS-Funktion aufrufen
	iRet = GSWriteRegionFile( iMode, lpcsFile, lpcsTemplate, lpcsPolySpec, lpcsRectSpec, iItems, (LPCSTR*)Arr );

	// Speicher freigeben
	if ( Arr )
		FreeNewLPSTRArr( Arr, iItems );

	return iRet;
}
