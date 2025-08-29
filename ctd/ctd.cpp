// ctd.cpp

// Includes
#include "ctd.h"

// Konstruktor
CCtd::CCtd( void )
{
	memset( this, 0, sizeof(CCtd) );

	// DLL laden
	TCHAR szDLL[MAX_PATH] = _T("");

	#ifdef TD11
	lstrcpy( szDLL, _T("CDLLI11.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD15
	lstrcpy( szDLL, _T("CDLLI15.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD20
	lstrcpy( szDLL, _T("CDLLI20.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD21
	lstrcpy( szDLL, _T("CDLLI21.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD30
	lstrcpy( szDLL, _T("CDLLI30.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD31
	lstrcpy( szDLL, _T("CDLLI31.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD40
	lstrcpy( szDLL, _T("CDLLI40.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD41
	lstrcpy( szDLL, _T("CDLLI41.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD42
	lstrcpy( szDLL, _T("CDLLI42.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD51
	lstrcpy( szDLL, _T("CDLLI51.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD52
	lstrcpy( szDLL, _T("CDLLI52.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD60
	lstrcpy( szDLL, _T("CDLLI60.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD61
	lstrcpy( szDLL, _T("CDLLI61.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD62
	lstrcpy( szDLL, _T("CDLLI62.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD63
	lstrcpy( szDLL, _T("CDLLI63.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD70
	lstrcpy( szDLL, _T("CDLLI70.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD71
	lstrcpy( szDLL, _T("CDLLI71.DLL") );
	m_hInst = LoadLibrary( szDLL );
	#elif TD72
	lstrcpy(szDLL, _T("CDLLI72.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#elif TD73
	lstrcpy(szDLL, _T("CDLLI73.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#elif TD74
	lstrcpy(szDLL, _T("CDLLI74.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#elif TD75
	lstrcpy(szDLL, _T("CDLLI75.DLL"));
	m_hInst = LoadLibrary(szDLL);
	#endif

	if ( m_hInst )
	{
		ArrayDimCount = (SALARRAYDIMCOUNT)GetProcAddress( m_hInst, "SalArrayDimCount" );
		ArrayGetLowerBound = (SALARRAYGETLOWERBOUND)GetProcAddress( m_hInst, "SalArrayGetLowerBound" );
		ArrayGetUpperBound = (SALARRAYGETUPPERBOUND)GetProcAddress( m_hInst, "SalArrayGetUpperBound" );
		ArrayIsEmpty = (SALARRAYISEMPTY)GetProcAddress( m_hInst, "SalArrayIsEmpty" );
		ColorGet = (SALCOLORGET)GetProcAddress( m_hInst, "SalColorGet" );
		CompileAndEvaluate = (SALCOMPILEANDEVALUATE)GetProcAddress( m_hInst, "SalCompileAndEvaluate" );
#if defined(TDASCII)
		CompileAndEvaluateA = (SALCOMPILEANDEVALUATEA)GetProcAddress( m_hInst, "SalCompileAndEvaluateA" );
#endif
		ContextCurrent = (SALCONTEXTCURRENT)GetProcAddress( m_hInst, "SalContextCurrent" );
		DateToStr = (SALDATETOSTR)GetProcAddress( m_hInst, "SalDateToStr" );
		DisableWindow = (SALDISABLEWINDOW)GetProcAddress( m_hInst, "SalDisableWindow" );
		DragDropStart = (SALDRAGDROPSTART)GetProcAddress( m_hInst, "SalDragDropStart" );
		DragDropStop = (SALDRAGDROPSTOP)GetProcAddress( m_hInst, "SalDragDropStop" );
		EnableWindow = (SALENABLEWINDOW)GetProcAddress( m_hInst, "SalEnableWindow" );
		FmtFieldToStr = (SALFMTFIELDTOSTR)GetProcAddress( m_hInst, "SalFmtFieldToStr" );
		FmtGetPicture = (SALFMTGETPICTURE)GetProcAddress( m_hInst, "SalFmtGetPicture" );
		FmtGetInputMask = (SALFMTGETINPUTMASK)GetProcAddress( m_hInst, "SalFmtGetInputMask" );
		FormUnitsToPixels = (SALFORMUNITSTOPIXELS)GetProcAddress( m_hInst, "SalFormUnitsToPixels" );
		GetDataType = (SALGETDATATYPE)GetProcAddress( m_hInst, "SalGetDataType" );
		GetItemName = (SALGETITEMNAME)GetProcAddress( m_hInst, "SalGetItemName" );
		GetMaxDataLength = (SALGETMAXDATALENGTH)GetProcAddress( m_hInst, "SalGetMaxDataLength" );
		GetType = (SALGETTYPE)GetProcAddress( m_hInst, "SalGetType" );
		GetVersion = (SALGETVERSION)GetProcAddress( m_hInst, "SalGetVersion" );
		HStringUnRef = (SALHSTRINGUNREF)GetProcAddress( m_hInst, "SalHStringUnRef" );
		InvalidateWindow = (SALINVALIDATEWINDOW)GetProcAddress( m_hInst, "SalInvalidateWindow" );
		IsWindowEnabled = (SALISWINDOWENABLED)GetProcAddress( m_hInst, "SalIsWindowEnabled" );
		IsWindowVisible = (SALISWINDOWVISIBLE)GetProcAddress( m_hInst, "SalIsWindowVisible" );
		ListAdd = (SALLISTADD)GetProcAddress( m_hInst, "SalListAdd" );
		ListClear = (SALLISTCLEAR)GetProcAddress( m_hInst, "SalListClear" );
		ListQueryCount = (SALLISTQUERYCOUNT)GetProcAddress( m_hInst, "SalListQueryCount" );
		ListQuerySelection = (SALLISTQUERYSELECTION)GetProcAddress( m_hInst, "SalListQuerySelection" );
		ListQueryText = (SALLISTQUERYTEXT)GetProcAddress( m_hInst, "SalListQueryText" );
		ListSetSelect = (SALLISTSETSELECT)GetProcAddress( m_hInst, "SalListSetSelect" );		
		NumberToHString = (SALNUMBERTOHSTRING)GetProcAddress( m_hInst, "SalNumberToHString" );
		ParentWindow = (SALPARENTWINDOW)GetProcAddress( m_hInst, "SalParentWindow" );
		PicGetString = (SALPICGETSTRING)GetProcAddress( m_hInst, "SalPicGetString" );
		PixelsToFormUnits = (SALPIXELSTOFORMUNITS)GetProcAddress( m_hInst, "SalPixelsToFormUnits" );
		QueryFieldEdit = (SALQUERYFIELDEDIT)GetProcAddress( m_hInst, "SalQueryFieldEdit" );
		SendValidateMsg = (SALSENDVALIDATEMSG)GetProcAddress( m_hInst, "SalSendValidateMsg" );
		SetWindowLoc = (SALSETWINDOWLOC)GetProcAddress( m_hInst, "SalSetWindowLoc" );
		SetWindowSize = (SALSETWINDOWSIZE)GetProcAddress( m_hInst, "SalSetWindowSize" );
		StatusSetText = (SALSTATUSSETTEXT)GetProcAddress( m_hInst, "SalStatusSetText" );
#if defined(TDASCII)
		StatusSetTextA = (SALSTATUSSETTEXTA)GetProcAddress( m_hInst, "SalStatusSetTextA" );
#endif
		StrIsValidNumber = (SALSTRISVALIDNUMBER)GetProcAddress( m_hInst, "SalStrIsValidNumber" );
#if defined(TDASCII)
		StrIsValidNumberA = (SALSTRISVALIDNUMBERA)GetProcAddress( m_hInst, "SalStrIsValidNumberA" );
#endif
		StrToDate = (SALSTRTODATE)GetProcAddress( m_hInst, "SalStrToDate");
#if defined(TDASCII)
		StrToDateA = (SALSTRTODATEA)GetProcAddress( m_hInst, "SalStrToDateA");
#endif
		StrToNumber = (SALSTRTONUMBER)GetProcAddress( m_hInst, "SalStrToNumber" );
#if defined(TDASCII)
		StrToNumberA = (SALSTRTONUMBERA)GetProcAddress( m_hInst, "SalStrToNumberA" );
#endif
#ifdef TDTHEMES
		ThemeGet = (SALTHEMEGET)GetProcAddress( m_hInst, "SalThemeGet" );
#endif
		UpdateWindow = (SALUPDATEWINDOW)GetProcAddress( m_hInst, "SalUpdateWindow" );
		WindowGetProperty = (SALWINDOWGETPROPERTY)GetProcAddress( m_hInst, "SalWindowGetProperty" );
#if defined(TDASCII)
		WindowGetPropertyA = (SALWINDOWGETPROPERTYA)GetProcAddress( m_hInst, "SalWindowGetPropertyA" );
#endif
		ValidateSet = (SALVALIDATESET)GetProcAddress( m_hInst, "SalValidateSet" );

		TblAnyRows = (SALTBLANYROWS)GetProcAddress( m_hInst, "SalTblAnyRows" );
		TblClearSelection = (SALTBLCLEARSELECTION)GetProcAddress( m_hInst, "SalTblClearSelection" );
		TblDefineCheckBoxColumn = (SALTBLDEFINECHECKBOXCOLUMN)GetProcAddress( m_hInst, "SalTblDefineCheckBoxColumn" );
		TblDefineDropDownListColumn = (SALTBLDEFINEDROPDOWNLISTCOLUMN)GetProcAddress( m_hInst, "SalTblDefineDropDownListColumn" );
		TblDefinePopupEditColumn = (SALTBLDEFINEPOPUPEDITCOLUMN)GetProcAddress( m_hInst, "SalTblDefinePopupEditColumn" );
		TblDefineRowHeader = (SALTBLDEFINEROWHEADER)GetProcAddress( m_hInst, "SalTblDefineRowHeader" );
		TblDefineSplitWindow = (SALTBLDEFINESPLITWINDOW)GetProcAddress( m_hInst, "SalTblDefineSplitWindow" );
		TblDeleteRow = (SALTBLDELETEROW)GetProcAddress( m_hInst, "SalTblDeleteRow" );
		TblFetchRow = (SALTBLFETCHROW)GetProcAddress( m_hInst, "SalTblFetchRow" );
		TblFindNextRow = (SALTBLFINDNEXTROW)GetProcAddress( m_hInst, "SalTblFindNextRow" );
		TblFindPrevRow = (SALTBLFINDPREVROW)GetProcAddress( m_hInst, "SalTblFindPrevRow" );
		TblGetColumnText = (SALTBLGETCOLUMNTEXT)GetProcAddress( m_hInst, "SalTblGetColumnText" );
		TblGetColumnTitle = (SALTBLGETCOLUMNTITLE)GetProcAddress( m_hInst, "SalTblGetColumnTitle" );
		TblGetColumnWindow = (SALTBLGETCOLUMNWINDOW)GetProcAddress( m_hInst, "SalTblGetColumnWindow" );
		TblInsertRow = (SALTBLINSERTROW)GetProcAddress( m_hInst, "SalTblInsertRow" );
		TblKillEdit = (SALTBLKILLEDIT)GetProcAddress( m_hInst, "SalTblKillEdit" );
		TblKillFocus = (SALTBLKILLFOCUS)GetProcAddress( m_hInst, "SalTblKillFocus" );
		TblObjectsFromPoint = (SALTBLOBJECTSFROMPOINT)GetProcAddress( m_hInst, "SalTblObjectsFromPoint" );
		TblQueryCheckBoxColumn = (SALTBLQUERYCHECKBOXCOLUMN)GetProcAddress( m_hInst, "SalTblQueryCheckBoxColumn" );
		TblQueryColumnCellType = (SALTBLQUERYCOLUMNCELLTYPE)GetProcAddress( m_hInst, "SalTblQueryColumnCellType" );
		TblQueryColumnFlags = (SALTBLQUERYCOLUMNFLAGS)GetProcAddress( m_hInst, "SalTblQueryColumnFlags" );
		TblQueryColumnID = (SALTBLQUERYCOLUMNID)GetProcAddress( m_hInst, "SalTblQueryColumnID" );
		TblQueryColumnPos = (SALTBLQUERYCOLUMNPOS)GetProcAddress( m_hInst, "SalTblQueryColumnPos" );
		TblQueryColumnWidth = (SALTBLQUERYCOLUMNWIDTH)GetProcAddress( m_hInst, "SalTblQueryColumnWidth" );
		TblQueryContext = (SALTBLQUERYCONTEXT)GetProcAddress( m_hInst, "SalTblQueryContext" );
		TblQueryDropDownListColumn = (SALTBLQUERYDROPDOWNLISTCOLUMN)GetProcAddress( m_hInst, "SalTblQueryDropDownListColumn" );
		TblQueryFocus = (SALTBLQUERYFOCUS)GetProcAddress( m_hInst, "SalTblQueryFocus" );
		TblQueryLinesPerRow = (SALTBLQUERYLINESPERROW)GetProcAddress( m_hInst, "SalTblQueryLinesPerRow" );
		TblQueryLockedColumns = (SALTBLQUERYLOCKEDCOLUMNS)GetProcAddress( m_hInst, "SalTblQueryLockedColumns" );
		TblQueryPopupEditColumn = (SALTBLQUERYPOPUPEDITCOLUMN)GetProcAddress( m_hInst, "SalTblQueryPopupEditColumn" );
		TblQueryRowFlags = (SALTBLQUERYROWFLAGS)GetProcAddress( m_hInst, "SalTblQueryRowFlags" );
		TblQueryRowHeader = (SALTBLQUERYROWHEADER)GetProcAddress( m_hInst, "SalTblQueryRowHeader" );
		TblQueryScroll = (SALTBLQUERYSCROLL)GetProcAddress( m_hInst, "SalTblQueryScroll" );
		TblQuerySplitWindow = (SALTBLQUERYSPLITWINDOW)GetProcAddress( m_hInst, "SalTblQuerySplitWindow" );
		TblQueryTableFlags = (SALTBLQUERYTABLEFLAGS)GetProcAddress( m_hInst, "SalTblQueryTableFlags" );
		TblQueryVisibleRange = (SALTBLQUERYVISIBLERANGE)GetProcAddress( m_hInst, "SalTblQueryVisibleRange" );
		TblScroll = (SALTBLSCROLL)GetProcAddress( m_hInst, "SalTblScroll" );
		TblSetColumnFlags = (SALTBLSETCOLUMNFLAGS)GetProcAddress( m_hInst, "SalTblSetColumnFlags" );
		TblSetColumnText = (SALTBLSETCOLUMNTEXT)GetProcAddress( m_hInst, "SalTblSetColumnText" );
#if defined(TDASCII)
		TblSetColumnTextA = (SALTBLSETCOLUMNTEXTA)GetProcAddress( m_hInst, "SalTblSetColumnTextA" );
#endif
		TblSetColumnTitle = (SALTBLSETCOLUMNTITLE)GetProcAddress( m_hInst, "SalTblSetColumnTitle" );
#if defined(TDASCII)
		TblSetColumnTitleA = (SALTBLSETCOLUMNTITLEA)GetProcAddress( m_hInst, "SalTblSetColumnTitleA" );
#endif
		TblSetColumnWidth = (SALTBLSETCOLUMNWIDTH)GetProcAddress( m_hInst, "SalTblSetColumnWidth" );
		TblSetContext = (SALTBLSETCONTEXT)GetProcAddress( m_hInst, "SalTblSetContext" );
		TblSetFocusCell = (SALTBLSETFOCUSCELL)GetProcAddress( m_hInst, "SalTblSetFocusCell" );
		TblSetFocusRow = (SALTBLSETFOCUSROW)GetProcAddress( m_hInst, "SalTblSetFocusRow" );
		TblSetLinesPerRow = (SALTBLSETLINESPERROW)GetProcAddress( m_hInst, "SalTblSetLinesPerRow" );
		TblSetRowFlags = (SALTBLSETROWFLAGS)GetProcAddress( m_hInst, "SalTblSetRowFlags" );
		TblSetTableFlags = (SALTBLSETTABLEFLAGS)GetProcAddress( m_hInst, "SalTblSetTableFlags" );

		CvtDoubleToNumber = (SWINCVTDOUBLETONUMBER)GetProcAddress( m_hInst, "SWinCvtDoubleToNumber" );
		CvtIntToNumber = (SWINCVTINTTONUMBER)GetProcAddress( m_hInst, "SWinCvtIntToNumber" );
		CvtLongToNumber = (SWINCVTLONGTONUMBER)GetProcAddress( m_hInst, "SWinCvtLongToNumber" );
#ifdef _WIN64
		CvtLongLongToNumber = (SWINCVTLONGLONGTONUMBER)GetProcAddress(m_hInst, "SWinCvtLongLongToNumber");
#endif
		CvtWordToNumber = (SWINCVTWORDTONUMBER)GetProcAddress( m_hInst, "SWinCvtWordToNumber" );
		CvtNumberToDouble = (SWINCVTNUMBERTODOUBLE)GetProcAddress( m_hInst, "SWinCvtNumberToDouble" );
		CvtNumberToInt = (SWINCVTNUMBERTOINT)GetProcAddress( m_hInst, "SWinCvtNumberToInt" );
		CvtNumberToLong = (SWINCVTNUMBERTOLONG)GetProcAddress( m_hInst, "SWinCvtNumberToLong" );
#ifdef _WIN64
		CvtNumberToLongLong = (SWINCVTNUMBERTOLONGLONG)GetProcAddress(m_hInst, "SWinCvtNumberToLongLong");
#endif
		CvtNumberToULong = (SWINCVTNUMBERTOULONG)GetProcAddress( m_hInst, "SWinCvtNumberToULong" );
#ifdef _WIN64
		CvtNumberToULongLong = (SWINCVTNUMBERTOULONGLONG)GetProcAddress(m_hInst, "SWinCvtNumberToULongLong");
#endif
		CvtNumberToWord = (SWINCVTNUMBERTOWORD)GetProcAddress( m_hInst, "SWinCvtNumberToWord" );
		CvtULongToNumber = (SWINCVTULONGTONUMBER)GetProcAddress( m_hInst, "SWinCvtULongToNumber" );
		HStringUnlock = (SWINHSTRINGUNLOCK)GetProcAddress(m_hInst, "SWinHStringUnlock");
		InitLPHSTRINGParam = (SWININITLPHSTRINGPARAM)GetProcAddress( m_hInst, "SWinInitLPHSTRINGParam" );
		MDArrayGetHandle = (SWINMDARRAYGETHANDLE)GetProcAddress( m_hInst, "SWinMDArrayGetHandle" );
		MDArrayGetHString = (SWINMDARRAYGETHSTRING)GetProcAddress( m_hInst, "SWinMDArrayGetHString" );
		MDArrayGetNumber = (SWINMDARRAYGETNUMBER)GetProcAddress( m_hInst, "SWinMDArrayGetNumber" );
		MDArrayGetUdv = (SWINMDARRAYGETUDV)GetProcAddress( m_hInst, "SWinMDArrayGetUdv" );
		MDArrayPutHandle = (SWINMDARRAYPUTHANDLE)GetProcAddress( m_hInst, "SWinMDArrayPutHandle" );
		MDArrayPutHString = (SWINMDARRAYPUTHSTRING)GetProcAddress( m_hInst, "SWinMDArrayPutHString" );
		MDArrayPutNumber = (SWINMDARRAYPUTNUMBER)GetProcAddress( m_hInst, "SWinMDArrayPutNumber" );
		StringGetBuffer = (SWINSTRINGGETBUFFER)GetProcAddress( m_hInst, "SWinStringGetBuffer" );
#if defined(TDASCII)
		StringGetBufferA = (SWINSTRINGGETBUFFERA)GetProcAddress( m_hInst, "SWinStringGetBufferA" );
#endif
		UdvDeref = (SWINUDVDEREF)GetProcAddress( m_hInst, "SWinUdvDeref" );
		LogLine = (SALLOGLINE)GetProcAddress( m_hInst, "SalLogLine" );
#if defined(TDASCII)
		LogLineA = (SALLOGLINEA)GetProcAddress( m_hInst, "SalLogLineA" );
#endif
		ResId = (SALRESID)GetProcAddress( m_hInst, "SalResId" );
#if defined(TDASCII)
		ResIdA = (SALRESIDA)GetProcAddress( m_hInst, "SalResIdA" );
#endif
		ResLoad = (SALRESLOAD)GetProcAddress( m_hInst, "SalResLoad" );
		WaitCursor = (SALWAITCURSOR)GetProcAddress(m_hInst, "SalWaitCursor");

		SalPicGetImage = (SALPICGETIMAGE)GetProcAddress( m_hInst, "SalPicGetImage" );

		lstrcpy( m_szDLL, szDLL );
	}
	else
	{
		TCHAR szMsg[1024] = _T("");
		wsprintf( szMsg, _T("Gupta runtime environment could not be loaded ( %s )."), szDLL );
		MessageBox( NULL, szMsg, _T("MICSTO"), MB_ICONSTOP | MB_OK );
		return;
	}

	m_Initialized = ArrayDimCount &&
					ArrayGetLowerBound &&
					ArrayGetUpperBound &&
					ArrayIsEmpty &&
					ColorGet &&
					CompileAndEvaluate &&
#if defined(TDASCII)
					CompileAndEvaluateA &&
#endif
					ContextCurrent &&
					DateToStr &&
					DisableWindow &&
					DragDropStart &&
					DragDropStop &&
					EnableWindow &&
					FmtFieldToStr &&
					FmtGetPicture &&
					FmtGetInputMask &&
					FormUnitsToPixels &&
					GetDataType &&
					GetItemName &&
					GetMaxDataLength &&
					GetType &&
					GetVersion &&
					HStringUnRef &&
					InvalidateWindow &&
					IsWindowEnabled &&
					IsWindowVisible &&
					ListAdd &&
					ListClear &&
					ListQueryCount &&
					ListQuerySelection &&
					ListQueryText &&
					ListSetSelect &&
					NumberToHString &&
					ParentWindow &&
					PicGetString &&
					PixelsToFormUnits &&
					QueryFieldEdit &&
					SendValidateMsg &&
					StatusSetText &&
#if defined(TDASCII)
					StatusSetTextA &&
#endif
					SetWindowLoc &&
					SetWindowSize &&
					StrIsValidNumber &&
#if defined(TDASCII)
					StrIsValidNumberA &&
#endif
					StrToDate &&
#if defined(TDASCII)
					StrToDateA &&
#endif
					StrToNumber &&
#if defined(TDASCII)
					StrToNumberA &&
#endif
					UpdateWindow &&
					WindowGetProperty &&
#if defined(TDASCII)
					WindowGetPropertyA &&
#endif
					ValidateSet &&

					TblAnyRows &&
					TblClearSelection &&
					TblDefineCheckBoxColumn &&
					TblDefineDropDownListColumn &&
					TblDefinePopupEditColumn &&
					TblDefineRowHeader &&
					TblDefineSplitWindow &&
					TblDeleteRow &&
					TblFetchRow &&
					TblFindNextRow &&
					TblFindPrevRow &&
					TblGetColumnText &&
					TblGetColumnTitle &&
					TblGetColumnWindow &&
					TblInsertRow &&
					TblKillEdit &&
					TblKillFocus &&
					TblObjectsFromPoint &&
					TblQueryCheckBoxColumn &&
					TblQueryColumnCellType &&
					TblQueryColumnFlags &&
					TblQueryColumnID &&
					TblQueryColumnPos &&
					TblQueryColumnWidth &&
					TblQueryContext &&
					TblQueryDropDownListColumn &&
					TblQueryFocus &&
					TblQueryLinesPerRow &&
					TblQueryLockedColumns &&
					TblQueryPopupEditColumn &&
					TblQueryRowFlags &&
					TblQueryRowHeader &&
					TblQueryScroll &&
					TblQuerySplitWindow &&
					TblQueryTableFlags &&
					TblQueryVisibleRange &&
					TblScroll &&
					TblSetColumnFlags &&
					TblSetColumnText &&
#if defined(TDASCII)
					TblSetColumnTextA &&
#endif
					TblSetColumnTitle &&
#if defined(TDASCII)
					TblSetColumnTitleA &&
#endif
					TblSetColumnWidth &&
					TblSetContext &&
					TblSetFocusCell &&
					TblSetFocusRow &&
					TblSetLinesPerRow &&
					TblSetRowFlags &&
					TblSetTableFlags &&

					CvtDoubleToNumber &&
					CvtIntToNumber &&
					CvtLongToNumber &&
#ifdef _WIN64
					CvtLongLongToNumber &&
#endif
					CvtWordToNumber &&
					CvtNumberToDouble &&
					CvtNumberToInt &&
					CvtNumberToLong &&
#ifdef _WIN64
					CvtNumberToLongLong &&
#endif
					CvtNumberToULong &&
#ifdef _WIN64
					CvtNumberToULongLong &&
#endif
					CvtNumberToWord &&
					CvtULongToNumber &&
#ifdef CTD7x
					HStringUnlock &&
#endif
					InitLPHSTRINGParam &&
					MDArrayGetHandle &&
					MDArrayGetHString &&
					MDArrayGetNumber &&
					MDArrayGetUdv &&
					MDArrayPutHandle &&
					MDArrayPutHString &&
					MDArrayPutNumber &&
					StringGetBuffer &&
#if defined(TDASCII)
					StringGetBufferA &&
#endif
					UdvDeref &&

					LogLine &&
#if defined(TDASCII)
					LogLineA &&
#endif
					ResId &&
#if defined(TDASCII)
					ResIdA &&
#endif
					ResLoad &&
					WaitCursor;

	if ( !m_Initialized )
		MessageBox( NULL, _T("Not all required Gupta functions could be loaded."), _T("MICSTO"), MB_ICONSTOP | MB_OK );
}


// Destruktor
CCtd::~CCtd( void )
{
	if ( m_hInst )
	{
		FreeLibrary( m_hInst );
		m_hInst = NULL;
	}
}

// HStrToNum
NUMBER CCtd::HStrToNum( HSTRING hs )
{
	long lLen;
	NUMBER nRet = {0};
	LPTSTR lps = StringGetBuffer( hs, &lLen );
	if ( lps )
	{
		nRet = StrToNumber( lps );
	}

	return nRet;
}

// HStrToTStr
BOOL CCtd::HStrToTStr(HSTRING hs, tstring &ts, BOOL bUnref /*=FALSE*/)
{
	BOOL bOk = FALSE;

	ts.clear();
	if (hs)
	{
		long lLen;
		LPTSTR lps = StringGetBuffer(hs, &lLen);
		if (lps)
		{
			ts = lps;
			bOk = TRUE;
		}
		HStrUnlock(hs);

		if (bUnref)
			HStringUnRef(hs);
	}

	return bOk;
}

// HStrUnlock
void CCtd::HStrUnlock(HSTRING hs)
{
	if (HStringUnlock)
		HStringUnlock(hs);
}

// NumGetNull
NUMBER CCtd::NumGetNull( )
{
	NUMBER n;
	memset( &n, 0, sizeof( NUMBER) );
	return n;
}

// NumIsEqual
BOOL CCtd::NumIsEqual( NUMBER &n1, NUMBER &n2 )
{
	if ( n1.numLength == 0 && n2.numLength == 0 )
		return TRUE;
	else if ( n1.numLength != n2.numLength )
		return FALSE;
	else
	{
		DOUBLE d1 = 0.0, d2 = 0.0;
		CvtNumberToDouble( &n1, &d1 );
		CvtNumberToDouble( &n2, &d2 );

		return d1 == d2;
	}
}

// NumToBool
BOOL CCtd::NumToBool( NUMBER &n )
{
	long l = NumToLong( n );
	return l ? TRUE : FALSE;
}

// NumToDouble
double CCtd::NumToDouble( NUMBER &n )
{	
	double dRet = 0.0;

	if ( n.numLength != 0 )
		CvtNumberToDouble( &n, &dRet );

	return dRet;
}

// NumToInt
int CCtd::NumToInt( NUMBER &n )
{	
	int iRet = 0;

	if ( n.numLength != 0 )
		CvtNumberToInt( &n, &iRet );

	return iRet;
}

// NumToLong
long CCtd::NumToLong( NUMBER &n )
{	
	long lRet = 0;

	if ( n.numLength != 0 )
		CvtNumberToLong( &n, &lRet );

	return lRet;
}

#ifdef _WIN64
// NumToLongLong
LONGLONG CCtd::NumToLongLong(NUMBER &n)
{
	LONGLONG lRet = 0;

	if (n.numLength != 0)
		CvtNumberToLongLong(&n, &lRet);

	return lRet;
}
#endif

// NumToULong
ULONG CCtd::NumToULong( NUMBER &n )
{	
	ULONG ulRet = 0;

	if ( n.numLength != 0 )
		CvtNumberToULong( &n, &ulRet );

	return ulRet;
}

#ifdef _WIN64
// NumToULongLong
ULONGLONG CCtd::NumToULongLong(NUMBER &n)
{
	ULONGLONG ulRet = 0;

	if (n.numLength != 0)
		CvtNumberToULongLong(&n, &ulRet);

	return ulRet;
}
#endif

// PicGetImage
long CCtd::PicGetImage( HWND hWndPic, LPHSTRING phs, LPINT piType )
{
	long lLen = 0;

	if ( SalPicGetImage )
		lLen =  SalPicGetImage( hWndPic, phs, piType );
	else if ( PicGetString )
	{
		lLen = PicGetString( hWndPic, PIC_FormatBitmap, phs );
		if ( lLen )
			*piType = PIC_ImageTypeBMP;
	}

	return lLen;
}

// TblGetCellType
int CCtd::TblGetCellType( HWND hWndCol )
{
	int iCellType = 0;
	TblQueryColumnCellType( hWndCol, &iCellType );
	return iCellType;
}

// TblGetFocusRow
long CCtd::TblGetFocusRow( HWND hWndTbl )
{
	HWND hWndCol;
	long lRow;

	if ( !TblQueryFocus( hWndTbl, &lRow, &hWndCol ) )
		return TBL_Error;

	return lRow;
}

// TblIsRow
BOOL CCtd::TblIsRow( HWND hWndTbl, long lRow )
{
	if ( lRow > TBL_MaxRow || lRow < TBL_MinSplitRow )
		 return FALSE;

	if ( TblQueryRowFlags( hWndTbl, lRow, ROW_InCache ) )
		return TRUE;
	else
	{
		long lContext = TblQueryContext( hWndTbl );
		BOOL bOk = TblSetContextOnly( hWndTbl, lRow );
		TblSetContextOnly( hWndTbl, lContext );
		return bOk;
	}
}

// TblSetContextEx
BOOL CCtd::TblSetContextEx( HWND hWndTbl, long lRow )
{
	if ( !TblSetContextOnly( hWndTbl, lRow ) )
		return FALSE;
	if ( lRow >= 0 && !TblQueryRowFlags( hWndTbl, lRow, ROW_InCache ) )
		SendMessage( hWndTbl, TIM_FetchRow, 0, lRow );
	return TRUE;
}

// TblSetFlagsRowRange
long CCtd::TblSetFlagsRowRange( HWND hWndTbl, long lRowFrom, long lRowTo, WORD wFlagsOn, WORD wFlagsOff, WORD wFlagsSet, BOOL bSet )
{
	long lSet = 0;
	if ( lRowFrom != TBL_Error && lRowTo != TBL_Error )
	{
		long lInc = lRowTo >= lRowFrom ? 1 : -1;
		long lMax = lRowTo + lInc;		

		for( long l = lRowFrom; l != lMax; l += lInc )
		{
			if ( TblIsRow( hWndTbl, l ) )
			{
				if ( wFlagsOn && !TblQueryRowFlags( hWndTbl, l, wFlagsOn ) )
					continue;

				if ( wFlagsOff && TblQueryRowFlags( hWndTbl, l, wFlagsOff ) )
					continue;

				TblSetRowFlags( hWndTbl, l, wFlagsSet, bSet );
				lSet++;
			}
		}
	}

	return lSet;
}