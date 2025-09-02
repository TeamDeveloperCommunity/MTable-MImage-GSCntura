#ifndef _INC_OPENOFFICE
#define _INC_OPENOFFICE

// Includes
#include "ooconst.h"
#include "export.h"
#include "mtexp.h"

// Klassen

// OOBorderLine
class OOBorderLine : public COleDispatchDriver
{
public:
	OOBorderLine() {}
	OOBorderLine(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL SetColor( long lValue );	
	BOOL SetInnerLineWidth( short shValue );
	BOOL SetLineDistance( short shValue );
	BOOL SetOuterLineWidth( short shValue );
};


// OOCell
class OOCell : public COleDispatchDriver
{
public:
	OOCell() {}
	OOCell(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL SetNumberFormat( long lFmtIndex );
	BOOL SetValue( LPCTSTR lpsValue, BOOL bFormula = FALSE );
	BOOL SetValue( double dValue );
};

// OOCellRange
class OOCellRange : public COleDispatchDriver
{
public:
	OOCellRange() {}
	OOCellRange(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL ClearContents( long lFlags );
	LPDISPATCH GetBorderLine( int iType );
	LPDISPATCH GetCellByPosition( long lCol, long lRow );
	short OOCellRange::GetParaIndent( );
	LPDISPATCH GetRangeAddress( );
	LPDISPATCH GetSpreadsheet( );
	LPDISPATCH GetStartCell( );	
	LPDISPATCH GetTableBorder( );
	BOOL Indent( long lSteps );
	BOOL Merge( BOOL bMerge );
	BOOL SetBackColor( long lValue );
	BOOL SetBorderLine( int iType, LPDISPATCH lpDisp );
	BOOL SetFontName( LPCTSTR lpsName );
	BOOL SetFontSize( float fSize );
	BOOL SetFontStyle( long lValue );
	BOOL SetFontWeight( float fWeight );
	BOOL SetHoriJustify( long lValue );
	BOOL SetIndentLevel( long lValue );
	BOOL SetNumberFormat( long lFmtIndex );
	BOOL SetParaIndent( short shValue );
	BOOL SetStrikeout( short shValue );
	BOOL SetTableBorder( LPDISPATCH lpDisp);
	BOOL SetTextColor( long lValue );
	BOOL SetTextWrap( BOOL bValue );
	BOOL SetUnderline( short shValue );
	BOOL SetValue( LPCTSTR lpsValue, BOOL bFormula = FALSE );
	BOOL SetValue( double dValue );
	BOOL SetVertJustify( long lValue );
};

// OOCellRangeAddress
class OOCellRangeAddress : public COleDispatchDriver
{
public:
	OOCellRangeAddress() {}
	OOCellRangeAddress(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	
	long EndColumn( );
	long EndRow( );
	long StartColumn( );
	long StartRow( );
};

// OOComponents
class OOComponents : public COleDispatchDriver
{
public:
	OOComponents() {}
	OOComponents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH CreateEnumeration( );
};


// OOConfirmer
class OOConfirmer
{
public:
	OOConfirmer( );

	static BOOL CALLBACK ConfirmProc( HWND hWnd, LPARAM lParam );
	static long Confirm( );
	BOOL IsStopped( );
	void SetStop( BOOL bSet );
	static UINT Start( LPVOID lpParam );
	void Stop( );
protected:
	BOOL m_bStop;
	CRITICAL_SECTION m_csStop;
};


// OODesktop
class OODesktop : public COleDispatchDriver
{
public:
	OODesktop() {}
	OODesktop(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetActiveFrame( );
	LPDISPATCH GetComponents( );
	LPDISPATCH GetCurrentFrame( );
	LPDISPATCH LoadComponentFromUrl( LPCTSTR lpsUrl, LPCTSTR lpsFrameName, long lSearchFlags );
};

// OODispatchHelper
class OODispatchHelper : public COleDispatchDriver
{
public:
	OODispatchHelper() {}
	OODispatchHelper(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	VARIANT ExecuteDispatch( LPDISPATCH lpDispProvider, LPCTSTR lpsUrl, LPCTSTR lpsFrame, long lSearchFlags, COleSafeArray &PropArray );
};


// OOEnumeration
class OOEnumeration : public COleDispatchDriver
{
public:
	OOEnumeration() {}
	OOEnumeration(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL HasMoreElements( );
	VARIANT NextElement( );
};


// OOFrame
class OOFrame : public COleDispatchDriver
{
public:
	OOFrame() {}
	OOFrame(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	CString GetName( );
};


// OOIdlClass
class OOIdlClass : public COleDispatchDriver
{
public:
	OOIdlClass() {}
	OOIdlClass(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH CreateObject( );
};


// OOIndexAccess
class OOIndexAccess : public COleDispatchDriver
{
public:
	OOIndexAccess() {}
	OOIndexAccess(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	VARIANT GetByIndex( long lIndex );
};


// OOModel
class OOModel : public COleDispatchDriver
{
public:
	OOModel() {}
	OOModel(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetCurrentController( );
};


// OONumberFormatter
class OONumberFormatter : public COleDispatchDriver
{
public:
	OONumberFormatter() {}
	OONumberFormatter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL AttachNumberFormatsSupplier( LPDISPATCH lpDisp );
	long DetectNumberFormat( long lKey, LPCTSTR lpsValue );
};


// OOPropertyValue
class OOPropertyValue : public COleDispatchDriver
{
public:
	OOPropertyValue() {}
	OOPropertyValue(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL SetName( LPCTSTR lpsName );
	BOOL SetValue( long Value );
	BOOL SetValue( LPCTSTR Value );
	
};


// OOReflection
class OOReflection : public COleDispatchDriver
{
public:
	OOReflection() {}
	OOReflection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH ForName( LPCTSTR lpsName );
};


// OOServiceManager
class OOServiceManager : public COleDispatchDriver
{
public:
	OOServiceManager() {}
	OOServiceManager(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH CreateInstance( LPCTSTR lpsName );
};


// OOSpreadsheet
class OOSpreadsheet : public COleDispatchDriver
{
public:
	OOSpreadsheet() {}
	OOSpreadsheet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetCellRangeByName( LPCTSTR lpsName );
	LPDISPATCH GetCellByPosition( long lCol, long lRow );
	LPDISPATCH GetCellRangeByPosition( long lColLeft, long lRowTop, long lColRight, long lRowBottom );
	LPDISPATCH GetColumns( );
	LPDISPATCH GetRows( );
};


// OOSpreadsheetDocument
class OOSpreadsheetDocument : public COleDispatchDriver
{
public:
	OOSpreadsheetDocument() {}
	OOSpreadsheetDocument(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetSheets( );
	static BOOL IsSpreadsheetDocument( LPDISPATCH lpDisp );
};


// OOSpreadsheetView
class OOSpreadsheetView : public COleDispatchDriver
{
public:
	OOSpreadsheetView() {}
	OOSpreadsheetView(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetActiveSheet( );
	LPDISPATCH GetFrame( );
	LPDISPATCH GetSelection( );
	BOOL InsertTransferable( LPDISPATCH pDisp );
	BOOL SetActiveSheet( LPDISPATCH pDisp );
	BOOL Select( LPDISPATCH pDisp );
};


// OOSpreadsheets
class OOSpreadsheets : public COleDispatchDriver
{
public:
	OOSpreadsheets() {}
	OOSpreadsheets(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	long Add( int iLanguage );
	LPDISPATCH GetByIndex( long lIndex );
	long GetCount( );
	VARIANT GetElementNames( );
	BOOL InsertNewByName( LPCTSTR lpsName, short nPos );
};


// OOSystemClipboard
class OOSystemClipboard : public COleDispatchDriver
{
public:
	OOSystemClipboard() {}
	OOSystemClipboard(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetContents( );
};


// OOTableBorder
class OOTableBorder : public COleDispatchDriver
{
public:
	OOTableBorder() {}
	OOTableBorder(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	LPDISPATCH GetBorderLine( int iType );
	BOOL SetBorderLine( int iType, LPDISPATCH lpDisp );
	BOOL SetLineValid( int iType, BOOL bValid );
};


// OOTableColumn
class OOTableColumn : public COleDispatchDriver
{
public:
	OOTableColumn() {}
	OOTableColumn(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL SetOptimalWidth( BOOL bSet );
};


// OOTableRow
class OOTableRow : public COleDispatchDriver
{
public:
	OOTableRow() {}
	OOTableRow(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}

	BOOL SetOptimalHeight( BOOL bSet );
};


// OOTransferable
class OOTransferable : public COleDispatchDriver
{
public:
	OOTransferable() {}
	OOTransferable(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
};

// Funktionen
void AttachRange( OOSpreadsheet & Sheet, OOCellRange & R, int iColFrom, int iRowFrom, int iColTo = -1, int iRowTo = -1 );
long MakeColor( long lColor );
void SetFont( OOCellRange & R, FONTINFO & fi );


#endif // _INC_OPENOFFICE
