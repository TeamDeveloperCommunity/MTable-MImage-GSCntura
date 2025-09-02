// cmtblcounter.h

#ifndef _INC_CMTBLCOUNTER
#define _INC_CMTBLCOUNTER

// Includes
#include <windows.h>

// Klasse CMTblCounter

class CMTblCounter
{
public:
	// Konstruktoren
	CMTblCounter( );
	// Destruktor
	~CMTblCounter( );
	// Funktionen
	long DecMergeCells( long lVal = 1 ) { return m_lMergeCells -= lVal; };
	long DecMergeRows( long lVal = 1 ) { return m_lMergeRows -= lVal; };
	long DecNoFocusUpdate( long lVal = 1 ) { return m_lNoFocusUpdate -= lVal; };
	long DecNoSelInv( long lVal = 1 ) { return m_lNoSelInv -= lVal; };
	long DecRowHeights( long lVal = 1 ) { return m_lRowHeights -= lVal; };
	long GetMergeCells( ) { return m_lMergeCells; };
	long GetMergeRows( ) { return m_lMergeRows; };
	long GetNoFocusUpdate( ) { return m_lNoFocusUpdate; };
	long GetNoSelInv( ) { return m_lNoSelInv; };
	long GetRowHeights( ) { return m_lRowHeights; };
	long IncMergeCells( long lVal = 1 ) { return m_lMergeCells += lVal; };
	long IncMergeRows( long lVal = 1 ) { return m_lMergeRows += lVal; };
	long IncNoFocusUpdate( long lVal = 1 ) { return m_lNoFocusUpdate += lVal; };
	long IncNoSelInv( long lVal = 1 ) { return m_lNoSelInv += lVal; };
	long IncRowHeights( long lVal = 1 ) { return m_lRowHeights += lVal; };
	void InitMergeCells( ) { m_lMergeCells = 0; };
	void InitMergeRows( ) { m_lMergeRows = 0; };
	void InitNoFocusUpdate( ) { m_lNoFocusUpdate = 0; };
	void InitNoSelInv( ) { m_lNoSelInv = 0; };
	void InitRowHeights( ) { m_lRowHeights = 0; };
private:
	long m_lMergeCells;
	long m_lMergeRows;
	long m_lNoSelInv;
	long m_lNoFocusUpdate;
	long m_lRowHeights;
};

#endif // _INC_CMTBLCOUNTER