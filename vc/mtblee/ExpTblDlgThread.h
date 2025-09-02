#if !defined(AFX_EXPTBLDLGTHREAD_H__0124ECE2_E96A_11D5_A9AB_00104B88A5A3__INCLUDED_)
#define AFX_EXPTBLDLGTHREAD_H__0124ECE2_E96A_11D5_A9AB_00104B88A5A3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ExpTblDlgThread.h : Header-Datei
//

#include "TblExpDlg.h"

/////////////////////////////////////////////////////////////////////////////
// Thread CExpTblDlgThread 

class CExpTblDlgThread : public CWinThread
{
	DECLARE_DYNCREATE(CExpTblDlgThread)
protected:
	CExpTblDlgThread();           // Dynamische Erstellung verwendet gesch�tzten Konstruktor

// Attribute
public:
	CTblExpDlg m_Dlg;
	int m_Language;
// Operationen
public:	
	HWND GetDlgHandle( );
	BOOL IsCanceled( );
	void MakeDlgTopmost( BOOL bTopmost );
	void SetLanguage( int iLanguage );
	

// �berschreibungen
	// Vom Klassen-Assistenten generierte virtuelle Funktions�berschreibungen
	//{{AFX_VIRTUAL(CExpTblDlgThread)
	public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementierung
protected:
	virtual ~CExpTblDlgThread();

	// Generierte Nachrichtenzuordnungsfunktionen
	//{{AFX_MSG(CExpTblDlgThread)
		// HINWEIS - Der Klassen-Assistent f�gt hier Member-Funktionen ein und entfernt diese.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ f�gt unmittelbar vor der vorhergehenden Zeile zus�tzliche Deklarationen ein.

#endif // AFX_EXPTBLDLGTHREAD_H__0124ECE2_E96A_11D5_A9AB_00104B88A5A3__INCLUDED_
