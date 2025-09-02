#if !defined(AFX_TBLEXPDLG_H__0124ECE1_E96A_11D5_A9AB_00104B88A5A3__INCLUDED_)
#define AFX_TBLEXPDLG_H__0124ECE1_E96A_11D5_A9AB_00104B88A5A3__INCLUDED_

// Includes
#include "resource.h"
#include "lang.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// TblExpDlg.h : Header-Datei
//

/////////////////////////////////////////////////////////////////////////////
// Dialogfeld CTblExpDlg 

class CTblExpDlg : public CDialog
{
// Konstruktion
public:
	void SetLanguage( int iLanguage );
	int m_Language;
	HWND m_StatusWnd;
	BOOL m_Cancel;
	CTblExpDlg(CWnd* pParent = NULL);   // Standardkonstruktor

// Dialogfelddaten
	//{{AFX_DATA(CTblExpDlg)
	enum { IDD = IDD_TBLEXP };
	CString	m_Status;
	//}}AFX_DATA


// Überschreibungen
	// Vom Klassen-Assistenten generierte virtuelle Funktionsüberschreibungen
	//{{AFX_VIRTUAL(CTblExpDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV-Unterstützung
	//}}AFX_VIRTUAL

// Implementierung
protected:

	// Generierte Nachrichtenzuordnungsfunktionen
	//{{AFX_MSG(CTblExpDlg)
	virtual void OnCancel();
	afx_msg LRESULT OnUser( WPARAM, LPARAM );
	virtual BOOL OnInitDialog();
	afx_msg void OnCancelMode();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedCancel();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ fügt unmittelbar vor der vorhergehenden Zeile zusätzliche Deklarationen ein.

#endif // AFX_TBLEXPDLG_H__0124ECE1_E96A_11D5_A9AB_00104B88A5A3__INCLUDED_
