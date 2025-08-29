// TblExpDlg.cpp: Implementierungsdatei
//

#include "stdafx.h"
#include "TblExpDlg.h"
#include "mtexcexpapi.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Dialogfeld CTblExpDlg 


CTblExpDlg::CTblExpDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTblExpDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTblExpDlg)
	m_Status = _T("");
	//}}AFX_DATA_INIT

	m_Cancel = FALSE;
	m_StatusWnd = NULL;
	m_Language = MTE_LNG_ENGLISH;
}


void CTblExpDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTblExpDlg)
	DDX_Text(pDX, IDC_STATUS, m_Status);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CTblExpDlg, CDialog)
	//{{AFX_MSG_MAP(CTblExpDlg)
	ON_MESSAGE( WM_USER, OnUser )
	ON_WM_CANCELMODE()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDCANCEL, &CTblExpDlg::OnBnClickedCancel)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Behandlungsroutinen für Nachrichten CTblExpDlg 


void CTblExpDlg::OnCancel() 
{
	m_Cancel = TRUE;
}

afx_msg LRESULT CTblExpDlg::OnUser( WPARAM wParam, LPARAM lParam )
{
	m_Status = LPTSTR(lParam);
	::SetWindowText( m_StatusWnd, m_Status );

	return 0;
}


BOOL CTblExpDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	/*if ( m_Language == MTE_LNG_ENGLISH )
	{
		HWND hWnd = NULL;
		GetDlgItem( IDCANCEL, &hWnd );
		::SetWindowText( hWnd, _T("Cancel") ); 
	}*/

	CString sUIStr;

	sUIStr = _T("");
	GetUIStr( m_Language, _T("StatDlg.Title"), sUIStr );
	::SetWindowText( m_hWnd, sUIStr );

	sUIStr = _T("");
	GetUIStr( m_Language, _T("StatDlg.BtnCancel"), sUIStr );
	HWND hWnd = NULL;
	GetDlgItem( IDCANCEL, &hWnd );
	::SetWindowText( hWnd, sUIStr ); 
	
	GetDlgItem( IDC_STATUS, &m_StatusWnd );

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX-Eigenschaftenseiten sollten FALSE zurückgeben
}

void CTblExpDlg::OnCancelMode() 
{
	CDialog::OnCancelMode();
	
	// TODO: Code für die Behandlungsroutine für Nachrichten hier einfügen
	
}

void CTblExpDlg::SetLanguage(int iLanguage)
{
	m_Language = iLanguage;
}


void CTblExpDlg::OnBnClickedCancel()
{
	// TODO: Fügen Sie hier Ihren Kontrollbehandlungscode für die Benachrichtigung ein.
	CDialog::OnCancel();
}
