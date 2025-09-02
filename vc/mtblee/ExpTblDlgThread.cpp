// ExpTblDlgThread.cpp: Implementierungsdatei
//

#include "stdafx.h"
#include "resource.h"
#include "ExpTblDlgThread.h"
#include "mtexcexpapi.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CExpTblDlgThread

IMPLEMENT_DYNCREATE(CExpTblDlgThread, CWinThread)

CExpTblDlgThread::CExpTblDlgThread()
{
	m_Language = MTE_LNG_ENGLISH;
}

CExpTblDlgThread::~CExpTblDlgThread()
{
}

BOOL CExpTblDlgThread::InitInstance()
{
	m_Dlg.SetLanguage( m_Language );
	m_Dlg.Create( IDD_TBLEXP );
	m_Dlg.SetWindowPos( &CWnd::wndTopMost, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE );
	m_pMainWnd = (CWnd*)&m_Dlg;
	return TRUE;
}

int CExpTblDlgThread::ExitInstance()
{
	// ZU ERLEDIGEN:  Bereinigungen für jeden Thread hier durchführen
	return CWinThread::ExitInstance();
}

BEGIN_MESSAGE_MAP(CExpTblDlgThread, CWinThread)
	//{{AFX_MSG_MAP(CExpTblDlgThread)
		// HINWEIS - Der Klassen-Assistent fügt hier Zuordnungsmakros ein und entfernt diese.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Behandlungsroutinen für Nachrichten CExpTblDlgThread 

BOOL CExpTblDlgThread::IsCanceled()
{
	return m_Dlg.m_Cancel;
}

HWND CExpTblDlgThread::GetDlgHandle()
{
	return m_Dlg.m_hWnd;
}

void CExpTblDlgThread::MakeDlgTopmost( BOOL bTopmost )
{
	m_Dlg.SetWindowPos( bTopmost ? &CWnd::wndTopMost : &CWnd::wndNoTopMost, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE );
}

void CExpTblDlgThread::SetLanguage(int iLanguage)
{
	m_Language = iLanguage;
}
