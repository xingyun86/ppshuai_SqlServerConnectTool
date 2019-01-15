// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "aboutdlg.h"
#include "MainDlg.h"

BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

BOOL CMainDlg::OnIdle()
{
	UIUpdateChildWindows();
	return FALSE;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// center the dialog on the screen
	CenterWindow();

	// set icons
	HICON hIcon = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR, ::GetSystemMetrics(SM_CXICON), ::GetSystemMetrics(SM_CYICON));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	UIAddChildWindowContainer(m_hWnd);

	return TRUE;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);
	pLoop->RemoveIdleHandler(this);

	// if UI is the last thread, no need to wait
	if(_Module.GetLockCount() == 1)
	{
		_Module.m_dwTimeOut = 0L;
		_Module.m_dwPause = 0L;
	}
	_Module.Unlock();

	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CAboutDlg dlg;
	dlg.DoModal();
	return 0;
}

LRESULT CMainDlg::OnOK(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	// TODO: Add validation code 
	CAdoSql as;
	if (SUCCEEDED(as.m_hResult))
	{
		try
		{
			//"Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=sa;Initial Catalog=QPServerInfoDB;Data Source=192.168.1.7,1433;";
			/*"Provider=SQLOLEDB.1;Password=123;Persist Security Info=False;User ID=sa;Initial Catalog=QPServerInfoDB;Data Source=PC201602252148";*/
			/*"Driver={sql server};server=PC201602252148;uid=sa;pwd=123"*/
			/*"Driver={sql server};server=192.168.0.7,1433;uid=sa;pwd=123"*/
			
			//as.m_hResult = as.InitConn(_T("driver={sql server};server=192.168.100.4;uid=sa;pwd=Admin_1880"));
			//as.m_hResult = as.InitConn(_T("Driver={sql server};server=192.168.100.4;uid=sa;pwd=Admin_1880"));
			as.m_hResult = as.InitConn(_T("Provider=SQLOLEDB.1;Password=Admin_1880;Persist Security Info=True;User ID=sa;Initial Catalog=master;Data Source=192.168.100.4\\.,1433;"));
			//as.m_hResult = as.InitConn(_T("Provider=SQLOLEDB.1;Password=Admin_1880;Persist Security Info=False;User ID=sa;Initial Catalog=master;Data Source=192.168.100.4"));
			as.m_hResult = as.ExitConn();
			OutputDebugString(_T("连接测试成功"));
		}
		catch (_com_error e)
		{
			OutputDebugString(e.ErrorMessage());
		}
	}
	//CloseDialog(wID);
	return 0;
}

LRESULT CMainDlg::OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CloseDialog(wID);
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	::PostQuitMessage(nVal);
}
