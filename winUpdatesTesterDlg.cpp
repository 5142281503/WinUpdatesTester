#include "stdafx.h"
#include "winUpdatesTester.h"
#include "winUpdatesTesterDlg.h"
#include <Wuapi.h>
#include <wuerror.h>
#include <sstream>

namespace
{
	CString GetSys32Folder()
	{
		wchar_t buf[MAX_PATH+2] ={ 0 };
		GetSystemDirectory(buf, MAX_PATH);
		return buf;
	}

	static const CString g_Sys32Folder( GetSys32Folder() + L"\\" );

	bool RunningOn64BitWindows()
	{
	#if defined(_WIN64)
		return true;
	#else
		BOOL bRunningAsWow64Process = FALSE;
		IsWow64Process( GetCurrentProcess(), &bRunningAsWow64Process ) ;
		return bRunningAsWow64Process;

	#endif
	}
	
}

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CwinUpdatesTesterDlg dialog




CwinUpdatesTesterDlg::CwinUpdatesTesterDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CwinUpdatesTesterDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CwinUpdatesTesterDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STATIC_MSG, m_Log);
}

BEGIN_MESSAGE_MAP(CwinUpdatesTesterDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CwinUpdatesTesterDlg::OnBnClickedGetEnabled)
	ON_BN_CLICKED(IDOK, &CwinUpdatesTesterDlg::OnBnClickedOkExit)
	ON_BN_CLICKED(IDC_BUTTON2, &CwinUpdatesTesterDlg::OnBnClickedGetSettings)
	ON_BN_CLICKED(IDC_BUTTON3, &CwinUpdatesTesterDlg::OnBnClickedSetSettings)
	ON_BN_CLICKED(IDC_BUTTON4, &CwinUpdatesTesterDlg::OnBnClickedCheckNow)
	ON_BN_CLICKED(IDC_BUTTON5, &CwinUpdatesTesterDlg::OnBnClickedGetLastUpdateTime)
	ON_BN_CLICKED(IDC_BUTTON6, &CwinUpdatesTesterDlg::OnBnClickedLastTimeUpdatesInstlled)
	ON_BN_CLICKED(IDC_BUTTON7, &CwinUpdatesTesterDlg::OnBnClickedDisplayWUUI)
	ON_BN_CLICKED(IDC_BUTTON8, &CwinUpdatesTesterDlg::OnBnClickedRebootPC)
	ON_BN_CLICKED(IDC_BUTTON9, &CwinUpdatesTesterDlg::OnBnClickedGetCurrentRebootState)
	ON_BN_CLICKED(IDC_BUTTON10, &CwinUpdatesTesterDlg::OnBnClickedSetEnabled)
	ON_BN_CLICKED(IDC_BUTTON11, &CwinUpdatesTesterDlg::OnBnClickedApplyNow)
END_MESSAGE_MAP()


// CwinUpdatesTesterDlg message handlers

BOOL CwinUpdatesTesterDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// init COM
	const HRESULT hResult = ::CoInitializeEx(	NULL, COINIT_APARTMENTTHREADED );	
	if ( hResult != S_OK && hResult != S_FALSE )
	{
		AfxMessageBox(L"COM ERROR !", MB_ICONSTOP);
	}

	// some info...
	if(!IsUserAnAdmin()){
		AfxMessageBox(L"You are not admin, some APIs may fail...", MB_ICONINFORMATION);
	}

	// set a radio button checked, doesn't matter which one
	((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(BST_CHECKED);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CwinUpdatesTesterDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CwinUpdatesTesterDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CwinUpdatesTesterDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CwinUpdatesTesterDlg::Log(const std::wstring& in_msg)
{
	CString s;
	m_Log.GetWindowTextW(s);
	m_Log.SetWindowTextW(s+in_msg.c_str()+L"\r\n");
	m_Log.LineScroll(INT_MAX-5, 0);
}

void CwinUpdatesTesterDlg::Log(const std::wstring& in_msg, HRESULT in_h)
{
	std::wstring m(in_msg);
	m += L":  => ";

	switch(in_h)
	{
	case S_OK:
		Log(m + L"OK");
		break;
	case E_ACCESSDENIED:
		Log(m + L"E_ACCESSDENIED");
		break;
	case WU_E_INSTALL_NOT_ALLOWED:
		Log(m + L"WU_E_INSTALL_NOT_ALLOWED");
		break;
	case WU_E_NO_UPDATE:
		Log(m + L"WU_E_NO_UPDATE");
		break;
	case WU_E_NOT_INITIALIZED :
		Log(m + L"WU_E_NOT_INITIALIZED");
		break;
	case WU_E_AU_NOSERVICE:
		Log(m + L"WU_E_AU_NOSERVICE");
		break;
	case WU_E_AU_PAUSED:
		Log(m + L"WU_E_AU_PAUSED");
		break;
	case WU_E_LEGACYSERVER:
		Log(m + L"WU_E_LEGACYSERVER");
		break;
	default:
		std::wostringstream str;
		str << std::hex << in_h;	
		Log(m + str.str());		
	}
}

void CwinUpdatesTesterDlg::OnBnClickedGetEnabled()
{
	CComPtr<IAutomaticUpdates> pAutomaticUpdates(NULL);
	const HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates, (void**) &pAutomaticUpdates);		
	if (FAILED(hr) || pAutomaticUpdates == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 
	VARIANT_BOOL vb = VARIANT_FALSE;
	if (S_OK != pAutomaticUpdates->get_ServiceEnabled(&vb))
	{
		Log(L"get_ServiceEnabled() failed !");
		return;
	}
	if (VARIANT_FALSE == vb)
	{
		Log(L"all the components that Automatic Updates requires are NOT available !");
		return;
	}
	Log(L"all the components that Automatic Updates requires are available");	

}

void CwinUpdatesTesterDlg::OnBnClickedOkExit()
{	
	OnOK();
}

void CwinUpdatesTesterDlg::OnBnClickedGetSettings()
{
	CComPtr<IAutomaticUpdates> pAutomaticUpdates(NULL);
	const HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates, (void**) &pAutomaticUpdates);		
	if (FAILED(hr) || pAutomaticUpdates == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 

	IAutomaticUpdatesSettings* s = NULL;
	if ((S_OK != pAutomaticUpdates->get_Settings(&s)) || NULL == s)
	{
		Log(L"get_Settings() failed !");
		return;
	}
	
	AutomaticUpdatesNotificationLevel l;
	if (S_OK != s->get_NotificationLevel(&l))
	{
		Log(L"get_NotificationLevel() failed !");
		return;
	}

	switch(l)
	{
	case aunlNotConfigured:
		Log(L"Automatic Updates is not configured by the user or by a Group Policy administrator. Users are periodically prompted to configure Automatic Updates.");
		break;
	case aunlDisabled:
		Log(L"Automatic Updates is disabled. Users are not notified of important updates for the computer");
		break;
	case aunlNotifyBeforeDownload:
		Log(L"Automatic Updates prompts users to approve updates before it downloads or installs the updates");
		break;
	case aunlNotifyBeforeInstallation:
		Log(L"Automatic Updates automatically downloads updates, but prompts users to approve the updates before installation");
		break;
	case aunlScheduledInstallation:
		Log(L"Automatic Updates automatically installs updates according to the schedule that is specified by the user....");
		break;
	}

}

void CwinUpdatesTesterDlg::OnBnClickedSetSettings()
{
	CComPtr<IAutomaticUpdates> pAutomaticUpdates(NULL);
	const HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates, (void**) &pAutomaticUpdates);		
	if (FAILED(hr) || pAutomaticUpdates == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 

	IAutomaticUpdatesSettings* s = NULL;
	if ((S_OK != pAutomaticUpdates->get_Settings(&s)) || NULL == s)
	{
		Log(L"get_Settings() failed !");
		return;
	}
	
	AutomaticUpdatesNotificationLevel l;
	if(BST_CHECKED == ((CButton*)GetDlgItem(IDC_RADIO1))->GetCheck())
	{
		l = aunlDisabled;
	}
	else if(BST_CHECKED == ((CButton*)GetDlgItem(IDC_RADIO2))->GetCheck())
	{
		l = aunlNotifyBeforeDownload;
	}
	else if(BST_CHECKED == ((CButton*)GetDlgItem(IDC_RADIO3))->GetCheck())
	{
		l = aunlNotifyBeforeInstallation;
	}
	else if(BST_CHECKED == ((CButton*)GetDlgItem(IDC_RADIO4))->GetCheck())
	{
		l = aunlScheduledInstallation;
	}
	else
	{
		Log(L"You broke my app !");
		return;
	}
		
	if (S_OK != s->put_NotificationLevel(l))
	{
		Log(L"put_NotificationLevel() failed !");
		return;
	}

	if (S_OK != s->Save())
	{
		Log(L"Save() failed !");
		return;
	}
}

void CwinUpdatesTesterDlg::OnBnClickedCheckNow()
{
	CComPtr<IAutomaticUpdates> pAutomaticUpdates(NULL);
	const HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates, (void**) &pAutomaticUpdates);
	if (FAILED(hr) || pAutomaticUpdates == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 
	
	const HRESULT h = pAutomaticUpdates->DetectNow();
	Log(L"pAutomaticUpdates->DetectNow()", h);
}



void CwinUpdatesTesterDlg::OnBnClickedGetLastUpdateTime()
{
	CComPtr<IAutomaticUpdates2> pAutomaticUpdates2(NULL);

	const HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates2, (void**) &pAutomaticUpdates2);
	if (FAILED(hr) || pAutomaticUpdates2 == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 

	IAutomaticUpdatesResults* pAutomaticUpdatesResults = NULL;
	if(S_OK != pAutomaticUpdates2->get_Results(&pAutomaticUpdatesResults) || NULL == pAutomaticUpdatesResults)
	{
		Log(L"pAutomaticUpdates2->Results failed!");
		return;
	}

	VARIANT date;	
	VariantInit(&date);

	if (S_OK != pAutomaticUpdatesResults->get_LastSearchSuccessDate(&date))
	{
		Log(L"get_LastSearchSuccessDate call failed !");
		return;
	}
	COleDateTime DateInfo(date);	
	Log( (CString("successfully searched for updates : As displayed in Windows UI, UTC: ") + DateInfo.Format()).GetBuffer());
}

void CwinUpdatesTesterDlg::OnBnClickedLastTimeUpdatesInstlled()
{
	CComPtr<IAutomaticUpdates2> pAutomaticUpdates2(NULL);

	const HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates2, (void**) &pAutomaticUpdates2);
	if (FAILED(hr) || pAutomaticUpdates2 == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 

	IAutomaticUpdatesResults* pAutomaticUpdatesResults = NULL;
	if(S_OK != pAutomaticUpdates2->get_Results(&pAutomaticUpdatesResults) || NULL == pAutomaticUpdatesResults)
	{
		Log(L"pAutomaticUpdates2->Results failed !");
		return;
	}

	VARIANT date;	
	VariantInit(&date);

	if (S_OK != pAutomaticUpdatesResults->get_LastInstallationSuccessDate(&date))
	{
		Log(L"get_LastInstallationSuccessDate call failed !");
		return;
	}
	COleDateTime DateInfo(date);	
	Log( (CString("successfully installed any updates, even if some failures occurred : As displayed in Windows UI, UTC: ") + DateInfo.Format()).GetBuffer());
}

void CwinUpdatesTesterDlg::OnBnClickedDisplayWUUI()
{	
	ExecuteWithUACAsync( CString(g_Sys32Folder + L"wuapp.exe").GetString(), L"", L"", SW_SHOW,L"");
}


bool
CwinUpdatesTesterDlg::ExecuteWithUACAsync(const std::wstring& in_App, const std::wstring& in_Arg, const std::wstring& in_Dir, const int in_ShowWindow, const std::wstring& in_Verb )
{
	SHELLEXECUTEINFO shellInfo = { sizeof(SHELLEXECUTEINFO)
			, SEE_MASK_FLAG_DDEWAIT
			, NULL
			, in_Verb.c_str()
			, in_App.c_str()
			, in_Arg.c_str()
			, in_Dir.c_str()
			, in_ShowWindow
			, NULL
			, NULL
			, NULL
			, NULL
			, 0
			, NULL
			, NULL };

	return ( ShellExecuteEx(&shellInfo) == TRUE );
}
void CwinUpdatesTesterDlg::OnBnClickedRebootPC()
{
	if(RebootComputer(true))
	{
		Log(L"OK");
	}
	else
	{
		Log(L"Failed !");
	}
}


bool CwinUpdatesTesterDlg::RebootComputer(bool bForce)
{
	HANDLE hProcessToken = NULL;
	if (OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hProcessToken))
	{
		TOKEN_PRIVILEGES tp;
		tp.PrivilegeCount = 1;
		tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

		if (LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &tp.Privileges[0].Luid))
		{
			AdjustTokenPrivileges(hProcessToken, FALSE, &tp, 0, NULL, NULL);
		}

		CloseHandle(hProcessToken);
	}

	unsigned long nFlags = EWX_REBOOT;
	if (bForce)
		nFlags |= EWX_FORCE;

	return ExitWindowsEx(nFlags, SHTDN_REASON_MAJOR_APPLICATION | SHTDN_REASON_FLAG_PLANNED) == TRUE;
}
void CwinUpdatesTesterDlg::OnBnClickedGetCurrentRebootState()
{
	CComPtr<ISystemInformation> pSystemInformation(NULL);

	const HRESULT hr = CoCreateInstance(CLSID_SystemInformation, NULL, CLSCTX_INPROC_SERVER, IID_ISystemInformation, (void**) &pSystemInformation);
	if (FAILED(hr) || pSystemInformation == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_SystemInformation) failed !");
		return;		
	} 

	VARIANT_BOOL vb = VARIANT_FALSE;
	if( S_OK != pSystemInformation->get_RebootRequired(&vb))
	{
		Log(L"get_RebootRequired failed !");
		return;		
	}
  
	if(VARIANT_TRUE == vb)
	{
		Log(L"YES, a system restart is required to complete the installation or uninstallation of one or more updates");
	}
	else
	{
		Log(L"NO");
	}
}
void CwinUpdatesTesterDlg::OnBnClickedSetEnabled()
{
	CComPtr<IAutomaticUpdates> pAutomaticUpdates(NULL);
	HRESULT hr = CoCreateInstance(CLSID_AutomaticUpdates, NULL, CLSCTX_INPROC_SERVER, IID_IAutomaticUpdates, (void**) &pAutomaticUpdates);		
	if (FAILED(hr) || pAutomaticUpdates == NULL)		
	{			
		Log(L"CoCreateInstance(CLSID_AutomaticUpdates) failed !");
		return;		
	} 
	
	hr = pAutomaticUpdates->EnableService();
	Log(L"pAutomaticUpdates->EnableService()", hr);
}

void CwinUpdatesTesterDlg::OnBnClickedApplyNow()
{		
	if(RunningOn64BitWindows()) // EXE is only in C:\Windows\System32, none in C:\Windows\SysWOW64
	{
		wchar_t strPath[ MAX_PATH ] = { 0 };
		SHGetSpecialFolderPath(
			0,       // Hwnd
			strPath, // String buffer.
			CSIDL_WINDOWS, // CSLID of folder
			FALSE ); // Create if doesn't exists?

		ExecuteWithUACAsync( CString(CString(strPath) + L"\\sysnative\\wuauclt.exe").GetString(), L"/UpdateNow", L"", SW_SHOW,L"");
	}
	else
	{
		ExecuteWithUACAsync( CString(g_Sys32Folder + L"wuauclt.exe").GetString(),                 L"/UpdateNow", L"", SW_SHOW,L"");		
	}
}
