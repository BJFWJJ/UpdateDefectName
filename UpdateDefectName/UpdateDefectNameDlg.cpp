// UpdateDefectNameDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "UpdateDefectName.h"
#include "UpdateDefectNameDlg.h"
#include "afxdialogex.h"
#include "direct.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CUpdateDefectNameDlg 对话框

CUpdateDefectNameDlg::CUpdateDefectNameDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_UPDATEDEFECTNAME_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	m_pAdo = nullptr;
	
	CString strPathConfig = _T("F:\\Inference\\path_config.ini");
	GetPrivateProfileString(_T("param"), _T("file_path"), _T("NULL"), m_strFilePath.GetBuffer(128), 128, strPathConfig);

}

CUpdateDefectNameDlg::~CUpdateDefectNameDlg()
{
	if (m_pAdo)
	{
		delete m_pAdo;
		m_pAdo = nullptr;
	}

	if (m_hLog.is_open())
	{
		m_hLog.close();
	}
}

void CUpdateDefectNameDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_INFO, m_listInfo);
}

BEGIN_MESSAGE_MAP(CUpdateDefectNameDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_BUTTON_OPEN, &CUpdateDefectNameDlg::OnBnClickedButtonOpen)
	ON_BN_CLICKED(IDCANCEL, &CUpdateDefectNameDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDOK, &CUpdateDefectNameDlg::OnBnClickedOk)
	ON_MESSAGE(WM_SHOWTASK, OnShowTask)


END_MESSAGE_MAP()

// CUpdateDefectNameDlg 消息处理程序
BOOL CUpdateDefectNameDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	for (int i = 0; i < 64; i++)
	{
		CString strKey;
		strKey.Format(_T("%d"), i);
		GetPrivateProfileString(_T("Param"), strKey, _T("无分类"), m_strDefectName[i].GetBuffer(64), 64, _T("D:\\DefectMap.ini"));
		m_strDefectName[i].ReleaseBuffer();
	}

	SetTimer(1, 10000, nullptr);		//每10秒执行一次
	SetTimer(3, 5000, nullptr);			//最小化到托盘

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CUpdateDefectNameDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

void CUpdateDefectNameDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	// 设置最小宽度
	lpMMI->ptMinTrackSize.x = 300;  // 最小宽度
	lpMMI->ptMinTrackSize.y = 200;  // 最小高度
}


//根据进程名字结束进程
bool TerminateProcessByName(const char *procName)
{
	DWORD procId = 0;

	// 获取进程ID  
	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	PROCESSENTRY32 processEntry;
	processEntry.dwSize = sizeof(PROCESSENTRY32);

	if (Process32First(snapshot, &processEntry))
	{
		do
		{
			if (!strcmp(processEntry.szExeFile, procName))
			{
				procId = processEntry.th32ProcessID;
				break;
			}
		} while (Process32Next(snapshot, &processEntry));
	}

	CloseHandle(snapshot);

	// 如果有找到进程，则终止它  
	if (procId != 0)
	{
		HANDLE procHandle = OpenProcess(PROCESS_ALL_ACCESS, FALSE, procId);
		if (procHandle)
		{
			TerminateProcess(procHandle, 1);
			CloseHandle(procHandle);
		}
		return true;
	}
	else
	{
		return false;
	}
}

void CUpdateDefectNameDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	switch (nIDEvent)
	{
	case 1:
	{
		ExistNewReel();
	}
	break;

	case 2:
	{
		HandleInferProcess();
		KillTimer(2);
	}
	break;

	case 3:
	{
		toTray();
		KillTimer(3);
	}
	break;

	default:
		break;
	}

	CDialog::OnTimer(nIDEvent);
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
HCURSOR CUpdateDefectNameDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

bool CUpdateDefectNameDlg::ConnectToDatabase(CADOLinkToBase *&pAdo)
{
	if (pAdo)
	{
		delete pAdo;
		pAdo = nullptr;
	}
	pAdo = new CADOLinkToBase;
	bool bDbConnect = pAdo->Connection(_T("10.169.70.170"), _T("DB_CENTRAL_UI"), _T("kexin2008"), true);
	return bDbConnect;
}

void CUpdateDefectNameDlg::DisConnectFromDatabase(CADOLinkToBase *&pAdo)
{
	if (pAdo)
	{
		delete pAdo;
		pAdo = nullptr;
	}
}

CString EscapeSingleQuotes(CString str)
{
	str.Replace(_T("'"), _T("''"));
	return str;
}

void CUpdateDefectNameDlg::UpdateDefectInfo(CADOLinkToBase *pAdo, CString strTable, CString strInfo, int nIndex, int nLevel, CString strRectCoordinate)
{
	CString strSQL;
	strSQL.Format(_T("UPDATE [%s] SET Feature_Name14 = '%s', Feature_nValue14 = %f ,Feature_Name13 = '%s' WHERE Index_InQue = %d"),
		strTable,
		EscapeSingleQuotes(strInfo),  
		static_cast<float>(nLevel),
		strRectCoordinate,
		nIndex);

	if (!pAdo->Execute(strSQL))
	{
		CString strErr;
		strErr.Format(_T("数据库更新失败！\nSQL语句:%s"), strSQL);
		AddOneMsg(strErr);
	}
}

void CUpdateDefectNameDlg::OnBnClickedButtonOpen()
{
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog dlgOpenFile(TRUE, 0, 0, OFN_NOCHANGEDIR, _TEXT("CSV File (*.ini)|*.ini||"), this);
	if (dlgOpenFile.DoModal() == IDOK)
	{
		CString strReelConfigPath;
		strReelConfigPath = dlgOpenFile.GetPathName();
		GetDlgItem(IDC_STATIC_PATH)->SetWindowText(strReelConfigPath);

		CString strReelTable,strReelName;
		GetPrivateProfileString(_T("param"), _T("reel_table"), _T("NULL"), strReelTable.GetBuffer(128), 128, strReelConfigPath);

		CString strInfo;
		strInfo.Format(_T("正在写入数据库..."));
		AddOneMsg(strInfo);

		bool bWriteToDB = WriteToDB(strReelTable);
		if (bWriteToDB)
		{
			strInfo.Format(_T("数据表：%s 完成手动更新！"), strReelTable);
		}
		else
		{
			strInfo.Format(_T("数据表：%s 手动更新失败，请重试或检查文件！"), strReelTable);
		}
		AddOneMsg(strInfo);
		strReelTable.ReleaseBuffer();
	}
}

//自动更新数据库
void CUpdateDefectNameDlg::AutomaticUpdateDatabase(CString strReelTable)
{
	CString strInfo;
	bool bWriteToDB = WriteToDB(strReelTable);
	if (bWriteToDB)
	{
		strInfo.Format(_T("数据表：%s 完成自动更新！"), strReelTable);
	}
	else
	{
		strInfo.Format(_T("数据表：%s 自动更新失败，请重试或手动更新！"), strReelTable);
	}
	AddOneMsg(strInfo);
	strReelTable.ReleaseBuffer();
}

bool CUpdateDefectNameDlg::WriteToDB(CString strReelTable)
{
	CString strCsvPath;
	strCsvPath.Format(_T("F:\\Inference\\%s\\result.csv"), strReelTable); 

	FILE *pFile = NULL;
	fopen_s(&pFile, strCsvPath, "r");
	if (pFile)
	{
		ConnectToDatabase(m_pAdo);

		int nFileIndex = 0;
		int nDefectIndex = 0;
		int nDefectLevel = 0;
		char strName[64];
		float fCentreX, fCentreY, fLength, fHeight;
		fCentreX = fCentreY = fLength = fHeight = -9.9999;
		while (fscanf_s(pFile, "%d,%d,%d,%f,%f,%f,%f\n", &nFileIndex, &nDefectIndex, &nDefectLevel, &fCentreX, &fCentreY, &fLength, &fHeight, 64) != EOF)
		{
			CString strRectCoordinate;
			strRectCoordinate.Format(_T("%f,%f,%f,%f"), fCentreX, fCentreY, fLength, fHeight);
			if (strRectCoordinate.Find(_T("-9.9999")) >= 0)	//找到了-9.9999这个值	
			{
				strRectCoordinate.Format(_T(""));
			}
			UpdateDefectInfo(m_pAdo, strReelTable, m_strDefectName[nDefectIndex], nFileIndex, nDefectLevel, strRectCoordinate);

			fCentreX = fCentreY = fLength = fHeight = -9.9999;
		}
		fclose(pFile);

		DisConnectFromDatabase(m_pAdo);

		WritePrivateProfileString(_T("param"), _T("is_execute"), "1", m_strReelConfigPath);
	}
	else
		return false;

	return true;
}

void CUpdateDefectNameDlg::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnCancel();
}


void CUpdateDefectNameDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	toTray();
}

void CUpdateDefectNameDlg::ExistNewReel()
{
	//1.判断是否获取到最新文件
	CString strPathConfig = _T("F:\\Inference\\path_config.ini");
	CString strGetFilePath;		//保存path_config.ini中最新获取到的文件路径
	CString strLog;
	GetPrivateProfileString(_T("param"), _T("file_path"), _T("NULL"), strGetFilePath.GetBuffer(128), 128, strPathConfig);

	if (m_strFilePath != strGetFilePath)
	{
		m_strFilePath = strGetFilePath;
		m_strReelConfigPath.Format(_T("%s\\config.ini"), strGetFilePath);
		//GetPrivateProfileString(_T("param"), _T("reel_name"), _T("NULL"), m_strReelName.GetBuffer(128), 128, m_strReelConfigPath);
		GetPrivateProfileString(_T("param"), _T("reel_table"), _T("NULL"), m_strReelTable.GetBuffer(128), 128, m_strReelConfigPath);
		if (m_strReelTable != _T("NULL"))
		{
			
			strLog.Format(_T("找到新文件：%s。"), m_strReelTable);
			AddOneMsg(strLog);

			//2.进行Python深度学习步骤
			bool b_IsExecute = GetPrivateProfileInt(_T("param"), _T("is_execute"), 0, m_strReelConfigPath);
			if (!b_IsExecute)		//IsExecute=0说明深度学习分类分级未执行完成，为1执行完成
			{
				ShellExecute(NULL, _T("open"), _T("C:\\Users\\adv\\Desktop\\test_script.sh"), NULL, NULL, SW_SHOWNORMAL);	//后续移动到配置文件里
				strLog.Format(_T("%s 正在进行缺陷分类分级。"), m_strReelTable);
				AddOneMsg(_T(strLog));
			}
			else
			{
				strLog.Format(_T("文件：%s 已经进行了缺陷分级分类。"), m_strReelTable);
				AddOneMsg(strLog);
			}
		}

		else
		{
			return;
		}
	}
	//3.循环判断Python深度学习步骤是否完成，完成立即关闭
	SetTimer(2, 2000, nullptr);
}

void CUpdateDefectNameDlg::HandleInferProcess()
{
	int nIsInfer = GetPrivateProfileInt(_T("param"), _T("is_infer"), 0, m_strReelConfigPath);
	if (nIsInfer == 1)
	{
		// 关闭外部程序		//.exe结尾的程序uccess
		TerminateProcessByName("mintty.exe");
		AddOneMsg(_T("缺陷分类完成，正在写入数据库..."));
		WritePrivateProfileString(_T("param"), _T("is_infer"), "0", m_strReelConfigPath);
		CString strReelTable;
		strReelTable = m_strReelTable;
		AutomaticUpdateDatabase(strReelTable);//执行自动更新数据库操作
	}
	else if (nIsInfer == 2)
	{
		TerminateProcessByName("mintty.exe");
		AddOneMsg(_T("训练异常，请重试！"));
	}
	else
	{
		return;
	}
}

void CUpdateDefectNameDlg::AddOneMsg(CString strInfo)
{
	CTime time = CTime::GetCurrentTime();
	CString strTime;
	strTime.Format(_T("%04d-%02d-%02d %02d:%02d:%02d"),
		time.GetYear(), time.GetMonth(), time.GetDay(),
		time.GetHour(), time.GetMinute(), time.GetSecond());

	strInfo = strTime + _T("   ") + strInfo;
	//m_listInfo.InsertItem(0, strInfo);
	m_listInfo.InsertString(m_listInfo.GetCount(), strInfo);
	m_listInfo.SetCurSel(m_listInfo.GetCount() - 1);

	CString strMonth, strDay;
	strMonth.Format(_T("%04d-%02d"), time.GetYear(), time.GetMonth());
	strDay.Format("%02d", time.GetDay());
	AddOneLog(strMonth, strDay, strInfo);

	AddHorizontalScrollBar(m_listInfo);
}

void CUpdateDefectNameDlg::AddOneLog(CString strMonth, CString strDay, CString strInfo)
{
	CString strPath = _T("..\\..\\Log\\") + strMonth + _T("\\");
	if (!PathFileExists(strPath))
		MakeDirectory(strPath);
	CString strFileDir = strPath + strDay + _T("log.txt");

	if (PathFileExists(strFileDir))
	{
		if (!m_hLog.is_open())
		{
			m_hLog.open(strFileDir, ios::in | ios::ate | ios::out);
		}
	}
	else
	{
		if (!m_hLog.is_open())
		{
			m_hLog.open(strFileDir);
		}
		else
		{
			m_hLog.close();
			m_hLog.open(strFileDir);
		}
	}
	strInfo += _T("\r\n");
	m_hLog << strInfo << flush;
}

bool CUpdateDefectNameDlg::MakeDirectory(CString strPathName)
{
	// 备份路径设置
	CString strTmpPathName = strPathName;

	// 补齐最后一个文件夹名，好做循环
	if (strTmpPathName.Right(1) != "\\")
		strTmpPathName += "\\";

	int nStringLength = strTmpPathName.GetLength();

	//	开始查找根目录
	int nCurCharPosition = strTmpPathName.Find('\\');

	//	跳过盘符
	if (strTmpPathName[nCurCharPosition - 1] == ':')
		nCurCharPosition = strTmpPathName.Find('\\', nCurCharPosition + 1);

	CString strSubPathName;

	while (nCurCharPosition != -1 && nCurCharPosition <= nStringLength)
	{
		strSubPathName = strTmpPathName.Left(nCurCharPosition + 1);

		if (!strSubPathName.IsEmpty())	//如果不为空
			_mkdir(strSubPathName);

		nCurCharPosition = strTmpPathName.Find('\\', nCurCharPosition + 1);
	}

	return true;
}

void CUpdateDefectNameDlg::toTray()//最小化到托盘
{
	NOTIFYICONDATA nid;
	nid.cbSize = (DWORD)sizeof(NOTIFYICONDATA);
	nid.hWnd = this->m_hWnd;
	nid.uID = IDR_MAINFRAME;
	nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	nid.uCallbackMessage = WM_SHOWTASK;//自定义的消息名称  
	nid.hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MAINFRAME));
	strcpy_s(nid.szTip, "更新数据库");//信息提示条
	Shell_NotifyIcon(NIM_ADD, &nid);//在托盘区添加图标  
	ShowWindow(SW_HIDE);//隐藏主窗口 
}

void CUpdateDefectNameDlg::DeleteTray()//销毁托盘图标
{
	NOTIFYICONDATA nid;
	nid.cbSize = (DWORD)sizeof(NOTIFYICONDATA);
	nid.hWnd = this->m_hWnd;
	nid.uID = IDR_MAINFRAME;
	nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	nid.uCallbackMessage = WM_SHOWTASK;
	nid.hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MAINFRAME));
	strcpy_s(nid.szTip, "更新数据库");
	Shell_NotifyIcon(NIM_DELETE, &nid);
}

LRESULT CUpdateDefectNameDlg::OnShowTask(WPARAM wParam, LPARAM lParam)
{//wParam接收的是图标的ID，而lParam接收的是鼠标的行为 
	if (wParam != IDR_MAINFRAME)
		return 1;
	switch (lParam)
	{
		case WM_LBUTTONDBLCLK://左键单击显示主界面
		{
			this->ShowWindow(SW_SHOW);
			SetForegroundWindow();
			DeleteTray();
		}
		break;
	}
	return 0;
}

void CUpdateDefectNameDlg::AddHorizontalScrollBar(CListBox& mybox)
{
	CString str;
	CSize sz;
	int dx = 0;
	CDC* pDC = mybox.GetDC();
	for (int j = 0; j < mybox.GetCount(); j++)
	{
		mybox.GetText(j, str);
		sz = pDC->GetTextExtent(str);

		if (sz.cx > dx)
			dx = sz.cx;
	}
	mybox.ReleaseDC(pDC);
	mybox.SetHorizontalExtent(dx);
}