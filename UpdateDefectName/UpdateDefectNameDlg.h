
// UpdateDefectNameDlg.h: 头文件
//

#pragma once
#include <iostream>
#include <fstream>
#include "ADOLinkToBase.h"
#include <tlhelp32.h>  
//#include <atlstr.h>
//#include <vector>
//#include <algorithm>
#define WM_SHOWTASK (WM_USER + 1)

using namespace std;

// CUpdateDefectNameDlg 对话框
class CUpdateDefectNameDlg : public CDialogEx
{
// 构造
public:
	CUpdateDefectNameDlg(CWnd* pParent = nullptr);	// 标准构造函数
	~CUpdateDefectNameDlg();
// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_UPDATEDEFECTNAME_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	CADOLinkToBase *m_pAdo;
	CString m_strDefectName[64];	

	bool ConnectToDatabase(CADOLinkToBase *&pAdo);
	void DisConnectFromDatabase(CADOLinkToBase *&pAdo);
	//void UpdateDefectInfo(CADOLinkToBase *pAdo, CString strTable, CString strInfo, int nIndex);
	void UpdateDefectInfo(CADOLinkToBase *pAdo, CString strTable, CString strInfo, int nIndex, int nLevel, CString strRectCoordinate);
	afx_msg void OnBnClickedButtonOpen();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedOk();
	void OnTimer(UINT_PTR nIDEvent);

	///////////////////////////////
	void ExistNewReel();	//判断是否有新卷号
	void AutomaticUpdateDatabase();
	void AddOneMsg(CString strInfo);
	void AddOneLog(CString strMonth, CString strDay, CString strInfo);
	bool MakeDirectory(CString strPathName);
	void toTray();			//最小化到托盘
	void DeleteTray();		//销毁托盘图标
	LRESULT OnShowTask(WPARAM wParam, LPARAM lParam);
	void AddHorizontalScrollBar(CListBox& mybox);
	


	CString m_strFilePath;					//用于存储上次找到的文件路径
	CString m_strNeelConfigPath;
	CString m_strReelName;
	//CListCtrl	m_listInfo;
	CListBox	m_listInfo;
	ofstream	m_hLog;

protected:
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);


	//////////////////////////////
};
