#include "pch.h"
#include "ADOLinkToBase.h"

CADOLinkToBase::CADOLinkToBase(void)
{
	::CoInitialize(NULL);
	m_hConnect = m_pConnection.CreateInstance("ADODB.Connection");
	m_hRecordset = m_pRecordset.CreateInstance( "ADODB.Recordset");
}

CADOLinkToBase::~CADOLinkToBase(void)
{
	DisConnect();
	::CoUninitialize();
}

void CADOLinkToBase::DisConnect()
{
	if (m_pConnection->GetState() == adStateOpen)
	{
		try
		{
			m_pConnection->Close();
		}
		catch (_com_error e)
		{
		}
	}
	if (m_pRecordset->GetState() == adStateOpen)
	{
		try
		{
			m_pRecordset->Close();
		}
		catch (_com_error e)
		{
		}
	}
}

bool CADOLinkToBase::Connection(LPCTSTR strServer,LPCTSTR strDataBase,LPCTSTR strPwd,bool bConMode)
{
	bool bCon = false;
	CString ConnectionString;
	if (bConMode)//远程IP登陆
	{
		ConnectionString.Format(_T("Provider=SQLOLEDB; Server=%s; Database=%s; uid=sa; pwd=%s;"),strServer,strDataBase,strPwd);
	} 
	else//本地Windows账户登陆
	{
		ConnectionString.Format(_T("Provider=SQLOLEDB; Server=%s; Database=%s; Integrated Security = SSPI;"),strServer,strDataBase);
	}
	//CString ConnectionString = _T("Provider=SQLOLEDB; Server=192.168.1.24; Database=1安全线综合表_be; uid=sa; pwd=12345678;");
	try
	{
		if (SUCCEEDED(m_hConnect))
		{
			m_hConnect = m_pConnection->Open(_bstr_t(ConnectionString), "", "", adModeUnknown);			
			if (m_pConnection->State) 
				bCon = true;
			else bCon = false;
		}
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(_T("数据库连接失败原因：%s") , e.ErrorMessage());
		AfxMessageBox(log);
	}
	return bCon;
}

bool CADOLinkToBase::Execute(LPCTSTR strSQL)
{
	if (m_pConnection == NULL || m_pRecordset == NULL)
	{
		return FALSE;
	}

	try

	{
		m_pRecordset = m_pConnection->Execute((_bstr_t)strSQL,NULL,adCmdText);
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(_T("数据库连接错误：%s") , e.ErrorMessage());
		//AfxMessageBox(log);
		return FALSE;
	}
	return TRUE;
}

bool CADOLinkToBase::GetCollect(LPCTSTR ColumnName, CString& strValue)
{
	try

	{
		if (m_pRecordset->adoEOF) //为空的判断
		{
			return FALSE;
		}
		_variant_t value = m_pRecordset->GetCollect(_variant_t(LPCTSTR(ColumnName)));
		strValue = var2str(value);
		return TRUE;
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(_T("数据库连接错误：%s") , e.ErrorMessage());
		AfxMessageBox(log);
		return FALSE;
	}

	return TRUE;
}

bool CADOLinkToBase::IsEmpty(void)
{
	return m_pRecordset->adoEOF;
}

void CADOLinkToBase::NextRecd(void)
{
	m_pRecordset->MoveNext();
}

/*-------------------------type transform --------------- 
【name】:var2str
【prameters】:var: input _variant_t 
【return】:CString 
【describe】:change _variant_t to CString
include var.vt = int/float/cstring/long
---------------------------------------------------*/
CString CADOLinkToBase::var2str(const _variant_t& var)
{
	CString strValue;
	switch (var.vt)
	{
	case VT_BSTR://字符串
	case VT_LPSTR://字符串
	case VT_LPWSTR://字符串
		strValue = (LPCTSTR)(_bstr_t)var;
		break;
	case VT_I1:
	case VT_UI1:
		strValue.Format(_T("%d"), var.bVal);
		break;
	case VT_I2://短整型
		strValue.Format(_T("%d"), var.iVal);
		break;
	case VT_UI2://无符号短整型
		strValue.Format(_T("%d"), var.uiVal);
		break;
	case VT_INT://整型
		strValue.Format(_T("%d"), var.intVal);
		break;
	case VT_I4: //整型
		strValue.Format(_T("%d"), var.lVal);
		break;
	case VT_I8: //长整型
		strValue.Format(_T("%d"), var.lVal);
		break;
	case VT_UINT://无符号整型
		strValue.Format(_T("%d"), var.uintVal);
		break;
	case VT_UI4: //无符号整型
		strValue.Format(_T("%d"), var.ulVal);
		break;
	case VT_UI8: //无符号长整型
		strValue.Format(_T("%d"), var.ulVal);
		break;
	case VT_VOID:
		strValue.Format(_T("%8x"), var.byref);
		break;
	case VT_R4://浮点型
		strValue.Format(_T("%.4f"), var.fltVal);
		break;
	case VT_R8://双精度型
		strValue.Format(_T("%.4f"), var.dblVal);
		break;
	case VT_DECIMAL: //小数
		strValue.Format(_T("%.4f"), (double)var);
		break;
	case VT_CY:
		{
			COleCurrency cy = var.cyVal;
			strValue = cy.Format();
		}
		break;
	case VT_BLOB:
	case VT_BLOB_OBJECT:
	case 0x2011:
		strValue = "[BLOB]";
		break;
	case VT_BOOL://布尔型

		strValue = var.boolVal ? "TRUE" : "FALSE";
		break;
	case VT_DATE: //日期型
		{
			DATE dt = var.date;
			COleDateTime da = COleDateTime(dt); 
			strValue = da.Format(_T("%Y-%m-%d %H:%M:%S"));
		}
		break;
	case VT_NULL://NULL值
		strValue = "";
		break;
	case VT_EMPTY://空
		strValue = "";
		break;
	case VT_UNKNOWN://未知类型
	default:
		strValue = "UN_KNOW";
		break;
	}
	return strValue;
}