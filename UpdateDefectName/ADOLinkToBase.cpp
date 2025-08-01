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
	if (bConMode)//Զ��IP��½
	{
		ConnectionString.Format(_T("Provider=SQLOLEDB; Server=%s; Database=%s; uid=sa; pwd=%s;"),strServer,strDataBase,strPwd);
	} 
	else//����Windows�˻���½
	{
		ConnectionString.Format(_T("Provider=SQLOLEDB; Server=%s; Database=%s; Integrated Security = SSPI;"),strServer,strDataBase);
	}
	//CString ConnectionString = _T("Provider=SQLOLEDB; Server=192.168.1.24; Database=1��ȫ���ۺϱ�_be; uid=sa; pwd=12345678;");
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
		log.Format(_T("���ݿ�����ʧ��ԭ��%s") , e.ErrorMessage());
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
		log.Format(_T("���ݿ����Ӵ���%s") , e.ErrorMessage());
		//AfxMessageBox(log);
		return FALSE;
	}
	return TRUE;
}

bool CADOLinkToBase::GetCollect(LPCTSTR ColumnName, CString& strValue)
{
	try

	{
		if (m_pRecordset->adoEOF) //Ϊ�յ��ж�
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
		log.Format(_T("���ݿ����Ӵ���%s") , e.ErrorMessage());
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
��name��:var2str
��prameters��:var: input _variant_t 
��return��:CString 
��describe��:change _variant_t to CString
include var.vt = int/float/cstring/long
---------------------------------------------------*/
CString CADOLinkToBase::var2str(const _variant_t& var)
{
	CString strValue;
	switch (var.vt)
	{
	case VT_BSTR://�ַ���
	case VT_LPSTR://�ַ���
	case VT_LPWSTR://�ַ���
		strValue = (LPCTSTR)(_bstr_t)var;
		break;
	case VT_I1:
	case VT_UI1:
		strValue.Format(_T("%d"), var.bVal);
		break;
	case VT_I2://������
		strValue.Format(_T("%d"), var.iVal);
		break;
	case VT_UI2://�޷��Ŷ�����
		strValue.Format(_T("%d"), var.uiVal);
		break;
	case VT_INT://����
		strValue.Format(_T("%d"), var.intVal);
		break;
	case VT_I4: //����
		strValue.Format(_T("%d"), var.lVal);
		break;
	case VT_I8: //������
		strValue.Format(_T("%d"), var.lVal);
		break;
	case VT_UINT://�޷�������
		strValue.Format(_T("%d"), var.uintVal);
		break;
	case VT_UI4: //�޷�������
		strValue.Format(_T("%d"), var.ulVal);
		break;
	case VT_UI8: //�޷��ų�����
		strValue.Format(_T("%d"), var.ulVal);
		break;
	case VT_VOID:
		strValue.Format(_T("%8x"), var.byref);
		break;
	case VT_R4://������
		strValue.Format(_T("%.4f"), var.fltVal);
		break;
	case VT_R8://˫������
		strValue.Format(_T("%.4f"), var.dblVal);
		break;
	case VT_DECIMAL: //С��
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
	case VT_BOOL://������

		strValue = var.boolVal ? "TRUE" : "FALSE";
		break;
	case VT_DATE: //������
		{
			DATE dt = var.date;
			COleDateTime da = COleDateTime(dt); 
			strValue = da.Format(_T("%Y-%m-%d %H:%M:%S"));
		}
		break;
	case VT_NULL://NULLֵ
		strValue = "";
		break;
	case VT_EMPTY://��
		strValue = "";
		break;
	case VT_UNKNOWN://δ֪����
	default:
		strValue = "UN_KNOW";
		break;
	}
	return strValue;
}