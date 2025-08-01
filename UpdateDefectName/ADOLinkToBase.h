#pragma once
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" no_namespace rename("EOF", "adoEOF")

class CADOLinkToBase
{
public:
	CADOLinkToBase(void);
	~CADOLinkToBase(void);

public:
	_ConnectionPtr m_pConnection;
	_RecordsetPtr m_pRecordset;
	HRESULT m_hConnect;
	HRESULT m_hRecordset;
public:
	bool Connection(LPCTSTR strServer,LPCTSTR strDataBase,LPCTSTR strPwd,bool bConMode);
	void DisConnect();
	bool Execute(LPCTSTR strSQL);
	bool GetCollect(LPCTSTR ColumnName, CString& strValue);
	bool IsEmpty(void);//�ж�Recordsetָ���Ƿ�ָ�����һ��
	void NextRecd(void);//��Recordsetָ��������һ��

public:
	CString var2str(const _variant_t& var); //ado _variant_t ��������ת�� 
};
