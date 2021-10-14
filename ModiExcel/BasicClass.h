#if !defined(_BASICCLASS_H__BA14C3D0_C5E8_4D24_90B8_DED6A84AEF10__INCLUDED_)
#define _BASICCLASS_H__BA14C3D0_C5E8_4D24_90B8_DED6A84AEF10__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "zjgBasic.h"			// MFC support for Windows Common Controls

class ExMethod;
class ExClass;
class ExClassLine;

class ExMethod
{
public:
	ExMethod();
	ExMethod(const ExMethod& Other);
	~ExMethod();
	const ExMethod& operator=(const ExMethod& Other);//��ֵ
	void Empty();												//���
	BOOL IsEmpty();												//�ж��Ƿ�Ϊ��

	void SetName(CString nName){mName = nName;};				//������
	CString GetName(){return mName;};							//������
	void SetOutType(CString n_OutType){m_OutType = n_OutType;};	//�������ص���������
	CString GetOutType(){return m_OutType;};					//�������ص���������
	void SetInType(CString n_InType){m_InType = n_InType;};		//�����������������������
	CString GetInType(){return m_InType;};						//�����������������������
	void SetVARTYPE(CString n_VARTYPE){m_VARTYPE = n_VARTYPE;};	//InvokeHelper�����ķ�������
	CString GetVARTYPE(){return m_VARTYPE;};					//InvokeHelper�����ķ�������
	void SetDISPID(CString n_DISPID){m_DISPID = n_DISPID;};		//InvokeHelper�����Ľӿ�����
	CString GetDISPID(){return m_DISPID;};						//InvokeHelper�����Ľӿ�����

	CString Get_VARENUM_Type(CString nVARENUM);		//��ȡVARENUM��Ӧ����������
	CString Get_Type_VARENUM(CString nVARENUM);		//��ȡ�������Ͷ�ӦVARENUM��ֵ

	CString GetMethodType();						//��ȡInvokeHelper�����ķ�������
	CString GetID();								//ͨ���ֽ�InvokeHelper��������ȡ�����ӿ����ʹ���

	BOOL DecomposeIvHp();							//�ֽ�InvokeHelper��䣬����д��Ӧ�Ĳ���
	BOOL DecomposeIvHp(CString strLine);			//�ֽ�InvokeHelper��䣬����д��Ӧ�Ĳ���
	BOOL DecomposeMeth(CString strLine);			//�ֽ⺯��������䣬����д������
	void Modify_Method(CList<CString,CString&> * pList);	//�Ե�ǰ����Ϊ�ο��࣬����pList��ָ����������޸�
	BOOL Modify_Method_PUT(CList<CString,CString&> * pList);//�Ե�ǰ����Ϊ�ο��࣬����pList��ָ����������޸�

public:
	CList<CString,CString&> mLine;
private:
	CString mName;						//������
	CString m_OutType;					//�����ķ������ͣ���m_VARTYPE�ж�Ӧ��ϵ
	CString m_VARTYPE;					//InvokeHelper�����еĵ�����������������������
	CString m_DISPID;					//InvokeHelper�����Ľӿ����ͣ�Ϊһ������
	CString m_InType;					//����������ο�����������

// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );

};

class ExClass
{
public:
	ExClass();
	ExClass(const ExClass& Other);
	~ExClass();

	const ExClass& operator=(const ExClass& Other);		//��ֵ
	void Empty();										//���
	BOOL IsEmpty();										//�ж��Ƿ��ǿ�
	
	void SetName(CString nName){mName = nName;};		//��������
	CString GetName(){return mName;};					//��ȡ����

	CString GetType(CString dwDispID);					//���ݷ���(����)�ı�ǣ���ȡ��������ֵ��Ӧ����������
	CString Get_VARENUM_Type(CString nVARENUM);			//��ȡVARENUM��Ӧ����������
	CString Get_Type_VARENUM(CString nVARENUM);			//��ȡ�������Ͷ�ӦVARENUM��ֵ

	POSITION Find_PUT(CString dwDispID);				//����dwDispID�����Ҷ�Ӧ�Ľӿڵ�PUT����
	void FindDispID(CString dwDispID,CList<ExMethod,ExMethod&>* pLine);			//����InvokeHelper�����Ľӿ����ͣ����Ҷ�Ӧ�ķ���

	int Modify_Method(CList<CString,CString&> * pList,CString * pstrMethod);	//�Ե�ǰ��Ϊ�ο��࣬����pList��ָ����������޸�
	int Modify_Method(ExMethod * pMethod);				//�Ե�ǰ��Ϊ�ο��࣬��(*pMethod)���������޸�

	void DecomposeClassName(CString strLine);			//�ֽ��ඨ����䣬����д��Ӧ�Ĳ���

public:
	CList<ExMethod,ExMethod&> mLine;
private:
	CString mName;		//����
	CString mBase;		//������
};

#define FILE_FINISHED		(0x0000)		//�������
#define FILE_NAME_ERROR		(0x0001)		//���ƴ���
#define FILE_READ_ERROR		(0x0002)		//������
#define FILE_WRITE_ERROR	(0x0004)		//д����
#define FILE_LESS_ERROR		(0x0008)		//ȱ��ԭ��
#define FILE_PART_ERROR		(0x0010)		//���ִ���
#define FILE_VARTYPE_ERR	(0x0020)		//�������ʹ���
#define FILE_C_LESSID_ERR	(0x0040)		//����ȱ�ٷ����ӿ�ID����
#define FILE_LN_LESSID_ERR	(0x0080)		//������ȱ�ٷ���������Ľӿ�ID����
#define FILE_LI_LESSID_ERR	(0x0100)		//������ȱ�ٽӿ�ID����
#define FILE_C_LESSPUT_ERR	(0x0200)		//����ȱ�ٳɶԵ�ID����


#define FILE_ERROR_ERR		(0x8000)				//�ļ����𻵵Ĵ���

class ExClassLine
{
public:
	ExClassLine();
	~ExClassLine();

public:
	void Empty();
	void SetFileName(CString nName){mFileName = nName;};
	CString GetFileName(){return mFileName;};

	
	BOOL LoadCPPFile(CString FileName);			//װ��.CPP��ʽ�ļ�
	int Load_H_File(CString FileName);			//װ��.h��ʽ�ļ�

	POSITION Find(CString nClassName);			//�����������������λ��
	POSITION Find_A(CString nClassName);		//���������������������������λ��
	POSITION Find_Method_PUT(ExMethod* pMethod);//���ݷ�������dwDispID�����Ҷ�Ӧ�Ľӿڵ�PUT����

	int Modify_H_File_A(CString OldFName,CString NewFName);	//�޸�.h��ʽ�ļ�
	int Modify_H_File_PUT(CString OldFileName,CString NewFileName);				//�޸�.h��ʽ�ļ�											
	                                            //����ֵ��1.�ļ�������ȷ��2.�ļ��򲻿���3.�ļ�����д��
	int Modi_Method_Name(ExMethod* pMethod);	//���ں�������ID��ͬ��ԭ������޸�
	int Modi_Method_ID(ExMethod* pMethod);		//����ID��ͬ��ԭ������޸�

	BOOL FillCBCtrl(CComboBox * pCBCtrl);							//���ο���������䵽��Ͽ���
	BOOL FillListCtrl(CListBox * pListCtrl,CString strFileName);	//���޸ĳ��ִ�����ļ�����䵽�б�
	BOOL FillStaticCtrl(CString * pCStrCtrl,CString strFileName);	//���޸ĳ��ִ�����ļ�����䵽�б�

	BOOL Save(CString FileName);

private:
	
public:
	CList<ExClass,ExClass&> mLine;
private:
	CString mFileName;


};

#define ENUM_NULL	0x00FFFFFA		//ö�ٿ�ֵ

#define ENUM_NOTHING				0x00000000		//�����û����Ӧ�ĵ���
#define ENUM_HAVE_ENUM				0x00000001		//������� "emun"����
#define ENUM_HAVE_LEFT				0x00000002		//������� "{"������
#define ENUM_HAVE_RIGHT				0x00000004		//������� "}"������
#define ENUM_HAVE_SEMICOLON			0x00000008		//������� ";"�ֺ�
#define ENUM_HAVE_RIGHTSEMICOLON	0x00000010		//����� "}"��";"���������м�ֻ�пո��TAB��

class EnumElement
{
public:
	EnumElement();
	~EnumElement();
	
	const EnumElement& operator=(const EnumElement& Other);			//��ֵ
	void Empty();													//���
	BOOL IsEmpty();													//�ж��Ƿ��ǿ�

	CString GetName(){return mName;};								//��Ա��
	void SetName(CString nName){mName = nName;};					//��Ա��
	int GetValue(){return mValue;};									//ֵ
	void SetValue(int nValue){mValue = nValue;};						//ֵ
	BOOL SetValue(CString nName);									//ֵ
	CString GetAnnotate(){return mAnnotate;};						//ע��
	void SetAnnotate(CString nAnnotate){mAnnotate = nAnnotate;};	//ע��

	int Decompose_CPP(CString strLine);			//��Ⲣ�ֽ�.CPP��.h�ļ���ö�ٶ������
	BOOL DecomposeValue(CString strLine);			//�ֽ��ำֵ��䣬����д��Ӧ�Ĳ���
	CString Compose();
	CString ComposeEnd();

private:
	CString mName;						//��Ա��
	int mValue;							//ֵ
	CString mAnnotate;					//ע��
};

class DecEnum
{
public:
	DecEnum();
	~DecEnum();
	
	const DecEnum& operator=(const DecEnum& Other);			//��ֵ
	void Empty();											//���
	BOOL IsEmpty();	

	CString GetName(){return mName;};
	void SetName(CString nName){mName = nName;};

	POSITION Find(CString strElementName);
	POSITION Find(int iElementValue);

	BOOL Compose(CList<CString,CString&> * pList);			//��ϳ�ö�����͵��ַ�������
	BOOL Decompose(CList<CString,CString&> * pList);		//�������е��ַ������ֽ��һ��ö����
	BOOL Marge(DecEnum* pOther);							//��pOther��ָ��ö����ϲ�����ǰ����
public:
	CList<EnumElement,EnumElement&> mLine; 
private:
	CString mName;
};

class DecEnumList
{
public:
	DecEnumList();
	~DecEnumList(){Empty();};

	void Empty();											//���
	BOOL IsEmpty();

	POSITION Find(CString DecEnumName);						//����ö����

	BOOL Load_CPP_H_File(CString FileName);					//װ��.CPP��.h��ʽ�ļ�
	BOOL Load_Txt_File(CString FileName);						//װ��.txt��ʽ�ļ�
	BOOL Create_H(CString FileName);						//����.H��ʽ�ļ�
	void Marge();											//ö����ϲ�

public:
	CList<DecEnum,DecEnum&> mLine; 
private:
	CString mName;											//�ļ���

};
/*
enum VARENUM
{
	VT_EMPTY	= 0,
	VT_NULL	= 1,
	VT_I2	= 2,
	VT_I4	= 3,
	VT_R4	= 4,
	VT_R8	= 5,
	VT_CY	= 6,
	VT_DATE	= 7,
	VT_BSTR	= 8,
	VT_DISPATCH	= 9,
	VT_ERROR	= 10,
	VT_BOOL	= 11,
	VT_VARIANT	= 12,
	VT_UNKNOWN	= 13,
	VT_DECIMAL	= 14,
	VT_I1	= 16,
	VT_UI1	= 17,
	VT_UI2	= 18,
	VT_UI4	= 19,
	VT_I8	= 20,
	VT_UI8	= 21,
	VT_INT	= 22,
	VT_UINT	= 23,
	VT_VOID	= 24,
	VT_HRESULT	= 25,
	VT_PTR	= 26,
	VT_SAFEARRAY	= 27,
	VT_CARRAY	= 28,
	VT_USERDEFINED	= 29,
	VT_LPSTR	= 30,
	VT_LPWSTR	= 31,
	VT_RECORD	= 36,
	VT_INT_PTR	= 37,
	VT_UINT_PTR	= 38,
	VT_FILETIME	= 64,
	VT_BLOB	= 65,
	VT_STREAM	= 66,
	VT_STORAGE	= 67,
	VT_STREAMED_OBJECT	= 68,
	VT_STORED_OBJECT	= 69,
	VT_BLOB_OBJECT	= 70,
	VT_CF	= 71,
	VT_CLSID	= 72,
	VT_VERSIONED_STREAM	= 73,
	VT_BSTR_BLOB	= 0xfff,
	VT_VECTOR	= 0x1000,
	VT_ARRAY	= 0x2000,
	VT_BYREF	= 0x4000,
	VT_RESERVED	= 0x8000,
	VT_ILLEGAL	= 0xffff,
	VT_ILLEGALMASKED	= 0xfff,
	VT_TYPEMASK	= 0xfff
    } ;
*/
// vtRet�� 
// VT_EMPTY		void 
// VT_I2		short 
// VT_I4		long 
// VT_R4		float 
// VT_R8		double 
// VT_CY		CY 
// VT_DATE		DATE 
// VT_BSTR		BSTR 
// VT_DISPATCH	LPDISPATCH 
// VT_ERROR		SCODE 
// VT_BOOL		BOOL 
// VT_VARIANT	VARIANT 
// VT_UNKNOWN	LPUNKNOWN

/*
������		����     
VT_EMPTY	ָʾδָ��ֵ    
VT_NULL		ָʾ��ֵ��������  SQL  �еĿ�ֵ��    
VT_I2		ָʾ  short  ����    
VT_I4		ָʾ  long  ����    
VT_R4   	ָʾ  float  ֵ    
VT_R8   	ָʾ  double  ֵ    
VT_CY   	ָʾ����ֵ    
VT_DATE   	ָʾ  DATE  ֵ    
VT_BSTR   	ָʾ  BSTR  �ַ���    
VT_DISPATCH	ָʾ  IDispatch  ָ��    
VT_ERROR   	ָʾ  SCODE   
VT_BOOL   	ָʾһ������ֵ    
VT_VARIANT 	ָʾ  VARIANTfar  ָ��    
VT_UNKNOWN 	ָʾ  IUnknown  ָ��    
VT_DECIMAL 	ָʾ  decimal  ֵ 
VT_UI1   	ָʾ  byte   
VT_UI2   	ָʾ  unsignedshort   
VT_UI4   	ָʾ  unsignedlong   
VT_I8   	ָʾ  64  λ����    
VT_UI8   	ָʾ  64  λ�޷�������    
VT_INT   	ָʾ����ֵ    
VT_UINT   	ָʾ  unsigned  ����ֵ    
VT_VOID   	ָʾ  C  ��ʽ  void   
VT_HRESULT  ָʾ  HRESULT   
VT_PTR   	ָʾָ������    
VT_SAFEARRAY   	ָʾ  SAFEARRAY   
VT_CARRAY   	ָʾ  C  ��ʽ����    
VT_USERDEFINED	ָʾ�û����������    
VT_LPSTR   		ָʾһ����  NULL  ��β���ַ���    
VT_LPWSTR   	ָʾ��  nullNothingnullptrnull   ���ã���  Visual Basic   ��Ϊ  Nothing ��   ��ֹ�Ŀ��ַ���    
VT_RECORD   	ָʾ�û����������    
VT_FILETIME   	ָʾ  FILETIME  ֵ    
VT_BLOB   		ָʾ�Գ���Ϊǰ׺���ֽ�    
VT_STREAM   	ָʾ�������������    
VT_STORAGE   	ָʾ����Ǵ洢������    
VT_STREAMED_OBJECT	ָʾ����������    
VT_STORED_OBJECT	ָʾ�洢��������    
VT_BLOB_OBJECT		ָʾ  Blob  ��������    
VT_CF   	ָʾ�������ʽ    
VT_CLSID   	ָʾ��ID   
VT_VECTOR   ָʾ�򵥵��Ѽ�������    
VT_ARRAY   	ָʾ  SAFEARRAY  ָ��    
VT_BYREF   	ָʾֵΪ���� 
��������������������������������
��Ȩ����������ΪCSDN������programmerES����ԭ�����£���ѭCC 4.0 BY-SA��ȨЭ�飬ת���븽��ԭ�ĳ������Ӽ���������
ԭ�����ӣ�https://blog.csdn.net/superlym2005/article/details/5838356
 */

#endif // !defined(AFX_STDAFX_H__BA14C3D0_C5E8_4D24_90B8_DED6A84AEF10__INCLUDED_)
