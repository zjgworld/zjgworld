
#include "stdafx.h"
#include "BasicClass.h"
#include "zjgBasic.h"

//////
using namespace ZJGName;

ExMethod::ExMethod()
{
}

ExMethod::ExMethod(const ExMethod& Other)
{
	*this = Other;
}


ExMethod::~ExMethod()
{
	this->mName.Empty();
	this->mLine.RemoveAll();
}

void ExMethod::Empty()
{
	this->mLine.RemoveAll();
	this->mName.Empty();
	this->m_DISPID.Empty();
	this->m_OutType.Empty();
	this->m_VARTYPE.Empty();
	this->m_InType.Empty();

}

BOOL ExMethod::IsEmpty()
{
	if( mLine.IsEmpty()		&& 
		mName.IsEmpty()		&&
		m_DISPID.IsEmpty()	&&
		m_OutType.IsEmpty()	&&
		m_InType.IsEmpty()	&&
		m_VARTYPE.IsEmpty()	) return true;;
	return false;
}


const ExMethod& ExMethod::operator=(const ExMethod& Other)//��ֵ
{
	this->mName		= Other.mName;
	this->m_DISPID	= Other.m_DISPID;
	this->m_OutType	= Other.m_OutType;
	this->m_InType	= Other.m_InType;
	this->m_VARTYPE	= Other.m_VARTYPE;

	this->mLine.RemoveAll();
	CString TemCStr;
	POSITION pPos;
	pPos = Other.mLine.GetHeadPosition();
	while(pPos)
	{
		TemCStr = Other.mLine.GetNext(pPos);
		mLine.AddTail(TemCStr);
	}
	return *this;
}
//**************************************************
// BOOL DecomposeMeth(CString strLine)
// ������CString strLine һ�����
//
// ����ֵ��CString ----����������ֵʱ�����Ǻ����������
//
//�ж��Ƿ��Ǻ����Ķ������,�ǣ����غ�����
//**************************************************/
//�����Ķ��������������ʽ
// 1.��.h�ļ��е���ʽ	a. float * put_Priority(long newValue)
//                      b. float * put_Priority(long& newValue)
//                      c. float * put_Priority(long* newValue)
//                      d. float get_Application()
// 2.��.cpp�ļ��е���ʽ a. LPDISPATCH Adjustments::GetApplication()	-----������
//                      b. float * Adjustments::GetItem(long Index)	-----������
//                      c. void put_Priority(long newValue)			-----��������
//						d. long GetCreator()						-----��������
// 3.�궨��             a. STDMETHOD(SetFirstPriority)()
//                      b. STDMETHOD(ModifyAppliesToRange)(Range * Range)

BOOL ExMethod::DecomposeMeth(CString strLine)		//�ֽ⺯��������䣬����д������
{
	CString LineCStr;
	LineCStr = strLine;
	LineCStr.TrimLeft();
	LineCStr.TrimRight();
	
	this->Empty();

	int iLeft;				//������
	int iRight;				//������
	int iSpace;				//�ո�
	int iColon;				//ð��
	int iDefine;			//�궨��
	int iAsterisk;			//�Ǻ�
	iLeft	= LineCStr.Find('(');
	iRight	= LineCStr.Find(')');
	iSpace	= LineCStr.Find(' ');
	iColon	= LineCStr.Find("::");
	iDefine	= LineCStr.Find("STDMETHOD");
	
	if(iDefine == 0 )								//�Ǻ궨�����
	{
		mName	= LineCStr.Mid(iLeft+1,iRight- iLeft - 1);
		//ȡ������������
		LineCStr= LineCStr.Mid(iRight+1);
		iLeft	= LineCStr.Find( '(' );			//�ڶ���С����
		iRight	= LineCStr.Find( ')' );			//�ڶ���С����
		if( (iLeft >= 0) && (iRight > iLeft))
		{
			m_InType = LineCStr.Mid(iLeft, (iRight - iLeft)+1);	//С�����ڵĲ��֣���������С����
		}
		this->m_OutType = "HRESULT";
	}
	else if((iSpace > 0) && (iLeft > iSpace) && (iRight > iLeft))
	{
		//�Ƿ�����ȡ������
		m_OutType	= LineCStr.Left(iLeft);		//������������ + [����+ :: + ] ������
		m_InType	= LineCStr.Mid(iLeft);		//( �������������� + ������ )
		//��ǰ�벿�ֽ��зֽ�
		iAsterisk	= m_OutType.ReverseFind('*');
		iSpace		= m_OutType.ReverseFind(' ');
		if(iAsterisk > iSpace)		//��'*'�ں���ʱ����'*'��Ϊ�ָ����������ÿո����ָ���
			iSpace = iAsterisk;	

		if (iColon < 0)									//������ĺ���
		{
			mName		= m_OutType.Mid(iSpace+1);		//������������֮��Ĳ��־��Ǻ�����
		}
		else if(iColon > iSpace &&  iColon < iLeft)		//����ĺ���ʱ
		{
			mName		= m_OutType.Mid(iColon+2);		//˫ð�ź���Ĳ����Ǻ�����
		}
		m_OutType	= m_OutType.Left(iSpace);			//������������
	}

	//ȡ�����������������������
	//��������ں����������С�����ڣ��ö��ŷָ����ӿں���ֻ����һ�����������	m_InType.TrimLeft();
	m_InType.TrimRight();
	iSpace = m_InType.ReverseFind(' ');
	iAsterisk = m_InType.ReverseFind('*');
	if(iSpace < iAsterisk) iSpace = iAsterisk;		//�ȳ���"*"��ʱ����"*"��Ϊ�ָ���
	if(iSpace > 0) m_InType = m_InType.Mid(1,iSpace);
	else	m_InType.Empty();

	iAsterisk = m_InType.ReverseFind('&');
	m_InType  = m_InType.Left(iAsterisk);
	iAsterisk = m_OutType.ReverseFind('&');
	m_InType  = m_OutType.Left(iAsterisk);

	mName.TrimLeft();
	mName.TrimRight();
	m_InType.TrimLeft();
	m_InType.TrimRight();
	m_OutType.TrimLeft();
	m_OutType.TrimRight();

	return !mName.IsEmpty();
}
//**************************************************
//
//
//
//
//
//
//**************************************************/
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
CString ExMethod::GetMethodType()			//���ݷ���(����)�ı�ǣ���ȡ��������ֵ��Ӧ����������
{
	CString TemNode;
	POSITION pPos;
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = mLine.GetNext(pPos);
		if(TemNode.Find("InvokeHelper(") >= 0)
		{
			this->DecomposeIvHp(TemNode);
			return this->m_OutType;
		}	
	}
	return "";
}

//**************************************************
//
//
//
//
//
//
//**************************************************/
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
CString ExMethod::GetID()			//ͨ���ֽ�InvokeHelper��������ȡ�����ӿ����ʹ���
{
	CString TemNode;
	POSITION pPos;
	int iPos1,iPos2;
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = mLine.GetNext(pPos);
		iPos1 = TemNode.Find("InvokeHelper(");
		if(iPos1 >= 0)
		{
			iPos1 += CString("InvokeHelper(").GetLength();		//ָ���Ƶ�С���ź���
			iPos2 = TemNode.Find(',');
			if(iPos2 >=0) TemNode = TemNode.Mid(iPos1,iPos2 - iPos1);
			else TemNode = "";
			return TemNode;
		}	
	}
	return "";
}

BOOL ExMethod::DecomposeIvHp( )			//�ֽ�InvokeHelper��䣬����д��Ӧ�Ĳ���
{
	CString TemNode;
	POSITION pPos;
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = mLine.GetNext(pPos);
		if(TemNode.Find("InvokeHelper") >= 0)
		{
			return DecomposeIvHp(TemNode);
		}	
	}
	return false;
}

BOOL ExMethod::DecomposeIvHp(CString strLine)			//�ֽ�InvokeHelper��䣬����д��Ӧ�Ĳ���
{
	int iPos1,iPos2;
	CString TemCStr;
	iPos1 = strLine.Find("InvokeHelper(");
	if(iPos1 < 0)	return false;					//λ�ò���ȷ��ֱ�ӷ���

	////�Ȼ�ȡInvokeHelper�����Ľӿ�����
	iPos1 += CString("InvokeHelper(").GetLength();	//ָ���Ƶ�С���ź���
	iPos2 = strLine.Find(',');						//��һ�����ŵ�λ��
	if(iPos2 < iPos1)	return false;				//λ�ò���ȷ��ֱ�ӷ���

	m_DISPID = strLine.Mid(iPos1,iPos2 - iPos1);	//�ӿ�����
	m_DISPID.TrimLeft();		
	m_DISPID.TrimRight();
		
	//������ӿڵķ����������ͣ��ڵڶ��������������֮��
	iPos1 = strLine.Find(',',iPos2+1);				//�ڶ������ŵ�λ��
	if(iPos1 < iPos2)	return false;				//λ�ò���ȷ��ֱ�ӷ���

	iPos2 = strLine.Find(',',iPos1+1);				//���������ŵ�λ��
	if(iPos2 < iPos1)	return false;				//λ�ò���ȷ��ֱ�ӷ���

	m_VARTYPE = strLine.Mid(iPos1+1,iPos2-iPos1-1);
	m_VARTYPE.TrimLeft();
	m_VARTYPE.TrimRight();

	if(m_VARTYPE.IsEmpty())
	{
		this->m_OutType.Empty();	
		return false;								//�ӿڵķ�����������Ϊ��
	}
	this->m_OutType = Get_VARENUM_Type(m_VARTYPE);
	return true;	
}

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
CString ExMethod::Get_VARENUM_Type(CString nVARENUM)			//��ȡVARENUM��Ӧ�ı�������
{
	if(nVARENUM == "VT_EMPTY")		return "void";
	if(nVARENUM == "VT_I2")			return "short";
	if(nVARENUM == "VT_I4")			return "long"; 
	if(nVARENUM == "VT_R4")			return "float"; 
	if(nVARENUM == "VT_R8")			return "double"; 
	if(nVARENUM == "VT_CY")			return "CY"; 
	if(nVARENUM == "VT_DATE")		return "DATE"; 
	if(nVARENUM == "VT_BSTR")		return "BSTR"; 
	if(nVARENUM == "VT_DISPATCH")	return "LPDISPATCH"; 
	if(nVARENUM == "VT_ERROR")		return "SCODE";
	if(nVARENUM == "VT_BOOL")		return "BOOL"; 
	if(nVARENUM == "VT_VARIANT")	return "VARIANT";
	if(nVARENUM == "VT_UNKNOWN")	return "LPUNKNOWN";
	return "";
}

CString ExMethod::Get_Type_VARENUM(CString nType)			//��ȡ�������Ͷ�ӦVARENUM��ֵ
{
	if(nType == "void")			return "VT_EMPTY";
	if(nType == "short")		return "VT_I2";
	if(nType == "long")			return "VT_I4"; 
	if(nType == "float")		return "VT_R4"; 
	if(nType == "double")		return "VT_R8"; 
	if(nType == "CY")			return "VT_CY"; 
	if(nType == "DATE")			return "VT_DATE"; 
	if(nType == "BSTR")			return "VT_BSTR"; 
	if(nType == "LPDISPATCH")	return "VT_DISPATCH"; 
	if(nType == "SCODE")		return "VT_ERROR";
	if(nType == "BOOL")			return "VT_BOOL"; 
	if(nType == "VARIANT")		return "VT_VARIANT";
	if(nType == "LPUNKNOWN")	return "VT_UNKNOWN";
	return "";
}
//**************************************************
// ����   Modify_Method(CList<CString,CString&> * pList)
// ���룺 ExMethod* pMethod  ����
//        
// ���أ� ����У����ض�ӦPUT�ͷ�����λ�á����򷵻ؼ�
//  
//
//**************************************************/
void ExMethod::Modify_Method(CList<CString,CString&> * pList)	//�Ե�ǰ����Ϊ�ο��࣬����pList��ָ����������޸�
{
	CString NodeCStr;

	POSITION pPos,pOldPos;
	pPos = pList->GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		NodeCStr = pList->GetNext(pPos);

		NodeCStr.Replace("RHS ",this->m_OutType+ ' ');
		NodeCStr.Replace(", ,",',' + this->m_VARTYPE + ',');
		pList->SetAt(pOldPos,NodeCStr);
	}
}
//**************************************************
// ����   Modify_Method_PUT(CList<CString,CString&> * pList)
// ���룺 ExMethod* pMethod  ����
//        
// ���أ� ����У����ض�ӦPUT�ͷ�����λ�á����򷵻ؼ�
//  
//
//**************************************************/
// ��PUT_�ຯ�������룬����GET_�ຯ���������
// �漰���������֣�a.�����������ͣ�b.�ӿں����еķ�������
BOOL ExMethod::Modify_Method_PUT(CList<CString,CString&> * pList)	//�Ե�ǰ����Ϊ�ο��࣬����pList��ָ����������޸�
{
	BOOL bRtn = true;
	CString NodeCStr,OutTypeCStr,VARTYPECStr;
	VARTYPECStr = this->Get_Type_VARENUM(m_InType);				//��PUT_�ຯ�������룬����GET_�ຯ�������
	if(VARTYPECStr.IsEmpty())
	{
		VARTYPECStr = "VT_I4";
		bRtn = false;
	}	
	POSITION pPos,pOldPos;
	pPos = pList->GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		NodeCStr = pList->GetNext(pPos);
		NodeCStr.Replace("RHS ",this->m_InType+ ' ');			//��PUT_�ຯ�������룬����GET_�ຯ�������
		NodeCStr.Replace(", ,",',' + VARTYPECStr + ',');
		pList->SetAt(pOldPos,NodeCStr);
	}
	return bRtn;
}

//////
ExClass::ExClass()
{

}
ExClass::ExClass(const ExClass& Other)
{
	*this = Other;
}
ExClass::~ExClass()
{
	this->mLine.RemoveAll();
	this->mName.Empty();
}

void ExClass::Empty()
{
	this->mLine.RemoveAll();
	this->mName.Empty();
	this->mBase.Empty();
}

const ExClass& ExClass::operator=(const ExClass& Other)//��ֵ
{
	this->mName = Other.mName;
	this->mBase = Other.mBase;		//������
	this->mLine.RemoveAll();
	ExMethod MethodNode;
	POSITION pPos;
	pPos = Other.mLine.GetHeadPosition();
	while(pPos)
	{
		MethodNode = Other.mLine.GetNext(pPos);
		mLine.AddTail(MethodNode);
	}
	return *this;
}

BOOL ExClass::IsEmpty()
{
	if(!(mName.IsEmpty()) ) return false;
	if(!(mLine.IsEmpty()) ) return false;
	return true;
}

CString ExClass::GetType(CString dwDispID)					//���ݷ���(����)�ı�ǣ���ȡ��������ֵ��Ӧ����������
{
	ExMethod TemNode;
	POSITION pPos;
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = mLine.GetNext(pPos);
		if(TemNode.GetID() == dwDispID)
		{
			return TemNode.GetMethodType();
		}
	}
	return "";
}

CString ExClass::Get_Type_VARENUM(CString nType)			//��ȡ�������Ͷ�ӦVARENUM��ֵ
{
	ExMethod TemNode;
	return TemNode.Get_Type_VARENUM(nType);

}
CString ExClass::Get_VARENUM_Type(CString nVARENUM)			//��ȡVARENUM��Ӧ�ı�������
{
	ExMethod TemNode;
	return TemNode.Get_VARENUM_Type(nVARENUM);
}
//**************************************************
// ����   Find_PUT(ExMethod* pMethod)
// ���룺 ExMethod* pMethod  ����
//        
// ���أ� ����У����ض�ӦPUT�ͷ�����λ�á����򷵻ؼ�
//  
//
//**************************************************/
POSITION ExClass::Find_PUT(CString dwDispID)				//����dwDispID�����Ҷ�Ӧ�Ľӿڵ�PUT����
{
	ExMethod TemNode;
	POSITION pPos,pOldPos;
	CString NameCStr;
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemNode = mLine.GetNext(pPos);
		NameCStr = TemNode.GetName().Left(3);
		NameCStr.MakeLower();
		if(	TemNode.GetDISPID() == dwDispID &&
			NameCStr == "put" )
			return pOldPos;
	}
	return pPos;
}
//**************************************************
// ����   FindDispID(CString dwDispID,CList<ExMethod,ExMethod&>* pLine)
// ���룺 dwDispID��InvokeHelper�����Ľӿ�����  
//        pLine����Ų��ҵ��ķ���������ָ��
// ���أ� ����У������档���򷵻ؼ�
//  
//
//**************************************************/
void ExClass::FindDispID(CString dwDispID,CList<ExMethod,ExMethod&>* pLine)	//����InvokeHelper�����Ľӿ����ͣ����Ҷ�Ӧ�ķ���
{
	pLine->RemoveAll();

	ExMethod TemNode;
	POSITION pPos;
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = mLine.GetNext(pPos);
		if(TemNode.GetDISPID() == dwDispID)
		{
			pLine->AddTail(TemNode);
		}
	}
	return ;
}

int ExClass::Modify_Method(ExMethod * pMethod)
{
	POSITION pPos,pOldPos;
	CString TemCStr;
	CString str_VARTYPE;
	CString str_Type;
	//��ȡ��Ӧ�ӿ����͵ķ���ֵ�ı�ʶ
	
	str_Type	= GetType(pMethod->GetDISPID());			//�ӿ����͵ķ���ֵ�ı�ʶ
	if(str_Type.IsEmpty()) return FILE_C_LESSID_ERR;		//��ǰ����ȱ����ͬID�ķ�����ֱ�ӷ��ز�����
	str_VARTYPE	= Get_Type_VARENUM(str_Type);				//�ӿ����͵ķ���ֵ�ı�ʶ
	if(str_VARTYPE.IsEmpty()) return FILE_VARTYPE_ERR;		//�޶�Ӧ���������ͣ�ֱ�ӷ��ز�����
		
	str_Type	= str_Type + " ";
	pPos = pMethod->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemCStr = pMethod->mLine.GetNext(pPos);			
		TemCStr.Replace("RHS ",str_Type);				//���ġ�RHS ��
		TemCStr.Replace(", ,",',' + str_VARTYPE + ',');	//�ڡ�, , ��������ӽӿ����͵ķ���ֵ�ı�ʶ
		pMethod->mLine.SetAt(pOldPos,TemCStr);
	}
	return FILE_FINISHED;
}

int ExClass::Modify_Method(CList<CString,CString&> * pList,CString * pstrMethod)
{
	//������ͬ�Ľӿں�������׼���滻��
	CString str_VARTYPE;
	CString str_Type;
	
	int iPos = pstrMethod->Find('(');	//�Ƿ��ǽӿڷ���
	int iColon = pstrMethod->Find(',');	//���ŵ�λ��
	CString TemCStr;
	POSITION pPos,pOldPos;
	if(iColon >= 0)
	{
		TemCStr = pstrMethod->Mid(iPos+1, iColon - iPos -1);
		str_VARTYPE	= GetType(TemCStr);							//�ӿ����͵ķ���ֵ�ı�ʶ
		if(str_VARTYPE.IsEmpty()) return FILE_C_LESSID_ERR;		//ȱ����ͬID�ķ���
		
		str_Type	=this->Get_VARENUM_Type(str_VARTYPE);		//����ֵ����������
		if(str_Type.IsEmpty()) return FILE_VARTYPE_ERR;			//�޶�Ӧ���������ͣ�ֱ�ӷ��ز�����
		
		str_Type	= str_Type + " ";
		//�������е��н����޸�
		pPos = pList->GetHeadPosition();
		while(pPos)
		{
			pOldPos = pPos;
			TemCStr = pList->GetNext(pPos);			
			TemCStr.Replace("RHS ",str_Type);				//���ġ�RHS ��
			TemCStr.Replace(", ,",',' + str_VARTYPE + ',');	//�ڡ�, , ��������ӽӿ����͵ķ���ֵ�ı�ʶ
			pList->SetAt(pOldPos,TemCStr);
		}
	}
	return FILE_FINISHED;
}

void ExClass::DecomposeClassName(CString LineCStr)			//�ֽ��ඨ����䣬����д��Ӧ�Ĳ���
{
	//ȡ����
	int iColon = LineCStr.Find(':');
	if(iColon > 0) this->mName = LineCStr.Mid(6,iColon - 6);
	else mName = LineCStr.Mid(6);
	mName.TrimRight();
}
////////////ExClassLine

ExClassLine::ExClassLine()
{

}

ExClassLine::~ExClassLine()
{
	Empty();
}

void ExClassLine::Empty()
{
	this->mLine.RemoveAll();
	this->mFileName.Empty();
}
///*********************************
//
//
//
//
//
//  ����װ��EXCEL�����ļ�
//*********************************
int ExClassLine::Load_H_File(CString FileName)			//װ��.h��ʽ�ļ�
{
	CString ReadLineCStr,LineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeRead)) return false;

	ExClass ExClassNode;
	ExMethod ExMethodNode;
	this->mLine.RemoveAll();
	this->mFileName = FileName;

	BOOL bClass;
	int iColon;				//ð��
	int ibrackets;			//������
	CString strClassName,TemType;;

	bClass = false;
	ibrackets = 0;
	ExClassNode.Empty();
	ExMethodNode.Empty();

	while(myFile.ReadString(ReadLineCStr))
	{
		LineCStr = ReadLineCStr;
		LineCStr.TrimLeft();
		
		if(ibrackets >= 2)						//�����ŵ��вŴ���
		{
			ExMethodNode.mLine.AddTail(ReadLineCStr);
			//�����еķ�����
			if(LineCStr.Find('{') >= 0)	ibrackets++;			
			if(LineCStr.Find('}') >= 0)
			{
				ibrackets--;
			}

			if(ibrackets < 2)						//���˷����Ĵ�����
			{
				TemType = ExMethodNode.GetOutType();
				ExMethodNode.DecomposeIvHp();
				ExMethodNode.SetOutType(TemType);				
				//������������ӵ���������
				ExClassNode.mLine.AddTail(ExMethodNode);
				ExMethodNode.Empty();
			}
		}
		else if(ibrackets == 1)						//�����ŵ��вŴ���
		{
			//�����еķ������д���
			if(LineCStr.Find('{') >= 0)	ibrackets++;

			if(ExMethodNode.GetName().IsEmpty())				//ǰһ��䲻�Ƕ�����䣬ֻ�����жϱ�����Ƿ��Ƕ������
			{
				ExMethodNode.DecomposeMeth(LineCStr);			//�ֽ���䣬���ж��Ƿ��Ǻ����������
			}
			if(ExMethodNode.GetName().IsEmpty() == NULL)		//���ж���䣬����һ�ж������˳��ģ������Ժϲ�
			{													//1�������ǰ����Ƕ�����䣬����ӡ�
				ExMethodNode.mLine.AddTail(ReadLineCStr);		//2��ǰһ����Ƕ������,Ҳ���
			}
			
			if(LineCStr.Find('}') >= 0)		ibrackets--;
			if(ibrackets < 1)
			{
				if(ExClassNode.IsEmpty() == NULL)
				{
					this->mLine.AddTail(ExClassNode);
				}
				ExClassNode.Empty();
			}
		}
		else if(bClass)											//�Ѿ������ඨ�����
		{
			ibrackets = 0;							//�����Ž�Ҫ��ʼ

			if(LineCStr.Find('{') >= 0)	ibrackets++;
			if(LineCStr.Find('}') >= 0)
			{	ibrackets--;
				if(ibrackets < 0)					//�����Ų��ɶԣ����ִ���
				{
					myFile.Close();	
					return FILE_ERROR_ERR;	
				}
			}
		}
		else if(LineCStr.Left(6).Compare("class ") == 0 )					//����
		{
			//ȡ����,�ඨ���ж�����ʽ��1��class ****
			//                         2��class **** : *******
			iColon = LineCStr.Find(':');
			if(iColon > 0) strClassName = LineCStr.Mid(6,iColon - 6);
			else strClassName = LineCStr.Mid(6);
			strClassName.TrimRight();
			strClassName.TrimLeft();

			bClass = true;
			ibrackets = 0;							//�����Ž�Ҫ��ʼ
			ExClassNode.Empty();
			ExClassNode.SetName(strClassName);

			if(LineCStr.Find('{') >= 0)	ibrackets++;
			if(LineCStr.Find('}') >= 0)
			{	ibrackets--;
				if(ibrackets < 0)					//�����Ų��ɶԣ����ִ���
				{
					myFile.Close();	
					return FILE_ERROR_ERR;	
				}
			}
		}
		else
		{
			//������ⲿ��������
			if(LineCStr.Find('{') >= 0 )	ibrackets++;

			if(LineCStr.Find('}') >= 0)		ibrackets--;
		}
	}
	myFile.Close();	
	if(ibrackets == 0)	return FILE_FINISHED;		//�����ųɶԣ���������
	else return FILE_ERROR_ERR;						//�����Ų��ɶԣ����ִ���
}

BOOL ExClassLine::LoadCPPFile(CString FileName)			//װ��.CPP��ʽ�ļ�
{
	CString ReadLineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeRead)) return false;
	int iSpace;				//�ո�
	int iLeft;				//������
	int iRight;				//������
	int iColon;				//ð��
	int ibrackets;			//������
	int iAnnotation;		//ע�ͺ�
	int iPos, iLineLen;		//һ�еĳ���
	CString ClassName;
	CString MethodName;
	CString SurplusCStr;

	ExClass ExClassNode;
	ExMethod ExMethodNode;
	this->mLine.RemoveAll();
	this->mFileName = FileName;

	ibrackets = 0;
	while(myFile.ReadString(ReadLineCStr))
	{
		ReadLineCStr.TrimLeft();
		ReadLineCStr.TrimRight();
		
		iLineLen = ReadLineCStr.GetLength();						//һ�еĳ���
		if(iLineLen <= 0) continue;									//���в�����
		if(iLineLen >= 2)
		{
			if(ReadLineCStr.GetAt(0) == '/' && ReadLineCStr.GetAt(1) == '/')
				continue;		//ע���в�����
			
			//ע�Ϳ鲻����
			iPos = ReadLineCStr.Find("/*");
			if(iPos >= 0)
			{
				iAnnotation++;
				SurplusCStr = ReadLineCStr.Left(iPos);				//һ����ʣ��Ĳ��֣��ڡ�/*��֮ǰ�Ĳ���
				ReadLineCStr= ReadLineCStr.Mid(iPos);				//һ����ʣ��Ĳ��֣��ڡ�/*��֮��Ĳ���
				do
				{
					iPos = ReadLineCStr.Find("*/");
					if(iPos >= 0) iAnnotation--;
				}while(iAnnotation > 0 && myFile.ReadString(ReadLineCStr));
			}
			if(SurplusCStr.GetLength() > 0) ReadLineCStr = SurplusCStr;
		}
		//�������в��ܶ����� { ClassName = ReadLineCStr.Mid(iSpace+1,/*�ֲ�ע��*/iColon - iSpace -1);} �����д���
		//����������һ�н��д���
		iSpace	= ReadLineCStr.Find(" ",	0);
		iLeft	= ReadLineCStr.Find("(",	0);
		iRight	= ReadLineCStr.Find(")",	0);
		iColon	= ReadLineCStr.Find("::",	0);

		if( (iSpace > 0) &&
			(iLeft < iRight) &&
			(iSpace < iColon) &&
			(iColon < iLeft) )
		{
			ClassName = ReadLineCStr.Mid(iSpace+1,iColon - iSpace -1);	//����
			MethodName= ReadLineCStr.Mid(iColon+2,iLeft - iColon -2);	//������
			//�µķ�����ʼ
			if(!ExMethodNode.IsEmpty())
			{
				ExMethodNode.GetMethodType();
				ExClassNode.mLine.AddTail(ExMethodNode);
			}
			ExMethodNode.Empty();
			ExMethodNode.SetName(MethodName);
			ExMethodNode.mLine.AddTail(ReadLineCStr);
			ibrackets = 0;
			if(ExClassNode.GetName() != ClassName)
			{
				//���࿪ʼ��
				//�����е��࣬��ӵ������У���������
				if(!ExClassNode.IsEmpty())	this->mLine.AddTail(ExClassNode);
				ExClassNode.Empty();
				ExClassNode.SetName(ClassName);
			}
			continue;
		}
		if(ReadLineCStr.Find("{",0) >= 0) ibrackets++;
		if(ReadLineCStr.Find("}",0) >= 0)
		{
			ibrackets--;
			if(ibrackets == 0) 
				ExMethodNode.mLine.AddTail(ReadLineCStr);
		}
		if(ibrackets <= 0) continue;		//���ڴ������ڵ��в�����
		ExMethodNode.mLine.AddTail(ReadLineCStr);
		
	}
	{
		//�������ļ�
		if(!ExMethodNode.IsEmpty())
		{
			ExMethodNode.GetMethodType();
			ExClassNode.mLine.AddTail(ExMethodNode);
		}
		if(!ExClassNode.IsEmpty())	this->mLine.AddTail(ExClassNode);
	}
	myFile.Close();

	return true;
}

///****************************************************
// \�������ƣ�Modify_H_File_PUT(CString FileName)
// \�������: FileName				----���޸ĺ����������ļ���
//		
// \����ֵ:   int					----����ֵΪ�������
//
// \˵��:	  1. �޸�.h��ʽ�ļ��з���ֵΪRHS�����ķ������͡��硰RHS get_Creator()��
//			  2. �޸ķ�������ΪRHS�ķ���ֵ�����͡��硰RHS result;��
//			  3. �ӿں�����InvokeHelper()���е�vtRet��ֵ���硰InvokeHelper(0x95, DISPATCH_PROPERTYGET, , (void*)&result, nullptr);��
//1����ͬһ�ļ��ڣ�����ͬ�ӿڵ� "put_" ������ȡ�����������������ͣ���Ϊ��Ӧ "get_" �����ķ������ͣ���������Ӧ���޸�
//2��һ���ļ�ֻ����һ���ࡣ
///****************************************************
int ExClassLine::Modify_H_File_PUT(CString OldFileName,CString NewFileName)
{
	//˼·��1.��װ���ࡣ
	//		2.�ٶ��ļ�����ѭ����A��������RHS������ʱ������ Find_Method_PUT(ExMethod* pMethod);
	//							B���ҵ���Ӧ�ġ�PUT_������ʱ�������޸ġ��Ҳ���ʱֱ�Ӹ��ƣ������ô������
	//							C. ������ǡ�RHS������ʱ��ֱ�Ӹ��ơ�
	//      
	POSITION pPos;
	CString TemCStr,ClassCStr;
	ExClass TemExClass,BaseExClass;
	ExMethod TemMethod,BaseMethod;
	
	this->Load_H_File(OldFileName);
	BaseExClass = this->mLine.GetHead();	//��Ϊ��׼��
	CString ReadLineCStr;
	CString LineCStr;
//	CString NewFileName;
	CStdioFile myOldFile;				//ԭ�ļ�
	CStdioFile myNewFile;				//���ļ�
	int iError = FILE_FINISHED;
	int iRtn  = FILE_FINISHED;

	int iPos;
	iPos = mFileName.Find(".h");						//����.h
	if(iPos <= 0) iPos = mFileName.Find(".H");		
	if(iPos <= 0) return FILE_NAME_ERROR;				//�ļ�������ȷ 

	if(!myOldFile.Open(OldFileName,CFile::modeRead)) return FILE_READ_ERROR;							//���ļ�����
	if(!myNewFile.Open(NewFileName,CFile::modeWrite | CFile::modeCreate)) return FILE_WRITE_ERROR;		//д�ļ�����
	
	int ibrackets = 0;					//������
	int iClass = 0;						//��Ĳ���
	int iAnnotation;					//ע�ͺ�
	BOOL bClass = false;				//������ı�־
	BOOL bWrite = true;					//���ԭ�����������ռ䶨�壬��Ϊ�١�

	TemExClass.Empty();
	this->Empty();
	this->mFileName = mFileName;
	bClass = false;

	while(myOldFile.ReadString(ReadLineCStr))
	{
		LineCStr = ReadLineCStr;
		LineCStr.TrimLeft();
		iAnnotation = LineCStr.Find("//");
		if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
		LineCStr.TrimRight();
		ClassCStr.Empty();

		if(!bClass)	//ǰ�治���ඨ�����
		{
			if(LineCStr.Left(6).Compare("class ") == 0)
			{
				bClass = true;//���ඨ�����
				ClassCStr = ";\n" + LineCStr;
			}
			else 
			{
				myNewFile.WriteString(ReadLineCStr +"\n");//�����ඨ�����,ֱ�Ӹ���
				continue;
			}
		}
		if(LineCStr.Find(';') > 0)						//
		{
			//��ǰ������ඨ�壬�����зֺţ�����������ˣ��磺class ExMethod;
			bClass = false;								//�������
			myNewFile.WriteString(ReadLineCStr +"\n");	//�������,ֱ�Ӹ���
			continue;
		}
		//ȡ����
		ClassCStr = ClassCStr + LineCStr;			//�ඨ����仹û�н��������Լ��ϵ�ǰ���
		TemExClass.DecomposeClassName(ClassCStr);
		if(TemExClass.GetName().IsEmpty())continue;	//û���������� class
													//             ****{
		do
		{
			LineCStr = ReadLineCStr;
			LineCStr.TrimLeft();
			iAnnotation = LineCStr.Find("//");
			if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
			LineCStr.TrimRight();
		
			//�����������ڽ���ѭ����
			//��׼����
			if(LineCStr.Find('{') < 0)
			{
				//��ǰ��䲻�ڴ�������ʱ��ֱ�Ӹ���
				myNewFile.WriteString(ReadLineCStr +"\n");
				continue;
			}
			if(ibrackets < 1)
			{
				myNewFile.WriteString(ReadLineCStr +"\n");		//��ǰ����Ǵ���{����ֱ�Ӹ���
			}
			ibrackets = 1;

			//*********************************
			//����������������ѭ��
			//*********************************
			TemMethod.Empty();

			//��ȡ��׼����
			while(myOldFile.ReadString(ReadLineCStr))
			{
				LineCStr = ReadLineCStr;
				LineCStr.TrimLeft();
				iAnnotation = LineCStr.Find("//");
				if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
				LineCStr.TrimRight();

				if(TemMethod.DecomposeMeth(LineCStr) == NULL)
				{
					//��ǰ��䲻�ڷ����Ĵ�������ʱ��ֱ�Ӹ���
					myNewFile.WriteString(ReadLineCStr +"\n");
					continue;
				}
				//*********************************
				//�����������ڷ�������ѭ��
				//*********************************
				do		//�ڷ�������ѭ��
				{
					LineCStr = ReadLineCStr;
					LineCStr.TrimLeft();
					iAnnotation = LineCStr.Find("//");
					if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
					LineCStr.TrimRight();
								
					TemMethod.mLine.AddTail(ReadLineCStr);
					if(LineCStr.Find('{') >= 0) ibrackets++;
					if(LineCStr.Find('}') >= 0)
					{
						ibrackets--;
						if(ibrackets == 1) break;			//��������ֻ��һ��ʱ,�Ѳ��ڷ�������
					}
				}while(myOldFile.ReadString(ReadLineCStr));	//�����ν���
				TemMethod.SetDISPID(TemMethod.GetID());		//���÷�����ID��
				//*********************************
				//		�����������ڷ�������ѭ�����ڲ�doѭ������
				//*********************************

				//�Է������д���
				if(TemMethod.GetOutType().Compare("RHS") == 0 )//�Ǵ��޸ĵķ���
				{
					//����Ҫ�޸ĵķ���
					pPos = BaseExClass.Find_PUT(TemMethod.GetDISPID());
					if(pPos)
					{
						BaseMethod = BaseExClass.mLine.GetAt(pPos);	//��׼����
						if(BaseMethod.Modify_Method_PUT(&TemMethod.mLine) == false)
						{//���÷��ش�������
							iRtn = iRtn | FILE_VARTYPE_ERR;		//û�ж�Ӧ����������
						}
					}
					else iRtn = iRtn | FILE_C_LESSPUT_ERR;		//����ȱ�ٳɶԵ�ID����
				}
	
				TemExClass.mLine.AddTail(TemMethod);
				//�������ڵ���д�뵽�µ��ļ�
				pPos = TemMethod.mLine.GetHeadPosition();
				while(pPos)
				{
					TemCStr = TemMethod.mLine.GetNext(pPos);
					myNewFile.WriteString(TemCStr +"\n");
				}
				//*********************************
				//		�����������ڷ�������ѭ�����ڲ�while(pPos)
				//*********************************
			}
	
			//*********************************
			//		����������������ѭ�������ڵ����while(myOldFile.ReadString(ReadLineCStr))
			//*********************************
		}while(myOldFile.ReadString(ReadLineCStr));	
		//*********************************
		//		�����������ڽ���ѭ�������doѭ�����˽���
		//*********************************
	}
	//*********************************
	//		�������������ļ���ѭ�����ļ��ڵ�ѭ��while(myOldFile.ReadString(ReadLineCStr))
	//*********************************

	myOldFile.Close();
	myNewFile.Close();

	return iRtn;
}


///****************************************************
// \�������ƣ�Modi_Method(ExMethod* pModiNode)
// \�������: ExMethod* pModiNode	----���޸ĺ�����ָ��
//		
// \����ֵ:   int					----����ֵΪ�������
//
// \˵��:	  1. �޸�.h��ʽ�ļ��з���ֵΪRHS�����ķ�������
//			  2. �޸ķ�������ΪRHS�ķ���ֵ������
//			  3. �ӿں�����InvokeHelper()���е�vtRet��ֵ
// ���ݽӿ�dwDispID��ͬ,�����޸ġ���ͬ�࣬��ͬ���������ͬ�ӿڵ�˳��
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
//
///****************************************************
int ExClassLine::Modify_H_File_A(CString OldFName,CString NewFName)
{
	//˼·��1.�ȶ��ļ����н���ѭ���������⣬ֱ�ӽ��и��ƣ����޸�
	//		2.�����ඨ��ʱ��A. ��ȡͬ��������࣬�Ա㰴ͬ���������Ϊ����ʱ���޸ġ�
	//                      B. ���������������ѭ�������Ƿ���ʱ��ֱ�Ӹ��ƣ����޸ġ�
	//						C. �Ƿ���ʱ���ж��ǲ�����Ҫ�޸��޸ĵķ���������ǣ����÷����޸ĺ���
	//                                          a. ����вο���ʱ�����òο����޸ĺ��������޸�
    //                                          b. û�вο��࣬����ͬ�ķ�������Ȼ����òο������޸ĺ��������޸�
	//                                          c. û����ͬ�ķ�����������ͬ�ӿڱ�ŵķ��������޸�
	//                      D. �����ڲ�������Ҫ�����޸ĵ��ֱ࣬�Ӹ��ơ�
	//      
	CString ReadLineCStr;
	CString LineCStr;
	CString NewFileName;
	CStdioFile myOldFile;				//ԭ�ļ�
	CStdioFile myNewFile;				//���ļ�
	int iError = FILE_FINISHED;
	int iRtn  = FILE_FINISHED;

	int iPos;
	iPos = OldFName.Find(".h");						//����.h
	if(iPos <= 0) iPos = OldFName.Find(".H");		
	if(iPos <= 0) return FILE_NAME_ERROR;														//�ļ�������ȷ 
	if(!myOldFile.Open(OldFName,CFile::modeRead)) return FILE_READ_ERROR;						//���ļ�����
	if(!myNewFile.Open(NewFName,CFile::modeWrite | CFile::modeCreate)) return FILE_WRITE_ERROR;	//д�ļ�����
	
	int ibrackets = 0;					//������
	int iClass = 0;						//��Ĳ���
	int iAnnotation;					//ע�ͺ�

	POSITION pPos;
	CString TemCStr;
	ExClass TemExClass,BaseExClass;
	ExMethod TemMethod;
	BOOL bFlag;

	TemExClass.Empty();
	while(myOldFile.ReadString(ReadLineCStr))
	{
		LineCStr = ReadLineCStr;
		LineCStr.TrimLeft();
		iAnnotation = LineCStr.Find("//");
		if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
		LineCStr.TrimRight();

		if(LineCStr.Left(6).Compare("class ") != 0 )	//�����ඨ�����,ֱ�Ӹ���
		{
			//������ʱ��ֱ�Ӹ���
			myNewFile.WriteString(ReadLineCStr +"\n");
			continue;
		}
		if(LineCStr.Find(';') > 0)						//
		{
			//��ǰ������ඨ�壬�����зֺţ�����������ˣ��磺class ExMethod;
			myNewFile.WriteString(ReadLineCStr +"\n");
			continue;
		}
		//ȡ����
		TemExClass.DecomposeClassName(LineCStr);
		if(TemExClass.GetName().IsEmpty()) continue;
		do
		{
			LineCStr = ReadLineCStr;
			LineCStr.TrimLeft();
			iAnnotation = LineCStr.Find("//");
			if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
			LineCStr.TrimRight();
		
			//�����������ڽ���ѭ����
			//��׼����
			if(LineCStr.Find('{') < 0)
			{
				//��ǰ��䲻�ڴ�������ʱ��ֱ�Ӹ���
				myNewFile.WriteString(ReadLineCStr +"\n");
				continue;
			}

			ibrackets = 1;
			myNewFile.WriteString(ReadLineCStr +"\n");	//д�뺬��{������

			//*********************************
			//����������������ѭ��
			//*********************************
			TemMethod.Empty();

			//��ȡ��׼����
			pPos = this->Find_A(TemExClass.GetName());
			if(pPos)  BaseExClass = this->mLine.GetAt(pPos);
			else BaseExClass.Empty();

			while(myOldFile.ReadString(ReadLineCStr))
			{
				LineCStr = ReadLineCStr;
				LineCStr.TrimLeft();
				iAnnotation = LineCStr.Find("//");
				if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
				LineCStr.TrimRight();

				if(TemMethod.DecomposeMeth(LineCStr) == NULL)
				{
					//��ǰ��䲻�ڷ����Ĵ�������ʱ��ֱ�Ӹ���
					myNewFile.WriteString(ReadLineCStr +"\n");
					continue;
				}
				//*********************************
				//�����������ڷ�������ѭ��
				//*********************************

				do		//�ڷ�������ѭ��
				{
					LineCStr = ReadLineCStr;
					LineCStr.TrimLeft();
					iAnnotation = LineCStr.Find("//");
					if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
					LineCStr.TrimRight();
								
					TemMethod.mLine.AddTail(ReadLineCStr);
					if(LineCStr.Find('{') >= 0) ibrackets++;
					if(LineCStr.Find('}') >= 0)
					{
						ibrackets--;
						if(ibrackets == 1) break;			//��������ֻ��һ��ʱ,�Ѳ��ڷ�������
					}
				}while(myOldFile.ReadString(ReadLineCStr));	//�����ν���
				TemMethod.SetDISPID(TemMethod.GetID());
				//*********************************
				//		�����������ڷ�������ѭ�����ڲ�doѭ������
				//*********************************

				//�Է������д���
				if(TemMethod.GetOutType().Compare("RHS") == 0 )//�Ǵ��޸ĵķ���
				{
					//����Ҫ�޸ĵķ���
					bFlag = BaseExClass.Modify_Method(&TemMethod);
					iRtn = iRtn | bFlag ;
					if(bFlag == FILE_C_LESSID_ERR )			//��׼����û�ж�Ӧ�ķ���
					{
						bFlag = this->Modi_Method_Name(&TemMethod);
						iRtn = iRtn | bFlag ;
						if(bFlag == FILE_LN_LESSID_ERR )	//������û�ж�Ӧ�ĺ������ķ���
						{
							bFlag = this->Modi_Method_ID(&TemMethod);
							iRtn = iRtn | bFlag ;
						}
					}
				}
				TemExClass.mLine.AddTail(TemMethod);
				//�������ڵ���д�뵽�µ��ļ�
				pPos = TemMethod.mLine.GetHeadPosition();
				while(pPos)
				{
					TemCStr = TemMethod.mLine.GetNext(pPos);
					myNewFile.WriteString(TemCStr +"\n");
				}
				//*********************************
				//		�����������ڷ�������ѭ�����ڲ�while(pPos)
				//*********************************
			}
			//*********************************
			//		����������������ѭ�������while(myOldFile.ReadString(ReadLineCStr))
			//*********************************
		}while(myOldFile.ReadString(ReadLineCStr));	
		//*********************************
		//		�����������ڽ���ѭ�������doѭ�����˽���
		//*********************************
	}
	myOldFile.Close();				//ԭ�ļ�
	myNewFile.Close();				//���ļ�
	return iRtn;
}

///****************************************************
// \�������ƣ�Modi_Method_Name(ExMethod* pModiNode)
// \�������: ExMethod* pModiNode	----���޸ĺ�����ָ��
//		
// \����ֵ:   int					----����ֵΪ�������
//
// \˵��:	  1. �޸�.h��ʽ�ļ��з���ֵΪRHS�����ķ�������
//			  2. �޸ķ�������ΪRHS�ķ���ֵ������
//			  3. �ӿں�����InvokeHelper()���е�vtRet��ֵ
// ���ݺ�������ͬ�����ƣ��ӿ�dwDispID��ͬ
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
//
///****************************************************
int ExClassLine::Modi_Method_Name(ExMethod* pMethod)		//���ݺ�������ͬ�����ƣ��ӿ�dwDispID��ͬ��ԭ������޸�
{
	ExMethod ExMethodNode;
	ExClass  ExClassNode;
	POSITION pPos;
	POSITION pMethodPos;
	CList<ExMethod,ExMethod&> TemLine;				//���ڴ����ͬDISPID����������

	int iPos;
	CString OldName_1,OldName_2,NewNameCStr;
	//�ҵ���ͬDISPID�ķ�����Ȼ��ȽϷ������������������ͬ�����޸�
	//�������� get_Application() �� GetApplication()
	//
	OldName_1 = pMethod->GetName();
	OldName_1.MakeUpper();
	iPos = OldName_1.Find('_');
	if(iPos > 0)
	{
		OldName_2 = OldName_1.Left(iPos) + OldName_1.Mid(iPos+1);
	}
	else OldName_2.Empty();

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		ExClassNode = mLine.GetNext(pPos);
		ExClassNode.FindDispID(pMethod->GetDISPID(),&TemLine);	//��ȡ��ͬID�ķ���
		//�ȽϷ��������ƣ��Ƿ�����
		pMethodPos = TemLine.GetHeadPosition();					//���ڴ����ͬDISPID����������
		while(pMethodPos)
		{
			ExMethodNode = TemLine.GetNext(pMethodPos);
			NewNameCStr = ExMethodNode.GetName();
			NewNameCStr.MakeUpper();
			if(	NewNameCStr.CompareNoCase(OldName_1) == 0 || 
				NewNameCStr.CompareNoCase(OldName_2) == 0 ) 
			{
				//���ƣ����Խ�������׼�������޸ġ�
				ExMethodNode.Modify_Method(&pMethod->mLine);
				return FILE_FINISHED;
			}
		}
	}
	return FILE_LN_LESSID_ERR;	//�����������ȱ����ͬID�ķ���
}
///****************************************************
// \�������ƣ�Modi_Method_ID(ExMethod* pModiNode)
// \�������: ExMethod* pModiNode	----���޸ĺ�����ָ��
//		
// \����ֵ:   int					----����ֵΪ�������
//
// \˵��:	  1. �޸�.h��ʽ�ļ��з���ֵΪRHS�����ķ�������
//			  2. �޸ķ�������ΪRHS�ķ���ֵ������
//			  3. �ӿں�����InvokeHelper()���е�vtRet��ֵ
// ���ݺ������ӿ�dwDispID��ͬ
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
//
///****************************************************
int ExClassLine::Modi_Method_ID(ExMethod* pMethod)		//����ID��ͬ��ԭ������޸�
{
	ExMethod ExMethodNode;
	ExClass  ExClassNode;
	POSITION pPos;
	POSITION pMethodPos;
	CList<ExMethod,ExMethod&> TemLine;				//���ڴ����ͬDISPID����������
	CString NewNameCStr;

	//�ҵ���ͬDISPID�ķ�����
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		ExClassNode = mLine.GetNext(pPos);
		ExClassNode.FindDispID(pMethod->GetDISPID(),&TemLine);	//��ȡ��ͬID�ķ���
		//�ȽϷ��������ƣ�
		pMethodPos = TemLine.GetHeadPosition();
		while(pMethodPos)
		{
			ExMethodNode = TemLine.GetNext(pMethodPos);
			NewNameCStr = ExMethodNode.GetName();
			NewNameCStr.MakeUpper();
			if(NewNameCStr.Find("GET") >= 0)
			{
				//���ƣ����Խ�������׼�������޸ġ�
				ExMethodNode.Modify_Method(&pMethod->mLine);
				return FILE_FINISHED;
			}
		}
	}
	return FILE_LI_LESSID_ERR;	//�����������ȱ����ͬID�ķ���
}

BOOL ExClassLine::Save(CString FileName)
{

	return true;
}


POSITION ExClassLine::Find(CString nClassName)			//�����������������λ��
{
	POSITION pPos,pOldPos;
	ExClass TemExClass;
	CString TemCStr;;

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemCStr = mLine.GetNext(pPos).GetName();		//��ȡ����
		if(TemCStr == nClassName )
		{
			return pOldPos;
		}
	}
	return NULL;
}

POSITION ExClassLine::Find_A(CString nClassName)		//���������������������������λ��
{
	POSITION pPos,pOldPos;
	CString nClassName_A;
	ExClass TemExClass;
	CString TemCStr;;

	if(nClassName.Left(1) == "C")
		nClassName_A = nClassName.Mid(1);

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemCStr = mLine.GetNext(pPos).GetName();		//��ȡ����
		if(TemCStr == nClassName || TemCStr == nClassName_A)
		{
			return pOldPos;
		}
	}
	return NULL;
}
//**************************************************
// ����   Find_Method_PUT(ExMethod* pMethod)
// ���룺 ExMethod* pMethod  ָ���ṩ�������ͽӿ�dwDispID�ķ���
//        
// ���أ� ����У����ض�ӦPUT�ͷ�����λ�á����򷵻ؼ�
//  
//
//**************************************************/
//��Ϊ�������нӿ�dwDispID��Ψһ����������ͬ�ӿ�dwDispID����£����ݷ�������ͬ��ӽ���ԭ�򣬽��в���
//
POSITION ExClassLine::Find_Method_PUT(ExMethod* pMethod)	//���ݷ�������dwDispID�����Ҷ�Ӧ�Ľӿڵ�PUT����
{
	if(pMethod == NULL) return NULL;				//��ָ�룬ֱ�ӷ��ء�
	
	CString dwDispID;
	CString dwName;

	dwDispID	= pMethod->GetDISPID();				//ָ����ָ������dwDispID
	dwName		= pMethod->GetName();				//ָ����ָ����������
//	dwName	= dwName

//	ExMethod TemMethodNode;
//	ExClass TemClassNode;
	POSITION pPos;
//	POSITION pMethodPos;

//	pPos = this->mLine.GetHeadPosition();
//	while(pPos)
//	{
//		pOldPos = pPos;
//		TemClassNode = mLine.GetNext(pPos);
//		pMethodPos = TemClassNode.Find_PUT(dwDispID);
		//�ж����������Ƿ���ͬ�����

//	}
	return pPos=NULL;
}	
BOOL ExClassLine::FillCBCtrl(CComboBox * pCBCtrl)		//���ο���������䵽��Ͽ���
{
	if(!pCBCtrl) return false;
	pCBCtrl->ResetContent();


	ExClass TemNod;
	POSITION pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNod = mLine.GetNext(pPos);
		pCBCtrl->AddString(TemNod.GetName());
	}
	return true;

}

BOOL ExClassLine::FillListCtrl(CListBox * pListCtrl,CString strFileName)	//���޸ĳ��ִ�����ļ�����䵽�б�
{
	if(!pListCtrl) return false;
	pListCtrl->ResetContent();
	
	CStdioFile  FileRead;
	if(!FileRead.Open(strFileName,CFile::modeRead)) return false;

	int iPathLen;
	iPathLen = strFileName.ReverseFind('\\')+1;		//·���ĳ���

	CString strReadLine;
	CString strTem;
	int iPos;
	while(FileRead.ReadString(strReadLine)) 
	{
		if(iPathLen > 0) strTem = strReadLine.Mid(iPathLen);	//�޳�·��
		
		iPos = strTem.Find("ȱ��ԭ��Ĵ���");
		if(iPos <= 0) continue;
		strTem = strTem.Left(iPos);				//�޳�˵������
		pListCtrl->AddString(strTem);
	}
	
	return true;
}

BOOL ExClassLine::FillStaticCtrl(CString * pCStrCtrl,CString strFileName)	//���޸ĳ��ִ�����ļ�����䵽�б�
{
	if(!pCStrCtrl) return false;
	pCStrCtrl->Empty();
	
	CStdioFile  FileRead;
	if(!FileRead.Open(strFileName,CFile::modeRead)) return false;

	CString strReadLine;
	while(FileRead.ReadString(strReadLine)) 
	{
		*pCStrCtrl += strReadLine + '\n';
	}
	
	return true;


}


EnumElement::EnumElement()
{
	Empty();
}
EnumElement::~EnumElement()
{
	Empty();
}

BOOL EnumElement::SetValue(CString nValue)
{
	nValue.TrimLeft();
	nValue.TrimRight();
	if(ZJGName::CStrIsNumber(nValue))
	{
		mValue = atoi(nValue);
		return true;
	}
	else
	{
		mValue = ENUM_NULL;
		return false;
	}
}

const EnumElement& EnumElement::operator=(const EnumElement& Other)			//��ֵ
{
	mName		= Other.mName;		//��Ա��
	mValue		= Other.mValue;		//ֵ
	mAnnotate	= Other.mAnnotate;	//ע��
	return *this;
}

void EnumElement::Empty()													//���
{
	mName.Empty();		//��Ա��
	mValue = ENUM_NULL;		//ֵ
	mAnnotate.Empty();	//ע��
}

BOOL EnumElement::IsEmpty()													//�ж��Ƿ��ǿ�
{
	if(!mName.IsEmpty())	return false;		//��Ա��
	if(mValue != ENUM_NULL)	return false;		//ֵ
	if(!mAnnotate.IsEmpty())return false;		//ע��
	return true;
}
//*********************
//
//
//
//
//
//*********************
//#define ENUM_NOTHING					0x00000000		//�����û����Ӧ�ĵ���
//#define ENUM_HAVE_ENUM				0x00000001		//������� "emun"����
//#define ENUM_HAVE_LEFT				0x00000002		//������� "{"������
//#define ENUM_HAVE_RIGHT				0x00000004		//������� "}"������
//#define ENUM_HAVE_SEMICOLON			0x00000008		//������� ";"�ֺ�
//#define ENUM_HAVE_RIGHTSEMICOLON		0x00000010		//����� "}"��";"���������м�ֻ�пո��TAB��
//	���䣺
//	1�� int iPos_Enum;					//enum���ʵ�λ��
//	2�� enum XlArabicModes{
//	3�� xlArabicBothStrict/*XXXXX*/	= 3 ,	//
//  4�� xlArabicBothStrict/*XXXXX	= 3 ,
//  5�� enum XlArabicModes//*********************
//  6�� 
int EnumElement::Decompose_CPP(CString strLine)			//��Ⲣ�ֽ�.CPP��.h�ļ���ö�ٶ������
{
	int result = 0;
	strLine.TrimLeft();
	strLine.TrimRight();

	int iPos_Enum;					//enum���ʵ�λ��
	int iPos_Left;					//"{"��λ��
	int iPos_Right;					//"}"��λ��
	int iPos_Notes;					//"//"��λ��
	int iPos_DQM;					//"��"��λ��  double quotation marks 
	int iPos_SQM;					//"��"��λ��  single  quotation marks 
	int iPos_Star_Start;			//"/*"��λ��
	int iPos_Star_End;				//"/*"��λ��
	int iNum_DQM;					//"��"������  double quotation marks
	int iNum_SQM;					//"��"������  single  quotation marks 
	
	iPos_Notes = strLine.Find("//");						//ע�ͺŵ�λ��
	if(iPos_Notes >= 0)	strLine = strLine.Left(iPos_Notes);	//ȥ��ע�Ͳ���
	
	iPos_Star_Start  = strLine.Find("/*");
	iPos_Star_End	 = strLine.Find("*/");
	if(iPos_Star_Start >= 0)										//����"/*"ʱ
	{
		if (iPos_Star_End - iPos_Star_Start >= 2)						//���ڳɶԵ�
		{
			strLine = strLine.Left(iPos_Star_Start) + ' ' + strLine.Mid(iPos_Star_End + 2);
		}
		else strLine = strLine.Left(iPos_Star_Start);
	}
	if(iPos_Star_End >= 0){};										//��ʱ������

	return result;
}

BOOL EnumElement::DecomposeValue(CString strLine)			//�ֽ��ำֵ��䣬����д��Ӧ�Ĳ���
{
	this->Empty();
	int iPos_1,iPos_2;					//λ��
	CString NameCStr,ValueCStr,AnnotationCStr;
	AnnotationCStr.Empty();
	ValueCStr.Empty();

	strLine.TrimLeft();
	NameCStr = strLine;
	iPos_1 = strLine.Find('\t');
	if(iPos_1 > 0)
	{
		NameCStr = strLine.Left(iPos_1);
		ValueCStr= strLine.Mid(iPos_1 + 1);
		
		ValueCStr.TrimLeft();
		iPos_2 = ValueCStr.Find('\t');

		if(iPos_2 > 0)
		{
			AnnotationCStr = ValueCStr.Mid(iPos_2 +1);	
			ValueCStr = ValueCStr.Left(iPos_2);
		}
		AnnotationCStr.TrimLeft();
		ValueCStr.TrimRight();
	}
	
	this->mAnnotate = AnnotationCStr;
	if(ValueCStr.IsEmpty()) this->mValue = ENUM_NULL;
	else 
	{
		if(ZJGName::CStrIsNumber(ValueCStr))	this->mValue = atoi(ValueCStr);
		else 
		{
			this->mValue = ENUM_NULL;
			return false;
		}
	}
	NameCStr.TrimRight();
	this->mName = NameCStr;
	return true;
}

CString EnumElement::Compose()
{
	CString result;
	result.Format("\t%s\t= %d ,\t//%s",mName,mValue,mAnnotate);
	return result;
}
CString EnumElement::ComposeEnd()
{
	CString result;
	result.Format("\t%s\t= %d \t//%s",mName,mValue,mAnnotate);
	return result;
}

//////
DecEnum::DecEnum()
{
	Empty();
}
DecEnum::~DecEnum()
{
	Empty();
}

const DecEnum& DecEnum::operator=(const DecEnum& Other)			//��ֵ
{
	this->mName = Other.mName;
	this->mLine.RemoveAll();

	EnumElement TemNode;
	POSITION pPos;
	pPos = Other.mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = Other.mLine.GetNext(pPos);
		mLine.AddTail(TemNode);
	}
	return *this;
}

void DecEnum::Empty()											//���
{
	this->mName.Empty();
	this->mLine.RemoveAll();
}

BOOL DecEnum::IsEmpty()
{
	if(! mName.IsEmpty() ) return false;
	return mLine.IsEmpty();
}

BOOL DecEnum::Compose(CList<CString,CString&> * pList)
{
	POSITION pPos;
	CString strLine;
	EnumElement ElementNode;

	pList->RemoveAll();

	if(this->mName.IsEmpty()) return false;

	strLine = "enum " + this->mName;
	pList->AddTail(strLine);
	pList->AddTail(CString("{"));

	pPos = mLine.GetHeadPosition();
	while(pPos)
	{
		ElementNode = mLine.GetNext(pPos);
		pList->AddTail(ElementNode.Compose());
	}

	if(mLine.IsEmpty() == NULL )				//���һ��ö��Ԫ�غ���û�ж���
	{
		pPos = pList->GetTailPosition();
		if(pPos)	pList->SetAt(pPos,mLine.GetTail().ComposeEnd());
	}
	pList->AddTail(CString("};"));
	return true;
}
/**************************************/
//	���������͵����� 
//	��һ�֣�
//		XlYesNoGuess
//		xlGuess	0	Excel ȷ���Ƿ��б��⣬����У��Ƿ���һ����
//		xlNo	2	Ĭ��ֵ�� Ӧ�����������������
//		xlYes	1	��Ӧ�����������������
//	�ڶ��֣�
//		XlYesNoGuess
//	�����֣�
//		XlArabicModes
//		XLARABICMODES
//		����	ֵ
//		xlArabicBothStrict	3
//		xlArabicNone	0
//		xlArabicStrictAlefHamza	1
//		xlArabicStrictFinalYaa	2
/**************************************/
BOOL DecEnum::Decompose(CList<CString,CString&> * pList)		//�������е��ַ������ֽ��һ��ö����
{
	this->Empty();						//���

	EnumElement ElementNod;
	POSITION pPos;
	CString NodeCStr;

	pPos = pList->GetHeadPosition();
	while(pPos)
	{
		NodeCStr = pList->GetNext(pPos);
		ElementNod.DecomposeValue(NodeCStr);
		//ȡ����
		if(	this->mName.IsEmpty()		)	   			//��ǰö���������Ϊ��
		{
			if(	ElementNod.GetValue() == ENUM_NULL &&	//ö��Ԫ�� ��ֵ
				ElementNod.GetAnnotate().IsEmpty() )	//ö��Ԫ�� ��ע��
			{
				this->mName = ElementNod.GetName();		//ö��Ԫ�ص����ƾ���ö���������
			}
			else return false;							//����ö����ʱ������ö��Ԫ�ز���������ʱ���Ǵ������롣
			continue;									
		}
		if(	ElementNod.GetValue() == ENUM_NULL &&		//ö��Ԫ�� ��ֵ
			ElementNod.GetAnnotate().IsEmpty() )		//ö��Ԫ�� ��ע��
		{
			continue;									//����ö��������ö��Ԫ�ؿ�ֵ���ظ��У�����
		}
		this->mLine.AddTail(ElementNod);
	}
	return true;
}

POSITION DecEnum::Find(CString strElementName)
{
	POSITION pPos,pOldPos;
	EnumElement TemNode;

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemNode = mLine.GetNext(pPos);
		if(TemNode.GetName() == strElementName)
			return pOldPos;
	}
	return pPos;
}

POSITION DecEnum::Find(int iElementValue)
{
	POSITION pPos,pOldPos;
	EnumElement TemNode;

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemNode = mLine.GetNext(pPos);
		if(TemNode.GetValue() == iElementValue)
			return pOldPos;
	}
	return pPos;
}
//
BOOL DecEnum::Marge(DecEnum* pOther)							//��pOther��ָ��ö����ϲ�����ǰ����
{
	POSITION pPos,pOldPos;
	EnumElement TemNode;
	if(pOther->GetName() != mName) return false;
	pPos = pOther->mLine.GetHeadPosition();
	while(pPos)
	{
		TemNode = pOther->mLine.GetNext(pPos);
		pOldPos = this->Find(TemNode.GetName());
		if(pOldPos) continue;
		this->mLine.AddTail(TemNode);
	}
	return true;
}
//////////////////
DecEnumList::DecEnumList()
{

}

void DecEnumList::Empty()										//���
{
	this->mName.Empty();
	this->mLine.RemoveAll();
}

BOOL DecEnumList::IsEmpty()
{
	if(!this->mName.IsEmpty()) return false;
	return this->mLine.IsEmpty();
}


POSITION DecEnumList::Find(CString DecEnumName)						//����ö����
{
	POSITION pPos,pOldPos;
	DecEnum TemNode;

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemNode = mLine.GetNext(pPos);
		if(TemNode.GetName() == DecEnumName)
			return pOldPos;
	}
	return pPos;;
}

///*******************************************
//
//
//
//
///*******************************************
// ���ļ���ö��������������ö��������
BOOL DecEnumList::Load_CPP_H_File(CString FileName)						//װ��.CPP��.h��ʽ�ļ�
{	
	//˼·��
	//	1�����ж���
	//	2������ö�ٶ��壬����ö�������ѭ����ö�ٶ���ı�־����enum ****�� 
	//	3������ö�ٶ��������䡰};��,�˳�ö�������ѭ��
	//	4�����ö���࣬��������

	
	CString ReadLineCStr,LineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeRead)) return false;

	CList<CString,CString&> TemList;
	
	DecEnum EnumNode,TemEnumNode;
	EnumElement ElementNode, ElementNode_1;
	this->mLine.RemoveAll();
	this->mName = FileName;

	EnumNode.Empty();
	ElementNode.Empty();
	ElementNode_1.Empty();
	TemList.RemoveAll();

	while(myFile.ReadString(LineCStr))
	{
		LineCStr.TrimLeft();
		if(!ElementNode.DecomposeValue(LineCStr)) continue;	//������ȷ���У�����
		if(ElementNode.GetName().IsEmpty()) continue;		//���У�����
		if(ElementNode_1.GetName().CompareNoCase(ElementNode.GetName()) == 0) continue;	//��ͬ������
		
		//��ͬʱΪ�ඨ��ʱ�����������������࿪ʼ��
		if(ElementNode.GetValue() == ENUM_NULL) 
		{
			if(EnumNode.Decompose(&TemList))			//����ö����
			{
				this->mLine.AddTail(EnumNode);			//ֱ�����
			}
			ElementNode_1 = ElementNode;
			TemList.RemoveAll();
			TemList.AddTail(LineCStr);
		}
		else	TemList.AddTail(LineCStr);
	}
	EnumNode.Decompose(&TemList);						//����ö����
	if(!EnumNode.GetName().IsEmpty())					//���ǿ���ʱ�����
		this->mLine.AddTail(EnumNode);
	myFile.Close();	

	Marge();											//ɾ���ظ���
	return true;
}

///*******************************************
//
//
//
//
///*******************************************
BOOL DecEnumList::Load_Txt_File(CString FileName)						//װ��.txt��ʽ�ļ�
{
	CString ReadLineCStr,LineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeRead)) return false;

	CList<CString,CString&> TemList;
	
	DecEnum EnumNode,TemEnumNode;
	EnumElement ElementNode, ElementNode_1;
	this->mLine.RemoveAll();
	this->mName = FileName;

	EnumNode.Empty();
	ElementNode.Empty();
	ElementNode_1.Empty();
	TemList.RemoveAll();

	while(myFile.ReadString(LineCStr))
	{
		LineCStr.TrimLeft();
		if(!ElementNode.DecomposeValue(LineCStr))	continue;					//������ȷ���У�����
		if(ElementNode.GetName().IsEmpty())			continue;					//���У�����
		if(ElementNode_1.GetName().CompareNoCase(ElementNode.GetName()) == 0)
			continue;															//��ͬ������
		
		//��ͬʱΪ�ඨ��ʱ�����������������࿪ʼ��
		if(ElementNode.GetValue() == ENUM_NULL) 
		{
			if(EnumNode.Decompose(&TemList))			//����ö����
			{
				this->mLine.AddTail(EnumNode);			//ֱ�����
			}
			ElementNode_1 = ElementNode;
			TemList.RemoveAll();
			TemList.AddTail(LineCStr);
		}
		else	TemList.AddTail(LineCStr);
	}
	EnumNode.Decompose(&TemList);						//����ö����
	if(!EnumNode.GetName().IsEmpty())					//���ǿ���ʱ�����
		this->mLine.AddTail(EnumNode);
	myFile.Close();	

	Marge();											//ɾ���ظ���
	return true;
}

BOOL DecEnumList::Create_H(CString FileName)						//����.H��ʽ�ļ�
{
	CString LineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeCreate | CFile::modeWrite)) return false;

	POSITION pPos,pListPos;	
	CList<CString,CString&> TemList;
	DecEnum EnumNode,TemEnumNode;

	EnumNode.Empty();

	pPos = mLine.GetHeadPosition();
	while(pPos)
	{
		EnumNode = mLine.GetNext(pPos);
		EnumNode.Compose(&TemList);

		pListPos = TemList.GetHeadPosition();
		while(pListPos)
		{
			LineCStr = TemList.GetNext(pListPos);
			myFile.WriteString(LineCStr + '\n');
		}
		myFile.WriteString("\n");
	}
	return true;
}
//
void DecEnumList::Marge()											//ö����ϲ�
{
	//˼·��1���Ӻ���ǰȡ���ϲ���ö����
	//		2�������ϲ����������е�����Ƚϣ��Ƚ�ʱ��ɨ���Ǵ�ǰ���
	//      3��������ظ�ö���࣬��ϲ�������������ͬʱɾ�����ϲ���ö����
	POSITION pPos,pOldPos,pMargePos;	
	DecEnum EnumNode,MargeNode;

	pPos = mLine.GetTailPosition();						//�Ӻ���ǰ��ʼɨ��
	while(pPos)
	{
		pOldPos  = pPos;
		EnumNode = mLine.GetPrev(pPos);					//ȡ���ϲ���ö����
		if(EnumNode.IsEmpty())							//�ǲ��ǿ���
		{
			mLine.RemoveAt(pOldPos);					//���࣬ɾ��
			continue;									//���������ѭ���Ρ�
		}
	
		pMargePos= this->Find(EnumNode.GetName());		//Ѱ���Ǵ�ǰ���ɨ��
		if(pOldPos == pMargePos) continue;				//��ͬ������û���ظ�������кϲ�����

		//����ͬ����ϲ�
		//
		MargeNode = mLine.GetAt(pMargePos);				//ȡ�ظ���ö����			
		MargeNode.Marge(&EnumNode);						//�ϲ�
		mLine.SetAt(pMargePos,MargeNode);				//��������
		mLine.RemoveAt(pOldPos);						//ɾ�����ϲ���ö����
	}
}