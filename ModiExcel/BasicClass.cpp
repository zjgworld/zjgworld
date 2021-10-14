
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


const ExMethod& ExMethod::operator=(const ExMethod& Other)//赋值
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
// 参数：CString strLine 一行语句
//
// 返回值：CString ----函数名，空值时，不是函数定义语句
//
//判断是否是函数的定义语句,是，返回函数名
//**************************************************/
//函数的定义语句有三种形式
// 1.在.h文件中的形式	a. float * put_Priority(long newValue)
//                      b. float * put_Priority(long& newValue)
//                      c. float * put_Priority(long* newValue)
//                      d. float get_Application()
// 2.在.cpp文件中的形式 a. LPDISPATCH Adjustments::GetApplication()	-----属于类
//                      b. float * Adjustments::GetItem(long Index)	-----属于类
//                      c. void put_Priority(long newValue)			-----不属于类
//						d. long GetCreator()						-----不属于类
// 3.宏定义             a. STDMETHOD(SetFirstPriority)()
//                      b. STDMETHOD(ModifyAppliesToRange)(Range * Range)

BOOL ExMethod::DecomposeMeth(CString strLine)		//分解函数定义语句，并填写函数名
{
	CString LineCStr;
	LineCStr = strLine;
	LineCStr.TrimLeft();
	LineCStr.TrimRight();
	
	this->Empty();

	int iLeft;				//左括号
	int iRight;				//右括号
	int iSpace;				//空格
	int iColon;				//冒号
	int iDefine;			//宏定义
	int iAsterisk;			//星号
	iLeft	= LineCStr.Find('(');
	iRight	= LineCStr.Find(')');
	iSpace	= LineCStr.Find(' ');
	iColon	= LineCStr.Find("::");
	iDefine	= LineCStr.Find("STDMETHOD");
	
	if(iDefine == 0 )								//是宏定义语句
	{
		mName	= LineCStr.Mid(iLeft+1,iRight- iLeft - 1);
		//取函数返回类型
		LineCStr= LineCStr.Mid(iRight+1);
		iLeft	= LineCStr.Find( '(' );			//第二层小括号
		iRight	= LineCStr.Find( ')' );			//第二层小括号
		if( (iLeft >= 0) && (iRight > iLeft))
		{
			m_InType = LineCStr.Mid(iLeft, (iRight - iLeft)+1);	//小括号内的部分，包含左右小括号
		}
		this->m_OutType = "HRESULT";
	}
	else if((iSpace > 0) && (iLeft > iSpace) && (iRight > iLeft))
	{
		//是方法，取方法名
		m_OutType	= LineCStr.Left(iLeft);		//函数返回类型 + [类名+ :: + ] 函数名
		m_InType	= LineCStr.Mid(iLeft);		//( 参数的数据类型 + 参数名 )
		//对前半部分进行分解
		iAsterisk	= m_OutType.ReverseFind('*');
		iSpace		= m_OutType.ReverseFind(' ');
		if(iAsterisk > iSpace)		//当'*'在后面时，用'*'作为分隔符。否则用空格作分隔符
			iSpace = iAsterisk;	

		if (iColon < 0)									//不是类的函数
		{
			mName		= m_OutType.Mid(iSpace+1);		//返回数据类型之后的部分就是函数名
		}
		else if(iColon > iSpace &&  iColon < iLeft)		//是类的函数时
		{
			mName		= m_OutType.Mid(iColon+2);		//双冒号后面的部分是函数名
		}
		m_OutType	= m_OutType.Left(iSpace);			//返回数据类型
	}

	//取函数的输入参数的数据类型
	//输入参数在函数名后面的小括号内，用逗号分隔。接口函数只考虑一个输入参数。	m_InType.TrimLeft();
	m_InType.TrimRight();
	iSpace = m_InType.ReverseFind(' ');
	iAsterisk = m_InType.ReverseFind('*');
	if(iSpace < iAsterisk) iSpace = iAsterisk;		//先出现"*"号时，以"*"作为分隔符
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
CString ExMethod::GetMethodType()			//根据方法(函数)的标记，获取函数返回值对应的数据类型
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
CString ExMethod::GetID()			//通过分解InvokeHelper函数，获取函数接口类型代码
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
			iPos1 += CString("InvokeHelper(").GetLength();		//指针移到小括号后面
			iPos2 = TemNode.Find(',');
			if(iPos2 >=0) TemNode = TemNode.Mid(iPos1,iPos2 - iPos1);
			else TemNode = "";
			return TemNode;
		}	
	}
	return "";
}

BOOL ExMethod::DecomposeIvHp( )			//分解InvokeHelper语句，并填写对应的参数
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

BOOL ExMethod::DecomposeIvHp(CString strLine)			//分解InvokeHelper语句，并填写对应的参数
{
	int iPos1,iPos2;
	CString TemCStr;
	iPos1 = strLine.Find("InvokeHelper(");
	if(iPos1 < 0)	return false;					//位置不正确，直接返回

	////先获取InvokeHelper函数的接口类型
	iPos1 += CString("InvokeHelper(").GetLength();	//指针移到小括号后面
	iPos2 = strLine.Find(',');						//第一个逗号的位置
	if(iPos2 < iPos1)	return false;				//位置不正确，直接返回

	m_DISPID = strLine.Mid(iPos1,iPos2 - iPos1);	//接口类型
	m_DISPID.TrimLeft();		
	m_DISPID.TrimRight();
		
	//下面求接口的返回数据类型，在第二个与第三个逗号之间
	iPos1 = strLine.Find(',',iPos2+1);				//第二个逗号的位置
	if(iPos1 < iPos2)	return false;				//位置不正确，直接返回

	iPos2 = strLine.Find(',',iPos1+1);				//第三个逗号的位置
	if(iPos2 < iPos1)	return false;				//位置不正确，直接返回

	m_VARTYPE = strLine.Mid(iPos1+1,iPos2-iPos1-1);
	m_VARTYPE.TrimLeft();
	m_VARTYPE.TrimRight();

	if(m_VARTYPE.IsEmpty())
	{
		this->m_OutType.Empty();	
		return false;								//接口的返回数据类型为空
	}
	this->m_OutType = Get_VARENUM_Type(m_VARTYPE);
	return true;	
}

// vtRet： 
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
CString ExMethod::Get_VARENUM_Type(CString nVARENUM)			//获取VARENUM对应的变量类型
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

CString ExMethod::Get_Type_VARENUM(CString nType)			//获取变量类型对应VARENUM的值
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
// 函数   Modify_Method(CList<CString,CString&> * pList)
// 输入： ExMethod* pMethod  方法
//        
// 返回： 如果有，返回对应PUT型方法的位置。否则返回假
//  
//
//**************************************************/
void ExMethod::Modify_Method(CList<CString,CString&> * pList)	//以当前方法为参考类，对于pList所指向的语句进行修改
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
// 函数   Modify_Method_PUT(CList<CString,CString&> * pList)
// 输入： ExMethod* pMethod  方法
//        
// 返回： 如果有，返回对应PUT型方法的位置。否则返回假
//  
//
//**************************************************/
// 用PUT_类函数的输入，代替GET_类函数的输出。
// 涉及到二个部分：a.返回数据类型；b.接口函数中的返回数据
BOOL ExMethod::Modify_Method_PUT(CList<CString,CString&> * pList)	//以当前方法为参考类，对于pList所指向的语句进行修改
{
	BOOL bRtn = true;
	CString NodeCStr,OutTypeCStr,VARTYPECStr;
	VARTYPECStr = this->Get_Type_VARENUM(m_InType);				//用PUT_类函数的输入，代替GET_类函数的输出
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
		NodeCStr.Replace("RHS ",this->m_InType+ ' ');			//用PUT_类函数的输入，代替GET_类函数的输出
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

const ExClass& ExClass::operator=(const ExClass& Other)//赋值
{
	this->mName = Other.mName;
	this->mBase = Other.mBase;		//基类名
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

CString ExClass::GetType(CString dwDispID)					//根据方法(函数)的标记，获取函数返回值对应的数据类型
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

CString ExClass::Get_Type_VARENUM(CString nType)			//获取变量类型对应VARENUM的值
{
	ExMethod TemNode;
	return TemNode.Get_Type_VARENUM(nType);

}
CString ExClass::Get_VARENUM_Type(CString nVARENUM)			//获取VARENUM对应的变量类型
{
	ExMethod TemNode;
	return TemNode.Get_VARENUM_Type(nVARENUM);
}
//**************************************************
// 函数   Find_PUT(ExMethod* pMethod)
// 输入： ExMethod* pMethod  方法
//        
// 返回： 如果有，返回对应PUT型方法的位置。否则返回假
//  
//
//**************************************************/
POSITION ExClass::Find_PUT(CString dwDispID)				//根据dwDispID，查找对应的接口的PUT函数
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
// 函数   FindDispID(CString dwDispID,CList<ExMethod,ExMethod&>* pLine)
// 输入： dwDispID：InvokeHelper函数的接口类型  
//        pLine：存放查找到的方法的链表指针
// 返回： 如果有，返回真。否则返回假
//  
//
//**************************************************/
void ExClass::FindDispID(CString dwDispID,CList<ExMethod,ExMethod&>* pLine)	//根据InvokeHelper函数的接口类型，查找对应的方法
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
	//获取对应接口类型的返回值的标识
	
	str_Type	= GetType(pMethod->GetDISPID());			//接口类型的返回值的标识
	if(str_Type.IsEmpty()) return FILE_C_LESSID_ERR;		//当前类中缺少相同ID的方法，直接返回并报错
	str_VARTYPE	= Get_Type_VARENUM(str_Type);				//接口类型的返回值的标识
	if(str_VARTYPE.IsEmpty()) return FILE_VARTYPE_ERR;		//无对应的数据类型，直接返回并报错
		
	str_Type	= str_Type + " ";
	pPos = pMethod->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemCStr = pMethod->mLine.GetNext(pPos);			
		TemCStr.Replace("RHS ",str_Type);				//更改“RHS ”
		TemCStr.Replace(", ,",',' + str_VARTYPE + ',');	//在“, , ”处，添加接口类型的返回值的标识
		pMethod->mLine.SetAt(pOldPos,TemCStr);
	}
	return FILE_FINISHED;
}

int ExClass::Modify_Method(CList<CString,CString&> * pList,CString * pstrMethod)
{
	//查找相同的接口函数，并准备替换。
	CString str_VARTYPE;
	CString str_Type;
	
	int iPos = pstrMethod->Find('(');	//是否是接口方法
	int iColon = pstrMethod->Find(',');	//逗号的位置
	CString TemCStr;
	POSITION pPos,pOldPos;
	if(iColon >= 0)
	{
		TemCStr = pstrMethod->Mid(iPos+1, iColon - iPos -1);
		str_VARTYPE	= GetType(TemCStr);							//接口类型的返回值的标识
		if(str_VARTYPE.IsEmpty()) return FILE_C_LESSID_ERR;		//缺少相同ID的方法
		
		str_Type	=this->Get_VARENUM_Type(str_VARTYPE);		//返回值的数据类型
		if(str_Type.IsEmpty()) return FILE_VARTYPE_ERR;			//无对应的数据类型，直接返回并报错
		
		str_Type	= str_Type + " ";
		//对链表中的行进行修改
		pPos = pList->GetHeadPosition();
		while(pPos)
		{
			pOldPos = pPos;
			TemCStr = pList->GetNext(pPos);			
			TemCStr.Replace("RHS ",str_Type);				//更改“RHS ”
			TemCStr.Replace(", ,",',' + str_VARTYPE + ',');	//在“, , ”处，添加接口类型的返回值的标识
			pList->SetAt(pOldPos,TemCStr);
		}
	}
	return FILE_FINISHED;
}

void ExClass::DecomposeClassName(CString LineCStr)			//分解类定义语句，并填写对应的参数
{
	//取类名
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
//  用于装载EXCEL的类文件
//*********************************
int ExClassLine::Load_H_File(CString FileName)			//装载.h格式文件
{
	CString ReadLineCStr,LineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeRead)) return false;

	ExClass ExClassNode;
	ExMethod ExMethodNode;
	this->mLine.RemoveAll();
	this->mFileName = FileName;

	BOOL bClass;
	int iColon;				//冒号
	int ibrackets;			//大括号
	CString strClassName,TemType;;

	bClass = false;
	ibrackets = 0;
	ExClassNode.Empty();
	ExMethodNode.Empty();

	while(myFile.ReadString(ReadLineCStr))
	{
		LineCStr = ReadLineCStr;
		LineCStr.TrimLeft();
		
		if(ibrackets >= 2)						//大括号的行才处理
		{
			ExMethodNode.mLine.AddTail(ReadLineCStr);
			//在类中的方法中
			if(LineCStr.Find('{') >= 0)	ibrackets++;			
			if(LineCStr.Find('}') >= 0)
			{
				ibrackets--;
			}

			if(ibrackets < 2)						//出了方法的大括号
			{
				TemType = ExMethodNode.GetOutType();
				ExMethodNode.DecomposeIvHp();
				ExMethodNode.SetOutType(TemType);				
				//方法结束，添加到类链表中
				ExClassNode.mLine.AddTail(ExMethodNode);
				ExMethodNode.Empty();
			}
		}
		else if(ibrackets == 1)						//大括号的行才处理
		{
			//对类中的方法进行处理
			if(LineCStr.Find('{') >= 0)	ibrackets++;

			if(ExMethodNode.GetName().IsEmpty())				//前一语句不是定义语句，只进行判断本语句是否是定义语句
			{
				ExMethodNode.DecomposeMeth(LineCStr);			//分解语句，并判断是否是函数定义语句
			}
			if(ExMethodNode.GetName().IsEmpty() == NULL)		//本判断语句，与上一判断语句是顺序的，不可以合并
			{													//1。如果当前语句是定义语句，则添加。
				ExMethodNode.mLine.AddTail(ReadLineCStr);		//2。前一语句是定义语句,也添加
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
		else if(bClass)											//已经有了类定义语句
		{
			ibrackets = 0;							//大括号将要开始

			if(LineCStr.Find('{') >= 0)	ibrackets++;
			if(LineCStr.Find('}') >= 0)
			{	ibrackets--;
				if(ibrackets < 0)					//大括号不成对，出现错误
				{
					myFile.Close();	
					return FILE_ERROR_ERR;	
				}
			}
		}
		else if(LineCStr.Left(6).Compare("class ") == 0 )					//是类
		{
			//取类名,类定义有二种形式：1。class ****
			//                         2。class **** : *******
			iColon = LineCStr.Find(':');
			if(iColon > 0) strClassName = LineCStr.Mid(6,iColon - 6);
			else strClassName = LineCStr.Mid(6);
			strClassName.TrimRight();
			strClassName.TrimLeft();

			bClass = true;
			ibrackets = 0;							//大括号将要开始
			ExClassNode.Empty();
			ExClassNode.SetName(strClassName);

			if(LineCStr.Find('{') >= 0)	ibrackets++;
			if(LineCStr.Find('}') >= 0)
			{	ibrackets--;
				if(ibrackets < 0)					//大括号不成对，出现错误
				{
					myFile.Close();	
					return FILE_ERROR_ERR;	
				}
			}
		}
		else
		{
			//在类的外部，空运行
			if(LineCStr.Find('{') >= 0 )	ibrackets++;

			if(LineCStr.Find('}') >= 0)		ibrackets--;
		}
	}
	myFile.Close();	
	if(ibrackets == 0)	return FILE_FINISHED;		//大括号成对，正常结束
	else return FILE_ERROR_ERR;						//大括号不成对，出现错误
}

BOOL ExClassLine::LoadCPPFile(CString FileName)			//装载.CPP格式文件
{
	CString ReadLineCStr;
	CStdioFile myFile;
	if(!myFile.Open(FileName,CFile::modeRead)) return false;
	int iSpace;				//空格
	int iLeft;				//左括号
	int iRight;				//右括号
	int iColon;				//冒号
	int ibrackets;			//大括号
	int iAnnotation;		//注释号
	int iPos, iLineLen;		//一行的长度
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
		
		iLineLen = ReadLineCStr.GetLength();						//一行的长度
		if(iLineLen <= 0) continue;									//空行不处理
		if(iLineLen >= 2)
		{
			if(ReadLineCStr.GetAt(0) == '/' && ReadLineCStr.GetAt(1) == '/')
				continue;		//注释行不处理
			
			//注释块不处理
			iPos = ReadLineCStr.Find("/*");
			if(iPos >= 0)
			{
				iAnnotation++;
				SurplusCStr = ReadLineCStr.Left(iPos);				//一行中剩余的部分，在“/*”之前的部分
				ReadLineCStr= ReadLineCStr.Mid(iPos);				//一行中剩余的部分，在“/*”之后的部分
				do
				{
					iPos = ReadLineCStr.Find("*/");
					if(iPos >= 0) iAnnotation--;
				}while(iAnnotation > 0 && myFile.ReadString(ReadLineCStr));
			}
			if(SurplusCStr.GetLength() > 0) ReadLineCStr = SurplusCStr;
		}
		//本程序中不能对类似 { ClassName = ReadLineCStr.Mid(iSpace+1,/*局部注释*/iColon - iSpace -1);} 语句进行处理
		//对于正常的一行进行处理
		iSpace	= ReadLineCStr.Find(" ",	0);
		iLeft	= ReadLineCStr.Find("(",	0);
		iRight	= ReadLineCStr.Find(")",	0);
		iColon	= ReadLineCStr.Find("::",	0);

		if( (iSpace > 0) &&
			(iLeft < iRight) &&
			(iSpace < iColon) &&
			(iColon < iLeft) )
		{
			ClassName = ReadLineCStr.Mid(iSpace+1,iColon - iSpace -1);	//类名
			MethodName= ReadLineCStr.Mid(iColon+2,iLeft - iColon -2);	//方法名
			//新的方法开始
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
				//新类开始，
				//将已有的类，添加到链表中，并更新类
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
		if(ibrackets <= 0) continue;		//不在大括号内的行不处理
		ExMethodNode.mLine.AddTail(ReadLineCStr);
		
	}
	{
		//结束读文件
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
// \函数名称：Modify_H_File_PUT(CString FileName)
// \输入参数: FileName				----待修改函数的所在文件名
//		
// \返回值:   int					----返回值为错误代码
//
// \说明:	  1. 修改.h格式文件中返回值为RHS函数的返回类型。如“RHS get_Creator()”
//			  2. 修改返回类型为RHS的返回值的类型。如“RHS result;”
//			  3. 接口函数（InvokeHelper()）中的vtRet的值。如“InvokeHelper(0x95, DISPATCH_PROPERTYGET, , (void*)&result, nullptr);”
//1。在同一文件内，对相同接口的 "put_" 函数，取函数参数的数据类型，作为对应 "get_" 函数的返回类型，并进行相应的修改
//2。一个文件只能有一个类。
///****************************************************
int ExClassLine::Modify_H_File_PUT(CString OldFileName,CString NewFileName)
{
	//思路：1.先装载类。
	//		2.再对文件进行循环。A。遇到“RHS”方法时，调用 Find_Method_PUT(ExMethod* pMethod);
	//							B。找到对应的“PUT_”函数时，进行修改。找不到时直接复制，并设置错误代码
	//							C. 如果不是“RHS”方法时，直接复制。
	//      
	POSITION pPos;
	CString TemCStr,ClassCStr;
	ExClass TemExClass,BaseExClass;
	ExMethod TemMethod,BaseMethod;
	
	this->Load_H_File(OldFileName);
	BaseExClass = this->mLine.GetHead();	//作为基准类
	CString ReadLineCStr;
	CString LineCStr;
//	CString NewFileName;
	CStdioFile myOldFile;				//原文件
	CStdioFile myNewFile;				//新文件
	int iError = FILE_FINISHED;
	int iRtn  = FILE_FINISHED;

	int iPos;
	iPos = mFileName.Find(".h");						//查找.h
	if(iPos <= 0) iPos = mFileName.Find(".H");		
	if(iPos <= 0) return FILE_NAME_ERROR;				//文件名不正确 

	if(!myOldFile.Open(OldFileName,CFile::modeRead)) return FILE_READ_ERROR;							//读文件错误
	if(!myNewFile.Open(NewFileName,CFile::modeWrite | CFile::modeCreate)) return FILE_WRITE_ERROR;		//写文件错误
	
	int ibrackets = 0;					//大括号
	int iClass = 0;						//类的层数
	int iAnnotation;					//注释号
	BOOL bClass = false;				//已是类的标志
	BOOL bWrite = true;					//如果原文已有命名空间定义，则为假。

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

		if(!bClass)	//前面不是类定义语句
		{
			if(LineCStr.Left(6).Compare("class ") == 0)
			{
				bClass = true;//是类定义语句
				ClassCStr = ";\n" + LineCStr;
			}
			else 
			{
				myNewFile.WriteString(ReadLineCStr +"\n");//不是类定义语句,直接复制
				continue;
			}
		}
		if(LineCStr.Find(';') > 0)						//
		{
			//当前语句是类定义，但是有分号，代表类结束了，如：class ExMethod;
			bClass = false;								//类结束了
			myNewFile.WriteString(ReadLineCStr +"\n");	//类结束了,直接复制
			continue;
		}
		//取类名
		ClassCStr = ClassCStr + LineCStr;			//类定义语句还没有结束，所以加上当前语句
		TemExClass.DecomposeClassName(ClassCStr);
		if(TemExClass.GetName().IsEmpty())continue;	//没有类名。如 class
													//             ****{
		do
		{
			LineCStr = ReadLineCStr;
			LineCStr.TrimLeft();
			iAnnotation = LineCStr.Find("//");
			if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
			LineCStr.TrimRight();
		
			//以下是在类内进行循环，
			//做准备。
			if(LineCStr.Find('{') < 0)
			{
				//当前语句不在大括号内时，直接复制
				myNewFile.WriteString(ReadLineCStr +"\n");
				continue;
			}
			if(ibrackets < 1)
			{
				myNewFile.WriteString(ReadLineCStr +"\n");		//当前语句是带“{”，直接复制
			}
			ibrackets = 1;

			//*********************************
			//下面程序段是在类内循环
			//*********************************
			TemMethod.Empty();

			//获取基准的类
			while(myOldFile.ReadString(ReadLineCStr))
			{
				LineCStr = ReadLineCStr;
				LineCStr.TrimLeft();
				iAnnotation = LineCStr.Find("//");
				if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
				LineCStr.TrimRight();

				if(TemMethod.DecomposeMeth(LineCStr) == NULL)
				{
					//当前语句不在方法的大括号内时，直接复制
					myNewFile.WriteString(ReadLineCStr +"\n");
					continue;
				}
				//*********************************
				//下面程序段是在方法段内循环
				//*********************************
				do		//在方法段内循环
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
						if(ibrackets == 1) break;			//当大括号只有一层时,已不在方法段内
					}
				}while(myOldFile.ReadString(ReadLineCStr));	//方法段结束
				TemMethod.SetDISPID(TemMethod.GetID());		//设置方法的ID号
				//*********************************
				//		上面程序段是在方法段内循环，内层do循环结束
				//*********************************

				//对方法进行处理。
				if(TemMethod.GetOutType().Compare("RHS") == 0 )//是待修改的方法
				{
					//是需要修改的方法
					pPos = BaseExClass.Find_PUT(TemMethod.GetDISPID());
					if(pPos)
					{
						BaseMethod = BaseExClass.mLine.GetAt(pPos);	//基准方法
						if(BaseMethod.Modify_Method_PUT(&TemMethod.mLine) == false)
						{//设置返回错误类型
							iRtn = iRtn | FILE_VARTYPE_ERR;		//没有对应的数据类型
						}
					}
					else iRtn = iRtn | FILE_C_LESSPUT_ERR;		//类中缺少成对的ID错误
				}
	
				TemExClass.mLine.AddTail(TemMethod);
				//将方法内的行写入到新的文件
				pPos = TemMethod.mLine.GetHeadPosition();
				while(pPos)
				{
					TemCStr = TemMethod.mLine.GetNext(pPos);
					myNewFile.WriteString(TemCStr +"\n");
				}
				//*********************************
				//		上面程序段是在方法段内循环，内层while(pPos)
				//*********************************
			}
	
			//*********************************
			//		上面程序段是在类内循环，类内的外层while(myOldFile.ReadString(ReadLineCStr))
			//*********************************
		}while(myOldFile.ReadString(ReadLineCStr));	
		//*********************************
		//		以上是在类内进行循环。外层do循环到此结束
		//*********************************
	}
	//*********************************
	//		上面程序段是在文件内循环，文件内的循环while(myOldFile.ReadString(ReadLineCStr))
	//*********************************

	myOldFile.Close();
	myNewFile.Close();

	return iRtn;
}


///****************************************************
// \函数名称：Modi_Method(ExMethod* pModiNode)
// \输入参数: ExMethod* pModiNode	----待修改函数的指针
//		
// \返回值:   int					----返回值为错误代码
//
// \说明:	  1. 修改.h格式文件中返回值为RHS函数的返回类型
//			  2. 修改返回类型为RHS的返回值的类型
//			  3. 接口函数（InvokeHelper()）中的vtRet的值
// 根据接口dwDispID相同,进行修改。先同类，再同方法，最后同接口的顺序
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
//
///****************************************************
int ExClassLine::Modify_H_File_A(CString OldFName,CString NewFName)
{
	//思路：1.先对文件的行进行循环，在类外，直接进行复制，不修改
	//		2.遇到类定义时，A. 先取同类或相似类，以便按同类或相似类为参照时行修改。
	//                      B. 在类内里，进行类内循环。不是方法时，直接复制，不修改。
	//						C. 是方法时，判断是不是需要修改修改的方法，如果是，调用方法修改函数
	//                                          a. 如果有参考类时，调用参考类修改函数进行修改
    //                                          b. 没有参考类，找相同的方法名，然后调用参考方法修改函数进行修改
	//                                          c. 没有相同的方法名，找相同接口编号的方法进行修改
	//                      D. 在类内部，不需要进行修改的类，直接复制。
	//      
	CString ReadLineCStr;
	CString LineCStr;
	CString NewFileName;
	CStdioFile myOldFile;				//原文件
	CStdioFile myNewFile;				//新文件
	int iError = FILE_FINISHED;
	int iRtn  = FILE_FINISHED;

	int iPos;
	iPos = OldFName.Find(".h");						//查找.h
	if(iPos <= 0) iPos = OldFName.Find(".H");		
	if(iPos <= 0) return FILE_NAME_ERROR;														//文件名不正确 
	if(!myOldFile.Open(OldFName,CFile::modeRead)) return FILE_READ_ERROR;						//读文件错误
	if(!myNewFile.Open(NewFName,CFile::modeWrite | CFile::modeCreate)) return FILE_WRITE_ERROR;	//写文件错误
	
	int ibrackets = 0;					//大括号
	int iClass = 0;						//类的层数
	int iAnnotation;					//注释号

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

		if(LineCStr.Left(6).Compare("class ") != 0 )	//不是类定义语句,直接复制
		{
			//不是类时，直接复制
			myNewFile.WriteString(ReadLineCStr +"\n");
			continue;
		}
		if(LineCStr.Find(';') > 0)						//
		{
			//当前语句是类定义，但是有分号，代表类结束了，如：class ExMethod;
			myNewFile.WriteString(ReadLineCStr +"\n");
			continue;
		}
		//取类名
		TemExClass.DecomposeClassName(LineCStr);
		if(TemExClass.GetName().IsEmpty()) continue;
		do
		{
			LineCStr = ReadLineCStr;
			LineCStr.TrimLeft();
			iAnnotation = LineCStr.Find("//");
			if(iAnnotation >=0 ) LineCStr = LineCStr.Left(iAnnotation);
			LineCStr.TrimRight();
		
			//以下是在类内进行循环，
			//做准备。
			if(LineCStr.Find('{') < 0)
			{
				//当前语句不在大括号内时，直接复制
				myNewFile.WriteString(ReadLineCStr +"\n");
				continue;
			}

			ibrackets = 1;
			myNewFile.WriteString(ReadLineCStr +"\n");	//写入含“{”的行

			//*********************************
			//下面程序段是在类内循环
			//*********************************
			TemMethod.Empty();

			//获取基准的类
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
					//当前语句不在方法的大括号内时，直接复制
					myNewFile.WriteString(ReadLineCStr +"\n");
					continue;
				}
				//*********************************
				//下面程序段是在方法段内循环
				//*********************************

				do		//在方法段内循环
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
						if(ibrackets == 1) break;			//当大括号只有一层时,已不在方法段内
					}
				}while(myOldFile.ReadString(ReadLineCStr));	//方法段结束
				TemMethod.SetDISPID(TemMethod.GetID());
				//*********************************
				//		上面程序段是在方法段内循环，内层do循环结束
				//*********************************

				//对方法进行处理。
				if(TemMethod.GetOutType().Compare("RHS") == 0 )//是待修改的方法
				{
					//是需要修改的方法
					bFlag = BaseExClass.Modify_Method(&TemMethod);
					iRtn = iRtn | bFlag ;
					if(bFlag == FILE_C_LESSID_ERR )			//基准类中没有对应的方法
					{
						bFlag = this->Modi_Method_Name(&TemMethod);
						iRtn = iRtn | bFlag ;
						if(bFlag == FILE_LN_LESSID_ERR )	//链表中没有对应的函数名的方法
						{
							bFlag = this->Modi_Method_ID(&TemMethod);
							iRtn = iRtn | bFlag ;
						}
					}
				}
				TemExClass.mLine.AddTail(TemMethod);
				//将方法内的行写入到新的文件
				pPos = TemMethod.mLine.GetHeadPosition();
				while(pPos)
				{
					TemCStr = TemMethod.mLine.GetNext(pPos);
					myNewFile.WriteString(TemCStr +"\n");
				}
				//*********************************
				//		上面程序段是在方法段内循环，内层while(pPos)
				//*********************************
			}
			//*********************************
			//		上面程序段是在类内循环，外层while(myOldFile.ReadString(ReadLineCStr))
			//*********************************
		}while(myOldFile.ReadString(ReadLineCStr));	
		//*********************************
		//		以上是在类内进行循环。外层do循环到此结束
		//*********************************
	}
	myOldFile.Close();				//原文件
	myNewFile.Close();				//新文件
	return iRtn;
}

///****************************************************
// \函数名称：Modi_Method_Name(ExMethod* pModiNode)
// \输入参数: ExMethod* pModiNode	----待修改函数的指针
//		
// \返回值:   int					----返回值为错误代码
//
// \说明:	  1. 修改.h格式文件中返回值为RHS函数的返回类型
//			  2. 修改返回类型为RHS的返回值的类型
//			  3. 接口函数（InvokeHelper()）中的vtRet的值
// 根据函数名相同或相似，接口dwDispID相同
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
//
///****************************************************
int ExClassLine::Modi_Method_Name(ExMethod* pMethod)		//根据函数名相同或相似，接口dwDispID相同的原则进行修改
{
	ExMethod ExMethodNode;
	ExClass  ExClassNode;
	POSITION pPos;
	POSITION pMethodPos;
	CList<ExMethod,ExMethod&> TemLine;				//用于存放相同DISPID方法的链表

	int iPos;
	CString OldName_1,OldName_2,NewNameCStr;
	//找到相同DISPID的方法，然后比较方法名，如果方法名相同，则修改
	//方法名如 get_Application() 和 GetApplication()
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
		ExClassNode.FindDispID(pMethod->GetDISPID(),&TemLine);	//获取相同ID的方法
		//比较方法的名称，是否相似
		pMethodPos = TemLine.GetHeadPosition();					//用于存放相同DISPID方法的链表
		while(pMethodPos)
		{
			ExMethodNode = TemLine.GetNext(pMethodPos);
			NewNameCStr = ExMethodNode.GetName();
			NewNameCStr.MakeUpper();
			if(	NewNameCStr.CompareNoCase(OldName_1) == 0 || 
				NewNameCStr.CompareNoCase(OldName_2) == 0 ) 
			{
				//相似，可以借用作标准，用来修改。
				ExMethodNode.Modify_Method(&pMethod->mLine);
				return FILE_FINISHED;
			}
		}
	}
	return FILE_LN_LESSID_ERR;	//函数名相近的缺少相同ID的方法
}
///****************************************************
// \函数名称：Modi_Method_ID(ExMethod* pModiNode)
// \输入参数: ExMethod* pModiNode	----待修改函数的指针
//		
// \返回值:   int					----返回值为错误代码
//
// \说明:	  1. 修改.h格式文件中返回值为RHS函数的返回类型
//			  2. 修改返回类型为RHS的返回值的类型
//			  3. 接口函数（InvokeHelper()）中的vtRet的值
// 根据函数名接口dwDispID相同
// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );
//
///****************************************************
int ExClassLine::Modi_Method_ID(ExMethod* pMethod)		//基于ID相同的原则进行修改
{
	ExMethod ExMethodNode;
	ExClass  ExClassNode;
	POSITION pPos;
	POSITION pMethodPos;
	CList<ExMethod,ExMethod&> TemLine;				//用于存放相同DISPID方法的链表
	CString NewNameCStr;

	//找到相同DISPID的方法，
	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		ExClassNode = mLine.GetNext(pPos);
		ExClassNode.FindDispID(pMethod->GetDISPID(),&TemLine);	//获取相同ID的方法
		//比较方法的名称，
		pMethodPos = TemLine.GetHeadPosition();
		while(pMethodPos)
		{
			ExMethodNode = TemLine.GetNext(pMethodPos);
			NewNameCStr = ExMethodNode.GetName();
			NewNameCStr.MakeUpper();
			if(NewNameCStr.Find("GET") >= 0)
			{
				//相似，可以借用作标准，用来修改。
				ExMethodNode.Modify_Method(&pMethod->mLine);
				return FILE_FINISHED;
			}
		}
	}
	return FILE_LI_LESSID_ERR;	//函数名相近的缺少相同ID的方法
}

BOOL ExClassLine::Save(CString FileName)
{

	return true;
}


POSITION ExClassLine::Find(CString nClassName)			//根据类名，查找类的位置
{
	POSITION pPos,pOldPos;
	ExClass TemExClass;
	CString TemCStr;;

	pPos = this->mLine.GetHeadPosition();
	while(pPos)
	{
		pOldPos = pPos;
		TemCStr = mLine.GetNext(pPos).GetName();		//获取类名
		if(TemCStr == nClassName )
		{
			return pOldPos;
		}
	}
	return NULL;
}

POSITION ExClassLine::Find_A(CString nClassName)		//根据类名和相似类名，查找类的位置
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
		TemCStr = mLine.GetNext(pPos).GetName();		//获取类名
		if(TemCStr == nClassName || TemCStr == nClassName_A)
		{
			return pOldPos;
		}
	}
	return NULL;
}
//**************************************************
// 函数   Find_Method_PUT(ExMethod* pMethod)
// 输入： ExMethod* pMethod  指向提供方法名和接口dwDispID的方法
//        
// 返回： 如果有，返回对应PUT型方法的位置。否则返回假
//  
//
//**************************************************/
//因为在链表中接口dwDispID不唯一，所以在相同接口dwDispID情况下，根据方法名相同或接近的原则，进行查找
//
POSITION ExClassLine::Find_Method_PUT(ExMethod* pMethod)	//根据方法名和dwDispID，查找对应的接口的PUT函数
{
	if(pMethod == NULL) return NULL;				//空指针，直接返回。
	
	CString dwDispID;
	CString dwName;

	dwDispID	= pMethod->GetDISPID();				//指针所指方法的dwDispID
	dwName		= pMethod->GetName();				//指针所指方法的名称
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
		//判定方法名称是否相同或相近

//	}
	return pPos=NULL;
}	
BOOL ExClassLine::FillCBCtrl(CComboBox * pCBCtrl)		//将参考用类名填充到组合框中
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

BOOL ExClassLine::FillListCtrl(CListBox * pListCtrl,CString strFileName)	//将修改出现错误的文件名填充到列表
{
	if(!pListCtrl) return false;
	pListCtrl->ResetContent();
	
	CStdioFile  FileRead;
	if(!FileRead.Open(strFileName,CFile::modeRead)) return false;

	int iPathLen;
	iPathLen = strFileName.ReverseFind('\\')+1;		//路径的长度

	CString strReadLine;
	CString strTem;
	int iPos;
	while(FileRead.ReadString(strReadLine)) 
	{
		if(iPathLen > 0) strTem = strReadLine.Mid(iPathLen);	//剔除路径
		
		iPos = strTem.Find("缺少原类的错误");
		if(iPos <= 0) continue;
		strTem = strTem.Left(iPos);				//剔除说明部分
		pListCtrl->AddString(strTem);
	}
	
	return true;
}

BOOL ExClassLine::FillStaticCtrl(CString * pCStrCtrl,CString strFileName)	//将修改出现错误的文件名填充到列表
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

const EnumElement& EnumElement::operator=(const EnumElement& Other)			//赋值
{
	mName		= Other.mName;		//成员名
	mValue		= Other.mValue;		//值
	mAnnotate	= Other.mAnnotate;	//注释
	return *this;
}

void EnumElement::Empty()													//清空
{
	mName.Empty();		//成员名
	mValue = ENUM_NULL;		//值
	mAnnotate.Empty();	//注释
}

BOOL EnumElement::IsEmpty()													//判断是否是空
{
	if(!mName.IsEmpty())	return false;		//成员名
	if(mValue != ENUM_NULL)	return false;		//值
	if(!mAnnotate.IsEmpty())return false;		//注释
	return true;
}
//*********************
//
//
//
//
//
//*********************
//#define ENUM_NOTHING					0x00000000		//语句中没有相应的单词
//#define ENUM_HAVE_ENUM				0x00000001		//语句中有 "emun"单词
//#define ENUM_HAVE_LEFT				0x00000002		//语句中有 "{"左括号
//#define ENUM_HAVE_RIGHT				0x00000004		//语句中有 "}"右括号
//#define ENUM_HAVE_SEMICOLON			0x00000008		//语句中有 ";"分号
//#define ENUM_HAVE_RIGHTSEMICOLON		0x00000010		//语句中 "}"和";"相连，或中间只有空格或TAB键
//	例句：
//	1。 int iPos_Enum;					//enum单词的位置
//	2。 enum XlArabicModes{
//	3。 xlArabicBothStrict/*XXXXX*/	= 3 ,	//
//  4。 xlArabicBothStrict/*XXXXX	= 3 ,
//  5。 enum XlArabicModes//*********************
//  6。 
int EnumElement::Decompose_CPP(CString strLine)			//检测并分解.CPP和.h文件中枚举定义语句
{
	int result = 0;
	strLine.TrimLeft();
	strLine.TrimRight();

	int iPos_Enum;					//enum单词的位置
	int iPos_Left;					//"{"的位置
	int iPos_Right;					//"}"的位置
	int iPos_Notes;					//"//"的位置
	int iPos_DQM;					//"“"的位置  double quotation marks 
	int iPos_SQM;					//"‘"的位置  single  quotation marks 
	int iPos_Star_Start;			//"/*"的位置
	int iPos_Star_End;				//"/*"的位置
	int iNum_DQM;					//"“"的数量  double quotation marks
	int iNum_SQM;					//"‘"的数量  single  quotation marks 
	
	iPos_Notes = strLine.Find("//");						//注释号的位置
	if(iPos_Notes >= 0)	strLine = strLine.Left(iPos_Notes);	//去掉注释部分
	
	iPos_Star_Start  = strLine.Find("/*");
	iPos_Star_End	 = strLine.Find("*/");
	if(iPos_Star_Start >= 0)										//存在"/*"时
	{
		if (iPos_Star_End - iPos_Star_Start >= 2)						//存在成对的
		{
			strLine = strLine.Left(iPos_Star_Start) + ' ' + strLine.Mid(iPos_Star_End + 2);
		}
		else strLine = strLine.Left(iPos_Star_Start);
	}
	if(iPos_Star_End >= 0){};										//暂时不考虑

	return result;
}

BOOL EnumElement::DecomposeValue(CString strLine)			//分解类赋值语句，并填写对应的参数
{
	this->Empty();
	int iPos_1,iPos_2;					//位置
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

const DecEnum& DecEnum::operator=(const DecEnum& Other)			//赋值
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

void DecEnum::Empty()											//清空
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

	if(mLine.IsEmpty() == NULL )				//最后一个枚举元素后面没有逗号
	{
		pPos = pList->GetTailPosition();
		if(pPos)	pList->SetAt(pPos,mLine.GetTail().ComposeEnd());
	}
	pList->AddTail(CString("};"));
	return true;
}
/**************************************/
//	有三种类型的输入 
//	第一种：
//		XlYesNoGuess
//		xlGuess	0	Excel 确定是否有标题，如果有，是否是一个。
//		xlNo	2	默认值。 应对整个区域进行排序。
//		xlYes	1	不应对整个区域进行排序。
//	第二种：
//		XlYesNoGuess
//	第三种：
//		XlArabicModes
//		XLARABICMODES
//		名称	值
//		xlArabicBothStrict	3
//		xlArabicNone	0
//		xlArabicStrictAlefHamza	1
//		xlArabicStrictFinalYaa	2
/**************************************/
BOOL DecEnum::Decompose(CList<CString,CString&> * pList)		//将链表中的字符串，分解成一个枚举类
{
	this->Empty();						//清空

	EnumElement ElementNod;
	POSITION pPos;
	CString NodeCStr;

	pPos = pList->GetHeadPosition();
	while(pPos)
	{
		NodeCStr = pList->GetNext(pPos);
		ElementNod.DecomposeValue(NodeCStr);
		//取名称
		if(	this->mName.IsEmpty()		)	   			//当前枚举类的名称为空
		{
			if(	ElementNod.GetValue() == ENUM_NULL &&	//枚举元素 空值
				ElementNod.GetAnnotate().IsEmpty() )	//枚举元素 无注释
			{
				this->mName = ElementNod.GetName();		//枚举元素的名称就是枚举类的类名
			}
			else return false;							//当无枚举名时，但是枚举元素不是名称行时，是错误输入。
			continue;									
		}
		if(	ElementNod.GetValue() == ENUM_NULL &&		//枚举元素 空值
			ElementNod.GetAnnotate().IsEmpty() )		//枚举元素 无注释
		{
			continue;									//已有枚举类名，枚举元素空值，重复行，跳过
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
BOOL DecEnum::Marge(DecEnum* pOther)							//将pOther所指的枚举类合并到当前类中
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

void DecEnumList::Empty()										//清空
{
	this->mName.Empty();
	this->mLine.RemoveAll();
}

BOOL DecEnumList::IsEmpty()
{
	if(!this->mName.IsEmpty()) return false;
	return this->mLine.IsEmpty();
}


POSITION DecEnumList::Find(CString DecEnumName)						//查找枚举类
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
// 将文件中枚举类读出，并组成枚举类链表
BOOL DecEnumList::Load_CPP_H_File(CString FileName)						//装载.CPP或.h格式文件
{	
	//思路：
	//	1。逐行读，
	//	2。遇到枚举定义，进入枚举类读出循环，枚举定义的标志，“enum ****” 
	//	3。遇到枚举定义结束语句“};”,退出枚举类读出循环
	//	4。组成枚举类，加入链表。

	
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
		if(!ElementNode.DecomposeValue(LineCStr)) continue;	//不是正确的行，跳过
		if(ElementNode.GetName().IsEmpty()) continue;		//空行，跳过
		if(ElementNode_1.GetName().CompareNoCase(ElementNode.GetName()) == 0) continue;	//相同，跳过
		
		//当同时为类定义时，代表旧类结束，新类开始，
		if(ElementNode.GetValue() == ENUM_NULL) 
		{
			if(EnumNode.Decompose(&TemList))			//产生枚举类
			{
				this->mLine.AddTail(EnumNode);			//直接添加
			}
			ElementNode_1 = ElementNode;
			TemList.RemoveAll();
			TemList.AddTail(LineCStr);
		}
		else	TemList.AddTail(LineCStr);
	}
	EnumNode.Decompose(&TemList);						//产生枚举类
	if(!EnumNode.GetName().IsEmpty())					//不是空类时，添加
		this->mLine.AddTail(EnumNode);
	myFile.Close();	

	Marge();											//删除重复项
	return true;
}

///*******************************************
//
//
//
//
///*******************************************
BOOL DecEnumList::Load_Txt_File(CString FileName)						//装载.txt格式文件
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
		if(!ElementNode.DecomposeValue(LineCStr))	continue;					//不是正确的行，跳过
		if(ElementNode.GetName().IsEmpty())			continue;					//空行，跳过
		if(ElementNode_1.GetName().CompareNoCase(ElementNode.GetName()) == 0)
			continue;															//相同，跳过
		
		//当同时为类定义时，代表旧类结束，新类开始，
		if(ElementNode.GetValue() == ENUM_NULL) 
		{
			if(EnumNode.Decompose(&TemList))			//产生枚举类
			{
				this->mLine.AddTail(EnumNode);			//直接添加
			}
			ElementNode_1 = ElementNode;
			TemList.RemoveAll();
			TemList.AddTail(LineCStr);
		}
		else	TemList.AddTail(LineCStr);
	}
	EnumNode.Decompose(&TemList);						//产生枚举类
	if(!EnumNode.GetName().IsEmpty())					//不是空类时，添加
		this->mLine.AddTail(EnumNode);
	myFile.Close();	

	Marge();											//删除重复项
	return true;
}

BOOL DecEnumList::Create_H(CString FileName)						//产生.H格式文件
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
void DecEnumList::Marge()											//枚举类合并
{
	//思路：1。从后向前取待合并的枚举类
	//		2。将待合并项与链表中的类相比较，比较时的扫描是从前向后。
	//      3。如果是重复枚举类，则合并，并更新链表，同时删除待合并的枚举类
	POSITION pPos,pOldPos,pMargePos;	
	DecEnum EnumNode,MargeNode;

	pPos = mLine.GetTailPosition();						//从后向前开始扫描
	while(pPos)
	{
		pOldPos  = pPos;
		EnumNode = mLine.GetPrev(pPos);					//取待合并的枚举类
		if(EnumNode.IsEmpty())							//是不是空类
		{
			mLine.RemoveAt(pOldPos);					//空类，删除
			continue;									//跳过下面的循环段。
		}
	
		pMargePos= this->Find(EnumNode.GetName());		//寻找是从前向后扫描
		if(pOldPos == pMargePos) continue;				//相同，代表没有重复项，不进行合并操作

		//有相同，则合并
		//
		MargeNode = mLine.GetAt(pMargePos);				//取重复的枚举类			
		MargeNode.Marge(&EnumNode);						//合并
		mLine.SetAt(pMargePos,MargeNode);				//更新链表
		mLine.RemoveAt(pOldPos);						//删除待合并的枚举类
	}
}