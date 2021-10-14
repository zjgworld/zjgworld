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
	const ExMethod& operator=(const ExMethod& Other);//赋值
	void Empty();												//清空
	BOOL IsEmpty();												//判断是否为空

	void SetName(CString nName){mName = nName;};				//函数名
	CString GetName(){return mName;};							//函数名
	void SetOutType(CString n_OutType){m_OutType = n_OutType;};	//函数返回的数据类型
	CString GetOutType(){return m_OutType;};					//函数返回的数据类型
	void SetInType(CString n_InType){m_InType = n_InType;};		//函数的输入参数的数据类型
	CString GetInType(){return m_InType;};						//函数的输入参数的数据类型
	void SetVARTYPE(CString n_VARTYPE){m_VARTYPE = n_VARTYPE;};	//InvokeHelper函数的返回类型
	CString GetVARTYPE(){return m_VARTYPE;};					//InvokeHelper函数的返回类型
	void SetDISPID(CString n_DISPID){m_DISPID = n_DISPID;};		//InvokeHelper函数的接口类型
	CString GetDISPID(){return m_DISPID;};						//InvokeHelper函数的接口类型

	CString Get_VARENUM_Type(CString nVARENUM);		//获取VARENUM对应的数据类型
	CString Get_Type_VARENUM(CString nVARENUM);		//获取数据类型对应VARENUM的值

	CString GetMethodType();						//获取InvokeHelper函数的返回类型
	CString GetID();								//通过分解InvokeHelper函数，获取函数接口类型代码

	BOOL DecomposeIvHp();							//分解InvokeHelper语句，并填写对应的参数
	BOOL DecomposeIvHp(CString strLine);			//分解InvokeHelper语句，并填写对应的参数
	BOOL DecomposeMeth(CString strLine);			//分解函数定义语句，并填写函数名
	void Modify_Method(CList<CString,CString&> * pList);	//以当前方法为参考类，对于pList所指向的语句进行修改
	BOOL Modify_Method_PUT(CList<CString,CString&> * pList);//以当前方法为参考类，对于pList所指向的语句进行修改

public:
	CList<CString,CString&> mLine;
private:
	CString mName;						//函数名
	CString m_OutType;					//函数的返回类型，与m_VARTYPE有对应关系
	CString m_VARTYPE;					//InvokeHelper函数中的第三个参数，返回数据类型
	CString m_DISPID;					//InvokeHelper函数的接口类型，为一组数据
	CString m_InType;					//函数的输入参考的数据类型

// void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ... );

};

class ExClass
{
public:
	ExClass();
	ExClass(const ExClass& Other);
	~ExClass();

	const ExClass& operator=(const ExClass& Other);		//赋值
	void Empty();										//清空
	BOOL IsEmpty();										//判断是否是空
	
	void SetName(CString nName){mName = nName;};		//设置类名
	CString GetName(){return mName;};					//获取类名

	CString GetType(CString dwDispID);					//根据方法(函数)的标记，获取函数返回值对应的数据类型
	CString Get_VARENUM_Type(CString nVARENUM);			//获取VARENUM对应的数据类型
	CString Get_Type_VARENUM(CString nVARENUM);			//获取数据类型对应VARENUM的值

	POSITION Find_PUT(CString dwDispID);				//根据dwDispID，查找对应的接口的PUT函数
	void FindDispID(CString dwDispID,CList<ExMethod,ExMethod&>* pLine);			//根据InvokeHelper函数的接口类型，查找对应的方法

	int Modify_Method(CList<CString,CString&> * pList,CString * pstrMethod);	//以当前类为参考类，对于pList所指向的语句进行修改
	int Modify_Method(ExMethod * pMethod);				//以当前类为参考类，对(*pMethod)方法进行修改

	void DecomposeClassName(CString strLine);			//分解类定义语句，并填写对应的参数

public:
	CList<ExMethod,ExMethod&> mLine;
private:
	CString mName;		//类名
	CString mBase;		//基类名
};

#define FILE_FINISHED		(0x0000)		//正常完成
#define FILE_NAME_ERROR		(0x0001)		//名称错误
#define FILE_READ_ERROR		(0x0002)		//读错误
#define FILE_WRITE_ERROR	(0x0004)		//写错误
#define FILE_LESS_ERROR		(0x0008)		//缺少原类
#define FILE_PART_ERROR		(0x0010)		//部分错误
#define FILE_VARTYPE_ERR	(0x0020)		//变量类型错误
#define FILE_C_LESSID_ERR	(0x0040)		//类中缺少方法接口ID错误
#define FILE_LN_LESSID_ERR	(0x0080)		//链表中缺少方法名相近的接口ID错误
#define FILE_LI_LESSID_ERR	(0x0100)		//链表中缺少接口ID错误
#define FILE_C_LESSPUT_ERR	(0x0200)		//类中缺少成对的ID错误


#define FILE_ERROR_ERR		(0x8000)				//文件已损坏的错误

class ExClassLine
{
public:
	ExClassLine();
	~ExClassLine();

public:
	void Empty();
	void SetFileName(CString nName){mFileName = nName;};
	CString GetFileName(){return mFileName;};

	
	BOOL LoadCPPFile(CString FileName);			//装载.CPP格式文件
	int Load_H_File(CString FileName);			//装载.h格式文件

	POSITION Find(CString nClassName);			//根据类名，查找类的位置
	POSITION Find_A(CString nClassName);		//根据类名和相似类名，查找类的位置
	POSITION Find_Method_PUT(ExMethod* pMethod);//根据方法名和dwDispID，查找对应的接口的PUT函数

	int Modify_H_File_A(CString OldFName,CString NewFName);	//修改.h格式文件
	int Modify_H_File_PUT(CString OldFileName,CString NewFileName);				//修改.h格式文件											
	                                            //返回值，1.文件名不正确。2.文件打不开。3.文件不能写。
	int Modi_Method_Name(ExMethod* pMethod);	//基于函数名和ID相同的原则进行修改
	int Modi_Method_ID(ExMethod* pMethod);		//基于ID相同的原则进行修改

	BOOL FillCBCtrl(CComboBox * pCBCtrl);							//将参考用类名填充到组合框中
	BOOL FillListCtrl(CListBox * pListCtrl,CString strFileName);	//将修改出现错误的文件名填充到列表
	BOOL FillStaticCtrl(CString * pCStrCtrl,CString strFileName);	//将修改出现错误的文件名填充到列表

	BOOL Save(CString FileName);

private:
	
public:
	CList<ExClass,ExClass&> mLine;
private:
	CString mFileName;


};

#define ENUM_NULL	0x00FFFFFA		//枚举空值

#define ENUM_NOTHING				0x00000000		//语句中没有相应的单词
#define ENUM_HAVE_ENUM				0x00000001		//语句中有 "emun"单词
#define ENUM_HAVE_LEFT				0x00000002		//语句中有 "{"左括号
#define ENUM_HAVE_RIGHT				0x00000004		//语句中有 "}"右括号
#define ENUM_HAVE_SEMICOLON			0x00000008		//语句中有 ";"分号
#define ENUM_HAVE_RIGHTSEMICOLON	0x00000010		//语句中 "}"和";"相连，或中间只有空格或TAB键

class EnumElement
{
public:
	EnumElement();
	~EnumElement();
	
	const EnumElement& operator=(const EnumElement& Other);			//赋值
	void Empty();													//清空
	BOOL IsEmpty();													//判断是否是空

	CString GetName(){return mName;};								//成员名
	void SetName(CString nName){mName = nName;};					//成员名
	int GetValue(){return mValue;};									//值
	void SetValue(int nValue){mValue = nValue;};						//值
	BOOL SetValue(CString nName);									//值
	CString GetAnnotate(){return mAnnotate;};						//注释
	void SetAnnotate(CString nAnnotate){mAnnotate = nAnnotate;};	//注释

	int Decompose_CPP(CString strLine);			//检测并分解.CPP和.h文件中枚举定义语句
	BOOL DecomposeValue(CString strLine);			//分解类赋值语句，并填写对应的参数
	CString Compose();
	CString ComposeEnd();

private:
	CString mName;						//成员名
	int mValue;							//值
	CString mAnnotate;					//注释
};

class DecEnum
{
public:
	DecEnum();
	~DecEnum();
	
	const DecEnum& operator=(const DecEnum& Other);			//赋值
	void Empty();											//清空
	BOOL IsEmpty();	

	CString GetName(){return mName;};
	void SetName(CString nName){mName = nName;};

	POSITION Find(CString strElementName);
	POSITION Find(int iElementValue);

	BOOL Compose(CList<CString,CString&> * pList);			//组合成枚举类型的字符串链表
	BOOL Decompose(CList<CString,CString&> * pList);		//将链表中的字符串，分解成一个枚举类
	BOOL Marge(DecEnum* pOther);							//将pOther所指的枚举类合并到当前类中
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

	void Empty();											//清空
	BOOL IsEmpty();

	POSITION Find(CString DecEnumName);						//查找枚举类

	BOOL Load_CPP_H_File(CString FileName);					//装载.CPP或.h格式文件
	BOOL Load_Txt_File(CString FileName);						//装载.txt格式文件
	BOOL Create_H(CString FileName);						//产生.H格式文件
	void Marge();											//枚举类合并

public:
	CList<DecEnum,DecEnum&> mLine; 
private:
	CString mName;											//文件名

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

/*
类型名		含义     
VT_EMPTY	指示未指定值    
VT_NULL		指示空值（类似于  SQL  中的空值）    
VT_I2		指示  short  整数    
VT_I4		指示  long  整数    
VT_R4   	指示  float  值    
VT_R8   	指示  double  值    
VT_CY   	指示货币值    
VT_DATE   	指示  DATE  值    
VT_BSTR   	指示  BSTR  字符串    
VT_DISPATCH	指示  IDispatch  指针    
VT_ERROR   	指示  SCODE   
VT_BOOL   	指示一个布尔值    
VT_VARIANT 	指示  VARIANTfar  指针    
VT_UNKNOWN 	指示  IUnknown  指针    
VT_DECIMAL 	指示  decimal  值 
VT_UI1   	指示  byte   
VT_UI2   	指示  unsignedshort   
VT_UI4   	指示  unsignedlong   
VT_I8   	指示  64  位整数    
VT_UI8   	指示  64  位无符号整数    
VT_INT   	指示整数值    
VT_UINT   	指示  unsigned  整数值    
VT_VOID   	指示  C  样式  void   
VT_HRESULT  指示  HRESULT   
VT_PTR   	指示指针类型    
VT_SAFEARRAY   	指示  SAFEARRAY   
VT_CARRAY   	指示  C  样式数组    
VT_USERDEFINED	指示用户定义的类型    
VT_LPSTR   		指示一个以  NULL  结尾的字符串    
VT_LPWSTR   	指示由  nullNothingnullptrnull   引用（在  Visual Basic   中为  Nothing ）   终止的宽字符串    
VT_RECORD   	指示用户定义的类型    
VT_FILETIME   	指示  FILETIME  值    
VT_BLOB   		指示以长度为前缀的字节    
VT_STREAM   	指示随后是流的名称    
VT_STORAGE   	指示随后是存储的名称    
VT_STREAMED_OBJECT	指示流包含对象    
VT_STORED_OBJECT	指示存储包含对象    
VT_BLOB_OBJECT		指示  Blob  包含对象    
VT_CF   	指示剪贴板格式    
VT_CLSID   	指示类ID   
VT_VECTOR   指示简单的已计数数组    
VT_ARRAY   	指示  SAFEARRAY  指针    
VT_BYREF   	指示值为引用 
————————————————
版权声明：本文为CSDN博主「programmerES」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/superlym2005/article/details/5838356
 */

#endif // !defined(AFX_STDAFX_H__BA14C3D0_C5E8_4D24_90B8_DED6A84AEF10__INCLUDED_)
