#include "stdafx.h"
#include "zjgBasic.h"

#include<stdio.h> 
#include <Windows.h> 
#include "math.h"

#include <AFXDLGS.h>
#include <AFX.h>

#include <WinSpool.h>     
#pragma comment(lib, "Winspool.lib")     
#include <shellapi.h>     
#pragma comment(lib, "shell32.lib")     

using namespace ZJGName;

CString ZJGName::DelDecimal(CString DecimalCStr)//删除字符型数值的小数部分
{
	if(DecimalCStr.Find(".",0)>0)
		return DecimalCStr.Left(DecimalCStr.Find(".",0));
	else 
		return DecimalCStr;
}

BOOL ZJGName::CStrIsNumber(CString InputCStr)//判断是否是数字
{
	BOOL bResult = CStrIsNumber(InputCStr.GetBuffer(InputCStr.GetLength()+1));
	InputCStr.ReleaseBuffer();
	return bResult;
}

BOOL ZJGName::CStrIsNumber(char * InputChar)//判断是否是数字
{
	StrDeleteSpace(InputChar);//删除首尾空格
	int i = 0;
	int iDot = 0;
	while(InputChar[i] == ' ' || InputChar[i] == '\t')
	{	//在最前面允许有空格和TAB
		i++;
	}
	if(	InputChar[i] == '+' ||
		InputChar[i] == '-')
		i++;	//只允许一个'+'号或'-'号
	while(InputChar[i] != '\0')
	{
		if(InputChar[i] >= '0' && InputChar[i] <= '9')
			i++;
		else if(InputChar[i] == '.' && iDot == 0)
		{//只允许一个小数点
			i++;
			iDot++;
		}
		else return FALSE;
		if(i > 1024) return FALSE;//字条长度超过1024，可能字符串出问题了，
	}
	return TRUE;
}

CString ZJGName::StrDeleteSpace(CString InputCStr,char DeleteChar)//去除首尾的空格
{
	//字符长度不超过256个。否则截长为256个字符。
	int i = 0 ;
	int iTotal = InputCStr.GetLength();
	//char DeleteChar = ' ';
	while((i< iTotal) && (InputCStr.GetAt(i) == DeleteChar))
	{
		i++;
	}
	InputCStr = InputCStr.Mid(i);
	i = InputCStr.GetLength() -1;
	while((i>= 0) && (InputCStr.GetAt(i) == DeleteChar))
	{
		i--;
	}
	return InputCStr.Left(i+1);
}
void ZJGName::StrDeleteSpace(char * InputChar,char DeleteChar)//去除首尾的空格
{
	//字符长度不超过256个。否则截长为256个字符。
	int i = 0;//删除前的指针
	int j = 0;//删除后的指针
	//char DeleteChar = ' ';
	//去除头部的空格
	while(InputChar[i] != '\0')
	{
		if(InputChar[i] != DeleteChar) break;//不是空格时跳出
		if( i >= 256)  InputChar[i+1] = '\0'; 
		i++;

	}
	j = 0;//删除后的指针
	while(InputChar[i] != '\0')
	{
		InputChar[j] = InputChar[i];
		j++;
		i++;
	}
	InputChar[j] = '\0';//加结束符
	//去除尾部的空格
	while(j >=0 )
	{
		if(InputChar[j] != DeleteChar) break;//不是空格时跳出
		InputChar[j] = '\0';//加结束符
		j--;
	}
}

CString ZJGName::StrDeleteEndChar(CString InputCStr,char DeleteChar)//去除尾部的指定字符
{
	//字符长度不超过256个。否则截长为256个字符。
	int i = 0 ;
	i = InputCStr.GetLength() -1;
	while((i>= 0) && (InputCStr.GetAt(i) == DeleteChar))
	{
		i--;
	}
	return InputCStr.Left(i+1);
}

void ZJGName::StrDeleteEndChar(char * InputChar,char DeleteChar)//去除尾部的指定字符
{
	//字符长度不超过256个。否则截长为256个字符。
	int i = 0;//删除前的指针
	int j = 0;//删除后的指针
	//char DeleteChar = ' ';
	//计算字符串的长度
	while(InputChar[i] != '\0')
	{
		if( i >= 256)  InputChar[i+1] = '\0'; 
		i++;
	}
	//去除尾部的空格
	while(i >=0 )
	{
		if(InputChar[i] != DeleteChar) break;//不是空格时跳出
		InputChar[i] = '\0';//加结束符
		i--;
	}
}
CString ZJGName::NumberToLetter(double Number)//将数值型数字转换成英文字符型数值
{
	long lNumber = (long)Number;
	long fNumber = (long)((Number - lNumber + 0.0049)*100);//四舍五入
	if(fNumber > 0)
		return	NumberToLetter(lNumber) + " AND CENTS " + NumberToLetter(fNumber);
	else return	NumberToLetter(lNumber);
}

CString ZJGName::NumberToLetter(long Number)//将数值型数字转换成英文字符型数值
{
	CString NumberLetter;
	NumberLetter.Empty();
	if(Number > 1000000000 || Number < 0)
	{
		return NumberLetter;
	}
	if(Number >= 1000000)
	{
		return NumberToLetter(long(Number / 1000000)) + "MILLION " + NumberToLetter(long(Number % 1000000));
	}
	if(Number >= 1000)
	{
		return NumberToLetter(long(Number / 1000)) + "THOUSAND " + NumberToLetter(long(Number % 1000));
	}
	if(Number >= 100)
	{
		return NumberToLetter(long(Number / 100)) + "HUNDRED " + NumberToLetter(long(Number % 100));
	}
	if( ((long (Number / 10)) * 10 == Number) && (Number >= 20))//大于20的整十位数 
	{
		switch(Number / 10)
		{
		case 2: return "TWENTY ";
		case 3: return "THIRTY ";
		case 4: return "FORTY ";
		case 5: return "FIFTY ";
		case 6: return "SIXTY ";
		case 7: return "SEVENTY ";
		case 8: return "EIGHTY ";
		case 9: return "NINETY ";
		}
	}
	if(Number < 20)
	{
		switch(Number)
		{
		case 0: return "";
		case 1: return "ONE ";
		case 2: return "TWO ";
		case 3: return "THREE ";
		case 4: return "FOUR ";
		case 5: return "FIVE ";
		case 6: return "SIX ";
		case 7: return "SEVEN ";
		case 8: return "EIGHT ";
		case 9: return "NINE ";
		case 10: return "TEN ";
		case 11: return "ELEVEN ";
		case 12: return "TWELVE ";
		case 13: return "THIRTEEN ";
		case 14: return "FOURTEEN ";
		case 15: return "FIFTEEN ";
		case 16: return "SIXTEEN ";
		case 17: return "SEVENTEEN ";
		case 18: return "EIGHTEEN ";
		case 19: return "NINETEEN ";
		default: return "";
		}
	}
	else return  StrDeleteSpace(NumberToLetter(long( Number / 10 * 10))) + "-" + NumberToLetter(long( Number%10));
	
}
int ZJGName::CutNameAdd(CString NameAddCStr,int MaxChar)//
{
	int CurrentIndex = MaxChar -1;//当前位置
	int TotalCharNumber = 0;

	if(NameAddCStr.GetAt(CurrentIndex) <0)//是汉字
	{		
		while(CurrentIndex >=0)
		{
			if(NameAddCStr.GetAt(CurrentIndex) > 0)//非汉字时
				break;
			CurrentIndex--;
			TotalCharNumber++;
		}
		if((TotalCharNumber % 2) == 1)//单数
			return MaxChar-1;
		else //双数
			return MaxChar;
	}
	if((NameAddCStr.GetAt(CurrentIndex) >= '0') && ( NameAddCStr.GetAt(CurrentIndex) <= '9'))//是数字
	{
		while((CurrentIndex >=0)&&(NameAddCStr.GetAt(CurrentIndex) >= '0') && ( NameAddCStr.GetAt(CurrentIndex) <= '9'))//是数字
			CurrentIndex--;
		return CurrentIndex + 1 ;
	}
	while(CurrentIndex >= 0)
	{	
		if(	NameAddCStr.GetAt(CurrentIndex) == ' ' ||//当空格时
			NameAddCStr.GetAt(CurrentIndex) == ',' ) //当','时
				break;
		CurrentIndex--;
	}
	if(CurrentIndex * 2 > MaxChar)	return CurrentIndex;
	CurrentIndex = MaxChar -1;
	while(CurrentIndex > 1)
	{	
		if( ((NameAddCStr.GetAt(CurrentIndex) >= 'a') && ( NameAddCStr.GetAt(CurrentIndex) <= 'z')) ||
			((NameAddCStr.GetAt(CurrentIndex) >= 'A') && ( NameAddCStr.GetAt(CurrentIndex) <= 'Z')) ||
			((NameAddCStr.GetAt(CurrentIndex) >= '0') && ( NameAddCStr.GetAt(CurrentIndex) <= '9')) )
		{
			CurrentIndex--;
		}
		else break;
	}
	return CurrentIndex;
	
}
int ZJGName::CutNameAdd(TCHAR * NameAddCStr,int MaxChar)//
{
	int CurrentIndex = MaxChar -1;//当前位置
	int TotalCharNumber = 0;

	if(NameAddCStr[CurrentIndex] <0)//是汉字
	{		
		while(CurrentIndex >=0)
		{
			if(NameAddCStr[CurrentIndex] > 0)//非汉字时
				break;
			CurrentIndex--;
			TotalCharNumber++;
		}
		if((TotalCharNumber % 2) == 1)//单数
			return MaxChar-1;
		else //双数
			return MaxChar;
	}
	if((NameAddCStr[CurrentIndex] >= '0') && ( NameAddCStr[CurrentIndex] <= '9'))//是数字
	{
		while((CurrentIndex >=0)&&(NameAddCStr[CurrentIndex] >= '0') && ( NameAddCStr[CurrentIndex] <= '9'))//是数字
			CurrentIndex--;
		return CurrentIndex + 1 ;
	}
	while(CurrentIndex >= 0)
	{	
		if(	NameAddCStr[CurrentIndex] == ' ' ||//当空格时
			NameAddCStr[CurrentIndex] == ',' ) //当','时
				break;
		CurrentIndex--;
	}
	if(CurrentIndex * 2 > MaxChar)	return CurrentIndex;
	CurrentIndex = MaxChar -1;
	while(CurrentIndex > 1)
	{	
		if( ((NameAddCStr[CurrentIndex] >= 'a') && ( NameAddCStr[CurrentIndex] <= 'z')) ||
			((NameAddCStr[CurrentIndex] >= 'A') && ( NameAddCStr[CurrentIndex] <= 'Z')) ||
			((NameAddCStr[CurrentIndex] >= '0') && ( NameAddCStr[CurrentIndex] <= '9')) )
		{
			CurrentIndex--;
		}
		else break;
	}
	return CurrentIndex;
}

CString ZJGName::MonthIToCStr(int iMonth,BOOL Simplify)//将数字月份转换成英文单词
{
	switch(iMonth)
	{
	case 1:  if(Simplify) return CString("Jan");
		     else         return CString("January");
			 break;
	case 2:  if(Simplify) return CString("Feb");
		     else         return CString("February");
			 break;
	case 3:  if(Simplify) return CString("Mar");
		     else         return CString("March");
			 break;
	case 4:  if(Simplify) return CString("Apr");
		     else         return CString("April");
			 break;
	case 5:  if(Simplify) return CString("May");
		     else         return CString("May");
			 break;
	case 6:  if(Simplify) return CString("Jun");
		     else         return CString("June");
			 break;
	case 7:  if(Simplify) return CString("Jul");
		     else         return CString("July");
			 break;
	case 8:  if(Simplify) return CString("Aug");
		     else         return CString("August");
			 break;
	case 9:  if(Simplify) return CString("Sept");
		     else         return CString("September");
			 break;
	case 10:  if(Simplify) return CString("Oct");
		     else         return CString("October");
			 break;
	case 11:  if(Simplify) return CString("Nov");
		     else         return CString("November");
			 break;
	case 12:  if(Simplify) return CString("Dec");
		     else         return CString("December");
			 break;
	default: CString TemCStr;
		TemCStr.Empty();
		return TemCStr;
	}
}
/*************************************************
Function    : ShowJPG
Description : 在DC上按图片原始尺寸显示JPG图片
Calls       : 
Called By   : 
Parameter   : [CDC* pDC] --- DC
            : [CString strPath] --- 要显示的图片路径，建议全路径
            : [int x] --- DC上显示的X位置 
            : [int y] --- DC上显示的Y位置
            : [bool OriginalSize] --- 是否按图片原始尺寸显示,false时将按DC大小缩放
Return      : bool --- 是否成功
Author      : Unknown
Date        : 2008-10-24
Modify      : 
*************************************************/
BOOL ZJGName::ShowJpg(CDC* pDC, CString strPath,CRect mRect) 
{   
	IStream *pStm;
	CFileStatus fstatus;
	CFile file;
	LONG cb;
	if (file.Open(strPath,CFile::modeRead)&&
		file.GetStatus(strPath,fstatus)&&
		((cb = fstatus.m_size) != -1))
	{      

		HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, cb); //从堆中分配一定数目的字节   
		LPVOID pvData = NULL;      
		if (hGlobal != NULL)
		{      
			pvData = GlobalLock(hGlobal);
			if (pvData != NULL)      
			{      
				file.Read(pvData, cb);      
				GlobalUnlock(hGlobal);      
				CreateStreamOnHGlobal(hGlobal, TRUE, &pStm);
			}
		}
	} 
	else
	{  
		return false;   
	}             
	IPicture *pPic;
	if(SUCCEEDED(OleLoadPicture(pStm,fstatus.m_size, TRUE,IID_IPicture,(LPVOID*)&pPic)))
	{       
		OLE_XSIZE_HIMETRIC jpgWidth;
		OLE_YSIZE_HIMETRIC jpgHeight;
		pPic->get_Width(&jpgWidth);
		pPic->get_Height(&jpgHeight);
		double fX,fY,fScale;      
		fX = (double)pDC->GetDeviceCaps(HORZRES)*
			 (double)jpgWidth /
			 ((double)pDC->GetDeviceCaps(HORZSIZE)*100.0);
		fY = (double)pDC->GetDeviceCaps(VERTRES)*
			 (double)jpgHeight/
			 ((double)pDC->GetDeviceCaps(VERTSIZE)*100.0);  
		
		DWORD mWidth = mRect.Width();
		DWORD mHeight = mRect.Height();
		DWORD Display_x;
		DWORD Display_y;

		if (fX < mWidth && fY < mHeight)
		{		//图片小于边框，按图片来显示，按边框的中心位置来计算图片显示位置点
			Display_x = mRect.CenterPoint().x - (DWORD)(fX / 2); //边框的中心位置
			Display_y = mRect.CenterPoint().y - (DWORD)(fY / 2); //边框的中心位置
			pPic->Render(*pDC,Display_x,Display_y,(DWORD)fX,(DWORD)fY,0,jpgHeight,jpgWidth,-jpgHeight,NULL); 
		}
		else	//图片大于边框，按边框来显示，按图片的高宽比来计算图片显示位置点及宽高
		{
			if( (fX / mWidth) > (fY / mHeight) )
			{		//图片宽度比高度大时，按宽度之比计算
				fScale = fX / mWidth;

			}
			else	//图片高度比长度大时，按高度之比计算
			{
				fScale = fY / mHeight;
			}
			mWidth	= (DWORD)(fX / fScale);
			mHeight	= (DWORD)(fY / fScale);
			Display_x = mRect.CenterPoint().x - (DWORD)(mWidth / 2); //边框的中心位置
			Display_y = mRect.CenterPoint().y - (DWORD)(mHeight / 2); //边框的中心位置
			pPic->Render(*pDC,Display_x,Display_y,mWidth,mHeight,0,jpgHeight,jpgWidth,-jpgHeight,NULL);
		}

		{		//图片大于边框，按边框来显示
//			pPic->Render(*pDC,mRect.left,mRect.top,mRect.Width(),mRect.Height(),0,jpgHeight,jpgWidth,-jpgHeight,NULL); 
			 
		}     
	}      
	return true;   
}
/****************************
功能：删除文件中每行指定字符前的内容
****************************/
BOOL ZJGName::DeleteFileCStr(CString mDelCStr)
{
	CFileDialog myFileDialog(TRUE,NULL,NULL,OFN_FILEMUSTEXIST,"C++File(*.cpp)|*.cpp|Head File(*.h)|*.h||");//TRUE:打开，FALSE：假
	//CFileDialog myFileDialog(TRUE);
	CString rFileName;
	CString wFileName;


	myFileDialog.DoModal();
	rFileName = myFileDialog.GetPathName();
	wFileName = myFileDialog.GetFileExt();
	wFileName = rFileName.Left(rFileName.Find(".")) +"-1." + wFileName;
	
	CStdioFile fileRead;///只读方式打开
	CStdioFile fileWrite;///只读方式打开
	CString ReadLineCStr;
    CString WriteLineCStr;
	///CFile::modeRead可改为 CFile::modeWrite(只写),CFile::modeReadWrite(读写),CFile::modeCreate(新建)
	if(!fileRead.Open(rFileName,CFile::modeRead)) return false;
	if(!fileWrite.Open(wFileName,CFile::modeCreate|CFile::modeWrite)) return false;

	int mDelCStrLength = mDelCStr.GetLength();//字符长度
	while(fileRead.ReadString(ReadLineCStr))
	{
		WriteLineCStr = ReadLineCStr.Mid(ReadLineCStr.Find(".") + mDelCStrLength) + "\r\n";
		fileWrite.WriteString(WriteLineCStr);
	}
	fileRead.Close();
	fileWrite.Close();
	return true;
}

BOOL ZJGName::IsDebug()		//判断当前可执行文件是否是调试文件目录 
{
	TCHAR SourceDir[256];
	::memset(SourceDir,256,'\0');
	GetModuleFileName(NULL,SourceDir,255);
	CString SourceFile(SourceDir);
	if(SourceFile.Find("\\Debug") >= 0)
		return true;
	else return false;
}

CString ZJGName::GetExeFileDir()//获取程序当前执行文件目录 
{
	TCHAR SourceDir[256];
	::memset(SourceDir,256,'\0');
	GetModuleFileName(NULL,SourceDir,255);
	CString SourceFile(SourceDir);
	int pPos = SourceFile.Find("\\Debug");
	if(pPos > 0)
	{
		SourceFile = SourceFile.Left(pPos + 1);
	}
	else
	{
		SourceFile.MakeReverse();
		SourceFile = SourceFile.Mid(SourceFile.Find("\\"));
		SourceFile.MakeReverse();
	}
	return SourceFile;
}

/****************
功能：检测文件夹是否存在
传入：文件夹名，
返回：真：存在
      假：不存在
*/
BOOL ZJGName::CheckDir(CString DirName) 
{
	DWORD dwAttr=GetFileAttributes(DirName);
	if(dwAttr==0xFFFFFFFF)  //临时文件夹不存在   
	{  
		return false;
	}
	return true;
}
/****************
功能：创建文件夹
传入：文件夹名，
返回：真：创建成功
      假：创建失败
*/
BOOL ZJGName::CreateDir(CString DirName) 
{
	TRY
	{
		int iStart;
		//检测最后一个字符是否是"\"，如果是，删除最后一个字符
		if(DirName.Right(1) == "\\")
			DirName = DirName.Left(DirName.GetLength() -1);
		iStart = 0;
		CString SubDirName;
		//以下循环是从根目录开始检查，是否存在，不存在就创建。
		while(iStart >= 0)
		{
			iStart = DirName.Find("\\",iStart+1);
			if(iStart >= 0)
				SubDirName = DirName.Left(iStart);
			else
				SubDirName = DirName;
			if(!CheckFile(SubDirName))	//不存在时，创建
			{
				if(SubDirName.Find("\\") < 0) 
					continue;//根目录不能创建

				if(!CreateDirectory(SubDirName,NULL)) //不存在时，创建
					return false;
			}
		}
		return true;  
	}
	CATCH(CFileException,e)
	{
		#ifdef _DEBUG
			afxDump<<"File could not be create"<<e->m_cause<<"\n";			
		#endif
		return false;
	}
	END_CATCH
}
/****************
功能：检测文件是否存在
传入：文件名，
返回：真：存在
      假：不存在
*/
BOOL ZJGName::CheckFile(CString FileName)		//检查文件是否存在
{
	CFileFind m_FileFind;
	return m_FileFind.FindFile(FileName);	
}

//得到列表控件的标头字符串内容
void GetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray) 
{
	LVCOLUMN lvcol;
    int   nColNum;
	char str[256];
    lvcol.mask = LVCF_TEXT;
	lvcol.cchTextMax = 256;
	lvcol.pszText = str;
	int nColCound = pListCtrl->GetHeaderCtrl()->GetItemCount();

	pHeaderArray->RemoveAll();
	pHeaderArray->SetSize(nColCound+1);//设置数组大小
	nColNum = 0;
	CString TemCStr;
	while(nColNum < nColCound)
	{
		pListCtrl->GetColumn(nColNum, &lvcol);
		TemCStr = lvcol.pszText;
        pHeaderArray->SetAt(nColNum,TemCStr);
		nColNum++;
    }
}

void ZJGName::GetListCtrlHeader(CListCtrl * pListCtrl,CList<CString,CString> * pHeaderList,int iStart,int iEnd)//得到列表控件的标头字符串内容
{
	LVCOLUMN lvcol;
    int   nColNum;
	char str[256];
    lvcol.mask = LVCF_TEXT;
	lvcol.cchTextMax = 256;
	lvcol.pszText = str;
	int nColCound = pListCtrl->GetHeaderCtrl()->GetItemCount();

	if(iEnd < 0) iEnd =  nColCound - 1;
	if(iStart < 0) iStart = 0;
	nColNum = iStart;
	pHeaderList->RemoveAll();
	CString TemCStr;
	while(nColNum <= iEnd)
	{
		pListCtrl->GetColumn(nColNum, &lvcol);
		pHeaderList->AddTail(CString(lvcol.pszText));
		nColNum++;
    }
}

void ZJGName::SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray,int iSizeCount)//设置列表控件的标头字符串内容
{
	LVCOLUMN lvcol;
    char  str[256];
    int   nColNum = 0;

    lvcol.mask = LVCF_TEXT;
    lvcol.cchTextMax = 256;
	lvcol.pszText = str;
	
	CString TemCStr;
	while(nColNum < iSizeCount)
    {
		TemCStr = pHeaderArray->GetAt(nColNum);
		::strcpy(lvcol.pszText,TemCStr.operator LPCTSTR());

        pListCtrl->SetColumn(nColNum, &lvcol);
		nColNum++;
    }
}
/*************************************************************************
 *
 * \函数名称：
 *   SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CList<CString,CString&> * pHeaderList)
 *
 * \输入参数:
 *		CListCtrl				* pListCtrl		列表框控件指针
 *		CList<CString,CString&> * pHeaderList	标头字符串表
 * \返回值:
 *		设置成功，返回真，否则为假	
 *
 * \说明:
 *   设置列表控件的标头字符串内容,依次将表中字符串填写到标题栏中
 *
 ************************************************************************
 */
/*
void ZJGName::SetListCtrlHeader(CListCtrl * pListCtrl,CList<CString,CString&> * pHeaderList)//设置列表控件的标头字符串内容
{
	LVCOLUMN lvcol;
    char  str[256];
    int   nColNum = 0;

    lvcol.mask = LVCF_TEXT;
    lvcol.cchTextMax = 256;
	lvcol.pszText = str;
	
	CString TemCStr;
	POSITION pPos;
	pPos = pHeaderList->GetHeadPosition();
	while(pPos)
    {
		TemCStr =pHeaderList->GetNext(pPos);
		::strcpy(lvcol.pszText,TemCStr.operator LPCTSTR());

        pListCtrl->SetColumn(nColNum, &lvcol);
		nColNum++;
    }
}
*/
/*************************************************************************
 *
 * \函数名称：
 *   SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString&>& HeaderCStr,CArray<int,int>& iHeader)
 *
 * \输入参数:
 *		CArray<CString,CString&>&	HeaderCStr  标头字符串集合
 *		CArray<int,int>&			iHeader     标头字符串对应列号集合
 * \返回值:
 *		设置成功，返回真，否则为假	
 *
 * \说明:
 *   设置列表控件的标头字符串内容
 *
 ************************************************************************
 */
BOOL ZJGName::SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString&>& HeaderCStr,CArray<int,int>& iHeader)//设置列表控件的标头字符串内容
{

	int nColCound = pListCtrl->GetHeaderCtrl()->GetItemCount();
	while(nColCound-- > 0)
	{
		if(! pListCtrl->DeleteColumn(nColCound)) return false;	//删除已有列
	}
	nColCound = HeaderCStr.GetSize();
	if(nColCound != iHeader.GetSize()) return false;//输入表头文字与宽度不一致
	for(int i = 0; i < nColCound ; i++)
	{
		pListCtrl->InsertColumn(i,HeaderCStr[i],2,iHeader[i]);
	}
	return true;
}
/*************************************************************************
 *
 * \函数名称：
 *   SetListCtrlRowPos(CListCtrl * pListCtrl,int OldTopRow)
 *
 * \输入参数:
 *		CListCtrl * pListCtrl	列表框控件指针
 *		int			OldTopRow   指定可见行     
 * \返回值:
 *		无	
 *
 * \说明:
 *   设置指定行，在列表控件中可见。
 *
 ************************************************************************
 */
void ZJGName::SetListCtrlRowPos(CListCtrl * pListCtrl,int OldTopRow)//设置指定行，在列表控件中可见。
{
//	this->m_LIST_Order_Ctrl.EnsureVisible(iCurSel,false);
	int iRowCount = pListCtrl->GetItemCount();		//总的行数
	iRowCount = iRowCount-1;						//最后一行的顺序号
	pListCtrl->EnsureVisible(iRowCount,false);		//最后一行可见
	pListCtrl->EnsureVisible(OldTopRow,false);		//原顶行可见	
}
/*************************************************************************
 *
 * \函数名称：
 *   GetFilePath()
 *
 * \输入参数:
 *		
 *		         
 * \返回值:
 *		返回对话框选择的文件路径	
 *
 * \说明:
 *   获取对话框选择的文件路径
 *
 ************************************************************************
 */
CString ZJGName::GetFilePath()
{
	CString strPath = "";
	BROWSEINFO bInfo;
	ZeroMemory(&bInfo, sizeof(bInfo));
	bInfo.hwndOwner = NULL;
	bInfo.lpszTitle = _T("请选择路径: ");
	bInfo.ulFlags = BIF_RETURNONLYFSDIRS;   
 
	LPITEMIDLIST lpDlist; //用来保存返回信息的IDList
	lpDlist = SHBrowseForFolder(&bInfo) ; //显示选择对话框
	if(lpDlist != NULL)  //用户按了确定按钮
	{
		TCHAR chPath[255]; //用来存储路径的字符串
		SHGetPathFromIDList(lpDlist, chPath);//把项目标识列表转化成字符串
		strPath = chPath; //将TCHAR类型的字符串转换为CString类型的字符串
	}
	return strPath;
}
/*************************************************************************
 *
 * \函数名称：
 *   zGetPrinter(CString CStrPrinter)
 *
 * \输入参数:
 *		CString CStrPrinter	  打印机简称，如"TSC TTP-244 Plus"
 *		         
 * \返回值:
 *		返回打印机详细名称	
 *
 * \说明:
 *   获取指定打印机详细名称
 *
 ************************************************************************
 */
CString ZJGName::zGetPrinter(CString CStrPrinter)//获取指定打印机名称
{
	//得到所有打印机   
	DWORD dwSize,dwPrinters;
	::EnumPrinters(PRINTER_ENUM_CONNECTIONS | PRINTER_ENUM_LOCAL,NULL,5,NULL,0,&dwSize,&dwPrinters);   
	BYTE *pBuffer=new BYTE[dwSize];   
	::EnumPrinters(PRINTER_ENUM_CONNECTIONS | PRINTER_ENUM_LOCAL,NULL,5,pBuffer,dwSize,&dwSize,&dwPrinters);   
	  
	CString sPrinter;  
	if(dwPrinters!=0)   
	{   
	    PRINTER_INFO_5 *pPrnInfo=(PRINTER_INFO_5 *)pBuffer;   
	    for(unsigned long i=0;i < dwPrinters;i++)   
		{   
			sPrinter.Format(_T("%s"),pPrnInfo-> pPrinterName);	//取打印机名称
			if(sPrinter.Find(CStrPrinter) >=0 )					//查找是否是条形码打印机
			{
				CString TemCStr;
				return sPrinter;
			}
			pPrnInfo++;//指针后移    
		}   
	}   
	delete []pBuffer; 
	sPrinter.Empty();
	return sPrinter;											//返回打印机名称
}
/*************************************************************************
 *
 * \函数名称：
 *   StrToTime(CString Date,char Separator)
 *
 * \输入参数:
 *		CString Date	字符串型时间。格式是：年-月-日---2020-11-28
 *		char Separator  分隔符         
 * \返回值:
 *		时间类对象	
 *
 * \说明:
 *   将字串型时间转换成时间类对象
 *
 ************************************************************************
 */
CTime ZJGName::StrToTime(CString nDateStr,char Separator)	//
{
	int iY,iM,iD,i;
	iY = iM = iD = 0;
	CString TemCStr;
    i = nDateStr.Find(Separator);
	if(i > 0)
	{
		TemCStr = nDateStr.Left(i);	//取左边的数值
		iY = atoi(TemCStr);
		nDateStr = nDateStr.Mid(i+1);
		i = nDateStr.Find(Separator);
		if(i > 0)
		{
			TemCStr = nDateStr.Left(i);	//取左边的数值
			iM = atoi(TemCStr);
			nDateStr = nDateStr.Mid(i+1);
			iD = atoi(nDateStr);
		}
		else return CTime(1997,1,1,0,0,0,0);
	}
	else return CTime(1997,1,1,0,0,0,0);

	if(iY < 1971) iY = 1971;
	if(iD < 1 || iD > 31) iD = 1;
	if(iM < 1 && iM > 12) iM = 1;

	return CTime(iY,iM,iD,0,0,0,0);
}
/*************************************************************************
 *
 * \函数名称：
 *   StrToTime(CString Date)
 *
 * \输入参数:
 *		CString Date	字符串型时间。格式是：年月日---20201128				8位
 *											  年月日---20201128-12-59-59	17位         
 * \返回值:
 *		时间类对象	
 *
 * \说明:
 *   将字串型时间转换成时间类对象
 *
 ************************************************************************
 */
CTime ZJGName::StrToTime(CString Date)
{
	int iYear,iMonth,iDay;
	int iHour,iMinute,iSecond;//小时，分钟，秒
	iYear = 1900;
	iMonth = 12;
	iDay = 1;
	iHour = iMinute = iSecond = 0;

	if(Date.GetLength() == 8)
	{
		iYear = atoi(Date.Left(4));
		iMonth= atoi(Date.Mid(4,2));
		iDay  = atoi(Date.Mid(6,2));
		iHour = iMinute = iSecond = 0;
	}
	else if(Date.GetLength() == 17)
	{
		iYear = atoi(Date.Left(4));
		iMonth= atoi(Date.Mid(4,2));
		iDay  = atoi(Date.Mid(6,2));
		iHour = atoi(Date.Mid(9,2));
		iMinute = atoi(Date.Mid(12,2));
		iSecond = atoi(Date.Mid(15,2));;
	}
	if(iYear < 1971) iYear = 1971;
	if(iDay < 1 || iDay > 31) iDay = 1;
	if(iMonth < 1 || iMonth > 12) iMonth = 1;
	
	return CTime(iYear,iMonth,iDay,iHour,iMinute,iSecond,0);
}
/*************************************************************************
 *
 * \函数名称：
 *   zDeleteFile(CString FileName)
 *
 * \输入参数:
 *		CString FileName	删除文件名
 *
 * \返回值:
 *		删除成功，返回真。否则返回假
 *
 * \说明:
 *   删除指定的文件
 *
 ************************************************************************
 */
BOOL ZJGName::zDeleteFile(CString FileName)		//删除文件
{
	TRY
	{
		if(CheckFile(FileName) && (DeleteFile(FileName) == false) )//文件存在且不能删除时，直接返回为假
			return false;
	}
	CATCH(CFileException,e)
	{
		e->Delete();
		e->ReportError();
		return false;
	}
	END_CATCH
	return true;
}



////CVolume类开始
CVolume::CVolume()
{
}

CVolume::CVolume(int nLength,int nWidth,int nHeight)
{
	Length = nLength;		//纸箱长，单位：毫米
	Width = nWidth;			//纸箱宽，单位：毫米
	Height = nHeight;		//纸箱高，单位：毫米
}

void CVolume::Empty()//清空
{
	Length = 0;		//纸箱长，单位：毫米
	Width = 0;		//纸箱宽，单位：毫米
	Height = 0;		//纸箱高，单位：毫米
}

BOOL CVolume::IsEmpty()	//判断空
{
	if(	Length == 0 && Width  == 0 && Height == 0 )
		return true;
	else return false;
}

BOOL CVolume::operator==(CVolume& Other)		//比较相同 
{
	if(	Height == Other.Height && 
		Length == Other.Length && 
		Width  == Other.Width	) 
		return true;
	else return false;
}

void CVolume::SetVolumeCStr(CString nVolumeCStr)	//将字符串型的纸箱类转换成纸箱类，长-宽-高，中间用“-”分开
{
	int iL;
	iL = nVolumeCStr.Find('-');
	this->Length = atoi(nVolumeCStr.Left(iL));
	nVolumeCStr = nVolumeCStr.Mid(iL+1);

	iL = nVolumeCStr.Find('-');
	this->Width = atoi(nVolumeCStr.Left(iL));
	nVolumeCStr = nVolumeCStr.Mid(iL+1);

	this->Height = atoi(nVolumeCStr);
}

CString CVolume::GetVolumeCStr()		//获取字符串型的纸箱类
{
	CString Resulf;
	Resulf.Format("%d-%d-%d",this->Length,this->Width,this->Height);
	return Resulf; 
}

CString ZJGName::LongToCStr(long n)	//将数值型数字转换成字符串型数字
{
	CString TemCStr;
	TemCStr.Format("%ld",n);
	return TemCStr;
}

CString ZJGName::IntToCStr(int n)//将数值型数字转换成字符串型数字
{
	CString TemCStr;
	TemCStr.Format("%d",n);
	return TemCStr;
}

CString ZJGName::DoubleToCStr(double n)//将数值型数字转换成字符串型数字
{
	CString TemCStr;
	TemCStr.Format("%lf",n);
	return TemCStr;
}

BOOL ZJGName::SendToClipboard(CString strText)
{
	// 打开剪贴板    
	if (!OpenClipboard(NULL)|| !EmptyClipboard())     
	{    
		return false;    
	}    
  
	HGLOBAL hMen;    
	// 分配全局内存
	
	hMen = GlobalAlloc(GMEM_MOVEABLE, ((strText.GetLength()+1)*sizeof(CHAR)));     
	if (!hMen)    
	{    
		CloseClipboard(); // 关闭剪切板     
   		return false;    
	}    
	// 把数据拷贝考全局内存中    
	// 锁住内存区     
	LPSTR lpStr = (LPSTR)GlobalLock(hMen);   
	// 内存复制    
	memcpy(lpStr, strText, ((strText.GetLength()+1)*sizeof(CHAR)));     
	// 字符结束符     
	lpStr[strlen(strText)] = (TCHAR)0;    
	// 释放锁     
	GlobalUnlock(hMen);    
  
	// 把内存中的数据放到剪切板上    
	SetClipboardData(CF_TEXT, hMen);    
	CloseClipboard();

	return true;
}

BOOL ZJGName::GetFromClipboard(CString & strText)
{
	if(OpenClipboard(NULL))//首先打开一个剪切板  
	{  
		if(IsClipboardFormatAvailable(CF_TEXT))//检查此时剪切板中数据是否为TEXT，可以设置成别的！  
		{  
			HANDLE hClip; //声明一个句柄  
			hClip=GetClipboardData(CF_TEXT);//得到剪切板的句柄  
			LPSTR pBuf;
			pBuf=(LPSTR)GlobalLock(hClip);//得到指向这块内存的指针  
			if(pBuf!=NULL) strText = pBuf;
			GlobalUnlock(hClip);//解除内存锁定  
			CloseClipboard();//关闭此剪切板 
			if(pBuf!=NULL)	return true;
			else return false;
		}
		else return false;
	}
	else return false;
}

void ZJGName::CStrToChar(char * TargetChar,CString OriginalCStr,BOOL bFlag/* = true*/)//将字条串对象转换成字符数组
{
	if(bFlag)
	{
		int iPosition = OriginalCStr.Find(":");//查找":"的位置
		if( iPosition >= 0)  OriginalCStr = OriginalCStr.Mid(iPosition+1);//如果有
	}
	memset(TargetChar,0,OriginalCStr.GetLength()+1);
	memcpy(TargetChar,OriginalCStr.GetBuffer(0),OriginalCStr.GetLength());
	TargetChar[OriginalCStr.GetLength()] = '\0';
	return;
}


void ZJGName::ShowAccJPG(CString PictureFile,CDC * pDC,CRect ACCPictureRect)
{	
//	GetDlgItem(IDC_STATIC_Picture)->GetWindowRect(&ACCPictureRect);//得到控件在屏幕上的位置大小
//    ScreenToClient(&ACCPictureRect);//得到控件在对话框中的位置大小
	//得到图片的路径名称

	COLORREF mColor = pDC->GetPixel(ACCPictureRect.left - 2, ACCPictureRect.top -2);//得到图片外的点的颜色
	CBrush mBrush, * pOldBrush;
	mBrush.CreateSolidBrush(mColor);
	pOldBrush = pDC->SelectObject(&mBrush);
	pDC->Rectangle(ACCPictureRect);				//将图片区恢复到原始色

	ShowJpg(pDC,PictureFile,ACCPictureRect + CRect(-6,-6,-6,-6));
	
	pDC->SelectObject(pOldBrush);
	pOldBrush->DeleteObject();
	mBrush.DeleteObject();
}
/****************
功能：将传入的一组数字凑合与成Number个的数字，其和为Total
传入：原始数字数组,原始数字数组的个数,凑合后的数组，凑合后的个数，总和
返回：真：存在
      假：不存在
如："1 2 3 3 8 4 6 4 7 3 6 3 3 2 3 4",和251，凑合后的个数9 -->"1,23,38,46,47,36,33,23,4"
*/
BOOL ZJGName::MakeDoNumber(int* pOriginal,int O_Number,int* pNew,int Number,long Total)
{
	//算法：采用递归算法
	//1。当个数为2时，用穷尽法，试凑
	//
	//2。当个数大于2时，
	if(Total < 0) return false;
	
	int i = 0,j;
	int iDigits = NumberOfDigits(Total);									//和的位数
	long iTotal = 0;
	//当原始数字数组的个数小于结果数组的个数时，补足
	if(O_Number < Number)
	{
		int * p = NULL;
		p = new int[Number];
		while( i < O_Number)
		{
			p[i] = pOriginal[i];
			i++;
		}
		while( i < Number )  p[i++] = 0;
		BOOL b = MakeDoNumber(p,O_Number,pNew,Number,Total);
		if(p) delete p;
		return b;
	}
	if(Number < 2)
	{
		pNew[0] = CombinationNumber(pOriginal,O_Number);
	}
	else if(Number == 2)
	{
		//应当要检查是否会越界，但是本段程序没有检查
		while(i <= iDigits)
		{
			pNew[0] = CombinationNumber(pOriginal,i);							//前一个数字
			pNew[1] = CombinationNumber(pOriginal + i,O_Number - i);			//后一个数字

			if(pNew[0] + pNew[1] == Total)
			{
				return true;													//相等时返回为真
			}
			i++;
		}
	}
	else
	{
		while(i <= iDigits)
		{
			pNew[0] = CombinationNumber(pOriginal,i);							//前一个数字
			MakeDoNumber(pOriginal + i,O_Number - i ,pNew + 1,Number -1,Total - pNew[0] );
			iTotal = pNew[0];
			j = 1;
			while( j < Number)
			{
				iTotal += pNew[j];
				j++;
			}
			if(iTotal == Total) return true;									//找到相等的一组数字时，返回为真
			i++;
		}
	}
	
	return false;
}

int ZJGName::CombinationNumber(int* p,int Number)
{
	if (Number < 1) return 0;
	int i = 1;
	int iTem;
	int Result = p[0];
	while(i < Number)
	{
		iTem = NumberOfDigits(p[i]);				//数字的位数
		Result = Result * int(pow(10,iTem)) + p[i];;
		i++;
	}
	return Result;
}

int ZJGName::NumberOfDigits(long Number)
{
	int i;
	for(i = 0; Number; Number /= 10,i++);
	return i;
}

int ZJGName::CountCharQTY(CString InputCStr,char CountChar)				//统计指定字符的个数		
{
	int iQTY = 0;
	int iLenght  = InputCStr.GetLength();
	char * pCStr = InputCStr.GetBuffer(iLenght);
	for(int i = 0; i < iLenght; i++)
	{
		if(pCStr[i] == CountChar)
			iQTY++;
	}

	return iQTY;
}

int ZJGName::CountNumbrtChar(CString InputCStr)			//统计数字字符的个数		
{
	int iQTY = 0;
	int iLenght  = InputCStr.GetLength();
	char * pCStr = InputCStr.GetBuffer(iLenght);
	for(int i = 0; i < iLenght; i++)
	{
		if(pCStr[i] >=  '0' && pCStr[i] <=  '9')
			iQTY++;
	}

	return iQTY;
}
/*************************************************************************
 * \函数名称：
 *	SizePos(CString nSize,CString nStandards)
 *
 * \输入参数:
 *	待定尺码， 尺码格式有：数字尺码：27,34,36, ... ,108,136.....
 *                         字符尺码：XXS,XS,S,M,L,XL,XXL
 *                         数字+字符：4Y,6Y, ... ,16Y
 *  标准尺码串,格式是：尺码1,尺码2,尺码3, ... ,尺码N。
 * \返回值:
 *  整数
 * \说明:
 *  查找待定尺码在标准尺码串中的位置。从0开始。
 *  当返回值小于0时，表示没有此尺码
 ************************************************************************
 */
int ZJGName::SizePos(CString nSize,CString nStandards)//求尺码在标准尺码中的位置
{
	int iPos;
	iPos = nStandards.Find(nSize);
	if(iPos < 0) return iPos;			//不存在,直接返回
	
	iPos = nStandards.Find(',' + nSize + ',');
	if(iPos > 0) return iPos+1;			//找到，直接返回

	iPos = nStandards.Find(nSize + ',');
	if(iPos == 0) return 0;				//是第一个尺码，返回

	iPos = nStandards.Find(',' + nSize);
	if(iPos + nSize.GetLength() + 1 == nStandards.GetLength()) 
		return iPos+1;					//是最后一个尺码，返回
    
	return -1;							//不存在。
}


// 1. 数字字符占空间：1.33
// 2. 空格字符占空间：1
// 3. "I"和":"字符占空间：0.8
// 4. 其它字符占空间：1.58
//输入：int iWord_H,int iPaper_W
//
double ZJGName::CentreLevel(CString& WordCStr,double iWord_H,double iPaper_W)//起始点坐标；iWord_H：字高；iPaper_W：纸宽。单位：mm。
{
	int iNumber,iSmall,iLarge,iSpace;
	double fDis ;
	iNumber= CountNumbrtChar(WordCStr);		//数字字符个数
	iSmall = CountCharQTY(WordCStr,'I');	//"I"字符个数
	iSmall += CountCharQTY(WordCStr,':');	//":"字符个数
	iSpace = CountCharQTY(WordCStr,' ');	//空格字符个数
	iLarge = WordCStr.GetLength() - iNumber - iSmall - iSpace;//其他字符个数

	fDis = iSmall * 0.8 + iSpace * 1 + iNumber * 1.33 + iLarge * 1.58; //字符宽度（无单位）
	fDis = fDis * iWord_H * 0.85 / 1.633;		//实际字符宽度，有单位
	return ((iPaper_W - fDis ) / 2);			//(纸宽-字符宽)/2。

}
