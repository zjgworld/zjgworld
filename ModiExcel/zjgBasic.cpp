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

CString ZJGName::DelDecimal(CString DecimalCStr)//ɾ���ַ�����ֵ��С������
{
	if(DecimalCStr.Find(".",0)>0)
		return DecimalCStr.Left(DecimalCStr.Find(".",0));
	else 
		return DecimalCStr;
}

BOOL ZJGName::CStrIsNumber(CString InputCStr)//�ж��Ƿ�������
{
	BOOL bResult = CStrIsNumber(InputCStr.GetBuffer(InputCStr.GetLength()+1));
	InputCStr.ReleaseBuffer();
	return bResult;
}

BOOL ZJGName::CStrIsNumber(char * InputChar)//�ж��Ƿ�������
{
	StrDeleteSpace(InputChar);//ɾ����β�ո�
	int i = 0;
	int iDot = 0;
	while(InputChar[i] == ' ' || InputChar[i] == '\t')
	{	//����ǰ�������пո��TAB
		i++;
	}
	if(	InputChar[i] == '+' ||
		InputChar[i] == '-')
		i++;	//ֻ����һ��'+'�Ż�'-'��
	while(InputChar[i] != '\0')
	{
		if(InputChar[i] >= '0' && InputChar[i] <= '9')
			i++;
		else if(InputChar[i] == '.' && iDot == 0)
		{//ֻ����һ��С����
			i++;
			iDot++;
		}
		else return FALSE;
		if(i > 1024) return FALSE;//�������ȳ���1024�������ַ����������ˣ�
	}
	return TRUE;
}

CString ZJGName::StrDeleteSpace(CString InputCStr,char DeleteChar)//ȥ����β�Ŀո�
{
	//�ַ����Ȳ�����256��������س�Ϊ256���ַ���
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
void ZJGName::StrDeleteSpace(char * InputChar,char DeleteChar)//ȥ����β�Ŀո�
{
	//�ַ����Ȳ�����256��������س�Ϊ256���ַ���
	int i = 0;//ɾ��ǰ��ָ��
	int j = 0;//ɾ�����ָ��
	//char DeleteChar = ' ';
	//ȥ��ͷ���Ŀո�
	while(InputChar[i] != '\0')
	{
		if(InputChar[i] != DeleteChar) break;//���ǿո�ʱ����
		if( i >= 256)  InputChar[i+1] = '\0'; 
		i++;

	}
	j = 0;//ɾ�����ָ��
	while(InputChar[i] != '\0')
	{
		InputChar[j] = InputChar[i];
		j++;
		i++;
	}
	InputChar[j] = '\0';//�ӽ�����
	//ȥ��β���Ŀո�
	while(j >=0 )
	{
		if(InputChar[j] != DeleteChar) break;//���ǿո�ʱ����
		InputChar[j] = '\0';//�ӽ�����
		j--;
	}
}

CString ZJGName::StrDeleteEndChar(CString InputCStr,char DeleteChar)//ȥ��β����ָ���ַ�
{
	//�ַ����Ȳ�����256��������س�Ϊ256���ַ���
	int i = 0 ;
	i = InputCStr.GetLength() -1;
	while((i>= 0) && (InputCStr.GetAt(i) == DeleteChar))
	{
		i--;
	}
	return InputCStr.Left(i+1);
}

void ZJGName::StrDeleteEndChar(char * InputChar,char DeleteChar)//ȥ��β����ָ���ַ�
{
	//�ַ����Ȳ�����256��������س�Ϊ256���ַ���
	int i = 0;//ɾ��ǰ��ָ��
	int j = 0;//ɾ�����ָ��
	//char DeleteChar = ' ';
	//�����ַ����ĳ���
	while(InputChar[i] != '\0')
	{
		if( i >= 256)  InputChar[i+1] = '\0'; 
		i++;
	}
	//ȥ��β���Ŀո�
	while(i >=0 )
	{
		if(InputChar[i] != DeleteChar) break;//���ǿո�ʱ����
		InputChar[i] = '\0';//�ӽ�����
		i--;
	}
}
CString ZJGName::NumberToLetter(double Number)//����ֵ������ת����Ӣ���ַ�����ֵ
{
	long lNumber = (long)Number;
	long fNumber = (long)((Number - lNumber + 0.0049)*100);//��������
	if(fNumber > 0)
		return	NumberToLetter(lNumber) + " AND CENTS " + NumberToLetter(fNumber);
	else return	NumberToLetter(lNumber);
}

CString ZJGName::NumberToLetter(long Number)//����ֵ������ת����Ӣ���ַ�����ֵ
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
	if( ((long (Number / 10)) * 10 == Number) && (Number >= 20))//����20����ʮλ�� 
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
	int CurrentIndex = MaxChar -1;//��ǰλ��
	int TotalCharNumber = 0;

	if(NameAddCStr.GetAt(CurrentIndex) <0)//�Ǻ���
	{		
		while(CurrentIndex >=0)
		{
			if(NameAddCStr.GetAt(CurrentIndex) > 0)//�Ǻ���ʱ
				break;
			CurrentIndex--;
			TotalCharNumber++;
		}
		if((TotalCharNumber % 2) == 1)//����
			return MaxChar-1;
		else //˫��
			return MaxChar;
	}
	if((NameAddCStr.GetAt(CurrentIndex) >= '0') && ( NameAddCStr.GetAt(CurrentIndex) <= '9'))//������
	{
		while((CurrentIndex >=0)&&(NameAddCStr.GetAt(CurrentIndex) >= '0') && ( NameAddCStr.GetAt(CurrentIndex) <= '9'))//������
			CurrentIndex--;
		return CurrentIndex + 1 ;
	}
	while(CurrentIndex >= 0)
	{	
		if(	NameAddCStr.GetAt(CurrentIndex) == ' ' ||//���ո�ʱ
			NameAddCStr.GetAt(CurrentIndex) == ',' ) //��','ʱ
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
	int CurrentIndex = MaxChar -1;//��ǰλ��
	int TotalCharNumber = 0;

	if(NameAddCStr[CurrentIndex] <0)//�Ǻ���
	{		
		while(CurrentIndex >=0)
		{
			if(NameAddCStr[CurrentIndex] > 0)//�Ǻ���ʱ
				break;
			CurrentIndex--;
			TotalCharNumber++;
		}
		if((TotalCharNumber % 2) == 1)//����
			return MaxChar-1;
		else //˫��
			return MaxChar;
	}
	if((NameAddCStr[CurrentIndex] >= '0') && ( NameAddCStr[CurrentIndex] <= '9'))//������
	{
		while((CurrentIndex >=0)&&(NameAddCStr[CurrentIndex] >= '0') && ( NameAddCStr[CurrentIndex] <= '9'))//������
			CurrentIndex--;
		return CurrentIndex + 1 ;
	}
	while(CurrentIndex >= 0)
	{	
		if(	NameAddCStr[CurrentIndex] == ' ' ||//���ո�ʱ
			NameAddCStr[CurrentIndex] == ',' ) //��','ʱ
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

CString ZJGName::MonthIToCStr(int iMonth,BOOL Simplify)//�������·�ת����Ӣ�ĵ���
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
Description : ��DC�ϰ�ͼƬԭʼ�ߴ���ʾJPGͼƬ
Calls       : 
Called By   : 
Parameter   : [CDC* pDC] --- DC
            : [CString strPath] --- Ҫ��ʾ��ͼƬ·��������ȫ·��
            : [int x] --- DC����ʾ��Xλ�� 
            : [int y] --- DC����ʾ��Yλ��
            : [bool OriginalSize] --- �Ƿ�ͼƬԭʼ�ߴ���ʾ,falseʱ����DC��С����
Return      : bool --- �Ƿ�ɹ�
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

		HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, cb); //�Ӷ��з���һ����Ŀ���ֽ�   
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
		{		//ͼƬС�ڱ߿򣬰�ͼƬ����ʾ�����߿������λ��������ͼƬ��ʾλ�õ�
			Display_x = mRect.CenterPoint().x - (DWORD)(fX / 2); //�߿������λ��
			Display_y = mRect.CenterPoint().y - (DWORD)(fY / 2); //�߿������λ��
			pPic->Render(*pDC,Display_x,Display_y,(DWORD)fX,(DWORD)fY,0,jpgHeight,jpgWidth,-jpgHeight,NULL); 
		}
		else	//ͼƬ���ڱ߿򣬰��߿�����ʾ����ͼƬ�ĸ߿��������ͼƬ��ʾλ�õ㼰���
		{
			if( (fX / mWidth) > (fY / mHeight) )
			{		//ͼƬ��ȱȸ߶ȴ�ʱ�������֮�ȼ���
				fScale = fX / mWidth;

			}
			else	//ͼƬ�߶ȱȳ��ȴ�ʱ�����߶�֮�ȼ���
			{
				fScale = fY / mHeight;
			}
			mWidth	= (DWORD)(fX / fScale);
			mHeight	= (DWORD)(fY / fScale);
			Display_x = mRect.CenterPoint().x - (DWORD)(mWidth / 2); //�߿������λ��
			Display_y = mRect.CenterPoint().y - (DWORD)(mHeight / 2); //�߿������λ��
			pPic->Render(*pDC,Display_x,Display_y,mWidth,mHeight,0,jpgHeight,jpgWidth,-jpgHeight,NULL);
		}

		{		//ͼƬ���ڱ߿򣬰��߿�����ʾ
//			pPic->Render(*pDC,mRect.left,mRect.top,mRect.Width(),mRect.Height(),0,jpgHeight,jpgWidth,-jpgHeight,NULL); 
			 
		}     
	}      
	return true;   
}
/****************************
���ܣ�ɾ���ļ���ÿ��ָ���ַ�ǰ������
****************************/
BOOL ZJGName::DeleteFileCStr(CString mDelCStr)
{
	CFileDialog myFileDialog(TRUE,NULL,NULL,OFN_FILEMUSTEXIST,"C++File(*.cpp)|*.cpp|Head File(*.h)|*.h||");//TRUE:�򿪣�FALSE����
	//CFileDialog myFileDialog(TRUE);
	CString rFileName;
	CString wFileName;


	myFileDialog.DoModal();
	rFileName = myFileDialog.GetPathName();
	wFileName = myFileDialog.GetFileExt();
	wFileName = rFileName.Left(rFileName.Find(".")) +"-1." + wFileName;
	
	CStdioFile fileRead;///ֻ����ʽ��
	CStdioFile fileWrite;///ֻ����ʽ��
	CString ReadLineCStr;
    CString WriteLineCStr;
	///CFile::modeRead�ɸ�Ϊ CFile::modeWrite(ֻд),CFile::modeReadWrite(��д),CFile::modeCreate(�½�)
	if(!fileRead.Open(rFileName,CFile::modeRead)) return false;
	if(!fileWrite.Open(wFileName,CFile::modeCreate|CFile::modeWrite)) return false;

	int mDelCStrLength = mDelCStr.GetLength();//�ַ�����
	while(fileRead.ReadString(ReadLineCStr))
	{
		WriteLineCStr = ReadLineCStr.Mid(ReadLineCStr.Find(".") + mDelCStrLength) + "\r\n";
		fileWrite.WriteString(WriteLineCStr);
	}
	fileRead.Close();
	fileWrite.Close();
	return true;
}

BOOL ZJGName::IsDebug()		//�жϵ�ǰ��ִ���ļ��Ƿ��ǵ����ļ�Ŀ¼ 
{
	TCHAR SourceDir[256];
	::memset(SourceDir,256,'\0');
	GetModuleFileName(NULL,SourceDir,255);
	CString SourceFile(SourceDir);
	if(SourceFile.Find("\\Debug") >= 0)
		return true;
	else return false;
}

CString ZJGName::GetExeFileDir()//��ȡ����ǰִ���ļ�Ŀ¼ 
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
���ܣ�����ļ����Ƿ����
���룺�ļ�������
���أ��棺����
      �٣�������
*/
BOOL ZJGName::CheckDir(CString DirName) 
{
	DWORD dwAttr=GetFileAttributes(DirName);
	if(dwAttr==0xFFFFFFFF)  //��ʱ�ļ��в�����   
	{  
		return false;
	}
	return true;
}
/****************
���ܣ������ļ���
���룺�ļ�������
���أ��棺�����ɹ�
      �٣�����ʧ��
*/
BOOL ZJGName::CreateDir(CString DirName) 
{
	TRY
	{
		int iStart;
		//������һ���ַ��Ƿ���"\"������ǣ�ɾ�����һ���ַ�
		if(DirName.Right(1) == "\\")
			DirName = DirName.Left(DirName.GetLength() -1);
		iStart = 0;
		CString SubDirName;
		//����ѭ���ǴӸ�Ŀ¼��ʼ��飬�Ƿ���ڣ������ھʹ�����
		while(iStart >= 0)
		{
			iStart = DirName.Find("\\",iStart+1);
			if(iStart >= 0)
				SubDirName = DirName.Left(iStart);
			else
				SubDirName = DirName;
			if(!CheckFile(SubDirName))	//������ʱ������
			{
				if(SubDirName.Find("\\") < 0) 
					continue;//��Ŀ¼���ܴ���

				if(!CreateDirectory(SubDirName,NULL)) //������ʱ������
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
���ܣ�����ļ��Ƿ����
���룺�ļ�����
���أ��棺����
      �٣�������
*/
BOOL ZJGName::CheckFile(CString FileName)		//����ļ��Ƿ����
{
	CFileFind m_FileFind;
	return m_FileFind.FindFile(FileName);	
}

//�õ��б�ؼ��ı�ͷ�ַ�������
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
	pHeaderArray->SetSize(nColCound+1);//���������С
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

void ZJGName::GetListCtrlHeader(CListCtrl * pListCtrl,CList<CString,CString> * pHeaderList,int iStart,int iEnd)//�õ��б�ؼ��ı�ͷ�ַ�������
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

void ZJGName::SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray,int iSizeCount)//�����б�ؼ��ı�ͷ�ַ�������
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
 * \�������ƣ�
 *   SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CList<CString,CString&> * pHeaderList)
 *
 * \�������:
 *		CListCtrl				* pListCtrl		�б��ؼ�ָ��
 *		CList<CString,CString&> * pHeaderList	��ͷ�ַ�����
 * \����ֵ:
 *		���óɹ��������棬����Ϊ��	
 *
 * \˵��:
 *   �����б�ؼ��ı�ͷ�ַ�������,���ν������ַ�����д����������
 *
 ************************************************************************
 */
/*
void ZJGName::SetListCtrlHeader(CListCtrl * pListCtrl,CList<CString,CString&> * pHeaderList)//�����б�ؼ��ı�ͷ�ַ�������
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
 * \�������ƣ�
 *   SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString&>& HeaderCStr,CArray<int,int>& iHeader)
 *
 * \�������:
 *		CArray<CString,CString&>&	HeaderCStr  ��ͷ�ַ�������
 *		CArray<int,int>&			iHeader     ��ͷ�ַ�����Ӧ�кż���
 * \����ֵ:
 *		���óɹ��������棬����Ϊ��	
 *
 * \˵��:
 *   �����б�ؼ��ı�ͷ�ַ�������
 *
 ************************************************************************
 */
BOOL ZJGName::SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString&>& HeaderCStr,CArray<int,int>& iHeader)//�����б�ؼ��ı�ͷ�ַ�������
{

	int nColCound = pListCtrl->GetHeaderCtrl()->GetItemCount();
	while(nColCound-- > 0)
	{
		if(! pListCtrl->DeleteColumn(nColCound)) return false;	//ɾ��������
	}
	nColCound = HeaderCStr.GetSize();
	if(nColCound != iHeader.GetSize()) return false;//�����ͷ�������Ȳ�һ��
	for(int i = 0; i < nColCound ; i++)
	{
		pListCtrl->InsertColumn(i,HeaderCStr[i],2,iHeader[i]);
	}
	return true;
}
/*************************************************************************
 *
 * \�������ƣ�
 *   SetListCtrlRowPos(CListCtrl * pListCtrl,int OldTopRow)
 *
 * \�������:
 *		CListCtrl * pListCtrl	�б��ؼ�ָ��
 *		int			OldTopRow   ָ���ɼ���     
 * \����ֵ:
 *		��	
 *
 * \˵��:
 *   ����ָ���У����б�ؼ��пɼ���
 *
 ************************************************************************
 */
void ZJGName::SetListCtrlRowPos(CListCtrl * pListCtrl,int OldTopRow)//����ָ���У����б�ؼ��пɼ���
{
//	this->m_LIST_Order_Ctrl.EnsureVisible(iCurSel,false);
	int iRowCount = pListCtrl->GetItemCount();		//�ܵ�����
	iRowCount = iRowCount-1;						//���һ�е�˳���
	pListCtrl->EnsureVisible(iRowCount,false);		//���һ�пɼ�
	pListCtrl->EnsureVisible(OldTopRow,false);		//ԭ���пɼ�	
}
/*************************************************************************
 *
 * \�������ƣ�
 *   GetFilePath()
 *
 * \�������:
 *		
 *		         
 * \����ֵ:
 *		���ضԻ���ѡ����ļ�·��	
 *
 * \˵��:
 *   ��ȡ�Ի���ѡ����ļ�·��
 *
 ************************************************************************
 */
CString ZJGName::GetFilePath()
{
	CString strPath = "";
	BROWSEINFO bInfo;
	ZeroMemory(&bInfo, sizeof(bInfo));
	bInfo.hwndOwner = NULL;
	bInfo.lpszTitle = _T("��ѡ��·��: ");
	bInfo.ulFlags = BIF_RETURNONLYFSDIRS;   
 
	LPITEMIDLIST lpDlist; //�������淵����Ϣ��IDList
	lpDlist = SHBrowseForFolder(&bInfo) ; //��ʾѡ��Ի���
	if(lpDlist != NULL)  //�û�����ȷ����ť
	{
		TCHAR chPath[255]; //�����洢·�����ַ���
		SHGetPathFromIDList(lpDlist, chPath);//����Ŀ��ʶ�б�ת�����ַ���
		strPath = chPath; //��TCHAR���͵��ַ���ת��ΪCString���͵��ַ���
	}
	return strPath;
}
/*************************************************************************
 *
 * \�������ƣ�
 *   zGetPrinter(CString CStrPrinter)
 *
 * \�������:
 *		CString CStrPrinter	  ��ӡ����ƣ���"TSC TTP-244 Plus"
 *		         
 * \����ֵ:
 *		���ش�ӡ����ϸ����	
 *
 * \˵��:
 *   ��ȡָ����ӡ����ϸ����
 *
 ************************************************************************
 */
CString ZJGName::zGetPrinter(CString CStrPrinter)//��ȡָ����ӡ������
{
	//�õ����д�ӡ��   
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
			sPrinter.Format(_T("%s"),pPrnInfo-> pPrinterName);	//ȡ��ӡ������
			if(sPrinter.Find(CStrPrinter) >=0 )					//�����Ƿ����������ӡ��
			{
				CString TemCStr;
				return sPrinter;
			}
			pPrnInfo++;//ָ�����    
		}   
	}   
	delete []pBuffer; 
	sPrinter.Empty();
	return sPrinter;											//���ش�ӡ������
}
/*************************************************************************
 *
 * \�������ƣ�
 *   StrToTime(CString Date,char Separator)
 *
 * \�������:
 *		CString Date	�ַ�����ʱ�䡣��ʽ�ǣ���-��-��---2020-11-28
 *		char Separator  �ָ���         
 * \����ֵ:
 *		ʱ�������	
 *
 * \˵��:
 *   ���ִ���ʱ��ת����ʱ�������
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
		TemCStr = nDateStr.Left(i);	//ȡ��ߵ���ֵ
		iY = atoi(TemCStr);
		nDateStr = nDateStr.Mid(i+1);
		i = nDateStr.Find(Separator);
		if(i > 0)
		{
			TemCStr = nDateStr.Left(i);	//ȡ��ߵ���ֵ
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
 * \�������ƣ�
 *   StrToTime(CString Date)
 *
 * \�������:
 *		CString Date	�ַ�����ʱ�䡣��ʽ�ǣ�������---20201128				8λ
 *											  ������---20201128-12-59-59	17λ         
 * \����ֵ:
 *		ʱ�������	
 *
 * \˵��:
 *   ���ִ���ʱ��ת����ʱ�������
 *
 ************************************************************************
 */
CTime ZJGName::StrToTime(CString Date)
{
	int iYear,iMonth,iDay;
	int iHour,iMinute,iSecond;//Сʱ�����ӣ���
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
 * \�������ƣ�
 *   zDeleteFile(CString FileName)
 *
 * \�������:
 *		CString FileName	ɾ���ļ���
 *
 * \����ֵ:
 *		ɾ���ɹ��������档���򷵻ؼ�
 *
 * \˵��:
 *   ɾ��ָ�����ļ�
 *
 ************************************************************************
 */
BOOL ZJGName::zDeleteFile(CString FileName)		//ɾ���ļ�
{
	TRY
	{
		if(CheckFile(FileName) && (DeleteFile(FileName) == false) )//�ļ������Ҳ���ɾ��ʱ��ֱ�ӷ���Ϊ��
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



////CVolume�࿪ʼ
CVolume::CVolume()
{
}

CVolume::CVolume(int nLength,int nWidth,int nHeight)
{
	Length = nLength;		//ֽ�䳤����λ������
	Width = nWidth;			//ֽ�����λ������
	Height = nHeight;		//ֽ��ߣ���λ������
}

void CVolume::Empty()//���
{
	Length = 0;		//ֽ�䳤����λ������
	Width = 0;		//ֽ�����λ������
	Height = 0;		//ֽ��ߣ���λ������
}

BOOL CVolume::IsEmpty()	//�жϿ�
{
	if(	Length == 0 && Width  == 0 && Height == 0 )
		return true;
	else return false;
}

BOOL CVolume::operator==(CVolume& Other)		//�Ƚ���ͬ 
{
	if(	Height == Other.Height && 
		Length == Other.Length && 
		Width  == Other.Width	) 
		return true;
	else return false;
}

void CVolume::SetVolumeCStr(CString nVolumeCStr)	//���ַ����͵�ֽ����ת����ֽ���࣬��-��-�ߣ��м��á�-���ֿ�
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

CString CVolume::GetVolumeCStr()		//��ȡ�ַ����͵�ֽ����
{
	CString Resulf;
	Resulf.Format("%d-%d-%d",this->Length,this->Width,this->Height);
	return Resulf; 
}

CString ZJGName::LongToCStr(long n)	//����ֵ������ת�����ַ���������
{
	CString TemCStr;
	TemCStr.Format("%ld",n);
	return TemCStr;
}

CString ZJGName::IntToCStr(int n)//����ֵ������ת�����ַ���������
{
	CString TemCStr;
	TemCStr.Format("%d",n);
	return TemCStr;
}

CString ZJGName::DoubleToCStr(double n)//����ֵ������ת�����ַ���������
{
	CString TemCStr;
	TemCStr.Format("%lf",n);
	return TemCStr;
}

BOOL ZJGName::SendToClipboard(CString strText)
{
	// �򿪼�����    
	if (!OpenClipboard(NULL)|| !EmptyClipboard())     
	{    
		return false;    
	}    
  
	HGLOBAL hMen;    
	// ����ȫ���ڴ�
	
	hMen = GlobalAlloc(GMEM_MOVEABLE, ((strText.GetLength()+1)*sizeof(CHAR)));     
	if (!hMen)    
	{    
		CloseClipboard(); // �رռ��а�     
   		return false;    
	}    
	// �����ݿ�����ȫ���ڴ���    
	// ��ס�ڴ���     
	LPSTR lpStr = (LPSTR)GlobalLock(hMen);   
	// �ڴ渴��    
	memcpy(lpStr, strText, ((strText.GetLength()+1)*sizeof(CHAR)));     
	// �ַ�������     
	lpStr[strlen(strText)] = (TCHAR)0;    
	// �ͷ���     
	GlobalUnlock(hMen);    
  
	// ���ڴ��е����ݷŵ����а���    
	SetClipboardData(CF_TEXT, hMen);    
	CloseClipboard();

	return true;
}

BOOL ZJGName::GetFromClipboard(CString & strText)
{
	if(OpenClipboard(NULL))//���ȴ�һ�����а�  
	{  
		if(IsClipboardFormatAvailable(CF_TEXT))//����ʱ���а��������Ƿ�ΪTEXT���������óɱ�ģ�  
		{  
			HANDLE hClip; //����һ�����  
			hClip=GetClipboardData(CF_TEXT);//�õ����а�ľ��  
			LPSTR pBuf;
			pBuf=(LPSTR)GlobalLock(hClip);//�õ�ָ������ڴ��ָ��  
			if(pBuf!=NULL) strText = pBuf;
			GlobalUnlock(hClip);//����ڴ�����  
			CloseClipboard();//�رմ˼��а� 
			if(pBuf!=NULL)	return true;
			else return false;
		}
		else return false;
	}
	else return false;
}

void ZJGName::CStrToChar(char * TargetChar,CString OriginalCStr,BOOL bFlag/* = true*/)//������������ת�����ַ�����
{
	if(bFlag)
	{
		int iPosition = OriginalCStr.Find(":");//����":"��λ��
		if( iPosition >= 0)  OriginalCStr = OriginalCStr.Mid(iPosition+1);//�����
	}
	memset(TargetChar,0,OriginalCStr.GetLength()+1);
	memcpy(TargetChar,OriginalCStr.GetBuffer(0),OriginalCStr.GetLength());
	TargetChar[OriginalCStr.GetLength()] = '\0';
	return;
}


void ZJGName::ShowAccJPG(CString PictureFile,CDC * pDC,CRect ACCPictureRect)
{	
//	GetDlgItem(IDC_STATIC_Picture)->GetWindowRect(&ACCPictureRect);//�õ��ؼ�����Ļ�ϵ�λ�ô�С
//    ScreenToClient(&ACCPictureRect);//�õ��ؼ��ڶԻ����е�λ�ô�С
	//�õ�ͼƬ��·������

	COLORREF mColor = pDC->GetPixel(ACCPictureRect.left - 2, ACCPictureRect.top -2);//�õ�ͼƬ��ĵ����ɫ
	CBrush mBrush, * pOldBrush;
	mBrush.CreateSolidBrush(mColor);
	pOldBrush = pDC->SelectObject(&mBrush);
	pDC->Rectangle(ACCPictureRect);				//��ͼƬ���ָ���ԭʼɫ

	ShowJpg(pDC,PictureFile,ACCPictureRect + CRect(-6,-6,-6,-6));
	
	pDC->SelectObject(pOldBrush);
	pOldBrush->DeleteObject();
	mBrush.DeleteObject();
}
/****************
���ܣ��������һ�����ִպ����Number�������֣����ΪTotal
���룺ԭʼ��������,ԭʼ��������ĸ���,�պϺ�����飬�պϺ�ĸ������ܺ�
���أ��棺����
      �٣�������
�磺"1 2 3 3 8 4 6 4 7 3 6 3 3 2 3 4",��251���պϺ�ĸ���9 -->"1,23,38,46,47,36,33,23,4"
*/
BOOL ZJGName::MakeDoNumber(int* pOriginal,int O_Number,int* pNew,int Number,long Total)
{
	//�㷨�����õݹ��㷨
	//1��������Ϊ2ʱ����������Դ�
	//
	//2������������2ʱ��
	if(Total < 0) return false;
	
	int i = 0,j;
	int iDigits = NumberOfDigits(Total);									//�͵�λ��
	long iTotal = 0;
	//��ԭʼ��������ĸ���С�ڽ������ĸ���ʱ������
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
		//Ӧ��Ҫ����Ƿ��Խ�磬���Ǳ��γ���û�м��
		while(i <= iDigits)
		{
			pNew[0] = CombinationNumber(pOriginal,i);							//ǰһ������
			pNew[1] = CombinationNumber(pOriginal + i,O_Number - i);			//��һ������

			if(pNew[0] + pNew[1] == Total)
			{
				return true;													//���ʱ����Ϊ��
			}
			i++;
		}
	}
	else
	{
		while(i <= iDigits)
		{
			pNew[0] = CombinationNumber(pOriginal,i);							//ǰһ������
			MakeDoNumber(pOriginal + i,O_Number - i ,pNew + 1,Number -1,Total - pNew[0] );
			iTotal = pNew[0];
			j = 1;
			while( j < Number)
			{
				iTotal += pNew[j];
				j++;
			}
			if(iTotal == Total) return true;									//�ҵ���ȵ�һ������ʱ������Ϊ��
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
		iTem = NumberOfDigits(p[i]);				//���ֵ�λ��
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

int ZJGName::CountCharQTY(CString InputCStr,char CountChar)				//ͳ��ָ���ַ��ĸ���		
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

int ZJGName::CountNumbrtChar(CString InputCStr)			//ͳ�������ַ��ĸ���		
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
 * \�������ƣ�
 *	SizePos(CString nSize,CString nStandards)
 *
 * \�������:
 *	�������룬 �����ʽ�У����ֳ��룺27,34,36, ... ,108,136.....
 *                         �ַ����룺XXS,XS,S,M,L,XL,XXL
 *                         ����+�ַ���4Y,6Y, ... ,16Y
 *  ��׼���봮,��ʽ�ǣ�����1,����2,����3, ... ,����N��
 * \����ֵ:
 *  ����
 * \˵��:
 *  ���Ҵ��������ڱ�׼���봮�е�λ�á���0��ʼ��
 *  ������ֵС��0ʱ����ʾû�д˳���
 ************************************************************************
 */
int ZJGName::SizePos(CString nSize,CString nStandards)//������ڱ�׼�����е�λ��
{
	int iPos;
	iPos = nStandards.Find(nSize);
	if(iPos < 0) return iPos;			//������,ֱ�ӷ���
	
	iPos = nStandards.Find(',' + nSize + ',');
	if(iPos > 0) return iPos+1;			//�ҵ���ֱ�ӷ���

	iPos = nStandards.Find(nSize + ',');
	if(iPos == 0) return 0;				//�ǵ�һ�����룬����

	iPos = nStandards.Find(',' + nSize);
	if(iPos + nSize.GetLength() + 1 == nStandards.GetLength()) 
		return iPos+1;					//�����һ�����룬����
    
	return -1;							//�����ڡ�
}


// 1. �����ַ�ռ�ռ䣺1.33
// 2. �ո��ַ�ռ�ռ䣺1
// 3. "I"��":"�ַ�ռ�ռ䣺0.8
// 4. �����ַ�ռ�ռ䣺1.58
//���룺int iWord_H,int iPaper_W
//
double ZJGName::CentreLevel(CString& WordCStr,double iWord_H,double iPaper_W)//��ʼ�����ꣻiWord_H���ָߣ�iPaper_W��ֽ����λ��mm��
{
	int iNumber,iSmall,iLarge,iSpace;
	double fDis ;
	iNumber= CountNumbrtChar(WordCStr);		//�����ַ�����
	iSmall = CountCharQTY(WordCStr,'I');	//"I"�ַ�����
	iSmall += CountCharQTY(WordCStr,':');	//":"�ַ�����
	iSpace = CountCharQTY(WordCStr,' ');	//�ո��ַ�����
	iLarge = WordCStr.GetLength() - iNumber - iSmall - iSpace;//�����ַ�����

	fDis = iSmall * 0.8 + iSpace * 1 + iNumber * 1.33 + iLarge * 1.58; //�ַ���ȣ��޵�λ��
	fDis = fDis * iWord_H * 0.85 / 1.633;		//ʵ���ַ���ȣ��е�λ
	return ((iPaper_W - fDis ) / 2);			//(ֽ��-�ַ���)/2��

}
