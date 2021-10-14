#if !defined(ZJGBASIC_H_INCLUDED_)
#define ZJGBASIC_H_INCLUDED_

namespace ZJGName
{
#include <afx.h>
#include "stdafx.h"
//#include "BasicClass.h"

#define BARCODEPRINTER "TSC TTP-244 Plus"
#define SIZETYPENUMBER	3		//尺码类型的数量
///////下面罗列了常用的自编程序

CString DelDecimal(CString DecimalCStr);	//删除字符型数值的小数部分
BOOL CStrIsNumber(CString InputCStr);		//判断是否是数字
BOOL CStrIsNumber(char * InputChar);		//判断是否是数字
CString StrDeleteSpace(CString InputCStr,char DeleteChar = ' ');//去除首尾的空格
void StrDeleteSpace(char * InputChar,char DeleteChar = ' ');	//去除首尾的空格
CString StrDeleteEndChar(CString InputCStr,char DeleteChar);	//去除首尾的指定字符
void StrDeleteEndChar(char * InputChar,char DeleteChar);		//去除首尾的指定字符
CString MonthIToCStr(int iMonth,BOOL Simplify);					//将数字月份转换成英文单词,Simplify为真，简略型
int CountCharQTY(CString InputCStr,char CountChar);				//统计指定字符的个数		
int CountNumbrtChar(CString InputCStr);			//统计数字字符的个数		
double CentreLevel(CString& WordCStr,double iWord_H,double iPaper_W);//起始点坐标；iWord_H：字高；iPaper_W：纸宽。单位：mm。

CString NumberToLetter(long Number);			//将数值型数字转换成英文字符型数值
CString NumberToLetter(double Number);			//将数值型数字转换成英文字符型数值,只保留两位小数

int CutNameAdd(CString NameAddCStr,int MaxChar);	//分割公司名与地址名
int CutNameAdd(TCHAR * NameAddCStr,int MaxChar);	//分割公司名与地址名
BOOL DeleteFileCStr(CString mDelCStr);	//删除文件中每行指定字符前的内容
CString GetExeFileDir();				//获取程序当前执行文件名
BOOL IsDebug();							//判断当前可执行文件是否是调试文件目录 
BOOL CheckDir(CString DirName);			//检查文件夹是否存在
BOOL CreateDir(CString DirName);		//创建文件夹
BOOL CheckFile(CString FileName);		//检查文件是否存在
BOOL zDeleteFile(CString FileName);		//删除指定文件

CString GetFilePath();								//获取对话框选择的文件路径
CTime StrToTime(CString nDateStr,char Separator);	//将字符串转换成时间
CTime StrToTime(CString Date);						//将字符串转换成时间

BOOL SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString&>& HeaderCStr,CArray<int,int>& iHeader);//设置列表控件的标头字符串内容
void SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray,int iSizeCount);		//设置列表控件的标头字符串内容
//void SetListCtrlHeader(CListCtrl * pListCtrl,);							//设置列表控件的标头字符串内容

void GetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray);						//得到列表控件的标头字符串内容
void GetListCtrlHeader(CListCtrl * pListCtrl,CList<CString,CString> * pHeaderList,int iStart = 0,int iEnd = -1);//得到列表控件的标头字符串内容
void SetListCtrlRowPos(CListCtrl * pListCtrl,int OldTopRow);//设置指定行，在列表控件中可见。

CString zGetPrinter(CString CStrPrinter);	//得到TSC打印机
CString LongToCStr(long n);					//将数值型数字转换成字符串型数字
CString IntToCStr(int n);					//将数值型数字转换成字符串型数字
CString DoubleToCStr(double n);				//将数值型数字转换成字符串型数字

BOOL SendToClipboard(CString strText);		//将文本内容复制到剪贴板
BOOL GetFromClipboard(CString & strText);	//将剪贴板内容复制到字符串

int CombinationNumber(int* p,int Number);		//将数字数组中的数字组合成一个大数字，如"2,4,7"-->"247"
BOOL MakeDoNumber(int* pOriginal,int O_Number,int* pNew,int Number,long Total);
int NumberOfDigits(long Number);				//求数字的位数

void CStrToChar(char * TargetChar,CString OriginalCStr,BOOL bFlag = true);		//将字条串对象转换成字符数组

BOOL ShowJpg(CDC* pDC, CString strPath,CRect mRect);//显示JPG图片
void ShowAccJPG(CString PictureFile,CDC * pDC,CRect ACCPictureRect);

int SizePos(CString nSize,CString nStandards);				//求尺码在标准尺码中的位置
}

class CVolume //纸箱体积类
{
public:
	CVolume();
	CVolume(int nLength,int nWidth,int nHeight);

	int GetLength(){return Length;};	//获取长
	int GetWidth(){return Width;};		//获取宽
	int GetHeight(){return Height;};	//获取高

	void SetLength(int nLength){ Length = nLength;};	//设置长
	void SetWidth(int nWidth){ Width = nWidth;};		//设置宽
	void SetHeight(int nHeight){ Height = nHeight;};	//设置高

	double GetVolume(){return Length/1000.0/1000000.0*Width*Height;};//体积，单位:立方米
	
	void SetVolumeCStr(CString nVolumeCStr);	//将字符串型的纸箱类转换成纸箱类
	CString GetVolumeCStr();		//获取字符串型的纸箱类
	void Empty();	//清空
	BOOL IsEmpty();	//判断空
//	CVolume& operator=(CVolume& Other);		//赋值
	BOOL operator==(CVolume& Other);		//比较相同 

private:
	int	Length;		//纸箱长，单位：毫米
	int	Width;		//纸箱宽，单位：毫米
	int	Height;		//纸箱高，单位：毫米
};

#endif // defined(ZJGBASIC_H_INCLUDED_)