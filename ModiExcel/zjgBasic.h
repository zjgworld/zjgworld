#if !defined(ZJGBASIC_H_INCLUDED_)
#define ZJGBASIC_H_INCLUDED_

namespace ZJGName
{
#include <afx.h>
#include "stdafx.h"
//#include "BasicClass.h"

#define BARCODEPRINTER "TSC TTP-244 Plus"
#define SIZETYPENUMBER	3		//�������͵�����
///////���������˳��õ��Ա����

CString DelDecimal(CString DecimalCStr);	//ɾ���ַ�����ֵ��С������
BOOL CStrIsNumber(CString InputCStr);		//�ж��Ƿ�������
BOOL CStrIsNumber(char * InputChar);		//�ж��Ƿ�������
CString StrDeleteSpace(CString InputCStr,char DeleteChar = ' ');//ȥ����β�Ŀո�
void StrDeleteSpace(char * InputChar,char DeleteChar = ' ');	//ȥ����β�Ŀո�
CString StrDeleteEndChar(CString InputCStr,char DeleteChar);	//ȥ����β��ָ���ַ�
void StrDeleteEndChar(char * InputChar,char DeleteChar);		//ȥ����β��ָ���ַ�
CString MonthIToCStr(int iMonth,BOOL Simplify);					//�������·�ת����Ӣ�ĵ���,SimplifyΪ�棬������
int CountCharQTY(CString InputCStr,char CountChar);				//ͳ��ָ���ַ��ĸ���		
int CountNumbrtChar(CString InputCStr);			//ͳ�������ַ��ĸ���		
double CentreLevel(CString& WordCStr,double iWord_H,double iPaper_W);//��ʼ�����ꣻiWord_H���ָߣ�iPaper_W��ֽ����λ��mm��

CString NumberToLetter(long Number);			//����ֵ������ת����Ӣ���ַ�����ֵ
CString NumberToLetter(double Number);			//����ֵ������ת����Ӣ���ַ�����ֵ,ֻ������λС��

int CutNameAdd(CString NameAddCStr,int MaxChar);	//�ָ˾�����ַ��
int CutNameAdd(TCHAR * NameAddCStr,int MaxChar);	//�ָ˾�����ַ��
BOOL DeleteFileCStr(CString mDelCStr);	//ɾ���ļ���ÿ��ָ���ַ�ǰ������
CString GetExeFileDir();				//��ȡ����ǰִ���ļ���
BOOL IsDebug();							//�жϵ�ǰ��ִ���ļ��Ƿ��ǵ����ļ�Ŀ¼ 
BOOL CheckDir(CString DirName);			//����ļ����Ƿ����
BOOL CreateDir(CString DirName);		//�����ļ���
BOOL CheckFile(CString FileName);		//����ļ��Ƿ����
BOOL zDeleteFile(CString FileName);		//ɾ��ָ���ļ�

CString GetFilePath();								//��ȡ�Ի���ѡ����ļ�·��
CTime StrToTime(CString nDateStr,char Separator);	//���ַ���ת����ʱ��
CTime StrToTime(CString Date);						//���ַ���ת����ʱ��

BOOL SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString&>& HeaderCStr,CArray<int,int>& iHeader);//�����б�ؼ��ı�ͷ�ַ�������
void SetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray,int iSizeCount);		//�����б�ؼ��ı�ͷ�ַ�������
//void SetListCtrlHeader(CListCtrl * pListCtrl,);							//�����б�ؼ��ı�ͷ�ַ�������

void GetListCtrlHeader(CListCtrl * pListCtrl,CArray<CString,CString> * pHeaderArray);						//�õ��б�ؼ��ı�ͷ�ַ�������
void GetListCtrlHeader(CListCtrl * pListCtrl,CList<CString,CString> * pHeaderList,int iStart = 0,int iEnd = -1);//�õ��б�ؼ��ı�ͷ�ַ�������
void SetListCtrlRowPos(CListCtrl * pListCtrl,int OldTopRow);//����ָ���У����б�ؼ��пɼ���

CString zGetPrinter(CString CStrPrinter);	//�õ�TSC��ӡ��
CString LongToCStr(long n);					//����ֵ������ת�����ַ���������
CString IntToCStr(int n);					//����ֵ������ת�����ַ���������
CString DoubleToCStr(double n);				//����ֵ������ת�����ַ���������

BOOL SendToClipboard(CString strText);		//���ı����ݸ��Ƶ�������
BOOL GetFromClipboard(CString & strText);	//�����������ݸ��Ƶ��ַ���

int CombinationNumber(int* p,int Number);		//�����������е�������ϳ�һ�������֣���"2,4,7"-->"247"
BOOL MakeDoNumber(int* pOriginal,int O_Number,int* pNew,int Number,long Total);
int NumberOfDigits(long Number);				//�����ֵ�λ��

void CStrToChar(char * TargetChar,CString OriginalCStr,BOOL bFlag = true);		//������������ת�����ַ�����

BOOL ShowJpg(CDC* pDC, CString strPath,CRect mRect);//��ʾJPGͼƬ
void ShowAccJPG(CString PictureFile,CDC * pDC,CRect ACCPictureRect);

int SizePos(CString nSize,CString nStandards);				//������ڱ�׼�����е�λ��
}

class CVolume //ֽ�������
{
public:
	CVolume();
	CVolume(int nLength,int nWidth,int nHeight);

	int GetLength(){return Length;};	//��ȡ��
	int GetWidth(){return Width;};		//��ȡ��
	int GetHeight(){return Height;};	//��ȡ��

	void SetLength(int nLength){ Length = nLength;};	//���ó�
	void SetWidth(int nWidth){ Width = nWidth;};		//���ÿ�
	void SetHeight(int nHeight){ Height = nHeight;};	//���ø�

	double GetVolume(){return Length/1000.0/1000000.0*Width*Height;};//�������λ:������
	
	void SetVolumeCStr(CString nVolumeCStr);	//���ַ����͵�ֽ����ת����ֽ����
	CString GetVolumeCStr();		//��ȡ�ַ����͵�ֽ����
	void Empty();	//���
	BOOL IsEmpty();	//�жϿ�
//	CVolume& operator=(CVolume& Other);		//��ֵ
	BOOL operator==(CVolume& Other);		//�Ƚ���ͬ 

private:
	int	Length;		//ֽ�䳤����λ������
	int	Width;		//ֽ�����λ������
	int	Height;		//ֽ��ߣ���λ������
};

#endif // defined(ZJGBASIC_H_INCLUDED_)