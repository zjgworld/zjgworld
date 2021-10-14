// ModiExcelDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ModiExcel.h"
#include "ModiExcelDlg.h"
#include "zjgBasic.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


using namespace ZJGName;

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CModiExcelDlg dialog

CModiExcelDlg::CModiExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CModiExcelDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CModiExcelDlg)
	m_Edit_ClassName_CStr = _T("");
	m_Static_SelFileName_CStr = _T("");
	m_List_FileName_CStr = _T("");
	m_Static_FileText_CStr = _T("");
	m_Static_SelRefFile_CStr = _T("");
	m_Static_ModiFile_CStr = _T("");
	m_CHECK_Sel_B = FALSE;
	m_CB_ClassName_CStr = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CModiExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CModiExcelDlg)
	DDX_Control(pDX, IDC_COMBO_ClassName, m_CB_ClassName_Ctrl);
//	DDX_Control(pDX, IDC_EDIT_ClassName, m_Edit_ClassName_Ctrl);
	DDX_Control(pDX, IDC_BUTTON_ModiFile, m_Btn_ModiFile_Ctrl);
	DDX_Control(pDX, IDC_BUTTON_RefFile, m_Btn_RelFile_Ctrl);
	DDX_Control(pDX, IDC_BUTTON_Test, m_Btn_Test_Ctrl);
	DDX_Control(pDX, IDC_CHECK_Sel, m_CHECK_Sel_Ctrl);
	DDX_Control(pDX, IDC_LIST_FileName, m_List_FileName_Ctrl);
	DDX_Control(pDX, IDC_BUTTON_Save, m_Btn_Save_Ctrl);
	DDX_Control(pDX, IDC_BUTTON_Next, m_Btn_Next_Ctrl);
	DDX_Control(pDX, IDC_BUTTON_All, m_Btn_All_Ctrl);
//	DDX_Text(pDX, IDC_EDIT_ClassName, m_Edit_ClassName_CStr);
	DDX_Text(pDX, IDC_STATIC_Selected, m_Static_SelFileName_CStr);
	DDX_LBString(pDX, IDC_LIST_FileName, m_List_FileName_CStr);
	DDX_Text(pDX, IDC_STATIC_FileText, m_Static_FileText_CStr);
	DDX_Text(pDX, IDC_STATIC_SelRefFile, m_Static_SelRefFile_CStr);
	DDX_Text(pDX, IDC_STATIC_ModiFile, m_Static_ModiFile_CStr);
	DDX_Check(pDX, IDC_CHECK_Sel, m_CHECK_Sel_B);
	DDX_CBString(pDX, IDC_COMBO_ClassName, m_CB_ClassName_CStr);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CModiExcelDlg, CDialog)
	//{{AFX_MSG_MAP(CModiExcelDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_Next, OnBUTTONNext)
	ON_BN_CLICKED(IDC_BUTTON_All, OnBUTTONAll)
	ON_BN_CLICKED(IDC_BUTTON_Save, OnBUTTONSave)
	ON_LBN_SELCHANGE(IDC_LIST_FileName, OnSelchangeLISTFileName)
	ON_BN_CLICKED(IDC_BUTTON_RefFile, OnBUTTONRefFile)
	ON_BN_CLICKED(IDC_BUTTON_ModiFile, OnBUTTONModiFile)
	ON_BN_CLICKED(IDC_BUTTON_Test, OnBUTTONTest)
	ON_BN_CLICKED(IDC_CHECK_Sel, OnCHECKSel)
	ON_CBN_SELCHANGE(IDC_COMBO_ClassName, OnSelchangeCOMBOClassName)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CModiExcelDlg message handlers

BOOL CModiExcelDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CModiExcelDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CModiExcelDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CModiExcelDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CModiExcelDlg::OnBUTTONNext() 
{
	// TODO: Add your control notification handler code here
	// TODO: Add your control notification handler code here
	int i = 0;
	int iCount = this->m_List_FileName_Ctrl.GetCount();
	int iResult; 
	int iSucceed = 0; 
	int iFail = 0;
	int iPos =0;
	CString strOldFileName;
	CString strNewFileName;
	CString strTem;
	ExClassLine CurExClassLine;
	CString strError;
	int iError[10];

	CStdioFile myFile;
	CStdioFile myIncludeFile;
	if(!myFile.Open(strFilePath + "-2\\修改记录.txt",CFile::modeCreate | CFile::modeWrite)  ) return;

	for( i = 0; i <10; i++) iError[i] =0;
	i = 0;
	CurExClassLine.Empty();
	while( i < iCount)
	{
		this->m_List_FileName_Ctrl.GetText(i,strTem);
		strOldFileName = strFilePath + "-1\\" + strTem;		//
		strNewFileName = strFilePath + "-2\\" + strTem;		//

		iResult = CurExClassLine.Modify_H_File_PUT(strOldFileName,strNewFileName);
		EroorPress(&myFile,strTem,iError,i,iResult,iSucceed,iFail);
		i++;
	}
	EroorSave(&myFile,iError,iCount,iSucceed,iFail);
	myFile.Close();
}


void CModiExcelDlg::OnBUTTONAll() 
{
	// TODO: Add your control notification handler code here
	int i = 0;
	int iCount = this->m_List_FileName_Ctrl.GetCount();
	int iResult; 
	int iSucceed = 0; 
	int iFail = 0;
	int iPos =0;
	CString strOldFileName;
	CString strNewFileName;
	CString strTem;

	int iError[10];
	for( i = 0; i <10; i++) iError[i] =0;

	i = 0;
	CStdioFile myFile;
	if(!myFile.Open(strFilePath + "-1\\修改记录.txt",CFile::modeCreate | CFile::modeWrite)  ) return;
	
	while( i < iCount)
	{
		this->m_List_FileName_Ctrl.GetText(i,strTem);
	
		strNewFileName = strFilePath + "-1\\" + strTem;		//
		strOldFileName = strFilePath + "\\" + strTem;
		
		iResult = mExClassLine.Modify_H_File_A(strOldFileName,strNewFileName);
		EroorPress(&myFile,strTem,iError,i,iResult,iSucceed,iFail);

		i++;
	}

	EroorSave(&myFile,iError,iCount,iSucceed,iFail);

	myFile.Close();

}
void CModiExcelDlg::EroorSave(CStdioFile* pFile,int * pError,int iCount,int iSucceed,int iFail)
{
	CString TemCStr;
	pFile->WriteString("\n\n");
	m_Static_FileText_CStr.Empty();
	m_Static_FileText_CStr.Format("已修改 %d 个文件，\n其中完全修改的有 %d 个文件,\n部分修改的有 %d 个文件",iCount,iSucceed,iFail);
	pFile->WriteString(m_Static_FileText_CStr);	

	m_Static_FileText_CStr = m_Static_FileText_CStr + "\n修改结束";
	this->UpdateData(false);

	TemCStr.Format(" %d 文件中存在 名称错误.\n",pError[0]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 读错误.\n",pError[1]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 写错误.\n",pError[2]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 缺少原类的错误.\n",pError[3]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 部分函数错误.\n",pError[4]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 变量类型不存在的错误.\n",pError[5]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 类中缺少方法接口ID的错误.\n",pError[6]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 链表中缺少方法名相近的接口ID错误.\n",pError[7]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 链表中缺少接口ID错误.\n",pError[8]);
	pFile->WriteString(TemCStr);
	TemCStr.Format(" %d 文件中存在 类中缺少成对的ID错误.\n",pError[9]);
	pFile->WriteString(TemCStr);
}

void CModiExcelDlg::EroorPress(CStdioFile* pFile,CString strTem,int * pError,
							   int iCur ,int iResult,int& iSucceed,int& iFail)
{
	if(iResult == FILE_FINISHED)	iSucceed++;
	else
	{
		iFail++;
		if(iResult & FILE_NAME_ERROR) 
		{
			pFile->WriteString(strTem + "含有 名称错误.\n");
			pError[0]++;
		}
		if(iResult & FILE_READ_ERROR) 
		{
			pFile->WriteString(strTem + "含有 读错误.\n");
			pError[1]++;
		}
		if(iResult & FILE_WRITE_ERROR) 
		{
			pFile->WriteString(strTem + "含有 写错误.\n");
			pError[2]++;
		}
		if(iResult & FILE_LESS_ERROR) 
		{
			pFile->WriteString(strTem + "含有 缺少原类.\n");
			pError[3]++;
		}
		if(iResult & FILE_PART_ERROR) 
		{
			pFile->WriteString(strTem + "含有 部分函数错误.\n");
			pError[4]++;
		}
		if(iResult & FILE_VARTYPE_ERR) 
		{
			pFile->WriteString(strTem + "含有 变量类型错误.\n");
			pError[5]++;
		}
		if(iResult & FILE_C_LESSID_ERR) 
		{
			pFile->WriteString(strTem + "含有 类中缺少方法接口ID错误.\n");
			pError[6]++;
		}
		if(iResult & FILE_LN_LESSID_ERR) 
		{
			pFile->WriteString(strTem + "含有 链表中缺少方法名相近的接口ID错误.\n");
			pError[7]++;
		}
		if(iResult & FILE_LI_LESSID_ERR) 
		{
			pFile->WriteString(strTem + "含有 链表中缺少接口ID错误.\n");
			pError[8]++;
		}
		if(iResult & FILE_C_LESSPUT_ERR) 
		{
			pFile->WriteString(strTem + "含有 类中缺少成对的ID错误.\n");
			pError[9]++;
		}
		pFile->WriteString("\n");
		}
		m_Static_FileText_CStr.Empty();
		m_Static_FileText_CStr.Format("已修改 %d 个文件，\n其中完整修改的有 %d 个文件,\n部分修改的有 %d 个文件",iCur,iSucceed,iFail);
		m_Static_FileText_CStr = m_Static_FileText_CStr + "\n正在修改的文件是" + strTem;
		this->UpdateData(false);
}

void CModiExcelDlg::OnBUTTONSave() 
{
	// TODO: Add your control notification handler code here
	CString strOldFileName;
	CString strNewFileName;
	CString strTem,ReadLineCStr;
	
	CStdioFile myReadFile,myWriteFile;
	strNewFileName = strFilePath + "-2\\Excel.h";		//
	if(! myWriteFile.Open(strNewFileName,CFile::modeCreate|CFile::modeWrite)  )  return;

	myWriteFile.WriteString("#pragma once\n");
//	myWriteFile.WriteString(CString("#include \"Excel_H.h\"\n"));

//	myWriteFile.WriteString(CString("namespace NS_Excel16\n"));
//	myWriteFile.WriteString(CString(_T("{\n")));
//	myWriteFile.WriteString(CString("\n"));

	int iCount = this->m_List_FileName_Ctrl.GetCount();
	int i = 0;
	while( i < iCount)
	{
		this->m_List_FileName_Ctrl.GetText(i,strTem);
		strOldFileName = strFilePath + "-2\\" + strTem;
		if(! myReadFile.Open(strOldFileName,CFile::modeRead)  )  continue;

		while(myReadFile.ReadString(ReadLineCStr))
		{
			if(ReadLineCStr.Find("#include ") >= 0)  continue;
			myWriteFile.WriteString(ReadLineCStr + "\n");
		}
		myWriteFile.WriteString("\n\n");
		myReadFile.Close();
		i++;
	
		m_Static_FileText_CStr.Empty();
		m_Static_FileText_CStr.Format("已合并 %d 个文件，共有 %d 个文件",i,iCount);
		m_Static_FileText_CStr = m_Static_FileText_CStr + "\n正在修改的文件是" + strTem;
		this->UpdateData(false);

	}
	m_Static_FileText_CStr.Format("已合并 %d 个文件，共有 %d 个文件",i,iCount);
	m_Static_FileText_CStr = m_Static_FileText_CStr + "\n合并结束";
	this->UpdateData(false);

	myWriteFile.WriteString(CString("\n}//namespace ZJGName命名空间结束\n\n"));
	myWriteFile.Close();		
}

void CModiExcelDlg::OnOK() 
{
	// TODO: Add extra validation here
//	OnBUTTONSave();
	CDialog::OnOK();
}

void CModiExcelDlg::OnCancel() 
{
	// TODO: Add extra cleanup here
	
	CDialog::OnCancel();
}


void CModiExcelDlg::OnSelchangeLISTFileName() 
{
	// TODO: Add your control notification handler code here
	this->UpdateData(true);
	//获取文件名
	
	int iCurSel = this->m_List_FileName_Ctrl.GetCurSel();
	if(iCurSel < 0 || iCurSel >= m_List_FileName_Ctrl.GetCount()) return;
	m_List_FileName_Ctrl.GetText(iCurSel,m_Static_SelFileName_CStr);


	this->mExClassLine.FillStaticCtrl(&m_Static_FileText_CStr,strFilePath+'\\'+m_Static_SelFileName_CStr);
	if(!(ZJGName::CheckDir(strFilePath + "-1\\")))
	{
	}

	this->UpdateData(false);
	
}

void CModiExcelDlg::OnBUTTONRefFile() 
{
	// TODO: Add your control notification handler code here
	CFileDialog myFileDialog(TRUE,NULL,"",OFN_HIDEREADONLY,"源程序文件(*.cpp)|*.cpp||头文件(*.h)|*.h",NULL);
	if(myFileDialog.DoModal()!=IDOK) return;
	this->m_Static_SelRefFile_CStr = myFileDialog.GetPathName();

	mExClassLine.LoadCPPFile(m_Static_SelRefFile_CStr);

	this->UpdateData(false);
		
}

void CModiExcelDlg::OnBUTTONModiFile() 
{
	// TODO: Add your control notification handler code here
	//装载文件名
	strFilePath = ZJGName::GetFilePath();
	m_CHECK_Sel_B = false;
	this->UpdateData(false);

	OnCHECKSel();
}

void CModiExcelDlg::OnBUTTONTest() 
{
	// TODO: Add your control notification handler code here
	enum enum_a{a1 = 2,a2=2,a3=3,a4=1};
	enum_a a,b,c,d;
	b = a1;
	c = a2;
	if(b==c) 
		d = b;
	else
		d = a3;
	DecEnumList myEnumList;
	myEnumList.Load_Txt_File("E:\\VC程序\\枚举-Excel.txt");
	myEnumList.Create_H("E:\\VC程序\\枚举-Excel.h");

	return;
}

void CModiExcelDlg::OnCHECKSel() 
{
	// TODO: Add your control notification handler code here
	this->UpdateData(true);
	m_List_FileName_Ctrl.ResetContent();
	if(this->m_CHECK_Sel_B)
	{
		//装载修改记录文件中的错误文件，到列表框中
		this->mExClassLine.FillListCtrl(&m_List_FileName_Ctrl,
										strFilePath + "\\修改记录.txt");
		this->mExClassLine.FillCBCtrl(&this->m_CB_ClassName_Ctrl);
	}
	else
	{
		//装载指定待修改文件夹中文件，到列表框中
		CFileFind myFindFile;
		BOOL bFindFile;
		CString TemCStr;

		TemCStr = strFilePath+"\\*.h";
		bFindFile = myFindFile.FindFile(TemCStr);
		while(bFindFile)
		{	
			bFindFile = myFindFile.FindNextFile();
			TemCStr = myFindFile.GetFileName();
			this->m_List_FileName_Ctrl.AddString(TemCStr);
		}
	}
}

void CModiExcelDlg::OnSelchangeCOMBOClassName() 
{
	// TODO: Add your control notification handler code here
	this->UpdateData(true);
	
	CString TemTypeName,TemCStr;
	int iCulSel = this->m_CB_ClassName_Ctrl.GetCurSel();
	if(iCulSel < 0) return;
	m_CB_ClassName_Ctrl.GetLBText(iCulSel,TemTypeName);
	m_CB_ClassName_CStr = TemTypeName;

	ExMethod TemExMethod;
	ExClass TemExClass;
	POSITION pPos,pMethodPos; 
	pPos = this->mExClassLine.Find(m_CB_ClassName_CStr);
	if(pPos) TemExClass = mExClassLine.mLine.GetAt(pPos);
	else return;
	
	m_Static_FileText_CStr.Empty();

	pPos = TemExClass.mLine.GetHeadPosition();
	while(pPos)
	{
		TemExMethod = TemExClass.mLine.GetNext(pPos);

		pMethodPos = TemExMethod.mLine.GetHeadPosition();
		while(pMethodPos)
		{
			TemCStr = TemExMethod.mLine.GetNext(pMethodPos);
			this->m_Static_FileText_CStr += TemCStr + " \n";
		}
		this->m_Static_FileText_CStr +=  " \n";
	}

	this->UpdateData(false);
}
