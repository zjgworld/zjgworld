// ModiExcelDlg.h : header file
//

#if !defined(AFX_MODIEXCELDLG_H__CBC59F1B_72FB_44FE_8586_D98FC367E746__INCLUDED_)
#define AFX_MODIEXCELDLG_H__CBC59F1B_72FB_44FE_8586_D98FC367E746__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CModiExcelDlg dialog
#include "BasicClass.h"

class CModiExcelDlg : public CDialog
{
// Construction
public:
	CModiExcelDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CModiExcelDlg)
	enum { IDD = IDD_MODIEXCEL_DIALOG };
	CComboBox	m_CB_ClassName_Ctrl;
	CEdit	m_Edit_ClassName_Ctrl;
	CButton	m_Btn_ModiFile_Ctrl;
	CButton	m_Btn_RelFile_Ctrl;
	CButton	m_Btn_Test_Ctrl;
	CButton	m_CHECK_Sel_Ctrl;
	CListBox	m_List_FileName_Ctrl;
	CButton	m_Btn_Save_Ctrl;
	CButton	m_Btn_Next_Ctrl;
	CButton	m_Btn_All_Ctrl;
	CString	m_Edit_ClassName_CStr;
	CString	m_Static_SelFileName_CStr;
	CString	m_List_FileName_CStr;
	CString	m_Static_FileText_CStr;
	CString	m_Static_SelRefFile_CStr;
	CString	m_Static_ModiFile_CStr;
	BOOL	m_CHECK_Sel_B;
	CString	m_CB_ClassName_CStr;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CModiExcelDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CModiExcelDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnBUTTONNext();
	afx_msg void OnBUTTONAll();
	afx_msg void OnBUTTONSave();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnSelchangeLISTFileName();
	afx_msg void OnBUTTONRefFile();
	afx_msg void OnBUTTONModiFile();
	afx_msg void OnBUTTONTest();
	afx_msg void OnCHECKSel();
	afx_msg void OnSelchangeCOMBOClassName();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	CString strFilePath;
	CString strRecord;

	ExClassLine mExClassLine;
	void EroorPress(CStdioFile* pFile,CString strTem,int * pError,
					 int iCur ,int iResult,int& iSucceed,int& iFail);
	void EroorSave(CStdioFile* pFile,int * pError,int iCount,int iSucceed,int iFail);

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MODIEXCELDLG_H__CBC59F1B_72FB_44FE_8586_D98FC367E746__INCLUDED_)
