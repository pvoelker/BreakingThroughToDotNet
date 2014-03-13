
// KeypadControlTestAppDlg.h : header file
//

#pragma once
#include "ckeypadcontrol.h"


// CKeypadControlTestAppDlg dialog
class CKeypadControlTestAppDlg : public CDialogEx
{
// Construction
public:
	CKeypadControlTestAppDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_KEYPADCONTROLTESTAPP_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CKeypadControl m_Keypad;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
};
