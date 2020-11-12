
// findsnDlg.h : 头文件
//
#include "workthread.h"
#pragma once
#include "CFont0.h"
#include<odbcinst.h> 
#include<afxdb.h>
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CApplication.h"


#define COLOR_RED		RGB(255, 0, 0)
#define COLOR_GREEN		RGB(0, 255 ,0)
#define COLOR_WHITE		RGB(0, 0, 0)
#define COLOR_YELLOW	RGB(255, 255, 0)
#define COLOR_WHITE		RGB(255, 255, 255)

#define SAVE_INFO_PASS_FILE '1'
#define SAVE_INFO_FAIL_FILE	  '2'

#define RET_EXCEL_START_APP_SUCCESS				0
#define RET_EXCEL_START_APP_FAIL				1
#define RET_EXCEL_START_OPEN_SOURCE_FILE_FAIL	2

// CfindsnDlg 对话框
class CfindsnDlg : public CDialogEx
{
// 构造
public:
	CfindsnDlg(CWnd* pParent = NULL);	// 标准构造函数
	afx_msg LRESULT OnUserMsg(WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	UINT8 startExcelApp();

	CString m_strEdit1;
	CEdit m_ctlEdit1;
	CString m_strEdit2;
	CEdit m_ctlEdit2;

	CRichEditCtrl	m_ctrlRedit;
	CString m_strRedit;
// 对话框数据
	enum { IDD = IDD_FINDSN_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持
	BOOL mOperate();
	threadInfo m_info;
	CApplication app; 
    CWorkbook book;  
    CWorkbooks books;  
    CWorksheet sheet;  
    CWorksheets sheets; 
    CRange range; 
	CRange tRange;
	CRange rangeFind;
	CRange rangeRow;
    CFont0 font;   
    CRange cols; 
    LPDISPATCH lpDisp;
	LPDISPATCH lpDispFind;

	char mSaveFilePassName[100];
	char mSaveFileFailName[100];
	char mLogFileName[100];
	char workPath[100];
	char snSrcFile[100];
// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	POINT	Old;
	void	resize();
};
